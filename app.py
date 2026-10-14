import streamlit as st
import pandas as pd
from sqlalchemy import create_engine
import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().with_name(".env"))

# --- Database Connection Setup ---
@st.cache_resource
def init_connection():
    # Use the synchronous MySQL dialect because pd.read_sql() is running synchronously.
    # mysql+aiomysql triggers the async SQLAlchemy stack and requires greenlet.
    db_url = os.getenv("DATABASE_URL")
    if not db_url:
        raise RuntimeError(
            "DATABASE_URL is not set. Add it to payment_check/.env or set it "
            "as an environment variable."
        )
    return create_engine(db_url)

engine = init_connection()

st.title("Telebirr Reconciliation Dashboard")
st.write("Upload the 'oct 10.xlsx' (or similar) Telebirr statement and select a date range to reconcile against the `customer_transactions` table.")

# --- UI Inputs ---
col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Start Date")
with col2:
    end_date = st.date_input("End Date")

uploaded_file = st.file_uploader("Upload Telebirr Statement (Excel)", type=["xlsx", "xls"])

if uploaded_file and st.button("Run Reconciliation"):
    try:
        # 1. Read the uploaded Excel file
        # The Telebirr statement has metadata rows above the actual table,
        # and the real column headers begin on row 6 (1-based indexing).
        excel_df = pd.read_excel(uploaded_file, header=5)

        # Normalize column names and support small variations in the header text
        excel_df.columns = [str(col).strip() for col in excel_df.columns]

        if 'Receipt No.' not in excel_df.columns:
            receipt_candidates = [
                c for c in excel_df.columns
                if 'receipt' in str(c).lower() and 'no' in str(c).lower()
            ]
            if receipt_candidates:
                excel_df = excel_df.rename(columns={receipt_candidates[0]: 'Receipt No.'})

        if 'Receipt No.' not in excel_df.columns:
            st.error("The uploaded Excel file must contain a 'Receipt No.' column in row 6.")
            st.stop()

        # Clean the Receipt No. column to ensure perfect matching
        excel_df['Receipt No.'] = excel_df['Receipt No.'].fillna('').astype(str).str.strip()
        excel_df = excel_df[excel_df['Receipt No.'] != '']

        # 2. Fetch data from DB based on selected date range and payment method
        # Note: Depending on your DB driver (MySQL vs PostgreSQL), you might need to change %s to ? or %(param)s
        query = """
            SELECT 
                ct.session_id,
                ct.order_id,
                ct.transaction_status,
                o.payment_method,
                ct.created_at
            FROM customer_transactions ct
            LEFT JOIN orders o ON o.id = ct.order_id
            WHERE DATE(ct.created_at) >= %s
              AND DATE(ct.created_at) <= %s
              AND o.payment_method IN ('telebirr-ussd', 'telebirr')
        """
        
        # Fetch dataframe
        db_df = pd.read_sql(query, engine, params=(start_date, end_date))
        
        # Clean session_id for matching
        db_df['session_id'] = db_df['session_id'].astype(str).str.strip()
        
        # Get only the successful transactions for baseline comparison
        db_success_df = db_df[db_df['transaction_status'].str.lower() == 'success']

        # --- RECONCILIATION LOGIC ---

        # Check 1: In Database as success/paid, but NOT in Telebirr Excel Statement
        st.subheader("Check 1: Paid in Database, Missing in Telebirr Statement")
        missing_in_telebirr = db_success_df[~db_success_df['session_id'].isin(excel_df['Receipt No.'])].copy()
        
        # Styling function for red highlights
        def highlight_red(row):
            return ['background-color: #ffcccc; color: #990000; font-weight: bold'] * len(row)

        if not missing_in_telebirr.empty:
            missing_in_telebirr['Mismatch Reason'] = "Doesn't exist in Telebirr statement"
            # Apply styling and display
            st.dataframe(
                missing_in_telebirr.style.apply(highlight_red, axis=1),
                use_container_width=True
            )
        else:
            st.success("All successful telebirr database transactions are present in the Excel statement!")

        # Check 2: In Telebirr Excel Statement, but NOT in Database (or not marked as 'success')
        st.subheader("Check 2: In Telebirr Statement, Missing or Failed in Database")
        missing_in_db = excel_df[~excel_df['Receipt No.'].isin(db_success_df['session_id'])].copy()

        if not missing_in_db.empty:
            missing_in_db['Mismatch Reason'] = "Not in customer transaction or status != success"
            
            # Map DB statuses to the excel view for context (optional but highly useful)
            # Find what the status actually is in the DB, if it exists at all
            db_status_map = db_df.set_index('session_id')['transaction_status'].to_dict()
            missing_in_db['Current DB Status'] = missing_in_db['Receipt No.'].map(db_status_map).fillna("Missing entirely")
            
            # Apply styling and display
            st.dataframe(
                missing_in_db.style.apply(highlight_red, axis=1),
                use_container_width=True
            )
        else:
            st.success("All receipts in the Telebirr statement matched successfully in the database!")

    except Exception as e:
        st.error(f"An error occurred during reconciliation: {e}")