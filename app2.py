import streamlit as st
import pandas as pd
from io import BytesIO

# -----------------------------------
# PAGE CONFIG
# -----------------------------------
st.set_page_config(
    page_title="Loan Repayment Schedule Generator",
    layout="wide"
)

# -----------------------------------
# SAFE TYPE HELPERS
# -----------------------------------
def safe_float(x, default=0.0):
    try:
        return float(str(x).replace(",", "").strip())
    except:
        return default

def safe_int(x, default=0):
    try:
        return int(float(str(x).replace(",", "").strip()))
    except:
        return default


# -----------------------------------
# LOAD LOAN DATA
# -----------------------------------
def load_loan_data(file):
    df = pd.read_excel(file, header=None)
    df = df.dropna(how="all")

    data = {}

    for _, row in df.iterrows():
        if pd.isna(row[0]) or pd.isna(row[1]):
            continue

        key = str(row[0]).strip().lower()
        value = row[1]

        data[key] = value

    return data


# -----------------------------------
# EMI CALCULATION (MULTI-FREQUENCY)
# -----------------------------------
def generate_schedule(principal, annual_rate, tenure_years, frequency):

    freq_map = {
        "Daily": 360,
        "Monthly": 12,
        "Quarterly": 4,
        "Yearly": 1
    }

    periods_per_year = freq_map[frequency]
    rate = annual_rate / periods_per_year / 100
    total_periods = tenure_years * periods_per_year

    if rate == 0:
        emi = principal / total_periods
    else:
        emi = (
            principal
            * rate
            * (1 + rate) ** total_periods
        ) / (
            (1 + rate) ** total_periods - 1
        )

    schedule = []
    opening_balance = principal
    total_interest = 0

    for i in range(1, int(total_periods) + 1):

        interest = opening_balance * rate
        principal_paid = min(emi - interest, opening_balance)
        closing_balance = opening_balance - principal_paid

        total_interest += interest

        schedule.append({
            "Installment No": i,
            "Opening Balance": round(opening_balance, 2),
            "EMI": round(emi, 2),
            "Interest": round(interest, 2),
            "Principal": round(principal_paid, 2),
            "Closing Balance": round(max(closing_balance, 0), 2)
        })

        opening_balance = closing_balance

        if opening_balance <= 0:
            break

    return pd.DataFrame(schedule), emi, total_interest


# -----------------------------------
# EXCEL EXPORT
# -----------------------------------
def create_excel_download(schedule_df, customer_name, principal, rate, tenure, emi, total_interest, frequency, installments):

    output = BytesIO()

    summary_df = pd.DataFrame({
        "Field": [
            "Customer Name",
            "Principal",
            "Interest Rate",
            "Tenure (Years)",
            "Repayment Frequency",
            "Number of Installments",
            "Monthly EMI Equivalent",
            "Total Interest"
        ],
        "Value": [
            customer_name,
            principal,
            rate,
            tenure,
            frequency,
            installments,
            round(emi, 2),
            round(total_interest, 2)
        ]
    })

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        summary_df.to_excel(writer, sheet_name="Loan Report", index=False, startrow=0)

        start_row = len(summary_df) + 3

        schedule_df.to_excel(
            writer,
            sheet_name="Loan Report",
            index=False,
            startrow=start_row
        )

    output.seek(0)
    return output


# -----------------------------------
# UI
# -----------------------------------
st.title("📊 Loan Repayment Schedule Generator")

uploaded_file = st.file_uploader(
    "Drag and Drop Excel File",
    type=["xlsx", "xls"]
)

# -----------------------------------
# PROCESS
# -----------------------------------
if uploaded_file:

    try:
        data = load_loan_data(uploaded_file)

        customer_name = str(data.get("customer name", "Unknown"))
        principal = safe_float(data.get("principal"))
        annual_rate = safe_float(data.get("interest rate"))
        tenure = safe_int(data.get("tenure"))

        # Frequency INSIDE summary logic (default monthly)
        frequency = "Monthly"

        freq_map = {
            "Daily": 360,
            "Monthly": 12,
            "Quarterly": 4,
            "Yearly": 1
        }

        periods_per_year = freq_map[frequency]
        installments = tenure * periods_per_year

        # -----------------------------
        # LOAN SUMMARY UI
        # -----------------------------
        st.subheader("📌 Loan Summary")

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric("Customer", customer_name)

        with col2:
            st.metric("Principal", f"{principal:,.2f}")

        with col3:
            st.metric("Interest Rate", f"{annual_rate}%")

        with col4:
            st.metric("Tenure", f"{tenure} Years")

        col5, col6 = st.columns(2)

        with col5:
            frequency = st.selectbox(
                "Repayment Frequency",
                ["Monthly", "Quarterly", "Yearly", "Daily"]
            )

        # recompute installments dynamically
        periods_per_year = freq_map[frequency]
        installments = tenure * periods_per_year

        with col6:
            st.metric("Number of Installments", installments)

        st.markdown("---")

        # -----------------------------------
        # CALCULATION
        # -----------------------------------
        schedule_df, emi, total_interest = generate_schedule(
            principal,
            annual_rate,
            tenure,
            frequency
        )

        # -----------------------------------
        # EXTRA SUMMARY METRICS
        # -----------------------------------
        col7, col8 = st.columns(2)

        with col7:
            st.metric("Per Installment Amount", f"{emi:,.2f}")

        with col8:
            st.metric("Total Interest", f"{total_interest:,.2f}")

        st.markdown("---")

        # -----------------------------------
        # TABLE
        # -----------------------------------
        st.subheader("📊 Repayment Schedule")

        st.dataframe(
            schedule_df,
            use_container_width=True,
            hide_index=True
        )

        # -----------------------------------
        # DOWNLOAD
        # -----------------------------------
        excel_file = create_excel_download(
            schedule_df,
            customer_name,
            principal,
            annual_rate,
            tenure,
            emi,
            total_interest,
            frequency,
            installments
        )

        st.download_button(
            label="⬇ Download Excel Report",
            data=excel_file,
            file_name=f"{customer_name}_loan_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error processing file: {e}")