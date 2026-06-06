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
# LOAD LOAN DATA
# -----------------------------------
def load_loan_data(file):
    df = pd.read_excel(file, header=None)
    df = df.dropna(how="all")

    data = {}
    current_label = None

    for _, row in df.iterrows():

        label = str(row[0]).strip() if pd.notna(row[0]) else ""
        value = row[1] if len(row) > 1 else None

        if pd.isna(value):
            if current_label:
                current_label += " " + label
            else:
                current_label = label
        else:
            if current_label:
                label = current_label + " " + label
                current_label = None

            data[label.lower().strip()] = value

    return data


# -----------------------------------
# EMI CALCULATION
# -----------------------------------
def generate_schedule(principal, annual_rate, months):

    monthly_rate = annual_rate / 12 / 100

    if monthly_rate == 0:
        emi = principal / months
    else:
        emi = (
            principal
            * monthly_rate
            * (1 + monthly_rate) ** months
        ) / (
            (1 + monthly_rate) ** months - 1
        )

    schedule = []
    opening_balance = principal
    total_interest = 0

    for i in range(1, months + 1):

        interest = opening_balance * monthly_rate
        principal_paid = emi - interest
        closing_balance = opening_balance - principal_paid

        if i == months:
            closing_balance = 0

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

    return pd.DataFrame(schedule), emi, total_interest


# -----------------------------------
# SINGLE SHEET EXPORT (FIXED)
# -----------------------------------
def create_excel_download(schedule_df, customer_name, principal, rate, tenure, emi, total_interest):

    output = BytesIO()

    # CUSTOMER SUMMARY AS DATAFRAME
    summary_df = pd.DataFrame({
        "Field": [
            "Customer Name",
            "Principal",
            "Interest Rate",
            "Tenure (Months)",
            "Monthly EMI",
            "Total Interest"
        ],
        "Value": [
            customer_name,
            principal,
            rate,
            tenure,
            round(emi, 2),
            round(total_interest, 2)
        ]
    })

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        # Write summary FIRST (top of sheet)
        summary_df.to_excel(writer, sheet_name="Loan Report", index=False, startrow=0)

        # Leave gap then write schedule
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
        principal = float(data.get("principal", 0))
        annual_rate = float(data.get("interest rate", 0))
        tenure = int(data.get("tenure", 0))

        schedule_df, emi, total_interest = generate_schedule(
            principal,
            annual_rate,
            tenure
        )

        # -----------------------------------
        # SUMMARY UI
        # -----------------------------------
        st.subheader("📌 Loan Summary")

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric("Customer", customer_name)

        with col2:
            st.metric("Principal", f"{principal:,.2f}")

        with col3:
            st.metric("Interest Rate", f"{annual_rate}%")

        with col4:
            st.metric("Tenure", f"{tenure} Months")

        col5, col6 = st.columns(2)

        with col5:
            st.metric("Monthly EMI", f"{emi:,.2f}")

        with col6:
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
        # DOWNLOAD (SINGLE SHEET)
        # -----------------------------------
        excel_file = create_excel_download(
            schedule_df,
            customer_name,
            principal,
            annual_rate,
            tenure,
            emi,
            total_interest
        )

        st.download_button(
            label="⬇ Download Excel Report",
            data=excel_file,
            file_name=f"{customer_name}_loan_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error processing file: {e}")