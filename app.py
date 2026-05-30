import streamlit as st
import pandas as pd
from datetime import datetime
import os

# =====================================
# CONFIGURATION
# =====================================

st.set_page_config(
    page_title="KalaMudra",
    layout="wide"
)

EXCEL_FILE = "investors.xlsx"


MAX_DURATION = 60

# =====================================
# CREATE EXCEL FILE IF NOT EXISTS
# =====================================

def create_excel():

    if not os.path.exists(EXCEL_FILE):

        columns = [
            "CLIENT NAME",
            "CLIENT CONTACT NO.",
            "DATE OF INVESTMENT",
            "INVESTED AMOUNT",
            "PLAN",
            "PLAN DURATION",
            "PROFIT AMT",
            "PROFIT PERCENTAGE",
            "CAPITAL RETURN",
            "CAPITAL_PAID"
        ]

        for i in range(1, MAX_DURATION + 1):

            columns.append(f"PROFIT DATE {i}")
            columns.append(f"PAID_{i}")
            columns.append(f"PAYMENT_DATE_{i}")

        pd.DataFrame(columns=columns).to_excel(
            EXCEL_FILE,
            index=False
        )

create_excel()

# =====================================
# LOAD DATA
# =====================================

def load_data():
    return pd.read_excel(EXCEL_FILE)

def save_data(df):

    df.to_excel(
        EXCEL_FILE,
        index=False
    )

    st.success("Data Saved")

df = load_data()

# Fix data types for payment columns
for i in range(1, MAX_DURATION + 1):

    paid_col = f"PAID_{i}"
    payment_col = f"PAYMENT_DATE_{i}"

    if paid_col in df.columns:

        df[paid_col] = (
            df[paid_col]
            .fillna(False)
            .astype(bool)
        )

    if payment_col in df.columns:

        df[payment_col] = (
            df[payment_col]
            .fillna("")
            .astype(str)
        )

if "CAPITAL_PAID" in df.columns:

    df["CAPITAL_PAID"] = (
        df["CAPITAL_PAID"]
        .fillna(False)
        .astype(bool)
    )

# Force payment date columns to string type
for i in range(1, MAX_DURATION + 1):

    payment_col = f"PAYMENT_DATE_{i}"

    if payment_col in df.columns:

        df[payment_col] = (
            df[payment_col]
            .fillna("")
            .astype(object)
        )
        


# =====================================
# TITLE
# =====================================

st.title("💰 KalaMudra")

# =====================================
# MENU
# =====================================

menu = st.sidebar.radio(
    "Menu",
    [
        "Dashboard",
        "Add Investor",
        "Investors",
        "Payment Tracker",
        "Reports"
    ]
)

# =====================================
# DASHBOARD
# =====================================

if menu == "Dashboard":

    total_investors = len(df)

    total_investment = pd.to_numeric(
        df["INVESTED AMOUNT"],
        errors="coerce"
    ).fillna(0).sum()

    today = pd.Timestamp.today().date()

    due_today = 0
    overdue = 0

    for _, row in df.iterrows():

        duration = int(
            row["PLAN DURATION"]
        ) if pd.notna(
            row["PLAN DURATION"]
        ) else 0

        for i in range(1, duration + 1):

            date_col = f"PROFIT DATE {i}"
            paid_col = f"PAID_{i}"

            if pd.notna(row[date_col]):

                due_date = pd.to_datetime(
                    row[date_col]
                ).date()

                if not row[paid_col]:

                    if due_date == today:
                        due_today += 1

                    elif due_date < today:
                        overdue += 1

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "Total Investors",
        total_investors
    )

    c2.metric(
        "Total Investment",
        f"₹{total_investment:,.0f}"
    )

    c3.metric(
        "Due Today",
        due_today
    )

    c4.metric(
        "Overdue",
        overdue
    )

# =====================================
# ADD INVESTOR
# =====================================

elif menu == "Add Investor":

    st.subheader("Add New Investor")

    col1, col2 = st.columns(2)

    with col1:

        client_name = st.text_input(
            "Client Name"
        )

        mobile = st.text_input(
            "Mobile Number"
        )

        investment_date = st.date_input(
            "Investment Date"
        )

        amount = st.number_input(
            "Invested Amount",
            min_value=0.0
        )

        plan = st.text_input(
            "Plan"
        )

    with col2:

        payout_type = st.selectbox(
            "Payout Type",
            [
                "Monthly",
                "Quarterly",
                "Half-Yearly",
                "Yearly"
            ]
        )

        cycles = st.number_input(
            "Number of Payouts",
            min_value=1,
            value=1
        )

        if payout_type == "Monthly":

            months_gap = 1

        elif payout_type == "Quarterly":

            months_gap = 3

        elif payout_type == "Half-Yearly":

            months_gap = 6

        elif payout_type == "Yearly":

            months_gap = 12


        profit_percentage = st.number_input(
            "Profit Percentage (%)",
            min_value=0.0,
            step=0.1
        )

        profit_amt = (
            amount * profit_percentage / 100
        )

        st.info(
            f"Total Profit Amount : ₹{profit_amt:,.2f}"
        )

        capital_return = st.date_input(
            "Capital Return Date"
        )

        Referral_name = st.text_input(
            "Referral Name"
        )


    monthly_profit = 0

    if cycles > 0:

        monthly_profit = (
            profit_amt / cycles
        )

    st.success(
        f"Profit Payout: ₹{monthly_profit:,.2f}"
    )


    from dateutil.relativedelta import relativedelta

    st.subheader(
        "Profit Payout Schedule"
    )

    payout_dates = []

    for i in range(cycles):

        default_date = (
            investment_date +
            relativedelta(
                months=(i+1)*months_gap
            )
        )

        payout_date = st.date_input(
            f"Profit Date {i+1}",
            value=default_date.date()
            if hasattr(default_date, "date")
            else default_date,
            key=f"profit_{i}"
        )
        payout_dates.append(
            payout_date
        )
        st.info(
            f"Profit Date {i+1} : "
            f"{payout_date.strftime('%d-%m-%Y')}"
        )

    if st.button(
        "Save Investor"
    ):

        row = {}

        for col in df.columns:
            row[col] = ""

        row["CLIENT NAME"] = client_name
        row["REFERRAL NAME"] = Referral_name
        row["CLIENT CONTACT NO."] = mobile
        row["DATE OF INVESTMENT"] = investment_date
        row["INVESTED AMOUNT"] = amount
        row["PLAN"] = plan
        row["PAYOUT TYPE"] = payout_type
        row["NUMBER OF PAYOUTS"] = cycles
        # row["PLAN DURATION"] = duration
        row["PROFIT AMT"] = profit_amt
        row["PROFIT PERCENTAGE"] = profit_percentage
        row["CAPITAL RETURN"] = capital_return
        row["CAPITAL_PAID"] = False

        for i in range(cycles):

            row[
                f"PROFIT DATE {i+1}"
            ] = payout_dates[i]

            row[
                f"PAID_{i+1}"
            ] = False

            row[
                f"PAYMENT_DATE_{i+1}"
            ] = ""

        df = pd.concat(
            [
                df,
                pd.DataFrame([row])
            ],
            ignore_index=True
        )

        save_data(df)

        st.success(
            "Investor Added Successfully"
        )

        st.rerun()

# =====================================
# INVESTORS
# =====================================

elif menu == "Investors":

    st.subheader(
        "Investor List"
    )

    latest_df = pd.read_excel(
        EXCEL_FILE
    )

    search = st.text_input(
        "Search Investor"
    )

    if search:

        latest_df = latest_df[
            latest_df[
                "CLIENT NAME"
            ]
            .astype(str)
            .str.contains(
                search,
                case=False,
                na=False
            )
        ]

    st.dataframe(
        latest_df,
        use_container_width=True
    )

# =====================================
# PAYMENT TRACKER
# =====================================

elif menu == "Payment Tracker":

    if len(df) == 0:

        st.warning("No investors found")

    else:

        client = st.selectbox(
            "Select Client",
            df["CLIENT NAME"],
            key="payment_tracker_client"
        )

        idx = df[
            df["CLIENT NAME"] == client
        ].index[0]

        duration = int(
            df.at[idx, "PLAN DURATION"]
        )

        total_profit = float(
            df.at[idx, "PROFIT AMT"]
        )

        installment_amount = (
            total_profit / duration
            if duration > 0 else 0
        )

        st.subheader(
            f"Investor : {client}"
        )

        c1, c2, c3 = st.columns(3)

        c1.metric(
            "Investment Amount",
            f"₹{float(df.at[idx,'INVESTED AMOUNT']):,.0f}"
        )

        c2.metric(
            "Profit Amount",
            f"₹{total_profit:,.0f}"
        )

        c3.metric(
            "Installment Amount",
            f"₹{installment_amount:,.0f}"
        )

        st.divider()

        payment_updates = {}

        today = datetime.now().date()

        for i in range(1, duration + 1):

            date_col = f"PROFIT DATE {i}"
            paid_col = f"PAID_{i}"
            payment_col = f"PAYMENT_DATE_{i}"

            due_date = pd.to_datetime(
                df.at[idx, date_col]
            )

            current_paid = bool(
                df.at[idx, paid_col]
            )

            st.markdown(
                f"### Installment {i}"
            )

            col1, col2, col3 = st.columns(
                [2,2,1]
            )

            col1.write(
                f"Due Date : {due_date.strftime('%d-%m-%Y')}"
            )

            col2.write(
                f"Amount : ₹{installment_amount:,.2f}"
            )

            payment_updates[i] = col3.checkbox(
                "Paid",
                value=current_paid,
                key=f"installment_{i}"
            )

            if current_paid:

                st.success(
                    f"Paid On : {df.at[idx,payment_col]}"
                )

            else:

                if due_date.date() < today:

                    st.error(
                        "🔴 OVERDUE"
                    )

                elif due_date.date() == today:

                    st.warning(
                        "🟡 DUE TODAY"
                    )

                else:

                    st.info(
                        "🔵 UPCOMING"
                    )

        st.divider()

        capital_paid = st.checkbox(
            "Capital Returned",
            value=bool(
                df.at[idx,"CAPITAL_PAID"]
            ),
            key="capital_returned"
        )

        if st.button(
            "Save Payment Status",
            type="primary"
        ):

            for i in range(
                1,
                duration + 1
            ):

                paid_col = f"PAID_{i}"
                payment_col = f"PAYMENT_DATE_{i}"

                df.loc[
                    idx,
                    paid_col
                ] = bool(
                    payment_updates[i]
                )

                if payment_updates[i]:

                    existing_date = str(
                        df.at[
                            idx,
                            payment_col
                        ]
                    )

                    if (
                        existing_date == ""
                        or existing_date == "nan"
                    ):

                        df.loc[
                            idx,
                            payment_col
                        ] = datetime.now().strftime(
                            "%d-%m-%Y %H:%M:%S"
                        )

                else:

                    df.loc[
                        idx,
                        payment_col
                    ] = ""

            df.loc[
                idx,
                "CAPITAL_PAID"
            ] = capital_paid

            save_data(df)

            st.success(
                "Payment Status Updated Successfully"
            )

            st.rerun()


# =====================================
# REPORTS
# =====================================

elif menu == "Reports":

    today = pd.Timestamp.today().date()

    report = []

    for _, row in df.iterrows():

        duration = int(
            row["PLAN DURATION"]
        ) if pd.notna(
            row["PLAN DURATION"]
        ) else 0

        for i in range(
            1,
            duration + 1
        ):

            date_col = (
                f"PROFIT DATE {i}"
            )

            paid_col = (
                f"PAID_{i}"
            )

            if pd.notna(
                row[date_col]
            ):

                due_date = pd.to_datetime(
                    row[date_col]
                ).date()

                if not row[
                    paid_col
                ]:

                    status = (
                        "OVERDUE"
                        if due_date < today
                        else "UPCOMING"
                    )

                    report.append(
                        {
                            "Client":
                            row[
                                "CLIENT NAME"
                            ],
                            "Installment":
                            i,
                            "Due Date":
                            due_date,
                            "Amount":
                            row[
                                "PROFIT AMT"
                            ] /
                            duration,
                            "Status":
                            status
                        }
                    )

    report_df = pd.DataFrame(
        report
    )

    st.dataframe(
        report_df,
        use_container_width=True
    )

    csv = report_df.to_csv(
        index=False
    )

    st.download_button(
        "Download Report",
        csv,
        "Investor_Report.csv",
        "text/csv"
    )