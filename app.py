import streamlit as st
import pandas as pd
import plotly.express as px
from openai import OpenAI
from werkzeug.security import check_password_hash
from pathlib import Path
import os
import re

# =====================================================
# OPENAI
# =====================================================

client = OpenAI(
    api_key=os.getenv("OPENAI_API_KEY")
)

# =====================================================
# PAGE CONFIG
# =====================================================

st.set_page_config(
    page_title="Nursery Intelligence Copilot Enterprise",
    layout="wide"
)

# =====================================================
# LOGIN CONFIG
# =====================================================

BUILTIN_ADMIN_USERNAME = "admin"

BUILTIN_ADMIN_PASSWORD_HASH = (
    "scrypt:32768:8:1$cer4Rc98UopaCXLm$4d12e32c06be3f94fab0e22a80237c5d552f4b85ec0167b4c7f422050af7228161b83c239f4eae0fdc3dcdc17068492cc04dbcbc1a87bd565df57d776993abba"
)

STANDARD_COLUMNS = [
    "Date",
    "Client Name",
    "Crop Name",
    "Variety",
    "Line",
    "Rep",
    "Quantity",
    "Amount"
]


def empty_standard_frame():

    return pd.DataFrame({
        "Date": pd.Series(dtype="datetime64[ns]"),
        "Client Name": pd.Series(dtype="string"),
        "Crop Name": pd.Series(dtype="string"),
        "Variety": pd.Series(dtype="string"),
        "Line": pd.Series(dtype="string"),
        "Rep": pd.Series(dtype="string"),
        "Quantity": pd.Series(dtype="float"),
        "Amount": pd.Series(dtype="float")
    })


def ensure_standard_columns(df):

    for col in STANDARD_COLUMNS:

        if col not in df.columns:

            if col == "Date":
                df[col] = pd.NaT

            elif col in ["Quantity", "Amount"]:
                df[col] = 0

            else:
                df[col] = ""

    df["Date"] = pd.to_datetime(
        df["Date"],
        errors="coerce"
    )

    df["Quantity"] = pd.to_numeric(
        df["Quantity"],
        errors="coerce"
    ).fillna(0)

    df["Amount"] = pd.to_numeric(
        df["Amount"],
        errors="coerce"
    ).fillna(0)

    return df

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False


def user_password_hashes():

    try:

        u = st.secrets.get("users")

        if u is None:
            return {}

        if hasattr(u, "to_dict"):
            u = u.to_dict()

        elif not isinstance(u, dict):
            u = dict(u)

        return {
            str(k): str(v)
            for k, v in u.items()
        }

    except Exception:
        return {}


def verify_login(username, password):

    uname = (username or "").strip()

    if not uname:
        return False

    users = user_password_hashes()

    h = users.get(uname)

    if h and check_password_hash(h, password):
        return True

    if (
        uname.lower() == BUILTIN_ADMIN_USERNAME.lower()
        and check_password_hash(
            BUILTIN_ADMIN_PASSWORD_HASH,
            password
        )
    ):
        return True

    return False


if not st.session_state.authenticated:

    st.title("🌱 Nursery Intelligence Copilot")

    with st.form("login_form"):

        username = st.text_input("Username")

        password = st.text_input(
            "Password",
            type="password"
        )

        submitted = st.form_submit_button(
            "Sign In"
        )

        if submitted:

            if verify_login(
                username,
                password
            ):

                st.session_state.authenticated = True
                st.rerun()

            else:

                st.error(
                    "Invalid username or password"
                )

    st.stop()

# =====================================================
# KNOWN LINES
# =====================================================

KNOWN_LINES = [
    "SEEDLINGS",
    "15CM COLOUR POTS",
    "12CM COLOUR POT",
    "12CM HERB/VEG",
    "PETUNIA HYBRIDS",
    "SIMPLY BEAUTIFUL",
    "PLANT TO PLATE",
    "MID RANGE",
    "CALIBRACHOA",
    "20CM AERO BOWL",
    "14CM SUCCULENTS",
    "25CM HANGING BASKET",
    "25CM COLOUR POT",
    "ARGYRANTHEMUM",
    "15CM RANUNCULUS",
    "15CM ANGELONIA",
    "12CM PATIO RANGE",
    "12CM ORNAMENTAL CHILLI",
    "9CM SUCCULENTS",
    "DAHLIA POTS",
    "12CM PETUNIA HYBRIDS",
    "32CM MIX AND MINGLE",
    "12CM PRIMROSE/OBCONICA",
    "CHILLI POT",
    "LAWN PLUGS",
    "12CM GRASS POTS",
    "17CM GERANIUM"
]

# =====================================================
# APP DIRECTORY
# =====================================================

def app_dir():

    return Path(__file__).resolve().parent

# =====================================================
# COLUMN FINDER
# =====================================================

def first_column(df, candidates):

    cols = {
        str(c).strip().lower(): c
        for c in df.columns
    }

    for cand in candidates:

        key = cand.lower().strip()

        if key in cols:
            return cols[key]

    for c in df.columns:

        cl = str(c).strip().lower()

        for cand in candidates:

            if cand.lower() in cl:
                return c

    return None

# =====================================================
# LOAD EXCEL FILE
# =====================================================

def load_excel_file(path):

    candidate = app_dir() / path

    if not candidate.is_file():
        return empty_standard_frame()

    try:

        df = pd.read_excel(candidate)

    except Exception:

        return empty_standard_frame()

    if len(df) == 0:
        return empty_standard_frame()

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    # =================================================
    # DATE
    # =================================================

    if "Date" in df.columns:

        df["Date"] = pd.to_datetime(
            df["Date"],
            errors="coerce"
        )

    else:

        df["Date"] = pd.NaT

    # =================================================
    # CLIENT
    # =================================================

    c_client = first_column(
        df,
        [
            "Client Name",
            "Client name",
            "Store",
            "Customer"
        ]
    )

    if c_client:

        df["Client Name"] = (
            df[c_client]
            .astype(str)
            .str.strip()
        )

    else:

        df["Client Name"] = ""

    # =================================================
    # CROP
    # =================================================

    c_crop = first_column(
        df,
        [
            "Crop Name",
            "Crop name",
            "Crop"
        ]
    )

    if c_crop:

        df["Crop Name"] = (
            df[c_crop]
            .astype(str)
            .str.upper()
            .str.strip()
        )

    else:

        df["Crop Name"] = ""

    # =================================================
    # VARIETY
    # =================================================

    c_variety = first_column(
        df,
        [
            "Variety",
            "Colour",
            "Color"
        ]
    )

    if c_variety:

        df["Variety"] = (
            df[c_variety]
            .astype(str)
            .str.strip()
        )

    else:

        df["Variety"] = ""

    # =================================================
    # LINE
    # =================================================

    c_line = first_column(
        df,
        [
            "Lines",
            "Line"
        ]
    )

    if c_line:

        df["Line"] = (
            df[c_line]
            .astype(str)
            .str.strip()
        )

    else:

        df["Line"] = ""

    # =================================================
    # REP
    # =================================================

    c_rep = first_column(
        df,
        [
            "User",
            "Rep",
            "Sales Rep"
        ]
    )

    if c_rep:

        df["Rep"] = (
            df[c_rep]
            .astype(str)
            .str.strip()
        )

    else:

        df["Rep"] = ""

    # =================================================
    # QUANTITY + AMOUNT HANDLING
    # =================================================

    filename_lower = str(path).lower()

    # ================================================
    # ORDERS FILE
    # ================================================

    if "orders" in filename_lower:

        # Amount column in orders = quantity

        if "Amount" in df.columns:

            df["Quantity"] = pd.to_numeric(
                df["Amount"],
                errors="coerce"
            ).fillna(0)

        else:

            c_qty = first_column(
                df,
                [
                    "Quantity",
                    "Qty",
                    "Units"
                ]
            )

            if c_qty:

                df["Quantity"] = pd.to_numeric(
                    df[c_qty],
                    errors="coerce"
                ).fillna(0)

            else:

                df["Quantity"] = 0

        # orders don't use rand sales values
        df["Amount"] = 0

    # ================================================
    # SALES + RETURNS
    # ================================================

    else:

        c_qty = first_column(
            df,
            [
                "Quantity",
                "Qty",
                "Units"
            ]
        )

        if c_qty:

            df["Quantity"] = pd.to_numeric(
                df[c_qty],
                errors="coerce"
            ).fillna(0)

        else:

            df["Quantity"] = 0

        c_amount = first_column(
            df,
            [
                "Amount",
                "Sales",
                "Rand",
                "Total",
                "Net Amount"
            ]
        )

        if c_amount:

            df["Amount"] = pd.to_numeric(
                df[c_amount],
                errors="coerce"
            ).fillna(0)

        else:

            df["Amount"] = 0

    return ensure_standard_columns(df)

# =====================================================
# LOAD DATA
# =====================================================

@st.cache_data
def load_orders():

    return load_excel_file(
        "master_orders.xlsx"
    )


@st.cache_data
def load_sales():

    return load_excel_file(
        "master_sales.xlsx"
    )


@st.cache_data
def load_returns():

    return load_excel_file(
        "master_returns.xlsx"
    )


df_orders = ensure_standard_columns(
    load_orders()
)

df_sales = ensure_standard_columns(
    load_sales()
)

df_returns = ensure_standard_columns(
    load_returns()
)

# =====================================================
# INTENT DETECTION
# =====================================================

def detect_intent(question):

    ql = question.lower()

    intent = {

        "compare": False,
        "top": False,
        "metric": "quantity",
        "source": "orders",

        "line": None,
        "crop": None,
        "variety": None,
        "client": None,
        "rep": None,

        "year": None,
        "month": None
    }

    # =================================================
    # COMPARE
    # =================================================

    if any(
        x in ql
        for x in [
            "compare",
            "vs",
            "versus"
        ]
    ):

        intent["compare"] = True

    # =================================================
    # TOP
    # =================================================

    if any(
        x in ql
        for x in [
            "top",
            "best",
            "highest",
            "most"
        ]
    ):

        intent["top"] = True

    # =================================================
    # SOURCE
    # =================================================

    if any(
        x in ql
        for x in [
            "sales",
            "sold",
            "revenue"
        ]
    ):

        intent["source"] = "sales"

    if any(
        x in ql
        for x in [
            "returns",
            "returned",
            "refund"
        ]
    ):

        intent["source"] = "returns"

    # =================================================
    # METRIC
    # =================================================

    if any(
        x in ql
        for x in [
            "rand",
            "revenue",
            "amount",
            "value"
        ]
    ):

        intent["metric"] = "amount"

    # =================================================
    # LINE
    # =================================================

    for line in KNOWN_LINES:

        if line.lower() in ql:

            intent["line"] = line

    # =================================================
    # YEAR
    # =================================================

    current_year = pd.Timestamp.today().year

    if "last year" in ql:

        intent["year"] = current_year - 1

    if "this year" in ql:

        intent["year"] = current_year

    years = re.findall(
        r"\b(20\d{2})\b",
        ql
    )

    if years:

        intent["year"] = int(years[0])

    # =================================================
    # MONTH
    # =================================================

    months = {

        "january": 1,
        "february": 2,
        "march": 3,
        "april": 4,
        "may": 5,
        "june": 6,
        "july": 7,
        "august": 8,
        "september": 9,
        "october": 10,
        "november": 11,
        "december": 12
    }

    for m, num in months.items():

        if m in ql:

            intent["month"] = num

    return intent

# =====================================================
# FILTERS
# =====================================================

def apply_filters(df, intent):

    d = df.copy()

    if len(d) == 0:
        return d

    if intent["line"]:

        d = d[
            d["Line"]
            .astype(str)
            .str.contains(
                intent["line"],
                case=False,
                na=False
            )
        ]

    if intent["year"]:

        d = d[
            d["Date"].dt.year
            == intent["year"]
        ]

    if intent["month"]:

        d = d[
            d["Date"].dt.month
            == intent["month"]
        ]

    return d

# =====================================================
# TITLE
# =====================================================

st.title(
    "🌱 Nursery Intelligence Copilot Enterprise"
)

# =====================================================
# SIDEBAR
# =====================================================

with st.sidebar:

    st.subheader("📁 Data")

    st.write(
        f"Orders Rows: {len(df_orders)}"
    )

    st.write(
        f"Sales Rows: {len(df_sales)}"
    )

    st.write(
        f"Returns Rows: {len(df_returns)}"
    )

# =====================================================
# AI SEARCH
# =====================================================

question = st.text_input(
    "Ask anything about orders, sales, returns, crops, varieties, clients or reps"
)

if question:

    intent = detect_intent(question)

    # =================================================
    # ACTIVE DATASET
    # =================================================

    if intent["source"] == "sales":

        df_active = df_sales
        source_name = "Sales"

    elif intent["source"] == "returns":

        df_active = df_returns
        source_name = "Returns"

    else:

        df_active = df_orders
        source_name = "Orders"

    st.caption(
        f"Using Dataset: {source_name}"
    )

    # =================================================
    # FILTERED DATA
    # =================================================

    df_temp = apply_filters(
        df_active,
        intent
    )

    # =================================================
    # COMPARE LOGIC
    # =================================================

    if intent["compare"]:

        ql = question.lower()

        compare_items = []

        for line in KNOWN_LINES:

            if line.lower() in ql:
                compare_items.append(line)

        # =============================================
        # SALES VS RETURNS
        # =============================================

        if (
            "sales vs returns" in ql
            or (
                "sales" in ql
                and "returns" in ql
            )
        ):

            sales_filtered = apply_filters(
                df_sales,
                intent
            )

            returns_filtered = apply_filters(
                df_returns,
                intent
            )

            sales_amount = (
                sales_filtered["Amount"]
                .sum()
            )

            returns_amount = (
                returns_filtered["Amount"]
                .sum()
            )

            sales_qty = (
                sales_filtered["Quantity"]
                .sum()
            )

            returns_qty = (
                returns_filtered["Quantity"]
                .sum()
            )

            net_sales = (
                sales_amount
                - returns_amount
            )

            compare_df = pd.DataFrame({

                "Metric": [

                    "Sales Amount",
                    "Returns Amount",
                    "Net Sales",
                    "Sales Quantity",
                    "Returns Quantity"
                ],

                "Value": [

                    sales_amount,
                    returns_amount,
                    net_sales,
                    sales_qty,
                    returns_qty
                ]
            })

            st.subheader(
                "📊 Sales vs Returns"
            )

            st.dataframe(compare_df)

            fig = px.bar(
                compare_df,
                x="Metric",
                y="Value"
            )

            st.plotly_chart(
                fig,
                use_container_width=True
            )

        # =============================================
        # LINE COMPARISON
        # =============================================

        elif len(compare_items) >= 2:

            results = []

            for item in compare_items:

                temp = df_temp[
                    df_temp["Line"]
                    .astype(str)
                    .str.contains(
                        item,
                        case=False,
                        na=False
                    )
                ]

                qty_total = (
                    temp["Quantity"]
                    .sum()
                )

                amount_total = (
                    temp["Amount"]
                    .sum()
                )

                results.append({

                    "Line": item,
                    "Quantity": qty_total,
                    "Amount": amount_total
                })

            results_df = pd.DataFrame(results)

            st.subheader(
                "📊 Comparison Results"
            )

            st.dataframe(results_df)

            metric_col = "Quantity"

            if intent["metric"] == "amount":

                metric_col = "Amount"

            fig = px.bar(
                results_df,
                x="Line",
                y=metric_col
            )

            st.plotly_chart(
                fig,
                use_container_width=True
            )

    # =================================================
    # TOP LOGIC
    # =================================================

    elif intent["top"]:

        group_col = "Line"

        if "client" in question.lower():

            group_col = "Client Name"

        if "crop" in question.lower():

            group_col = "Crop Name"

        if "variety" in question.lower():

            group_col = "Variety"

        if "rep" in question.lower():

            group_col = "Rep"

        result = (

            df_temp
            .groupby(group_col)["Quantity"]
            .sum()
            .sort_values(
                ascending=False
            )
            .head(10)
        )

        st.subheader(
            f"🏆 Top {group_col}"
        )

        st.dataframe(result)

        fig = px.bar(
            x=result.index,
            y=result.values
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )

    # =================================================
    # STANDARD TOTALS
    # =================================================

    else:

        # =============================================
        # ORDERS
        # =============================================

        if intent["source"] == "orders":

            qty_total = (
                df_temp["Quantity"]
                .sum()
            )

            st.success(
                f"""
Total stock ordered:
{qty_total:,.0f} units
"""
            )

        # =============================================
        # SALES / RETURNS
        # =============================================

        else:

            sales_filtered = apply_filters(
                df_sales,
                intent
            )

            returns_filtered = apply_filters(
                df_returns,
                intent
            )

            sold_qty = (
                sales_filtered["Quantity"]
                .sum()
            )

            sold_amount = (
                sales_filtered["Amount"]
                .sum()
            )

            returned_qty = (
                returns_filtered["Quantity"]
                .sum()
            )

            returned_amount = (
                returns_filtered["Amount"]
                .sum()
            )

            net_sales = (
                sold_amount
                - returned_amount
            )

            st.success(
                f"""
Sold Quantity:
{sold_qty:,.0f}

Returned Quantity:
{returned_qty:,.0f}

Sales Revenue:
R{sold_amount:,.2f}

Returns Value:
R{returned_amount:,.2f}

Net Sales:
R{net_sales:,.2f}
"""
            )

    # =================================================
    # MATCHING DATA
    # =================================================

    st.subheader(
        "📋 Matching Data"
    )

    st.dataframe(
        df_temp,
        use_container_width=True
    )

# =====================================================
# DASHBOARD
# =====================================================

st.header("📊 Dashboard")

dashboard = st.selectbox(

    "Choose Dashboard",

    [
        "Orders",
        "Sales",
        "Returns"
    ]
)

# =====================================================
# ORDERS DASHBOARD
# =====================================================

if dashboard == "Orders":

    df_orders_dashboard = ensure_standard_columns(
        df_orders.copy()
    )

    col1, col2 = st.columns(2)

    col1.metric(
        "Total Orders",
        len(df_orders_dashboard)
    )

    col2.metric(
        "Total Ordered Qty",
        f"{df_orders_dashboard['Quantity'].sum():,.0f}"
    )

    st.subheader(
        "📦 Orders By Line"
    )

    if len(df_orders_dashboard) > 0:

        orders_lines = (

            df_orders_dashboard
            .groupby("Line")["Quantity"]
            .sum()
            .sort_values(
                ascending=False
            )
        )

        st.dataframe(orders_lines)

        fig = px.bar(
            x=orders_lines.index,
            y=orders_lines.values
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )

# =====================================================
# SALES DASHBOARD
# =====================================================

elif dashboard == "Sales":

    col1, col2 = st.columns(2)

    col1.metric(
        "Sales Records",
        len(df_sales)
    )

    col2.metric(
        "Sales Revenue",
        f"R{df_sales['Amount'].sum():,.2f}"
    )

    if len(df_sales) > 0:

        st.subheader(
            "💰 Sales By Line"
        )

        sales_lines = (

            df_sales
            .groupby("Line")["Amount"]
            .sum()
            .sort_values(
                ascending=False
            )
        )

        st.dataframe(sales_lines)

        fig = px.bar(
            x=sales_lines.index,
            y=sales_lines.values
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )

# =====================================================
# RETURNS DASHBOARD
# =====================================================

elif dashboard == "Returns":

    col1, col2 = st.columns(2)

    col1.metric(
        "Returns Records",
        len(df_returns)
    )

    col2.metric(
        "Returns Value",
        f"R{df_returns['Amount'].sum():,.2f}"
    )

    if len(df_returns) > 0:

        st.subheader(
            "🔄 Returns By Line"
        )

        returns_lines = (

            df_returns
            .groupby("Line")["Amount"]
            .sum()
            .sort_values(
                ascending=False
            )
        )

        st.dataframe(returns_lines)

        fig = px.bar(
            x=returns_lines.index,
            y=returns_lines.values
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )
