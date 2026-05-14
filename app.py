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
    page_title="Nursery Intelligence Copilot Enterprise v3.1",
    layout="wide"
)

# =====================================================
# LOGIN
# =====================================================

BUILTIN_ADMIN_USERNAME = "admin"

BUILTIN_ADMIN_PASSWORD_HASH = (
    "scrypt:32768:8:1$cer4Rc98UopaCXLm$4d12e32c06be3f94fab0e22a80237c5d552f4b85ec0167b4c7f422050af7228161b83c239f4eae0fdc3dcdc17068492cc04dbcbc1a87bd565df57d776993abba"
)

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


def verify_login(username: str, password: str):

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

    st.title("🌱 Nursery Intelligence Copilot Enterprise")

    st.info(
        "Default Login: admin / 1234"
    )

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
                username.strip(),
                password
            ):

                st.session_state.authenticated = True
                st.rerun()

            else:

                st.error(
                    "Invalid username or password."
                )

    st.stop()

# =====================================================
# APP PATH
# =====================================================

def app_dir():

    return Path(__file__).resolve().parent

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

        cl = str(c).lower().strip()

        for cand in candidates:

            if cand.lower() in cl:
                return c

    return None

# =====================================================
# LOAD EXCEL FILE
# =====================================================

def load_excel_file(filename):

    path = app_dir() / filename

    if not path.exists():

        return pd.DataFrame()

    df = pd.read_excel(path)

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    # =================================================
    # DATE
    # =================================================

    c_date = first_column(
        df,
        ["Date"]
    )

    if c_date:

        df["Date"] = pd.to_datetime(
            df[c_date],
            errors="coerce"
        )

    else:

        df["Date"] = pd.NaT

    # =================================================
    # LINE
    # =================================================

    c_line = first_column(
        df,
        ["Lines", "Line"]
    )

    if c_line:

        df["Line"] = (
            df[c_line]
            .astype(str)
            .str.upper()
            .str.strip()
        )

    else:

        df["Line"] = ""

    # =================================================
    # CROP
    # =================================================

    c_crop = first_column(
        df,
        ["Crop Name", "Crop"]
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
        ["Variety", "Colour", "Color"]
    )

    if c_variety:

        df["Variety"] = (
            df[c_variety]
            .astype(str)
            .str.upper()
            .str.strip()
        )

    else:

        df["Variety"] = ""

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
    # QUANTITY
    # =================================================

    c_qty = first_column(
        df,
        [
            "Quantity",
            "Qty",
            "Units",
            "Amount"
        ]
    )

    if c_qty:

        df["Quantity"] = pd.to_numeric(
            df[c_qty],
            errors="coerce"
        ).fillna(0)

    else:

        df["Quantity"] = 0

    # =================================================
    # AMOUNT
    # =================================================

    c_amount = first_column(
        df,
        [
            "Sales",
            "Rand",
            "Total",
            "Value",
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

    return df

# =====================================================
# LOAD DATASETS
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

df_orders = load_orders()
df_sales = load_sales()
df_returns = load_returns()

# =====================================================
# SIDEBAR
# =====================================================

with st.sidebar:

    st.header("📁 Data Status")

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
# MAIN TITLE
# =====================================================

st.title(
    "🌱 Nursery Intelligence Copilot Enterprise v3.1"
)

# =====================================================
# AI SEARCH
# =====================================================

question = st.text_input(
    "Ask anything about orders, sales, returns, crops, varieties, reps or stores"
)

# =====================================================
# QUERY ENGINE
# =====================================================

if question:

    q_lower = question.lower()

    # =================================================
    # DATASET DETECTION
    # =================================================

    active_dataset = "orders"

    if any(
        x in q_lower
        for x in [
            "sales",
            "sold",
            "revenue",
            "returns",
            "returned",
            "net sales",
            "actual sales"
        ]
    ):

        active_dataset = "sales"

    # =================================================
    # BASE DATAFRAME
    # =================================================

    if active_dataset == "orders":

        df = df_orders.copy()

    else:

        df = df_sales.copy()

    # =================================================
    # LINE FILTER
    # =================================================

    found_lines = []

    for line in KNOWN_LINES:

        if line.lower() in q_lower:

            found_lines.append(line)

    if len(found_lines) == 1:

        df = df[
            df["Line"]
            .astype(str)
            .str.contains(
                found_lines[0],
                case=False,
                na=False
            )
        ]

    # =================================================
    # CROP FILTER
    # =================================================

    crop_keywords = [
        "petunia",
        "calibrachoa",
        "geranium",
        "angelonia",
        "dahlia",
        "primrose",
        "ranunculus",
        "succulent",
        "chilli"
    ]

    for crop in crop_keywords:

        if crop in q_lower:

            df = df[
                df["Crop Name"]
                .astype(str)
                .str.contains(
                    crop,
                    case=False,
                    na=False
                )
            ]

    # =================================================
    # VARIETY FILTER
    # =================================================

    variety_keywords = [
        "pink",
        "red",
        "white",
        "purple",
        "yellow",
        "orange",
        "blue",
        "eagle"
    ]

    for var in variety_keywords:

        if var in q_lower:

            df = df[
                df["Variety"]
                .astype(str)
                .str.contains(
                    var,
                    case=False,
                    na=False
                )
            ]

    # =================================================
    # YEAR FILTER
    # =================================================

    current_year = pd.Timestamp.today().year

    if "last year" in q_lower:

        df = df[
            df["Date"].dt.year
            == current_year - 1
        ]

    if "this year" in q_lower:

        df = df[
            df["Date"].dt.year
            == current_year
        ]

    year_match = re.findall(
        r"\b(20\d{2})\b",
        q_lower
    )

    if year_match:

        df = df[
            df["Date"].dt.year
            == int(year_match[0])
        ]

    # =================================================
    # MONTH FILTER
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

    for month_name, month_num in months.items():

        if month_name in q_lower:

            df = df[
                df["Date"].dt.month
                == month_num
            ]

    # =================================================
    # COMPARE
    # =================================================

    if (
        "compare" in q_lower
        or "vs" in q_lower
        or "versus" in q_lower
    ):

        if len(found_lines) < 2:

            st.warning(
                "Please mention at least 2 lines to compare"
            )

        else:

            results = {}

            for line in found_lines:

                temp = df_orders.copy()

                temp = temp[
                    temp["Line"]
                    .astype(str)
                    .str.contains(
                        line,
                        case=False,
                        na=False
                    )
                ]

                results[line] = temp[
                    "Quantity"
                ].sum()

            result_df = pd.DataFrame({
                "Line": results.keys(),
                "Quantity Ordered": results.values()
            })

            st.subheader(
                "📊 Comparison Results"
            )

            st.dataframe(result_df)

            fig = px.bar(
                result_df,
                x="Line",
                y="Quantity Ordered"
            )

            st.plotly_chart(
                fig,
                use_container_width=True
            )

    # =================================================
    # TOP RESULTS
    # =================================================

    elif any(
        x in q_lower
        for x in [
            "top",
            "best",
            "highest",
            "most"
        ]
    ):

        group_col = "Client Name"

        if "crop" in q_lower:
            group_col = "Crop Name"

        elif "variety" in q_lower:
            group_col = "Variety"

        elif "line" in q_lower:
            group_col = "Line"

        elif "rep" in q_lower:
            group_col = "Rep"

        metric_col = (
            "Quantity"
            if active_dataset == "orders"
            else "Amount"
        )

        result = (
            df.groupby(group_col)[metric_col]
            .sum()
            .sort_values(ascending=False)
            .head(10)
        )

        top_item = result.index[0]
        top_value = result.iloc[0]

        if active_dataset == "orders":

            result_text = f"""
Top {group_col}: {top_item}

Total Quantity Ordered: {top_value:,.0f}
"""

        else:

            result_text = f"""
Top {group_col}: {top_item}

Sales Value: R{top_value:,.2f}
"""

        ai_prompt = f"""
You are a senior nursery business analyst.

Question:
{question}

Results:
{result_text}

Write a concise professional management insight.
"""

        ai_response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "user",
                    "content": ai_prompt
                }
            ]
        )

        st.success(
            ai_response.choices[0].message.content
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
    # STANDARD RESULTS
    # =================================================

    else:

        if active_dataset == "orders":

            total_qty = df["Quantity"].sum()

            result_text = f"""
Total Quantity Ordered:
{total_qty:,.0f}
"""

        else:

            sold_amount = df["Amount"].sum()

            returns_filtered = df_returns.copy()

            returned_amount = (
                returns_filtered["Amount"]
                .sum()
            )

            net_sales = (
                sold_amount
                - returned_amount
            )

            result_text = f"""
Sales Amount:
R{sold_amount:,.2f}

Returns Amount:
R{returned_amount:,.2f}

Net Sales:
R{net_sales:,.2f}
"""

        ai_prompt = f"""
You are a senior nursery business intelligence manager.

Answer professionally and naturally.

Question:
{question}

Results:
{result_text}
"""

        ai_response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "user",
                    "content": ai_prompt
                }
            ]
        )

        st.subheader("🧠 AI Insight")

        st.success(
            ai_response.choices[0].message.content
        )

        st.subheader("📋 Matching Records")

        st.dataframe(
            df,
            use_container_width=True
        )

# =====================================================
# DASHBOARDS
# =====================================================

st.header("📊 Business Dashboards")

tab1, tab2, tab3 = st.tabs(
    [
        "📦 Orders",
        "💰 Sales",
        "🔄 Returns"
    ]
)

# =====================================================
# ORDERS DASHBOARD
# =====================================================

with tab1:

    col1, col2 = st.columns(2)

    col1.metric(
        "Total Orders",
        len(df_orders)
    )

    col2.metric(
        "Total Quantity Ordered",
        f"{df_orders['Quantity'].sum():,.0f}"
    )

    st.subheader(
        "📦 Quantity Ordered By Line"
    )

    orders_line = (
        df_orders.groupby("Line")[
            "Quantity"
        ]
        .sum()
        .sort_values(
            ascending=False
        )
    )

    st.dataframe(orders_line)

    fig1 = px.bar(
        x=orders_line.index,
        y=orders_line.values
    )

    st.plotly_chart(
        fig1,
        use_container_width=True
    )

    st.subheader(
        "🌱 Quantity Ordered By Crop"
    )

    orders_crop = (
        df_orders.groupby("Crop Name")[
            "Quantity"
        ]
        .sum()
        .sort_values(
            ascending=False
        )
    )

    st.dataframe(orders_crop)

    fig2 = px.bar(
        x=orders_crop.index,
        y=orders_crop.values
    )

    st.plotly_chart(
        fig2,
        use_container_width=True
    )

    st.subheader(
        "🏪 Quantity Ordered By Store"
    )

    orders_client = (
        df_orders.groupby("Client Name")[
            "Quantity"
        ]
        .sum()
        .sort_values(
            ascending=False
        )
    )

    st.dataframe(orders_client)

# =====================================================
# SALES DASHBOARD
# =====================================================

with tab2:

    sold = df_sales["Amount"].sum()

    returned = df_returns["Amount"].sum()

    net = sold - returned

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "Sales",
        f"R{sold:,.2f}"
    )

    col2.metric(
        "Returns",
        f"R{returned:,.2f}"
    )

    col3.metric(
        "Net Sales",
        f"R{net:,.2f}"
    )

    sales_line = (
        df_sales.groupby("Line")[
            "Amount"
        ]
        .sum()
        .sort_values(
            ascending=False
        )
    )

    st.subheader(
        "💰 Sales By Line"
    )

    st.dataframe(sales_line)

# =====================================================
# RETURNS DASHBOARD
# =====================================================

with tab3:

    returns_line = (
        df_returns.groupby("Line")[
            "Amount"
        ]
        .sum()
        .sort_values(
            ascending=False
        )
    )

    st.subheader(
        "🔄 Returns By Line"
    )

    st.dataframe(returns_line)