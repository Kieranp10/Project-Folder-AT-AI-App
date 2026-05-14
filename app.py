import streamlit as st
import pandas as pd
import plotly.express as px
from openai import OpenAI
from werkzeug.security import check_password_hash
from pathlib import Path
import os
import re
import json

# =====================================================
# CONFIG
# =====================================================

client = OpenAI(
    api_key=os.getenv("OPENAI_API_KEY")
)

st.set_page_config(
    page_title="Nursery Intelligence Copilot v2.2",
    layout="wide"
)

# =====================================================
# LOGIN CONFIG
# =====================================================

BUILTIN_ADMIN_USERNAME = "admin"

BUILTIN_ADMIN_PASSWORD_HASH = (
    "scrypt:32768:8:1$cer4Rc98UopaCXLm$4d12e32c06be3f94fab0e22a80237c5d552f4b85ec0167b4c7f422050af7228161b83c239f4eae0fdc3dcdc17068492cc04dbcbc1a87bd565df57d776993abba"
)

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

# =====================================================
# LOGIN FUNCTIONS
# =====================================================

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
        uname.lower()
        == BUILTIN_ADMIN_USERNAME.lower()
        and check_password_hash(
            BUILTIN_ADMIN_PASSWORD_HASH,
            password
        )
    ):
        return True

    return False

# =====================================================
# LOGIN PAGE
# =====================================================

if not st.session_state.authenticated:

    st.title("🌱 Nursery Intelligence Copilot")

    st.info(
        "Login using admin / 1234"
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
# APP DIRECTORY
# =====================================================

def _app_dir():
    return Path(__file__).resolve().parent

# =====================================================
# FIND COLUMN
# =====================================================

def _first_column(df, candidates):

    cols = {
        str(c).strip().lower(): c
        for c in df.columns
    }

    for cand in candidates:

        key = cand.strip().lower()

        if key in cols:
            return cols[key]

    for c in df.columns:

        cl = str(c).strip().lower()

        for cand in candidates:

            if cand.strip().lower() in cl:
                return c

    return None

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
# LOAD DATA FILES
# =====================================================

def _load_master_file(path, dataset_type="orders"):

    candidate = _app_dir() / path

    if not candidate.is_file():
        candidate = Path(path)

    if not candidate.is_file():
        return None

    try:
        df = pd.read_excel(candidate)

    except Exception:
        return None

    df.columns = df.columns.astype(str).str.strip()

    # =================================================
    # DATE
    # =================================================

    date_col = _first_column(
        df,
        ["Date"]
    )

    if date_col:

        df["Date"] = pd.to_datetime(
            df[date_col],
            errors="coerce"
        )

    else:
        df["Date"] = pd.NaT

    # =================================================
    # LINE
    # =================================================

    line_col = _first_column(
        df,
        [
            "Lines",
            "Line",
            "Product/service",
            "Item",
            "Description"
        ]
    )

    if line_col:

        df["Line"] = (
            df[line_col]
            .astype(str)
            .str.strip()
        )

    else:
        df["Line"] = ""

    # =================================================
    # CLIENT
    # =================================================

    client_col = _first_column(
        df,
        [
            "Client name",
            "Client Name",
            "Store",
            "Customer"
        ]
    )

    if client_col:

        df["Client Name"] = (
            df[client_col]
            .astype(str)
            .str.strip()
        )

    else:
        df["Client Name"] = ""

    # =================================================
    # CROP
    # =================================================

    crop_col = _first_column(
        df,
        [
            "Crop name",
            "Crop Name",
            "Crop"
        ]
    )

    if crop_col:

        df["Crop Name"] = (
            df[crop_col]
            .astype(str)
            .str.strip()
            .str.upper()
        )

    else:
        df["Crop Name"] = ""

    # =================================================
    # VARIETY
    # =================================================

    variety_col = _first_column(
        df,
        [
            "Variety",
            "Colour",
            "Color"
        ]
    )

    if variety_col:

        df["Variety"] = (
            df[variety_col]
            .astype(str)
            .str.strip()
        )

    else:
        df["Variety"] = ""

    # =================================================
    # REP / USER
    # =================================================

    rep_col = _first_column(
        df,
        [
            "User",
            "Rep",
            "Sales Rep"
        ]
    )

    if rep_col:

        df["Rep"] = (
            df[rep_col]
            .astype(str)
            .str.strip()
        )

    else:
        df["Rep"] = ""

    # =================================================
    # ORDERS DATA
    # =================================================

    if dataset_type == "orders":

        qty_col = _first_column(
            df,
            ["Amount"]
        )

        if qty_col:

            df["Quantity"] = pd.to_numeric(
                df[qty_col],
                errors="coerce"
            ).fillna(0)

        else:
            df["Quantity"] = 0

        df["Amount"] = 0

    # =================================================
    # SALES / RETURNS
    # =================================================

    else:

        qty_col = _first_column(
            df,
            [
                "Quantity",
                "Qty",
                "Units"
            ]
        )

        if qty_col:

            df["Quantity"] = pd.to_numeric(
                df[qty_col],
                errors="coerce"
            ).fillna(0)

        else:
            df["Quantity"] = 0

        amount_col = _first_column(
            df,
            [
                "Amount",
                "Sales",
                "Total",
                "Rand"
            ]
        )

        if amount_col:

            df["Amount"] = pd.to_numeric(
                df[amount_col],
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

    df = _load_master_file(
        "master_orders.xlsx",
        "orders"
    )

    if df is None:

        return pd.DataFrame(
            columns=[
                "Date",
                "Line",
                "Quantity",
                "Crop Name",
                "Variety",
                "Client Name",
                "Rep"
            ]
        )

    return df

@st.cache_data
def load_sales():

    df = _load_master_file(
        "master_sales.xlsx",
        "sales"
    )

    if df is None:

        return pd.DataFrame(
            columns=[
                "Date",
                "Line",
                "Quantity",
                "Amount",
                "Crop Name",
                "Variety",
                "Client Name",
                "Rep"
            ]
        )

    return df

@st.cache_data
def load_returns():

    df = _load_master_file(
        "master_returns.xlsx",
        "returns"
    )

    if df is None:

        return pd.DataFrame(
            columns=[
                "Date",
                "Line",
                "Quantity",
                "Amount",
                "Crop Name",
                "Variety",
                "Client Name",
                "Rep"
            ]
        )

    return df

df_orders = load_orders()
df_sales = load_sales()
df_returns = load_returns()

# =====================================================
# AI INTENT ENGINE
# =====================================================

def detect_intent(question):

    ql = question.lower()

    intent = {
        "compare": False,
        "top": False,
        "line": [],
        "crop": None,
        "variety": None,
        "client": None,
        "rep": None,
        "year": None,
        "month": None,
        "dataset": "orders"
    }

    # =================================================
    # DATASET
    # =================================================

    if any(
        x in ql
        for x in [
            "sales",
            "sold",
            "revenue",
            "rand"
        ]
    ):
        intent["dataset"] = "sales"

    if any(
        x in ql
        for x in [
            "returns",
            "returned",
            "refund"
        ]
    ):
        intent["dataset"] = "returns"

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
    # LINE DETECTION
    # =================================================

    for line in KNOWN_LINES:

        if line.lower() in ql:
            intent["line"].append(line)

    # =================================================
    # YEAR
    # =================================================

    current_year = pd.Timestamp.today().year

    if "last year" in ql:
        intent["year"] = current_year - 1

    if "this year" in ql:
        intent["year"] = current_year

    year_match = re.findall(
        r"\b(20\d{2})\b",
        ql
    )

    if year_match:
        intent["year"] = int(year_match[0])

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

    for m, v in months.items():

        if m in ql:
            intent["month"] = v

    return intent

# =====================================================
# FILTERS
# =====================================================

def apply_filters(df, intent):

    d = df.copy()

    if len(intent["line"]) > 0:

        pattern = "|".join(
            intent["line"]
        )

        d = d[
            d["Line"]
            .astype(str)
            .str.contains(
                pattern,
                case=False,
                na=False
            )
        ]

    if intent["crop"]:

        d = d[
            d["Crop Name"]
            .astype(str)
            .str.contains(
                intent["crop"],
                case=False,
                na=False
            )
        ]

    if intent["variety"]:

        d = d[
            d["Variety"]
            .astype(str)
            .str.contains(
                intent["variety"],
                case=False,
                na=False
            )
        ]

    if intent["client"]:

        d = d[
            d["Client Name"]
            .astype(str)
            .str.contains(
                intent["client"],
                case=False,
                na=False
            )
        ]

    if intent["rep"]:

        d = d[
            d["Rep"]
            .astype(str)
            .str.contains(
                intent["rep"],
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
# MAIN APP
# =====================================================

st.title(
    "🌱 Nursery Intelligence Copilot v2.2"
)

tab1, tab2 = st.tabs(
    [
        "🧠 AI Assistant",
        "📊 Dashboards"
    ]
)

# =====================================================
# AI TAB
# =====================================================

with tab1:

    question = st.text_input(
        "Ask anything about orders, sales, returns, crops, reps, varieties, clients or lines"
    )

    if question:

        intent = detect_intent(question)

        # =============================================
        # DATASET
        # =============================================

        if intent["dataset"] == "sales":

            df_active = df_sales
            dataset_name = "Sales"

        elif intent["dataset"] == "returns":

            df_active = df_returns
            dataset_name = "Returns"

        else:

            df_active = df_orders
            dataset_name = "Orders"

        st.caption(
            f"Using Dataset: {dataset_name}"
        )

        df_temp = apply_filters(
            df_active,
            intent
        )

        # =============================================
        # EMPTY CHECK
        # =============================================

        if len(df_temp) == 0:

            st.warning(
                "No matching data found"
            )

        else:

            # =========================================
            # COMPARE
            # =========================================

            if (
                intent["compare"]
                and len(intent["line"]) >= 2
            ):

                results = {}

                for line in intent["line"]:

                    temp = df_temp[
                        df_temp["Line"]
                        .astype(str)
                        .str.contains(
                            line,
                            case=False,
                            na=False
                        )
                    ]

                    if dataset_name == "Orders":

                        results[line] = (
                            temp["Quantity"]
                            .sum()
                        )

                    else:

                        results[line] = (
                            temp["Amount"]
                            .sum()
                        )

                st.subheader(
                    "Comparison Results"
                )

                compare_df = pd.DataFrame(
                    {
                        "Line": results.keys(),
                        "Value": results.values()
                    }
                )

                st.dataframe(compare_df)

                fig = px.bar(
                    compare_df,
                    x="Line",
                    y="Value"
                )

                st.plotly_chart(
                    fig,
                    use_container_width=True
                )

            # =========================================
            # TOP ANALYSIS
            # =========================================

            elif intent["top"]:

                group_col = "Line"

                ql = question.lower()

                if "client" in ql:
                    group_col = "Client Name"

                elif "crop" in ql:
                    group_col = "Crop Name"

                elif "variety" in ql:
                    group_col = "Variety"

                elif "rep" in ql:
                    group_col = "Rep"

                metric_col = (
                    "Quantity"
                    if dataset_name == "Orders"
                    else "Amount"
                )

                result = (
                    df_temp
                    .groupby(group_col)[metric_col]
                    .sum()
                    .sort_values(
                        ascending=False
                    )
                    .head(10)
                )

                st.subheader(
                    f"Top {group_col}"
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

            # =========================================
            # TOTALS
            # =========================================

            else:

                if dataset_name == "Orders":

                    total_qty = (
                        df_temp["Quantity"]
                        .sum()
                    )

                    ai_text = f"""
Question:
{question}

Results:
Total Ordered Quantity:
{total_qty:,.0f}

Dataset:
Orders
"""

                else:

                    total_sales = (
                        df_temp["Amount"]
                        .sum()
                    )

                    total_qty = (
                        df_temp["Quantity"]
                        .sum()
                    )

                    ai_text = f"""
Question:
{question}

Results:
Total Value:
R{total_sales:,.2f}

Total Quantity:
{total_qty:,.0f}

Dataset:
{dataset_name}
"""

                natural_response = client.chat.completions.create(
                    model="gpt-4o-mini",

                    messages=[
                        {
                            "role": "system",
                            "content": """
You are a professional nursery business intelligence manager.

Answer professionally.

Be concise and analytical.

Do not use greetings or sign-offs.
"""
                        },

                        {
                            "role": "user",
                            "content": ai_text
                        }
                    ]
                )

                st.subheader(
                    "Managerial Insight"
                )

                st.success(
                    natural_response
                    .choices[0]
                    .message
                    .content
                )

            # =========================================
            # DATA TABLE
            # =========================================

            st.subheader(
                "Matching Data"
            )

            st.dataframe(
                df_temp,
                use_container_width=True
            )

# =====================================================
# DASHBOARDS
# =====================================================

with tab2:

    dashboard_tab = st.selectbox(
        "Select Dashboard",
        [
            "Orders",
            "Sales",
            "Returns"
        ]
    )

    # =================================================
    # ORDERS
    # =================================================

    if dashboard_tab == "Orders":

        st.header(
            "📦 Orders Dashboard"
        )

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "Order Records",
            len(df_orders)
        )

        col2.metric(
            "Total Quantity Ordered",
            f"{df_orders['Quantity'].sum():,.0f}"
        )

        col3.metric(
            "Clients",
            df_orders[
                "Client Name"
            ].nunique()
        )

        st.subheader(
            "Top Ordered Lines"
        )

        top_lines = (
            df_orders
            .groupby("Line")["Quantity"]
            .sum()
            .sort_values(
                ascending=False
            )
            .head(10)
        )

        st.dataframe(top_lines)

        fig1 = px.bar(
            x=top_lines.index,
            y=top_lines.values
        )

        st.plotly_chart(
            fig1,
            use_container_width=True
        )

        st.subheader(
            "Top Ordered Crops"
        )

        top_crops = (
            df_orders
            .groupby("Crop Name")["Quantity"]
            .sum()
            .sort_values(
                ascending=False
            )
            .head(10)
        )

        st.dataframe(top_crops)

        st.subheader(
            "Top Ordering Clients"
        )

        top_clients = (
            df_orders
            .groupby("Client Name")["Quantity"]
            .sum()
            .sort_values(
                ascending=False
            )
            .head(10)
        )

        st.dataframe(top_clients)

    # =================================================
    # SALES
    # =================================================

    elif dashboard_tab == "Sales":

        st.header(
            "💰 Sales Dashboard"
        )

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "Sales Records",
            len(df_sales)
        )

        col2.metric(
            "Sales Value",
            f"R{df_sales['Amount'].sum():,.2f}"
        )

        col3.metric(
            "Quantity Sold",
            f"{df_sales['Quantity'].sum():,.0f}"
        )

        top_sales = (
            df_sales
            .groupby("Line")["Amount"]
            .sum()
            .sort_values(
                ascending=False
            )
            .head(10)
        )

        st.subheader(
            "Top Sales Lines"
        )

        st.dataframe(top_sales)

        fig2 = px.bar(
            x=top_sales.index,
            y=top_sales.values
        )

        st.plotly_chart(
            fig2,
            use_container_width=True
        )

    # =================================================
    # RETURNS
    # =================================================

    elif dashboard_tab == "Returns":

        st.header(
            "🔄 Returns Dashboard"
        )

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "Return Records",
            len(df_returns)
        )

        col2.metric(
            "Return Value",
            f"R{df_returns['Amount'].sum():,.2f}"
        )

        col3.metric(
            "Returned Quantity",
            f"{df_returns['Quantity'].sum():,.0f}"
        )

        top_returns = (
            df_returns
            .groupby("Line")["Amount"]
            .sum()
            .sort_values(
                ascending=False
            )
            .head(10)
        )

        st.subheader(
            "Top Returned Lines"
        )

        st.dataframe(top_returns)

        fig3 = px.bar(
            x=top_returns.index,
            y=top_returns.values
        )

        st.plotly_chart(
            fig3,
            use_container_width=True
        )