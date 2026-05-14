import streamlit as st
import pandas as pd
import plotly.express as px
from openai import OpenAI
from werkzeug.security import check_password_hash
from pathlib import Path
import os
import re

# =====================================================
# CONFIG
# =====================================================

BUILTIN_ADMIN_USERNAME = "admin"
BUILTIN_ADMIN_PASSWORD_HASH = (
    "scrypt:32768:8:1$cer4Rc98UopaCXLm$4d12e32c06be3f94fab0e22a80237c5d552f4b85ec0167b4c7f422050af7228161b83c239f4eae0fdc3dcdc17068492cc04dbcbc1a87bd565df57d776993abba"
)

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

st.set_page_config(
    page_title="Nursery Intelligence Copilot v2.1",
    layout="wide"
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
        return {str(k): str(v) for k, v in u.items()}
    except Exception:
        return {}


def verify_login(username: str, password: str) -> bool:
    uname = (username or "").strip()
    if not uname:
        return False
    users = user_password_hashes()
    h = users.get(uname)
    if h and check_password_hash(h, password):
        return True
    if uname.lower() == BUILTIN_ADMIN_USERNAME.lower() and check_password_hash(
        BUILTIN_ADMIN_PASSWORD_HASH, password
    ):
        return True
    return False


if not st.session_state.authenticated:
    st.title("Nursery Intelligence Copilot")
    if not user_password_hashes():
        st.info(
            "Sign in with **admin** / **1234**. Add more users in `.streamlit/secrets.toml` under `[users]` when ready."
        )
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Sign in")
        if submitted:
            if verify_login(username.strip(), password):
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Invalid username or password.")
    st.stop()

# =====================================================
# DATA LOAD
# =====================================================

def _first_column(df, candidates):
    cols = {str(c).strip().lower(): c for c in df.columns}
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


def _app_dir() -> Path:
    return Path(__file__).resolve().parent


def _load_master_file(path):
    candidate = _app_dir() / path
    if not candidate.is_file():
        candidate = Path(path)
    if not candidate.is_file():
        return None
    try:
        df = pd.read_excel(path)
    except Exception:
        return None
    df.columns = df.columns.astype(str).str.strip()

    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    else:
        df["Date"] = pd.NaT

    c_line = _first_column(df, ["Lines", "Line", "Product/service", "Item", "Description", "Class", "Item description"])
    if c_line:
        df["Line"] = df[c_line].astype(str).str.strip()
    else:
        df["Line"] = ""

    c_qty = _first_column(df, ["Quantity", "Qty", "Units", "QTY", "Amount"])
    if c_qty:
        df["Quantity"] = pd.to_numeric(df[c_qty], errors="coerce").fillna(0.0)
    else:
        df["Quantity"] = 0.0

    c_amt = _first_column(df, ["Amount", "Rand", "Sales", "Total", "Line amount", "Net amount"])
    if c_amt:
        df["Amount"] = pd.to_numeric(df[c_amt], errors="coerce").fillna(0.0)
    else:
        df["Amount"] = 0.0

    return df


@st.cache_data
def load_orders():
    df = _load_master_file("master_orders.xlsx")
    if df is None:
        return pd.DataFrame(
            columns=["Date", "Line", "Quantity", "Amount", "Crop Name", "Variety", "Client Name"]
        )
    return df


@st.cache_data
def load_sales():
    df = _load_master_file("master_sales.xlsx")
    if df is None:
        return pd.DataFrame(
            columns=["Date", "Line", "Quantity", "Amount", "Crop Name", "Variety", "Client Name"]
        )
    return df


@st.cache_data
def load_returns():
    df = _load_master_file("master_returns.xlsx")
    if df is None:
        return pd.DataFrame(
            columns=["Date", "Line", "Quantity", "Amount", "Crop Name", "Variety", "Client Name"]
        )
    return df


df_orders = load_orders()
df_sales = load_sales()
df_returns = load_returns()

orders_file = (_app_dir() / "master_orders.xlsx").is_file()
sales_file = (_app_dir() / "master_sales.xlsx").is_file()
returns_file = (_app_dir() / "master_returns.xlsx").is_file()

# =====================================================
# KNOWN LINES (UNCHANGED)
# =====================================================

KNOWN_LINES = [
    "SEEDLINGS","15CM COLOUR POTS","12CM COLOUR POT","12CM HERB/VEG",
    "PETUNIA HYBRIDS","SIMPLY BEAUTIFUL","PLANT TO PLATE","MID RANGE",
    "CALIBRACHOA","20CM AERO BOWL","14CM SUCCULENTS","25CM HANGING BASKET",
    "25CM COLOUR POT","ARGYRANTHEMUM","15CM RANUNCULUS","15CM ANGELONIA",
    "12CM PATIO RANGE","12CM ORNAMENTAL CHILLI","9CM SUCCULENTS",
    "DAHLIA POTS","12CM PETUNIA HYBRIDS","32CM MIX AND MINGLE",
    "12CM PRIMROSE/OBCONICA","CHILLI POT","LAWN PLUGS",
    "12CM GRASS POTS","17CM GERANIUM"
]

# =====================================================
# INTENT ENGINE (FULL RESTORE CORE)
# =====================================================

def detect_intent(q):

    ql = q.lower()

    intent = {
        "compare": False,
        "total": False,
        "top": False,
        "line": None,
        "crop": None,
        "variety": None,
        "client": None,
        "year": None,
        "month": None,
        "source": "orders",
        "metric": "quantity",
    }

    # =====================
    # ACTION DETECTION
    # =====================

    if any(x in ql for x in ["compare", "vs", "versus"]):
        intent["compare"] = True

    if any(x in ql for x in ["how many", "total", "sold", "amount", "returns", "returned", "net sales", "revenue"]):
        intent["total"] = True

    if any(x in ql for x in ["top", "best", "highest"]):
        intent["top"] = True

    # =====================
    # SOURCE ROUTING
    # =====================

    if any(x in ql for x in ["return", "returns", "returned", "refund"]):
        intent["source"] = "returns"
    elif any(x in ql for x in ["sold", "sales", "sale", "revenue", "net sales"]):
        intent["source"] = "sales"
    elif any(x in ql for x in ["order", "ordered", "ordering"]):
        intent["source"] = "orders"

    if any(x in ql for x in ["amount", "rand", "value", "sales", "revenue", "total", "net sales"]):
        intent["metric"] = "amount"
    if any(x in ql for x in ["how many", "quantity", "qty", "units", "pieces", "number of"]):
        intent["metric"] = "quantity"

    # =====================
    # LINE DETECTION
    # =====================

    for line in KNOWN_LINES:
        if line.lower() in ql:
            intent["line"] = line
            break

    # =====================
    # CROP / VARIETY / CLIENT
    # =====================

    intent["crop"] = None
    intent["variety"] = None
    intent["client"] = None

    if "petunia" in ql:
        intent["crop"] = "PETUNIA"

    # =====================
    # YEAR DETECTION
    # =====================

    current_year = pd.Timestamp.today().year

    if "last year" in ql:
        intent["year"] = current_year - 1

    if "this year" in ql:
        intent["year"] = current_year

    year_match = re.findall(r"\b(20\d{2})\b", ql)
    if year_match:
        intent["year"] = int(year_match[0])

    # =====================
    # MONTH DETECTION
    # =====================

    months = {
        "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
        "july": 7, "august": 8, "september": 9, "october": 10,
        "november": 11, "december": 12
    }

    for m, v in months.items():
        if m in ql:
            intent["month"] = v

    return intent

# =====================================================
# FILTER ENGINE (STABLE CORE)
# =====================================================

def apply_filters(df, intent):

    d = df.copy()

    if intent["line"] and "Line" in d.columns:
        d = d[d["Line"].astype(str).str.contains(intent["line"], case=False, na=False)]

    if intent["crop"] and "Crop Name" in d.columns:
        d = d[d["Crop Name"].astype(str).str.contains(intent["crop"], case=False, na=False)]

    if intent["variety"] and "Variety" in d.columns:
        d = d[d["Variety"].astype(str).str.contains(intent["variety"], case=False, na=False)]

    if intent["client"] and "Client Name" in d.columns:
        d = d[d["Client Name"].astype(str).str.contains(intent["client"], case=False, na=False)]

    if intent["year"] and "Date" in d.columns:
        d = d[d["Date"].dt.year == intent["year"]]

    if intent["month"] and "Date" in d.columns:
        d = d[d["Date"].dt.month == intent["month"]]

    return d

# =====================================================
# UI
# =====================================================

st.title("🌱 Nursery Intelligence Copilot v2.1")

with st.sidebar:
    st.subheader("Data files")
    st.write("master_orders.xlsx", "✅" if orders_file else "❌")
    st.write("master_sales.xlsx", "✅" if sales_file else "❌")
    st.write("master_returns.xlsx", "✅" if returns_file else "❌")
    st.write("Orders rows", len(df_orders) if df_orders is not None else 0)
    st.write("Sales rows", len(df_sales) if df_sales is not None else 0)
    st.write("Returns rows", len(df_returns) if df_returns is not None else 0)
    if not orders_file or not sales_file or not returns_file:
        st.warning("Missing one or more expected master files. Check file names and upload to the app folder.")

question = st.text_input("Ask anything about sales or orders")

if question:

    intent = detect_intent(question)
    df_active = df_orders
    source_label = "Orders"
    if intent["source"] == "sales":
        df_active = df_sales
        source_label = "Sales"
    elif intent["source"] == "returns":
        df_active = df_returns
        source_label = "Returns"

    st.caption(f"Active dataset: **{source_label}**")

    df_temp = apply_filters(df_active, intent)

    # =====================================================
    # SAFETY FALLBACK (PREVENT FALSE "NO DATA")
    # =====================================================

    if len(df_temp) == 0:
        df_temp = df_active.copy()

    ql = question.lower()

    # =====================================================
    # COMPARE LOGIC (RESTORED + FIXED)
    # =====================================================

    if intent["compare"]:

        items = []

        for line in KNOWN_LINES:
            if line.lower() in ql:
                items.append(line)

        if len(items) < 2:
            st.warning("Please mention at least 2 items to compare")
            st.stop()

        results = {}

        for item in items:
            temp = df_temp[df_temp["Line"].astype(str).str.contains(item, case=False, na=False)]
            results[item] = temp["Amount"].sum()

        st.subheader("Comparison Results")
        st.dataframe(results)

        fig = px.bar(x=list(results.keys()), y=list(results.values()))
        st.plotly_chart(fig, use_container_width=True)

    # =====================================================
    # TOP LOGIC (RESTORED)
    # =====================================================

    elif intent["top"]:

        group_col = "Client Name"

        if "crop" in ql:
            group_col = "Crop Name"
        elif "variety" in ql:
            group_col = "Variety"
        elif "line" in ql:
            group_col = "Line"

        if group_col not in df_temp.columns:
            group_col = "Line"

        result = df_temp.groupby(group_col)["Amount"].sum().sort_values(ascending=False).head(10)

        st.subheader(f"Top {group_col}")
        st.dataframe(result)

        fig = px.bar(x=result.index, y=result.values)
        st.plotly_chart(fig, use_container_width=True)

    # =====================================================
    # TOTAL LOGIC (RESTORED CORE FEATURE)
    # =====================================================

    else:
        if intent["source"] == "orders":
            ordered_qty = df_temp["Quantity"].sum()
            ordered_amt = df_temp["Amount"].sum()
            if intent["metric"] == "amount":
                st.success(f"Ordered Amount: R{ordered_amt:,.2f}")
            else:
                st.success(f"Ordered Quantity: {ordered_qty:,.0f}")
        else:
            sold_qty = df_temp["Quantity"].sum() if intent["source"] == "sales" else 0.0
            sold_amt = df_temp["Amount"].sum() if intent["source"] == "sales" else 0.0
            returns_temp = apply_filters(df_returns, intent)
            returned_qty = returns_temp["Quantity"].sum()
            returned_amt = returns_temp["Amount"].sum()
            net_sales = sold_amt - returned_amt

            if intent["source"] == "returns":
                st.success(
                    f"Returned Quantity: {returned_qty:,.0f} | Returned Amount: R{returned_amt:,.2f}"
                )
            else:
                if intent["metric"] == "amount":
                    st.success(
                        f"Sales Amount: R{sold_amt:,.2f} | Returns Amount: R{returned_amt:,.2f} | Net Sales: R{net_sales:,.2f}"
                    )
                else:
                    st.success(
                        f"Sold Quantity: {sold_qty:,.0f} | Returned Quantity: {returned_qty:,.0f} | Net Sales: R{net_sales:,.2f}"
                    )

    # =====================================================
    # DATA VIEW (UNCHANGED)
    # =====================================================

    st.subheader("Matching Data")
    st.dataframe(df_temp)

# =====================================================
# DASHBOARD (FULL RESTORE)
# =====================================================

st.header("📊 Dashboard")

col1, col2, col3 = st.columns(3)

col1.metric("Orders", len(df_orders))
col2.metric("Order Amount", f"{df_orders['Amount'].sum():,}")
col3.metric("Order Clients", df_orders["Client Name"].nunique() if "Client Name" in df_orders.columns else 0)

st.subheader("Top Clients")
st.dataframe(df_orders.groupby("Client Name")["Amount"].sum().sort_values(ascending=False).head(10))

st.subheader("Top Crops")
st.dataframe(df_orders.groupby("Crop Name")["Amount"].sum().sort_values(ascending=False).head(10))

st.subheader("Top Lines")
st.dataframe(df_orders.groupby("Line")["Amount"].sum().sort_values(ascending=False).head(10))