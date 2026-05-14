import streamlit as st
import pandas as pd
import plotly.express as px
from openai import OpenAI
from werkzeug.security import check_password_hash
from pathlib import Path
import os
import re

# =====================================================
# LOGIN CONFIG
# =====================================================

BUILTIN_ADMIN_USERNAME = "admin"
BUILTIN_ADMIN_PASSWORD_HASH = (
    "scrypt:32768:8:1$cer4Rc98UopaCXLm$4d12e32c06be3f94fab0e22a80237c5d552f4b85ec0167b4c7f422050af7228161b83c239f4eae0fdc3dcdc17068492cc04dbcbc1a87bd565df57d776993abba"
)

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
    """Find column in df that matches any of the candidates (flexible matching)"""
    cols = {str(c).strip().lower(): c for c in df.columns}
    
    # Exact match first
    for cand in candidates:
        key = cand.strip().lower()
        if key in cols:
            return cols[key]
    
    # Substring match
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
        df = pd.read_excel(candidate)
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

    c_qty = _first_column(df, ["Quantity", "Qty", "Units", "QTY"])
    if c_qty:
        df["Quantity"] = pd.to_numeric(df[c_qty], errors="coerce").fillna(0.0)
    else:
        df["Quantity"] = 0.0

    c_amt = _first_column(df, ["Amount", "Rand", "Sales", "Total", "Line amount", "Net amount"])
    if c_amt:
        df["Amount"] = pd.to_numeric(df[c_amt], errors="coerce").fillna(0.0)
    else:
        df["Amount"] = 0.0

    c_client = _first_column(df, ["Client name", "Client Name", "Store", "Customer", "Buyer"])
    if c_client:
        df["Client Name"] = df[c_client].astype(str).str.strip()
    else:
        df["Client Name"] = ""

    c_crop = _first_column(df, ["Crop name", "Crop Name", "Crop", "Product Type"])
    if c_crop:
        df["Crop Name"] = df[c_crop].astype(str).str.strip().str.upper()
    else:
        df["Crop Name"] = ""

    c_variety = _first_column(df, ["Variety", "Colour", "Color", "Type", "Shade"])
    if c_variety:
        df["Variety"] = df[c_variety].astype(str).str.strip()
    else:
        df["Variety"] = ""

    c_rep = _first_column(df, ["User", "Rep", "Representative", "Sales Rep", "Salesperson"])
    if c_rep:
        df["Rep"] = df[c_rep].astype(str).str.strip()
    else:
        df["Rep"] = ""

    return df


@st.cache_data
def load_orders():
    df = _load_master_file("master_orders.xlsx")
    if df is None:
        return pd.DataFrame(
            columns=["Date", "Line", "Quantity", "Amount", "Crop Name", "Variety", "Client Name", "Rep"]
        )
    return df


@st.cache_data
def load_sales():
    df = _load_master_file("master_sales.xlsx")
    if df is None:
        return pd.DataFrame(
            columns=["Date", "Line", "Quantity", "Amount", "Crop Name", "Variety", "Client Name", "Rep"]
        )
    return df


@st.cache_data
def load_returns():
    df = _load_master_file("master_returns.xlsx")
    if df is None:
        return pd.DataFrame(
            columns=["Date", "Line", "Quantity", "Amount", "Crop Name", "Variety", "Client Name", "Rep"]
        )
    return df


df_orders = load_orders()
df_sales = load_sales()
df_returns = load_returns()

orders_file = (_app_dir() / "master_orders.xlsx").is_file()
sales_file = (_app_dir() / "master_sales.xlsx").is_file()
returns_file = (_app_dir() / "master_returns.xlsx").is_file()

# =====================================================
# KNOWN LINES
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
# INTENT ENGINE
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
        "rep": None,
        "year": None,
        "month": None,
        "source": "orders",
        "metric": "quantity",
        "detailed": False,
    }

    # ACTION DETECTION
    if any(x in ql for x in ["compare", "vs", "versus", "ordered vs", "sold vs", "ordered and sold", "ordered and returned"]):
        intent["compare"] = True
    
    if any(x in ql for x in ["how many", "total", "sold", "amount", "returns", "returned", "net sales", "revenue", "ordered", "tell me"]):
        intent["total"] = True
    
    if any(x in ql for x in ["top", "best", "highest", "most", "leading"]):
        intent["top"] = True

    # DETAILED MANAGERIAL QUERY
    if any(x in ql for x in ["what", "how much", "tell me", "give me", "show me", "report", "analysis", "breakdown"]):
        intent["detailed"] = True

    # SOURCE ROUTING
    if any(x in ql for x in ["return", "returns", "returned", "refund"]):
        intent["source"] = "returns"
    elif any(x in ql for x in ["sold", "sales", "sale", "revenue", "net sales"]):
        intent["source"] = "sales"
    elif any(x in ql for x in ["order", "ordered", "ordering"]):
        intent["source"] = "orders"

    # If asking about "ordered vs sold" or comparison, mark as compare
    if "ordered" in ql and ("sold" in ql or "sold" in ql):
        intent["compare"] = True

    # METRIC DETECTION
    if any(x in ql for x in ["amount", "rand", "value", "sales", "revenue", "total", "net sales", "worth"]):
        intent["metric"] = "amount"
    if any(x in ql for x in ["how many", "quantity", "qty", "units", "pieces", "number of", "count"]):
        intent["metric"] = "quantity"

    # LINE DETECTION
    for line in KNOWN_LINES:
        if line.lower() in ql:
            intent["line"] = line
            break

    # CROP DETECTION (more flexible)
    crops = ["petunia", "calibrachoa", "argyranthemum", "dahlia", "ranunculus", "angelonia", "succulent", "chilli", "primrose", "geranium"]
    for crop in crops:
        if crop in ql:
            intent["crop"] = crop.upper()
            break

    # VARIETY/COLOUR DETECTION (look for common colours/varieties)
    varieties = ["pink", "red", "white", "purple", "yellow", "orange", "blue", "hybrid", "colour", "color", "shade", "variegated"]
    for var in varieties:
        if var in ql:
            intent["variety"] = var
            break

    # REP DETECTION (look for "rep named X" or "user X ordered" or similar)
    rep_patterns = [
        r"rep\s+(?:named\s+)?(\w+)",
        r"user\s+(\w+)",
        r"rep\s+(\w+)",
        r"(?:sales\s+)?rep\s+(\w+)",
    ]
    for pattern in rep_patterns:
        match = re.search(pattern, ql)
        if match:
            intent["rep"] = match.group(1).title()
            break

    # DATE DETECTION
    current_year = pd.Timestamp.today().year
    if "last year" in ql:
        intent["year"] = current_year - 1
    if "this year" in ql:
        intent["year"] = current_year
    year_match = re.findall(r"\b(20\d{2})\b", ql)
    if year_match:
        intent["year"] = int(year_match[0])

    # MONTH DETECTION
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
# FILTER ENGINE
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

    if intent["rep"] and "Rep" in d.columns:
        d = d[d["Rep"].astype(str).str.contains(intent["rep"], case=False, na=False)]

    if intent["year"] and "Date" in d.columns:
        d = d[d["Date"].dt.year == intent["year"]]

    if intent["month"] and "Date" in d.columns:
        d = d[d["Date"].dt.month == intent["month"]]

    return d


# =====================================================
# UI MAIN
# =====================================================

st.title("🌱 Nursery Intelligence Copilot v2.1")

with st.sidebar:
    st.subheader("📁 Data Files")
    st.write("✅ master_orders.xlsx" if orders_file else "❌ master_orders.xlsx")
    st.write("✅ master_sales.xlsx" if sales_file else "❌ master_sales.xlsx")
    st.write("✅ master_returns.xlsx" if returns_file else "❌ master_returns.xlsx")
    
    st.subheader("📊 Rows Loaded")
    st.write(f"Orders: {len(df_orders)}")
    st.write(f"Sales: {len(df_sales)}")
    st.write(f"Returns: {len(df_returns)}")

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

    st.caption(f"📌 Active dataset: **{source_label}**")

    df_temp = apply_filters(df_active, intent)

    if len(df_temp) == 0:
        df_temp = df_active.copy()

    ql = question.lower()

    # COMPARE LOGIC
    if intent["compare"]:
        items = []
        for line in KNOWN_LINES:
            if line.lower() in ql:
                items.append(line)

        if len(items) < 2:
            st.warning("Please mention at least 2 items to compare")
        else:
            results = {}
            for item in items:
                temp = df_temp[df_temp["Line"].astype(str).str.contains(item, case=False, na=False)]
                results[item] = temp["Amount"].sum()

            st.subheader("Comparison Results")
            st.dataframe(results)
            fig = px.bar(x=list(results.keys()), y=list(results.values()))
            st.plotly_chart(fig, use_container_width=True)

    # TOP LOGIC
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

        if len(df_temp) > 0:
            result = df_temp.groupby(group_col)["Amount"].sum().sort_values(ascending=False).head(10)
            st.subheader(f"Top {group_col}")
            st.dataframe(result)
            fig = px.bar(x=result.index, y=result.values)
            st.plotly_chart(fig, use_container_width=True)

    # TOTAL / SUMMARY LOGIC
    else:
        if intent["source"] == "orders":
            ordered_qty = df_temp["Quantity"].sum() if len(df_temp) > 0 else 0
            ordered_amt = df_temp["Amount"].sum() if len(df_temp) > 0 else 0
            if intent["metric"] == "amount":
                st.success(f"📦 Ordered Amount: R{ordered_amt:,.2f}")
            else:
                st.success(f"📦 Ordered Quantity: {ordered_qty:,.0f}")
        else:
            df_sales_filtered = apply_filters(df_sales, intent)
            df_returns_filtered = apply_filters(df_returns, intent)
            
            sold_qty = df_sales_filtered["Quantity"].sum() if len(df_sales_filtered) > 0 else 0.0
            sold_amt = df_sales_filtered["Amount"].sum() if len(df_sales_filtered) > 0 else 0.0
            returned_qty = df_returns_filtered["Quantity"].sum() if len(df_returns_filtered) > 0 else 0.0
            returned_amt = df_returns_filtered["Amount"].sum() if len(df_returns_filtered) > 0 else 0.0
            net_sales = sold_amt - returned_amt

            if intent["source"] == "returns":
                st.success(f"🔄 Returned Quantity: {returned_qty:,.0f} | Returned Amount: R{returned_amt:,.2f}")
            else:
                if intent["metric"] == "amount":
                    st.success(
                        f"💰 Sales Amount: R{sold_amt:,.2f} | Returns Amount: R{returned_amt:,.2f} | Net Sales: R{net_sales:,.2f}"
                    )
                else:
                    st.success(
                        f"📊 Sold Quantity: {sold_qty:,.0f} | Returned Quantity: {returned_qty:,.0f} | Net Sales: R{net_sales:,.2f}"
                    )

    # DATA VIEW
    st.subheader("📋 Matching Data")
    st.dataframe(df_temp, use_container_width=True)


# =====================================================
# DASHBOARDS
# =====================================================

st.header("📊 Dashboards")

dashboard_tab = st.selectbox("Select Dashboard", ["Orders", "Sales", "Returns"])

if dashboard_tab == "Orders":
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Orders", len(df_orders))
    col2.metric("Total Qty", f"{df_orders['Quantity'].sum():,.0f}")
    col3.metric("Total Amount", f"R{df_orders['Amount'].sum():,.2f}")

    if len(df_orders) > 0:
        st.subheader("Orders by Line")
        st.dataframe(df_orders.groupby("Line").agg({"Quantity": "sum", "Amount": "sum"}).sort_values("Amount", ascending=False))

        st.subheader("Orders by Crop")
        if "Crop Name" in df_orders.columns:
            st.dataframe(df_orders.groupby("Crop Name").agg({"Quantity": "sum", "Amount": "sum"}).sort_values("Amount", ascending=False))

        st.subheader("Orders by Client")
        if "Client Name" in df_orders.columns:
            st.dataframe(df_orders.groupby("Client Name").agg({"Quantity": "sum", "Amount": "sum"}).sort_values("Amount", ascending=False))

elif dashboard_tab == "Sales":
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Sales Records", len(df_sales))
    col2.metric("Total Qty Sold", f"{df_sales['Quantity'].sum():,.0f}")
    col3.metric("Total Sales Amount", f"R{df_sales['Amount'].sum():,.2f}")

    if len(df_sales) > 0:
        st.subheader("Sales by Line")
        st.dataframe(df_sales.groupby("Line").agg({"Quantity": "sum", "Amount": "sum"}).sort_values("Amount", ascending=False))

        st.subheader("Sales by Crop")
        if "Crop Name" in df_sales.columns:
            st.dataframe(df_sales.groupby("Crop Name").agg({"Quantity": "sum", "Amount": "sum"}).sort_values("Amount", ascending=False))

        st.subheader("Sales by Client")
        if "Client Name" in df_sales.columns:
            st.dataframe(df_sales.groupby("Client Name").agg({"Quantity": "sum", "Amount": "sum"}).sort_values("Amount", ascending=False))

elif dashboard_tab == "Returns":
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Returns Records", len(df_returns))
    col2.metric("Total Qty Returned", f"{df_returns['Quantity'].sum():,.0f}")
    col3.metric("Total Returns Amount", f"R{df_returns['Amount'].sum():,.2f}")

    if len(df_returns) > 0:
        st.subheader("Returns by Line")
        st.dataframe(df_returns.groupby("Line").agg({"Quantity": "sum", "Amount": "sum"}).sort_values("Amount", ascending=False))

        st.subheader("Returns by Crop")
        if "Crop Name" in df_returns.columns:
            st.dataframe(df_returns.groupby("Crop Name").agg({"Quantity": "sum", "Amount": "sum"}).sort_values("Amount", ascending=False))

        st.subheader("Returns by Client")
        if "Client Name" in df_returns.columns:
            st.dataframe(df_returns.groupby("Client Name").agg({"Quantity": "sum", "Amount": "sum"}).sort_values("Amount", ascending=False))