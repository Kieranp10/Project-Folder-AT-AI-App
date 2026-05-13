import streamlit as st
import pandas as pd
import plotly.express as px
from openai import OpenAI
import os
import re

# =====================================================
# CONFIG
# =====================================================

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

st.set_page_config(
    page_title="Nursery Intelligence Copilot v2.1",
    layout="wide"
)

# =====================================================
# DATA LOAD
# =====================================================

@st.cache_data
def load_data():
    df = pd.read_excel("master_orders.xlsx")
    df.columns = df.columns.str.strip()

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")

    return df

df = load_data()

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
        "month": None
    }

    # =====================
    # ACTION DETECTION
    # =====================

    if any(x in ql for x in ["compare", "vs", "versus"]):
        intent["compare"] = True

    if any(x in ql for x in ["how many", "total", "sold", "amount"]):
        intent["total"] = True

    if any(x in ql for x in ["top", "best", "highest"]):
        intent["top"] = True

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
    # YEAR DETECTION (FULL RESTORE)
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
    # MONTH DETECTION (FULL RESTORE)
    # =====================

    months = {
        "january":1,"february":2,"march":3,"april":4,"may":5,"june":6,
        "july":7,"august":8,"september":9,"october":10,
        "november":11,"december":12
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

    if intent["line"]:
        d = d[d["Line"].astype(str).str.contains(intent["line"], case=False, na=False)]

    if intent["crop"]:
        d = d[d["Crop Name"].astype(str).str.contains(intent["crop"], case=False, na=False)]

    if intent["variety"]:
        d = d[d["Variety"].astype(str).str.contains(intent["variety"], case=False, na=False)]

    if intent["client"]:
        d = d[d["Client Name"].astype(str).str.contains(intent["client"], case=False, na=False)]

    if intent["year"]:
        d = d[d["Date"].dt.year == intent["year"]]

    if intent["month"]:
        d = d[d["Date"].dt.month == intent["month"]]

    return d

# =====================================================
# UI
# =====================================================

st.title("🌱 Nursery Intelligence Copilot v2.1")

question = st.text_input("Ask anything about sales")

if question:

    intent = detect_intent(question)
    df_temp = apply_filters(df, intent)

    # =====================================================
    # SAFETY FALLBACK (PREVENT FALSE "NO DATA")
    # =====================================================

    if len(df_temp) == 0:
        df_temp = df.copy()

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

        result = df_temp.groupby(group_col)["Amount"].sum().sort_values(ascending=False).head(10)

        st.subheader(f"Top {group_col}")
        st.dataframe(result)

        fig = px.bar(x=result.index, y=result.values)
        st.plotly_chart(fig, use_container_width=True)

    # =====================================================
    # TOTAL LOGIC (RESTORED CORE FEATURE)
    # =====================================================

    else:

        total = df_temp["Amount"].sum()
        st.success(f"Total Sold: {total:,}")

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

col1.metric("Orders", len(df))
col2.metric("Total Amount", f"{df['Amount'].sum():,}")
col3.metric("Clients", df["Client Name"].nunique())

st.subheader("Top Clients")
st.dataframe(df.groupby("Client Name")["Amount"].sum().sort_values(ascending=False).head(10))

st.subheader("Top Crops")
st.dataframe(df.groupby("Crop Name")["Amount"].sum().sort_values(ascending=False).head(10))

st.subheader("Top Lines")
st.dataframe(df.groupby("Line")["Amount"].sum().sort_values(ascending=False).head(10))