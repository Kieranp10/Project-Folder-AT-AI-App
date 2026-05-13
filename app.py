import streamlit as st
import pandas as pd
import plotly.express as px
from openai import OpenAI
import json
import os

# =====================================================
# CONFIG
# =====================================================

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

st.set_page_config(
    page_title="Nursery Intelligence System",
    layout="wide"
)

# =====================================================
# KNOWN LINES (UNCHANGED)
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
# TEXT NORMALISATION
# =====================================================

def normalize(text):
    return str(text).upper().replace("-", " ").strip()

# =====================================================
# MATCH LINE (YOUR ORIGINAL WORKING LOGIC RESTORED)
# =====================================================

def match_line(text):
    text = normalize(text)

    for line in KNOWN_LINES:
        ln = normalize(line)

        if ln in text:
            return line
        if text in ln:
            return line

    return None

# =====================================================
# LOAD DATA
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
# LOGIN
# =====================================================

USERS = {
    "admin": "1234",
    "manager": "pass123",
    "viewer": "view123"
}

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:

    st.title("🔐 Login")

    u = st.text_input("Username")
    p = st.text_input("Password", type="password")

    if st.button("Login"):
        if USERS.get(u) == p:
            st.session_state.logged_in = True
            st.rerun()
        else:
            st.error("Invalid login")

    st.stop()

# =====================================================
# APP
# =====================================================

st.title("🌱 Nursery Intelligence System")

tab1, tab2 = st.tabs(["AI Assistant", "Dashboard"])

# =====================================================
# AI TAB (NOW ONLY HELPS, DOES NOT CONTROL LOGIC)
# =====================================================

with tab1:

    st.header("Ask Your Data")

    question = st.text_input("Ask anything")

    if question:

        df_temp = df.copy()
        q_lower = question.lower()

        # =================================================
        # 1. LINE FILTER (ROBUST)
        # =================================================

        detected_line = None
        for line in KNOWN_LINES:
            if line.lower() in q_lower:
                detected_line = line
                break

        if detected_line:
            df_temp = df_temp[df_temp["Line"].astype(str).str.upper() == detected_line.upper()]

        # =================================================
        # 2. CROP / VARIETY / CLIENT (SIMPLE + RELIABLE)
        # =================================================

        words = question.lower()

        if "petunia" in words:
            df_temp = df_temp[df_temp["Crop Name"].str.contains("petunia", case=False, na=False)]

        if "variety" in words:
            # optional future expansion
            pass

        if "client" in words or "store" in words:
            pass

        # =================================================
        # 3. DATE ENGINE (FIXED - THIS WAS YOUR MAIN ISSUE)
        # =================================================

        df_temp["Date"] = pd.to_datetime(df_temp["Date"], errors="coerce")

        latest_year = df["Date"].dt.year.max()

        # YEAR
        if "last year" in q_lower:
            df_temp = df_temp[df_temp["Date"].dt.year == latest_year - 1]

        if "this year" in q_lower:
            df_temp = df_temp[df_temp["Date"].dt.year == latest_year]

        # SPECIFIC YEAR
        for y in range(2000, 2100):
            if str(y) in q_lower:
                df_temp = df_temp[df_temp["Date"].dt.year == y]

        # MONTHS (CRITICAL FIX FOR YOUR ISSUE)
        months = {
            "january":1,"february":2,"march":3,"april":4,"may":5,"june":6,
            "july":7,"august":8,"september":9,"october":10,"november":11,"december":12
        }

        detected_month = None
        for m, num in months.items():
            if m in q_lower:
                df_temp = df_temp[df_temp["Date"].dt.month == num]
                detected_month = m

        # =================================================
        # 4. SMART FALLBACK (PREVENTS "NO DATA")
        # =================================================

        if len(df_temp) == 0:
            st.warning("No matching data found - loosening filters")

            df_temp = df.copy()

            if detected_line:
                df_temp = df_temp[df_temp["Line"].astype(str).str.contains(detected_line, case=False, na=False)]

        # =================================================
        # 5. LOGIC: TOTAL
        # =================================================

        if "how many" in q_lower or "total" in q_lower:

            total = df_temp["Amount"].sum()

            st.success(f"Total Sold: {total:,}")

        # =================================================
        # 6. LOGIC: COMPARE (FIXED PROPERLY)
        # =================================================

        elif "compare" in q_lower or "vs" in q_lower:

            items = []

            for line in KNOWN_LINES:
                if line.lower() in q_lower:
                    items.append(line)

            results = {}

            for item in items:
                temp = df_temp[df_temp["Line"].astype(str).str.contains(item, case=False, na=False)]
                results[item] = temp["Amount"].sum()

            st.subheader("Comparison")
            st.dataframe(results)
            st.plotly_chart(px.bar(x=list(results.keys()), y=list(results.values())))

        # =================================================
        # 7. DEFAULT OUTPUT
        # =================================================

        else:
            st.success(f"Total Sold: {df_temp['Amount'].sum():,}")

        # =================================================
        # DATA VIEW
        # =================================================

        st.subheader("Matching Data")
        st.dataframe(df_temp)

# =====================================================
# DASHBOARD (UNCHANGED)
# =====================================================

with tab2:

    st.header("📊 Dashboard")

    st.metric("Orders", len(df))
    st.metric("Total Sales", f"{df['Amount'].sum():,}")
    st.metric("Clients", df["Client Name"].nunique())

    st.subheader("Top Clients")
    st.dataframe(df.groupby("Client Name")["Amount"].sum().sort_values(ascending=False).head(10))

    st.subheader("Top Crops")
    st.dataframe(df.groupby("Crop Name")["Amount"].sum().sort_values(ascending=False).head(10))

    st.subheader("Top Lines")
    st.dataframe(df.groupby("Line")["Amount"].sum().sort_values(ascending=False).head(10))