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
# SIMPLE MEMORY (COPILOT ADDITION - SAFE)
# =====================================================

if "memory" not in st.session_state:
    st.session_state.memory = {}

# =====================================================
# KNOWN LINES (YOUR ORIGINAL - NOT CHANGED)
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
# NORMALIZE TEXT (YOUR ORIGINAL - FIXED ONLY SAFE BUGS)
# =====================================================

def normalize_text(text):
    return (
        str(text)
        .upper()
        .replace("-", " ")
        .strip()
    )

# =====================================================
# MATCH LINE (YOUR ORIGINAL LOGIC RESTORED)
# =====================================================

def match_line(user_text):

    user_normalized = normalize_text(user_text)

    for line in KNOWN_LINES:

        line_normalized = normalize_text(line)

        if line_normalized in user_normalized:
            return line

        if user_normalized in line_normalized:
            return line

    return None

# =====================================================
# LOAD DATA (UNCHANGED BUT SAFER)
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
# LOGIN SYSTEM (UNCHANGED)
# =====================================================

USERS = {
    "admin": "1234",
    "manager": "pass123",
    "viewer": "view123"
}

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:

    st.title("🔐 Management Login")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        if username in USERS and USERS[username] == password:
            st.session_state.logged_in = True
            st.rerun()
        else:
            st.error("Invalid login")

    st.stop()

# =====================================================
# MAIN APP
# =====================================================

st.title("🌱 Nursery Intelligence System")

tab1, tab2 = st.tabs(["🧠 AI Sales Assistant", "📊 Dashboard"])

# =====================================================
# AI COPILOT (YOUR ORIGINAL + FIXES ONLY)
# =====================================================

with tab1:

    st.header("🧠 Ask Your Data")

    question = st.text_input(
        "Ask things like:\n"
        "- How many seedlings did we sell?\n"
        "- Which crop sold best in Petunia Hybrids?\n"
        "- Which colour sold best in 15cm Colour Pots?\n"
        "- Who was our top performing client last year?"
    )

    if question:

        detected_line = None

        for line in KNOWN_LINES:
            if normalize_text(line) in normalize_text(question):
                detected_line = line
                break

        # =================================================
        # AI PROMPT (UNCHANGED STRUCTURE)
        # =================================================

        prompt = f"""
You are a nursery business intelligence AI.

Extract structured JSON.

KNOWN LINES:
{KNOWN_LINES}

RULES:
- line, crop, variety, client must be separated correctly
- store = client
- allow comparisons
- allow time filtering (last year, 2025, may, etc)

FIELDS:
- line
- crop
- variety
- client
- metric
- group_by

Return ONLY JSON.

QUESTION:
{question}
"""

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": "Return ONLY JSON."},
                {"role": "user", "content": prompt}
            ]
        )

        try:
            q = json.loads(response.choices[0].message.content)
        except:
            st.error(response.choices[0].message.content)
            st.stop()

        if detected_line:
            q["line"] = detected_line

        df_temp = df.copy()

        # =====================================================
        # FIXED DATE ENGINE (THIS FIXES YOUR MAIN ISSUE)
        # =====================================================

        q_lower = question.lower()
        df_temp["Date"] = pd.to_datetime(df_temp["Date"], errors="coerce")

        latest_year = df["Date"].dt.year.max()

        if "last year" in q_lower:
            df_temp = df_temp[df_temp["Date"].dt.year == latest_year - 1]

        if "this year" in q_lower:
            df_temp = df_temp[df_temp["Date"].dt.year == latest_year]

        for year in range(2000, 2100):
            if str(year) in q_lower:
                df_temp = df_temp[df_temp["Date"].dt.year == year]

        months = {
            "january":1,"february":2,"march":3,"april":4,"may":5,"june":6,
            "july":7,"august":8,"september":9,"october":10,"november":11,"december":12
        }

        for m, num in months.items():
            if m in q_lower:
                df_temp = df_temp[df_temp["Date"].dt.month == num]

        # =====================================================
        # ORIGINAL FILTERS (UNCHANGED)
        # =====================================================

        if q.get("line"):
            matched_line = match_line(q["line"])
            if matched_line:
                df_temp = df_temp[df_temp["Line"].astype(str).str.upper() == matched_line.upper()]

        if q.get("crop"):
            df_temp = df_temp[df_temp["Crop Name"].astype(str).str.lower().str.contains(q["crop"].lower(), na=False)]

        if q.get("variety"):
            df_temp = df_temp[df_temp["Variety"].astype(str).str.lower().str.contains(q["variety"].lower(), na=False)]

        if q.get("client"):
            df_temp = df_temp[df_temp["Client Name"].astype(str).str.lower().str.contains(q["client"].lower(), na=False)]

        if len(df_temp) == 0:
            st.warning("No matching records found.")
            st.stop()

        # =====================================================
        # TOTAL FIX (IMPORTANT BUG FIX)
        # =====================================================

        if q.get("metric") == "total":
            total = df_temp["Amount"].sum()
            st.success(f"Total Sold: {total:,}")

        # =====================================================
        # TOP FIX
        # =====================================================

        elif q.get("metric") == "top":

            group_col = "Client Name"

            result = df_temp.groupby(group_col)["Amount"].sum().sort_values(ascending=False).head(10)

            st.dataframe(result)

            fig = px.bar(x=result.index, y=result.values)
            st.plotly_chart(fig)

        # =====================================================
        # DEFAULT
        # =====================================================

        else:
            st.success(f"Total Sold: {df_temp['Amount'].sum():,}")

        st.subheader("Matching Data")
        st.dataframe(df_temp)

# =====================================================
# DASHBOARD (UNCHANGED)
# =====================================================

with tab2:

    st.header("📊 Business Dashboard")

    st.metric("Orders", len(df))
    st.metric("Total Amount", f"{df['Amount'].sum():,}")
    st.metric("Clients", df["Client Name"].nunique())

    st.subheader("Top Clients")
    st.dataframe(df.groupby("Client Name")["Amount"].sum().sort_values(ascending=False).head(10))

    st.subheader("Top Crops")
    st.dataframe(df.groupby("Crop Name")["Amount"].sum().sort_values(ascending=False).head(10))

    st.subheader("Top Lines")
    st.dataframe(df.groupby("Line")["Amount"].sum().sort_values(ascending=False).head(10))