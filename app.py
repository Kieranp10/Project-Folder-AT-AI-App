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
    page_title="Nursery AI Copilot",
    layout="wide"
)

# =====================================================
# MEMORY SYSTEM (COPILOT)
# =====================================================

if "memory" not in st.session_state:
    st.session_state.memory = {}

if "last_query" not in st.session_state:
    st.session_state.last_query = None

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
# LOAD DATA
# =====================================================

@st.cache_data
def load_data():
    df = pd.read_excel("master_orders.xlsx")
    df.columns = df.columns.str.strip()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")
    df = df.dropna(subset=["Amount"])
    return df

df = load_data()

# =====================================================
# LOGIN SYSTEM
# =====================================================

USERS = {"admin": "1234",
         "Kieran":"Kr557"
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
            st.error("Wrong login")

    st.stop()

# =====================================================
# APP
# =====================================================

st.title("🧠 Nursery AI Copilot")

tab1, tab2 = st.tabs(["AI Copilot", "Dashboard"])

# =====================================================
# COPILOT TAB
# =====================================================

with tab1:

    st.header("Ask Your Nursery Data")

    question = st.text_input("Ask anything")

    if question:

        st.session_state.last_query = question

        # =================================================
        # AI PARSER (EXPANDED - NOT REMOVED ANY FEATURES)
        # =================================================

        prompt = f"""
You are a nursery AI copilot.

Convert question into structured JSON.

KNOWN LINES:
{list(df['Line'].unique())}

You MUST detect:
- line, crop, variety, client
- multiple filters allowed
- comparisons allowed
- time filters allowed

METRICS:
- total
- top
- compare
- monthly_top
- top_client

IMPORTANT:
- "store/client/shop" = Client Name
- allow multiple filters at once

Return ONLY JSON:

{{
  "metric": null,
  "filters": {{
    "line": [],
    "crop": [],
    "variety": [],
    "client": []
  }},
  "date": {{
    "year": null,
    "month": null
  }},
  "compare": {{
    "items": []
  }}
}}

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

        q = json.loads(response.choices[0].message.content)

        df_temp = df.copy()

        # =====================================================
        # SAFE DATE ENGINE (FULL VERSION - NO REMOVAL)
        # =====================================================

        q_lower = question.lower()
        df_temp["Date"] = pd.to_datetime(df_temp["Date"], errors="coerce")

        latest_year = df["Date"].dt.year.max()

        # YEAR LOGIC
        if "last year" in q_lower:
            df_temp = df_temp[df_temp["Date"].dt.year == latest_year - 1]

        if "this year" in q_lower:
            df_temp = df_temp[df_temp["Date"].dt.year == latest_year]

        for year in range(2000, 2100):
            if str(year) in q_lower:
                df_temp = df_temp[df_temp["Date"].dt.year == year]

        # MONTH LOGIC
        months = {
            "january":1,"february":2,"march":3,"april":4,"may":5,"june":6,
            "july":7,"august":8,"september":9,"october":10,"november":11,"december":12
        }

        for m, num in months.items():
            if m in q_lower:
                df_temp = df_temp[df_temp["Date"].dt.month == num]

        # WEEK LOGIC
        if "last 7 days" in q_lower:
            df_temp = df_temp[df_temp["Date"] >= df["Date"].max() - pd.Timedelta(days=7)]

        if "last 30 days" in q_lower:
            df_temp = df_temp[df_temp["Date"] >= df["Date"].max() - pd.Timedelta(days=30)]

        # =====================================================
        # FILTER ENGINE (UNCHANGED STRUCTURE)
        # =====================================================

        for f in q.get("filters", {}).get("line", []):
            df_temp = df_temp[df_temp["Line"].str.contains(f, case=False, na=False)]

        for f in q.get("filters", {}).get("crop", []):
            df_temp = df_temp[df_temp["Crop Name"].str.contains(f, case=False, na=False)]

        for f in q.get("filters", {}).get("variety", []):
            df_temp = df_temp[df_temp["Variety"].str.contains(f, case=False, na=False)]

        for f in q.get("filters", {}).get("client", []):
            df_temp = df_temp[df_temp["Client Name"].str.contains(f, case=False, na=False)]

        if len(df_temp) == 0:
            st.warning("No matching data found")
            st.stop()

        # =====================================================
        # METRICS (FULL ORIGINAL STYLE KEPT)
        # =====================================================

        metric = q.get("metric")

        # ---------------- TOTAL ----------------
        if metric == "total":
            st.success(f"Total Sales: {df_temp['Amount'].sum():,}")

        # ---------------- TOP ----------------
        elif metric == "top":
            result = df_temp.groupby("Client Name")["Amount"].sum().sort_values(ascending=False).head(10)

            st.subheader("Top Results")
            st.dataframe(result)

            fig = px.bar(x=result.index, y=result.values)
            st.plotly_chart(fig)

        # ---------------- TOP CLIENT ----------------
        elif metric == "top_client":
            result = df_temp.groupby("Client Name")["Amount"].sum().sort_values(ascending=False).head(10)

            st.subheader("Top Clients")
            st.dataframe(result)

            fig = px.bar(x=result.index, y=result.values)
            st.plotly_chart(fig)

        # ---------------- MONTHLY ----------------
        elif metric == "monthly_top":
            result = df_temp.groupby(df_temp["Date"].dt.month)["Amount"].sum()

            st.subheader("Monthly Performance")
            st.dataframe(result)

            fig = px.bar(x=result.index, y=result.values)
            st.plotly_chart(fig)

        # ---------------- COMPARE (FULL FIXED ENGINE) ----------------
        elif metric == "compare":

            items = q.get("compare", {}).get("items", [])

            results = {}

            for item in items:

                mask = (
                    df_temp["Line"].str.contains(item, case=False, na=False) |
                    df_temp["Crop Name"].str.contains(item, case=False, na=False) |
                    df_temp["Variety"].str.contains(item, case=False, na=False) |
                    df_temp["Client Name"].str.contains(item, case=False, na=False)
                )

                temp = df_temp[mask]

                results[item] = temp["Amount"].sum()

            st.subheader("Comparison Results")
            st.dataframe(results)

            fig = px.bar(x=list(results.keys()), y=list(results.values()))
            st.plotly_chart(fig)

        # ---------------- DEFAULT ----------------
        else:
            st.success(f"Total Sales: {df_temp['Amount'].sum():,}")

        # =====================================================
        # DATA VIEW
        # =====================================================

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