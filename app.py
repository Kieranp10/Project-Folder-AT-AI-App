import streamlit as st
import pandas as pd
import plotly.express as px
from openai import OpenAI
import json
import os

# =====================================================
# CONFIG
# =====================================================

client = OpenAI(
    api_key=os.getenv("OPENAI_API_KEY")
)

st.set_page_config(
    page_title="Nursery Intelligence System",
    layout="wide"
)

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

    # FIX: ensure numeric totals work properly
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")
    df = df.dropna(subset=["Amount"])

    return df

df = load_data()

# =====================================================
# LOGIN SYSTEM
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
# AI SALES ASSISTANT (FIXED + ENHANCED)
# =====================================================

with tab1:

    st.header("🧠 Ask Your Data")

    question = st.text_input("Ask your question")

    if question:

        # =================================================
        # AI QUERY ENGINE (IMPROVED)
        # =================================================

        prompt = f"""
You are a nursery business intelligence parser.

Convert the question into structured JSON.

KNOWN LINES:
{KNOWN_LINES}

IMPORTANT RULES:
- A question may contain multiple filters
- Support line, crop, variety, client/store
- Support comparisons (X vs Y)
- Support time filters (month, year, last year, this year)

CLIENT RULES:
- store = Client Name
- client = Client Name
- shop = Client Name

METRICS:
- total
- top
- compare
- monthly_top
- top_client

GROUP BY:
- line
- crop
- variety
- client

COMPARE RULE:
If comparing items, put them in compare.items

OUTPUT ONLY VALID JSON:

{{
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
  "metric": null,
  "group_by": null,
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

        raw = response.choices[0].message.content

        try:
            q = json.loads(raw)
        except:
            st.error("AI returned invalid JSON")
            st.write(raw)
            st.stop()

        # =================================================
        # SMART FALLBACK (IMPORTANT)
        # =================================================

        if not q.get("metric"):

            q_lower = question.lower()

            if "compare" in q_lower:
                q["metric"] = "compare"
            elif "client" in q_lower or "store" in q_lower or "who bought" in q_lower:
                q["metric"] = "top_client"
            elif "month" in q_lower:
                q["metric"] = "monthly_top"
            elif "most" in q_lower or "best" in q_lower:
                q["metric"] = "top"
            elif "how many" in q_lower:
                q["metric"] = "total"
            else:
                q["metric"] = "total"

        df_temp = df.copy()

        # =================================================
        # FILTERING
        # =================================================

        for line in q.get("filters", {}).get("line", []):
            df_temp = df_temp[df_temp["Line"].str.contains(line, case=False, na=False)]

        for crop in q.get("filters", {}).get("crop", []):
            df_temp = df_temp[df_temp["Crop Name"].str.contains(crop, case=False, na=False)]

        for var in q.get("filters", {}).get("variety", []):
            df_temp = df_temp[df_temp["Variety"].str.contains(var, case=False, na=False)]

        for client_name in q.get("filters", {}).get("client", []):
            df_temp = df_temp[df_temp["Client Name"].str.contains(client_name, case=False, na=False)]

        # =================================================
        # DATE FILTERING
        # =================================================

        q_lower = question.lower()

        months = {
            "january": 1, "february": 2, "march": 3, "april": 4,
            "may": 5, "june": 6, "july": 7, "august": 8,
            "september": 9, "october": 10, "november": 11, "december": 12
        }

        for m, num in months.items():
            if m in q_lower:
                df_temp = df_temp[df_temp["Date"].dt.month == num]

        if "last year" in q_lower:
            latest = df["Date"].dt.year.max()
            df_temp = df_temp[df_temp["Date"].dt.year == latest - 1]

        if "this year" in q_lower:
            latest = df["Date"].dt.year.max()
            df_temp = df_temp[df_temp["Date"].dt.year == latest]

        if len(df_temp) == 0:
            st.warning("No matching data")
            st.stop()

        # =================================================
        # METRICS
        # =================================================

        metric = q.get("metric")

        # -------------------------
        # TOTAL
        # -------------------------
        if metric == "total":
            total = df_temp["Amount"].sum()
            st.success(f"Total Sales: {total:,}")

        # -------------------------
        # TOP
        # -------------------------
        elif metric == "top":

            group_map = {
                "line": "Line",
                "crop": "Crop Name",
                "variety": "Variety",
                "client": "Client Name"
            }

            group_col = group_map.get(q.get("group_by", "client"), "Client Name")

            result = df_temp.groupby(group_col)["Amount"].sum().sort_values(ascending=False).head(10)

            st.subheader(f"Top {group_col}")
            st.dataframe(result)

            fig = px.bar(x=result.index, y=result.values)
            st.plotly_chart(fig, use_container_width=True)

        # -------------------------
        # TOP CLIENT (FIXED)
        # -------------------------
        elif metric == "top_client":

            result = df_temp.groupby("Client Name")["Amount"].sum().sort_values(ascending=False).head(10)

            st.subheader("Top Clients / Stores")
            st.dataframe(result)

            fig = px.bar(x=result.index, y=result.values)
            st.plotly_chart(fig, use_container_width=True)

        # -------------------------
        # MONTHLY TOP
        # -------------------------
        elif metric == "monthly_top":

            result = df_temp.groupby(df_temp["Date"].dt.month)["Amount"].sum().sort_values(ascending=False)

            st.subheader("Best Months")
            st.dataframe(result)

            fig = px.bar(x=result.index, y=result.values)
            st.plotly_chart(fig)

        # -------------------------
        # COMPARE (FIXED + MULTI-COLUMN SMART)
        # -------------------------
        elif metric == "compare":

            items = q.get("compare", {}).get("items", [])

            results = {}

            def detect_column(item):

                item_lower = item.lower()

                for line in KNOWN_LINES:
                    if item_lower in line.lower():
                        return "Line"

                if df["Crop Name"].str.contains(item, case=False, na=False).any():
                    return "Crop Name"

                if df["Variety"].str.contains(item, case=False, na=False).any():
                    return "Variety"

                return "Client Name"

            for item in items:

                col = detect_column(item)

                temp = df_temp[df_temp[col].str.contains(item, case=False, na=False)]

                results[f"{item} ({col})"] = temp["Amount"].sum()

            st.subheader("Comparison Results")
            st.dataframe(results)

            fig = px.bar(x=list(results.keys()), y=list(results.values()))
            st.plotly_chart(fig, use_container_width=True)

        # -------------------------
        # DEFAULT
        # -------------------------
        else:
            st.success(f"Total Sales: {df_temp['Amount'].sum():,}")

        # =================================================
        # MATCHING DATA
        # =================================================

        st.subheader("Matching Data")
        st.dataframe(df_temp)

# =====================================================
# DASHBOARD (UNCHANGED)
# =====================================================

with tab2:

    st.header("📊 Business Dashboard")

    col1, col2, col3 = st.columns(3)

    col1.metric("Orders", len(df))
    col2.metric("Total Amount", f"{df['Amount'].sum():,}")
    col3.metric("Clients", df["Client Name"].nunique())

    st.subheader("🏆 Top Clients")

    top_clients = df.groupby("Client Name")["Amount"].sum().sort_values(ascending=False).head(10)
    st.dataframe(top_clients)

    fig1 = px.bar(x=top_clients.index, y=top_clients.values)
    st.plotly_chart(fig1, use_container_width=True)

    st.subheader("🌱 Top Crops")

    top_crops = df.groupby("Crop Name")["Amount"].sum().sort_values(ascending=False).head(10)
    st.dataframe(top_crops)

    fig2 = px.bar(x=top_crops.index, y=top_crops.values)
    st.plotly_chart(fig2, use_container_width=True)

    st.subheader("📦 Top Lines")

    top_lines = df.groupby("Line")["Amount"].sum().sort_values(ascending=False).head(10)
    st.dataframe(top_lines)

    fig3 = px.bar(x=top_lines.index, y=top_lines.values)
    st.plotly_chart(fig3, use_container_width=True)

    st.subheader("👨‍💼 Rep Performance")

    rep = df.groupby("User")["Amount"].sum().sort_values(ascending=False)
    st.dataframe(rep)

    fig4 = px.bar(x=rep.index, y=rep.values)
    st.plotly_chart(fig4, use_container_width=True)