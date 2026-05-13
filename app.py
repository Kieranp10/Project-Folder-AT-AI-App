import streamlit as st
import pandas as pd
import plotly.express as px
import os
from openai import OpenAI
import json

# =====================================================
# CONFIG
# =====================================================

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

st.set_page_config(
    page_title="Nursery Intelligence Copilot v2",
    layout="wide"
)

# =====================================================
# DATA
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
# INTENT ENGINE (NO MORE AI FAILURE HERE)
# =====================================================

def detect_intent(q):

    ql = q.lower()

    intent = {
        "compare": False,
        "total": False,
        "line": None,
        "crop": None,
        "client": None,
        "year": None,
        "month": None
    }

    # compare
    if any(x in ql for x in ["compare", "vs", "versus"]):
        intent["compare"] = True

    # totals
    if "how many" in ql or "total" in ql:
        intent["total"] = True

    # line detection
    for line in KNOWN_LINES:
        if line.lower() in ql:
            intent["line"] = line
            break

    # crop detection (simple but reliable)
    if "petunia" in ql:
        intent["crop"] = "petunia"

    # client
    if "client" in ql or "store" in ql:
        intent["client"] = "client"

    # year
    for y in range(2000, 2100):
        if str(y) in ql:
            intent["year"] = int(y)

    if "last year" in ql:
        intent["year"] = df["Date"].dt.year.max() - 1

    if "this year" in ql:
        intent["year"] = df["Date"].dt.year.max()

    # month
    months = {
        "january":1,"february":2,"march":3,"april":4,"may":5,
        "june":6,"july":7,"august":8,"september":9,
        "october":10,"november":11,"december":12
    }

    for m, v in months.items():
        if m in ql:
            intent["month"] = v

    return intent

# =====================================================
# FILTER ENGINE (PURE + RELIABLE)
# =====================================================

def apply_filters(df, intent):

    d = df.copy()

    if intent["line"]:
        d = d[d["Line"].astype(str).str.contains(intent["line"], case=False, na=False)]

    if intent["crop"]:
        d = d[d["Crop Name"].astype(str).str.contains(intent["crop"], case=False, na=False)]

    if intent["year"]:
        d = d[d["Date"].dt.year == intent["year"]]

    if intent["month"]:
        d = d[d["Date"].dt.month == intent["month"]]

    return d

# =====================================================
# UI
# =====================================================

st.title("🌱 Nursery Intelligence Copilot v2")

question = st.text_input("Ask anything")

if question:

    intent = detect_intent(question)
    df_temp = apply_filters(df, intent)

    # fallback safety
    if len(df_temp) == 0:
        st.warning("No exact match — expanding search")
        df_temp = df.copy()

    # =================================================
    # COMPARE
    # =================================================

    if intent["compare"]:

        items = []

        for line in KNOWN_LINES:
            if line.lower() in question.lower():
                items.append(line)

        results = {}

        for item in items:
            temp = df_temp[df_temp["Line"].astype(str).str.contains(item, case=False, na=False)]
            results[item] = temp["Amount"].sum()

        st.subheader("Comparison Results")
        st.dataframe(results)
        st.plotly_chart(px.bar(x=list(results.keys()), y=list(results.values())))

    # =================================================
    # TOTAL
    # =================================================

    elif intent["total"]:

        total = df_temp["Amount"].sum()
        st.success(f"Total Sold: {total:,}")

    # =================================================
    # DEFAULT
    # =================================================

    else:

        st.success(f"Total Sold: {df_temp['Amount'].sum():,}")

    # =================================================
    # DATA VIEW
    # =================================================

    st.subheader("Matching Data")
    st.dataframe(df_temp)