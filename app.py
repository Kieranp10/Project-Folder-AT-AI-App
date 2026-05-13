import streamlit as st
import pandas as pd
import plotly.express as px
from openai import OpenAI
import json

# =====================================================
# CONFIG
# =====================================================
client = OpenAI(api_key="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")

st.set_page_config(
    page_title="Nursery Intelligence System",
    layout="wide"
)

# =====================================================
# KNOWN LINES (FROM EXCEL)
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
# NORMALIZE TEXT
# =====================================================
def normalize_text(text):

    return (
        str(text)
        .upper()
        .replace("S", "")
        .replace("-", " ")
        .strip()
    )

# =====================================================
# MATCH LINE
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
# LOAD DATA
# =====================================================
@st.cache_data
def load_data():

    df = pd.read_excel("master_orders.xlsx")

    df.columns = df.columns.str.strip()

    df['Date'] = pd.to_datetime(
        df['Date'],
        errors='coerce'
    )

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

    password = st.text_input(
        "Password",
        type="password"
    )

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

tab1, tab2 = st.tabs(
    [
        "🧠 AI Sales Assistant",
        "📊 Dashboard"
    ]
)

# =====================================================
# =====================================================
# AI SALES ASSISTANT
# =====================================================
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

        # =================================================
        # DETECT LINE
        # =================================================
        detected_line = None

        for line in KNOWN_LINES:

            if normalize_text(line) in normalize_text(question):
                detected_line = line
                break

        # =================================================
        # AI PROMPT
        # =================================================
        prompt = f"""
You are a nursery business intelligence AI.

The user is asking questions about nursery sales data.

IMPORTANT:
These are official Line names:
{KNOWN_LINES}

RULES:

- If a line is mentioned, extract it exactly

- Crop Name is separate from Line

- Variety is separate from Crop Name

- Never confuse line with crop

- "store" means client

- "stores" means client

- "customer" means client

- "customers" means client

- "top performing" means metric = "top"

- "best performing" means metric = "top"

- "highest selling" means metric = "top"

- If user asks:
    "who was our top client"
    → metric = "top"
    → group_by = "client"

- If user asks:
    "top store"
    → metric = "top"
    → group_by = "client"

- If user asks:
    "top crop"
    → metric = "top"
    → group_by = "crop"

- If user asks:
    "top variety"
    → metric = "top"
    → group_by = "variety"

- If user asks:
    "how many"
    → metric = "total"

FIELDS:
- line
- crop
- variety
- client
- metric
- group_by

metric options:
- total
- top

group_by options:
- line
- crop
- variety
- client

If not mentioned use null.

Return ONLY valid JSON.
No markdown.
No explanations.

QUESTION:
{question}
"""

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            response_format={"type": "json_object"},
            messages=[
                {
                    "role": "system",
                    "content": """
You are a structured nursery ERP AI.

You ONLY return valid JSON.
No explanations.
No markdown.
"""
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ]
        )

        try:

            q = json.loads(
                response.choices[0].message.content
            )

        except Exception as e:

            st.error(f"AI Error: {e}")

            st.write(
                response.choices[0].message.content
            )

            st.stop()

        # =================================================
        # FORCE DETECTED LINE
        # =================================================
        if detected_line:
            q["line"] = detected_line

        # =================================================
        # FILTER DATA
        # =================================================
        df_temp = df.copy()

        # LINE
        if q.get("line"):

            matched_line = match_line(q["line"])

            if matched_line:

                df_temp = df_temp[
                    df_temp["Line"]
                    .astype(str)
                    .str.upper()
                    == matched_line.upper()
                ]

        # CROP
        if q.get("crop"):

            df_temp = df_temp[
                df_temp["Crop Name"]
                .astype(str)
                .str.lower()
                .str.contains(
                    q["crop"].lower(),
                    na=False
                )
            ]

        # VARIETY
        if q.get("variety"):

            df_temp = df_temp[
                df_temp["Variety"]
                .astype(str)
                .str.lower()
                .str.contains(
                    q["variety"].lower(),
                    na=False
                )
            ]

        # CLIENT
        if q.get("client"):

            df_temp = df_temp[
                df_temp["Client Name"]
                .astype(str)
                .str.lower()
                .str.contains(
                    q["client"].lower(),
                    na=False
                )
            ]

        # =================================================
        # DATE FILTERS
        # =================================================
        q_lower = question.lower()

        if "last year" in q_lower:

            latest_year = df['Date'].dt.year.max()

            df_temp = df_temp[
                df_temp['Date'].dt.year == latest_year - 1
            ]

        if "this year" in q_lower:

            latest_year = df['Date'].dt.year.max()

            df_temp = df_temp[
                df_temp['Date'].dt.year == latest_year
            ]

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

                df_temp = df_temp[
                    df_temp['Date'].dt.month == month_num
                ]

        # =================================================
        # EMPTY CHECK
        # =================================================
        if len(df_temp) == 0:

            st.warning("No matching records found.")

            st.stop()

        # =================================================
        # TOTAL QUESTIONS
        # =================================================
        if q.get("metric") == "total":

            total = df_temp["Amount"].sum()

            result_text = f"""
Total Sold: {total:,}
"""

            natural_prompt = f"""
You are a nursery sales employee speaking to management.

Answer naturally and professionally.

DO NOT:
- write emails
- use greetings
- use subject lines
- sign off
- say dear sir/madam

Be concise and business-like.

Question:
{question}

Results:
{result_text}
"""

            natural_response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {
                        "role": "system",
                        "content": """
You are a professional nursery business analyst.
"""
                    },
                    {
                        "role": "user",
                        "content": natural_prompt
                    }
                ]
            )

            st.subheader("AI Insight")

            st.success(
                natural_response.choices[0].message.content
            )

        # =================================================
        # TOP QUESTIONS
        # =================================================
        elif q.get("metric") == "top":

            group_by = q.get("group_by")

            if group_by == "line":
                group_col = "Line"

            elif group_by == "crop":
                group_col = "Crop Name"

            elif group_by == "variety":
                group_col = "Variety"

            elif group_by == "client":
                group_col = "Client Name"

            else:
                group_col = "Client Name"

            result = (
                df_temp.groupby(group_col)["Amount"]
                .sum()
                .sort_values(ascending=False)
                .head(20)
            )

            top_item = result.index[0]
            top_amount = result.iloc[0]

            result_text = f"""
Top {group_col}: {top_item}
Amount Sold: {top_amount:,}
"""

            natural_prompt = f"""
You are a nursery sales employee speaking to management.

Answer naturally and professionally.

DO NOT:
- write emails
- use greetings
- use subject lines
- sign off

Be concise and business-like.

Question:
{question}

Results:
{result_text}
"""

            natural_response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {
                        "role": "system",
                        "content": """
You are a professional nursery business analyst.
"""
                    },
                    {
                        "role": "user",
                        "content": natural_prompt
                    }
                ]
            )

            st.subheader("AI Insight")

            st.success(
                natural_response.choices[0].message.content
            )

            st.subheader(
                f"Top Performing {group_col}"
            )

            st.dataframe(result)

            fig = px.bar(
                x=result.index,
                y=result.values,
                labels={
                    "x": group_col,
                    "y": "Amount"
                }
            )

            st.plotly_chart(
                fig,
                use_container_width=True
            )

        # =================================================
        # DEFAULT
        # =================================================
        else:

            total = df_temp["Amount"].sum()

            st.subheader("Result")

            st.success(
                f"Total Sold: {total:,}"
            )

        # =================================================
        # MATCHING DATA
        # =================================================
        st.subheader("Matching Data")

        st.dataframe(
            df_temp[
                [
                    'Date',
                    'Client Name',
                    'Crop Name',
                    'Variety',
                    'Line',
                    'Amount'
                ]
            ]
        )

# =====================================================
# DASHBOARD
# =====================================================
with tab2:

    st.header("📊 Business Dashboard")

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "Orders",
        len(df)
    )

    col2.metric(
        "Total Amount",
        f"{df['Amount'].sum():,}"
    )

    col3.metric(
        "Clients",
        df['Client Name'].nunique()
    )

    # =================================================
    # TOP CLIENTS
    # =================================================
    st.subheader("🏆 Top Clients")

    top_clients = (
        df.groupby('Client Name')['Amount']
        .sum()
        .sort_values(ascending=False)
        .head(10)
    )

    st.dataframe(top_clients)

    fig1 = px.bar(
        x=top_clients.index,
        y=top_clients.values
    )

    st.plotly_chart(
        fig1,
        use_container_width=True
    )

    # =================================================
    # TOP CROPS
    # =================================================
    st.subheader("🌱 Top Crops")

    top_crops = (
        df.groupby('Crop Name')['Amount']
        .sum()
        .sort_values(ascending=False)
        .head(10)
    )

    st.dataframe(top_crops)

    fig2 = px.bar(
        x=top_crops.index,
        y=top_crops.values
    )

    st.plotly_chart(
        fig2,
        use_container_width=True
    )

    # =================================================
    # TOP LINES
    # =================================================
    st.subheader("📦 Top Lines")

    top_lines = (
        df.groupby('Line')['Amount']
        .sum()
        .sort_values(ascending=False)
        .head(10)
    )

    st.dataframe(top_lines)

    fig3 = px.bar(
        x=top_lines.index,
        y=top_lines.values
    )

    st.plotly_chart(
        fig3,
        use_container_width=True
    )

    # =================================================
    # REP PERFORMANCE
    # =================================================
    st.subheader("👨‍💼 Rep Performance")

    rep = (
        df.groupby('User')['Amount']
        .sum()
        .sort_values(ascending=False)
    )

    st.dataframe(rep)

    fig4 = px.bar(
        x=rep.index,
        y=rep.values
    )

    st.plotly_chart(
        fig4,
        use_container_width=True
    )