from __future__ import annotations

import os
import re
import streamlit as st
import pandas as pd
import plotly.express as px
from openai import OpenAI
from werkzeug.security import check_password_hash

# =====================================================
# CONFIG
# =====================================================

# Always available: admin / 1234. Add more accounts in .streamlit/secrets.toml under [users].
BUILTIN_ADMIN_USERNAME = "admin"
BUILTIN_ADMIN_PASSWORD_HASH = (
    "scrypt:32768:8:1$cer4Rc98UopaCXLm$4d12e32c06be3f94fab0e22a80237c5d552f4b85ec0167b4c7f422050af7228161b83c239f4eae0fdc3dcdc17068492cc04dbcbc1a87bd565df57d776993abba"
)


def openai_api_key():
    env = os.getenv("OPENAI_API_KEY")
    if env:
        return env
    try:
        return st.secrets.get("openai_api_key", "") or ""
    except Exception:
        return ""


def user_password_hashes():
    """Streamlit wraps `[users]` in AttrDict, not plain dict — coerce before use."""
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
        pass
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


st.set_page_config(
    page_title="Nursery Intelligence Copilot v2.1",
    layout="wide",
)

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

# =====================================================
# LOGIN
# =====================================================

if not st.session_state.authenticated:
    st.title("Nursery Intelligence Copilot")
    if not user_password_hashes():
        st.info(
            "Sign in with **admin** / **1234**. When you are ready, add more users in "
            "`.streamlit/secrets.toml` under `[users]` (password hashes from Werkzeug)."
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


def logout_sidebar():
    with st.sidebar:
        if st.button("Log out"):
            st.session_state.authenticated = False
            st.rerun()


logout_sidebar()

# =====================================================
# OPENAI CLIENT (optional summaries)
# =====================================================

_openai_client = None


def get_openai_client():
    global _openai_client
    key = openai_api_key()
    if not key:
        return None
    if _openai_client is None:
        _openai_client = OpenAI(api_key=key)
    return _openai_client


def summarize_comparison(question: str, result_lines: list) -> str | None:
    client = get_openai_client()
    if not client or not result_lines:
        return None
    try:
        body = "\n".join(result_lines)
        r = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You assist nursery staff. Interpret sales comparison numbers briefly "
                        "and practically (2–5 short bullets, plain English)."
                    ),
                },
                {
                    "role": "user",
                    "content": f"User question: {question}\n\nData:\n{body}\n\nSummarize insights.",
                },
            ],
            max_tokens=350,
        )
        return (r.choices[0].message.content or "").strip() or None
    except Exception:
        return None


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
    "17CM GERANIUM",
]

MONTH_NAMES = {
    1: "January",
    2: "February",
    3: "March",
    4: "April",
    5: "May",
    6: "June",
    7: "July",
    8: "August",
    9: "September",
    10: "October",
    11: "November",
    12: "December",
}

MONTH_WORDS = {
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
    "december": 12,
}


@st.cache_data
def sorted_crops_and_varieties(_df: pd.DataFrame):
    crops = sorted(
        _df["Crop Name"].dropna().astype(str).unique(),
        key=lambda x: len(str(x)),
        reverse=True,
    )
    varieties = sorted(
        _df["Variety"].dropna().astype(str).unique(),
        key=lambda x: len(str(x)),
        reverse=True,
    )
    return crops, varieties


def extract_years_from_query(ql: str) -> list[int]:
    cy = pd.Timestamp.today().year
    years = []
    if "last year" in ql:
        years.append(cy - 1)
    if "this year" in ql:
        years.append(cy)
    years.extend(int(y) for y in re.findall(r"\b(20\d{2})\b", ql))
    return sorted(set(years))


def extract_months_from_query(ql: str) -> list[int]:
    found = []
    for name, num in MONTH_WORDS.items():
        if name in ql:
            found.append(num)
    return sorted(set(found))


def detect_lines_in_query(ql: str) -> list[str]:
    ql = ql.lower()
    return [line for line in KNOWN_LINES if line.lower() in ql]


def terms_from_query(ql: str, catalog: list[str], min_len: int = 3) -> list[str]:
    ql_lower = ql.lower()
    matched = []
    for term in catalog:
        s = str(term).strip()
        if len(s) < min_len:
            continue
        if s.lower() in ql_lower:
            matched.append(s)
    return matched


def filter_sales(
    d: pd.DataFrame,
    *,
    line=None,
    crop=None,
    variety=None,
    client=None,
    years=None,
    months=None,
):
    out = d.copy()
    if line:
        out = out[out["Line"].astype(str).str.contains(line, case=False, na=False)]
    if crop:
        out = out[out["Crop Name"].astype(str).str.contains(crop, case=False, na=False)]
    if variety:
        out = out[out["Variety"].astype(str).str.contains(variety, case=False, na=False)]
    if client:
        out = out[out["Client Name"].astype(str).str.contains(client, case=False, na=False)]
    if years:
        out = out[out["Date"].dt.year.isin(years)]
    if months:
        out = out[out["Date"].dt.month.isin(months)]
    return out


# =====================================================
# INTENT ENGINE
# =====================================================


def detect_intent(q: str):
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
    }

    if any(x in ql for x in ["compare", "vs", "versus"]):
        intent["compare"] = True

    if any(x in ql for x in ["how many", "total", "sold", "amount"]):
        intent["total"] = True

    if any(x in ql for x in ["top", "best", "highest"]):
        intent["top"] = True

    for line in KNOWN_LINES:
        if line.lower() in ql:
            intent["line"] = line
            break

    intent["crop"] = None
    intent["variety"] = None
    intent["client"] = None

    if "petunia" in ql:
        intent["crop"] = "PETUNIA"

    current_year = pd.Timestamp.today().year

    if "last year" in ql:
        intent["year"] = current_year - 1

    if "this year" in ql:
        intent["year"] = current_year

    year_match = re.findall(r"\b(20\d{2})\b", ql)
    if year_match:
        intent["year"] = int(year_match[0])

    for m, v in MONTH_WORDS.items():
        if m in ql:
            intent["month"] = v
            break

    return intent


def apply_filters(df_in, intent):
    d = df_in.copy()

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


def run_comparison(question: str, df_in: pd.DataFrame):
    ql = question.lower()
    lines = detect_lines_in_query(ql)
    years_q = extract_years_from_query(ql)
    months_q = extract_months_from_query(ql)
    crops_catalog, varieties_catalog = sorted_crops_and_varieties(df_in)
    crops_found = terms_from_query(ql, crops_catalog)
    varieties_found = terms_from_query(ql, varieties_catalog)

    scope_line = lines[0] if len(lines) == 1 else None
    scope_crop = crops_found[0] if len(crops_found) == 1 else None
    scope_variety = varieties_found[0] if len(varieties_found) == 1 else None

    year_filter = years_q if len(years_q) >= 1 else None
    month_filter = months_q if len(months_q) >= 1 else None

    results = {}
    mode_label = ""

    def scoped_base():
        return filter_sales(
            df_in,
            line=scope_line,
            crop=scope_crop,
            variety=scope_variety,
            years=year_filter,
            months=month_filter,
        )

    if len(lines) >= 2:
        mode_label = "Compare lines (amount)"
        for line in lines:
            seg = filter_sales(
                df_in,
                line=line,
                years=year_filter,
                months=month_filter,
            )
            results[line] = float(seg["Amount"].sum())

    elif len(crops_found) >= 2:
        mode_label = "Compare crop names (amount)"
        line_scope = scope_line if len(lines) == 1 else None
        for crop in crops_found:
            seg = filter_sales(
                df_in,
                crop=crop,
                line=line_scope,
                years=year_filter,
                months=month_filter,
            )
            results[crop] = float(seg["Amount"].sum())

    elif len(varieties_found) >= 2:
        mode_label = "Compare varieties (amount)"
        line_scope = scope_line if len(lines) == 1 else None
        crop_scope = scope_crop if len(crops_found) == 1 else None
        for var in varieties_found:
            seg = filter_sales(
                df_in,
                variety=var,
                line=line_scope,
                crop=crop_scope,
                years=year_filter,
                months=month_filter,
            )
            results[var] = float(seg["Amount"].sum())

    elif len(lines) == 1 and len(years_q) >= 2:
        mode_label = "Compare years for line (amount)"
        for y in years_q:
            seg = filter_sales(df_in, line=scope_line, years=[y], months=month_filter)
            results[str(y)] = float(seg["Amount"].sum())

    elif len(lines) == 1 and len(months_q) >= 2:
        mode_label = "Compare months for line (amount)"
        for m in months_q:
            seg = filter_sales(df_in, line=scope_line, years=year_filter, months=[m])
            label = MONTH_NAMES.get(m, str(m))
            results[label] = float(seg["Amount"].sum())

    elif len(crops_found) == 1 and len(years_q) >= 2:
        mode_label = "Compare years for crop (amount)"
        for y in years_q:
            seg = filter_sales(df_in, crop=scope_crop, years=[y], months=month_filter)
            results[str(y)] = float(seg["Amount"].sum())

    elif len(varieties_found) == 1 and len(years_q) >= 2:
        mode_label = "Compare years for variety (amount)"
        for y in years_q:
            seg = filter_sales(df_in, variety=scope_variety, years=[y], months=month_filter)
            results[str(y)] = float(seg["Amount"].sum())

    elif len(years_q) >= 2:
        mode_label = "Compare calendar years (amount)"
        base = scoped_base()
        for y in years_q:
            seg = base[base["Date"].dt.year == y]
            results[str(y)] = float(seg["Amount"].sum())

    elif len(months_q) >= 2:
        mode_label = "Compare months (amount)"
        base = scoped_base()
        for m in months_q:
            seg = base[base["Date"].dt.month == m]
            label = MONTH_NAMES.get(m, str(m))
            results[label] = float(seg["Amount"].sum())

    else:
        st.warning(
            "For a comparison, name at least two lines, two crops, two varieties, "
            "two months, or two years — or one line/crop/variety with two years/months."
        )
        return

    st.subheader(mode_label)
    st.dataframe(pd.Series(results, name="Amount"))

    fig = px.bar(x=list(results.keys()), y=list(results.values()))
    st.plotly_chart(fig, use_container_width=True)

    summary_lines = [f"{k}: {v:,.2f}" for k, v in results.items()]
    insight = summarize_comparison(question, summary_lines)
    if insight:
        st.info(insight)
    elif not openai_api_key():
        st.caption("Add `OPENAI_API_KEY` or `openai_api_key` in secrets for AI summaries.")

    st.caption(f"Filters applied — years: {year_filter or 'any'}, months: {month_filter or 'any'}")


# =====================================================
# UI
# =====================================================

st.title("Nursery Intelligence Copilot v2.1")

question = st.text_input("Ask anything about sales")

if question:

    intent = detect_intent(question)
    df_temp = apply_filters(df, intent)

    if len(df_temp) == 0:
        df_temp = df.copy()

    ql = question.lower()

    if intent["compare"]:
        run_comparison(question, df)
        df_temp = df.copy()

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

    else:

        total = df_temp["Amount"].sum()
        st.success(f"Total Sold: {total:,}")

    st.subheader("Matching Data")
    if intent["compare"]:
        st.caption(
            "Full data preview. Comparisons use years, months, lines, crops, and varieties detected in your question."
        )
    st.dataframe(df_temp)

st.header("Dashboard")

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
