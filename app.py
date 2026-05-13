from __future__ import annotations

import os
import re
import unicodedata
from difflib import SequenceMatcher
from pathlib import Path

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


def summarize_insights(question: str, data: str, system_prompt: str) -> str | None:
    client = get_openai_client()
    if not client or not data:
        return None
    try:
        r = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": system_prompt,
                },
                {
                    "role": "user",
                    "content": f"Question: {question}\n\nData:\n{data}\n\nProvide insights.",
                },
            ],
            max_tokens=200,
        )
        return (r.choices[0].message.content or "").strip() or None
    except Exception:
        return None


# =====================================================
# DATA LOAD
# =====================================================

def _first_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
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


def _map_qb_line_to_canonical(raw: str) -> str:
    s = str(raw).strip()
    if not s or s.lower() == "nan":
        return ""
    sl = normalize_for_match(s)
    for kl in sorted(KNOWN_LINES, key=len, reverse=True):
        kln = normalize_for_match(kl)
        if kln in sl or sl in kln:
            return kl
        if line_tokens_match_query(kl, s):
            return kl
    return s.strip()


@st.cache_data
def load_data():
    df = pd.read_excel("master_orders.xlsx")
    df.columns = df.columns.str.strip()

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")
    # For master orders, Amount represents the quantity ordered (no rand values)
    df["Quantity"] = df["Amount"]

    # Map line column
    c_line = _first_column(df, ["Line", "Product/service", "Item", "Description", "Class", "Item description"])
    if c_line:
        df["Line"] = df[c_line].map(_map_qb_line_to_canonical)
    else:
        df["Line"] = ""

    return df


df_orders = load_data()


def _app_dir() -> Path:
    return Path(__file__).resolve().parent


def discover_quickbooks_excel_path() -> str | None:
    """Resolve QB export next to this script (not cwd). Tries env QUICKBOOKS_XLSX then common filenames."""
    base = _app_dir()
    env = os.getenv("QUICKBOOKS_XLSX", "").strip()
    if env:
        p = Path(env)
        if not p.is_absolute():
            p = base / p
        if p.is_file():
            return str(p)
    for name in (
        "quickbooks_lines.xlsx",
        "quickbooks_line.xlsx",
        "QuickBooks_Lines.xlsx",
        "QuickBooks_Line.xlsx",
    ):
        fp = base / name
        if fp.is_file():
            return str(fp)
    for fp in sorted(base.glob("quickbooks*.xlsx"), key=lambda x: str(x).lower()):
        if fp.is_file():
            return str(fp)
    for fp in sorted(base.glob("quickbooks*.xls"), key=lambda x: str(x).lower()):
        if fp.is_file():
            return str(fp)
    return None


@st.cache_data
def load_quickbooks(resolved_path: str, _file_mtime: float) -> pd.DataFrame | None:
    """Load QuickBooks line export (invoices + credit notes). `_file_mtime` busts cache when the file changes."""
    if not resolved_path or not os.path.isfile(resolved_path):
        return None
    try:
        raw = pd.read_excel(resolved_path)
    except Exception:
        return None
    raw.columns = raw.columns.astype(str).str.strip()

    c_date = _first_column(raw, ["Date", "Txn date", "Transaction date", "Invoice date", "Posting date"])
    c_line = _first_column(raw, ["Line", "Product/service", "Item", "Description", "Class", "Item description"])
    c_amt = _first_column(raw, ["Amount", "Net amount", "Line amount", "Sales price", "Total"])
    if not c_date or not c_amt:
        return None
    if not c_line:
        c_line = c_amt

    c_type = _first_column(
        raw,
        ["Transaction type", "Txn type", "Type", "Document type", "Memo", "Source"],
    )
    c_customer = _first_column(raw, ["Customer", "Customer name", "Name"])
    c_doc = _first_column(raw, ["Num", "No", "Doc no", "Invoice no", "Reference"])
    c_qty = _first_column(raw, ["Qty", "Quantity", "Units", "QTY"])

    out = pd.DataFrame()
    out["Date"] = pd.to_datetime(raw[c_date], errors="coerce")
    out["Line"] = raw[c_line].map(_map_qb_line_to_canonical)
    amt = pd.to_numeric(raw[c_amt], errors="coerce").fillna(0.0)
    if c_type:
        types = raw[c_type].astype(str).str.lower()
        is_credit = types.str.contains(r"credit|refund|return|credit memo|credit note", regex=True, na=False)
        amt = amt.where(~is_credit | (amt <= 0), -amt.abs())
    out["Amount"] = amt
    if c_qty:
        qty = pd.to_numeric(raw[c_qty], errors="coerce").fillna(0.0)
        # For credits, quantity might be positive, keep as is
        out["Quantity"] = qty
    else:
        out["Quantity"] = 0.0
    out["QB_DocType"] = raw[c_type] if c_type else ""
    out["Client Name"] = raw[c_customer] if c_customer else ""
    out["QB_DocNo"] = raw[c_doc] if c_doc else ""
    out["Crop Name"] = ""
    out["Variety"] = ""
    out = out.dropna(subset=["Date"], how="all")
    out = out[out["Line"].astype(str).str.len() > 0]
    return out


_qb_resolved = discover_quickbooks_excel_path()
if _qb_resolved:
    df_qb = load_quickbooks(_qb_resolved, os.path.getmtime(_qb_resolved))
else:
    df_qb = None
df = df_orders


def catalog_subsumed_by_line(line: str, catalog_name: str) -> bool:
    """True if a crop/variety label repeats the line name in a looser/shorter form (e.g. Petunia vs PETUNIA HYBRIDS)."""
    ln = normalize_for_match(line)
    nn = normalize_for_match(catalog_name)
    if not nn or not ln:
        return False
    if nn == ln:
        return True
    if len(nn) < 4:
        return False
    if nn not in ln:
        return False
    return len(nn) < len(ln) - 2


def catalog_entry_matches(term: str, question: str, min_chars: int = 2) -> bool:
    s = str(term).strip()
    if len(s) < min_chars:
        return False
    sn = normalize_for_match(s)
    qn = normalize_for_match(question)
    if sn in qn:
        return True
    if re.sub(r"\s+", "", sn) in re.sub(r"\s+", "", qn):
        return True
    stoks = [t for t in tokenize_phrase(s) if len(t) >= min_chars or t.isdigit()]
    if not stoks:
        return False
    qtoks = tokenize_phrase(question)
    if not all(any(tokens_match(st, qt) for qt in qtoks) for st in stoks):
        return False
    if len(stoks) == 1 and len(stoks[0]) < 5:
        return sn in qn or re.sub(r"\s+", "", sn) in re.sub(r"\s+", "", qn)
    return True


def catalog_matches_fuzzy(question: str, catalog: list[str], min_chars: int = 2) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for term in catalog:
        s = str(term).strip()
        if s in seen:
            continue
        if catalog_entry_matches(s, question, min_chars=min_chars):
            seen.add(s)
            out.append(s)
    return out


def filter_sales(
    d: pd.DataFrame,
    *,
    line=None,
    crop=None,
    variety=None,
    client=None,
    years=None,
    months=None,
    days=None,
    doc_no=None,
    doc_type=None,
):
    out = d.copy()
    if line:
        out = out[out["Line"].astype(str).str.contains(line, case=False, na=False)]
    if crop and "Crop Name" in out.columns:
        out = out[out["Crop Name"].astype(str).str.contains(crop, case=False, na=False)]
    if variety and "Variety" in out.columns:
        out = out[out["Variety"].astype(str).str.contains(variety, case=False, na=False)]
    if client and "Client Name" in out.columns:
        out = out[out["Client Name"].astype(str).str.contains(client, case=False, na=False)]
    if doc_no and "QB_DocNo" in out.columns:
        out = out[out["QB_DocNo"].astype(str).str.contains(str(doc_no), case=False, na=False)]
    if doc_type:
        if "QB_DocType" in out.columns:
            if doc_type == "credit":
                out = out[out["QB_DocType"].astype(str).str.contains(r"credit|refund|return|memo", case=False, na=False)]
            elif doc_type == "invoice":
                out = out[~out["QB_DocType"].astype(str).str.contains(r"credit|refund|return|memo", case=False, na=False)]
        elif "Amount" in out.columns:
            if doc_type == "credit":
                out = out[out["Amount"] < 0]
            elif doc_type == "invoice":
                out = out[out["Amount"] >= 0]
    if years:
        out = out[out["Date"].dt.year.isin(years)]
    if months:
        out = out[out["Date"].dt.month.isin(months)]
    if days:
        out = out[out["Date"].dt.day.isin(days)]
    return out


# =====================================================
# INTENT ENGINE
# =====================================================


def detect_intent(q: str, df_in: pd.DataFrame | None = None):
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
        "day": None,
        "doc_no": None,
        "doc_type": None,
        "metric": "amount",
    }

    if any(x in ql for x in ["compare", "vs", "versus"]):
        intent["compare"] = True

    if any(x in ql for x in ["how many", "total", "sold", "amount", "invoice", "credit note", "credit memo"]):
        intent["total"] = True

    if any(x in ql for x in ["top", "best", "highest"]):
        intent["top"] = True

    intent["line"] = primary_line_from_query(q)

    intent["crop"] = None
    intent["variety"] = None
    intent["client"] = None

    if df_in is not None:
        crops_cat, var_cat = sorted_crops_and_varieties(df_in)
        pline = intent["line"]
        crop_hits = catalog_matches_fuzzy(q, crops_cat, min_chars=3)
        var_hits = catalog_matches_fuzzy(q, var_cat, min_chars=2)
        if pline:
            crop_hits = [c for c in crop_hits if not catalog_subsumed_by_line(pline, c)]
            var_hits = [v for v in var_hits if not catalog_subsumed_by_line(pline, v)]
        if len(crop_hits) == 1:
            intent["crop"] = crop_hits[0]
        elif "petunia" in ql and not crop_hits and not pline:
            intent["crop"] = "PETUNIA"
        if len(var_hits) == 1:
            intent["variety"] = var_hits[0]
        if "Client Name" in df_in.columns:
            client_hits = catalog_matches_fuzzy(q, df_in["Client Name"].dropna().astype(str).unique().tolist(), min_chars=3)
            if len(client_hits) == 1:
                intent["client"] = client_hits[0]
    elif "petunia" in ql:
        intent["crop"] = "PETUNIA"

    date_parts = extract_day_month_year_from_query(ql)
    if date_parts["day"]:
        intent["day"] = date_parts["day"]
    if date_parts["month"]:
        intent["month"] = date_parts["month"]
    if date_parts["year"]:
        intent["year"] = date_parts["year"]

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

    intent["doc_no"] = extract_invoice_number_from_query(ql)
    if is_credit_query(ql):
        intent["doc_type"] = "credit"
    elif is_invoice_query(ql) or is_sales_query(ql):
        intent["doc_type"] = "invoice"

    if any(x in ql for x in ["how many", "quantity", "qty", "units", "pieces", "number of"]):
        intent["metric"] = "quantity"
    else:
        intent["metric"] = "amount"

    return intent


def apply_filters(df_in, intent):
    return filter_sales(
        df_in,
        line=intent.get("line"),
        crop=intent.get("crop"),
        variety=intent.get("variety"),
        client=intent.get("client"),
        years=[intent["year"]] if intent.get("year") else None,
        months=[intent["month"]] if intent.get("month") else None,
        days=[intent["day"]] if intent.get("day") else None,
        doc_no=intent.get("doc_no"),
        doc_type=intent.get("doc_type"),
    )


def run_comparison(question: str, df_in: pd.DataFrame, metric: str = "amount"):
    ql = question.lower()
    lines = lines_matching_query(question)
    years_q = extract_years_from_query(ql)
    months_q = extract_months_from_query(ql)
    crops_catalog, varieties_catalog = sorted_crops_and_varieties(df_in)
    crops_found = catalog_matches_fuzzy(question, crops_catalog, min_chars=3)
    varieties_found = catalog_matches_fuzzy(question, varieties_catalog, min_chars=2)

    pl = primary_line_from_query(question)
    if pl:
        crops_found = [c for c in crops_found if not catalog_subsumed_by_line(pl, c)]
        varieties_found = [v for v in varieties_found if not catalog_subsumed_by_line(pl, v)]

    line_ref = pl or (lines[0] if len(lines) == 1 else None)
    scope_line = line_ref
    scope_crop = crops_found[0] if len(crops_found) == 1 else None
    scope_variety = varieties_found[0] if len(varieties_found) == 1 else None

    year_filter = years_q if len(years_q) >= 1 else None
    month_filter = months_q if len(months_q) >= 1 else None

    results = {}
    mode_label = ""
    col = "Amount" if metric == "amount" else "Quantity"

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
        mode_label = f"Compare lines ({metric})"
        for line in lines:
            seg = filter_sales(
                df_in,
                line=line,
                years=year_filter,
                months=month_filter,
            )
            results[line] = float(seg[col].sum())

    elif len(months_q) >= 2 and line_ref:
        mode_label = f"Compare months for line ({metric})"
        for m in months_q:
            seg = filter_sales(df_in, line=line_ref, years=year_filter, months=[m])
            label = MONTH_NAMES.get(m, str(m))
            results[label] = float(seg[col].sum())

    elif len(months_q) >= 2 and len(crops_found) == 1:
        mode_label = f"Compare months for crop ({metric})"
        for m in months_q:
            seg = filter_sales(df_in, crop=scope_crop, years=year_filter, months=[m])
            label = MONTH_NAMES.get(m, str(m))
            results[label] = float(seg[col].sum())

    elif len(months_q) >= 2 and len(varieties_found) == 1:
        mode_label = f"Compare months for variety ({metric})"
        for m in months_q:
            seg = filter_sales(df_in, variety=scope_variety, years=year_filter, months=[m])
            label = MONTH_NAMES.get(m, str(m))
            results[label] = float(seg[col].sum())

    elif len(crops_found) >= 2:
        mode_label = f"Compare crop names ({metric})"
        line_scope = line_ref if len(lines) == 1 else None
        for crop in crops_found:
            seg = filter_sales(
                df_in,
                crop=crop,
                line=line_scope,
                years=year_filter,
                months=month_filter,
            )
            results[crop] = float(seg[col].sum())

    elif len(varieties_found) >= 2:
        mode_label = f"Compare varieties ({metric})"
        line_scope = line_ref if len(lines) == 1 else None
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
            results[var] = float(seg[col].sum())

    elif line_ref and len(years_q) >= 2:
        mode_label = f"Compare years for line ({metric})"
        for y in years_q:
            seg = filter_sales(df_in, line=line_ref, years=[y], months=month_filter)
            results[str(y)] = float(seg[col].sum())

    elif len(crops_found) == 1 and len(years_q) >= 2:
        mode_label = f"Compare years for crop ({metric})"
        for y in years_q:
            seg = filter_sales(df_in, crop=scope_crop, years=[y], months=month_filter)
            results[str(y)] = float(seg[col].sum())

    elif len(varieties_found) == 1 and len(years_q) >= 2:
        mode_label = f"Compare years for variety ({metric})"
        for y in years_q:
            seg = filter_sales(df_in, variety=scope_variety, years=[y], months=month_filter)
            results[str(y)] = float(seg[col].sum())

    elif len(years_q) >= 2:
        mode_label = f"Compare calendar years ({metric})"
        base = scoped_base()
        for y in years_q:
            seg = base[base["Date"].dt.year == y]
            results[str(y)] = float(seg[col].sum())

    elif len(months_q) >= 2:
        mode_label = f"Compare months ({metric})"
        base = scoped_base()
        for m in months_q:
            seg = base[base["Date"].dt.month == m]
            label = MONTH_NAMES.get(m, str(m))
            results[label] = float(seg[col].sum())

    else:
        st.warning(
            "For a comparison, name at least two lines, two crops, two varieties, "
            "two months, or two years — or one line, crop, or variety with two months/years "
            "(e.g. Petunia Hybrids in May vs June last year)."
        )
        return

    st.subheader(mode_label)
    st.dataframe(pd.Series(results, name=col))

    fig = px.bar(x=list(results.keys()), y=list(results.values()))
    st.plotly_chart(fig, use_container_width=True)

    summary_lines = [f"{k}: {v:,.2f}" for k, v in results.items()]
    insight = summarize_comparison(question, summary_lines)
    if insight:
        st.info(insight)
    elif not openai_api_key():
        st.caption("Add `OPENAI_API_KEY` or `openai_api_key` in secrets for AI summaries.")

    st.caption(f"Filters applied — years: {year_filter or 'any'}, months: {month_filter or 'any'}")


QB_KEYWORDS = (
    "quickbooks",
    "qb ",
    " qb",
    "credit note",
    "credit memo",
    "invoice total",
    "invoice",
    "return",
    "returns",
    "refund",
    "sales",
    "sold",
    "rand",
    "rands",
    "value",
    "revenue",
    "made",
    "actual sales",
    "financial",
    "accurate sales",
    "from qb",
)
ORDER_KEYWORDS = (
    "rep order",
    "app order",
    "what we ordered",
    "ordered for",
    "master order",
    "store order",
    "from app",
)


def question_prefers_orders(ql: str) -> bool:
    if any(k in ql for k in ORDER_KEYWORDS):
        return True
    if re.search(r"\b(order|ordered|quantity|qty|units|pieces)\b", ql):
        return True
    return False


def question_prefers_quickbooks(ql: str) -> bool:
    if any(k in ql for k in QB_KEYWORDS):
        return True
    if "store" in ql and re.search(r"\b(sales|sold|returns|return|invoice|credit|rand|rands|value|revenue|made)\b", ql):
        return True
    return False


def resolve_active_dataframe(
    question: str,
    sidebar_choice: str,
    orders_df: pd.DataFrame,
    qb_df: pd.DataFrame | None,
) -> tuple[pd.DataFrame, str, bool]:
    """Pick which dataframe powers the search bar. Returns (df, label, compare_sources_mode)."""
    ql = (question or "").lower()
    has_qb = qb_df is not None and len(qb_df) > 0

    if has_qb and sidebar_choice.startswith("Compare"):
        return orders_df, "compare_sources", True

    if has_qb and sidebar_choice.startswith("QuickBooks"):
        return qb_df, "QuickBooks (invoices & credits, line level)", False

    if has_qb and question_prefers_quickbooks(ql) and not question_prefers_orders(ql):
        return qb_df, "QuickBooks (from your question)", False

    if question_prefers_orders(ql) and has_qb and not sidebar_choice.startswith("QuickBooks"):
        return orders_df, "Rep orders (from your question)", False

    return orders_df, "Rep orders (app)", False


def render_orders_vs_quickbooks(question: str, orders_df: pd.DataFrame, qb_df: pd.DataFrame):
    ql = (question or "").lower()
    lines = lines_matching_query(question) if ql else []
    years = extract_years_from_query(ql)
    months = extract_months_from_query(ql)
    yf = years if years else None
    mf = months if months else None

    st.subheader("Rep orders vs QuickBooks (line level)")
    if lines:
        rows = []
        for line in lines[:15]:
            o = float(filter_sales(orders_df, line=line, years=yf, months=mf)["Amount"].sum())
            q = float(filter_sales(qb_df, line=line, years=yf, months=mf)["Amount"].sum())
            rows.append({"Line": line, "Rep_app_Amount": o, "QB_net_Amount": q})
        tdf = pd.DataFrame(rows)
    else:
        o_s = orders_df.groupby("Line", dropna=False)["Amount"].sum()
        q_s = qb_df.groupby("Line", dropna=False)["Amount"].sum()
        tdf = pd.DataFrame({"Rep_app_Amount": o_s, "QB_net_Amount": q_s}).fillna(0)
        tdf = tdf.sort_values("Rep_app_Amount", ascending=False).head(30)
        tdf = tdf.reset_index()

    st.dataframe(tdf, use_container_width=True)
    if len(tdf) > 0:
        id_col = "Line" if "Line" in tdf.columns else tdf.columns[0]
        tdf_m = tdf.melt(
            id_vars=[id_col],
            value_vars=["Rep_app_Amount", "QB_net_Amount"],
            var_name="Source",
            value_name="Amount",
        )
        fig = px.bar(
            tdf_m,
            x=id_col,
            y="Amount",
            color="Source",
            barmode="group",
            labels={id_col: "Line"},
        )
        st.plotly_chart(fig, use_container_width=True)

    st.caption(
        "Rep app: ordering data (crop/place detail; can include duplicates or replenishment). "
        "QuickBooks: invoice + credit line amounts in your file’s currency (net sales vs returns at line level)."
    )


# =====================================================
# UI
# =====================================================

st.title("Nursery Intelligence Copilot v2.1")

with st.sidebar:
    st.subheader("Data source")
    qb_ok = df_qb is not None and len(df_qb) > 0
    if not qb_ok:
        st.caption(
            "Put **quickbooks_line.xlsx** or **quickbooks_lines.xlsx** in the same folder as "
            "`app.py` (not only your working directory), or set **QUICKBOOKS_XLSX** to the full path."
        )
    opts = ["Rep orders (app — crop & store detail)"]
    if qb_ok:
        opts.append("QuickBooks (accurate money — line level)")
        opts.append("Compare rep orders vs QuickBooks (by line)")
    data_choice = st.radio("Use for the search bar", opts, index=0, key="data_source_radio")

with st.expander("How rep orders vs QuickBooks fit together"):
    st.markdown(
        """
**Rep orders (app)** — What reps ordered for stores: best for **which crops/varieties** went to **which places**.  
Totals are **not** the same as final sales (duplicates, swaps, top-ups).

**QuickBooks** — **Invoices and credit notes**: best for **Rand (or file currency)**, **sales vs returns**, and **true line-level revenue**.  
QuickBooks usually **does not** break out each crop the way your app does.

Use **Compare** to put the two side by side on **Line** (same names as in your app where possible).
        """
    )

question = st.text_input("Ask anything about sales")

df_active, source_label, compare_sources = resolve_active_dataframe(
    question, data_choice, df_orders, df_qb
)
st.caption(f"Active dataset: **{source_label}**")

if question:

    intent = detect_intent(question, df_active)

    if compare_sources and qb_ok:
        render_orders_vs_quickbooks(question, df_orders, df_qb)

    df_temp = apply_filters(df_active, intent)

    if len(df_temp) == 0:
        df_temp = df_active.copy()

    ql = question.lower()

    # Combined summary for line/crop/variety queries
    if (intent["line"] or intent["crop"] or intent["variety"]) and df_qb is not None and df_orders is not None and not compare_sources:
        summary_intent = {
            "line": intent["line"],
            "crop": intent["crop"],
            "variety": intent["variety"],
            "client": intent["client"],
            "year": intent["year"],
            "month": intent["month"],
            "day": intent["day"],
        }
        # Ordered quantity
        ordered_intent = summary_intent.copy()
        ordered_intent["metric"] = "quantity"
        ordered_df = apply_filters(df_orders, ordered_intent)
        ordered_qty = ordered_df["Quantity"].sum()

        # Sales amount
        sales_intent = summary_intent.copy()
        sales_intent["doc_type"] = "invoice"
        sales_intent["metric"] = "amount"
        sales_df = apply_filters(df_qb, sales_intent)
        sales_amt = sales_df["Amount"].sum()

        # Returns amount
        returns_intent = summary_intent.copy()
        returns_intent["doc_type"] = "credit"
        returns_df = apply_filters(df_qb, returns_intent)
        returns_amt = abs(returns_df["Amount"].sum())

        net_sales = sales_amt - returns_amt

        item_name = intent["line"] or intent["crop"] or intent["variety"] or "items"
        data = f"Ordered Quantity: {ordered_qty:,.0f}\nSales Amount: R{sales_amt:,.2f}\nReturns Amount: R{returns_amt:,.2f}\nNet Sales: R{net_sales:,.2f}"
        system_prompt = "You assist nursery management. Provide a concise, insightful summary of the ordering and sales performance for the product, highlighting key metrics and implications in 2-4 sentences."
        insight = summarize_insights(question, data, system_prompt)
        if insight:
            st.info(insight)
        else:
            st.info(f"I see there were {ordered_qty:,.0f} {item_name} ordered, but there was R{sales_amt:,.2f} worth of sales and R{returns_amt:,.2f} in returns, giving R{net_sales:,.2f} in net sales.")

    if intent["compare"] and not compare_sources:
        run_comparison(question, df_active, intent["metric"])
        df_temp = df_active.copy()

    elif intent["top"]:

        group_col = "Client Name"
        if "crop" in ql and "Crop Name" in df_temp.columns:
            group_col = "Crop Name"
        elif "variety" in ql and "Variety" in df_temp.columns:
            group_col = "Variety"
        elif "line" in ql or "Crop Name" not in df_temp.columns:
            group_col = "Line"
        elif "client" in ql and "Client Name" in df_temp.columns:
            group_col = "Client Name"
        else:
            group_col = "Line"

        col = "Amount" if intent["metric"] == "amount" else "Quantity"
        result = df_temp.groupby(group_col)[col].sum().sort_values(ascending=False).head(10)

        st.subheader(f"Top {group_col} ({intent['metric']})")
        st.dataframe(result)

        fig = px.bar(x=result.index, y=result.values)
        st.plotly_chart(fig, use_container_width=True)

    elif not compare_sources:

        col = "Amount" if intent["metric"] == "amount" else "Quantity"
        total = df_temp[col].sum()
        unit = "R" if intent["metric"] == "amount" and df_active is df_qb else ""
        st.success(f"Total {intent['metric']}: {unit}{total:,.2f}")

    st.subheader("Matching Data")
    if intent["compare"]:
        st.caption(
            "Full data preview. Comparisons use years, months, lines, crops, and varieties detected in your question."
        )
    st.dataframe(df_temp)

st.header("Dashboard")

tab_orders, tab_qb, tab_about = st.tabs(["Rep orders", "QuickBooks", "About the data"])

with tab_orders:
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Orders", len(df_orders))
    col2.metric("Total Amount", f"{df_orders['Amount'].sum():,}")
    col3.metric("Total Quantity", f"{df_orders['Quantity'].sum():,}")
    col4.metric("Clients", df_orders["Client Name"].nunique())
    st.subheader("Top Clients")
    st.dataframe(df_orders.groupby("Client Name")["Amount"].sum().sort_values(ascending=False).head(10))
    st.subheader("Top Crops")
    st.dataframe(df_orders.groupby("Crop Name")["Amount"].sum().sort_values(ascending=False).head(10))
    st.subheader("Top Lines")
    st.dataframe(df_orders.groupby("Line")["Amount"].sum().sort_values(ascending=False).head(10))

with tab_qb:
    if not qb_ok:
        st.info(
            "Put an Excel export in the **same folder as app.py** named **quickbooks_line.xlsx** or "
            "**quickbooks_lines.xlsx** (or any **quickbooks*.xlsx**), or set **QUICKBOOKS_XLSX** to the full file path. "
            "Expected columns include **Date**, a **line/item** column, and **Amount**; optional **Transaction type** "
            "(rows with Credit / Credit memo / Refund are treated as returns)."
        )
    else:
        sales_amt = df_qb.loc[df_qb["Amount"] > 0, "Amount"].sum()
        ret_amt = df_qb.loc[df_qb["Amount"] < 0, "Amount"].sum()
        sales_qty = df_qb.loc[df_qb["Amount"] > 0, "Quantity"].sum()
        ret_qty = df_qb.loc[df_qb["Amount"] < 0, "Quantity"].sum()
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("QB rows", len(df_qb))
        c2.metric("Net total", f"{df_qb['Amount'].sum():,.2f}")
        c3.metric("Sales (positive lines)", f"{sales_amt:,.2f}")
        c4.metric("Returns (negative lines)", f"{ret_amt:,.2f}")
        c5.metric("Sales Qty", f"{sales_qty:,.0f}")
        c6.metric("Returns Qty", f"{ret_qty:,.0f}")
        st.subheader("Net by line")
        st.dataframe(df_qb.groupby("Line")["Amount"].sum().sort_values(ascending=False))
        st.subheader("Sample rows")
        st.dataframe(df_qb.head(50))

with tab_about:
    st.markdown(
        """
| Source | Best for | Limitation |
|--------|-----------|------------|
| **Rep orders (app)** | Crop, variety, customer/place, what was **ordered** (quantity & order amount, not rands) | Not exact sell-through; can double-count or replace stock |
| **QuickBooks** | **Money**: sales vs credits, **line** rollups (quantity & amount) | Usually no per-crop breakdown like the app |

The search bar uses the **sidebar data source**, unless your question clearly asks for the other (e.g. "QuickBooks" / "invoices" vs "orders" / "crops").

You can ask for **quantity** (how many) or **amount** (rand value). For returns, use "credit note" or "return"; for sales, "invoice" or "sold".
        """
    )
