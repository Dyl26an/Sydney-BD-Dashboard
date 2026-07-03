import io
import re
from typing import Optional, List, Dict, Tuple

import msoffcrypto
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Sydney Growth Intelligence V5", layout="wide")

# -----------------------------
# Helpers
# -----------------------------

def norm(s: str) -> str:
    return re.sub(r"\s+", "", str(s).lower())


def find_col(cols: List[str], include_any: List[str], exclude_any: Optional[List[str]] = None) -> Optional[str]:
    exclude_any = exclude_any or []
    nmap = {c: norm(c) for c in cols}
    for c, n in nmap.items():
        if any(norm(k) in n for k in include_any) and not any(norm(x) in n for x in exclude_any):
            return c
    return None


def find_best_col(cols: List[str], candidates: List[str], exclude_any: Optional[List[str]] = None) -> Optional[str]:
    exclude_any = exclude_any or []
    ncols = {norm(c): c for c in cols}
    for cand in candidates:
        nc = norm(cand)
        if nc in ncols:
            return ncols[nc]
    for cand in candidates:
        col = find_col(cols, [cand], exclude_any)
        if col:
            return col
    return None


def to_num(s):
    if s is None:
        return pd.Series(dtype=float)
    return pd.to_numeric(
        pd.Series(s).astype(str).str.replace("%", "", regex=False).str.replace(",", "", regex=False).str.replace("$", "", regex=False),
        errors="coerce",
    )


def as_rate(series: pd.Series) -> pd.Series:
    x = to_num(series)
    # If most non-null values are > 1, treat as percentage points (e.g. 7.5 means 7.5%).
    non = x.dropna()
    if len(non) and non.quantile(0.75) > 1:
        x = x / 100.0
    return x.fillna(0)


def safe_div(a, b):
    return float(a) / float(b) if b and pd.notna(b) and float(b) != 0 else 0.0


def fmt_money(x):
    try:
        return f"${float(x):,.0f}"
    except Exception:
        return "$0"


def fmt_int(x):
    try:
        return f"{float(x):,.0f}"
    except Exception:
        return "0"


def fmt_pct(x):
    try:
        return f"{float(x) * 100:.2f}%"
    except Exception:
        return "0.00%"


def read_excel(uploaded_file, password: str) -> pd.DataFrame:
    raw = uploaded_file.read()
    bio = io.BytesIO(raw)
    decrypted = io.BytesIO()
    try:
        office = msoffcrypto.OfficeFile(bio)
        office.load_key(password=password)
        office.decrypt(decrypted)
        decrypted.seek(0)
        return pd.read_excel(decrypted, sheet_name=0)
    except Exception:
        # Some files may not actually be encrypted.
        bio.seek(0)
        return pd.read_excel(bio, sheet_name=0)


def detect_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    cols = list(df.columns)
    return {
        "bd_name": find_best_col(cols, ["bd姓名", "BD姓名", "bd name", "owner name", "负责人姓名", "bd名称", "姓名"]),
        "bd_id": find_best_col(cols, ["bd工号", "BD工号", "bd id", "owner", "负责人", "bd"]),
        "merchant_name": find_best_col(cols, ["商户名称", "店铺名称", "门店名称", "merchant name", "shop name", "store name", "餐厅名称"]),
        "merchant_id": find_best_col(cols, ["店铺id", "商户id", "门店id", "merchant id", "store id", "shop id"]),
        "area": find_best_col(cols, ["商圈", "区域", "area", "suburb", "business district"]),
        "category": find_best_col(cols, ["品类", "主营品类", "category", "cuisine", "菜系"]),
        "month": find_best_col(cols, ["月份", "月", "month", "reporting month", "数据月份"]),
        "gmv": find_best_col(cols, ["gmv", "实付gmv", "交易额", "销售额"], exclude_any=["同比", "环比", "率"]),
        "orders": find_best_col(cols, ["订单数_排除mm的均单", "订单数", "有效订单", "orders", "order count"], exclude_any=["配送类型", "转化率", "率", "必选"]),
        "exposure": find_best_col(cols, ["平均曝光人数", "曝光人数", "曝光", "impression", "exposure"], exclude_any=["转化率", "率"]),
        "visit": find_best_col(cols, ["平均进店人数", "进店人数", "进店", "visit"], exclude_any=["转化率", "率"]),
        "cart": find_best_col(cols, ["平均加购人数", "加购人数", "加购", "cart"], exclude_any=["转化率", "率"]),
        "exposure_order_rate": find_best_col(cols, ["曝光下单转化率", "曝光到下单", "exposure order" ]),
        "exposure_visit_rate": find_best_col(cols, ["曝光进店转化率", "曝光到进店", "exposure visit" ]),
        "visit_cart_rate": find_best_col(cols, ["进店加购转化率", "进店到加购", "visit cart" ]),
        "cart_order_rate": find_best_col(cols, ["加购下单转化率", "加购到下单", "cart order" ]),
        "promo": find_best_col(cols, ["是否有折扣", "有折扣", "折扣", "活动", "promo", "campaign"], exclude_any=["率", "金额"]),
        "material": find_best_col(cols, ["是否有物料", "有物料", "物料", "material"], exclude_any=["率"]),
        "visit_record": find_best_col(cols, ["拜访记录", "是否拜访", "拜访", "visit record"], exclude_any=["率"]),
    }


def truthy_rate(series: pd.Series) -> float:
    if series is None or len(series) == 0:
        return 0.0
    s = series.fillna("").astype(str).str.strip().str.lower()
    positive = s.isin(["1", "true", "yes", "y", "有", "是", "已配置", "已拜访", "已上传", "available", "active"])
    # numeric fallback: > 0
    nums = pd.to_numeric(s.str.replace("%", "", regex=False), errors="coerce")
    positive = positive | (nums.fillna(0) > 0)
    return positive.mean() if len(positive) else 0.0


def prep(df: pd.DataFrame, c: Dict[str, Optional[str]]) -> pd.DataFrame:
    out = df.copy()
    out["_bd_display"] = "Unknown BD"
    if c.get("bd_name"):
        out["_bd_display"] = out[c["bd_name"]].fillna("").astype(str).str.strip()
    if c.get("bd_id"):
        out["_bd_display"] = out["_bd_display"].where(out["_bd_display"].ne(""), out[c["bd_id"]].fillna("").astype(str).str.strip())
    out["_bd_display"] = out["_bd_display"].replace("", "Unknown BD")

    out["_merchant"] = out[c["merchant_name"]].fillna("").astype(str) if c.get("merchant_name") else "Unknown Merchant"
    out["_merchant_id"] = out[c["merchant_id"]].fillna("").astype(str) if c.get("merchant_id") else ""
    out["_area"] = out[c["area"]].fillna("Unknown").astype(str) if c.get("area") else "Unknown"
    out["_category"] = out[c["category"]].fillna("Unknown").astype(str) if c.get("category") else "Unknown"

    for key in ["gmv", "orders", "exposure", "visit", "cart"]:
        out[f"_{key}"] = to_num(out[c[key]]).fillna(0) if c.get(key) else 0

    # Official conversion-rate fields are monthly merchant-level rates. Aggregate by weighted average.
    for key in ["exposure_order_rate", "exposure_visit_rate", "visit_cart_rate", "cart_order_rate"]:
        out[f"_{key}"] = as_rate(out[c[key]]) if c.get(key) else 0

    # Opportunity flags
    out["_promo_flag"] = 0
    out["_material_flag"] = 0
    out["_visit_record_flag"] = 0
    if c.get("promo"):
        s = out[c["promo"]].fillna("").astype(str).str.strip().str.lower()
        nums = pd.to_numeric(s.str.replace("%", "", regex=False), errors="coerce")
        out["_promo_flag"] = (s.isin(["1", "true", "yes", "y", "有", "是", "已配置", "active"]) | (nums.fillna(0) > 0)).astype(int)
    if c.get("material"):
        s = out[c["material"]].fillna("").astype(str).str.strip().str.lower()
        nums = pd.to_numeric(s.str.replace("%", "", regex=False), errors="coerce")
        out["_material_flag"] = (s.isin(["1", "true", "yes", "y", "有", "是", "已配置", "active"]) | (nums.fillna(0) > 0)).astype(int)
    if c.get("visit_record"):
        s = out[c["visit_record"]].fillna("").astype(str).str.strip().str.lower()
        nums = pd.to_numeric(s.str.replace("%", "", regex=False), errors="coerce")
        out["_visit_record_flag"] = (s.isin(["1", "true", "yes", "y", "有", "是", "已拜访", "active"]) | (nums.fillna(0) > 0)).astype(int)

    return out


def weighted_rate(df: pd.DataFrame, rate_col: str, weight_col: str) -> float:
    if len(df) == 0:
        return 0.0
    w = df[weight_col].fillna(0)
    r = df[rate_col].fillna(0)
    return safe_div((r * w).sum(), w.sum())


def summary_for(df: pd.DataFrame) -> Dict[str, float]:
    return {
        "Merchants": len(df),
        "GMV": df["_gmv"].sum(),
        "Orders": df["_orders"].sum(),
        "GMV / Store": safe_div(df["_gmv"].sum(), len(df)),
        "High GMV Stores": int((df["_gmv"] >= df["_gmv"].quantile(0.75)).sum()) if len(df) else 0,
        "Exposure → Visit": weighted_rate(df, "_exposure_visit_rate", "_exposure"),
        "Visit → Cart": weighted_rate(df, "_visit_cart_rate", "_visit"),
        "Cart → Order": weighted_rate(df, "_cart_order_rate", "_cart"),
        "Exposure → Order": weighted_rate(df, "_exposure_order_rate", "_exposure"),
        "Promo Rate": df["_promo_flag"].mean() if len(df) else 0,
        "Material Rate": df["_material_flag"].mean() if len(df) else 0,
        "Visit Record Rate": df["_visit_record_flag"].mean() if len(df) else 0,
    }


def bd_summary(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    p75 = df["_gmv"].quantile(0.75) if len(df) else 0
    for bd, g in df.groupby("_bd_display", dropna=False):
        s = summary_for(g)
        s["BD Name"] = bd
        s["High GMV Stores"] = int((g["_gmv"] >= p75).sum())
        rows.append(s)
    out = pd.DataFrame(rows)
    if len(out) == 0:
        return out
    out = out.sort_values("GMV", ascending=False).reset_index(drop=True)
    out.insert(0, "Rank", range(1, len(out) + 1))
    return out


def merchant_score(df: pd.DataFrame) -> pd.DataFrame:
    x = df.copy()
    # percent-rank style scores, stable and explainable
    def pct_rank(s):
        return s.rank(pct=True).fillna(0)
    x["Health Score"] = (
        30 * pct_rank(x["_gmv"]) +
        20 * pct_rank(x["_exposure_order_rate"]) +
        15 * pct_rank(x["_exposure_visit_rate"]) +
        10 * x["_promo_flag"] +
        10 * x["_material_flag"] +
        5 * x["_visit_record_flag"] +
        10 * pct_rank(x["_orders"])
    ).round(0).clip(0, 100)
    x["Learning Score"] = (
        35 * pct_rank(x["_gmv"]) +
        25 * pct_rank(x["_exposure_order_rate"]) +
        15 * pct_rank(x["_orders"]) +
        10 * x["_promo_flag"] +
        10 * x["_material_flag"] +
        5 * x["_visit_record_flag"]
    ).round(0).clip(0, 100)
    med_gmv = x["_gmv"].median() if len(x) else 0
    x["Opportunity Score"] = (
        30 * pct_rank(x["_exposure"]) +
        20 * (1 - pct_rank(x["_exposure_order_rate"])) +
        15 * (1 - pct_rank(x["_visit_cart_rate"])) +
        10 * (1 - x["_promo_flag"]) +
        10 * (1 - x["_material_flag"]) +
        10 * (1 - x["_visit_record_flag"]) +
        5 * (x["_gmv"] < med_gmv).astype(int)
    ).round(0).clip(0, 100)
    return x


def display_cols(df: pd.DataFrame) -> pd.DataFrame:
    cols = ["_merchant", "_merchant_id", "_bd_display", "_area", "_category", "_gmv", "_orders", "_exposure_order_rate", "_exposure_visit_rate", "_visit_cart_rate", "_cart_order_rate", "Health Score", "Learning Score", "Opportunity Score"]
    avail = [c for c in cols if c in df.columns]
    out = df[avail].copy()
    rename = {
        "_merchant": "Merchant",
        "_merchant_id": "Merchant ID",
        "_bd_display": "BD",
        "_area": "Area",
        "_category": "Category",
        "_gmv": "GMV",
        "_orders": "Orders",
        "_exposure_order_rate": "Exposure → Order",
        "_exposure_visit_rate": "Exposure → Visit",
        "_visit_cart_rate": "Visit → Cart",
        "_cart_order_rate": "Cart → Order",
    }
    return out.rename(columns=rename)


def format_table(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in out.columns:
        if c in ["GMV", "GMV / Store"]:
            out[c] = out[c].map(fmt_money)
        elif c in ["Orders", "Merchants", "High GMV Stores"]:
            out[c] = out[c].map(fmt_int)
        elif "Rate" in c or "→" in c:
            out[c] = out[c].map(fmt_pct)
    return out

# -----------------------------
# UI
# -----------------------------

st.title("Sydney Growth Intelligence V5")
st.caption("Monthly merchant report → learn from top stores → identify scalable actions")

with st.sidebar:
    st.header("Upload")
    uploaded = st.file_uploader("Encrypted Excel file", type=["xlsx", "xls"])
    password = st.text_input("Password", type="password")
    month_input = st.text_input("Reporting month", value="2026-06", help="Example: 2026-06. This will be shown across the report.")
    st.divider()
    st.caption("After upload, choose your BD from the dropdown.")

if not uploaded:
    st.info("Upload your monthly encrypted Excel file, enter password, then analyse.")
    st.stop()

try:
    raw_df = read_excel(uploaded, password)
except Exception as e:
    st.error("Could not open the Excel file. Please check password and file format.")
    st.stop()

cols = detect_columns(raw_df)
df = prep(raw_df, cols)
df = merchant_score(df)

bd_options = sorted(df["_bd_display"].dropna().unique().tolist())
with st.sidebar:
    default_idx = bd_options.index("Yuan Dong") if "Yuan Dong" in bd_options else 0
    selected_bd = st.selectbox("BD / Owner", bd_options, index=default_idx)
    top_n = st.slider("Top N merchants", 10, 100, 20, 5)

my_df = df[df["_bd_display"] == selected_bd].copy()
bd_rank = bd_summary(df)
my_rank_row = bd_rank[bd_rank["BD Name"] == selected_bd]
my_rank = int(my_rank_row["Rank"].iloc[0]) if len(my_rank_row) else None
leader = bd_rank.iloc[0] if len(bd_rank) else None
my_sum = summary_for(my_df)
syd_sum = summary_for(df)

st.success(f"Loaded {len(df):,} rows and {len(df.columns):,} derived/raw columns. Reporting month: {month_input}")

# Executive Summary
st.header("Executive Summary")
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Sydney GMV", fmt_money(syd_sum["GMV"]))
c2.metric("Sydney Orders", fmt_int(syd_sum["Orders"]))
c3.metric("Your Rank", f"#{my_rank}" if my_rank else "N/A")
c4.metric("Your GMV", fmt_money(my_sum["GMV"]))
if leader is not None:
    gap = float(leader["GMV"]) - float(my_sum["GMV"])
    c5.metric("Gap to #1", fmt_money(max(gap, 0)))
else:
    c5.metric("Gap to #1", "N/A")

if leader is not None and selected_bd != leader["BD Name"]:
    st.info(f"{month_input}: {selected_bd} ranks #{my_rank}. Gap to #1 ({leader['BD Name']}) is {fmt_money(max(float(leader['GMV']) - my_sum['GMV'], 0))}. Focus on high-opportunity merchants and learn from high Learning Score merchants.")
else:
    st.info(f"{month_input}: {selected_bd} is currently the top BD by GMV. Focus on defending top merchants and reducing funnel leakage.")

# Visuals
st.subheader("Visual overview")
vc1, vc2 = st.columns(2)
with vc1:
    top_bd_chart = bd_rank.head(8).copy()
    fig = px.bar(top_bd_chart.sort_values("GMV"), x="GMV", y="BD Name", orientation="h", title="Top BD by GMV")
    st.plotly_chart(fig, use_container_width=True)
with vc2:
    area_top = df.groupby("_area")["_gmv"].sum().sort_values(ascending=False).head(10).reset_index()
    area_top.columns = ["Area", "GMV"]
    fig = px.bar(area_top.sort_values("GMV"), x="GMV", y="Area", orientation="h", title="Top areas by GMV")
    st.plotly_chart(fig, use_container_width=True)

# Metric dictionary
with st.expander("Metric Dictionary — how each number is calculated", expanded=False):
    metric_dict = pd.DataFrame([
        ["GMV", "Σ merchant GMV", "Sum", "Final commercial output."],
        ["Orders", "Σ monthly order count", "Sum", "Order volume."],
        ["GMV / Store", "ΣGMV ÷ merchant count", "Weighted ratio", "Quality of merchant portfolio."],
        ["Exposure → Visit", "Σ(rate × exposure) ÷ Σexposure", "Weighted average", "Whether traffic enters the store."],
        ["Visit → Cart", "Σ(rate × visit) ÷ Σvisit", "Weighted average", "Whether menu/price/product drives intent."],
        ["Cart → Order", "Σ(rate × cart) ÷ Σcart", "Weighted average", "Whether checkout/offer/delivery closes orders."],
        ["Exposure → Order", "Σ(rate × exposure) ÷ Σexposure", "Weighted average", "End-to-end traffic conversion."],
        ["Promo Rate", "Stores with promo ÷ total stores", "Store-level share", "Commercial setup coverage."],
        ["Material Rate", "Stores with material ÷ total stores", "Store-level share", "Marketing asset coverage."],
        ["Visit Record Rate", "Stores with visit record ÷ total stores", "Store-level share", "BD touch coverage."],
        ["Learning Score", "GMV + conversion + orders + setup signals", "0–100 score", "Best merchants worth learning from."],
        ["Opportunity Score", "High exposure + low conversion/setup gaps", "0–100 score", "Merchants most worth improving."],
    ], columns=["Metric", "Formula", "Aggregation", "Why it matters"])
    st.dataframe(metric_dict, use_container_width=True, hide_index=True)

# BD Ranking
st.header("BD Ranking — weighted metrics")
st.dataframe(format_table(bd_rank), use_container_width=True, hide_index=True)

# Compare Me
st.header("Compare Me vs Top BD")
if leader is not None and len(my_rank_row):
    compare = pd.DataFrame([leader.to_dict(), my_rank_row.iloc[0].to_dict()])
    st.dataframe(format_table(compare[["BD Name", "Rank", "Merchants", "GMV", "Orders", "GMV / Store", "High GMV Stores", "Exposure → Visit", "Visit → Cart", "Cart → Order", "Exposure → Order", "Promo Rate", "Material Rate", "Visit Record Rate"]]), use_container_width=True, hide_index=True)

# Learn from best
st.header("Learn From Best — which stores are worth copying?")
learn = df.sort_values("Learning Score", ascending=False).head(top_n)
st.dataframe(format_table(display_cols(learn)), use_container_width=True, hide_index=True)

lc1, lc2 = st.columns(2)
with lc1:
    fig = px.scatter(df, x="_exposure_order_rate", y="_gmv", size="_orders", hover_name="_merchant", color="_bd_display", title="Merchant GMV vs exposure-to-order conversion")
    st.plotly_chart(fig, use_container_width=True)
with lc2:
    best_cat = learn.groupby("_category")["Learning Score"].mean().sort_values(ascending=False).head(10).reset_index()
    best_cat.columns = ["Category", "Avg Learning Score"]
    fig = px.bar(best_cat.sort_values("Avg Learning Score"), x="Avg Learning Score", y="Category", orientation="h", title="Categories with strongest learnable stores")
    st.plotly_chart(fig, use_container_width=True)

# Opportunity Finder
st.header("Opportunity Finder — highest ROI stores to improve")
opp = df[df["_bd_display"] == selected_bd].sort_values("Opportunity Score", ascending=False).head(top_n)
st.dataframe(format_table(display_cols(opp)), use_container_width=True, hide_index=True)

st.subheader("Recommended action logic")
actions = []
for _, r in opp.head(15).iterrows():
    reasons = []
    if r["_exposure"] > df["_exposure"].median():
        reasons.append("high exposure")
    if r["_exposure_order_rate"] < df["_exposure_order_rate"].median():
        reasons.append("low exposure→order")
    if r["_visit_cart_rate"] < df["_visit_cart_rate"].median():
        reasons.append("low visit→cart")
    if r["_promo_flag"] == 0:
        reasons.append("missing promo")
    if r["_material_flag"] == 0:
        reasons.append("missing material")
    if r["_visit_record_flag"] == 0:
        reasons.append("needs visit")
    action = "Prioritise merchant visit, review menu ranking, add bundle/promo, and benchmark against similar high Learning Score stores."
    actions.append({
        "Month": month_input,
        "Merchant": r["_merchant"],
        "Merchant ID": r["_merchant_id"],
        "BD": r["_bd_display"],
        "Area": r["_area"],
        "Category": r["_category"],
        "GMV": r["_gmv"],
        "Orders": r["_orders"],
        "Opportunity Score": r["Opportunity Score"],
        "Reason": ", ".join(reasons) if reasons else "portfolio improvement candidate",
        "Action": action,
    })
action_df = pd.DataFrame(actions)
st.dataframe(format_table(action_df), use_container_width=True, hide_index=True)
st.download_button("Download action plan CSV", action_df.to_csv(index=False).encode("utf-8-sig"), file_name=f"action_plan_{selected_bd}_{month_input}.csv", mime="text/csv")

# Area / category intelligence
st.header("Area & Category Intelligence")
tab_area, tab_cat, tab_merch = st.tabs(["Area board", "Category board", "Merchant search"])
with tab_area:
    area_rows = []
    for area, g in df.groupby("_area"):
        s = summary_for(g)
        s["Area"] = area
        area_rows.append(s)
    area_df = pd.DataFrame(area_rows).sort_values("GMV", ascending=False)
    st.dataframe(format_table(area_df[["Area", "Merchants", "GMV", "Orders", "GMV / Store", "Exposure → Visit", "Visit → Cart", "Cart → Order", "Exposure → Order"]].head(50)), use_container_width=True, hide_index=True)
with tab_cat:
    cat_rows = []
    for cat, g in df.groupby("_category"):
        s = summary_for(g)
        s["Category"] = cat
        cat_rows.append(s)
    cat_df = pd.DataFrame(cat_rows).sort_values("GMV", ascending=False)
    st.dataframe(format_table(cat_df[["Category", "Merchants", "GMV", "Orders", "GMV / Store", "Exposure → Visit", "Visit → Cart", "Cart → Order", "Exposure → Order"]].head(50)), use_container_width=True, hide_index=True)
with tab_merch:
    q = st.text_input("Search merchant name / ID")
    search_df = df.copy()
    if q:
        search_df = search_df[search_df["_merchant"].astype(str).str.contains(q, case=False, na=False) | search_df["_merchant_id"].astype(str).str.contains(q, case=False, na=False)]
    st.dataframe(format_table(display_cols(search_df.sort_values("_gmv", ascending=False).head(100))), use_container_width=True, hide_index=True)

# Monthly note
st.header("Monthly Trend — next step")
st.caption("Current version analyses one monthly file at a time and labels it with Reporting Month. V5.1 can add multi-file upload so June/July/August can be compared in one trend chart.")

with st.expander("Detected source columns"):
    st.json(cols)
