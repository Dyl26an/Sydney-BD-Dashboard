import io
import re
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st

try:
    import msoffcrypto
except Exception:
    msoffcrypto = None

st.set_page_config(
    page_title="Sydney Growth Intelligence",
    page_icon="📈",
    layout="wide",
)

# -----------------------------
# Styling
# -----------------------------
st.markdown(
    """
    <style>
    .main .block-container {padding-top: 2rem; padding-bottom: 3rem;}
    .metric-card {background: #111827; border: 1px solid #273244; padding: 16px; border-radius: 14px;}
    .small-muted {color:#9CA3AF; font-size: 0.92rem;}
    .good {color:#22c55e; font-weight:700;}
    .warn {color:#f59e0b; font-weight:700;}
    .bad {color:#ef4444; font-weight:700;}
    .explain-box {background:#0f172a; border:1px solid #334155; padding:14px 16px; border-radius:12px; margin-bottom:10px;}
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------
# Helpers
# -----------------------------
def normalize_col(c: str) -> str:
    return str(c).strip().lower().replace(" ", "").replace("_", "")


def pick_col(df: pd.DataFrame, candidates: List[str], must_not: Optional[List[str]] = None) -> Optional[str]:
    must_not = must_not or []
    cols = list(df.columns)
    norm_map = {c: normalize_col(c) for c in cols}
    cand_norm = [normalize_col(x) for x in candidates]
    not_norm = [normalize_col(x) for x in must_not]

    # exact normalized match first
    for cand in cand_norm:
        for c, n in norm_map.items():
            if n == cand and not any(bad in n for bad in not_norm):
                return c
    # contains match second
    for cand in cand_norm:
        for c, n in norm_map.items():
            if cand in n and not any(bad in n for bad in not_norm):
                return c
    return None


def pick_numeric_col(df: pd.DataFrame, candidates: List[str], must_not: Optional[List[str]] = None) -> Optional[str]:
    c = pick_col(df, candidates, must_not=must_not)
    if c:
        return c
    return None


def to_num(s):
    if s is None:
        return pd.Series(dtype=float)
    return pd.to_numeric(
        pd.Series(s).astype(str).str.replace(r"[^0-9.\-]", "", regex=True).replace("", np.nan),
        errors="coerce",
    ).fillna(0)


def to_rate(s):
    if s is None:
        return pd.Series(dtype=float)
    raw = pd.Series(s).astype(str).str.strip()
    pct_mask = raw.str.contains("%", na=False)
    val = pd.to_numeric(raw.str.replace("%", "", regex=False).str.replace(r"[^0-9.\-]", "", regex=True), errors="coerce").fillna(0)
    val = np.where(pct_mask, val / 100, val)
    val = pd.Series(val)
    # If values look like 0-100 percentages, convert to 0-1
    if val.max() > 1.5 and val.max() <= 100:
        val = val / 100
    return val.fillna(0)


def read_excel(uploaded_file, password: str) -> pd.DataFrame:
    content = uploaded_file.read()
    bio = io.BytesIO(content)

    # Try encrypted first if password provided
    if password and msoffcrypto is not None:
        try:
            office_file = msoffcrypto.OfficeFile(io.BytesIO(content))
            office_file.load_key(password=password)
            decrypted = io.BytesIO()
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            return pd.read_excel(decrypted, sheet_name=0)
        except Exception:
            pass

    # Fallback regular Excel
    bio.seek(0)
    return pd.read_excel(bio, sheet_name=0)


def detect_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    """Detect columns with priority for this merchant performance Excel.
    Important: count columns must be picked before categorical fields such as 必选_订单配送类型.
    Funnel rates should use the existing rate columns where available, because raw exposure/visit/cart/order
    columns can be mixed monthly totals vs daily averages.
    """
    return {
        "bd_name": pick_col(df, ["bd姓名", "bd name", "owner name", "owner", "负责人", "业务员", "姓名"]),
        "bd_id": pick_col(df, ["bd工号", "bd id", "bdid", "工号"]),
        "merchant_name": pick_col(df, ["商户名称", "店铺名称", "门店名称", "merchant name", "store name", "restaurant name", "name"]),
        "merchant_id": pick_col(df, ["店铺id", "商户id", "merchant id", "store id", "poi id", "门店id"]),
        "area": pick_col(df, ["商圈", "区域", "area", "district", "suburb"]),
        "category": pick_col(df, ["店铺二级分类", "品类", "主营品类", "category", "cuisine", "菜系"]),
        "gmv": pick_numeric_col(df, ["gmv", "交易额", "成交额", "销售额", "总gmv"]),
        # This file uses 订单数_排除mm的均单 as the reliable order-count field.
        "orders": pick_numeric_col(df, ["订单数_排除mm的均单", "订单数", "有效订单", "店铺下单人数", "下单人数_排除mm的均单", "orders", "order"], must_not=["配送类型", "必选", "转化率", "率"]),
        "exposure": pick_numeric_col(df, ["平均曝光人数", "曝光人数", "曝光次数", "曝光uv", "曝光", "impression", "exposure"], must_not=["转化率", "率", "高曝光"]),
        "visit": pick_numeric_col(df, ["平均进店人数", "进店人数", "访问人数", "进店", "访问", "visit", "store visit"], must_not=["转化率", "率"]),
        "cart": pick_numeric_col(df, ["平均加购人数", "加购人数", "购物车人数", "加购", "购物车", "cart", "add to cart"], must_not=["转化率", "率"]),
        "exposure_visit_rate": pick_col(df, ["曝光进店转化率", "曝光到进店率", "曝光->进店", "exposure visit"]),
        "exposure_order_rate": pick_col(df, ["曝光下单转化率", "曝光到下单率", "曝光->下单", "exposure order", "曝光下单"]),
        "visit_cart_rate": pick_col(df, ["进店加购转化率", "进店到加购", "visit cart"]),
        "cart_order_rate": pick_col(df, ["加购下单转化率", "加购到下单率", "加购下单", "cart order"]),
        "material": pick_col(df, ["是否有物料", "物料", "material"]),
        "visit_record": pick_col(df, ["拜访记录", "是否拜访", "visit record", "拜访"]),
        "promo": pick_col(df, ["折扣来数", "营销活动", "活动", "优惠券", "promo", "campaign", "discount"]),
    }

def enrich(df: pd.DataFrame, cols: Dict[str, Optional[str]]) -> pd.DataFrame:
    out = df.copy()
    n = len(out)

    def num_col(key):
        c = cols.get(key)
        return to_num(out[c]) if c else pd.Series([0] * n)

    out["_bd_name"] = out[cols["bd_name"]].astype(str).str.strip() if cols.get("bd_name") else "Unknown"
    out["_bd_id"] = out[cols["bd_id"]].astype(str).str.strip() if cols.get("bd_id") else out["_bd_name"]
    out["_bd_display"] = np.where(out["_bd_name"].isin(["", "nan", "None", "Unknown"]), out["_bd_id"], out["_bd_name"])
    out["_merchant_name"] = out[cols["merchant_name"]].astype(str).str.strip() if cols.get("merchant_name") else "Unknown Merchant"
    out["_merchant_id"] = out[cols["merchant_id"]].astype(str).str.strip() if cols.get("merchant_id") else ""
    out["_area"] = out[cols["area"]].astype(str).str.strip() if cols.get("area") else "Unknown"
    out["_category"] = out[cols["category"]].astype(str).str.strip() if cols.get("category") else "Unknown"

    out["_gmv"] = num_col("gmv")
    out["_orders"] = num_col("orders")
    out["_exposure"] = num_col("exposure")
    out["_visit"] = num_col("visit")
    out["_cart"] = num_col("cart")

    # Prefer rate columns from the source file. They are more reliable than recomputing with mixed
    # monthly totals and daily-average people counts.
    if cols.get("exposure_order_rate"):
        out["_exp_order_rate"] = to_rate(out[cols["exposure_order_rate"]])
    else:
        out["_exp_order_rate"] = np.where(out["_exposure"] > 0, out["_orders"] / out["_exposure"], 0)

    if cols.get("exposure_visit_rate"):
        out["_visit_rate"] = to_rate(out[cols["exposure_visit_rate"]])
    else:
        out["_visit_rate"] = np.where(out["_exposure"] > 0, out["_visit"] / out["_exposure"], 0)

    if cols.get("visit_cart_rate"):
        out["_cart_rate"] = to_rate(out[cols["visit_cart_rate"]])
    else:
        out["_cart_rate"] = np.where(out["_visit"] > 0, out["_cart"] / out["_visit"], 0)

    if cols.get("cart_order_rate"):
        out["_cart_order_rate"] = to_rate(out[cols["cart_order_rate"]])
    else:
        out["_cart_order_rate"] = np.where(out["_cart"] > 0, out["_orders"] / out["_cart"], 0)

    def boolish(key):
        c = cols.get(key)
        if not c:
            return pd.Series([False] * n)
        s = out[c]
        if pd.api.types.is_numeric_dtype(s):
            return s.fillna(0) > 0
        txt = s.astype(str).str.lower().str.strip()
        return ~(txt.isin(["", "0", "no", "false", "nan", "none", "无", "否", "没有"]))

    out["_has_material"] = boolish("material")
    out["_has_visit_record"] = boolish("visit_record")
    out["_has_promo"] = boolish("promo")
    return out


def bd_summary(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("_bd_display", dropna=False).agg(
        Merchants=("_merchant_name", "count"),
        GMV=("_gmv", "sum"),
        Orders=("_orders", "sum"),
        Exposure=("_exposure", "sum"),
        Visit=("_visit", "sum"),
        Cart=("_cart", "sum"),
        Promo_Rate=("_has_promo", "mean"),
        Material_Rate=("_has_material", "mean"),
        Visit_Record_Rate=("_has_visit_record", "mean"),
    ).reset_index().rename(columns={"_bd_display": "BD Name"})
    g["GMV / Store"] = np.where(g["Merchants"] > 0, g["GMV"] / g["Merchants"], 0)
    g["Exposure → Visit"] = np.where(g["Exposure"] > 0, g["Visit"] / g["Exposure"], 0)
    g["Visit → Cart"] = np.where(g["Visit"] > 0, g["Cart"] / g["Visit"], 0)
    g["Cart → Order"] = np.where(g["Cart"] > 0, g["Orders"] / g["Cart"], 0)
    g["Exposure → Order"] = np.where(g["Exposure"] > 0, g["Orders"] / g["Exposure"], 0)
    # Count high-GMV stores per BD. Use merge rather than assigning raw .values,
    # because some rows can have blank/NaN BD names and groupby(dropna=False) order/length
    # may not match the summary table exactly on Streamlit Cloud.
    high_cutoff = df["_gmv"].quantile(0.75) if len(df) else 0
    high_counts = (
        df.assign(_high=df["_gmv"] >= high_cutoff)
        .groupby("_bd_display", dropna=False)["_high"]
        .sum()
        .reset_index()
        .rename(columns={"_bd_display": "BD Name", "_high": "High GMV Stores"})
    )
    g = g.merge(high_counts, on="BD Name", how="left")
    g["High GMV Stores"] = g["High GMV Stores"].fillna(0).astype(int)
    g = g.sort_values("GMV", ascending=False).reset_index(drop=True)
    g.insert(0, "Rank", np.arange(1, len(g) + 1))
    return g


def area_summary(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("_area", dropna=False).agg(
        Merchants=("_merchant_name", "count"),
        GMV=("_gmv", "sum"),
        Orders=("_orders", "sum"),
        Exposure=("_exposure", "sum"),
        Visit=("_visit", "sum"),
        Cart=("_cart", "sum"),
    ).reset_index().rename(columns={"_area": "Area"})
    g["GMV / Store"] = np.where(g["Merchants"] > 0, g["GMV"] / g["Merchants"], 0)
    g["Exposure → Order"] = np.where(g["Exposure"] > 0, g["Orders"] / g["Exposure"], 0)
    return g.sort_values("GMV", ascending=False)


def make_action_plan(df: pd.DataFrame, selected_bd: str, limit: int = 30) -> pd.DataFrame:
    d = df[df["_bd_display"] == selected_bd].copy() if selected_bd != "All Sydney" else df.copy()
    if d.empty:
        return pd.DataFrame()

    # Score opportunity: high exposure/gmv potential + weak conversion + missing promo/material/visit record
    exp_norm = d["_exposure"] / max(d["_exposure"].quantile(0.95), 1)
    gmv_norm = d["_gmv"] / max(d["_gmv"].quantile(0.95), 1)
    conv_gap = np.maximum(0, df["_exp_order_rate"].median() - d["_exp_order_rate"])
    conv_norm = conv_gap / max(df["_exp_order_rate"].median(), 0.0001)
    missing = (~d["_has_promo"]).astype(int) * 0.18 + (~d["_has_material"]).astype(int) * 0.12 + (~d["_has_visit_record"]).astype(int) * 0.12
    d["Opportunity Score"] = (exp_norm.clip(0, 1) * 35 + gmv_norm.clip(0, 1) * 25 + conv_norm.clip(0, 2) * 20 + missing * 100).round(1)

    reasons = []
    actions = []
    for _, r in d.iterrows():
        rs, ac = [], []
        if r["_exposure"] > df["_exposure"].quantile(0.7) and r["_exp_order_rate"] < df["_exp_order_rate"].median():
            rs.append("High exposure but weak order conversion")
            ac.append("Review menu first screen, hero image, combo structure and delivery fee incentive")
        if not r["_has_promo"]:
            rs.append("No/weak promotion setup")
            ac.append("Add discount, bundle or new-customer offer")
        if not r["_has_material"]:
            rs.append("No material")
            ac.append("Add in-store/platform material to improve visibility and trust")
        if not r["_has_visit_record"]:
            rs.append("No recent visit record")
            ac.append("Schedule merchant visit and collect menu/campaign blockers")
        if r["_gmv"] > df["_gmv"].quantile(0.75):
            rs.append("High-value merchant")
            ac.append("Protect relationship and upsell deeper campaign package")
        reasons.append("; ".join(rs) if rs else "Stable merchant; monitor performance")
        actions.append("; ".join(dict.fromkeys(ac)) if ac else "Monitor weekly; no urgent action")

    d["Opportunity Reason"] = reasons
    d["Recommended Action"] = actions
    cols = ["_merchant_name", "_merchant_id", "_bd_display", "_area", "_category", "_gmv", "_orders", "_exposure", "_visit", "_cart", "_exp_order_rate", "_has_promo", "_has_material", "_has_visit_record", "Opportunity Score", "Opportunity Reason", "Recommended Action"]
    out = d[cols].rename(columns={
        "_merchant_name": "Merchant Name",
        "_merchant_id": "Merchant ID",
        "_bd_display": "BD Name",
        "_area": "Area",
        "_category": "Category",
        "_gmv": "GMV",
        "_orders": "Orders",
        "_exposure": "Exposure",
        "_visit": "Visit",
        "_cart": "Cart",
        "_exp_order_rate": "Exposure → Order",
        "_has_promo": "Has Promo",
        "_has_material": "Has Material",
        "_has_visit_record": "Has Visit Record",
    })
    return out.sort_values("Opportunity Score", ascending=False).head(limit)


def fmt_money(x):
    return f"${x:,.0f}"


def fmt_pct(x):
    return f"{x:.2%}"


def show_metric(label, value, help_text=None):
    st.metric(label, value, help=help_text)


def display_df(df: pd.DataFrame, money_cols=None, pct_cols=None):
    money_cols = money_cols or []
    pct_cols = pct_cols or []
    fmt = {}
    for c in money_cols:
        if c in df.columns:
            fmt[c] = "${:,.0f}"
    for c in pct_cols:
        if c in df.columns:
            fmt[c] = "{:.2%}"
    st.dataframe(df.style.format(fmt), use_container_width=True, hide_index=True)


FIELD_EXPLAIN = {
    "GMV": "Gross Merchandise Value，总成交额。老板最关注的结果指标，但不能单独判断BD能力，因为它受商圈、店铺质量和历史积累影响。",
    "Orders": "订单数。反映真实成交量。和GMV一起看，可以判断是客单价高还是订单量强。",
    "GMV / Store": "单店GMV。非常关键。它能判断BD是不是在经营高质量店铺，而不是只靠店铺数量堆结果。",
    "Exposure → Visit": "曝光到进店率。用户看到店铺后是否愿意点进去。受店铺首图、评分、配送费、店名、菜系吸引力影响。",
    "Visit → Cart": "进店到加购率。用户进店后是否愿意把商品加入购物车。受菜单结构、图片、价格、爆品、套餐影响。",
    "Cart → Order": "加购到下单率。用户加购后是否最终支付。受配送费、满减、优惠券、起送价、结算体验影响。",
    "Exposure → Order": "曝光到下单转化率。最完整的转化效率指标，衡量从看到店铺到最终下单的整体效率。",
    "Promo_Rate": "活动覆盖率。这里指店铺是否有促销/折扣/优惠配置。高不一定代表健康，但低通常说明还有运营空间。",
    "Material_Rate": "物料覆盖率。指是否有物料/展示资源。它不是最终结果，但能帮助提升可见度和信任感。",
    "Visit_Record_Rate": "拜访记录覆盖率。指BD是否有拜访记录。高说明触达频率更好，但也要和GMV提升结合看。",
    "High GMV Stores": "高GMV店铺数量。判断BD是否拥有和培养足够多的高价值商户。",
}

# -----------------------------
# Sidebar
# -----------------------------
st.sidebar.title("Upload")
uploaded = st.sidebar.file_uploader("Encrypted Excel file", type=["xlsx", "xls"])
password = st.sidebar.text_input("Password", type="password")

st.title("Sydney Growth Intelligence V3")
st.caption("Upload encrypted Excel → compare BD performance → learn from top BD → generate practical action plan")

if not uploaded:
    st.info("Upload your monthly encrypted Excel, enter password, then click Analyse.")
    st.subheader("What V3 focuses on")
    st.markdown(
        """
        - Explain what each KPI means and whether it is business-critical
        - Compare you against Top BD by name, not ID
        - Show merchant names in opportunity lists
        - Find why stronger BDs are stronger
        - Turn data into batch learning and action plan
        """
    )
    st.stop()

if st.sidebar.button("Analyse", type="primary"):
    st.session_state["run"] = True

if not st.session_state.get("run"):
    st.info("Click Analyse in the sidebar to start.")
    st.stop()

try:
    raw = read_excel(uploaded, password)
except Exception as e:
    st.error(f"Could not read the Excel file. Check password or file format. Error: {e}")
    st.stop()

cols = detect_columns(raw)
df = enrich(raw, cols)
bd_names = sorted([x for x in df["_bd_display"].dropna().unique().tolist() if str(x).strip() and str(x) != "nan"])
selected_bd = st.sidebar.selectbox("BD / Owner", bd_names, index=bd_names.index("Yuan Dong") if "Yuan Dong" in bd_names else 0)

st.success(f"Loaded {len(df):,} rows and {len(df.columns):,} derived/raw columns.")

bd_rank = bd_summary(df)
areas = area_summary(df)
my_row = bd_rank[bd_rank["BD Name"] == selected_bd]
top_row = bd_rank.iloc[0] if not bd_rank.empty else None
my_data = df[df["_bd_display"] == selected_bd]

# -----------------------------
# Executive summary
# -----------------------------
st.header("Executive Summary")
c1, c2, c3, c4 = st.columns(4)
with c1:
    show_metric("Sydney GMV", fmt_money(df["_gmv"].sum()), FIELD_EXPLAIN["GMV"])
with c2:
    show_metric("Sydney Orders", f"{df['_orders'].sum():,.0f}", FIELD_EXPLAIN["Orders"])
with c3:
    show_metric("Merchants", f"{len(df):,}")
with c4:
    show_metric("BD Count", f"{len(bd_names):,}")

if not my_row.empty:
    r = my_row.iloc[0]
    gap = float(top_row["GMV"] - r["GMV"]) if top_row is not None else 0
    st.subheader(f"My Performance — {selected_bd}")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        show_metric("Rank", f"#{int(r['Rank'])}")
    with c2:
        show_metric("My GMV", fmt_money(r["GMV"]))
    with c3:
        show_metric("My Merchants", f"{int(r['Merchants']):,}")
    with c4:
        show_metric("GMV / Store", fmt_money(r["GMV / Store"]), FIELD_EXPLAIN["GMV / Store"])
    with c5:
        show_metric("Gap to #1", fmt_money(max(gap, 0)))

    if gap > 0:
        st.info(f"You are {fmt_money(gap)} behind #{int(top_row['Rank'])} {top_row['BD Name']}. The fastest way to close this gap is not more merchants first; it is lifting GMV/store and converting high-exposure merchants.")
    else:
        st.success("You are currently #1 by GMV. Next focus: protect top merchants and build repeatable playbook for the team.")
else:
    st.warning("Selected BD not found after data cleaning.")

# -----------------------------
# Tabs
# -----------------------------
tabs = st.tabs([
    "📘 KPI Dictionary",
    "🏆 BD Ranking",
    "🔍 Compare Me",
    "🏪 Merchant Intelligence",
    "📍 Area Intelligence",
    "🎯 Opportunity Finder",
    "🧠 Learn From Best",
])

with tabs[0]:
    st.subheader("KPI Dictionary — what each header means")
    for k, v in FIELD_EXPLAIN.items():
        st.markdown(f"<div class='explain-box'><b>{k}</b><br><span class='small-muted'>{v}</span></div>", unsafe_allow_html=True)
    st.markdown("### Are these the most decisive BD metrics?")
    st.write("The most decisive metrics are not just Promo_Rate or Visit_Record_Rate. For BD performance, focus on this order:")
    st.markdown(
        """
        1. **GMV / Store** — tells whether you manage quality merchants, not just many merchants.  
        2. **High GMV Stores** — tells whether you have enough strong merchants.  
        3. **Exposure → Order** — tells whether traffic is being converted efficiently.  
        4. **Funnel gaps**: Exposure → Visit, Visit → Cart, Cart → Order — tells where the merchant is leaking customers.  
        5. **Promo / Material / Visit Record** — these are action levers, not final outcomes. They matter because they explain what you can improve.
        """
    )
    st.markdown("### Detected columns")
    st.json(cols)

with tabs[1]:
    st.subheader("BD Ranking — Name Display")
with st.expander("字段解释 / What these columns mean", expanded=False):
    st.markdown("""
    - **Orders**：订单数。本版本优先读取 `订单数_排除mm的均单`，不再误抓 `必选_订单配送类型`。
    - **Exposure → Visit**：曝光进店转化率。衡量用户看到店铺后是否愿意点进店。
    - **Visit → Cart**：进店加购转化率。衡量菜单、价格、图片、套餐是否有吸引力。
    - **Cart → Order**：加购下单转化率。衡量用户是否在最后一步流失，常和配送费、优惠、起送价有关。
    - **Exposure → Order**：曝光下单转化率，最综合的转化结果指标。
    - **Promo_Rate**：有活动/折扣配置的店铺比例。
    - **Material_Rate**：有物料记录的店铺比例；如果源数据全空，这里会显示 0%。
    - **Visit_Record_Rate**：有拜访记录的店铺比例。
    """)

    rank_view = bd_rank[["Rank", "BD Name", "Merchants", "GMV", "Orders", "GMV / Store", "High GMV Stores", "Exposure → Visit", "Visit → Cart", "Cart → Order", "Exposure → Order", "Promo_Rate", "Material_Rate", "Visit_Record_Rate"]].copy()
    display_df(rank_view, money_cols=["GMV", "GMV / Store"], pct_cols=["Exposure → Visit", "Visit → Cart", "Cart → Order", "Exposure → Order", "Promo_Rate", "Material_Rate", "Visit_Record_Rate"])

with tabs[2]:
    st.subheader(f"Compare Me — {selected_bd} vs Top BD")
    if not my_row.empty and top_row is not None:
        me = my_row.iloc[0]
        compare = pd.DataFrame({
            "KPI": ["GMV", "Merchants", "GMV / Store", "High GMV Stores", "Exposure → Visit", "Visit → Cart", "Cart → Order", "Exposure → Order", "Promo_Rate", "Material_Rate", "Visit_Record_Rate"],
            selected_bd: [me["GMV"], me["Merchants"], me["GMV / Store"], me["High GMV Stores"], me["Exposure → Visit"], me["Visit → Cart"], me["Cart → Order"], me["Exposure → Order"], me["Promo_Rate"], me["Material_Rate"], me["Visit_Record_Rate"]],
            str(top_row["BD Name"]): [top_row["GMV"], top_row["Merchants"], top_row["GMV / Store"], top_row["High GMV Stores"], top_row["Exposure → Visit"], top_row["Visit → Cart"], top_row["Cart → Order"], top_row["Exposure → Order"], top_row["Promo_Rate"], top_row["Material_Rate"], top_row["Visit_Record_Rate"]],
        })
        compare["Gap"] = compare[str(top_row["BD Name"])] - compare[selected_bd]
        st.dataframe(compare, use_container_width=True, hide_index=True)

        st.markdown("### Diagnosis")
        gaps = []
        if me["GMV / Store"] < top_row["GMV / Store"]:
            gaps.append(f"Your GMV/store is lower by {fmt_money(top_row['GMV / Store'] - me['GMV / Store'])}. This means the key gap is merchant quality or merchant monetisation, not only merchant count.")
        if me["High GMV Stores"] < top_row["High GMV Stores"]:
            gaps.append(f"You have {int(top_row['High GMV Stores'] - me['High GMV Stores'])} fewer high-GMV merchants. Build more mid-tier stores into top-tier stores.")
        if me["Exposure → Order"] < top_row["Exposure → Order"]:
            gaps.append("Your exposure-to-order conversion is weaker than the top BD. Prioritise high-exposure low-conversion stores.")
        if me["Visit_Record_Rate"] < top_row["Visit_Record_Rate"]:
            gaps.append("Your visit record coverage is lower. Increase structured visits for high-potential merchants, not random visits.")
        if not gaps:
            gaps.append("You are close to or ahead of the top BD on major operating indicators. Focus on protecting key merchants and replicating your playbook.")
        for g in gaps:
            st.write("- " + g)

with tabs[3]:
    st.subheader("Merchant Intelligence — with merchant names")
    merchant_view = df[df["_bd_display"] == selected_bd].copy()
    q = st.text_input("Search merchant name / area / category", "")
    if q:
        mask = merchant_view["_merchant_name"].str.contains(q, case=False, na=False) | merchant_view["_area"].str.contains(q, case=False, na=False) | merchant_view["_category"].str.contains(q, case=False, na=False)
        merchant_view = merchant_view[mask]
    merchant_out = merchant_view[["_merchant_name", "_merchant_id", "_area", "_category", "_gmv", "_orders", "_exposure", "_visit", "_cart", "_exp_order_rate", "_has_promo", "_has_material", "_has_visit_record"]].rename(columns={
        "_merchant_name": "Merchant Name", "_merchant_id": "Merchant ID", "_area": "Area", "_category": "Category", "_gmv": "GMV", "_orders": "Orders", "_exposure": "Exposure", "_visit": "Visit", "_cart": "Cart", "_exp_order_rate": "Exposure → Order", "_has_promo": "Has Promo", "_has_material": "Has Material", "_has_visit_record": "Has Visit Record"
    }).sort_values("GMV", ascending=False)
    display_df(merchant_out.head(200), money_cols=["GMV"], pct_cols=["Exposure → Order"])

with tabs[4]:
    st.subheader("Area Intelligence")
    display_df(areas.head(30), money_cols=["GMV", "GMV / Store"], pct_cols=["Exposure → Order"])
    st.bar_chart(areas.set_index("Area")["GMV"].head(15))

with tabs[5]:
    st.subheader("Opportunity Finder — practical visit/action list")
    action = make_action_plan(df, selected_bd, limit=50)
    if action.empty:
        st.warning("No action plan generated for this BD.")
    else:
        display_df(action, money_cols=["GMV"], pct_cols=["Exposure → Order"])
        csv = action.to_csv(index=False).encode("utf-8-sig")
        st.download_button("Download Action Plan CSV", csv, file_name=f"{selected_bd}_action_plan.csv", mime="text/csv")

with tabs[6]:
    st.subheader("Learn From Best")
    if top_row is not None:
        top_bd = top_row["BD Name"]
        st.markdown(f"### Top BD: {top_bd}")
        top_merchants = df[df["_bd_display"] == top_bd].sort_values("_gmv", ascending=False).head(30)
        top_out = top_merchants[["_merchant_name", "_area", "_category", "_gmv", "_orders", "_exp_order_rate", "_has_promo", "_has_material", "_has_visit_record"]].rename(columns={
            "_merchant_name": "Merchant Name", "_area": "Area", "_category": "Category", "_gmv": "GMV", "_orders": "Orders", "_exp_order_rate": "Exposure → Order", "_has_promo": "Has Promo", "_has_material": "Has Material", "_has_visit_record": "Has Visit Record"
        })
        display_df(top_out, money_cols=["GMV"], pct_cols=["Exposure → Order"])

        st.markdown("### What to copy in batch")
        common_area = top_merchants["_area"].value_counts().head(3).index.tolist()
        common_cat = top_merchants["_category"].value_counts().head(3).index.tolist()
        promo_rate = top_merchants["_has_promo"].mean()
        visit_rate = top_merchants["_has_visit_record"].mean()
        mat_rate = top_merchants["_has_material"].mean()
        st.write(f"- Top BD's strongest areas: {', '.join(map(str, common_area)) if common_area else 'N/A'}")
        st.write(f"- Top BD's strongest categories: {', '.join(map(str, common_cat)) if common_cat else 'N/A'}")
        st.write(f"- Promo coverage among top merchants: {promo_rate:.1%}")
        st.write(f"- Material coverage among top merchants: {mat_rate:.1%}")
        st.write(f"- Visit record coverage among top merchants: {visit_rate:.1%}")
        st.info("Batch learning idea: find your merchants in the same area/category with lower GMV/store or lower conversion, then copy the top merchants' campaign and menu structure first.")

# Download full enriched dataset
st.divider()
st.subheader("Download")
out_csv = df.to_csv(index=False).encode("utf-8-sig")
st.download_button("Download enriched data CSV", out_csv, file_name="sydney_growth_enriched_data.csv", mime="text/csv")
