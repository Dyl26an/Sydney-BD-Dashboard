import io
import re
from typing import Dict, Optional, List, Tuple

import pandas as pd
import streamlit as st

try:
    import msoffcrypto
except Exception:
    msoffcrypto = None

st.set_page_config(page_title="Sydney Growth Intelligence V4", layout="wide")

# ----------------------------
# Helpers
# ----------------------------
def norm_text(x: str) -> str:
    return re.sub(r"\s+", "", str(x).lower())


def read_excel(uploaded_file, password: str | None) -> pd.DataFrame:
    data = uploaded_file.read()
    bio = io.BytesIO(data)
    # Try encrypted first if password exists
    if password and msoffcrypto is not None:
        try:
            office_file = msoffcrypto.OfficeFile(bio)
            office_file.load_key(password=password)
            decrypted = io.BytesIO()
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            return pd.read_excel(decrypted, sheet_name=0)
        except Exception:
            pass
    # Fallback normal Excel
    bio.seek(0)
    return pd.read_excel(bio, sheet_name=0)


def find_col(cols: List[str], include_any: List[str], exclude_any: List[str] = None, prefer_exact: List[str] = None) -> Optional[str]:
    exclude_any = exclude_any or []
    prefer_exact = prefer_exact or []
    norm_cols = {c: norm_text(c) for c in cols}
    for exact in prefer_exact:
        for c in cols:
            if norm_cols[c] == norm_text(exact):
                return c
    candidates = []
    for c, nc in norm_cols.items():
        if any(norm_text(k) in nc for k in include_any) and not any(norm_text(k) in nc for k in exclude_any):
            candidates.append(c)
    return candidates[0] if candidates else None


def detect_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    cols = list(df.columns)
    return {
        "bd_name": find_col(cols, ["bd姓名", "bdname", "bd名称", "姓名", "负责人姓名"], prefer_exact=["bd姓名", "BD姓名", "BD Name"]),
        "bd_id": find_col(cols, ["bd工号", "bdid", "工号", "负责人"], exclude_any=["姓名"], prefer_exact=["bd工号", "BD工号"]),
        "merchant_name": find_col(cols, ["商户名称", "店铺名称", "门店名称", "merchantname", "restaurantname", "name"], prefer_exact=["商户名称", "店铺名称", "merchant_name"]),
        "merchant_id": find_col(cols, ["店铺id", "商户id", "merchantid", "restaurantid"], prefer_exact=["店铺id", "商户id"]),
        "area": find_col(cols, ["商圈", "区域", "area", "suburb"], prefer_exact=["商圈"]),
        "category": find_col(cols, ["品类", "菜系", "category", "cuisine"], prefer_exact=["品类"]),
        "gmv": find_col(cols, ["gmv", "销售额", "交易额"], prefer_exact=["gmv", "GMV"]),
        # Strict priority: real order count. Avoid delivery type / ratio fields.
        "orders": find_col(cols, ["订单数_排除mm的均单", "订单数", "ordercount", "orders"], exclude_any=["转化率", "配送类型", "必选", "均单价", "aov", "类型"], prefer_exact=["订单数_排除mm的均单", "订单数", "orders"]),
        "exposure": find_col(cols, ["曝光人数", "曝光用户", "曝光uv", "曝光", "impression", "exposure"], exclude_any=["转化率", "rate", "%"], prefer_exact=["曝光人数", "曝光"]),
        "visit": find_col(cols, ["进店人数", "进店用户", "进店uv", "进店", "visit"], exclude_any=["转化率", "rate", "%"], prefer_exact=["进店人数", "进店"]),
        "cart": find_col(cols, ["加购人数", "加购用户", "加购uv", "加购", "cart"], exclude_any=["转化率", "rate", "%"], prefer_exact=["加购人数", "加购"]),
        "promo": find_col(cols, ["折扣", "活动", "优惠", "promo", "campaign"], exclude_any=["来数", "gmv"], prefer_exact=["折扣", "活动"]),
        "material": find_col(cols, ["是否有物料", "物料", "material"], prefer_exact=["是否有物料"]),
        "visit_record": find_col(cols, ["拜访记录", "是否拜访", "visitrecord", "visit_record"], prefer_exact=["拜访记录"]),
        # Source rate columns only for audit, not used for group weighted metrics
        "src_e2o_rate": find_col(cols, ["曝光下单转化率"], prefer_exact=["曝光下单转化率"]),
        "src_e2v_rate": find_col(cols, ["曝光进店转化率"], prefer_exact=["曝光进店转化率"]),
        "src_v2c_rate": find_col(cols, ["进店加购转化率"], prefer_exact=["进店加购转化率"]),
        "src_c2o_rate": find_col(cols, ["加购下单转化率"], prefer_exact=["加购下单转化率"]),
    }


def to_num(s: pd.Series) -> pd.Series:
    if s is None:
        return pd.Series(dtype=float)
    if s.dtype == object:
        return pd.to_numeric(s.astype(str).str.replace("%", "", regex=False).str.replace(",", "", regex=False), errors="coerce")
    return pd.to_numeric(s, errors="coerce")


def yes_rate(series: pd.Series) -> float:
    if series is None or len(series) == 0:
        return 0.0
    s = series.fillna("").astype(str).str.strip().str.lower()
    yes = s.isin(["1", "yes", "y", "true", "有", "是", "已配置", "configured", "active"])
    # If numeric 0/1
    num = pd.to_numeric(series, errors="coerce")
    if num.notna().sum() > len(series) * 0.5:
        return float((num.fillna(0) > 0).mean())
    return float(yes.mean())


def safe_div(a: float, b: float) -> float:
    return float(a) / float(b) if b and pd.notna(b) and float(b) != 0 else 0.0


def prepare(df: pd.DataFrame, col: Dict[str, Optional[str]]) -> pd.DataFrame:
    out = df.copy()
    out["_bd_display"] = "Unknown BD"
    if col.get("bd_name") and col.get("bd_id"):
        name = out[col["bd_name"]].fillna("").astype(str).str.strip()
        bid = out[col["bd_id"]].fillna("").astype(str).str.strip()
        out["_bd_display"] = name.where(name.ne("") & name.ne("nan"), bid)
    elif col.get("bd_name"):
        out["_bd_display"] = out[col["bd_name"]].fillna("Unknown BD").astype(str).str.strip()
    elif col.get("bd_id"):
        out["_bd_display"] = out[col["bd_id"]].fillna("Unknown BD").astype(str).str.strip()

    out["_merchant_name"] = out[col["merchant_name"]].fillna("").astype(str) if col.get("merchant_name") else ""
    out["_merchant_id"] = out[col["merchant_id"]].fillna("").astype(str) if col.get("merchant_id") else ""
    out["_area"] = out[col["area"]].fillna("Unknown").astype(str) if col.get("area") else "Unknown"
    out["_category"] = out[col["category"]].fillna("Unknown").astype(str) if col.get("category") else "Unknown"
    for key in ["gmv", "orders", "exposure", "visit", "cart"]:
        source = col.get(key)
        out[f"_{key}"] = to_num(out[source]).fillna(0) if source else 0
    return out


def group_metrics(g: pd.DataFrame) -> pd.Series:
    merchants = len(g)
    gmv = g["_gmv"].sum()
    orders = g["_orders"].sum()
    exposure = g["_exposure"].sum()
    visit = g["_visit"].sum()
    cart = g["_cart"].sum()
    p75 = st.session_state.get("global_gmv_p75", g["_gmv"].quantile(0.75))
    return pd.Series({
        "Merchants": merchants,
        "GMV": gmv,
        "Orders": orders,
        "GMV / Store": safe_div(gmv, merchants),
        "High GMV Stores": int((g["_gmv"] >= p75).sum()),
        "Exposure": exposure,
        "Visit": visit,
        "Cart": cart,
        "Exposure → Visit": safe_div(visit, exposure),
        "Visit → Cart": safe_div(cart, visit),
        "Cart → Order": safe_div(orders, cart),
        "Exposure → Order": safe_div(orders, exposure),
    })


def add_coverage_metrics(rank: pd.DataFrame, df: pd.DataFrame, col: Dict[str, Optional[str]]) -> pd.DataFrame:
    result = rank.copy()
    for label, source_key in [("Promo Rate", "promo"), ("Material Rate", "material"), ("Visit Record Rate", "visit_record")]:
        source = col.get(source_key)
        if source:
            coverage = df.groupby("_bd_display")[source].apply(yes_rate)
            result[label] = result["BD Name"].map(coverage).fillna(0)
        else:
            result[label] = 0.0
    return result


def fmt_money(x):
    return f"${x:,.0f}"


def fmt_pct(x):
    return f"{x:.2%}"


def style_metrics(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in ["GMV", "GMV / Store"]:
        if c in out: out[c] = out[c].map(fmt_money)
    for c in ["Exposure → Visit", "Visit → Cart", "Cart → Order", "Exposure → Order", "Promo Rate", "Material Rate", "Visit Record Rate"]:
        if c in out: out[c] = out[c].map(fmt_pct)
    for c in ["Orders", "Exposure", "Visit", "Cart"]:
        if c in out: out[c] = out[c].map(lambda v: f"{v:,.0f}")
    return out


METRIC_DICT = pd.DataFrame([
    ["GMV", "Σ GMV", "汇总", "最终经营结果，但不能单独判断BD能力。"],
    ["Orders", "Σ 订单数", "汇总", "成交量。必须使用订单数原始列，不用配送类型或转化率列。"],
    ["GMV / Store", "Σ GMV ÷ 店铺数", "加权/汇总后计算", "衡量店铺质量和经营深度。"],
    ["Exposure → Visit", "Σ 进店人数 ÷ Σ 曝光人数", "加权平均", "衡量曝光是否能吸引用户点进店铺。"],
    ["Visit → Cart", "Σ 加购人数 ÷ Σ 进店人数", "加权平均", "衡量菜单、图片、价格是否让用户产生购买意愿。"],
    ["Cart → Order", "Σ 订单数 ÷ Σ 加购人数", "加权平均", "衡量临门一脚，受价格、配送费、优惠券影响大。"],
    ["Exposure → Order", "Σ 订单数 ÷ Σ 曝光人数", "加权平均", "最终转化效率。不是每家店转化率的算术平均。"],
    ["Promo Rate", "有活动店铺数 ÷ 总店铺数", "店铺覆盖率", "衡量BD推动活动配置能力。"],
    ["Material Rate", "有物料店铺数 ÷ 总店铺数", "店铺覆盖率", "衡量物料覆盖，不是GMV指标。"],
    ["Visit Record Rate", "有拜访记录店铺数 ÷ 总店铺数", "店铺覆盖率", "衡量拜访覆盖度。"],
], columns=["Metric", "Formula", "Method", "Why it matters"])


# ----------------------------
# UI
# ----------------------------
st.title("Sydney Growth Intelligence V4")
st.caption("Metric dictionary first: weighted funnel metrics, merchant names, and explainable BD comparison.")

with st.sidebar:
    st.header("Upload")
    uploaded = st.file_uploader("Encrypted Excel file", type=["xlsx", "xls"])
    password = st.text_input("Password", type="password")
    st.caption("V4 does not store your raw Excel. Data is processed only during this session.")

if not uploaded:
    st.info("Upload the monthly encrypted Excel, enter password, then view V4 analysis.")
    st.subheader("What changed in V4")
    st.dataframe(METRIC_DICT, use_container_width=True, hide_index=True)
    st.stop()

try:
    raw = read_excel(uploaded, password)
except Exception as e:
    st.error("Could not open the Excel file. Check the password and file format.")
    st.stop()

cols = detect_columns(raw)
df = prepare(raw, cols)
st.session_state["global_gmv_p75"] = df["_gmv"].quantile(0.75) if len(df) else 0

bd_options = sorted([x for x in df["_bd_display"].dropna().unique().tolist() if str(x).strip()])
with st.sidebar:
    selected_bd = st.selectbox("BD / Owner", bd_options, index=bd_options.index("Yuan Dong") if "Yuan Dong" in bd_options else 0)
    top_n = st.slider("Top merchants", 5, 50, 20)

st.success(f"Loaded {len(df):,} rows and {len(raw.columns):,} raw columns.")

with st.expander("Detected columns / 字段识别结果", expanded=False):
    st.json(cols)

with st.expander("Metric Dictionary / 指标口径说明", expanded=True):
    st.dataframe(METRIC_DICT, use_container_width=True, hide_index=True)

# Overall metrics
sydney_metrics = group_metrics(df)
my_df = df[df["_bd_display"] == selected_bd].copy()
my_metrics = group_metrics(my_df) if len(my_df) else group_metrics(df.iloc[0:0].copy())

bd_rank = df.groupby("_bd_display", dropna=False).apply(group_metrics, include_groups=False).reset_index().rename(columns={"_bd_display": "BD Name"})
bd_rank = add_coverage_metrics(bd_rank, df, cols)
bd_rank = bd_rank.sort_values("GMV", ascending=False).reset_index(drop=True)
bd_rank.insert(0, "Rank", range(1, len(bd_rank)+1))
my_rank = int(bd_rank.loc[bd_rank["BD Name"] == selected_bd, "Rank"].iloc[0]) if selected_bd in bd_rank["BD Name"].values else None

st.header("Executive Summary")
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Sydney GMV", fmt_money(sydney_metrics["GMV"]))
c2.metric("Sydney Orders", f"{sydney_metrics['Orders']:,.0f}")
c3.metric("Merchants", f"{sydney_metrics['Merchants']:,.0f}")
c4.metric(f"{selected_bd} Rank", f"#{my_rank}" if my_rank else "N/A")
c5.metric("Gap to #1", fmt_money(max(bd_rank.iloc[0]["GMV"] - my_metrics["GMV"], 0)) if len(bd_rank) else "$0")

st.header(f"My Performance — {selected_bd}")
c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("GMV", fmt_money(my_metrics["GMV"]))
c2.metric("Orders", f"{my_metrics['Orders']:,.0f}")
c3.metric("Merchants", f"{my_metrics['Merchants']:,.0f}")
c4.metric("GMV / Store", fmt_money(my_metrics["GMV / Store"]))
c5.metric("Exposure → Order", fmt_pct(my_metrics["Exposure → Order"]))
c6.metric("High GMV Stores", f"{my_metrics['High GMV Stores']:,.0f}")

st.header("BD Ranking — weighted metrics")
rank_display_cols = ["Rank", "BD Name", "Merchants", "GMV", "Orders", "GMV / Store", "High GMV Stores", "Exposure → Visit", "Visit → Cart", "Cart → Order", "Exposure → Order", "Promo Rate", "Material Rate", "Visit Record Rate"]
st.dataframe(style_metrics(bd_rank[rank_display_cols]), use_container_width=True, hide_index=True)

st.header("Compare Me vs Top BD")
if len(bd_rank) > 0:
    top_bd = bd_rank.iloc[0]
    me = bd_rank[bd_rank["BD Name"] == selected_bd].iloc[0]
    compare = pd.DataFrame({
        "Metric": ["GMV", "Orders", "Merchants", "GMV / Store", "High GMV Stores", "Exposure → Visit", "Visit → Cart", "Cart → Order", "Exposure → Order", "Promo Rate", "Visit Record Rate"],
        selected_bd: [me["GMV"], me["Orders"], me["Merchants"], me["GMV / Store"], me["High GMV Stores"], me["Exposure → Visit"], me["Visit → Cart"], me["Cart → Order"], me["Exposure → Order"], me["Promo Rate"], me["Visit Record Rate"]],
        f"Top BD: {top_bd['BD Name']}": [top_bd["GMV"], top_bd["Orders"], top_bd["Merchants"], top_bd["GMV / Store"], top_bd["High GMV Stores"], top_bd["Exposure → Visit"], top_bd["Visit → Cart"], top_bd["Cart → Order"], top_bd["Exposure → Order"], top_bd["Promo Rate"], top_bd["Visit Record Rate"]],
    })
    compare["Gap"] = compare[f"Top BD: {top_bd['BD Name']}"] - compare[selected_bd]
    # Format manually
    formatted = compare.copy()
    money_metrics = ["GMV", "GMV / Store"]
    pct_metrics = ["Exposure → Visit", "Visit → Cart", "Cart → Order", "Exposure → Order", "Promo Rate", "Visit Record Rate"]
    for idx, row in formatted.iterrows():
        if row["Metric"] in money_metrics:
            for c in formatted.columns[1:]: formatted.at[idx, c] = fmt_money(compare.at[idx, c])
        elif row["Metric"] in pct_metrics:
            for c in formatted.columns[1:]: formatted.at[idx, c] = fmt_pct(compare.at[idx, c])
        else:
            for c in formatted.columns[1:]: formatted.at[idx, c] = f"{compare.at[idx, c]:,.0f}"
    st.dataframe(formatted, use_container_width=True, hide_index=True)

st.header("Merchant Intelligence")
merchant_cols = ["_merchant_name", "_merchant_id", "_bd_display", "_area", "_category", "_gmv", "_orders", "_exposure", "_visit", "_cart"]
merch = df[merchant_cols].copy()
merch.columns = ["Merchant Name", "Merchant ID", "BD Name", "Area", "Category", "GMV", "Orders", "Exposure", "Visit", "Cart"]
merch["Exposure → Visit"] = merch.apply(lambda r: safe_div(r["Visit"], r["Exposure"]), axis=1)
merch["Visit → Cart"] = merch.apply(lambda r: safe_div(r["Cart"], r["Visit"]), axis=1)
merch["Cart → Order"] = merch.apply(lambda r: safe_div(r["Orders"], r["Cart"]), axis=1)
merch["Exposure → Order"] = merch.apply(lambda r: safe_div(r["Orders"], r["Exposure"]), axis=1)
my_merch = merch[merch["BD Name"] == selected_bd].sort_values("GMV", ascending=False)
st.dataframe(style_metrics(my_merch.head(top_n)), use_container_width=True, hide_index=True)

st.header("Opportunity Finder")
opp = my_merch.copy()
# score focuses on high exposure, weak final conversion, enough GMV potential
sydney_e2o = sydney_metrics["Exposure → Order"]
opp["Opportunity Score"] = (
    opp["Exposure"].rank(pct=True).fillna(0) * 40
    + (1 - (opp["Exposure → Order"] / sydney_e2o).clip(0, 2).fillna(0) / 2) * 35
    + opp["GMV"].rank(pct=True).fillna(0) * 25
).round(1)
opp["Reason"] = opp.apply(lambda r: "High exposure but low final conversion" if r["Exposure → Order"] < sydney_e2o and r["Exposure"] > my_merch["Exposure"].median() else "Potential merchant", axis=1)
opp_cols = ["Merchant Name", "Merchant ID", "Area", "Category", "GMV", "Orders", "Exposure", "Exposure → Visit", "Visit → Cart", "Cart → Order", "Exposure → Order", "Opportunity Score", "Reason"]
st.dataframe(style_metrics(opp.sort_values("Opportunity Score", ascending=False)[opp_cols].head(top_n)), use_container_width=True, hide_index=True)

st.header("Area Intelligence")
area_rank = df.groupby("_area", dropna=False).apply(group_metrics, include_groups=False).reset_index().rename(columns={"_area":"Area"}).sort_values("GMV", ascending=False)
st.dataframe(style_metrics(area_rank[["Area", "Merchants", "GMV", "Orders", "GMV / Store", "Exposure → Visit", "Visit → Cart", "Cart → Order", "Exposure → Order"]].head(30)), use_container_width=True, hide_index=True)

# downloads
st.header("Download")
out = io.BytesIO()
with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
    bd_rank.to_excel(writer, sheet_name="BD Ranking", index=False)
    my_merch.to_excel(writer, sheet_name="My Merchants", index=False)
    opp.sort_values("Opportunity Score", ascending=False).to_excel(writer, sheet_name="Opportunity", index=False)
    area_rank.to_excel(writer, sheet_name="Area", index=False)
    METRIC_DICT.to_excel(writer, sheet_name="Metric Dictionary", index=False)
out.seek(0)
st.download_button("Download V4 Excel report", data=out, file_name="sydney_growth_intelligence_v4_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
