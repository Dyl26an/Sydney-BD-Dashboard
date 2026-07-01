import io
from typing import Optional, List

import msoffcrypto
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Sydney BD Growth Intelligence", layout="wide")

BD_NAME_CANDIDATES = ["bd姓名", "BD姓名", "bd name", "BD Name", "负责人姓名", "业务员姓名", "Owner Name"]
BD_ID_CANDIDATES = ["bd工号", "BD工号", "bd id", "BD ID", "工号", "员工号", "Owner ID"]
MERCHANT_NAME_CANDIDATES = ["商户名称", "店铺名称", "门店名称", "商家名称", "主店名称", "merchant name", "Merchant Name", "shop name", "Shop Name"]
MERCHANT_ID_CANDIDATES = ["店铺id", "店铺ID", "商户id", "商户ID", "merchant id", "Merchant ID"]
AREA_CANDIDATES = ["商圈", "区域", "area", "Area", "suburb", "Suburb"]
CATEGORY_CANDIDATES = ["主营类目", "主类目", "品类", "category", "Category"]
GMV_CANDIDATES = ["gmv", "GMV", "交易额", "销售额", "实付GMV", "支付GMV"]
ORDER_CANDIDATES = ["订单数_排除mm的均单", "订单数", "完成订单数", "有效订单", "orders", "Orders"]
EXPOSURE_CANDIDATES = ["平均曝光人数", "曝光人数", "曝光次数", "曝光量", "impressions", "Impressions"]
VISIT_CANDIDATES = ["平均进店人数", "进店人数", "访问人数", "visits", "Visits"]
CART_CANDIDATES = ["平均加购人数", "加购人数", "加购数", "cart", "Cart"]
ORDER_USER_CANDIDATES = ["平均下单人数_埋点", "下单人数_排除mm的均单", "平均下单人数"]
EXPOSURE_ORDER_RATE_CANDIDATES = ["曝光下单转化率", "曝光到下单率", "Exposure → Order", "exposure_to_order"]
EXPOSURE_VISIT_RATE_CANDIDATES = ["曝光进店转化率", "曝光到进店率"]
VISIT_CART_RATE_CANDIDATES = ["进店加购转化率", "进店到加购率"]
CART_ORDER_RATE_CANDIDATES = ["加购下单转化率", "加购到下单率"]
MATERIAL_CANDIDATES = ["是否有物料", "有物料", "material", "Material"]
VISIT_RECORD_CANDIDATES = ["是否有拜访", "拜访记录", "是否拜访", "visit record", "Visit Record"]
PROMO_BOOL_PREFIXES = ("是否设置", "是否配置", "门店是否配置")
PROMO_KEYWORDS = ("折扣", "满减", "红包", "优惠", "券", "首单", "新客", "运费", "活动")


def decrypt_excel(uploaded_file, password: str) -> bytes:
    data = uploaded_file.read()
    bio = io.BytesIO(data)
    out = io.BytesIO()
    try:
        office_file = msoffcrypto.OfficeFile(bio)
        office_file.load_key(password=password)
        office_file.decrypt(out)
        return out.getvalue()
    except Exception:
        return data


def read_excel_bytes(excel_bytes: bytes) -> pd.DataFrame:
    xl = pd.ExcelFile(io.BytesIO(excel_bytes), engine="openpyxl")
    best = None
    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet)
        if best is None or df.shape[0] * df.shape[1] > best.shape[0] * best.shape[1]:
            best = df
    return best if best is not None else pd.DataFrame()


def find_col(df: pd.DataFrame, candidates: List[str], contains: bool = False) -> Optional[str]:
    cols = list(df.columns)
    lower_map = {str(c).strip().lower(): c for c in cols}
    for cand in candidates:
        key = cand.strip().lower()
        if key in lower_map:
            return lower_map[key]
    if contains:
        for cand in candidates:
            key = cand.strip().lower()
            for c in cols:
                if key in str(c).strip().lower():
                    return c
    return None


def numeric_series(df: pd.DataFrame, col: Optional[str]) -> pd.Series:
    if not col or col not in df.columns:
        return pd.Series([0] * len(df), index=df.index, dtype="float64")
    s = df[col]
    if s.dtype == object:
        s = (
            s.astype(str)
            .str.replace("$", "", regex=False)
            .str.replace(",", "", regex=False)
            .str.replace("%", "", regex=False)
            .str.strip()
        )
    out = pd.to_numeric(s, errors="coerce").fillna(0)
    # If a rate was stored as 1.5 rather than 0.015, convert only obvious percent columns.
    if col and ("率" in str(col) or "rate" in str(col).lower() or "%" in str(col)) and out.max() > 1:
        out = out / 100
    return out


def truthy_series(df: pd.DataFrame, col: Optional[str]) -> pd.Series:
    if not col or col not in df.columns:
        return pd.Series([False] * len(df), index=df.index)
    txt = df[col].astype(str).str.lower().str.strip()
    return txt.isin(["1", "yes", "y", "true", "有", "是", "已", "done", "开", "开启"])


def detect_columns(df: pd.DataFrame) -> dict:
    bd_name_col = find_col(df, BD_NAME_CANDIDATES, contains=True)
    bd_id_col = find_col(df, BD_ID_CANDIDATES, contains=True)
    merchant_col = find_col(df, MERCHANT_NAME_CANDIDATES, contains=True)
    merchant_id_col = find_col(df, MERCHANT_ID_CANDIDATES, contains=True)
    cols = {
        "bd_name_col": bd_name_col,
        "bd_id_col": bd_id_col,
        "owner_col": bd_name_col or bd_id_col,
        "merchant_col": merchant_col,
        "merchant_id_col": merchant_id_col,
        "area_col": find_col(df, AREA_CANDIDATES, contains=True),
        "category_col": find_col(df, CATEGORY_CANDIDATES, contains=True),
        "gmv_col": find_col(df, GMV_CANDIDATES, contains=True),
        "order_col": find_col(df, ORDER_CANDIDATES, contains=True),
        "exposure_col": find_col(df, EXPOSURE_CANDIDATES, contains=True),
        "visit_col": find_col(df, VISIT_CANDIDATES, contains=True),
        "cart_col": find_col(df, CART_CANDIDATES, contains=True),
        "order_user_col": find_col(df, ORDER_USER_CANDIDATES, contains=True),
        "exp_order_rate_col": find_col(df, EXPOSURE_ORDER_RATE_CANDIDATES, contains=True),
        "exp_visit_rate_col": find_col(df, EXPOSURE_VISIT_RATE_CANDIDATES, contains=True),
        "visit_cart_rate_col": find_col(df, VISIT_CART_RATE_CANDIDATES, contains=True),
        "cart_order_rate_col": find_col(df, CART_ORDER_RATE_CANDIDATES, contains=True),
        "material_col": find_col(df, MATERIAL_CANDIDATES, contains=True),
        "visit_record_col": find_col(df, VISIT_RECORD_CANDIDATES, contains=True),
    }
    promo_cols = []
    for c in df.columns:
        cs = str(c)
        if cs.startswith(PROMO_BOOL_PREFIXES) and any(k in cs for k in PROMO_KEYWORDS):
            promo_cols.append(c)
    cols["promo_bool_cols"] = promo_cols
    return cols


def add_metrics(df: pd.DataFrame, cols: dict) -> pd.DataFrame:
    out = df.copy()
    out["_bd_display"] = out[cols["owner_col"]].astype(str).str.strip() if cols.get("owner_col") else "Unknown"
    out["_bd_id"] = out[cols["bd_id_col"]].astype(str).str.strip() if cols.get("bd_id_col") else ""
    out["_merchant"] = out[cols["merchant_col"]].astype(str).str.strip() if cols.get("merchant_col") else out.index.astype(str)
    out["_merchant_id"] = out[cols["merchant_id_col"]].astype(str).str.strip() if cols.get("merchant_id_col") else ""
    out["_area"] = out[cols["area_col"]].astype(str).str.strip() if cols.get("area_col") else "Unknown"
    out["_category"] = out[cols["category_col"]].astype(str).str.strip() if cols.get("category_col") else "Unknown"
    out["_gmv"] = numeric_series(out, cols.get("gmv_col"))
    out["_orders"] = numeric_series(out, cols.get("order_col"))
    out["_exposure"] = numeric_series(out, cols.get("exposure_col"))
    out["_visits"] = numeric_series(out, cols.get("visit_col"))
    out["_cart"] = numeric_series(out, cols.get("cart_col"))
    out["_order_users"] = numeric_series(out, cols.get("order_user_col"))

    out["_exposure_to_order"] = numeric_series(out, cols.get("exp_order_rate_col")) if cols.get("exp_order_rate_col") else out["_orders"] / out["_exposure"].replace(0, pd.NA)
    out["_visit_rate"] = numeric_series(out, cols.get("exp_visit_rate_col")) if cols.get("exp_visit_rate_col") else out["_visits"] / out["_exposure"].replace(0, pd.NA)
    out["_cart_rate"] = numeric_series(out, cols.get("visit_cart_rate_col")) if cols.get("visit_cart_rate_col") else out["_cart"] / out["_visits"].replace(0, pd.NA)
    out["_cart_to_order"] = numeric_series(out, cols.get("cart_order_rate_col")) if cols.get("cart_order_rate_col") else out["_orders"] / out["_cart"].replace(0, pd.NA)

    promo_cols = cols.get("promo_bool_cols", [])
    if promo_cols:
        bool_df = pd.concat([truthy_series(out, c) for c in promo_cols], axis=1)
        out["_promo_count"] = bool_df.sum(axis=1)
        out["_has_promo"] = out["_promo_count"] > 0
    else:
        out["_promo_count"] = 0
        out["_has_promo"] = False
    out["_has_material"] = truthy_series(out, cols.get("material_col"))
    out["_has_visit_record"] = truthy_series(out, cols.get("visit_record_col"))
    return out


def bd_ranking(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("_bd_display", dropna=False).agg(
        Merchants=("_merchant", "count"),
        GMV=("_gmv", "sum"),
        Orders=("_orders", "sum"),
        Exposure=("_exposure", "sum"),
        Visits=("_visits", "sum"),
        Cart=("_cart", "sum"),
        Promo_Rate=("_has_promo", "mean"),
        Material_Rate=("_has_material", "mean"),
        Visit_Record_Rate=("_has_visit_record", "mean"),
    ).reset_index().rename(columns={"_bd_display": "BD Name"})
    g["GMV / Store"] = g["GMV"] / g["Merchants"].replace(0, pd.NA)
    g["Exposure → Order"] = g["Orders"] / g["Exposure"].replace(0, pd.NA)
    # Use weighted average of provided conversion rate when order/exposure are low or missing.
    weighted = df.groupby("_bd_display").apply(lambda x: (x["_exposure_to_order"].fillna(0) * x["_exposure"].replace(0, 1)).sum() / x["_exposure"].replace(0, 1).sum()).reset_index(name="Weighted Exposure → Order")
    g = g.merge(weighted, left_on="BD Name", right_on="_bd_display", how="left").drop(columns=["_bd_display"])
    g["Exposure → Order"] = g["Exposure → Order"].fillna(0)
    g.loc[g["Exposure → Order"] == 0, "Exposure → Order"] = g.loc[g["Exposure → Order"] == 0, "Weighted Exposure → Order"]
    g = g.sort_values("GMV", ascending=False).reset_index(drop=True)
    g.insert(0, "Rank", range(1, len(g) + 1))
    return g


def area_ranking(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("_area").agg(Merchants=("_merchant", "count"), GMV=("_gmv", "sum"), Orders=("_orders", "sum"), Exposure=("_exposure", "sum")).reset_index().rename(columns={"_area": "Area"})
    g["GMV / Store"] = g["GMV"] / g["Merchants"].replace(0, pd.NA)
    g["Exposure → Order"] = g["Orders"] / g["Exposure"].replace(0, pd.NA)
    return g.sort_values("GMV", ascending=False)


def merchant_table(df: pd.DataFrame, selected_bd: Optional[str] = None, top_n: int = 100) -> pd.DataFrame:
    d = df.copy() if not selected_bd else df[df["_bd_display"] == selected_bd].copy()
    d = d.sort_values("_gmv", ascending=False).head(top_n)
    return pd.DataFrame({
        "BD Name": d["_bd_display"],
        "Merchant Name": d["_merchant"],
        "Merchant ID": d["_merchant_id"],
        "Area": d["_area"],
        "Category": d["_category"],
        "GMV": d["_gmv"],
        "Orders": d["_orders"],
        "Exposure": d["_exposure"],
        "Visits": d["_visits"],
        "Cart": d["_cart"],
        "Exposure → Order": d["_exposure_to_order"],
        "Visit Rate": d["_visit_rate"],
        "Cart Rate": d["_cart_rate"],
        "Cart → Order": d["_cart_to_order"],
        "Promo Count": d["_promo_count"],
        "Has Material": d["_has_material"],
        "Has Visit Record": d["_has_visit_record"],
    })


def build_action_plan(df: pd.DataFrame, selected_bd: str, top_n: int = 30) -> pd.DataFrame:
    me = df[df["_bd_display"] == selected_bd].copy()
    if me.empty:
        return pd.DataFrame()
    median_conv = df["_exposure_to_order"].fillna(0).median()
    me["Opportunity Score"] = (
        me["_gmv"].rank(pct=True) * 20
        + me["_exposure"].rank(pct=True) * 25
        + (me["_exposure_to_order"].fillna(0) < median_conv).astype(int) * 25
        + (~me["_has_promo"]).astype(int) * 10
        + (~me["_has_material"]).astype(int) * 10
        + (~me["_has_visit_record"]).astype(int) * 10
    )
    def reason(row):
        r = []
        if row["_exposure"] > me["_exposure"].median(): r.append("曝光高")
        if row["_exposure_to_order"] < median_conv: r.append("曝光下单转化低")
        if not row["_has_promo"]: r.append("活动弱/无活动")
        if not row["_has_material"]: r.append("无物料")
        if not row["_has_visit_record"]: r.append("无拜访记录")
        return "；".join(r) or "表现稳定"
    def action(row):
        a = []
        if row["_exposure_to_order"] < median_conv: a.append("检查首图、菜单排序、爆品价格和套餐结构")
        if not row["_has_promo"]: a.append("补充折扣券/满减/新客立减/运费券")
        if not row["_has_material"]: a.append("安排店内物料")
        if not row["_has_visit_record"]: a.append("本周安排拜访")
        return "；".join(a) or "保持跟进"
    me = me.sort_values("Opportunity Score", ascending=False).head(top_n)
    plan = merchant_table(me, None, top_n=top_n)
    plan["Opportunity Score"] = me["Opportunity Score"].values
    plan["Opportunity Reason"] = me.apply(reason, axis=1).values
    plan["Recommended Action"] = me.apply(action, axis=1).values
    return plan


def format_money(x):
    try:
        return f"${float(x):,.0f}"
    except Exception:
        return "$0"

def fmt_pct(x):
    try:
        return f"{float(x):.2%}"
    except Exception:
        return "-"

st.title("Sydney Growth Intelligence — V2.2")
st.caption("修复订单/曝光字段识别 · 显示商户名称 · 更完整的店铺表和学习对比")

with st.sidebar:
    st.header("Upload")
    uploaded = st.file_uploader("Encrypted Excel", type=["xlsx", "xlsm", "xls"])
    password = st.text_input("Password", type="password")

if not uploaded:
    st.info("Upload your monthly Excel report to start.")
    st.stop()

try:
    excel_bytes = decrypt_excel(uploaded, password)
    raw = read_excel_bytes(excel_bytes)
    cols = detect_columns(raw)
    df = add_metrics(raw, cols)
except Exception as e:
    st.error(f"Failed to read file: {e}")
    st.stop()

st.success(f"Loaded {len(df):,} rows and {len(raw.columns):,} columns.")
if not cols.get("owner_col"):
    st.error("Could not find BD name or BD ID column.")
    st.stop()

bd_options = sorted([x for x in df["_bd_display"].dropna().astype(str).unique() if x and x.lower() != "nan"])
def_idx = next((i for i, n in enumerate(bd_options) if "yuan" in n.lower() and "dong" in n.lower()), 0)
selected_bd = st.sidebar.selectbox("BD Name", bd_options, index=def_idx if bd_options else 0)

ranking = bd_ranking(df)
my_row = ranking[ranking["BD Name"] == selected_bd]

st.subheader("Overview")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Sydney GMV", format_money(df["_gmv"].sum()))
c2.metric("Sydney Orders", f"{df['_orders'].sum():,.0f}")
c3.metric("Merchants", f"{len(df):,}")
c4.metric("BD Count", f"{len(bd_options):,}")

if not my_row.empty:
    r = my_row.iloc[0]
    st.subheader(f"My Performance — {selected_bd}")
    a,b,c,d,e = st.columns(5)
    a.metric("Rank", f"#{int(r['Rank'])}")
    b.metric("My GMV", format_money(r["GMV"]))
    c.metric("My Merchants", f"{int(r['Merchants']):,}")
    d.metric("GMV / Store", format_money(r["GMV / Store"]))
    e.metric("Exposure → Order", fmt_pct(r["Exposure → Order"]))

tab1, tab2, tab3, tab4, tab5 = st.tabs(["BD Ranking", "Merchant List", "Compare Me", "Learn From Best", "Action Plan"])

with tab1:
    st.subheader("BD Ranking")
    show = ranking[["Rank", "BD Name", "Merchants", "GMV", "Orders", "GMV / Store", "Exposure → Order", "Promo_Rate", "Material_Rate", "Visit_Record_Rate"]]
    st.dataframe(show.style.format({"GMV":"${:,.0f}", "Orders":"{:,.0f}", "GMV / Store":"${:,.0f}", "Exposure → Order":"{:.2%}", "Promo_Rate":"{:.1%}", "Material_Rate":"{:.1%}", "Visit_Record_Rate":"{:.1%}"}), use_container_width=True, hide_index=True)
    st.subheader("Area Ranking")
    ar = area_ranking(df)
    st.dataframe(ar.style.format({"GMV":"${:,.0f}", "Orders":"{:,.0f}", "GMV / Store":"${:,.0f}", "Exposure → Order":"{:.2%}"}), use_container_width=True, hide_index=True)

with tab2:
    st.subheader(f"Merchant List — {selected_bd}")
    mt = merchant_table(df, selected_bd, top_n=200)
    st.dataframe(mt.style.format({"GMV":"${:,.0f}", "Orders":"{:,.0f}", "Exposure":"{:,.0f}", "Visits":"{:,.0f}", "Cart":"{:,.0f}", "Exposure → Order":"{:.2%}", "Visit Rate":"{:.2%}", "Cart Rate":"{:.2%}", "Cart → Order":"{:.2%}"}), use_container_width=True, hide_index=True)

with tab3:
    st.subheader("Compare Me vs Top BD")
    if not my_row.empty and not ranking.empty:
        m = my_row.iloc[0]
        t = ranking.iloc[0]
        compare = pd.DataFrame({
            "KPI": ["GMV", "Orders", "Merchants", "GMV / Store", "Exposure → Order", "Promo Rate", "Material Rate", "Visit Record Rate"],
            selected_bd: [m["GMV"], m["Orders"], m["Merchants"], m["GMV / Store"], m["Exposure → Order"], m["Promo_Rate"], m["Material_Rate"], m["Visit_Record_Rate"]],
            f"Top BD: {t['BD Name']}": [t["GMV"], t["Orders"], t["Merchants"], t["GMV / Store"], t["Exposure → Order"], t["Promo_Rate"], t["Material_Rate"], t["Visit_Record_Rate"]],
        })
        st.dataframe(compare, use_container_width=True, hide_index=True)
        gaps = []
        if m["GMV / Store"] < t["GMV / Store"]: gaps.append("单店GMV")
        if m["Exposure → Order"] < t["Exposure → Order"]: gaps.append("曝光→下单转化")
        if m["Visit_Record_Rate"] < t["Visit_Record_Rate"]: gaps.append("拜访覆盖")
        if m["Material_Rate"] < t["Material_Rate"]: gaps.append("物料覆盖")
        st.info("优先学习方向：" + ("、".join(gaps) if gaps else "当前核心指标接近Top BD"))

with tab4:
    top_bd = ranking.iloc[0]["BD Name"] if not ranking.empty else selected_bd
    st.subheader(f"Learn From Best — {top_bd}")
    best = merchant_table(df, top_bd, top_n=30)
    st.dataframe(best.style.format({"GMV":"${:,.0f}", "Orders":"{:,.0f}", "Exposure":"{:,.0f}", "Exposure → Order":"{:.2%}", "Visit Rate":"{:.2%}", "Cart Rate":"{:.2%}", "Cart → Order":"{:.2%}"}), use_container_width=True, hide_index=True)
    st.write("学习重点：优先看 Top BD 高GMV店铺的品类、商圈、活动数、拜访覆盖与转化率，再对照你的同品类店铺复制。")

with tab5:
    st.subheader(f"Action Plan — {selected_bd}")
    plan = build_action_plan(df, selected_bd, top_n=50)
    if plan.empty:
        st.warning("No merchant found for selected BD.")
    else:
        st.dataframe(plan.style.format({"GMV":"${:,.0f}", "Orders":"{:,.0f}", "Exposure":"{:,.0f}", "Exposure → Order":"{:.2%}", "Visit Rate":"{:.2%}", "Cart Rate":"{:.2%}", "Cart → Order":"{:.2%}", "Opportunity Score":"{:.1f}"}), use_container_width=True, hide_index=True)

output = io.BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    ranking.to_excel(writer, index=False, sheet_name="BD Ranking")
    merchant_table(df, selected_bd, top_n=500).to_excel(writer, index=False, sheet_name="My Merchants")
    merchant_table(df, None, top_n=1000).to_excel(writer, index=False, sheet_name="All Merchants")
    build_action_plan(df, selected_bd, top_n=100).to_excel(writer, index=False, sheet_name="Action Plan")
st.download_button("Download V2.2 Excel Report", output.getvalue(), "sydney_bd_dashboard_v22_report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with st.expander("Detected columns"):
    st.json(cols)
