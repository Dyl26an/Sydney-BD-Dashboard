import io
from typing import Optional, List, Tuple

import msoffcrypto
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Sydney BD Growth Intelligence", layout="wide")

NAME_COL_CANDIDATES = [
    "bd姓名", "BD姓名", "bd name", "BD Name", "BD_NAME", "owner name", "Owner Name",
    "负责人姓名", "业务员姓名", "BD", "bd", "负责人", "owner", "Owner",
]
ID_COL_CANDIDATES = [
    "bd工号", "BD工号", "bd id", "BD ID", "BD_ID", "owner id", "Owner ID", "工号", "员工号",
]
MERCHANT_COL_CANDIDATES = [
    "店铺名称", "门店名称", "商家名称", "merchant name", "Merchant Name", "shop name", "Shop Name", "name", "Name"
]
AREA_COL_CANDIDATES = ["商圈", "区域", "area", "Area", "suburb", "Suburb", "城市区域"]
GMV_COL_CANDIDATES = ["GMV", "gmv", "交易额", "销售额", "实付GMV", "支付GMV", "订单金额"]
ORDER_COL_CANDIDATES = ["订单", "订单数", "orders", "Orders", "有效订单", "完成订单数"]
EXPOSURE_COL_CANDIDATES = ["曝光", "曝光人数", "曝光次数", "impression", "Impressions", "曝光量"]
VISIT_COL_CANDIDATES = ["进店", "进店人数", "访问", "visits", "Visits", "店铺访问"]
CART_COL_CANDIDATES = ["加购", "加购人数", "cart", "Cart", "Add to cart", "加购数"]
MATERIAL_COL_CANDIDATES = ["物料", "是否有物料", "material", "Material"]
VISIT_RECORD_COL_CANDIDATES = ["拜访", "拜访记录", "是否拜访", "visit record", "Visit Record"]
PROMO_COL_KEYWORDS = ["折扣", "满减", "优惠", "券", "promotion", "promo", "discount", "rebate", "活动"]


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
        # If file is not encrypted, try returning original bytes.
        return data


def read_excel_bytes(excel_bytes: bytes) -> pd.DataFrame:
    xl = pd.ExcelFile(io.BytesIO(excel_bytes), engine="openpyxl")
    # Prefer first sheet with rows/columns
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
        for c in cols:
            lc = str(c).strip().lower()
            for cand in candidates:
                if cand.strip().lower() in lc:
                    return c
    return None


def numeric_series(df: pd.DataFrame, col: Optional[str]) -> pd.Series:
    if not col or col not in df.columns:
        return pd.Series([0] * len(df), index=df.index, dtype="float64")
    s = df[col]
    if s.dtype == object:
        s = s.astype(str).str.replace("$", "", regex=False).str.replace(",", "", regex=False).str.replace("%", "", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0)


def detect_columns(df: pd.DataFrame) -> dict:
    bd_name_col = find_col(df, NAME_COL_CANDIDATES, contains=True)
    bd_id_col = find_col(df, ID_COL_CANDIDATES, contains=True)
    # If no explicit name column, use ID column as fallback.
    owner_col = bd_name_col or bd_id_col
    merchant_col = find_col(df, MERCHANT_COL_CANDIDATES, contains=True)
    area_col = find_col(df, AREA_COL_CANDIDATES, contains=True)
    gmv_col = find_col(df, GMV_COL_CANDIDATES, contains=True)
    order_col = find_col(df, ORDER_COL_CANDIDATES, contains=True)
    exposure_col = find_col(df, EXPOSURE_COL_CANDIDATES, contains=True)
    visit_col = find_col(df, VISIT_COL_CANDIDATES, contains=True)
    cart_col = find_col(df, CART_COL_CANDIDATES, contains=True)
    material_col = find_col(df, MATERIAL_COL_CANDIDATES, contains=True)
    visit_record_col = find_col(df, VISIT_RECORD_COL_CANDIDATES, contains=True)
    promo_cols = [c for c in df.columns if any(k.lower() in str(c).lower() for k in PROMO_COL_KEYWORDS)]
    return {
        "bd_name_col": bd_name_col,
        "bd_id_col": bd_id_col,
        "owner_col": owner_col,
        "merchant_col": merchant_col,
        "area_col": area_col,
        "gmv_col": gmv_col,
        "order_col": order_col,
        "exposure_col": exposure_col,
        "visit_col": visit_col,
        "cart_col": cart_col,
        "material_col": material_col,
        "visit_record_col": visit_record_col,
        "promo_cols": promo_cols,
    }


def add_metrics(df: pd.DataFrame, cols: dict) -> pd.DataFrame:
    out = df.copy()
    out["_bd_display"] = out[cols["owner_col"]].astype(str).str.strip() if cols["owner_col"] else "Unknown"
    out["_merchant"] = out[cols["merchant_col"]].astype(str).str.strip() if cols["merchant_col"] else out.index.astype(str)
    out["_area"] = out[cols["area_col"]].astype(str).str.strip() if cols["area_col"] else "Unknown"
    out["_gmv"] = numeric_series(out, cols["gmv_col"])
    out["_orders"] = numeric_series(out, cols["order_col"])
    out["_exposure"] = numeric_series(out, cols["exposure_col"])
    out["_visits"] = numeric_series(out, cols["visit_col"])
    out["_cart"] = numeric_series(out, cols["cart_col"])
    out["_exposure_to_order"] = out["_orders"] / out["_exposure"].replace(0, pd.NA)
    out["_visit_rate"] = out["_visits"] / out["_exposure"].replace(0, pd.NA)
    out["_cart_rate"] = out["_cart"] / out["_visits"].replace(0, pd.NA)
    out["_cart_to_order"] = out["_orders"] / out["_cart"].replace(0, pd.NA)

    promo_cols = cols.get("promo_cols", [])
    if promo_cols:
        promo_numeric = pd.concat([numeric_series(out, c) for c in promo_cols], axis=1)
        out["_promo_count"] = (promo_numeric > 0).sum(axis=1)
        out["_has_promo"] = out["_promo_count"] > 0
    else:
        out["_promo_count"] = 0
        out["_has_promo"] = False

    def truthy(s):
        if s is None or s not in out.columns:
            return pd.Series([False] * len(out), index=out.index)
        txt = out[s].astype(str).str.lower().str.strip()
        return txt.isin(["1", "yes", "y", "true", "有", "是", "已", "done"])

    out["_has_material"] = truthy(cols.get("material_col"))
    out["_has_visit_record"] = truthy(cols.get("visit_record_col"))
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
    g["Visit Rate"] = g["Visits"] / g["Exposure"].replace(0, pd.NA)
    g["Cart Rate"] = g["Cart"] / g["Visits"].replace(0, pd.NA)
    g = g.sort_values("GMV", ascending=False).reset_index(drop=True)
    g.insert(0, "Rank", range(1, len(g) + 1))
    return g


def build_action_plan(df: pd.DataFrame, selected_bd: str, top_n: int = 30) -> pd.DataFrame:
    me = df[df["_bd_display"] == selected_bd].copy()
    if me.empty:
        return pd.DataFrame()
    # opportunity score: high exposure, low conversion, missing promo/material/visit record
    conv = me["_exposure_to_order"].fillna(0)
    median_conv = df["_exposure_to_order"].replace([pd.NA], 0).fillna(0).median()
    me["Opportunity Score"] = (
        me["_exposure"].rank(pct=True) * 40
        + (conv < median_conv).astype(int) * 25
        + (~me["_has_promo"]).astype(int) * 15
        + (~me["_has_material"]).astype(int) * 10
        + (~me["_has_visit_record"]).astype(int) * 10
    )

    def reason(row):
        r = []
        if row["_exposure"] > me["_exposure"].median(): r.append("high exposure")
        if pd.notna(row["_exposure_to_order"]) and row["_exposure_to_order"] < median_conv: r.append("low conversion")
        if not row["_has_promo"]: r.append("weak/no promo")
        if not row["_has_material"]: r.append("no material")
        if not row["_has_visit_record"]: r.append("no visit record")
        return ", ".join(r) or "stable merchant"

    def action(row):
        actions = []
        if not row["_has_promo"]: actions.append("add coupon / bundle / discount campaign")
        if pd.notna(row["_exposure_to_order"]) and row["_exposure_to_order"] < median_conv: actions.append("review hero image, menu order and pricing")
        if not row["_has_material"]: actions.append("arrange in-store material")
        if not row["_has_visit_record"]: actions.append("schedule merchant visit this week")
        return "; ".join(actions) or "maintain and monitor"

    plan = me.sort_values("Opportunity Score", ascending=False).head(top_n).copy()
    plan["BD Name"] = plan["_bd_display"]
    plan["Merchant Name"] = plan["_merchant"]
    plan["Area"] = plan["_area"]
    plan["GMV"] = plan["_gmv"]
    plan["Orders"] = plan["_orders"]
    plan["Exposure"] = plan["_exposure"]
    plan["Exposure → Order"] = plan["_exposure_to_order"]
    plan["Reason"] = plan.apply(reason, axis=1)
    plan["Recommended Action"] = plan.apply(action, axis=1)
    return plan[["BD Name", "Merchant Name", "Area", "GMV", "Orders", "Exposure", "Exposure → Order", "Opportunity Score", "Reason", "Recommended Action"]]


def format_money(x):
    try:
        return f"${x:,.0f}"
    except Exception:
        return "$0"


def format_pct(x):
    try:
        return f"{x:.2%}"
    except Exception:
        return "-"


st.title("Sydney Growth Intelligence — V2.1")
st.caption("BD names first · Compare with top BD · Learn from best merchants · Merchant-name action plan")

with st.sidebar:
    st.header("Upload")
    uploaded = st.file_uploader("Encrypted Excel", type=["xlsx", "xlsm", "xls"])
    password = st.text_input("Password", type="password")
    st.caption("V2.1 will prefer BD name columns over BD ID/work number columns.")

if not uploaded:
    st.info("Upload your monthly Excel report to start.")
    st.stop()

try:
    excel_bytes = decrypt_excel(uploaded, password)
    raw = read_excel_bytes(excel_bytes)
    cols = detect_columns(raw)
    df = add_metrics(raw, cols)
except Exception as e:
    st.error(f"Failed to read the file: {e}")
    st.stop()

st.success(f"Loaded {len(df):,} rows and {len(raw.columns):,} columns.")

if not cols["owner_col"]:
    st.error("Could not find BD name/ID column. Please check if the file has bd姓名, BD Name, bd工号, or Owner columns.")
    st.stop()

bd_options = sorted([x for x in df["_bd_display"].dropna().astype(str).unique() if x and x.lower() != "nan"])
def_idx = 0
for i, name in enumerate(bd_options):
    if "yuan" in name.lower() and "dong" in name.lower():
        def_idx = i
        break

selected_bd = st.sidebar.selectbox("BD Name", bd_options, index=def_idx if bd_options else 0)

ranking = bd_ranking(df)
my_row = ranking[ranking["BD Name"] == selected_bd]
top_row = ranking.iloc[[0]] if not ranking.empty else pd.DataFrame()

st.subheader("Overview")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Sydney GMV", format_money(df["_gmv"].sum()))
c2.metric("Sydney Orders", f"{df['_orders'].sum():,.0f}")
c3.metric("Merchants", f"{len(df):,}")
c4.metric("BD Count", f"{len(bd_options):,}")

if not my_row.empty:
    r = my_row.iloc[0]
    st.subheader(f"My Performance — {selected_bd}")
    a, b, c, d, e = st.columns(5)
    a.metric("Rank", f"#{int(r['Rank'])}")
    b.metric("My GMV", format_money(r["GMV"]))
    c.metric("My Merchants", f"{int(r['Merchants']):,}")
    d.metric("GMV / Store", format_money(r["GMV / Store"]))
    e.metric("Exposure → Order", format_pct(r["Exposure → Order"]))

st.subheader("BD Ranking — Name Display")
show_cols = ["Rank", "BD Name", "Merchants", "GMV", "Orders", "GMV / Store", "Exposure → Order", "Promo_Rate", "Material_Rate", "Visit_Record_Rate"]
st.dataframe(
    ranking[show_cols].style.format({
        "GMV": "${:,.0f}", "GMV / Store": "${:,.0f}", "Exposure → Order": "{:.2%}",
        "Promo_Rate": "{:.1%}", "Material_Rate": "{:.1%}", "Visit_Record_Rate": "{:.1%}",
    }),
    use_container_width=True,
    hide_index=True,
)

st.subheader("Compare Me vs Top BD")
if not my_row.empty and not top_row.empty:
    m, t = my_row.iloc[0], top_row.iloc[0]
    compare = pd.DataFrame({
        "KPI": ["GMV", "Merchants", "GMV / Store", "Exposure → Order", "Promo Rate", "Material Rate", "Visit Record Rate"],
        selected_bd: [m["GMV"], m["Merchants"], m["GMV / Store"], m["Exposure → Order"], m["Promo_Rate"], m["Material_Rate"], m["Visit_Record_Rate"]],
        f"Top BD: {t['BD Name']}": [t["GMV"], t["Merchants"], t["GMV / Store"], t["Exposure → Order"], t["Promo_Rate"], t["Material_Rate"], t["Visit_Record_Rate"]],
    })
    st.dataframe(compare, use_container_width=True, hide_index=True)
    gaps = []
    if m["Promo_Rate"] < t["Promo_Rate"]: gaps.append("promotion coverage")
    if m["Material_Rate"] < t["Material_Rate"]: gaps.append("in-store material coverage")
    if m["Visit_Record_Rate"] < t["Visit_Record_Rate"]: gaps.append("merchant visit coverage")
    if m["Exposure → Order"] < t["Exposure → Order"]: gaps.append("exposure-to-order conversion")
    st.info("Priority learning areas: " + (", ".join(gaps) if gaps else "you are close to the top BD on the tracked metrics."))

st.subheader("Learn From Best")
if not top_row.empty:
    top_bd = top_row.iloc[0]["BD Name"]
    top_merchants = df[df["_bd_display"] == top_bd].sort_values("_gmv", ascending=False).head(20)
    best = top_merchants.assign(
        **{
            "BD Name": top_merchants["_bd_display"],
            "Merchant Name": top_merchants["_merchant"],
            "Area": top_merchants["_area"],
            "GMV": top_merchants["_gmv"],
            "Orders": top_merchants["_orders"],
            "Exposure → Order": top_merchants["_exposure_to_order"],
            "Promo Count": top_merchants["_promo_count"],
            "Has Material": top_merchants["_has_material"],
        }
    )[["BD Name", "Merchant Name", "Area", "GMV", "Orders", "Exposure → Order", "Promo Count", "Has Material"]]
    st.write(f"Top merchants managed by **{top_bd}**:")
    st.dataframe(best.style.format({"GMV": "${:,.0f}", "Exposure → Order": "{:.2%}"}), use_container_width=True, hide_index=True)

st.subheader(f"Action Plan — {selected_bd}")
plan = build_action_plan(df, selected_bd, top_n=30)
if plan.empty:
    st.warning("No merchant found for this BD. Please choose another BD name from the dropdown.")
else:
    st.dataframe(plan.style.format({"GMV": "${:,.0f}", "Exposure → Order": "{:.2%}", "Opportunity Score": "{:.1f}"}), use_container_width=True, hide_index=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        ranking.to_excel(writer, index=False, sheet_name="BD Ranking")
        plan.to_excel(writer, index=False, sheet_name="Action Plan")
        df[["_bd_display", "_merchant", "_area", "_gmv", "_orders", "_exposure", "_exposure_to_order", "_has_promo", "_has_material", "_has_visit_record"]].to_excel(writer, index=False, sheet_name="Merchant Metrics")
    st.download_button(
        "Download V2.1 Excel Report",
        data=output.getvalue(),
        file_name="sydney_bd_dashboard_v21_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with st.expander("Detected columns"):
    st.json(cols)
