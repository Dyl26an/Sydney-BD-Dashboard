import io
from typing import Optional, List, Dict

import pandas as pd
import streamlit as st
import msoffcrypto

st.set_page_config(page_title="Sydney BD AI Copilot", layout="wide")
st.title("Sydney BD AI Copilot")
st.caption("Upload encrypted Excel → analyse merchant performance → generate action plan")

# -----------------------------
# Helpers
# -----------------------------
def decrypt_excel(uploaded_file, password: str) -> io.BytesIO:
    encrypted = io.BytesIO(uploaded_file.getvalue())
    decrypted = io.BytesIO()
    office_file = msoffcrypto.OfficeFile(encrypted)
    office_file.load_key(password=password)
    office_file.decrypt(decrypted)
    decrypted.seek(0)
    return decrypted


def clean_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def find_col(df: pd.DataFrame, exact: List[str], keywords: List[str]) -> Optional[str]:
    cols = list(df.columns)
    for e in exact:
        if e in cols:
            return e
    norm = {str(c).lower().replace(" ", "").replace("_", ""): c for c in cols}
    for kw in keywords:
        k = kw.lower().replace(" ", "").replace("_", "")
        for nc, original in norm.items():
            if k in nc:
                return original
    return None


def detect_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    return {
        "owner_name": find_col(df, ["bd姓名", "BD姓名", "负责人姓名"], ["bdname", "ownername", "负责人姓名", "姓名"]),
        "owner_code": find_col(df, ["bd工号", "BD工号", "bd id", "BD ID"], ["bd工号", "bdid", "ownerid", "工号"]),
        "merchant_id": find_col(df, ["店铺id", "商户id", "merchant_id"], ["店铺id", "商户id", "merchantid", "storeid"]),
        "merchant_name": find_col(df, ["商户名称", "店铺名称", "门店名称", "主店名称"], ["商户名称", "店铺名称", "门店名称", "merchantname", "storename", "restaurantname", "name"]),
        "area": find_col(df, ["商圈", "区域", "地区", "suburb", "area"], ["商圈", "区域", "地区", "suburb", "area", "zone"]),
        "category": find_col(df, ["主营类目", "店铺二级分类"], ["主营", "类目", "category"]),
        "gmv": find_col(df, ["gmv", "GMV"], ["gmv", "sales", "revenue", "交易额", "销售额"]),
        "orders": find_col(df, ["订单数_排除mm的均单", "订单数", "必选_订单配送类型"], ["订单数", "orders", "下单"]),
        "exposure_rate": find_col(df, ["曝光下单转化率"], ["曝光下单", "曝光→下单", "exposureorder"]),
        "visit_rate": find_col(df, ["曝光进店转化率"], ["曝光进店", "exposurevisit"]),
        "cart_rate": find_col(df, ["进店加购转化率"], ["进店加购", "visitcart"]),
        "cart_order_rate": find_col(df, ["加购下单转化率"], ["加购下单", "cartorder"]),
        "exposure": find_col(df, ["平均曝光人数", "曝光人数", "曝光"], ["平均曝光", "曝光人数", "exposure", "impression"]),
        "visit": find_col(df, ["平均进店人数", "进店人数"], ["平均进店", "进店人数", "visit"]),
        "cart": find_col(df, ["平均加购人数", "加购人数"], ["平均加购", "加购人数", "cart"]),
        "order_users": find_col(df, ["平均下单人数_埋点", "店铺下单人数"], ["平均下单人数", "店铺下单人数", "下单人数"]),
        "material": find_col(df, ["是否有物料", "是否有物料 ", "店铺历史物料合格率"], ["物料", "material"]),
        "visit_record": find_col(df, ["是否有拜访", "拜访记录", "近30天是否拜访"], ["拜访", "visited", "visitrecord"]),
        "single_discount": find_col(df, ["是否设置了单品折扣"], ["单品折扣"]),
        "new_customer_discount": find_col(df, ["是否设置了门店新客折扣"], ["新客折扣", "新客"]),
        "delivery_coupon": find_col(df, ["门店是否配置运费减免"], ["运费减免", "运费券", "delivery"]),
        "store_discount": find_col(df, ["是否配置独享全店折扣"], ["全店折扣", "门店折扣"]),
    }


def to_num(s):
    return pd.to_numeric(s.astype(str).str.replace(r"[$,% ,]", "", regex=True), errors="coerce")


def yes_rate(s: pd.Series) -> float:
    if s is None or len(s) == 0:
        return 0.0
    vals = s.astype(str).str.strip().str.lower()
    yes = vals.isin(["是", "yes", "y", "1", "true", "有"])
    return float(yes.mean())


def money(x):
    if x is None or pd.isna(x):
        return "-"
    return f"A${x:,.0f}"


def pct(x):
    if x is None or pd.isna(x):
        return "-"
    return f"{x:.2%}"


def prepare(df: pd.DataFrame, cols: Dict[str, Optional[str]]) -> pd.DataFrame:
    df = clean_cols(df)
    numeric_keys = ["gmv", "orders", "exposure_rate", "visit_rate", "cart_rate", "cart_order_rate", "exposure", "visit", "cart", "order_users"]
    for k in numeric_keys:
        c = cols.get(k)
        if c and c in df.columns:
            df[c] = to_num(df[c])
    return df


def filter_owner(df: pd.DataFrame, cols: Dict[str, Optional[str]], selected: str) -> pd.DataFrame:
    if selected == "All Sydney" or not selected:
        return df.copy()
    masks = []
    for k in ["owner_name", "owner_code"]:
        c = cols.get(k)
        if c and c in df.columns:
            masks.append(df[c].astype(str).str.strip().eq(selected))
            masks.append(df[c].astype(str).str.contains(selected, case=False, na=False, regex=False))
    if not masks:
        return df.iloc[0:0].copy()
    mask = masks[0]
    for m in masks[1:]:
        mask = mask | m
    return df[mask].copy()


def owner_options(df: pd.DataFrame, cols: Dict[str, Optional[str]]) -> List[str]:
    opts = ["All Sydney"]
    c = cols.get("owner_name") or cols.get("owner_code")
    if c and c in df.columns:
        vals = sorted([v for v in df[c].dropna().astype(str).str.strip().unique() if v and v.lower() != "nan"])
        opts.extend(vals)
    return opts


def add_action_plan(my: pd.DataFrame, all_df: pd.DataFrame, cols: Dict[str, Optional[str]]) -> pd.DataFrame:
    df = my.copy()
    if len(df) == 0:
        return df
    gmv, orders, exp, exp_rate = cols.get("gmv"), cols.get("orders"), cols.get("exposure"), cols.get("exposure_rate")
    visit_rate, cart_rate, cart_order_rate = cols.get("visit_rate"), cols.get("cart_rate"), cols.get("cart_order_rate")
    material, visit_record = cols.get("material"), cols.get("visit_record")
    promo_cols = [cols.get("single_discount"), cols.get("new_customer_discount"), cols.get("delivery_coupon"), cols.get("store_discount")]
    promo_cols = [c for c in promo_cols if c and c in df.columns]

    score = pd.Series(0.0, index=df.index)
    reason = pd.Series("", index=df.index, dtype="object")
    action = pd.Series("Review merchant setup and compare against similar stores.", index=df.index, dtype="object")

    if exp and exp_rate:
        avg = all_df[exp_rate].median(skipna=True)
        high_exp = df[exp] >= all_df[exp].quantile(0.65)
        low_conv = df[exp_rate] < avg
        mask = high_exp & low_conv
        score += high_exp.fillna(False) * 25 + low_conv.fillna(False) * 25
        reason[mask] = "曝光高，但曝光下单转化率低于悉尼中位数"
        action[mask] = "优先拜访：检查主图、菜单排序、爆品露出、配送费和活动组合。"

    if cart_rate and cart_order_rate:
        avg_cart_order = all_df[cart_order_rate].median(skipna=True)
        mask = (df[cart_rate] >= all_df[cart_rate].quantile(0.60)) & (df[cart_order_rate] < avg_cart_order)
        score += mask.fillna(False) * 20
        reason[mask & reason.eq("")] = "用户愿意加购，但加购到下单流失偏高"
        action[mask] = "重点优化成交：设置满减/运费减免/新客折扣，检查起送价和配送费。"

    if promo_cols:
        no_promo = pd.Series(True, index=df.index)
        for c in promo_cols:
            vals = df[c].astype(str).str.strip().str.lower()
            no_promo = no_promo & (~vals.isin(["是", "yes", "y", "1", "true", "有"]))
        score += no_promo * 15
        reason[no_promo & reason.eq("")] = "活动配置弱或缺失"
        action[no_promo] = "补齐活动：至少配置一个高感知活动，例如新客折扣、运费减免或全店折扣。"

    if material and material in df.columns:
        vals = df[material].astype(str).str.strip().str.lower()
        no_mat = ~vals.isin(["是", "yes", "y", "1", "true", "有", "合格"])
        score += no_mat * 10
        reason[no_mat & reason.eq("")] = "物料状态需要跟进"

    if gmv and gmv in df.columns:
        score += df[gmv].rank(pct=True).fillna(0) * 20
    if orders and orders in df.columns:
        score += df[orders].rank(pct=True).fillna(0) * 10

    reason[reason.eq("")] = "综合表现有优化空间"
    df["机会原因"] = reason
    df["建议动作"] = action
    df["机会分"] = score.round(1)
    return df


def build_display(df: pd.DataFrame, cols: Dict[str, Optional[str]], n=30) -> pd.DataFrame:
    wanted = [
        cols.get("merchant_name"), cols.get("merchant_id"), cols.get("area"), cols.get("owner_name"),
        cols.get("gmv"), cols.get("orders"), cols.get("exposure"), cols.get("visit"), cols.get("cart"),
        cols.get("exposure_rate"), cols.get("visit_rate"), cols.get("cart_rate"), cols.get("cart_order_rate"),
        "机会原因", "建议动作", "机会分"
    ]
    wanted = [c for c in wanted if c and c in df.columns]
    out = df.sort_values("机会分", ascending=False)[wanted].head(n).copy() if "机会分" in df.columns else df[wanted].head(n).copy()
    rename = {
        cols.get("merchant_name"): "店铺名称",
        cols.get("merchant_id"): "店铺ID",
        cols.get("area"): "商圈",
        cols.get("owner_name"): "BD",
        cols.get("gmv"): "GMV",
        cols.get("orders"): "订单",
        cols.get("exposure"): "曝光",
        cols.get("visit"): "进店",
        cols.get("cart"): "加购",
        cols.get("exposure_rate"): "曝光下单率",
        cols.get("visit_rate"): "曝光进店率",
        cols.get("cart_rate"): "进店加购率",
        cols.get("cart_order_rate"): "加购下单率",
    }
    out = out.rename(columns={k: v for k, v in rename.items() if k})
    return out


def group_table(df, by, cols):
    gmv, orders, exp_rate = cols.get("gmv"), cols.get("orders"), cols.get("exposure_rate")
    if not by or by not in df.columns:
        return pd.DataFrame()
    agg = {by: "count"}
    if gmv: agg[gmv] = "sum"
    if orders: agg[orders] = "sum"
    if exp_rate: agg[exp_rate] = "median"
    t = df.groupby(by, dropna=False).agg(agg).rename(columns={by: "店铺数"}).reset_index()
    rename = {by: "商圈/BD", gmv: "GMV", orders: "订单", exp_rate: "曝光下单率中位数"}
    t = t.rename(columns={k:v for k,v in rename.items() if k})
    if "GMV" in t.columns:
        t = t.sort_values("GMV", ascending=False)
    return t

# -----------------------------
# Sidebar controls
# -----------------------------
with st.sidebar:
    st.header("Upload")
    uploaded = st.file_uploader("Encrypted Excel file", type=["xlsx", "xls"])
    password = st.text_input("Password", type="password")
    st.caption("公司原始Excel不要上传到GitHub，只在这个网页上传分析。")

if not uploaded or not password:
    st.info("Upload your monthly encrypted Excel, enter password, then the dashboard will load.")
    st.subheader("V1.1 upgraded")
    st.write("- 自动识别 BD 姓名，不再只识别 bd工号")
    st.write("- 增加 BD 下拉选择")
    st.write("- Action Plan 增加店铺名称")
    st.write("- 增加商圈、BD排名、转化漏斗、活动/物料/拜访覆盖率")
    st.stop()

try:
    file_obj = decrypt_excel(uploaded, password)
    xls = pd.ExcelFile(file_obj)
    sheet = st.sidebar.selectbox("Sheet", xls.sheet_names, index=0)
    raw = pd.read_excel(xls, sheet_name=sheet)
    raw = clean_cols(raw)
    cols = detect_columns(raw)
    df = prepare(raw, cols)
except Exception as e:
    st.error(f"Could not read the file. Check the password or file format. Error: {e}")
    st.stop()

owners = owner_options(df, cols)
default_idx = owners.index("Yuan Dong") if "Yuan Dong" in owners else 0
selected_owner = st.sidebar.selectbox("BD / Owner", owners, index=default_idx)
my_df = filter_owner(df, cols, selected_owner)
planned = add_action_plan(my_df, df, cols)

st.success(f"Loaded {len(df):,} rows and {len(df.columns):,} columns. Selected: {selected_owner} ({len(my_df):,} merchants).")

# -----------------------------
# Overview
# -----------------------------
gmv, orders = cols.get("gmv"), cols.get("orders")
exp_rate, visit_rate, cart_rate, cart_order_rate = cols.get("exposure_rate"), cols.get("visit_rate"), cols.get("cart_rate"), cols.get("cart_order_rate")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Sydney Merchants", f"{len(df):,}")
c2.metric("Your Merchants", f"{len(my_df):,}")
c3.metric("Your GMV", money(my_df[gmv].sum()) if gmv else "-")
c4.metric("Your Orders", f"{my_df[orders].sum():,.0f}" if orders else "-")

c5, c6, c7, c8 = st.columns(4)
c5.metric("曝光下单率", pct(my_df[exp_rate].median()) if exp_rate else "-")
c6.metric("曝光进店率", pct(my_df[visit_rate].median()) if visit_rate else "-")
c7.metric("进店加购率", pct(my_df[cart_rate].median()) if cart_rate else "-")
c8.metric("加购下单率", pct(my_df[cart_order_rate].median()) if cart_order_rate else "-")

# -----------------------------
# Detected columns and owner sanity check
# -----------------------------
with st.expander("Detected Columns / 系统识别到的字段", expanded=False):
    st.json(cols)
    if cols.get("owner_name"):
        st.write("BD examples:", df[cols["owner_name"]].dropna().astype(str).unique()[:30].tolist())

# -----------------------------
# Main tabs
# -----------------------------
tab1, tab2, tab3, tab4, tab5 = st.tabs(["Top Action Plan", "商圈分析", "BD排名", "覆盖率", "下载/周报"])

with tab1:
    st.subheader("Top Action Plan / 优先跟进店铺")
    action_display = build_display(planned, cols, 50)
    st.dataframe(action_display, use_container_width=True, hide_index=True)

with tab2:
    st.subheader("商圈表现")
    area_tbl = group_table(my_df, cols.get("area"), cols)
    st.dataframe(area_tbl.head(50), use_container_width=True, hide_index=True)

with tab3:
    st.subheader("全悉尼 BD 排名")
    bd_tbl = group_table(df, cols.get("owner_name") or cols.get("owner_code"), cols)
    st.dataframe(bd_tbl.head(50), use_container_width=True, hide_index=True)

with tab4:
    st.subheader("活动 / 物料 / 拜访覆盖率")
    cover_rows = []
    mapping = {
        "单品折扣覆盖率": cols.get("single_discount"),
        "新客折扣覆盖率": cols.get("new_customer_discount"),
        "运费减免覆盖率": cols.get("delivery_coupon"),
        "全店折扣覆盖率": cols.get("store_discount"),
        "物料覆盖率": cols.get("material"),
        "拜访覆盖率": cols.get("visit_record"),
    }
    for name, col in mapping.items():
        if col and col in my_df.columns:
            cover_rows.append({"指标": name, "字段": col, "覆盖率": pct(yes_rate(my_df[col]))})
    st.dataframe(pd.DataFrame(cover_rows), use_container_width=True, hide_index=True)

with tab5:
    st.subheader("AI-style Weekly Summary")
    top_area = "-"
    if cols.get("area") and gmv and len(my_df):
        tmp = my_df.groupby(cols["area"])[gmv].sum().sort_values(ascending=False)
        if len(tmp): top_area = str(tmp.index[0])
    summary = f"This week, {selected_owner} has {len(my_df):,} merchants, total GMV {money(my_df[gmv].sum()) if gmv else '-'}, and total orders {my_df[orders].sum():,.0f} if order data is available. The highest GMV area is {top_area}. Next focus should be the Top Action Plan merchants, especially high-exposure low-conversion stores and merchants with weak promotion setup."
    st.write(summary)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        action_display.to_excel(writer, index=False, sheet_name="Top Action Plan")
        if len(my_df):
            my_df.to_excel(writer, index=False, sheet_name="Selected Merchants")
        if 'area_tbl' in locals() and not area_tbl.empty:
            area_tbl.to_excel(writer, index=False, sheet_name="Area Ranking")
        if 'bd_tbl' in locals() and not bd_tbl.empty:
            bd_tbl.to_excel(writer, index=False, sheet_name="BD Ranking")
    st.download_button("Download Full Analysis Excel", data=output.getvalue(), file_name="sydney_bd_analysis_v11.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
