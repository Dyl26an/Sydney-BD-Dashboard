import io
import re
from typing import Optional, List, Dict, Tuple

import pandas as pd
import streamlit as st

try:
    import msoffcrypto
except Exception:
    msoffcrypto = None

st.set_page_config(page_title="Sydney BD Growth Intelligence", layout="wide")

# -----------------------------
# Helpers
# -----------------------------

def norm(s: str) -> str:
    return re.sub(r"\s+", "", str(s).lower())


def find_col(df: pd.DataFrame, candidates: List[str], must_contain_any: Optional[List[str]] = None) -> Optional[str]:
    cols = list(df.columns)
    ncols = {c: norm(c) for c in cols}
    for cand in candidates:
        nc = norm(cand)
        for c, n in ncols.items():
            if n == nc or nc in n or n in nc:
                return c
    if must_contain_any:
        terms = [norm(x) for x in must_contain_any]
        for c, n in ncols.items():
            if any(t in n for t in terms):
                return c
    return None


def find_metric_col(df: pd.DataFrame, include_terms: List[str], exclude_terms: Optional[List[str]] = None) -> Optional[str]:
    exclude_terms = exclude_terms or []
    best = None
    best_score = -1
    for c in df.columns:
        n = norm(c)
        if any(norm(x) in n for x in exclude_terms):
            continue
        score = sum(1 for t in include_terms if norm(t) in n)
        if score > best_score and score > 0:
            best = c
            best_score = score
    return best


def to_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(
        s.astype(str).str.replace("$", "", regex=False).str.replace(",", "", regex=False).str.replace("%", "", regex=False),
        errors="coerce",
    ).fillna(0)


def decrypt_excel(uploaded_file, password: str) -> bytes:
    raw = uploaded_file.read()
    if not password:
        return raw
    if msoffcrypto is None:
        raise RuntimeError("msoffcrypto-tool is not installed. Please check requirements.txt")
    office_file = msoffcrypto.OfficeFile(io.BytesIO(raw))
    office_file.load_key(password=password)
    out = io.BytesIO()
    office_file.decrypt(out)
    return out.getvalue()


def read_excel(uploaded_file, password: str) -> pd.DataFrame:
    data = decrypt_excel(uploaded_file, password)
    xls = pd.ExcelFile(io.BytesIO(data), engine="openpyxl")
    sheet = xls.sheet_names[0]
    df = pd.read_excel(io.BytesIO(data), sheet_name=sheet, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def detect_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    return {
        "bd": find_col(df, ["bd工号", "bd姓名", "BD", "BD Name", "owner", "负责人", "业务员", "account manager"]),
        "merchant": find_col(df, ["店铺名称", "门店名称", "商家名称", "merchant name", "store name", "restaurant name", "name"]),
        "area": find_col(df, ["商圈", "区域", "城市区域", "area", "district", "suburb", "zone"]),
        "category": find_col(df, ["品类", "菜系", "category", "cuisine", "business type"]),
        "gmv": find_metric_col(df, ["gmv"], ["rate", "率", "%"]),
        "orders": find_metric_col(df, ["订单"], ["率", "rate", "%"]) or find_metric_col(df, ["orders"], ["rate", "%"]),
        "exposure": find_metric_col(df, ["曝光"], ["率", "rate", "%"]) or find_metric_col(df, ["impression", "exposure"], ["rate", "%"]),
        "visit": find_metric_col(df, ["进店"], ["率", "rate", "%"]) or find_metric_col(df, ["visit"], ["rate", "%"]),
        "cart": find_metric_col(df, ["加购"], ["率", "rate", "%"]) or find_metric_col(df, ["cart"], ["rate", "%"]),
        "material": find_col(df, ["物料", "material", "poster", "posm"]),
        "visit_record": find_col(df, ["拜访", "visit record", "last visit", "到访"]),
        "discount": find_col(df, ["折扣", "discount", "优惠", "coupon", "满减", "运费券", "活动"]),
    }


def add_metrics(df: pd.DataFrame, c: Dict[str, Optional[str]]) -> pd.DataFrame:
    d = df.copy()
    for key in ["gmv", "orders", "exposure", "visit", "cart"]:
        if c.get(key):
            d[f"__{key}"] = to_num(d[c[key]])
        else:
            d[f"__{key}"] = 0
    d["__exp_to_order"] = d["__orders"] / d["__exposure"].replace(0, pd.NA)
    d["__exp_to_visit"] = d["__visit"] / d["__exposure"].replace(0, pd.NA)
    d["__visit_to_cart"] = d["__cart"] / d["__visit"].replace(0, pd.NA)
    d["__cart_to_order"] = d["__orders"] / d["__cart"].replace(0, pd.NA)
    for col in ["__exp_to_order", "__exp_to_visit", "__visit_to_cart", "__cart_to_order"]:
        d[col] = d[col].fillna(0).clip(lower=0, upper=10)
    return d


def coverage_rate(s: pd.Series) -> float:
    if s is None or len(s) == 0:
        return 0.0
    text = s.astype(str).str.strip().str.lower()
    missing = text.isin(["", "0", "nan", "none", "否", "无", "no", "false", "n"])
    return float((~missing).mean())


def bd_ranking(d: pd.DataFrame, c: Dict[str, Optional[str]]) -> pd.DataFrame:
    bd_col = c["bd"]
    if not bd_col:
        return pd.DataFrame()
    rows = []
    for bd, g in d.groupby(bd_col, dropna=False):
        if str(bd).strip() in ["", "nan", "None"]:
            continue
        exposure = g["__exposure"].sum()
        orders = g["__orders"].sum()
        row = {
            "BD": str(bd),
            "Merchants": len(g),
            "GMV": g["__gmv"].sum(),
            "Orders": orders,
            "Exposure": exposure,
            "GMV / Store": g["__gmv"].sum() / max(len(g), 1),
            "Exposure → Order": orders / exposure if exposure else 0,
            "Avg Store Conversion": g["__exp_to_order"].mean(),
        }
        if c.get("material"):
            row["Material Coverage"] = coverage_rate(g[c["material"]])
        if c.get("visit_record"):
            row["Visit Coverage"] = coverage_rate(g[c["visit_record"]])
        if c.get("discount"):
            row["Promotion Coverage"] = coverage_rate(g[c["discount"]])
        rows.append(row)
    r = pd.DataFrame(rows)
    if len(r):
        r = r.sort_values("GMV", ascending=False).reset_index(drop=True)
        r.insert(0, "Rank", range(1, len(r) + 1))
    return r


def opportunity_list(d: pd.DataFrame, c: Dict[str, Optional[str]], selected_bd: Optional[str], top_n=30) -> pd.DataFrame:
    g = d.copy()
    if selected_bd and c.get("bd"):
        g = g[g[c["bd"]].astype(str) == str(selected_bd)]
    if len(g) == 0:
        return pd.DataFrame()
    exp_q = g["__exposure"].quantile(0.65) if g["__exposure"].sum() else 0
    conv_med = d["__exp_to_order"].median() if len(d) else 0
    g["Opportunity Score"] = 0.0
    g.loc[g["__exposure"] >= exp_q, "Opportunity Score"] += 35
    g.loc[g["__exp_to_order"] < conv_med, "Opportunity Score"] += 35
    g["Opportunity Score"] += (g["__gmv"].rank(pct=True) * 20).fillna(0)
    if c.get("material"):
        mat_missing = g[c["material"]].astype(str).str.strip().str.lower().isin(["", "0", "nan", "none", "否", "无", "no", "false", "n"])
        g.loc[mat_missing, "Opportunity Score"] += 10
    reason = []
    action = []
    for _, row in g.iterrows():
        rs = []
        ac = []
        if row["__exposure"] >= exp_q and row["__exp_to_order"] < conv_med:
            rs.append("曝光高但下单转化低")
            ac.append("优先检查首图、菜单排序、爆品套餐和满减/运费券")
        if row["__visit"] > 0 and row["__cart"] / max(row["__visit"], 1) < d["__visit_to_cart"].median():
            rs.append("进店后加购偏弱")
            ac.append("优化菜单前10项、图片、价格带和套餐命名")
        if row["__cart"] > 0 and row["__orders"] / max(row["__cart"], 1) < d["__cart_to_order"].median():
            rs.append("加购后下单偏弱")
            ac.append("补优惠券、满减门槛、配送费补贴，降低最后一步流失")
        reason.append("；".join(rs) or "综合潜力较高")
        action.append("；".join(ac) or "安排拜访，复制同品类高转化店的活动与菜单结构")
    g["Reason"] = reason
    g["Recommended Action"] = action
    cols = []
    rename = {}
    for key, label in [("merchant", "Store Name"), ("bd", "BD"), ("area", "Area"), ("category", "Category")]:
        if c.get(key):
            cols.append(c[key]); rename[c[key]] = label
    cols += ["__gmv", "__orders", "__exposure", "__exp_to_order", "Opportunity Score", "Reason", "Recommended Action"]
    out = g.sort_values("Opportunity Score", ascending=False)[cols].head(top_n).rename(columns=rename)
    out = out.rename(columns={"__gmv": "GMV", "__orders": "Orders", "__exposure": "Exposure", "__exp_to_order": "Exposure → Order"})
    return out


def best_practice(d: pd.DataFrame, c: Dict[str, Optional[str]], selected_bd: Optional[str]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    top = d.sort_values("__exp_to_order", ascending=False).head(100)
    mine = d[d[c["bd"]].astype(str) == str(selected_bd)] if selected_bd and c.get("bd") else pd.DataFrame()
    rows = []
    for label, data in [("Top 100 conversion stores", top), ("My stores", mine)]:
        if len(data) == 0: continue
        row = {"Group": label, "Stores": len(data), "Avg GMV": data["__gmv"].mean(), "Avg Orders": data["__orders"].mean(), "Avg Conversion": data["__exp_to_order"].mean()}
        if c.get("material"): row["Material Coverage"] = coverage_rate(data[c["material"]])
        if c.get("visit_record"): row["Visit Coverage"] = coverage_rate(data[c["visit_record"]])
        if c.get("discount"): row["Promotion Coverage"] = coverage_rate(data[c["discount"]])
        rows.append(row)
    top_cols = []
    rename = {}
    for key, label in [("merchant", "Store Name"), ("bd", "BD"), ("area", "Area"), ("category", "Category")]:
        if c.get(key): top_cols.append(c[key]); rename[c[key]] = label
    top_cols += ["__gmv", "__orders", "__exposure", "__exp_to_order"]
    top_stores = top[top_cols].rename(columns=rename).rename(columns={"__gmv":"GMV", "__orders":"Orders", "__exposure":"Exposure", "__exp_to_order":"Exposure → Order"})
    return pd.DataFrame(rows), top_stores


def format_df(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in out.columns:
        if col in ["GMV", "GMV / Store", "Avg GMV"]:
            out[col] = out[col].map(lambda x: f"${x:,.0f}" if pd.notna(x) else "")
        elif col in ["Exposure → Order", "Avg Store Conversion", "Avg Conversion", "Material Coverage", "Visit Coverage", "Promotion Coverage"]:
            out[col] = out[col].map(lambda x: f"{x:.2%}" if pd.notna(x) else "")
        elif col in ["Opportunity Score"]:
            out[col] = out[col].map(lambda x: f"{x:.1f}" if pd.notna(x) else "")
    return out


def make_excel(ranking, opp, pattern, topstores) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        ranking.to_excel(writer, index=False, sheet_name="BD Ranking")
        opp.to_excel(writer, index=False, sheet_name="Action Plan")
        pattern.to_excel(writer, index=False, sheet_name="Best Practice")
        topstores.to_excel(writer, index=False, sheet_name="Top Stores")
    return output.getvalue()

# -----------------------------
# UI
# -----------------------------

st.title("Sydney BD Growth Intelligence")
st.caption("V2: BD Ranking · Compare Me · Learn From Best · Action Plan")

with st.sidebar:
    uploaded = st.file_uploader("Upload encrypted Excel", type=["xlsx", "xlsm", "xls"])
    password = st.text_input("Password", type="password")
    analyse = st.button("Analyse", type="primary")

if not uploaded:
    st.info("Upload your monthly Excel file, enter password, then click Analyse.")
    st.stop()

if analyse or uploaded:
    try:
        df = read_excel(uploaded, password)
        cols = detect_columns(df)
        d = add_metrics(df, cols)
    except Exception as e:
        st.error(f"Could not read this file: {e}")
        st.stop()

    st.success(f"Loaded {len(df):,} rows and {len(df.columns):,} columns")
    with st.expander("Detected columns"):
        st.json(cols)

    bd_col = cols.get("bd")
    bd_options = []
    if bd_col:
        bd_options = sorted([x for x in d[bd_col].dropna().astype(str).unique().tolist() if x.strip() and x.strip().lower() != "nan"])
    selected_bd = st.sidebar.selectbox("Choose BD / Owner", bd_options, index=0 if bd_options else None) if bd_options else None

    total_gmv = d["__gmv"].sum()
    total_orders = d["__orders"].sum()
    total_exp = d["__exposure"].sum()
    my = d[d[bd_col].astype(str) == str(selected_bd)] if selected_bd and bd_col else pd.DataFrame()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Sydney GMV", f"${total_gmv:,.0f}")
    c2.metric("Sydney Orders", f"{total_orders:,.0f}")
    c3.metric("Sydney Conversion", f"{(total_orders/total_exp if total_exp else 0):.2%}")
    c4.metric("Merchants", f"{len(d):,}")

    if selected_bd and len(my):
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("My Merchants", f"{len(my):,}")
        m2.metric("My GMV", f"${my['__gmv'].sum():,.0f}")
        m3.metric("My Orders", f"{my['__orders'].sum():,.0f}")
        m4.metric("My Conversion", f"{(my['__orders'].sum()/my['__exposure'].sum() if my['__exposure'].sum() else 0):.2%}")

    ranking = bd_ranking(d, cols)
    opp = opportunity_list(d, cols, selected_bd, 30)
    pattern, topstores = best_practice(d, cols, selected_bd)

    tab1, tab2, tab3, tab4 = st.tabs(["🏆 BD Ranking", "🔍 Compare Me", "📍 Learn From Best", "🎯 Action Plan"])

    with tab1:
        st.subheader("Sydney BD Ranking")
        st.dataframe(format_df(ranking), use_container_width=True, hide_index=True)

    with tab2:
        st.subheader("Compare Me vs Top BD")
        if selected_bd and len(ranking):
            top_bd = ranking.iloc[0]
            me_row = ranking[ranking["BD"] == str(selected_bd)]
            if len(me_row):
                me_row = me_row.iloc[0]
                compare_rows = []
                for metric in ["GMV", "Merchants", "GMV / Store", "Exposure → Order", "Material Coverage", "Visit Coverage", "Promotion Coverage"]:
                    if metric in ranking.columns:
                        compare_rows.append({"Metric": metric, "Me": me_row.get(metric, 0), "Top BD": top_bd.get(metric, 0), "Gap": me_row.get(metric, 0) - top_bd.get(metric, 0)})
                comp = pd.DataFrame(compare_rows)
                st.dataframe(format_df(comp), use_container_width=True, hide_index=True)
                st.markdown("### Quick diagnosis")
                weak = []
                for _, r in comp.iterrows():
                    if isinstance(r["Me"], (int, float)) and isinstance(r["Top BD"], (int, float)) and r["Top BD"] and r["Me"] < r["Top BD"] * 0.85:
                        weak.append(r["Metric"])
                if weak:
                    st.write("Your biggest gaps are: **" + "**, **".join(weak[:3]) + "**. Focus on these before chasing more merchants.")
                else:
                    st.write("You are close to top BD on the main measurable metrics. Next step is store-level optimization.")
            else:
                st.warning("Selected BD was not found in ranking.")
        else:
            st.warning("No BD column detected.")

    with tab3:
        st.subheader("Best Practice Pattern")
        st.dataframe(format_df(pattern), use_container_width=True, hide_index=True)
        st.subheader("Top conversion stores to learn from")
        st.dataframe(format_df(topstores.head(50)), use_container_width=True, hide_index=True)

    with tab4:
        st.subheader("Priority Action Plan")
        st.dataframe(format_df(opp), use_container_width=True, hide_index=True)
        if len(opp):
            st.markdown("### This week focus")
            st.write("Start with the top 10 stores. The best batch move is: fix menu first, add/adjust promotion second, then visit stores with high exposure but low conversion.")

    report = make_excel(ranking, opp, pattern, topstores)
    st.download_button("Download full Excel report", report, file_name="sydney_bd_growth_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
