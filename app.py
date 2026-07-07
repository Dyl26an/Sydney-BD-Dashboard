import io
import re
from typing import Optional, List, Dict, Tuple

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

try:
    import msoffcrypto
except Exception:
    msoffcrypto = None

st.set_page_config(page_title="Sydney Growth Intelligence V6", layout="wide")

# -----------------------------
# Helpers
# -----------------------------
def norm(s: str) -> str:
    return re.sub(r"\s+", "", str(s).lower())

def find_col(df: pd.DataFrame, candidates: List[str], must_numeric: bool = False) -> Optional[str]:
    cols = list(df.columns)
    nmap = {norm(c): c for c in cols}
    for cand in candidates:
        if norm(cand) in nmap:
            col = nmap[norm(cand)]
            if not must_numeric or pd.to_numeric(df[col], errors="coerce").notna().sum() > 0:
                return col
    # fuzzy contains
    for cand in candidates:
        nc = norm(cand)
        for c in cols:
            if nc and nc in norm(c):
                if not must_numeric or pd.to_numeric(df[c], errors="coerce").notna().sum() > 0:
                    return c
    return None

def to_num(s):
    if s is None:
        return pd.Series(dtype=float)
    return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False).str.replace("%", "", regex=False).str.replace("$", "", regex=False), errors="coerce")

def pct_series(s):
    x = to_num(s)
    # if source stores 0.015 as 1.5%, keep; if 1.5 means 1.5%, divide by 100 later for display? We store as fraction.
    if x.dropna().median() > 1:
        x = x / 100.0
    return x

def weighted_rate(df, rate_col, weight_col):
    if not rate_col or not weight_col:
        return np.nan
    r = pct_series(df[rate_col])
    w = to_num(df[weight_col]).fillna(0)
    mask = r.notna() & w.gt(0)
    if not mask.any():
        return np.nan
    return float((r[mask] * w[mask]).sum() / w[mask].sum())

def money(x):
    if pd.isna(x): return "-"
    return f"${x:,.0f}"

def pct(x):
    if pd.isna(x): return "-"
    return f"{x*100:.2f}%"

def safe_div(a, b):
    return np.nan if not b or pd.isna(b) or b == 0 else a / b

def decrypt_excel(uploaded, password: str):
    raw = uploaded.read()
    bio = io.BytesIO(raw)
    if password and msoffcrypto is not None:
        try:
            office = msoffcrypto.OfficeFile(bio)
            office.load_key(password=password)
            decrypted = io.BytesIO()
            office.decrypt(decrypted)
            decrypted.seek(0)
            return pd.read_excel(decrypted, sheet_name=None, engine="openpyxl")
        except Exception:
            pass
    bio.seek(0)
    return pd.read_excel(bio, sheet_name=None, engine="openpyxl")

@st.cache_data(show_spinner=False)
def prepare_data_from_bytes(raw: bytes, password: str):
    bio = io.BytesIO(raw)
    if password and msoffcrypto is not None:
        try:
            office = msoffcrypto.OfficeFile(bio)
            office.load_key(password=password)
            decrypted = io.BytesIO()
            office.decrypt(decrypted)
            decrypted.seek(0)
            sheets = pd.read_excel(decrypted, sheet_name=None, engine="openpyxl")
        except Exception:
            bio.seek(0)
            sheets = pd.read_excel(bio, sheet_name=None, engine="openpyxl")
    else:
        sheets = pd.read_excel(bio, sheet_name=None, engine="openpyxl")
    df = next(iter(sheets.values()))
    df.columns = [str(c).strip() for c in df.columns]
    return df

def detect_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    c = {}
    c["month"] = find_col(df, ["月份", "月", "统计月份", "reporting month", "month", "date"])
    c["merchant"] = find_col(df, ["商户名称", "店铺名称", "门店名称", "merchant name", "restaurant name", "name"])
    c["merchant_id"] = find_col(df, ["商户id", "门店id", "merchant id", "shop id", "store id"])
    c["bd_name"] = find_col(df, ["bd姓名", "BD姓名", "负责人姓名", "BD Name", "owner name", "负责人"])
    c["bd_code"] = find_col(df, ["bd工号", "BD工号", "bd", "owner", "BD"])
    c["area"] = find_col(df, ["商圈", "区域", "area", "suburb", "zone"])
    c["category"] = find_col(df, ["品类", "菜系", "主营品类", "category", "cuisine"])
    c["level"] = find_col(df, ["商户等级", "店铺等级", "等级", "merchant level", "level"])
    c["gmv"] = find_col(df, ["GMV", "交易额", "营业额", "销售额", "实付金额"], True)
    c["orders"] = find_col(df, ["订单数_排除mm的均单", "订单数", "有效订单", "orders", "order count"], True)
    c["exposure"] = find_col(df, ["平均曝光人数", "曝光人数", "曝光", "exposure", "impression"], True)
    c["visit"] = find_col(df, ["平均进店人数", "进店人数", "进店", "visit", "store visit"], True)
    c["cart"] = find_col(df, ["平均加购人数", "加购人数", "加购", "cart", "add to cart"], True)
    c["e2o_rate"] = find_col(df, ["曝光下单转化率", "曝光-下单", "exposure order", "impression order"], True)
    c["e2v_rate"] = find_col(df, ["曝光进店转化率", "曝光-进店", "exposure visit", "impression visit"], True)
    c["v2c_rate"] = find_col(df, ["进店加购转化率", "进店-加购", "visit cart"], True)
    c["c2o_rate"] = find_col(df, ["加购下单转化率", "加购-下单", "cart order"], True)
    c["promo"] = find_col(df, ["是否有活动", "活动", "折扣", "优惠", "promo", "promotion", "campaign"])
    c["material"] = find_col(df, ["物料", "是否有物料", "material"])
    c["visit_record"] = find_col(df, ["拜访记录", "拜访", "visit record"])
    c["aov"] = find_col(df, ["客单价", "均单价", "AOV", "average order value"], True)
    return c

def add_metrics(df: pd.DataFrame, c: Dict[str, Optional[str]]) -> pd.DataFrame:
    out = df.copy()
    def n(key):
        return to_num(out[c[key]]) if c.get(key) else pd.Series(np.nan, index=out.index)
    out["_merchant"] = out[c["merchant"]].astype(str) if c.get("merchant") else out.index.astype(str)
    out["_bd"] = out[c["bd_name"]].astype(str) if c.get("bd_name") else (out[c["bd_code"]].astype(str) if c.get("bd_code") else "Unknown")
    out["_area"] = out[c["area"]].fillna("Unknown").astype(str) if c.get("area") else "Unknown"
    out["_category"] = out[c["category"]].fillna("Unknown").astype(str) if c.get("category") else "Unknown"
    out["_level"] = out[c["level"]].fillna("Unknown").astype(str) if c.get("level") else "Auto"
    out["_gmv"] = n("gmv").fillna(0)
    out["_orders"] = n("orders").fillna(0)
    out["_exposure"] = n("exposure")
    out["_visit"] = n("visit")
    out["_cart"] = n("cart")
    out["_aov"] = n("aov") if c.get("aov") else out["_gmv"] / out["_orders"].replace(0, np.nan)
    for k, new in [("e2o_rate","_e2o"),("e2v_rate","_e2v"),("v2c_rate","_v2c"),("c2o_rate","_c2o")]:
        out[new] = pct_series(out[c[k]]) if c.get(k) else np.nan
    # fallback direct ratios only if same-period raw counts make sense, capped to reasonable range
    out["_e2v"] = out["_e2v"].fillna(out["_visit"] / out["_exposure"].replace(0, np.nan))
    out["_v2c"] = out["_v2c"].fillna(out["_cart"] / out["_visit"].replace(0, np.nan))
    out["_e2o"] = out["_e2o"].where(out["_e2o"].between(0,1), np.nan)
    out["_e2v"] = out["_e2v"].where(out["_e2v"].between(0,1), np.nan)
    out["_v2c"] = out["_v2c"].where(out["_v2c"].between(0,1), np.nan)
    out["_c2o"] = out["_c2o"].where(out["_c2o"].between(0,1), np.nan)
    def flag(col):
        if not col:
            return pd.Series(False, index=out.index)
        s = out[col].astype(str).str.lower().str.strip()
        return ~(s.isin(["", "0", "nan", "none", "no", "否", "无", "false"]))
    out["_promo"] = flag(c.get("promo"))
    out["_material"] = flag(c.get("material"))
    out["_visit_record"] = flag(c.get("visit_record"))
    if not c.get("level"):
        q = out["_gmv"].quantile([.25,.5,.75]).to_dict()
        out["_level"] = pd.cut(out["_gmv"], bins=[-np.inf,q.get(.25,0),q.get(.5,0),q.get(.75,0),np.inf], labels=["D","C","B","A"]).astype(str)
    # health score
    gmv_score = out["_gmv"].rank(pct=True).fillna(0) * 30
    conv_score = out["_e2o"].rank(pct=True).fillna(0) * 25
    funnel_score = out["_e2v"].rank(pct=True).fillna(0)*10 + out["_v2c"].rank(pct=True).fillna(0)*10 + out["_c2o"].rank(pct=True).fillna(0)*10
    op_score = out["_promo"].astype(int)*5 + out["_material"].astype(int)*5 + out["_visit_record"].astype(int)*5
    out["_health_score"] = (gmv_score + conv_score + funnel_score + op_score).clip(0,100)
    return out

def summary_metrics(df: pd.DataFrame) -> Dict[str, float]:
    return {
        "gmv": df["_gmv"].sum(),
        "orders": df["_orders"].sum(),
        "merchants": df["_merchant"].nunique(),
        "e2o": weighted_avg(df, "_e2o", "_exposure"),
        "e2v": weighted_avg(df, "_e2v", "_exposure"),
        "v2c": weighted_avg(df, "_v2c", "_visit"),
        "c2o": weighted_avg(df, "_c2o", "_cart"),
        "promo_rate": df["_promo"].mean(),
        "material_rate": df["_material"].mean(),
        "visit_rate": df["_visit_record"].mean(),
    }

def weighted_avg(df, val_col, weight_col):
    v = df[val_col]
    w = df[weight_col].fillna(0)
    m = v.notna() & w.gt(0)
    if m.any(): return float((v[m]*w[m]).sum()/w[m].sum())
    return float(v.mean()) if v.notna().any() else np.nan

def learning_candidates(df: pd.DataFrame, target_idx: int) -> pd.DataFrame:
    target = df.loc[target_idx]
    cand = df.drop(index=target_idx).copy()
    if cand.empty: return cand
    # similarity pieces
    cat = (cand["_category"] == target["_category"]).astype(float)
    area = (cand["_area"] == target["_area"]).astype(float)
    level = (cand["_level"] == target["_level"]).astype(float)
    def closeness(col):
        t = target[col]
        x = cand[col]
        if pd.isna(t) or t == 0:
            return pd.Series(0.5, index=cand.index)
        return (1 - (np.log1p(x.fillna(0)) - np.log1p(t)) .abs() / 3).clip(0,1)
    exposure_close = closeness("_exposure")
    aov_close = closeness("_aov")
    # performance score: learn from merchants that are stronger than target
    perf = (cand["_gmv"].rank(pct=True).fillna(0)*0.45 + cand["_e2o"].rank(pct=True).fillna(0)*0.35 + cand["_health_score"].rank(pct=True).fillna(0)*0.20)
    sim = cat*0.35 + level*0.15 + area*0.15 + exposure_close*0.15 + aov_close*0.10 + perf*0.10
    # boost if better than target
    better = ((cand["_gmv"] > target["_gmv"]).astype(float)*0.05 + (cand["_e2o"] > target["_e2o"]).astype(float)*0.05)
    cand["Learning Score"] = ((sim + better).clip(0,1) * 100).round(1)
    cand["Why learn"] = np.select(
        [cand["_category"].eq(target["_category"]) & cand["_gmv"].gt(target["_gmv"]), cand["_e2o"].gt(target["_e2o"]), cand["_health_score"].gt(target["_health_score"])],
        ["同品类且GMV更高", "转化率更强", "综合健康度更高"],
        default="相似店铺，可对比"
    )
    return cand.sort_values("Learning Score", ascending=False)

def make_gap_table(target: pd.Series, top5: pd.DataFrame) -> pd.DataFrame:
    rows = []
    specs = [
        ("GMV", "_gmv", money, "高"),
        ("Orders", "_orders", lambda x: f"{x:,.0f}" if pd.notna(x) else "-", "高"),
        ("Exposure → Visit", "_e2v", pct, "高"),
        ("Visit → Cart", "_v2c", pct, "高"),
        ("Cart → Order", "_c2o", pct, "高"),
        ("Exposure → Order", "_e2o", pct, "高"),
        ("Promo", "_promo", lambda x: "Yes" if x else "No", "有"),
        ("Material", "_material", lambda x: "Yes" if x else "No", "有"),
        ("Visit Record", "_visit_record", lambda x: "Yes" if x else "No", "有"),
    ]
    for label, col, fmt, _ in specs:
        if col in ["_promo","_material","_visit_record"]:
            mine = bool(target[col])
            avg = float(top5[col].mean()) if len(top5) else np.nan
            gap = (1.0 if mine else 0.0) - avg if pd.notna(avg) else np.nan
            rows.append({"Metric": label, "My merchant": fmt(mine), "Top5 avg": pct(avg), "Gap": f"{gap*100:+.1f} pp" if pd.notna(gap) else "-"})
        else:
            mine = target[col]
            avg = top5[col].mean() if len(top5) else np.nan
            gap = mine - avg if pd.notna(mine) and pd.notna(avg) else np.nan
            rows.append({"Metric": label, "My merchant": fmt(mine), "Top5 avg": fmt(avg), "Gap": (money(gap) if col == "_gmv" else (f"{gap:+,.0f}" if col == "_orders" else f"{gap*100:+.2f} pp" if pd.notna(gap) else "-"))})
    return pd.DataFrame(rows)

def action_plan(target: pd.Series, top5: pd.DataFrame) -> List[str]:
    plans = []
    def avg(col): return top5[col].mean() if len(top5) else np.nan
    if pd.notna(avg("_v2c")) and pd.notna(target["_v2c"]) and target["_v2c"] < avg("_v2c")*0.9:
        plans.append("优先提升 Visit → Cart：检查菜单首屏、爆品排序、套餐组合和图片吸引力。")
    if pd.notna(avg("_c2o")) and pd.notna(target["_c2o"]) and target["_c2o"] < avg("_c2o")*0.9:
        plans.append("提升 Cart → Order：增加满减、运费券或下单门槛更低的套餐，减少用户加购后流失。")
    if pd.notna(avg("_e2v")) and pd.notna(target["_e2v"]) and target["_e2v"] < avg("_e2v")*0.9:
        plans.append("提升 Exposure → Visit：优化店铺封面、店名展示、评分/配送时间和活动标签露出。")
    if not bool(target["_promo"]): plans.append("补活动：至少配置一个主活动，例如折扣、满减、新客券或运费券。")
    if not bool(target["_material"]): plans.append("补物料：争取官方物料/店内展示/平台视觉资源，提高用户信任感。")
    if not bool(target["_visit_record"]): plans.append("安排一次线下拜访：用同品类 Top5 店铺作为案例，和商户讨论菜单、套餐、活动。")
    if not plans:
        plans.append("这家店基础表现不错。下一步建议学习 Top5 的套餐结构和高转化 SKU，做精细化提升。")
    return plans[:6]

# -----------------------------
# UI
# -----------------------------
st.title("Sydney Growth Intelligence V6")
st.caption("新增 Merchant AI Coach：输入店铺名，找到同品类最值得学习的 Top 5，并生成提升建议。")

with st.sidebar:
    st.header("Upload")
    uploaded = st.file_uploader("Upload monthly Excel", type=["xlsx", "xls"])
    password = st.text_input("Excel password", type="password")

if not uploaded:
    st.info("请先上传月度 Excel。")
    st.stop()

raw = uploaded.getvalue()
try:
    df0 = prepare_data_from_bytes(raw, password)
except Exception as e:
    st.error(f"文件读取失败：{e}")
    st.stop()

cols = detect_columns(df0)
df = add_metrics(df0, cols)

month = "Unknown"
if cols.get("month"):
    mvals = df0[cols["month"]].dropna().astype(str).unique()
    if len(mvals): month = mvals[0]

st.subheader(f"Reporting Month: {month}")

# Global filters
with st.expander("Global filters", expanded=False):
    lvls = sorted([x for x in df["_level"].dropna().unique()])
    areas = sorted([x for x in df["_area"].dropna().unique()])
    selected_lvls = st.multiselect("Merchant Level", lvls, default=lvls)
    selected_areas = st.multiselect("Area", areas, default=areas)
filtered = df[df["_level"].isin(selected_lvls) & df["_area"].isin(selected_areas)].copy()

# Top metrics
s = summary_metrics(filtered)
c1,c2,c3,c4,c5 = st.columns(5)
c1.metric("GMV", money(s["gmv"]))
c2.metric("Orders", f"{s['orders']:,.0f}")
c3.metric("Merchants", f"{s['merchants']:,.0f}")
c4.metric("Exposure → Order", pct(s["e2o"]))
c5.metric("Cart → Order", pct(s["c2o"]))

tabs = st.tabs(["🏆 Overview", "🔍 Merchant AI Coach", "⭐ Learn From Best", "🎯 Opportunity Finder", "📖 Metric Dictionary"])

with tabs[0]:
    st.markdown("### BD ranking")
    bd = filtered.groupby("_bd").agg(GMV=("_gmv","sum"), Orders=("_orders","sum"), Merchants=("_merchant","nunique"), Avg_Health=("_health_score","mean")).reset_index()
    bd["GMV per Store"] = bd["GMV"] / bd["Merchants"].replace(0,np.nan)
    bd = bd.sort_values("GMV", ascending=False)
    st.dataframe(bd.head(20).style.format({"GMV":"${:,.0f}", "Orders":"{:,.0f}", "GMV per Store":"${:,.0f}", "Avg_Health":"{:.1f}"}), use_container_width=True)
    if len(bd):
        fig = px.bar(bd.head(10).sort_values("GMV"), x="GMV", y="_bd", orientation="h", title="Top 10 BD by GMV")
        st.plotly_chart(fig, use_container_width=True)

with tabs[1]:
    st.markdown("### Merchant AI Coach")
    names = filtered["_merchant"].dropna().astype(str).unique().tolist()
    query = st.text_input("输入你的店铺名称", placeholder="例如：东北咱家菜")
    matches = []
    if query:
        nq = norm(query)
        matches = [n for n in names if nq in norm(n)][:30]
    choice = st.selectbox("选择匹配店铺", matches if matches else names[:200])
    if choice:
        target_idx = filtered[filtered["_merchant"].astype(str)==choice].index[0]
        target = filtered.loc[target_idx]
        candidates = learning_candidates(filtered, target_idx)
        # prefer same category for top5 if enough
        same_cat = candidates[candidates["_category"] == target["_category"]]
        top5 = (same_cat if len(same_cat)>=5 else candidates).head(5)
        st.markdown(f"#### {target['_merchant']} — Merchant Profile")
        a,b,c,d,e = st.columns(5)
        a.metric("GMV", money(target["_gmv"]))
        b.metric("Orders", f"{target['_orders']:,.0f}")
        c.metric("Category", str(target["_category"]))
        d.metric("Area", str(target["_area"]))
        e.metric("Level", str(target["_level"]))

        st.markdown("#### Top 5 learning merchants")
        show = top5[["_merchant","_bd","_area","_category","_level","_gmv","_orders","_e2o","_e2v","_v2c","_c2o","Learning Score","Why learn"]].rename(columns={"_merchant":"Merchant","_bd":"BD","_area":"Area","_category":"Category","_level":"Level","_gmv":"GMV","_orders":"Orders","_e2o":"Exposure → Order","_e2v":"Exposure → Visit","_v2c":"Visit → Cart","_c2o":"Cart → Order"})
        st.dataframe(show.style.format({"GMV":"${:,.0f}", "Orders":"{:,.0f}", "Exposure → Order":"{:.2%}", "Exposure → Visit":"{:.2%}", "Visit → Cart":"{:.2%}", "Cart → Order":"{:.2%}", "Learning Score":"{:.1f}"}), use_container_width=True)

        st.markdown("#### Gap analysis")
        gap = make_gap_table(target, top5)
        st.dataframe(gap, use_container_width=True)

        st.markdown("#### Funnel comparison")
        chart_df = pd.DataFrame({
            "Metric":["Exposure → Visit","Visit → Cart","Cart → Order","Exposure → Order"],
            "My merchant":[target["_e2v"],target["_v2c"],target["_c2o"],target["_e2o"]],
            "Top5 avg":[top5["_e2v"].mean(),top5["_v2c"].mean(),top5["_c2o"].mean(),top5["_e2o"].mean()],
        })
        fig = px.bar(chart_df.melt(id_vars="Metric", var_name="Group", value_name="Rate"), x="Metric", y="Rate", color="Group", barmode="group", title="My merchant vs Top5 avg conversion")
        fig.update_yaxes(tickformat=".1%")
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### Best practice checklist")
        bp = pd.DataFrame({
            "Item":["Promotion", "Material", "Visit record"],
            "My merchant":["Yes" if target["_promo"] else "No", "Yes" if target["_material"] else "No", "Yes" if target["_visit_record"] else "No"],
            "Top5 coverage":[pct(top5["_promo"].mean()), pct(top5["_material"].mean()), pct(top5["_visit_record"].mean())],
            "Priority":["High" if (not target["_promo"] and top5["_promo"].mean()>0.5) else "Medium", "High" if (not target["_material"] and top5["_material"].mean()>0.5) else "Medium", "High" if (not target["_visit_record"] and top5["_visit_record"].mean()>0.5) else "Medium"]
        })
        st.dataframe(bp, use_container_width=True)

        st.markdown("#### AI Action Plan")
        for i, p in enumerate(action_plan(target, top5), 1):
            st.write(f"{i}. {p}")

        st.markdown("#### Visit Brief")
        brief = f"""
Merchant Visit Brief — {target['_merchant']}

Reporting month: {month}
BD: {target['_bd']}
Area: {target['_area']}
Category: {target['_category']}
Level: {target['_level']}

Current performance:
- GMV: {money(target['_gmv'])}
- Orders: {target['_orders']:,.0f}
- Exposure → Visit: {pct(target['_e2v'])}
- Visit → Cart: {pct(target['_v2c'])}
- Cart → Order: {pct(target['_c2o'])}
- Exposure → Order: {pct(target['_e2o'])}

Top merchants to learn from:
{chr(10).join([f"- {r['_merchant']} ({r['_bd']}): Learning Score {r['Learning Score']} — {r['Why learn']}" for _, r in top5.iterrows()])}

Recommended discussion points:
{chr(10).join([f"{i}. {p}" for i, p in enumerate(action_plan(target, top5), 1)])}
"""
        st.download_button("Download Visit Brief", brief, file_name=f"visit_brief_{re.sub(r'[^A-Za-z0-9]+','_',str(target['_merchant']))}.txt")

with tabs[2]:
    st.markdown("### Learn From Best")
    top_merch = filtered.sort_values(["_health_score","_gmv"], ascending=False).head(50)
    st.dataframe(top_merch[["_merchant","_bd","_area","_category","_level","_gmv","_orders","_e2o","_health_score"]].rename(columns={"_merchant":"Merchant","_bd":"BD","_area":"Area","_category":"Category","_level":"Level","_gmv":"GMV","_orders":"Orders","_e2o":"Exposure → Order","_health_score":"Health Score"}).style.format({"GMV":"${:,.0f}", "Orders":"{:,.0f}", "Exposure → Order":"{:.2%}", "Health Score":"{:.1f}"}), use_container_width=True)
    cat = st.selectbox("Choose category for best practice", sorted(filtered["_category"].dropna().unique().tolist()))
    sub = filtered[filtered["_category"]==cat]
    top = sub.sort_values("_gmv", ascending=False).head(max(10, min(50, len(sub))))
    st.write(f"**{cat} best practice pattern**")
    cc1,cc2,cc3,cc4 = st.columns(4)
    cc1.metric("Top merchants", len(top))
    cc2.metric("Promo coverage", pct(top["_promo"].mean()))
    cc3.metric("Material coverage", pct(top["_material"].mean()))
    cc4.metric("Avg Exposure → Order", pct(weighted_avg(top,"_e2o","_exposure")))

with tabs[3]:
    st.markdown("### Opportunity Finder")
    opp = filtered.copy()
    # high exposure/low conversion + missing ops = opportunity
    opp["Opportunity Score"] = (
        opp["_exposure"].rank(pct=True).fillna(0)*35 +
        (1-opp["_e2o"].rank(pct=True).fillna(0))*25 +
        (1-opp["_v2c"].rank(pct=True).fillna(0))*15 +
        (~opp["_promo"]).astype(int)*10 +
        (~opp["_material"]).astype(int)*10 +
        (~opp["_visit_record"]).astype(int)*5
    ).round(1)
    opp = opp.sort_values("Opportunity Score", ascending=False).head(50)
    st.dataframe(opp[["_merchant","_bd","_area","_category","_level","_gmv","_orders","_exposure","_e2o","_v2c","_c2o","Opportunity Score"]].rename(columns={"_merchant":"Merchant","_bd":"BD","_area":"Area","_category":"Category","_level":"Level","_gmv":"GMV","_orders":"Orders","_exposure":"Exposure","_e2o":"Exposure → Order","_v2c":"Visit → Cart","_c2o":"Cart → Order"}).style.format({"GMV":"${:,.0f}", "Orders":"{:,.0f}", "Exposure":"{:,.0f}", "Exposure → Order":"{:.2%}", "Visit → Cart":"{:.2%}", "Cart → Order":"{:.2%}", "Opportunity Score":"{:.1f}"}), use_container_width=True)

with tabs[4]:
    st.markdown("### Metric Dictionary")
    md = pd.DataFrame([
        ["Exposure → Visit", "加权平均：Σ(曝光进店转化率 × 曝光人数) ÷ Σ曝光人数", "衡量店铺封面、标签、评分、配送等是否能把曝光转为进店。"],
        ["Visit → Cart", "加权平均：Σ(进店加购转化率 × 进店人数) ÷ Σ进店人数", "衡量菜单、图片、价格、套餐是否能促使用户加购。"],
        ["Cart → Order", "加权平均：Σ(加购下单转化率 × 加购人数) ÷ Σ加购人数", "衡量优惠、配送费、起送价、结算体验是否让用户完成下单。"],
        ["Exposure → Order", "加权平均：Σ(曝光下单转化率 × 曝光人数) ÷ Σ曝光人数", "整体漏斗效率，适合用于BD/商圈/等级对比。"],
        ["Learning Score", "同品类、同等级、同商圈、曝光规模接近、AOV接近、且表现更好", "用于找到最值得复制学习的店，而不是简单找GMV最高的店。"],
        ["Opportunity Score", "高曝光 + 低转化 + 缺活动/物料/拜访", "用于找最值得优先拜访和优化的店。"],
        ["Health Score", "GMV、转化漏斗、活动、物料、拜访综合评分", "用于识别优秀店和风险店。"],
    ], columns=["Metric", "Formula", "Why it matters"])
    st.dataframe(md, use_container_width=True)

st.caption("V6 note: conversion rates use source conversion fields with weighted aggregation. Orders are not divided by average cart people because the file is monthly snapshot data.")
