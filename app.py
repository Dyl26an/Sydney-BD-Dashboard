import io
import re
from typing import Optional, Tuple

import pandas as pd
import streamlit as st

try:
    import msoffcrypto
except Exception:
    msoffcrypto = None

st.set_page_config(page_title="Sydney BD AI Copilot", layout="wide")

st.title("Sydney BD AI Copilot")
st.caption("Upload encrypted Excel → analyse merchant performance → generate action plan")


def decrypt_excel(uploaded_file, password: str) -> io.BytesIO:
    data = uploaded_file.read()
    bio = io.BytesIO(data)
    if msoffcrypto is None:
        raise RuntimeError("msoffcrypto-tool is not installed. Please check requirements.txt")
    office_file = msoffcrypto.OfficeFile(bio)
    office_file.load_key(password=password)
    out = io.BytesIO()
    office_file.decrypt(out)
    out.seek(0)
    return out


def read_excel(uploaded_file, password: str) -> pd.DataFrame:
    try:
        if password.strip():
            decrypted = decrypt_excel(uploaded_file, password.strip())
            return pd.read_excel(decrypted)
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file)
    except Exception as e:
        raise RuntimeError(f"Unable to open Excel. Please check password/file. Details: {e}")


def find_col(df: pd.DataFrame, patterns) -> Optional[str]:
    cols = list(df.columns)
    for pat in patterns:
        rgx = re.compile(pat, re.I)
        for c in cols:
            if rgx.search(str(c)):
                return c
    return None


def pick_columns(df: pd.DataFrame):
    return {
        "owner": find_col(df, [r"owner|bd|负责人|销售|account.*manager|am"]),
        "merchant": find_col(df, [r"merchant.*name|store.*name|shop.*name|店铺|商户|门店|name"]),
        "area": find_col(df, [r"area|zone|suburb|商圈|区域|district"]),
        "gmv": find_col(df, [r"gmv|sales|销售额|交易额"]),
        "orders": find_col(df, [r"orders?$|order.*count|订单"]),
        "exposure": find_col(df, [r"exposure|impression|曝光"]),
        "visit": find_col(df, [r"visit|进店|store.*view"]),
        "cart": find_col(df, [r"cart|加购"]),
        "material": find_col(df, [r"material|物料"]),
        "campaign": find_col(df, [r"campaign|promo|discount|voucher|活动|折扣|券|满减"]),
    }


def to_num(s):
    return pd.to_numeric(s, errors="coerce").fillna(0)


def analyse(df: pd.DataFrame, owner_name: str, cols: dict) -> Tuple[pd.DataFrame, pd.DataFrame]:
    work = df.copy()
    for k in ["gmv", "orders", "exposure", "visit", "cart"]:
        if cols.get(k):
            work[f"__{k}"] = to_num(work[cols[k]])
        else:
            work[f"__{k}"] = 0

    work["曝光到下单率"] = work.apply(lambda r: r["__orders"] / r["__exposure"] if r["__exposure"] else 0, axis=1)
    work["进店到加购率"] = work.apply(lambda r: r["__cart"] / r["__visit"] if r["__visit"] else 0, axis=1)
    work["加购到下单率"] = work.apply(lambda r: r["__orders"] / r["__cart"] if r["__cart"] else 0, axis=1)

    if owner_name.strip() and cols.get("owner"):
        my = work[work[cols["owner"]].astype(str).str.contains(owner_name.strip(), case=False, na=False)].copy()
    else:
        my = work.copy()

    def reason(row):
        reasons = []
        if row["__exposure"] > work["__exposure"].quantile(0.75) and row["曝光到下单率"] < work["曝光到下单率"].median():
            reasons.append("曝光高但下单转化低")
        if row["__visit"] > 0 and row["进店到加购率"] < work["进店到加购率"].median():
            reasons.append("进店后加购弱，菜单/图片/价格可能需要优化")
        if row["__cart"] > 0 and row["加购到下单率"] < work["加购到下单率"].median():
            reasons.append("加购后未下单，建议检查运费/满减/券")
        if cols.get("material") and str(row.get(cols["material"], "")).strip() in ["", "0", "无", "nan", "None"]:
            reasons.append("物料缺失")
        if cols.get("campaign") and str(row.get(cols["campaign"], "")).strip() in ["", "0", "无", "nan", "None"]:
            reasons.append("活动配置弱")
        return "；".join(reasons) if reasons else "稳定，可观察"

    my["机会原因"] = my.apply(reason, axis=1)
    my["机会分"] = (
        my["__exposure"].rank(pct=True) * 35
        + my["__gmv"].rank(pct=True) * 25
        + (1 - my["曝光到下单率"].rank(pct=True)) * 25
        + (1 - my["进店到加购率"].rank(pct=True)) * 15
    )

    out_cols = []
    for k in ["merchant", "area", "owner"]:
        if cols.get(k):
            out_cols.append(cols[k])
    out_cols += ["__gmv", "__orders", "__exposure", "__visit", "__cart", "曝光到下单率", "进店到加购率", "加购到下单率", "机会原因", "机会分"]
    action = my[out_cols].sort_values("机会分", ascending=False).head(30).copy()
    action = action.rename(columns={"__gmv":"GMV", "__orders":"订单", "__exposure":"曝光", "__visit":"进店", "__cart":"加购"})
    return my, action


with st.sidebar:
    st.header("Upload")
    uploaded = st.file_uploader("Encrypted Excel file", type=["xlsx", "xls"])
    password = st.text_input("Password", type="password")
    owner_name = st.text_input("BD / Owner name", value="Yuan Dong")
    analyse_btn = st.button("Analyse", type="primary")

if uploaded and analyse_btn:
    try:
        df = read_excel(uploaded, password)
        cols = pick_columns(df)
        my, action = analyse(df, owner_name, cols)

        st.success(f"Loaded {len(df):,} rows and {len(df.columns):,} columns.")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Sydney Merchants", f"{len(df):,}")
        c2.metric("Your Merchants", f"{len(my):,}")
        c3.metric("Your GMV", f"${my['__gmv'].sum():,.0f}")
        c4.metric("Your Orders", f"{my['__orders'].sum():,.0f}")

        st.subheader("Detected Columns")
        st.write(cols)

        st.subheader("Top Action Plan")
        st.dataframe(action, use_container_width=True)

        csv = action.to_csv(index=False).encode("utf-8-sig")
        st.download_button("Download Action Plan CSV", csv, "action_plan.csv", "text/csv")

        st.subheader("Weekly Summary Draft")
        st.text_area(
            "Copy this into your weekly report",
            value=(
                f"This week I analysed {len(df):,} Sydney merchants and {len(my):,} merchants under {owner_name}. "
                f"The priority focus is the Top 30 opportunity merchants with high exposure, weak conversion, missing material, or weak campaign setup. "
                "Next actions: visit high-exposure low-conversion merchants first, improve menu image/order structure, push voucher or bundle setup, and track conversion change in the next report."
            ),
            height=130,
        )
    except Exception as e:
        st.error(str(e))
else:
    st.info("Upload your monthly encrypted Excel, enter password, then click Analyse.")
    st.markdown("""
### What this V1 does
- Opens password-protected Excel files
- Detects likely columns automatically
- Filters by BD/owner name
- Finds high-opportunity merchants
- Generates a practical action plan

### Important
Do not upload company raw Excel files into GitHub. Upload only this app code.
""")
