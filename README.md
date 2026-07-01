# Sydney Growth Intelligence V2.2

Fixes:
- Correctly detects real order column such as `订单数_排除mm的均单`
- Correctly detects exposure / visit / cart volume columns instead of conversion-rate columns
- Shows merchant names using `商户名称`
- Adds Merchant List tab with merchant name, area, category, GMV, orders, funnel metrics, promo/material/visit flags
- Keeps BD name display and comparison modules

Update on GitHub by replacing:
- app.py
- requirements.txt
- README.md

Then reboot the Streamlit app.
