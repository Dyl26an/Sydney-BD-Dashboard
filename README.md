# Sydney Growth Intelligence V4

V4 rebuilds the metric logic using a clear Metric Dictionary.

Key changes:
- Funnel rates are weighted, calculated from raw counts: sum(numerator) / sum(denominator)
- No simple average of conversion-rate columns for BD ranking
- Orders, exposure, visit, cart columns are detected with safer priority rules
- Merchant name is displayed in action plans and merchant intelligence
- Metric explanations are shown in the app

Deploy on Streamlit with:
- Main file: app.py
