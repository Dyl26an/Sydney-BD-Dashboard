# Sydney Growth Intelligence V4.1

Fixes V4 Compare Me formatting crash and keeps weighted metric definitions.

Core metric rules:
- Exposure → Visit = Σ Visit / Σ Exposure
- Visit → Cart = Σ Cart / Σ Visit
- Cart → Order = Σ Orders / Σ Cart
- Exposure → Order = Σ Orders / Σ Exposure

Do not use simple average of conversion-rate columns for BD or Sydney level summaries.
