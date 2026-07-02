# Sydney Growth Intelligence V4.2

Fixes funnel metric logic.

Important correction:
- The report has monthly orders but average exposure/visit/cart user counts.
- Therefore `Orders / Cart` and `Orders / Exposure` are not valid.
- V4.2 uses official source conversion-rate fields and aggregates them using weighted averages.

Weighted methods:
- Exposure → Visit = Σ(曝光进店转化率 × 平均曝光人数) / Σ平均曝光人数
- Visit → Cart = Σ(进店加购转化率 × 平均进店人数) / Σ平均进店人数
- Cart → Order = Σ(加购下单转化率 × 平均加购人数) / Σ平均加购人数
- Exposure → Order = Σ(曝光下单转化率 × 平均曝光人数) / Σ平均曝光人数
