# Project Variance Analysis: Who Really Owns the Raw Material Cost Overrun?

Used **Excel** to build a **3-factor raw material variance analysis (Yield – Mix – Price)** for a 330ml canned beer SKU at **Sài Thành Brewery**. It settles a June cost-review dispute between the CEO and the Purchasing department with data.

<img width="982" height="187" alt="image" src="https://github.com/user-attachments/assets/acc262b3-dd58-4221-9073-9da4bdb6b65b" />
<img width="785" height="307" alt="image" src="https://github.com/user-attachments/assets/dc5b628f-84f5-4b8d-8e3b-cdb6408df09d" />

## Business Context
At the June cost review, management accounting reported that the **actual brewing material cost per can was above standard**. With **5,000,000 cans/month**, the gap was eroding the plant's gross margin.

- **The CEO** looked only at the total variance and blamed Purchasing: *"Imported malt and hop prices went up. Purchasing must take full responsibility!"*
- **The Purchasing Director** pushed back. His team had negotiated hop prices **below** standard, and he suspected most of the gap came from operations: more material per can, and a recipe mix that drifted from the formula. But he had no data to prove it.

**Task:** Break down the variance to answer: *Who is really responsible, and for how many VND?*

## Dataset
Standard vs actual brewing materials **per 330ml can** (sheet `Case study`):

| RM Code | Ingredient | Std Qty (g) | Std Price (VND/g) | Actual Qty (g) | Actual Price (VND/g) |
|:---|:---|---:|---:|---:|---:|
| B001 | Barley malt | 42 | 18.00 | 46 | 18.30 |
| R001 | Rice grits | 18 | 12.00 | 21 | 12.40 |
| H001 | Hops pellets | 0.40 | 900.00 | 0.45 | 890.00 |
| W001 | Brewing water | 340 | 0.80 | 350 | 0.85 |
| | **Total** | **400.40** | | **417.45** | |

**Scope:** only the 4 main brewing ingredients. Cans, lids, cartons, yeast, labour, utilities, depreciation, overhead, excise tax and VAT are excluded. The figures are the **direct material cost of the beer liquid**, not the full cost of a can.

## Preview
| | Standard | Actual | Variance |
|:---|---:|---:|---:|
| Material cost per can (VND) | 1,604.0 | 1,800.2 | **(196.2)** |
| Impact on monthly cost (VND) | | | **(981,000,000)** |

→ **The CEO is wrong.** Purchase price accounts for only **17.9%** of the overrun. **82.1%** comes from operations: the quantity of material used and the recipe mix.

## How to implement

### Step 1: Pull data into the analysis sheet
All inputs are linked from `Case study` with `XLOOKUP` on the RM Code, so the analysis updates automatically when the raw data changes:
```excel
=XLOOKUP($C11,'Case study'!$B$20:$B$23,'Case study'!$D$20:$D$23)
```

### Step 2: Calculate mix % and amounts
For both Standard and Actual:
- **Mix %** = ingredient quantity ÷ total quantity → `=F11/$F$15`
- **Amount** = quantity × unit price → `=F11*H11`

<img width="432" height="227" alt="image" src="https://github.com/user-attachments/assets/e211649d-be1c-4710-8a07-9c7b214b4a43" />

<img width="417" height="235" alt="image" src="https://github.com/user-attachments/assets/28b747d9-b473-4ff1-ab1d-2bafdfff28d2" />


### Step 3: Split the variance into 3 factors
**Sign convention:** Variance = Standard − Actual (impact on profit). Negative = **Adverse**, positive = **Favorable**.

| Factor | Question it answers | Formula (per ingredient) |
|:---|:---|:---|
| **① ΔYield** | Using the standard mix and price, how much extra cost comes from using **417.45g instead of 400.40g** of material? | `Std Mix × Std Price × (Std Total Qty − Actual Total Qty)` |
| **② ΔMix** | At standard prices, how much does blending the 417.45g at the **actual ratio instead of the standard ratio** cost? | `Std Price × Actual Total Qty × (Std Mix − Actual Mix)` |
| **ΔUsage** | Total quantity effect | `ΔYield + ΔMix` |
| **③ ΔPrice** | How much does the **purchase price** change cost for the actual quantity used? | `Actual Qty × (Std Price − Actual Price)` |
| **ΔTotal** | Total variance | `ΔUsage + ΔPrice` |

A **check cell** confirms that `ΔTotal = Standard Amount − Actual Amount` (returns `True`).

<img width="1222" height="232" alt="image" src="https://github.com/user-attachments/assets/91bd5ead-5773-49e0-8d36-ae3bbad1896e" />

### Step 4: Results by ingredient (VND per can)
| Ingredient | ΔYield | ΔMix | ΔUsage | ΔPrice | ΔTotal |
|:---|---:|---:|---:|---:|---:|
| Barley malt | (32.2) | (39.8) | (72.0) | (13.8) | (85.8) |
| Rice grits | (9.2) | (26.8) | (36.0) | (8.4) | (44.4) |
| Hops pellets | (15.3) | (29.7) | (45.0) | **+4.5** | (40.5) |
| Brewing water | (11.6) | +3.6 | (8.0) | (17.5) | (25.5) |
| **Total** | **(68.3)** | **(92.7)** | **(161.0)** | **(35.2)** | **(196.2)** |

### Step 5: Visualise with a waterfall chart
A **waterfall chart** walks from the Standard cost (1,604.0) through ΔYield (+68.3), ΔMix (+92.7) and ΔPrice (+35.2) to the Actual cost (1,800.2).
<img width="626" height="366" alt="image" src="https://github.com/user-attachments/assets/5380b39c-a363-4927-ab36-3fa854b4c863" />

## Key Findings – Who is responsible?
| Factor | VND / can | Share | Favorable / Adverse | Responsible department |
|:---|---:|---:|:---|:---|
| ① ΔYield | 68.3 | 34.8% | Adverse | **Production** |
| ② ΔMix | 92.7 | 47.2% | Adverse | **Quality Control / R&D** |
| ③ ΔPrice | 35.2 | 17.9% | Adverse | **Purchasing** |
| **Total** | **196.2** | **100%** | **Adverse** | |

1. **Purchasing is responsible for only 17.9%** (35.2 VND/can, about 176M VND/month), not the entire overrun. Hops pellets even show a **favorable** price variance of +4.5 VND, which confirms the Purchasing Director's claim that he negotiated below-standard hop prices.
2. **The mix variance is the biggest driver (47.2%).** The blend drifted from the recipe. Barley malt had the largest impact (−39.8 VND).
3. **The yield variance (34.8%)** comes from loading 417.45g of material per can instead of 400.40g, which points to losses in brewing, filtering and filling.
4. Altogether, operations (ΔUsage) account for **161 VND/can (82.1%)**, about **805M VND/month**.

<img width="787" height="312" alt="image" src="https://github.com/user-attachments/assets/1ae8a5fb-8bec-4320-8020-d9e108f11bcb" />

## Recommendations
- **Production:** review losses in brewing, filtering and filling. Set a loss KPI of ≤ 400.4 g/can.
- **QC / R&D:** tighten weighing tolerances for hops and malt, and re-check the blend ratio for every batch.
- **Purchasing:** keep the hop price-locking strategy, and extend long-term contracts to malt and rice grits.

> 💡 **Key takeaway:** A single total variance hides who is accountable. Splitting it into **Yield, Mix and Price** turns a blame argument into clear, measurable actions for each department.

## Assumptions & Limitations
- Mix % is calculated by **weight, including brewing water** (~85% of total weight). This dilutes the mix effect of the high-value dry ingredients.
- The brewing water price includes allocated treatment chemicals and energy. Part of its −17.5 VND price variance may belong to Utilities/Production rather than Purchasing.
- The monthly impact assumes every one of the 5,000,000 cans carries the same per-can variance.

# Author
Huynh Truong Giang
Email: *truonggiang281005@gmail.com*
