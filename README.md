# India Voice & Language AI APIs: Bottom-up Market Sizing

**Question:** how big is the Indian market for Indian-language voice and language AI APIs, which customer segments matter most, and where should a sales team spend its time?

Prices come from **Sarvam AI's public API price list** (docs.sarvam.ai, accessed 14 Nov 2026). Account counts come from public government and regulator data. Every assumption sits in a yellow cell in the workbook, so anyone can change it and watch the totals move.

## Files
| File | What it is |
|---|---|
| `market_sizing_model.xlsx` | The model. Tabs: **Inputs** (prices and call assumptions), **Segments** (bottom-up market), **Sensitivity**, **Funnel Plan**, **Sources** |
| `build_model.py` | Rebuilds the workbook from scratch (every number is a live formula) |
| `check_model.py` | Recomputes every headline number in plain Python and checks it against the workbook |
| `results/summary.json` | Headline outputs |

## Method
1. **Cost of one AI voice-call minute at Sarvam prices.** Speech to text (₹0.49/min) + text to speech (Bulbul v3, ₹3.00 per 1,000 characters × characters the agent speaks) + LLM (Sarvam 105B, ₹15/₹60 per 1M input/output tokens × tokens per turn × turns per minute).
2. **Segments × accounts.** Banks (RBI categories), insurers (IRDAI), RBI upper-layer NBFCs, states and UTs, private TV channels (I&B Ministry) and DPIIT-recognised startups.
3. **Annual spend per adopting account** = monthly volume × 12 × price per unit. Volume is call minutes for voice segments, dubbed minutes for media, and ₹ of API spend for startups.
4. **Full-adoption market → serviceable market → revenue potential** = accounts × spend → × 3-year adoption → × assumed vendor share.

## Findings (base case)
- **One AI voice minute costs about ₹1.72** at Sarvam's prices. **Text to speech is 65% of that cost**, so switching to Bulbul v4 Flash (₹1.00 per 1,000 characters) cuts it **44%, to ₹0.97**. TTS price is the biggest lever on what voice customers pay.
- **Serviceable market of about ₹63 Cr a year**, of which an assumed 20% vendor share is about ₹13 Cr.
- **232 enterprise accounts** (banks, insurers, NBFCs, government) make up **about half** of that market. A large bank adopting is worth **about 172×** a startup. Sales time belongs on BFSI and government accounts; startups belong in a self-serve funnel.
- **The startup long tail is the most uncertain number.** Moving startup adoption from 1% to 4% and large-bank minutes from 2.5 to 10 lakh a month moves the serviceable market from **₹41 Cr to ₹107 Cr**. That is the first assumption to test with real signup and usage data.

## Measured vs assumed
| Measured (cited in Sources tab) | Assumed (yellow cells) |
|---|---|
| All Sarvam unit prices and the ₹100 free credit | How much of each minute the agent speaks, speech rate, tokens per turn |
| Number of banks, insurers, upper-layer NBFCs, TV channels, startups, states | Monthly call minutes per adopting account |
| | 3-year adoption % by segment, vendor share |

## Funnel plan
The **Funnel Plan** tab sets out how to trace a self-serve user from signup to revenue: sign up → verify (₹100 free credits) → first API call → half the credits used → credits exhausted → first payment → spend growth. Each stage has the event to log, its conversion metric, the CRM (HubSpot) property to store, and the most likely leak to test. It is a measurement plan, not measured data.

## Limits
- Usage and adoption are assumptions; the model is built so they can be replaced with real data.
- Counts the largest NBFCs only (RBI upper layer), not the long tail, so the NBFC segment is conservative.
- Prices are list prices before GST and volume discounts; enterprise deals usually price lower.
- Voice and dubbing only: excludes document AI, translation and LLM-only use, so the market is understated.

## Run
```bash
python build_model.py     # rebuild the workbook
# open market_sizing_model.xlsx in Excel/Sheets/LibreOffice to calculate, then:
python check_model.py     # verifies every headline number
```
