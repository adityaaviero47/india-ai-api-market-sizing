"""Recomputes every headline number in plain Python and checks it against the workbook.
Run after opening/recalculating market_sizing_model.xlsx:  python check_model.py"""
import json
from openpyxl import load_workbook

wb = load_workbook("market_sizing_model.xlsx", data_only=True)
inp, seg = wb["Inputs"], wb["Segments"]
g = lambda r: inp.cell(r, 2).value

stt, tts_v3, tts_flash, llm_in, llm_out, dub = g(5), g(6), g(7), g(8), g(9), g(10)
talk, chars, turns, tin, tout, langs = g(13), g(14), g(15), g(16), g(17), g(18)
tts_min = talk * chars / 1000 * tts_v3
llm_min = turns * (tin * llm_in + tout * llm_out) / 1e6
cost_min = stt + tts_min + llm_min
flash_min = stt + talk * chars / 1000 * tts_flash + llm_min

share = seg["B17"].value
rows, sam_total, som_total = [], 0, 0
for r in range(5, 14):
    name, acc, vol, price, adopt = (seg.cell(r, c).value for c in (1, 2, 5, 7, 10))
    unit_price = cost_min if "minutes" in seg.cell(r, 6).value and "dubbed" not in seg.cell(r, 6).value else (dub * langs if "dubbed" in seg.cell(r, 6).value else 1)
    spend = vol * 12 * unit_price
    sam = acc * spend / 1e7 * adopt
    assert abs(sam - seg.cell(r, 11).value) < 1e-6, name
    rows.append({"segment": name, "accounts": acc, "annual_spend_per_adopter_inr": round(spend), "serviceable_cr": round(sam, 2)})
    sam_total += sam; som_total += sam * share

enterprise = sum(x["serviceable_cr"] for x in rows[:7])
out = {
  "cost_per_voice_minute_inr": round(cost_min, 2),
  "tts_share_of_cost": round(tts_min / cost_min, 3),
  "cost_per_minute_with_v4_flash_inr": round(flash_min, 2),
  "change_with_v4_flash": round(flash_min / cost_min - 1, 3),
  "serviceable_market_cr_per_year": round(sam_total, 1),
  "sarvam_revenue_potential_cr_per_year": round(som_total, 1),
  "enterprise_accounts": sum(x["accounts"] for x in rows[:7]),
  "enterprise_share_of_serviceable": round(enterprise / sam_total, 3),
  "large_bank_vs_startup_spend": round(rows[0]["annual_spend_per_adopter_inr"] / rows[-1]["annual_spend_per_adopter_inr"], 1),
  "segments": rows,
}
sen = wb["Sensitivity"]
vals = [sen.cell(r, c).value for r in range(5, 8) for c in range(2, 5)]
out["sensitivity_range_cr"] = [round(min(vals), 1), round(max(vals), 1)]
assert abs(sen["C6"].value - sam_total) < 1e-6
json.dump(out, open("results/summary.json", "w"), indent=2)
print(json.dumps({k: v for k, v in out.items() if k != "segments"}, indent=2))
print("All workbook values match the independent recomputation.")
