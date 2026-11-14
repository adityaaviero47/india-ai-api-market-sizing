"""Builds market_sizing_model.xlsx: a bottom-up sizing of the Indian enterprise market
for Indian-language voice and language AI APIs, priced on Sarvam AI's public price list.
Every count is sourced (Sources tab); every assumption is in a yellow cell (Inputs tab)."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

F = "Calibri"
BLUE = Font(name=F, color="0000FF"); BLACK = Font(name=F); GREEN = Font(name=F, color="008000")
BOLD = Font(name=F, bold=True); HEAD = Font(name=F, bold=True); TITLE = Font(name=F, bold=True, size=13)
YELLOW = PatternFill("solid", fgColor="FFFF00"); DARK = PatternFill("solid", fgColor="D9D9D9"); GREY = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
WRAP = Alignment(wrap_text=True, vertical="top")
RS = '#,##0;(#,##0);"-"'; RS2 = '#,##0.00;(#,##0.00);"-"'; CR = '#,##0.00;(#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'; NUM = '#,##0;(#,##0);"-"'

wb = Workbook()

# ---------------- Sources ----------------
src = wb.active; src.title = "Sources"
src_rows = [
 ("S1","Sarvam API prices (Pay as you go, excl. GST): STT ₹0.49/min; Bulbul v3 TTS ₹3.00 per 1K chars; Bulbul v4 Flash ₹1.00 per 1K chars; Sarvam 105B ₹15 in / ₹60 out per 1M tokens; Dubbing ₹40/min per target language; ₹100 free credits for every verified user","Sarvam AI docs, API pricing (primary)","https://docs.sarvam.ai/api/getting-started/pricing","14 Nov 2026"),
 ("S2","Banks in India: 12 public sector, 21 private sector, 11 small finance, 5 payments, 28 regional rural, 44 foreign","Reserve Bank of India, 'Banks in India' list (primary)","https://rbi.org.in/Commonman/English/Scripts/BanksInIndia.aspx","14 Nov 2026"),
 ("S3","Insurers as of 31 Mar 2025: 26 life, 25 general, 8 standalone health (33 non-life in total)","IRDAI Annual Report 2024-25, as summarised by Ditto Insurance","https://www.joinditto.in/articles/general/irdai-annual-report/","14 Nov 2026"),
 ("S4","NBFCs in the 2026-27 Upper Layer list: 17","Reserve Bank of India press release, 6 Aug 2026 (primary)","https://rbi.org.in/Scripts/BS_PressReleaseDisplay.aspx?prid=63301","14 Nov 2026"),
 ("S5","Private satellite TV channels: 908 in 2024-25","Ministry of I&B reply in Lok Sabha, reported by News On AIR (24 Jul 2025)","https://www.newsonair.gov.in/indias-media-landscape-expands-over-1-55-lakh-publications-908-private-tv-channels-ib-ministry","14 Nov 2026"),
 ("S6","DPIIT-recognised startups: 2,12,283 as on 31 Jan 2026","Press Information Bureau, Ministry of Commerce & Industry, 17 Mar 2026 (primary)","https://pib.gov.in/PressReleasePage.aspx?PRID=2241313","14 Nov 2026"),
 ("S7","States and Union Territories of India: 28 states + 8 UTs = 36","Government of India, Know India portal (primary)","https://knowindia.india.gov.in/states-uts/","14 Nov 2026"),
]
src["A1"] = "Sources"; src["A1"].font = TITLE
for j,h in enumerate(["ID","Fact used","Source","URL","Accessed"],1):
    c=src.cell(3,j,h); c.font=HEAD; c.fill=DARK
for i,r in enumerate(src_rows,4):
    for j,v in enumerate(r,1):
        c=src.cell(i,j,v); c.font=BLACK; c.alignment=WRAP; c.border=BOX
for col,w in zip("ABCDE",[6,70,42,60,12]): src.column_dimensions[col].width=w

# ---------------- Inputs ----------------
inp = wb.create_sheet("Inputs", 0)
inp["A1"] = "Inputs"; inp["A1"].font = TITLE
inp["A2"] = "blue = input, yellow = my assumption (change these), black = formula"; inp["A2"].font = Font(name=F, color="595959")
rows = [
 (4,"Sarvam prices (from their pricing page)",None,None,None,True),
 (5,"STT (Rs/min)",0.49,"S1",RS2,False),
 (6,"TTS Bulbul v3 (Rs/1K chars)",3.00,"S1",RS2,False),
 (7,"TTS Bulbul v4 Flash (Rs/1K chars)",1.00,"S1",RS2,False),
 (8,"Sarvam 105B input (Rs/1M tokens)",15,"S1",RS2,False),
 (9,"Sarvam 105B output (Rs/1M tokens)",60,"S1",RS2,False),
 (10,"Dubbing (Rs/min per language)",40,"S1",RS2,False),
 (12,"Per-minute call assumptions",None,None,None,True),
 (13,"Agent talk share",0.5,"roughly half the call",PCT,False),
 (14,"Chars spoken per min",750,"~150 wpm x 5 chars",NUM,False),
 (15,"LLM turns per min",4,"1 reply every ~15 sec",NUM,False),
 (16,"Input tokens per turn",1500,"guess, prompt + history",NUM,False),
 (17,"Output tokens per turn",60,"short spoken reply",NUM,False),
 (18,"Languages per dubbed min",2,"e.g. Hindi + 1 regional",NUM,False),
 (19,"TTS model (1 = v3, 2 = v4 Flash)",1,"switch to 2 to test v4 Flash",NUM,False),
]
for r,label,val,note,fmt,hdr in rows:
    a=inp.cell(r,1,label)
    if hdr: a.font=BOLD; a.fill=GREY; inp.cell(r,2).fill=GREY; inp.cell(r,3).fill=GREY; continue
    a.font=BLACK; b=inp.cell(r,2,val); b.font=BLUE; b.number_format=fmt; b.border=BOX
    if r>=13: b.fill=YELLOW
    inp.cell(r,3,note).font=Font(name=F, color="595959")
inp["A21"]="Cost per voice minute"; inp["A21"].font=BOLD; inp["A21"].fill=GREY; inp["B21"].fill=GREY; inp["C21"].fill=GREY
der = [
 (22,"STT","=B5"),
 (23,"TTS","=B13*B14/1000*IF(B19=2,B7,B6)"),
 (24,"LLM","=B15*(B16*B8+B17*B9)/1000000"),
 (25,"Total Rs/min","=SUM(B22:B24)"),
 (26,"TTS % of cost","=IF(B25=0,0,B23/B25)"),
 (27,"Rs/min with v4 Flash","=B22+B13*B14/1000*B7+B24"),
 (28,"Change vs base","=IF(B25=0,0,B27/B25-1)"),
 (29,"Dubbing Rs per source min","=B10*B18"),
]
for r,label,f in der:
    inp.cell(r,1,label).font=BLACK; c=inp.cell(r,2,f); c.font=BLACK; c.border=BOX
    c.number_format = PCT if r in (26,28) else RS2
inp["A25"].font=BOLD; inp["B25"].font=BOLD
inp.column_dimensions["A"].width=34; inp.column_dimensions["B"].width=12; inp.column_dimensions["C"].width=30

# ---------------- Segments ----------------
seg = wb.create_sheet("Segments", 1)
seg["A1"]="Bottom-up sizing (accounts x usage x price x adoption)"; seg["A1"].font=TITLE
seg["A2"]="Account counts from Sources tab. Yellow = assumptions. Cr = crore."; seg["A2"].font=Font(name=F, color="595959")
hdr=["Segment","Accounts","Source","Use case","Monthly volume / adopter","Unit","Rs / unit","Annual spend / adopter (Rs)","TAM (Rs Cr/yr)","3-yr adoption","SAM (Rs Cr/yr)","Sarvam share","SOM (Rs Cr/yr)","Rank (spend)","Lead product"]
for j,h in enumerate(hdr,1):
    c=seg.cell(4,j,h); c.font=HEAD; c.fill=DARK; c.alignment=Alignment(wrap_text=True,vertical="center")
seg.row_dimensions[4].height=48
V="=Inputs!$B$25"; D="=Inputs!$B$29"
data=[
 ("Banks: public and private",33,"S2: 12 + 21","Voice support and collections",500000,"call minutes",V,0.5),
 ("Banks: small finance, RRB, payments",44,"S2: 11 + 28 + 5","Voice support and collections",50000,"call minutes",V,0.3),
 ("Banks: foreign (India operations)",44,"S2","Voice support",10000,"call minutes",V,0.2),
 ("Insurers: life",26,"S3","Renewal and claims calls",100000,"call minutes",V,0.4),
 ("Insurers: general and health",33,"S3","Renewal and claims calls",100000,"call minutes",V,0.4),
 ("NBFCs: RBI upper layer",17,"S4","Collections and onboarding calls",200000,"call minutes",V,0.5),
 ("Government: states and UTs",36,"S7","Citizen helplines",200000,"call minutes",V,0.3),
 ("Media: private TV channels",908,"S5","Dubbing into Indian languages",300,"source minutes dubbed",D,0.2),
 ("Startups: DPIIT recognised",212283,"S6","Self-serve API use",5000,"Rs of API spend",1,0.02),
]
first=5
for i,(name,acc,s,use,vol,unit,price,adopt) in enumerate(data):
    r=first+i
    seg.cell(r,1,name).font=BLACK
    c=seg.cell(r,2,acc); c.font=BLUE; c.number_format=NUM
    seg.cell(r,3,s).font=Font(name=F, color="595959")
    seg.cell(r,4,use).font=BLACK
    c=seg.cell(r,5,vol); c.font=BLUE; c.fill=YELLOW; c.number_format=NUM
    seg.cell(r,6,unit).font=BLACK
    c=seg.cell(r,7,price); c.font=GREEN if isinstance(price,str) else BLUE; c.number_format=RS2
    c=seg.cell(r,8,f"=E{r}*12*G{r}"); c.number_format=RS
    c=seg.cell(r,9,f"=B{r}*H{r}/10000000"); c.number_format=CR
    c=seg.cell(r,10,adopt); c.font=BLUE; c.fill=YELLOW; c.number_format=PCT
    c=seg.cell(r,11,f"=I{r}*J{r}"); c.number_format=CR
    c=seg.cell(r,12,"=$B$17"); c.font=BLACK; c.number_format=PCT
    c=seg.cell(r,13,f"=K{r}*L{r}"); c.number_format=CR
    c=seg.cell(r,14,f"=RANK(H{r},$H${first}:$H${first+len(data)-1})"); c.number_format="0"
    prod={"Banks":"Samvaad voice agents","Insurers":"Samvaad voice agents","NBFCs":"Samvaad voice agents","Government":"Samvaad voice agents","Media":"Content Studio (dubbing)","Startups":"Model APIs (self-serve)"}
    seg.cell(r,15,prod[name.split(":")[0]]).font=BLACK
    for j in range(1,16): seg.cell(r,j).border=BOX
last=first+len(data)-1; tot=last+1
seg.cell(tot,1,"Total").font=BOLD
for col in "BIKM":
    c=seg[f"{col}{tot}"]; c.value=f"=SUM({col}{first}:{col}{last})"; c.font=BOLD
    c.number_format = NUM if col=="B" else CR
seg["A16"]="Share"; seg["A16"].font=BOLD
seg["A17"]="Sarvam share of serviceable (assumed)"; seg["B17"]=0.2; seg["B17"].font=BLUE; seg["B17"].fill=YELLOW; seg["B17"].number_format=PCT
seg["A19"]="Quick checks"; seg["A19"].font=BOLD
find=[
 ("Enterprise share (banks, insurers, NBFCs, govt)",f"=SUM(K{first}:K{first+6})/K{tot}",PCT),
 ("Startup share",f"=K{last}/K{tot}",PCT),
 ("Large bank vs startup spend",f"=H{first}/H{last}",'0"x"'),
 ("Enterprise accounts",f"=SUM(B{first}:B{first+6})",NUM),
]
for k,(lab,f,fmt) in enumerate(find,20):
    seg.cell(k,1,lab).font=BLACK; c=seg.cell(k,2,f); c.number_format=fmt; c.font=BOLD
widths=[36,11,16,30,16,18,11,16,16,11,16,10,16,10,26]
for j,w in enumerate(widths,1): seg.column_dimensions[get_column_letter(j)].width=w

# ---------------- Sensitivity ----------------
sen = wb.create_sheet("Sensitivity", 2)
sen["A1"]="Sensitivity: serviceable market (Rs Cr/yr)"; sen["A1"].font=TITLE
sen["A2"]="rows = startup adoption, cols = large-bank minutes/month; everything else at base case"; sen["A2"].font=Font(name=F, color="595959")
sen["A4"]="Startup adoption / bank min"; sen["A4"].font=BOLD
cols=[250000,500000,1000000]; rws=[0.01,0.02,0.04]
for j,v in enumerate(cols,2):
    c=sen.cell(4,j,v); c.font=BLUE; c.number_format=NUM
for i,v in enumerate(rws,5):
    c=sen.cell(i,1,v); c.font=BLUE; c.number_format=PCT
    for j in range(2,5):
        cl=get_column_letter(j)
        f=(f"=Segments!$K${tot}-Segments!$K${last}-Segments!$K${first}"
           f"+Segments!$K${last}*$A{i}/Segments!$J${last}"
           f"+Segments!$K${first}*{cl}$4/Segments!$E${first}")
        c=sen.cell(i,j,f); c.number_format=CR; c.border=BOX
sen["A9"]="Startup adoption swings the total the most, so test it first with real signup data."; sen["A9"].font=Font(name=F, color="595959")
sen.column_dimensions["A"].width=44
for c in "BCD": sen.column_dimensions[c].width=18

# ---------------- Funnel plan ----------------
fun = wb.create_sheet("Funnel Plan", 3)
fun["A1"]="Self-serve funnel: what to track"; fun["A1"].font=TITLE
fun["A2"]="Based on the Rs 100 free credits (S1). Plan only, no real data."; fun["A2"].font=Font(name=F, color="595959")
fh=["Stage","Event","Definition","Metric","HubSpot property","Possible leak"]
for j,h in enumerate(fh,1):
    c=fun.cell(4,j,h); c.font=HEAD; c.fill=DARK
steps=[
 ("1. Sign up","account_created","Account created on the dashboard","—","signup_date, signup_source","Traffic that never signs up"),
 ("2. Verify","account_verified","Verified, ₹100 free credits granted","Verified ÷ signed up","verified_date","Verification friction"),
 ("3. Activate","first_api_call","First successful API call (any product)","First call ÷ verified (7 days)","first_call_date, first_product","Unclear value, docs or key setup"),
 ("4. Engage","credits_50pct_used","₹50 of the free credits used","50% used ÷ first call (30 days)","credits_used, products_used","One test call, then nothing"),
 ("5. Exhaust","credits_exhausted","Free credits run out","Exhausted ÷ engaged","credits_exhausted_date","Quality or language gaps"),
 ("6. Pay","first_payment","First top-up or paid plan","Paid ÷ exhausted (14 days)","plan, first_payment_value","Price, or no plan that fits"),
 ("7. Expand","monthly_spend_up","Spend grows month on month","Net revenue retention","mrr, mrr_change, account_owner","Usage plateaus; route to sales"),
]
for i,s in enumerate(steps,5):
    for j,v in enumerate(s,1):
        c=fun.cell(i,j,v); c.font=BLACK; c.alignment=WRAP; c.border=BOX
fun["A13"]="Routing idea: if monthly spend or company size crosses a threshold, create the account in HubSpot with an owner so sales picks it up."; fun["A13"].font=Font(name=F, color="595959")
for col,w in zip("ABCDEF",[13,20,36,28,30,30]): fun.column_dimensions[col].width=w

for ws in wb.worksheets:
    ws.sheet_view.showGridLines = True
wb.save("market_sizing_model.xlsx")
print("saved")
