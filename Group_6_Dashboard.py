"""
NYC Property Sales Dashboard - Phase 4
Approach: Pre-compute all filter combos in Python → write to hidden LookupData sheet
          → ChartSource sheet uses INDEX/MATCH formulas referencing Dashboard filter dropdowns
          → Charts read ChartSource → changing a dropdown recalculates formulas → chart updates
"""
import pandas as pd
import numpy as np
from scipy import stats
import xlsxwriter, warnings
warnings.filterwarnings("ignore")

# ── LOAD & CLEAN ─────────────────────────────────────────────────────────────
df = pd.read_csv("nyc-property-sales.csv", low_memory=False)
BMAP = {"1":"Manhattan","2":"Bronx","3":"Brooklyn","4":"Queens","5":"Staten Island"}
df["Borough"]    = df["BOROUGH"].astype(str).str.strip().map(BMAP)
df["Sale Price"] = pd.to_numeric(df["SALE PRICE"].astype(str).str.strip().str.replace(",",""), errors="coerce")
df["Sale Date"]  = pd.to_datetime(df["SALE DATE"], errors="coerce")
# Normalize building class (new dataset has inconsistent spacing)
df["BUILDING CLASS CATEGORY"] = df["BUILDING CLASS CATEGORY"].astype(str).str.strip().str.replace(r'\s+',' ',regex=True)

valid = df[(df["Sale Price"]>10000) & df["Sale Date"].notna() & df["Borough"].notna()].copy()
valid["Sale Year"] = valid["Sale Date"].dt.year.astype(int).astype(str)

BOROUGHS = ["Manhattan","Bronx","Brooklyn","Queens","Staten Island"]
top6     = valid["BUILDING CLASS CATEGORY"].value_counts().head(6).index.tolist()
CLASSES  = [c[:28] for c in top6]
valid["CLS"] = valid["BUILDING CLASS CATEGORY"].str[:28]
# MONTHS holds YEARS for Q2 trend analysis (2003-2023)
MONTHS   = [str(y) for y in sorted(valid["Sale Year"].dropna().astype(int).unique())]

# filter option lists (first option = "All")
FC = ["All"] + CLASSES   # 7 items → cols B-H in LookupData
FB = ["All"] + BOROUGHS  # 6 items → cols B-G in LookupData

# ── PRE-COMPUTE LOOKUP TABLES ─────────────────────────────────────────────────
def med(mask): v = valid.loc[mask,"Sale Price"]; return float(v.median()) if len(v) else 0
def cnt(mask): return int(mask.sum())
def tot(mask): v = valid.loc[mask,"Sale Price"]; return float(v.sum()) if len(v) else 0

# Q1: median price  — rows=boroughs(5), cols=FC(7)
q1 = {(b,c): med((valid.Borough==b) & (True if c=="All" else (valid.CLS==c))) for b in BOROUGHS for c in FC}

# Q2: transaction count — rows=years(21), cols=FB(6)
q2 = {(m,b): cnt((valid["Sale Year"]==m) & (True if b=="All" else (valid.Borough==b))) for m in MONTHS for b in FB}

# Q3: median price — rows=classes(6), cols=FB(6)
q3 = {(c,b): med((valid.CLS==c) & (True if b=="All" else (valid.Borough==b))) for c in CLASSES for b in FB}

# Q4: total value  — rows=boroughs(5), cols=FC(7)
q4 = {(b,c): tot((valid.Borough==b) & (True if c=="All" else (valid.CLS==c))) for b in BOROUGHS for c in FC}

# ── Q5: CORRELATION & LINEAR REGRESSION ──────────────────────────────────────
valid["Gross Sq Ft"] = pd.to_numeric(
    df.loc[valid.index, "GROSS SQUARE FEET"].astype(str).str.strip().str.replace(",",""), errors="coerce")
valid["Land Sq Ft"] = pd.to_numeric(
    df.loc[valid.index, "LAND SQUARE FEET"].astype(str).str.strip().str.replace(",",""), errors="coerce")
valid["Total Units"]  = pd.to_numeric(df.loc[valid.index, "TOTAL UNITS"],  errors="coerce")
valid["Year Built"]   = pd.to_numeric(df.loc[valid.index, "YEAR BUILT"],    errors="coerce")
valid["Res Units"]    = pd.to_numeric(df.loc[valid.index, "RESIDENTIAL UNITS"], errors="coerce")

q5_num_cols = ["Gross Sq Ft", "Land Sq Ft", "Total Units", "Year Built", "Res Units"]
q5_df = valid[["Sale Price"] + q5_num_cols].replace(0, np.nan).dropna()
# Cap outliers at 99th pct for cleaner analysis
for col in ["Sale Price"] + q5_num_cols:
    q5_df = q5_df[q5_df[col] <= q5_df[col].quantile(0.99)]

# Pearson correlation with Sale Price
corr_vals = {c: float(q5_df["Sale Price"].corr(q5_df[c])) for c in q5_num_cols}
corr_sorted = sorted(corr_vals.items(), key=lambda x: abs(x[1]), reverse=True)

# Linear regression: Sale Price ~ Gross Sq Ft (top predictor)
top_pred = corr_sorted[0][0]
slope, intercept, r_value, p_value, std_err = stats.linregress(
    q5_df[top_pred], q5_df["Sale Price"])
r_squared = r_value ** 2

# Scatter sample (500 pts for performance)
scatter = q5_df[[top_pred, "Sale Price"]].sample(min(500, len(q5_df)), random_state=42)
SCATTER_N = len(scatter)
scatter_x = scatter[top_pred].tolist()
scatter_y = scatter["Sale Price"].tolist()
print(f"[Q5] Top predictor: {top_pred}  |  R²={r_squared:.3f}  |  p={p_value:.2e}")

# ── WORKBOOK SETUP ────────────────────────────────────────────────────────────
wb  = xlsxwriter.Workbook("Group_6_NYC_Dashboard.xlsx")

# Formats
fT  = wb.add_format({"bold":True,"font_size":20,"font_color":"#FFFFFF","bg_color":"#1F3864","align":"center","valign":"vcenter"})
fS  = wb.add_format({"italic":True,"font_size":9,"font_color":"#FFFFFF","bg_color":"#2E75B6","align":"center","valign":"vcenter"})
fH  = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#2E75B6","align":"center","border":1})
fHD = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#1F3864","align":"center","border":1})
fLB = wb.add_format({"bold":True,"font_color":"#1F3864","font_size":10})
fFL = wb.add_format({"bold":True,"font_color":"#1F3864","bg_color":"#D9E1F2","border":1,"align":"center","font_size":9})
fFV = wb.add_format({"border":2,"align":"center","font_color":"#1F3864","bold":True,"bg_color":"#FFFFFF","font_size":10})
fKL = wb.add_format({"bold":True,"bg_color":"#1F3864","font_color":"#FFFFFF","align":"center","border":1,"font_size":9})
fKV = wb.add_format({"bg_color":"#2E75B6","font_color":"#FFFFFF","align":"center","border":1,"font_size":9,"num_format":"$#,##0"})
fKN = wb.add_format({"bg_color":"#2E75B6","font_color":"#FFFFFF","align":"center","border":1,"font_size":9,"num_format":"#,##0"})
fKT = wb.add_format({"bg_color":"#2E75B6","font_color":"#FFFFFF","align":"center","border":1,"font_size":9})
fCH = wb.add_format({"bold":True,"font_size":11,"font_color":"#1F3864","bg_color":"#EBF3FB","align":"center","valign":"vcenter","border":1})
fNote= wb.add_format({"italic":True,"font_color":"#888888","font_size":8})

# ── SHEET 1: LookupData (hidden) ──────────────────────────────────────────────
LD = wb.add_worksheet("LookupData")
LD.hide()

# Row offsets (0-indexed)
Q1H=0;  Q1D=1   # Q1 header row, data start
Q2H=8;  Q2D=9   # Q2
Q3H=Q2D+len(MONTHS)+1; Q3D=Q3H+1
Q4H=Q3D+6+1;    Q4D=Q4H+1

def write_lookup(sheet, hdr_row, dat_row, row_labels, col_labels, table):
    sheet.write(hdr_row, 0, "Label", fHD)
    for ci,cl in enumerate(col_labels):
        sheet.write(hdr_row, ci+1, cl, fH)
    for ri,rl in enumerate(row_labels):
        sheet.write(dat_row+ri, 0, rl)
        for ci,cl in enumerate(col_labels):
            sheet.write(dat_row+ri, ci+1, table[(rl,cl)])

write_lookup(LD, Q1H, Q1D, BOROUGHS, FC, q1)
write_lookup(LD, Q2H, Q2D, MONTHS,   FB, q2)
write_lookup(LD, Q3H, Q3D, CLASSES,  FB, q3)
write_lookup(LD, Q4H, Q4D, BOROUGHS, FC, q4)

# Excel row numbers (1-indexed) for formula strings
def er(r): return r+1  # 0-indexed → Excel 1-indexed

# ── SHEET 2: ChartSource (hidden) — INDEX/MATCH formulas ─────────────────────
# Filter cells on Dashboard sheet (0-indexed → Excel):
# Q1 Building Class filter: row=7,col=5 → F8
# Q2 Borough filter:        row=7,col=17 → R8
# Q3 Borough filter:        row=31,col=5 → F32
# Q4 Building Class filter: row=31,col=17 → R32

CS = wb.add_worksheet("ChartSource")
CS.hide()

# Helper: Excel col letter from 0-indexed col number
def col_letter(c):
    s=""
    c+=1
    while c>0:
        c,r=divmod(c-1,26)
        s=chr(65+r)+s
    return s

# Q1 data — col A (0), col B (1)
# INDEX range: LookupData B{Q1D+1}:H{Q1D+5}  (5 boroughs × 7 classes)
# MATCH range: LookupData B{Q1H+1}:H{Q1H+1}  (class headers)
q1_idx = f"LookupData!$B${er(Q1D)}:$H${er(Q1D+4)}"
q1_mch = f"LookupData!$B${er(Q1H)}:$H${er(Q1H)}"
CS.write(0,0,"Borough",fH); CS.write(0,1,"Median Price (USD)",fH)
for i,b in enumerate(BOROUGHS):
    CS.write(i+1, 0, b)
    CS.write_formula(i+1, 1,
        f'=INDEX({q1_idx},{i+1},MATCH(Dashboard!$F$8,{q1_mch},0))')

# Q2 data — col D (3), col E (4)
# INDEX range: LookupData B{Q2D+1}:G{Q2D+nM}  (months × 6 boroughs)
# MATCH range: LookupData B{Q2H+1}:G{Q2H+1}
nM = len(MONTHS)
q2_idx = f"LookupData!$B${er(Q2D)}:$G${er(Q2D+nM-1)}"
q2_mch = f"LookupData!$B${er(Q2H)}:$G${er(Q2H)}"
CS.write(0,3,"Month",fH); CS.write(0,4,"Transactions",fH)
for i,m in enumerate(MONTHS):
    CS.write(i+1, 3, m)
    CS.write_formula(i+1, 4,
        f'=INDEX({q2_idx},{i+1},MATCH(Dashboard!$R$8,{q2_mch},0))')

# Q3 data — col G (6), col H (7)
# INDEX range: LookupData B{Q3D+1}:G{Q3D+5}  (6 classes × 6 boroughs)
# MATCH range: LookupData B{Q3H+1}:G{Q3H+1}
q3_idx = f"LookupData!$B${er(Q3D)}:$G${er(Q3D+5)}"
q3_mch = f"LookupData!$B${er(Q3H)}:$G${er(Q3H)}"
CS.write(0,6,"Building Class",fH); CS.write(0,7,"Avg Price (USD)",fH)
for i,c in enumerate(CLASSES):
    CS.write(i+1, 6, c)
    CS.write_formula(i+1, 7,
        f'=INDEX({q3_idx},{i+1},MATCH(Dashboard!$F$32,{q3_mch},0))')

# Q4 data — col J (9), col K (10)
# INDEX range: LookupData B{Q4D+1}:H{Q4D+4}  (5 boroughs × 7 classes)
# MATCH range: LookupData B{Q4H+1}:H{Q4H+1}
q4_idx = f"LookupData!$B${er(Q4D)}:$H${er(Q4D+4)}"
q4_mch = f"LookupData!$B${er(Q4H)}:$H${er(Q4H)}"
CS.write(0,9,"Borough",fH); CS.write(0,10,"Total Sales Value (USD)",fH)
for i,b in enumerate(BOROUGHS):
    CS.write(i+1, 9, b)
    CS.write_formula(i+1,10,
        f'=INDEX({q4_idx},{i+1},MATCH(Dashboard!$R$32,{q4_mch},0))')

# Q5 scatter data — col M (12)=x, col N (13)=y
CS.write(0,12,top_pred,fH); CS.write(0,13,"Sale Price",fH)
for i,(x,y) in enumerate(zip(scatter_x, scatter_y)):
    CS.write(i+1,12,x)
    CS.write(i+1,13,y)

# ── SHEET 3: DASHBOARD ────────────────────────────────────────────────────────
ws = wb.add_worksheet("Dashboard")
ws.hide_gridlines(2)
ws.set_zoom(85)
ws.set_column("A:A",2); ws.set_column("B:B",14); ws.set_column("C:Z",11)
ws.set_row(0,6); ws.set_row(1,52); ws.set_row(2,28); ws.set_row(3,10)

# Title (row 1) and summary (row 2) — Excel rows 2-3
ws.merge_range("B2:V2","NYC Property Sales Dashboard  |  2003 – 2023",fT)
ws.merge_range("B3:V3",
    f"Analyzing {len(valid):,} transactions across 5 NYC Boroughs  "
    "·  Q1: Median Price by Borough  "
    "·  Q2: Yearly Volume & Price Trends (2003–2023)  "
    "·  Q3: Avg Price by Building Class  "
    "·  Q4: Total Market Value by Borough  "
    "·  Use the dropdown filters on each chart to explore the data",fS)

ws.set_row(3,8)

# ── Top-left area: Chart 1 (Q1) ───────────────────────────────────────────────
ws.set_row(4,20); ws.set_row(5,20); ws.set_row(6,20); ws.set_row(7,26)
ws.merge_range("B5:K5","Q1 — Median Sale Price by Borough",fCH)
ws.write("B8","Building Class:",fFL)
ws.write("F8","All",fFV)   # Q1 filter cell (col=5,row=7)
ws.data_validation("F8",{
    "validate":"list","source":FC,
    "input_title":"Q1 Filter","input_message":"Select a building class (or All)"
})
ws.write("G8","<- change to filter chart",fNote)

# ── Top-right area: Chart 2 (Q2) ─────────────────────────────────────────────
ws.merge_range("M5:V5","Q2 — Yearly Transaction Volume (2003–2023)",fCH)
ws.write("Q8","Borough:",fFL)
ws.write("R8","All",fFV)   # Q2 filter cell (col=17,row=7)
ws.data_validation("R8",{
    "validate":"list","source":FB,
    "input_title":"Q2 Filter","input_message":"Select a borough (or All)"
})
ws.write("S8","<- change to filter chart",fNote)

# ── Bottom-left area: Chart 3 (Q3) ───────────────────────────────────────────
ws.set_row(29,20); ws.set_row(30,20); ws.set_row(31,26)
ws.merge_range("B29:K29","Q3 — Avg Sale Price by Building Class",fCH)
ws.write("E32","Borough:",fFL)
ws.write("F32","All",fFV)   # Q3 filter cell (col=5,row=31)
ws.data_validation("F32",{
    "validate":"list","source":FB,
    "input_title":"Q3 Filter","input_message":"Select a borough (or All)"
})
ws.write("G32","<- change to filter chart",fNote)

# ── Bottom-right area: Chart 4 (Q4) ──────────────────────────────────────────
ws.merge_range("M29:V29","Q4 — Total Sales Value by Borough",fCH)
ws.write("Q32","Building Class:",fFL)
ws.write("R32","All",fFV)   # Q4 filter cell (col=17,row=31)
ws.data_validation("R32",{
    "validate":"list","source":FC,
    "input_title":"Q4 Filter","input_message":"Select a building class (or All)"
})
ws.write("S32","<- change to filter chart",fNote)

# ── BUILD CHARTS ──────────────────────────────────────────────────────────────
CHART_W, CHART_H = 480, 300

# Chart 1 — Horizontal Bar (Q1: median price by borough)
c1 = wb.add_chart({"type":"bar"})
c1.add_series({
    "name":"Median Sale Price",
    "categories":["ChartSource",1,0,5,0],
    "values":    ["ChartSource",1,1,5,1],
    "fill":{"color":"#2E75B6"},"gap":90
})
c1.set_title({"name":"Median Sale Price by Borough"})
c1.set_x_axis({"name":"Median Price (USD)","num_format":"$#,##0"})
c1.set_y_axis({"name":"Borough"})
c1.set_legend({"none":True})
c1.set_size({"width":CHART_W,"height":CHART_H})
c1.set_chartarea({"fill":{"color":"#F5F9FF"},"border":{"color":"#BDD7EE"}})
ws.insert_chart("B9",c1,{"x_offset":4,"y_offset":4})

# Chart 2 — Line (Q2: monthly transactions)
c2 = wb.add_chart({"type":"line"})
c2.add_series({
    "name":"Transactions",
    "categories":["ChartSource",1,3,nM,3],
    "values":    ["ChartSource",1,4,nM,4],
    "line":{"color":"#ED7D31","width":2.5},
    "marker":{"type":"circle","size":5,"fill":{"color":"#ED7D31"},"border":{"color":"#C55A11"}}
})
c2.set_title({"name":"Yearly Transaction Volume (2003-2023)"})
c2.set_x_axis({"name":"Year","text_axis":True})
c2.set_y_axis({"name":"Number of Sales"})
c2.set_legend({"none":True})
c2.set_size({"width":CHART_W,"height":CHART_H})
c2.set_chartarea({"fill":{"color":"#FFFBF5"},"border":{"color":"#F4B183"}})
ws.insert_chart("M9",c2,{"x_offset":4,"y_offset":4})

# Chart 3 — Horizontal Bar (Q3: avg price by building class)
c3 = wb.add_chart({"type":"bar"})
c3.add_series({
    "name":"Avg Sale Price",
    "categories":["ChartSource",1,6,6,6],
    "values":    ["ChartSource",1,7,6,7],
    "fill":{"color":"#70AD47"},"gap":90
})
c3.set_title({"name":"Avg Sale Price by Building Class"})
c3.set_x_axis({"name":"Avg Price (USD)","num_format":"$#,##0"})
c3.set_y_axis({"name":"Building Class"})
c3.set_legend({"none":True})
c3.set_size({"width":CHART_W,"height":CHART_H})
c3.set_chartarea({"fill":{"color":"#F6FFF0"},"border":{"color":"#A9D18E"}})
ws.insert_chart("B33",c3,{"x_offset":4,"y_offset":4})

# Chart 4 — Column (Q4: total sales value by borough)
c4 = wb.add_chart({"type":"column"})
c4.add_series({
    "name":"Total Sales Value",
    "categories":["ChartSource",1,9,5,9],
    "values":    ["ChartSource",1,10,5,10],
    "fill":{"color":"#7030A0"},"gap":80
})
c4.set_title({"name":"Total Sales Value by Borough"})
c4.set_x_axis({"name":"Borough"})
c4.set_y_axis({"name":"Total Value (USD)","num_format":'$#,##0,,"M"'})
c4.set_legend({"none":True})
c4.set_size({"width":CHART_W,"height":CHART_H})
c4.set_chartarea({"fill":{"color":"#FAF5FF"},"border":{"color":"#B39CD0"}})
ws.insert_chart("M33",c4,{"x_offset":4,"y_offset":4})

# ── KPI FOOTER ────────────────────────────────────────────────────────────────
total_txn    = len(valid)
median_price = float(valid["Sale Price"].median())
total_val    = float(valid["Sale Price"].sum())
top_borough  = valid.groupby("Borough")["Sale Price"].sum().idxmax()

ws.set_row(54,8); ws.set_row(55,28)
ws.write("B56","Total Valid Sales",fKL); ws.write("C56",total_txn,fKN)
ws.write("D56","Median Sale Price",fKL); ws.write("E56",median_price,fKV)
ws.write("F56","Total Market Value",fKL); ws.write("G56",total_val,fKV)
ws.write("H56","Top Borough (by value)",fKL); ws.write("I56",top_borough,fKT)


# ── Q5 SECTION ON DASHBOARD ──────────────────────────────────────────────────
fQ5T  = wb.add_format({"bold":True,"font_size":14,"font_color":"#FFFFFF",
                       "bg_color":"#375623","align":"center","valign":"vcenter"})
fQ5S  = wb.add_format({"italic":True,"font_size":9,"font_color":"#FFFFFF",
                       "bg_color":"#548235","align":"center"})
fQ5H  = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#548235",
                       "align":"center","border":1})
fQ5R  = wb.add_format({"bg_color":"#E2EFDA","border":1,"align":"center"})
fQ5RL = wb.add_format({"bg_color":"#E2EFDA","border":1,"bold":True})
fQ5RN = wb.add_format({"bg_color":"#E2EFDA","border":1,"align":"center","num_format":"0.000"})
fQ5V  = wb.add_format({"bg_color":"#C6EFCE","border":1,"align":"center",
                       "bold":True,"num_format":"0.000"})
fStat = wb.add_format({"bg_color":"#DEEAF1","border":1,"font_size":9,"align":"center"})
fStatL= wb.add_format({"bg_color":"#BDD7EE","border":1,"font_size":9,"bold":True,"align":"center"})
fStatV= wb.add_format({"bg_color":"#DEEAF1","border":1,"font_size":9,"align":"center","num_format":"0.000"})

# Q5 banner rows 58-60 (0-indexed 57-59)
ws.set_row(57,8); ws.set_row(58,40); ws.set_row(59,22); ws.set_row(60,8)
ws.merge_range("B59:V59","Q5 — Correlation & Linear Regression: Which Numeric Attributes Drive Sale Price?",fQ5T)
ws.merge_range("B60:V60",
    f"Dataset: NYC Property Sales 2003-2023  |  Top predictor: {top_pred}  |  R\u00b2 = {r_squared:.3f}  |  "
    f"Equation: Sale Price = {slope:,.0f} x {top_pred} + {intercept:,.0f}  "
    f"|  p-value = {p_value:.2e}  |  Sample n = {SCATTER_N}",fQ5S)

# Left side: scatter chart — insert at row 61 (0-idx=60), col B
c5 = wb.add_chart({"type":"scatter"})
c5.add_series({
    "name":       "Sampled Sales",
    "categories": ["ChartSource",1,12,SCATTER_N,12],
    "values":     ["ChartSource",1,13,SCATTER_N,13],
    "marker":{"type":"circle","size":4,
              "fill":{"color":"#4472C4","transparency":50},
              "border":{"color":"#4472C4"}},
    "line":{"none":True},
    "trendline":{
        "type":"linear",
        "display_equation":True,
        "display_r_squared":True,
        "line":{"color":"#FF0000","width":2}
    }
})
c5.set_title({"name":f"Q5 — Sale Price vs {top_pred}"})
c5.set_x_axis({"name":top_pred,"num_format":"#,##0"})
c5.set_y_axis({"name":"Sale Price (USD)","num_format":"$#,##0"})
c5.set_legend({"none":True})
c5.set_size({"width":580,"height":340})
c5.set_chartarea({"fill":{"color":"#F6FFF0"},"border":{"color":"#A9D18E"}})
ws.insert_chart("B61",c5,{"x_offset":4,"y_offset":4})

# Right side: Correlation table + regression stats — col N onward, row 61
ws.set_row(60,20); ws.set_row(61,20)
ws.write("N61","Predictor",fQ5H)
ws.write("O61","Pearson r",fQ5H)
ws.write("P61","Strength",fQ5H)
for i,(pred,r) in enumerate(corr_sorted):
    strength = "Strong" if abs(r)>0.5 else ("Moderate" if abs(r)>0.3 else "Weak")
    base_fmt = fQ5V if i==0 else fQ5R
    ws.write(62+i, 13, pred,     fQ5RL)
    ws.write(62+i, 14, r,        fQ5RN)
    ws.write(62+i, 15, strength, base_fmt)

# Regression stats box
reg_row = 62 + len(corr_sorted) + 1
ws.merge_range(reg_row, 13, reg_row, 15, "Linear Regression Summary", fQ5H)
stats_data = [
    ("Dependent Variable", "Sale Price",       fStat),
    ("Independent Variable", top_pred,          fStat),
    ("R (correlation)",   round(r_value,4),     fStatV),
    ("R² (explained var.)",round(r_squared,4),  fStatV),
    ("Slope (coefficient)",round(slope,2),      fStatV),
    ("Intercept",          round(intercept,2),  fStatV),
    ("p-value",            round(p_value,6),    fStatV),
    ("Sample Size",        SCATTER_N,           fStat),
]
for j,(lbl,val,fmt) in enumerate(stats_data):
    ws.write(reg_row+1+j, 13, lbl, fStatL)
    ws.merge_range(reg_row+1+j, 14, reg_row+1+j, 15, val, fmt)

# ── DATA REFERENCE SHEET ──────────────────────────────────────────────────────
dr = wb.add_worksheet("Data Reference")
dr.set_column("A:B",22); dr.set_column("C:F",18)
dr.merge_range("A1:B1","Q1 Median Sale Price by Borough (All Classes)",fHD)
dr.write("A2","Borough",fH); dr.write("B2","Median Price",fH)
for i,b in enumerate(BOROUGHS):
    dr.write(i+2,0,b); dr.write(i+2,1,q1[(b,"All")])

dr.merge_range("D1:E1","Q4 Total Sales Value by Borough (All Classes)",fHD)
dr.write("D2","Borough",fH); dr.write("E2","Total Value",fH)
for i,b in enumerate(BOROUGHS):
    dr.write(i+2,3,b); dr.write(i+2,4,q4[(b,"All")])

dr.merge_range("A9:C9","Q2 Monthly Transactions (All Boroughs)",fHD)
dr.write("A10","Month",fH); dr.write("B10","Transactions",fH)
for i,m in enumerate(MONTHS):
    dr.write(i+10,0,m); dr.write(i+10,1,q2[(m,"All")])

# Q5 detail in Data Reference
dr.merge_range("D9:F9","Q5 Correlation with Sale Price",fHD)
dr.write("D10","Predictor",fH); dr.write("E10","Pearson r",fH); dr.write("F10","R squared",fH)
for i,(pred,r) in enumerate(corr_sorted):
    dr.write(10+i,3,pred); dr.write(10+i,4,round(r,4)); dr.write(10+i,5,round(r**2,4))

wb.close()
print("[OK] Group_6_NYC_Dashboard.xlsx created")
print(f"     Valid records : {total_txn:,}")
print(f"     Median price  : ${median_price:,.0f}")
print(f"     Total value   : ${total_val/1e9:.2f}B")
print()
print("HOW FILTERS WORK:")
print("  - Chart 1 (Q1): change cell F8  (Building Class dropdown)")
print("  - Chart 2 (Q2): change cell R8  (Borough dropdown)")
print("  - Chart 3 (Q3): change cell F32 (Borough dropdown)")
print("  - Chart 4 (Q4): change cell R32 (Building Class dropdown)")
print("  Charts update automatically via INDEX/MATCH formulas.")
