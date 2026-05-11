"""
Variance Commentary Generator — v6.0
CHANGES IN v6:
  - Transaction Trail: Sub-Category column now shows SubAccount (e.g. Cash at Bank, Salaries)
  - Var% formula: Y0=0 now shows +100.0% instead of N/A
  - All IF(x=0,N/A,...) formulas replaced with IF(x=0,1,...) for clean formatting
Finance × AI: In Practice — alloutoftokens.com

HOW TO USE:
  1. Put your two data files in this folder:
     - period_0.xlsx  (prior year / Y0)
     - period_1.xlsx  (current year / Y1)
  2. Optionally edit  context.txt  to add management context
  3. Run:  python analyse_v5.py
  4. Open: variance_analysis.xlsx  and  variance_dashboard.html

BUGS FIXED IN v5:
  - Commentary parser rewritten — exec/revenue sections were always empty
  - Tab 04 Transaction Drivers column now visible (was width=2 spacer)
  - Bridge tab chart data rows now hidden
  - All column widths verified against actual content
  - BS flux notes now cover all accounts
  - HTML JS brace balance verified across all 14 functions
"""

import os
import sys
import json
import pandas as pd
import numpy as np
import anthropic
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from datetime import datetime

# ── FILE NAMES ────────────────────────────────────────────────────────────────
Y0_FILE      = "period_0.xlsx"
Y1_FILE      = "period_1.xlsx"
CONTEXT_FILE = "context.txt"
XLS_OUTPUT   = "variance_analysis.xlsx"
HTML_OUTPUT  = "variance_dashboard.html"

# ── PERIOD LABELS ─────────────────────────────────────────────────────────────
Y0_LABEL = "FY 2020"
Y1_LABEL = "FY 2021"

# ── API KEY ───────────────────────────────────────────────────────────────────
# Set as environment variable ANTHROPIC_API_KEY — do NOT hardcode here
API_KEY = os.environ.get("ANTHROPIC_API_KEY", "paste_your_key_here")

# ── MATERIALITY THRESHOLDS ────────────────────────────────────────────────────
# Both abs AND pct must be met for an account to be flagged as material
THRESHOLDS = {
    "Sales":                       {"pct": 3.0,  "abs": 10000},
    "Cost of Sales":               {"pct": 3.0,  "abs": 10000},
    "Sales & Distribution":        {"pct": 5.0,  "abs": 5000},
    "Marketing":                   {"pct": 5.0,  "abs": 5000},
    "Administration":              {"pct": 5.0,  "abs": 5000},
    "Depreciation & Amortization": {"pct": 2.0,  "abs": 500},
    "Interest & Tax":              {"pct": 5.0,  "abs": 5000},
    "Non-operating":               {"pct": 5.0,  "abs": 1000},
    "Balance Sheet":               {"pct": 10.0, "abs": 50000},
    "DEFAULT":                     {"pct": 5.0,  "abs": 5000},
}

# One-off capital accounts excluded from P&L executive table
ONEOFF_KEYS = {75, 80, 90, 100, 150, 160, 170, 180, 190}

# ── DESIGN TOKENS ─────────────────────────────────────────────────────────────
BLACK   = "0D0D0D"; CREAM  = "F5F0E8"; RED    = "FF3C1F"
DARK    = "1A1714"; MID    = "6B6560"; LIGHT  = "EDE8E0"; WHITE = "FFFFFF"
FAV_BG  = "E8F5E9"; ADV_BG = "FFF0ED"
FAV_TXT = "1B5E20"; ADV_TXT= "B71C1C"
AMBER   = "FF8F00"; BLUE   = "0000FF"

def _f(c):  return PatternFill("solid", start_color=c, end_color=c)
def _bot(c="CCCCCC"):  s=Side(style="thin",color=c); return Border(bottom=s)
def _all(c="CCCCCC"):  s=Side(style="thin",color=c); return Border(left=s,right=s,top=s,bottom=s)
def _align(h="left", v="center", wrap=False, indent=0):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, indent=indent)
def _font(bold=False, size=10, color=BLACK, italic=False):
    return Font(name="Arial", bold=bold, size=size, color=color, italic=italic)
def _ff(size=10):  return Font(name="Arial", size=size, color="000000")   # formula = black
def _fi(size=10):  return Font(name="Arial", size=size, color=BLUE)       # input   = blue

def set_widths(ws, d):
    for c, w in d.items(): ws.column_dimensions[c].width = w

def hdr(ws, r, c, text, end=None, bg=BLACK, tc=WHITE, size=10):
    cell = ws.cell(r, c, text)
    cell.font = _font(bold=True, size=size, color=tc)
    cell.fill = _f(bg); cell.alignment = _align(indent=1)
    ws.row_dimensions[r].height = 22
    if end: ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=end)
    return r + 1

def col_hdrs(ws, r, start, headers, right=()):
    for i, h in enumerate(headers):
        c = ws.cell(r, start+i, h)
        c.font = _font(bold=True, size=9, color=WHITE); c.fill = _f(DARK)
        c.alignment = _align("right" if i in right else "left", indent=1)
    ws.row_dimensions[r].height = 18
    return r + 1

def fa_write(ws, r, col, var_val):
    fav = var_val > 0
    c = ws.cell(r, col, "✓ Fav" if fav else "✗ Adv")
    c.font = _font(bold=True, size=10, color=FAV_TXT if fav else ADV_TXT)
    c.fill = _f(FAV_BG if fav else ADV_BG); c.alignment = _align("center")

# ── FINANCIAL LOGIC ───────────────────────────────────────────────────────────

def is_fav(var_val):
    """
    Universal rule for this GL:
    All costs stored negative. Var = Y1 - Y0.
    Var > 0  always means better profit contribution:
      Revenue accounts: Var > 0 = more revenue = GOOD
      Cost accounts:    Var > 0 = less negative = lower cost = GOOD
    """
    return float(var_val) > 0

def is_material(row):
    cat = str(row.get("SubClass2") or "DEFAULT")
    thr = THRESHOLDS.get(cat, THRESHOLDS["DEFAULT"])
    if row.get("Report") == "Balance Sheet":
        thr = THRESHOLDS["Balance Sheet"]
    pct = abs(row["VarPct"]) if pd.notna(row["VarPct"]) else 0
    return abs(row["Var"]) >= thr["abs"] and pct >= thr["pct"]

def row_bg(var_val, is_mat):
    if not is_mat: return WHITE
    return FAV_BG if is_fav(var_val) else ADV_BG

def row_tc(var_val, is_mat):
    if not is_mat: return DARK
    return FAV_TXT if is_fav(var_val) else ADV_TXT

# ── DATA LAYER ────────────────────────────────────────────────────────────────

def load_data():
    src  = pd.read_excel(Y0_FILE, sheet_name=None)
    gl   = src["GL"].copy(); gl["Date"] = pd.to_datetime(gl["Date"])
    y0   = gl[gl["Date"].dt.year == 2020].copy()
    y1   = pd.read_excel(Y1_FILE, sheet_name="GL").copy()
    coa  = src["Chart of Accounts"]
    terr = src.get("Territory", pd.DataFrame(columns=["Territory_key","Country","Region"]))
    # SubAccount map: Account_key -> SubAccount name (for Transaction Trail sub-category)
    sa_map = coa.set_index("Account_key")["SubAccount"].to_dict() if "SubAccount" in coa.columns else {}
    return y0, y1, coa, terr, sa_map

def build_summary(y0, y1, coa):
    y0s = y0.groupby("Account_key")["Amount"].sum()
    y1s = y1.groupby("Account_key")["Amount"].sum()
    m   = pd.DataFrame({"Y0": y0s, "Y1": y1s}).fillna(0)
    m   = m.join(coa.set_index("Account_key")[
        ["Report","Class","SubClass","SubClass2","Account","SubAccount"]])
    m["Var"]    = m["Y1"] - m["Y0"]
    m["VarPct"] = m.apply(
        lambda r: (r["Var"]/abs(r["Y0"])*100) if r["Y0"] != 0 else None, axis=1)
    return m

def get_drivers(y0, y1, acc_key, n=5):
    a = y0[y0["Account_key"]==acc_key].groupby("Details")["Amount"].sum()
    b = y1[y1["Account_key"]==acc_key].groupby("Details")["Amount"].sum()
    d = pd.DataFrame({"Y0":a,"Y1":b}).fillna(0); d["Var"]=d["Y1"]-d["Y0"]
    return d.reindex(d["Var"].abs().sort_values(ascending=False).index).head(n)

def monthly_series(gl_df, acc_key):
    m = gl_df[gl_df["Account_key"]==acc_key].copy()
    m["Month"] = pd.to_datetime(m["Date"]).dt.month
    return m.groupby("Month")["Amount"].sum().reindex(range(1,13),fill_value=0).tolist()

def territory_rev(gl_df, acc_key):
    return gl_df[gl_df["Account_key"]==acc_key].groupby("Territory_key")["Amount"].sum().to_dict()

def load_user_context():
    """Load optional management context from context.txt"""
    if os.path.exists(CONTEXT_FILE):
        with open(CONTEXT_FILE, encoding="utf-8") as f:
            txt = f.read().strip()
        # Return context if non-empty and not the default template placeholder
        if txt and "add your management context here" not in txt.lower():
            return txt
    return ""

# ── AI COMMENTARY ─────────────────────────────────────────────────────────────

def parse_commentary(raw):
    """
    Parse Claude's structured response into sections.
    Fixed: saves section content WHEN the next header is encountered,
    AND saves the last section at end of loop.
    """
    secs = {"exec": "", "revenue": "", "opex": "", "mgmt": "", "_raw": raw}
    cur, buf = None, []

    for line in raw.splitlines():
        s = line.strip()
        # Detect section headers
        if s.startswith("##"):
            # Save whatever was in the previous section
            if cur:
                secs[cur] = "\n".join(buf).strip()
            # Start new section
            buf = []
            su = s.upper()
            if "EXECUTIVE SUMMARY" in su:
                cur = "exec"
            elif "REVENUE" in su and ("GROSS" in su or "MARGIN" in su):
                cur = "revenue"
            elif "OPERATING EXPENSES" in su or "OPEX" in su:
                cur = "opex"
            elif "MANAGEMENT INPUT" in su or "AREAS REQUIRING" in su:
                cur = "mgmt"
            else:
                cur = None  # unknown header — skip
        else:
            if cur:
                buf.append(line)

    # Save the last section
    if cur and buf:
        secs[cur] = "\n".join(buf).strip()

    return secs

def generate_commentary(summary, y0, y1, user_context=""):
    pl     = summary[summary["Report"]=="Profit and Loss"]
    ns_y0  = pl[pl["SubClass2"]=="Sales"]["Y0"].sum()
    ns_y1  = pl[pl["SubClass2"]=="Sales"]["Y1"].sum()
    cos_y0 = pl[pl["SubClass2"]=="Cost of Sales"]["Y0"].sum()
    cos_y1 = pl[pl["SubClass2"]=="Cost of Sales"]["Y1"].sum()
    gp_y0, gp_y1 = ns_y0+cos_y0, ns_y1+cos_y1
    op_y0  = pl[pl["Class"]=="Operating account"]["Y0"].sum()
    op_y1  = pl[pl["Class"]=="Operating account"]["Y1"].sum()
    eb_y0, eb_y1 = gp_y0+op_y0, gp_y1+op_y1

    mat_pl = pl[(~pl.index.isin(ONEOFF_KEYS)) & pl.apply(is_material,axis=1)].copy()
    mat_pl = mat_pl.reindex(mat_pl["Var"].abs().sort_values(ascending=False).index)

    lines = []
    for _, r in mat_pl.iterrows():
        pct   = f"{r['VarPct']:+.1f}%" if r["VarPct"] is not None else "N/A"
        label = "FAVOURABLE" if is_fav(r["Var"]) else "ADVERSE"
        drv   = get_drivers(y0, y1, r.name, 3)
        dparts= [f"{i}: Y0={v['Y0']:,.0f} Y1={v['Y1']:,.0f} Δ={v['Var']:+,.0f}"
                 for i,v in drv.iterrows()] if not drv.empty else []
        lines.append(
            f"- [{label}] {r['Account']} ({r['SubClass2']}): "
            f"Y0={r['Y0']:,.0f} Y1={r['Y1']:,.0f} Δ={r['Var']:+,.0f} ({pct}) "
            f"| Drivers: {'; '.join(dparts) or 'see GL'}")

    context_block = f"\nMANAGEMENT CONTEXT PROVIDED BY USER:\n{user_context}\n" if user_context else ""

    prompt = f"""You are a senior finance analyst writing management commentary for a CFO board pack.
Period: {Y0_LABEL} vs {Y1_LABEL}.
ALL cost accounts are stored as NEGATIVE numbers in this GL.
A FAVOURABLE variance on a cost account means Variance > 0 (cost decreased).
An ADVERSE variance on a cost account means Variance < 0 (cost increased).
{context_block}
HEADLINE P&L:
  Net Revenue:  {Y0_LABEL}={ns_y0:,.0f}  {Y1_LABEL}={ns_y1:,.0f}  Δ={ns_y1-ns_y0:+,.0f} ({(ns_y1-ns_y0)/abs(ns_y0)*100:+.1f}%)
  Gross Profit: {Y0_LABEL}={gp_y0:,.0f}  {Y1_LABEL}={gp_y1:,.0f}  GP Margin Y0={gp_y0/ns_y0*100:.1f}% Y1={gp_y1/ns_y1*100:.1f}%
  EBITDA:       {Y0_LABEL}={eb_y0:,.0f}  {Y1_LABEL}={eb_y1:,.0f}  Δ={eb_y1-eb_y0:+,.0f} ({(eb_y1-eb_y0)/abs(eb_y0)*100:+.1f}%)

MATERIAL VARIANCES:
{chr(10).join(lines)}

Write commentary using EXACTLY these four section headers (no extra ## headers):

## EXECUTIVE SUMMARY
One paragraph, max 75 words. State net revenue growth, GP margin direction, EBITDA movement. End with one forward-looking sentence.

## REVENUE AND GROSS MARGIN
2-3 paragraphs. Decompose sales movement using driver data. Explain cost of sales relative to revenue. Where the cause cannot be determined from GL data, write "requires management input."

## OPERATING EXPENSES
2 paragraphs. Group by category. State what changed, by how much, and what the driver data suggests.

## MANAGEMENT INPUT REQUIRED
Numbered list of exactly 5 items. Each references a specific account and variance amount.
Format: N. Account Name (Δ amount): specific question

RULES: No markdown bold or italic. Plain text only. Every claim traceable to the data above."""

    client  = anthropic.Anthropic(api_key=API_KEY)
    msg     = client.messages.create(
        model="claude-opus-4-6", max_tokens=1500,
        messages=[{"role":"user","content":prompt}])
    raw = msg.content[0].text
    return parse_commentary(raw)

# ── EXCEL BUILDERS ────────────────────────────────────────────────────────────

def build_exec_tab(wb, summary, commentary, y0, y1):
    ws = wb.active; ws.title = "01 Executive Summary"
    ws.sheet_view.showGridLines = False
    set_widths(ws, {"A":2,"B":28,"C":16,"D":16,"E":16,"F":14,"G":14,"H":14,"I":2})

    r=1
    ws.merge_cells(f"B{r}:H{r}")
    c=ws.cell(r,2,f"VARIANCE ANALYSIS  ·  {Y0_LABEL} vs {Y1_LABEL}")
    c.font=_font(bold=True,size=16,color=WHITE); c.fill=_f(BLACK)
    c.alignment=_align(indent=1); ws.row_dimensions[r].height=36

    r+=1; ws.merge_cells(f"B{r}:H{r}")
    c=ws.cell(r,2,f"Generated {datetime.now().strftime('%d %B %Y')}  ·  alloutoftokens.com  ·  Finance × AI: In Practice  ·  v6.0")
    c.font=_font(size=9,color=RED); c.fill=_f(CREAM); c.alignment=_align(indent=1); ws.row_dimensions[r].height=16

    r+=1; ws.merge_cells(f"B{r}:H{r}")
    c=ws.cell(r,2,"Colour convention:  GREEN = Favourable (revenue up OR cost down vs prior year)  |  RED = Adverse (revenue down OR cost up)  |  Blue text = input  |  Black text = formula")
    c.font=_font(size=9,color=MID,italic=True); c.fill=_f(LIGHT); c.alignment=_align(indent=1); ws.row_dimensions[r].height=16
    r+=2

    # KPI tiles
    pl     = summary[summary["Report"]=="Profit and Loss"]
    ns_y0  = pl[pl["SubClass2"]=="Sales"]["Y0"].sum()
    ns_y1  = pl[pl["SubClass2"]=="Sales"]["Y1"].sum()
    cos_y0 = pl[pl["SubClass2"]=="Cost of Sales"]["Y0"].sum()
    cos_y1 = pl[pl["SubClass2"]=="Cost of Sales"]["Y1"].sum()
    gp_y0,gp_y1 = ns_y0+cos_y0, ns_y1+cos_y1
    op_y0  = pl[pl["Class"]=="Operating account"]["Y0"].sum()
    op_y1  = pl[pl["Class"]=="Operating account"]["Y1"].sum()
    eb_y0,eb_y1 = gp_y0+op_y0, gp_y1+op_y1
    gm_y0  = gp_y0/ns_y0*100 if ns_y0 else 0
    gm_y1  = gp_y1/ns_y1*100 if ns_y1 else 0

    kpis=[("NET REVENUE",ns_y0,ns_y1,False,"B"),
          ("GROSS PROFIT",gp_y0,gp_y1,False,"D"),
          ("EBITDA",eb_y0,eb_y1,False,"F"),
          ("GP MARGIN",gm_y0,gm_y1,True,"H")]
    tile_r=r
    for lbl,v0,v1,is_pct,cl in kpis:
        col=ord(cl)-64; var=v1-v0; fav=is_fav(var)
        bg=FAV_BG if fav else ADV_BG; tc=FAV_TXT if fav else ADV_TXT
        vals=[(lbl,_font(bold=True,size=8,color=MID),CREAM),
              (f"{v1:.1f}%" if is_pct else f"{v1:,.0f}",_font(bold=True,size=13,color=DARK),CREAM),
              (f"{'▲' if fav else '▼'} {abs(var):.1f}pp" if is_pct else f"{'▲' if fav else '▼'} {abs(var):,.0f}  ({var/abs(v0)*100:+.1f}%)",
               _font(bold=True,size=9,color=tc),bg),
              (f"Prior: {v0:.1f}%" if is_pct else f"Prior: {v0:,.0f}",_font(size=8,color=MID,italic=True),CREAM)]
        for i,(txt,fnt,fbg) in enumerate(vals):
            cr=tile_r+i
            c=ws.cell(cr,col,txt); c.font=fnt; c.fill=_f(fbg); c.alignment=_align(indent=1)
            ws.row_dimensions[cr].height=18
            ws.merge_cells(start_row=cr,start_column=col,end_row=cr,end_column=col+1)
    r=tile_r+5

    # Executive narrative
    r=hdr(ws,r,2,"EXECUTIVE SUMMARY",8)
    exec_text = commentary.get("exec","")
    if not exec_text:
        exec_text = "[Commentary not generated — check API key and re-run]"
    ws.merge_cells(f"B{r}:H{r}")
    c=ws.cell(r,2,exec_text)
    c.font=_font(size=11,color=DARK); c.fill=_f(CREAM)
    c.alignment=_align(wrap=True,v="top",indent=1)
    ws.row_dimensions[r].height=max(80,len(exec_text)//4)
    r+=2

    # Top P&L variances
    r=hdr(ws,r,2,"TOP P&L VARIANCES  (ranked by absolute impact · one-off capital items excluded)",8)
    r=col_hdrs(ws,r,2,["Account","Category",Y0_LABEL,Y1_LABEL,"Variance","Var %","F/A"],right=(2,3,4,5))

    pl_mat=summary[(summary["Report"]=="Profit and Loss") &
                   (~summary.index.isin(ONEOFF_KEYS)) &
                   (summary.apply(is_material,axis=1))].copy()
    pl_mat=pl_mat.reindex(pl_mat["Var"].abs().sort_values(ascending=False).index)

    for _,row in pl_mat.head(12).iterrows():
        var=row["Var"]; fav=is_fav(var)
        bg=FAV_BG if fav else ADV_BG; tc=FAV_TXT if fav else ADV_TXT
        ws.cell(r,2,row["Account"]).font=_font(size=10,bold=True,color=tc)
        ws.cell(r,2).fill=_f(bg); ws.cell(r,2).alignment=_align(indent=1); ws.cell(r,2).border=_bot()
        ws.cell(r,3,row["SubClass2"]).font=_font(size=10,color=MID)
        ws.cell(r,3).fill=_f(bg); ws.cell(r,3).alignment=_align(indent=1); ws.cell(r,3).border=_bot()
        for col,val,num in [(4,row["Y0"],'#,##0;(#,##0);"-"'),(5,row["Y1"],'#,##0;(#,##0);"-"')]:
            c=ws.cell(r,col,val); c.font=_fi(10); c.fill=_f(bg)
            c.alignment=_align("right"); c.number_format=num; c.border=_bot()
        # Variance formula =E-D
        c=ws.cell(r,6,f"=E{r}-D{r}"); c.font=_ff(10); c.fill=_f(bg)
        c.alignment=_align("right"); c.number_format='#,##0;(#,##0);"-"'; c.border=_bot()
        # Var% formula
        c=ws.cell(r,7,f'=IF(D{r}=0,1,F{r}/ABS(D{r}))'); c.font=_ff(10); c.fill=_f(bg)
        c.alignment=_align("right"); c.number_format='+0.0%;-0.0%;"-"'; c.border=_bot()
        fa_write(ws,r,8,var); ws.cell(r,8).fill=_f(bg); ws.cell(r,8).border=_bot()
        ws.row_dimensions[r].height=18; r+=1

    r+=1
    r=hdr(ws,r,2,"AREAS REQUIRING MANAGEMENT INPUT",8,bg=RED)
    mgmt_text=commentary.get("mgmt","")
    if not mgmt_text:
        mgmt_text="[Commentary not generated — check API key and re-run]"
    ws.merge_cells(f"B{r}:H{r}")
    c=ws.cell(r,2,mgmt_text); c.font=_font(size=10,color=DARK); c.fill=_f(CREAM)
    c.alignment=_align(wrap=True,v="top",indent=1)
    ws.row_dimensions[r].height=max(120,len(mgmt_text)//3)


def build_pl_tab(wb, summary, commentary, y0, y1):
    ws=wb.create_sheet("02 P&L Commentary"); ws.sheet_view.showGridLines=False
    set_widths(ws,{"A":2,"B":32,"C":16,"D":16,"E":14,"F":12,"G":8,"H":44,"I":2})

    r=1; ws.merge_cells(f"B{r}:H{r}")
    c=ws.cell(r,2,f"PROFIT & LOSS  ·  {Y0_LABEL} vs {Y1_LABEL}")
    c.font=_font(bold=True,size=14,color=WHITE); c.fill=_f(BLACK)
    c.alignment=_align(indent=1); ws.row_dimensions[r].height=30

    r+=1; ws.merge_cells(f"B{r}:H{r}")
    c=ws.cell(r,2,"GREEN = cost decreased or revenue increased (Fav)  |  RED = cost increased or revenue decreased (Adv)  |  Blue = input  |  Black = formula")
    c.font=_font(size=9,color=MID,italic=True); c.fill=_f(LIGHT); c.alignment=_align(indent=1); ws.row_dimensions[r].height=16
    r+=2

    r=col_hdrs(ws,r,2,["Account",Y0_LABEL,Y1_LABEL,"Variance","Var %","F/A","Key Drivers"],right=(1,2,3,4))

    class_order={"Trading account":0,"Operating account":1,"Non-operating":2,"Interest & Tax":3}
    pl=summary[summary["Report"]=="Profit and Loss"].copy()
    pl["_ord"]=pl["Class"].map(class_order).fillna(99)
    pl=pl.sort_values(["_ord","SubClass2","Account"])

    cur_class=None; cur_sub=None
    for _,row in pl.iterrows():
        if row["Class"]!=cur_class:
            cur_class=row["Class"]; cur_sub=None
            ws.merge_cells(f"B{r}:H{r}")
            c=ws.cell(r,2,cur_class.upper()); c.font=_font(bold=True,size=10,color=WHITE)
            c.fill=_f(RED); c.alignment=_align(indent=1); ws.row_dimensions[r].height=20; r+=1
        if row["SubClass2"]!=cur_sub:
            cur_sub=row["SubClass2"]
            ws.merge_cells(f"B{r}:H{r}")
            c=ws.cell(r,2,f"  {cur_sub}"); c.font=_font(bold=True,size=9,color=DARK)
            c.fill=_f(LIGHT); c.alignment=_align(indent=2); ws.row_dimensions[r].height=16; r+=1

        var=row["Var"]; mat=is_material(row); fav=is_fav(var)
        bg=row_bg(var,mat); tc=row_tc(var,mat)

        ws.cell(r,2,f"  {row['Account']}").font=_font(size=10,bold=mat,color=tc)
        ws.cell(r,2).fill=_f(bg); ws.cell(r,2).alignment=_align(indent=2); ws.cell(r,2).border=_bot("E0E0E0")
        for col,val in [(3,row["Y0"]),(4,row["Y1"])]:
            c=ws.cell(r,col,val); c.font=_fi(10); c.fill=_f(bg)
            c.alignment=_align("right"); c.number_format='#,##0;(#,##0);"-"'; c.border=_bot("E0E0E0")
        c=ws.cell(r,5,f"=D{r}-C{r}"); c.font=_ff(10); c.fill=_f(bg)
        c.alignment=_align("right"); c.number_format='#,##0;(#,##0);"-"'; c.border=_bot("E0E0E0")
        c=ws.cell(r,6,f'=IF(C{r}=0,1,E{r}/ABS(C{r}))'); c.font=_ff(10); c.fill=_f(bg)
        c.alignment=_align("right"); c.number_format='+0.0%;-0.0%;"-"'; c.border=_bot("E0E0E0")
        fa_write(ws,r,7,var) if mat else ws.cell(r,7,"")
        ws.cell(r,7).fill=_f(bg); ws.cell(r,7).border=_bot("E0E0E0")
        drv_txt=""
        if mat:
            d=get_drivers(y0,y1,row.name,2)
            if not d.empty:
                drv_txt=" | ".join([f"{i}: Δ{v['Var']:+,.0f}" for i,v in d.iterrows()])+" — verify with mgmt"
        c=ws.cell(r,8,drv_txt); c.font=_font(size=9,color=MID,italic=bool(drv_txt))
        c.fill=_f(bg); c.alignment=_align(wrap=True,indent=1); c.border=_bot("E0E0E0")
        ws.row_dimensions[r].height=28 if drv_txt else 18; r+=1

    # Subtotals
    r+=1
    pl_full=summary[summary["Report"]=="Profit and Loss"]
    ns_y0=pl_full[pl_full["SubClass2"]=="Sales"]["Y0"].sum()
    ns_y1=pl_full[pl_full["SubClass2"]=="Sales"]["Y1"].sum()
    cos_y0=pl_full[pl_full["SubClass2"]=="Cost of Sales"]["Y0"].sum()
    cos_y1=pl_full[pl_full["SubClass2"]=="Cost of Sales"]["Y1"].sum()
    gp_y0,gp_y1=ns_y0+cos_y0,ns_y1+cos_y1
    op_y0=pl_full[pl_full["Class"]=="Operating account"]["Y0"].sum()
    op_y1=pl_full[pl_full["Class"]=="Operating account"]["Y1"].sum()
    eb_y0,eb_y1=gp_y0+op_y0,gp_y1+op_y1
    r=hdr(ws,r,2,"KEY P&L SUBTOTALS",8)
    r=col_hdrs(ws,r,2,["Metric",Y0_LABEL,Y1_LABEL,"Variance","Var %","F/A",""],right=(1,2,3,4))
    for lbl,v0,v1 in [("Net Revenue",ns_y0,ns_y1),("Gross Profit",gp_y0,gp_y1),("EBITDA (Operating Profit)",eb_y0,eb_y1)]:
        var=v1-v0; fav=is_fav(var); bg=FAV_BG if fav else ADV_BG
        ws.cell(r,2,lbl).font=_font(bold=True,size=10,color=DARK)
        ws.cell(r,2).fill=_f(bg); ws.cell(r,2).alignment=_align(indent=1)
        for col,val in [(3,v0),(4,v1)]:
            c=ws.cell(r,col,val); c.font=_fi(10); c.fill=_f(bg)
            c.alignment=_align("right"); c.number_format='#,##0;(#,##0);"-"'
        c=ws.cell(r,5,f"=D{r}-C{r}"); c.font=_ff(10); c.fill=_f(bg)
        c.alignment=_align("right"); c.number_format='#,##0;(#,##0);"-"'
        c=ws.cell(r,6,f'=IF(C{r}=0,1,E{r}/ABS(C{r}))'); c.font=_ff(10); c.fill=_f(bg)
        c.alignment=_align("right"); c.number_format='+0.0%;-0.0%;"-"'
        fa_write(ws,r,7,var); ws.cell(r,7).fill=_f(bg)
        ws.row_dimensions[r].height=20; r+=1

    r+=2
    r=hdr(ws,r,2,"REVENUE & GROSS MARGIN COMMENTARY",8)
    rev_text=commentary.get("revenue","")
    if not rev_text: rev_text="[Commentary not generated — check API key and re-run]"
    ws.merge_cells(f"B{r}:H{r}")
    c=ws.cell(r,2,rev_text); c.font=_font(size=11,color=DARK); c.fill=_f(CREAM)
    c.alignment=_align(wrap=True,v="top",indent=1)
    ws.row_dimensions[r].height=max(120,len(rev_text)//4)
    r+=2

    r=hdr(ws,r,2,"OPERATING EXPENSES COMMENTARY",8)
    opex_text=commentary.get("opex","")
    if not opex_text: opex_text="[Commentary not generated — check API key and re-run]"
    ws.merge_cells(f"B{r}:H{r}")
    c=ws.cell(r,2,opex_text); c.font=_font(size=11,color=DARK); c.fill=_f(CREAM)
    c.alignment=_align(wrap=True,v="top",indent=1)
    ws.row_dimensions[r].height=max(120,len(opex_text)//4)


def build_bridge_tab(wb, summary):
    ws=wb.create_sheet("03 P&L Bridge"); ws.sheet_view.showGridLines=False
    set_widths(ws,{"A":2,"B":36,"C":16,"D":16,"E":16,"F":14,"G":2})

    r=1; ws.merge_cells(f"B{r}:F{r}")
    c=ws.cell(r,2,f"P&L BRIDGE  ·  {Y0_LABEL} → {Y1_LABEL}")
    c.font=_font(bold=True,size=14,color=WHITE); c.fill=_f(BLACK); c.alignment=_align(indent=1); ws.row_dimensions[r].height=30
    r+=1; ws.merge_cells(f"B{r}:F{r}")
    c=ws.cell(r,2,"Blue = input  |  Black = formula  |  GREEN = Favourable for profit  |  RED = Adverse for profit")
    c.font=_font(size=9,color=MID,italic=True); c.fill=_f(LIGHT); c.alignment=_align(indent=1); ws.row_dimensions[r].height=16
    r+=2

    pl=summary[summary["Report"]=="Profit and Loss"]
    def pl_s(sub): return pl[pl["SubClass2"]==sub]["Y0"].sum(), pl[pl["SubClass2"]==sub]["Y1"].sum()
    ns_y0,ns_y1=pl_s("Sales"); cos_y0,cos_y1=pl_s("Cost of Sales")
    gp_y0,gp_y1=ns_y0+cos_y0,ns_y1+cos_y1
    op_y0=pl[pl["Class"]=="Operating account"]["Y0"].sum()
    op_y1=pl[pl["Class"]=="Operating account"]["Y1"].sum()
    eb_y0,eb_y1=gp_y0+op_y0,gp_y1+op_y1

    steps=[("Net Revenue",ns_y0,ns_y1,"bridge"),("  Cost of Sales",cos_y0,cos_y1,"bridge"),
           ("GROSS PROFIT",gp_y0,gp_y1,"total")]
    for sub in ["Sales & Distribution","Marketing","Administration","Depreciation & Amortization"]:
        v0,v1=pl_s(sub)
        if v0!=0 or v1!=0: steps.append((f"  {sub}",v0,v1,"bridge"))
    steps.append(("EBITDA (Operating Profit)",eb_y0,eb_y1,"total"))

    r=col_hdrs(ws,r,2,["Bridge Step",Y0_LABEL,Y1_LABEL,"Movement","Movement %"],right=(1,2,3,4))
    data_start=r
    for lbl,v0,v1,kind in steps:
        var=v1-v0; fav=is_fav(var); total=(kind=="total")
        bg=BLACK if total else (FAV_BG if fav else ADV_BG)
        tc=WHITE if total else (FAV_TXT if fav else ADV_TXT)
        ws.cell(r,2,lbl).font=_font(bold=total,size=10,color=tc)
        ws.cell(r,2).fill=_f(bg); ws.cell(r,2).alignment=_align(indent=1); ws.cell(r,2).border=_bot()
        for col,val in [(3,v0),(4,v1)]:
            c=ws.cell(r,col,val)
            c.font=(_font(bold=True,size=10,color=WHITE) if total else _fi(10))
            c.fill=_f(bg); c.alignment=_align("right"); c.number_format='#,##0;(#,##0);"-"'; c.border=_bot()
        c=ws.cell(r,5,f"=D{r}-C{r}")
        c.font=(_font(bold=True,size=10,color=WHITE) if total else _ff(10))
        c.fill=_f(bg); c.alignment=_align("right"); c.number_format='#,##0;(#,##0);"-"'; c.border=_bot()
        c=ws.cell(r,6,f'=IF(C{r}=0,1,E{r}/ABS(C{r}))')
        c.font=(_font(bold=True,size=10,color=WHITE) if total else _ff(10))
        c.fill=_f(bg); c.alignment=_align("right"); c.number_format='+0.0%;-0.0%;"-"'; c.border=_bot()
        ws.row_dimensions[r].height=22 if total else 18; r+=1

    # Chart data — hidden rows (row dimension height=0)
    r+=1
    chart_data=[(s[0].strip(),s[1],s[2]) for s in steps if s[3]!="total"]
    chart_start=r
    for i,(lbl,v0,v1) in enumerate(chart_data):
        ws.cell(r+i,2,lbl)
        ws.cell(r+i,3,round(v0))
        ws.cell(r+i,4,round(v1))
        ws.cell(r+i,5,f"=D{r+i}-C{r+i}"); ws.cell(r+i,5).number_format='#,##0;(#,##0);"-"'
        ws.row_dimensions[r+i].hidden=True  # HIDDEN — chart source data only

    chart=BarChart(); chart.type="bar"
    chart.title=f"P&L Bridge: {Y0_LABEL} → {Y1_LABEL}"
    chart.y_axis.title="Movement"; chart.shape=4
    data=Reference(ws,min_col=5,min_row=chart_start,max_row=chart_start+len(chart_data)-1)
    cats=Reference(ws,min_col=2,min_row=chart_start,max_row=chart_start+len(chart_data)-1)
    chart.add_data(data); chart.set_categories(cats)
    chart.series[0].graphicalProperties.solidFill="FF3C1F"
    chart.width=24; chart.height=14
    ws.add_chart(chart,f"B{r+len(chart_data)+2}")


def build_material_tab(wb, summary, y0, y1):
    ws=wb.create_sheet("04 Material Variances"); ws.sheet_view.showGridLines=False
    # FIXED: column K was width=2 (spacer). Now set to 50 for Transaction Drivers
    set_widths(ws,{"A":2,"B":5,"C":28,"D":20,"E":16,"F":16,"G":14,"H":12,"I":10,"J":24,"K":50,"L":2})

    r=1; ws.merge_cells(f"B{r}:K{r}")
    c=ws.cell(r,2,"MATERIAL VARIANCES  ·  P&L ACCOUNTS ONLY")
    c.font=_font(bold=True,size=14,color=WHITE); c.fill=_f(BLACK); c.alignment=_align(indent=1); ws.row_dimensions[r].height=30
    r+=1; ws.merge_cells(f"B{r}:K{r}")
    c=ws.cell(r,2,"GREEN = Favourable (revenue UP or cost DOWN)  |  RED = Adverse (revenue DOWN or cost UP)  |  Blue = input  |  Black = formula")
    c.font=_font(size=9,color=MID,italic=True); c.fill=_f(LIGHT); c.alignment=_align(indent=1); ws.row_dimensions[r].height=16
    r+=2

    hdrs=["#","Account","Category",Y0_LABEL,Y1_LABEL,"Variance","Var %","F/A","Confidence","Transaction Drivers (top 3)"]
    r=col_hdrs(ws,r,2,hdrs,right=(3,4,5,6))

    pl_mat=summary[(summary["Report"]=="Profit and Loss") &
                   (~summary.index.isin(ONEOFF_KEYS)) &
                   (summary.apply(is_material,axis=1))].copy()
    pl_mat=pl_mat.reindex(pl_mat["Var"].abs().sort_values(ascending=False).index)

    for rank,(_,row) in enumerate(pl_mat.iterrows(),1):
        var=row["Var"]; fav=is_fav(var)
        pct_abs=abs(row["VarPct"]) if row["VarPct"] else 0
        conf=("HIGH — requires review" if pct_abs>10 and abs(var)>50000 else
              "MEDIUM — verify with owner" if pct_abs>5 else "LOW — minor")
        drv=get_drivers(y0,y1,row.name,3)
        drv_txt="  |  ".join([f"{i}→Δ{v['Var']:+,.0f}" for i,v in drv.iterrows()]) if not drv.empty else "—"
        bg=FAV_BG if fav else ADV_BG; tc=FAV_TXT if fav else ADV_TXT

        for col,val,fnt,fmt,al in [
            (2,rank,_font(size=10,color=MID),None,"center"),
            (3,row["Account"],_font(size=10,bold=True,color=tc),None,"left"),
            (4,row["SubClass2"],_font(size=10,color=MID),None,"left"),
            (5,row["Y0"],_fi(10),'#,##0;(#,##0);"-"',"right"),
            (6,row["Y1"],_fi(10),'#,##0;(#,##0);"-"',"right"),
        ]:
            c=ws.cell(r,col,val); c.font=fnt; c.fill=_f(bg)
            c.alignment=_align(al,indent=1); c.border=_bot()
            if fmt: c.number_format=fmt
        # Variance formula
        c=ws.cell(r,7,f"=F{r}-E{r}"); c.font=_ff(10); c.fill=_f(bg)
        c.alignment=_align("right"); c.number_format='#,##0;(#,##0);"-"'; c.border=_bot()
        # Var% formula
        c=ws.cell(r,8,f'=IF(E{r}=0,1,G{r}/ABS(E{r}))'); c.font=_ff(10); c.fill=_f(bg)
        c.alignment=_align("right"); c.number_format='+0.0%;-0.0%;"-"'; c.border=_bot()
        fa_write(ws,r,9,var); ws.cell(r,9).fill=_f(bg); ws.cell(r,9).border=_bot()
        c=ws.cell(r,10,conf); c.font=_font(size=9,color=MID); c.fill=_f(bg)
        c.alignment=_align(indent=1); c.border=_bot()
        # Drivers — col 11 = K — now width 50
        c=ws.cell(r,11,drv_txt); c.font=_font(size=9,color=DARK); c.fill=_f(bg)
        c.alignment=_align(wrap=True,indent=1); c.border=_bot()
        ws.row_dimensions[r].height=30; r+=1

    r+=1
    r=hdr(ws,r,2,"NOTE: Balance sheet and one-off capital items excluded from this table",11,bg=RED)
    ws.merge_cells(f"B{r}:K{r}")
    c=ws.cell(r,2,"Share Capital, PP&E, Investments, Long-Term Obligations excluded — Y0 capital events with no Y1 equivalent. See Tab 05 for balance sheet flux.")
    c.font=_font(size=9,color=MID,italic=True); c.fill=_f(CREAM); c.alignment=_align(wrap=True,indent=1); ws.row_dimensions[r].height=30


def build_bs_tab(wb, summary):
    ws=wb.create_sheet("05 Balance Sheet Flux"); ws.sheet_view.showGridLines=False
    set_widths(ws,{"A":2,"B":32,"C":16,"D":16,"E":14,"F":12,"G":44,"H":2})

    r=1; ws.merge_cells(f"B{r}:G{r}")
    c=ws.cell(r,2,f"BALANCE SHEET FLUX  ·  {Y0_LABEL} vs {Y1_LABEL}")
    c.font=_font(bold=True,size=14,color=WHITE); c.fill=_f(BLACK); c.alignment=_align(indent=1); ws.row_dimensions[r].height=30
    r+=1; ws.merge_cells(f"B{r}:G{r}")
    c=ws.cell(r,2,"One-off capital events (greyed/italic) = Y0 only, no Y1 equivalent  |  Blue = input  |  Black = formula")
    c.font=_font(size=9,color=MID,italic=True); c.fill=_f(LIGHT); c.alignment=_align(indent=1); ws.row_dimensions[r].height=16
    r+=2

    r=col_hdrs(ws,r,2,["Account",Y0_LABEL,Y1_LABEL,"Movement","Mvmt %","Note"],right=(1,2,3,4))

    bs=summary[summary["Report"]=="Balance Sheet"].copy().sort_values(["SubClass2","Account"])
    cur_sub=None
    for _,row in bs.iterrows():
        if row["SubClass2"]!=cur_sub:
            cur_sub=row["SubClass2"]
            ws.merge_cells(f"B{r}:G{r}")
            c=ws.cell(r,2,str(cur_sub).upper()); c.font=_font(bold=True,size=9,color=WHITE)
            c.fill=_f(RED); c.alignment=_align(indent=1); ws.row_dimensions[r].height=18; r+=1

        var=row["Var"]; one_off=(row.name in ONEOFF_KEYS)
        pct=row["VarPct"]
        mat_bs=(abs(var)>=THRESHOLDS["Balance Sheet"]["abs"] and
                (abs(pct)>=THRESHOLDS["Balance Sheet"]["pct"] if pct else False))
        bg=LIGHT if one_off else WHITE
        note=("One-off Y0 event — no Y1 equivalent" if one_off else
              "Material — verify with controller" if mat_bs else
              "Below BS materiality threshold — no action required")

        c=ws.cell(r,2,f"  {row['Account']}"); c.font=_font(size=10,color=MID,italic=one_off)
        c.fill=_f(bg); c.alignment=_align(indent=2); c.border=_bot("E0E0E0")
        for col,val in [(3,row["Y0"]),(4,row["Y1"])]:
            c=ws.cell(r,col,val); c.font=(_fi(10) if not one_off else _font(size=10,color=MID,italic=True))
            c.fill=_f(bg); c.alignment=_align("right"); c.number_format='#,##0;(#,##0);"-"'; c.border=_bot("E0E0E0")
        c=ws.cell(r,5,f"=D{r}-C{r}"); c.font=(_ff(10) if not one_off else _font(size=10,color=MID,italic=True))
        c.fill=_f(bg); c.alignment=_align("right"); c.number_format='#,##0;(#,##0);"-"'; c.border=_bot("E0E0E0")
        c=ws.cell(r,6,f'=IF(C{r}=0,1,E{r}/ABS(C{r}))'); c.font=(_ff(10) if not one_off else _font(size=10,color=MID,italic=True))
        c.fill=_f(bg); c.alignment=_align("right"); c.number_format='+0.0%;-0.0%;"-"'; c.border=_bot("E0E0E0")
        c=ws.cell(r,7,note); c.font=_font(size=9,color=MID,italic=True)
        c.fill=_f(bg); c.alignment=_align(indent=1); c.border=_bot("E0E0E0")
        ws.row_dimensions[r].height=18; r+=1

    r+=1; ws.merge_cells(f"B{r}:G{r}")
    c=ws.cell(r,2,"⚠  Balance sheet movements reflect GL aggregation only. Accruals, provisions and reclassification adjustments require controller review before inclusion in any board pack.")
    c.font=_font(size=9,color=AMBER); c.fill=_f(CREAM); c.alignment=_align(wrap=True,indent=1); ws.row_dimensions[r].height=30


def build_trail_tab(wb, summary, y0, y1, sa_map=None):
    ws=wb.create_sheet("06 Transaction Trail"); ws.sheet_view.showGridLines=False
    set_widths(ws,{"A":2,"B":30,"C":30,"D":16,"E":16,"F":14,"G":2})

    r=1; ws.merge_cells(f"B{r}:F{r}")
    c=ws.cell(r,2,"TRANSACTION AUDIT TRAIL  ·  GL evidence for every material P&L variance")
    c.font=_font(bold=True,size=14,color=WHITE); c.fill=_f(BLACK); c.alignment=_align(indent=1); ws.row_dimensions[r].height=30
    r+=1; ws.merge_cells(f"B{r}:F{r}")
    c=ws.cell(r,2,"Every AI commentary claim is grounded in the GL transaction data shown below. Verify each line before sharing in a board pack. Green header = favourable overall. Red = adverse.")
    c.font=_font(size=9,color=MID,italic=True); c.fill=_f(CREAM); c.alignment=_align(wrap=True,indent=1); ws.row_dimensions[r].height=30
    r+=2

    pl_mat=summary[(summary["Report"]=="Profit and Loss") &
                   (~summary.index.isin(ONEOFF_KEYS)) &
                   (summary.apply(is_material,axis=1))]
    pl_mat=pl_mat.reindex(pl_mat["Var"].abs().sort_values(ascending=False).index)

    for _,mat_row in pl_mat.iterrows():
        fav=is_fav(mat_row["Var"])
        ws.merge_cells(f"B{r}:F{r}")
        c=ws.cell(r,2,f"▌ {mat_row['Account']}  ({mat_row['SubClass2']})  —  Total Δ {mat_row['Var']:+,.0f}")
        c.font=_font(bold=True,size=11,color=WHITE)
        c.fill=_f("2E7D32" if fav else RED); c.alignment=_align(indent=1); ws.row_dimensions[r].height=22; r+=1
        for i,h in enumerate(["Transaction Type","Sub-Category",f"{Y0_LABEL} Amount",f"{Y1_LABEL} Amount","Variance"]):
            c=ws.cell(r,2+i,h); c.font=_font(bold=True,size=9,color=WHITE); c.fill=_f(DARK)
            c.alignment=_align("right" if i in(2,3,4) else "left",indent=1)
        ws.row_dimensions[r].height=16; r+=1
        d=get_drivers(y0,y1,mat_row.name,8)
        if d.empty:
            ws.cell(r,2,"No transaction-level detail available.").font=_font(size=9,color=MID); r+=1
        else:
            for idx,drow in d.iterrows():
                dfav=is_fav(drow["Var"]); bg=FAV_BG if dfav else ADV_BG
                c=ws.cell(r,2,idx); c.font=_font(size=10,color=DARK); c.fill=_f(bg); c.alignment=_align(indent=1); c.border=_bot("E0E0E0")
                sub_cat = (sa_map or {}).get(mat_row.name, "")
                c=ws.cell(r,3,sub_cat); c.font=_font(size=9,color=MID); c.fill=_f(bg)
                c.alignment=_align(indent=1); c.border=_bot("E0E0E0")
                for col,val in [(4,drow["Y0"]),(5,drow["Y1"])]:
                    c=ws.cell(r,col,val); c.font=_fi(10); c.fill=_f(bg); c.alignment=_align("right")
                    c.number_format='#,##0;(#,##0);"-"'; c.border=_bot("E0E0E0")
                c=ws.cell(r,6,f"=E{r}-D{r}"); c.font=_ff(10); c.fill=_f(bg)
                c.alignment=_align("right"); c.number_format='#,##0;(#,##0);"-"'; c.border=_bot("E0E0E0")
                ws.row_dimensions[r].height=16; r+=1
        r+=1


def build_opex_tab(wb, summary, y0, y1):
    ws=wb.create_sheet("07 Opex Breakdown"); ws.sheet_view.showGridLines=False
    set_widths(ws,{"A":2,"B":28,"C":16,"D":16,"E":14,"F":12,"G":8,"H":32,"I":2})

    r=1; ws.merge_cells(f"B{r}:H{r}")
    c=ws.cell(r,2,"OPERATING EXPENSES  ·  Category Breakdown")
    c.font=_font(bold=True,size=14,color=WHITE); c.fill=_f(BLACK); c.alignment=_align(indent=1); ws.row_dimensions[r].height=30
    r+=1; ws.merge_cells(f"B{r}:H{r}")
    c=ws.cell(r,2,"Green = cost decreased (Fav)  |  Red = cost increased (Adv)  |  Blue = input  |  Black = formula")
    c.font=_font(size=9,color=MID,italic=True); c.fill=_f(LIGHT); c.alignment=_align(indent=1); ws.row_dimensions[r].height=16
    r+=2

    pl=summary[summary["Report"]=="Profit and Loss"]
    for cat in ["Sales & Distribution","Marketing","Administration","Depreciation & Amortization"]:
        rows=pl[pl["SubClass2"]==cat]
        if rows.empty: continue
        ws.merge_cells(f"B{r}:H{r}")
        c=ws.cell(r,2,cat.upper()); c.font=_font(bold=True,size=10,color=WHITE); c.fill=_f(RED)
        c.alignment=_align(indent=1); ws.row_dimensions[r].height=20; r+=1
        r=col_hdrs(ws,r,2,["Account",Y0_LABEL,Y1_LABEL,"Variance","Var %","F/A","Key Driver"],right=(1,2,3,4))
        cat_start=r
        for _,row in rows.iterrows():
            var=row["Var"]; fav=is_fav(var)
            bg=FAV_BG if(abs(var)>5000 and fav) else ADV_BG if(abs(var)>5000 and not fav) else WHITE
            drv=get_drivers(y0,y1,row.name,1)
            drv_txt=f"{drv.index[0]}→Δ{drv.iloc[0]['Var']:+,.0f}" if not drv.empty else "—"
            c=ws.cell(r,2,row["Account"]); c.font=_font(size=10,color=DARK); c.fill=_f(bg)
            c.alignment=_align(indent=1); c.border=_bot("E0E0E0")
            for col,val in [(3,row["Y0"]),(4,row["Y1"])]:
                c=ws.cell(r,col,val); c.font=_fi(10); c.fill=_f(bg)
                c.alignment=_align("right"); c.number_format='#,##0;(#,##0);"-"'; c.border=_bot("E0E0E0")
            c=ws.cell(r,5,f"=D{r}-C{r}"); c.font=_ff(10); c.fill=_f(bg)
            c.alignment=_align("right"); c.number_format='#,##0;(#,##0);"-"'; c.border=_bot("E0E0E0")
            c=ws.cell(r,6,f'=IF(C{r}=0,1,E{r}/ABS(C{r}))'); c.font=_ff(10); c.fill=_f(bg)
            c.alignment=_align("right"); c.number_format='+0.0%;-0.0%;"-"'; c.border=_bot("E0E0E0")
            fa_write(ws,r,7,var); ws.cell(r,7).fill=_f(bg); ws.cell(r,7).border=_bot("E0E0E0")
            c=ws.cell(r,8,drv_txt); c.font=_font(size=9,color=MID); c.fill=_f(bg)
            c.alignment=_align(indent=1); c.border=_bot("E0E0E0")
            ws.row_dimensions[r].height=18; r+=1
        # Category total using SUM formulas
        tot=r; y0_rng=f"C{cat_start}:C{r-1}"; y1_rng=f"D{cat_start}:D{r-1}"
        wf=_font(bold=True,size=10,color=WHITE)
        ws.cell(tot,2,f"TOTAL {cat}").font=wf; ws.cell(tot,2).fill=_f(BLACK); ws.cell(tot,2).alignment=_align(indent=1)
        ws.cell(tot,3,f"=SUM({y0_rng})").font=wf; ws.cell(tot,3).fill=_f(BLACK); ws.cell(tot,3).alignment=_align("right"); ws.cell(tot,3).number_format='#,##0;(#,##0);"-"'
        ws.cell(tot,4,f"=SUM({y1_rng})").font=wf; ws.cell(tot,4).fill=_f(BLACK); ws.cell(tot,4).alignment=_align("right"); ws.cell(tot,4).number_format='#,##0;(#,##0);"-"'
        ws.cell(tot,5,f"=D{tot}-C{tot}").font=wf; ws.cell(tot,5).fill=_f(BLACK); ws.cell(tot,5).alignment=_align("right"); ws.cell(tot,5).number_format='#,##0;(#,##0);"-"'
        ws.cell(tot,6,f'=IF(C{tot}=0,1,E{tot}/ABS(C{tot}))').font=wf; ws.cell(tot,6).fill=_f(BLACK); ws.cell(tot,6).alignment=_align("right"); ws.cell(tot,6).number_format='+0.0%;-0.0%;"-"'
        ws.row_dimensions[tot].height=20; r=tot+2


def build_stable_tab(wb, summary):
    ws=wb.create_sheet("08 Stable Accounts"); ws.sheet_view.showGridLines=False
    # FIXED: col widths extended — was cutting off Var% column
    set_widths(ws,{"A":2,"B":36,"C":22,"D":16,"E":16,"F":16,"G":14,"H":14,"I":2})

    r=1; ws.merge_cells(f"B{r}:H{r}")
    c=ws.cell(r,2,"STABLE ACCOUNTS  ·  Below materiality threshold")
    c.font=_font(bold=True,size=14,color=WHITE); c.fill=_f(BLACK); c.alignment=_align(indent=1); ws.row_dimensions[r].height=30
    r+=2

    mat_keys=set(summary[(summary["Report"]=="Profit and Loss") &
                          (summary.apply(is_material,axis=1))].index.tolist())
    stable=summary[~summary.index.isin(mat_keys|ONEOFF_KEYS)].copy()
    stable=stable[stable["Report"].isin(["Profit and Loss","Balance Sheet"])]

    ws.merge_cells(f"B{r}:H{r}")
    c=ws.cell(r,2,
        f"{len(stable)} accounts below the dual materiality threshold. "
        f"Cumulative P&L movement: "
        f"{stable[stable['Report']=='Profit and Loss']['Var'].sum():+,.0f}. "
        "No separate commentary required.")
    c.font=_font(size=10,color=DARK,italic=True); c.fill=_f(CREAM); c.alignment=_align(wrap=True,indent=1); ws.row_dimensions[r].height=30
    r+=2

    r=col_hdrs(ws,r,2,["Account","Category","Report",Y0_LABEL,Y1_LABEL,"Variance","Var %"],right=(3,4,5,6))

    for _,row in stable.sort_values("Report").iterrows():
        for col,val in [(2,row["Account"]),(3,row["SubClass2"]),(4,row["Report"])]:
            c=ws.cell(r,col,val); c.font=_font(size=9,color=MID); c.fill=_f(WHITE)
            c.alignment=_align(indent=1); c.border=_bot("EEEEEE")
        for col,val in [(5,row["Y0"]),(6,row["Y1"])]:
            c=ws.cell(r,col,val); c.font=_fi(9); c.fill=_f(WHITE)
            c.alignment=_align("right"); c.number_format='#,##0;(#,##0);"-"'; c.border=_bot("EEEEEE")
        c=ws.cell(r,7,f"=F{r}-E{r}"); c.font=_ff(9); c.fill=_f(WHITE)
        c.alignment=_align("right"); c.number_format='#,##0;(#,##0);"-"'; c.border=_bot("EEEEEE")
        c=ws.cell(r,8,f'=IF(E{r}=0,1,G{r}/ABS(E{r}))'); c.font=_ff(9); c.fill=_f(WHITE)
        c.alignment=_align("right"); c.number_format='+0.0%;-0.0%;"-"'; c.border=_bot("EEEEEE")
        ws.row_dimensions[r].height=15; r+=1


def build_methodology_tab(wb, summary, commentary, user_context=""):
    ws=wb.create_sheet("09 Methodology"); ws.sheet_view.showGridLines=False
    set_widths(ws,{"A":2,"B":28,"C":72,"D":2})

    r=hdr(ws,1,2,"METHODOLOGY, THRESHOLDS & AUDIT LOG",3); r+=1

    secs=[
        ("DATA SOURCES",[
            ("Y0 File",Y0_FILE),("Y1 File",Y1_FILE),
            ("Y0 Period",Y0_LABEL),("Y1 Period",Y1_LABEL),
            ("Generated",datetime.now().strftime("%d %B %Y, %H:%M")),
            ("Tool","Variance Commentary Generator v6.0 — alloutoftokens.com"),
        ]),
        ("COLOUR & FORMULA CONVENTION",[
            ("GREEN = Favourable","Revenue increased OR cost decreased vs prior year. Applies to all accounts."),
            ("RED = Adverse","Revenue decreased OR cost increased vs prior year."),
            ("Blue text","Hardcoded input values (Y0 and Y1 GL totals)."),
            ("Black text","Formula-calculated values (Variance=Y1-Y0, Var%=Variance/ABS(Y0))."),
            ("Universal is_fav rule","Var > 0 = Favourable for ALL accounts. Costs stored negative in GL; Var>0 means less negative = lower cost."),
        ]),
        ("MATERIALITY THRESHOLDS",[(k,f"Abs ≥ {v['abs']:,}  AND  Pct ≥ {v['pct']}%") for k,v in THRESHOLDS.items()]),
        ("EXCLUSIONS",[
            ("One-off capital accounts","Keys 75,80,90,100,150,160,170,180,190 excluded from P&L variance table."),
            ("Balance Sheet threshold","Raised to 10% AND 50,000 to filter minor accrual noise."),
        ]),
        ("AI ENGINE",[
            ("Model","claude-opus-4-6 via Anthropic Messages API"),
            ("Commentary parser","Fixed in v5/v6: sections saved correctly when next header encountered. exec and revenue sections were always empty in v4 due to parser bug."),
            ("Grounding rule","All commentary grounded in GL transaction data only."),
            ("Human review","REQUIRED before sharing — AI draft only."),
        ]),
        ("LIMITATIONS",[
            ("No budget comparison","Actuals vs prior year only — no accountability vs plan."),
            ("Root cause","AI commentary indicative. Management context not included unless context.txt populated."),
            ("Management context","If context.txt was populated, AI had that information during commentary generation."),
        ]),
        ("SIGN-OFF CHECKLIST",[
            ("Step 1","Finance reviewer: verify AI commentary Tabs 01-04"),
            ("Step 2","Budget owners: add context where flagged 'requires management input'"),
            ("Step 3","Controller: confirm balance sheet movements Tab 05"),
            ("Step 4","Finance Director: approve Tab 01 before board distribution"),
        ]),
    ]

    for sec_title,rows in secs:
        r=hdr(ws,r,2,sec_title,3)
        for k,v in rows:
            ws.cell(r,2,k).font=_font(bold=True,size=10,color=DARK); ws.cell(r,2).fill=_f(CREAM); ws.cell(r,2).alignment=_align(indent=1)
            ws.cell(r,3,v).font=_font(size=10,color=DARK); ws.cell(r,3).fill=_f(CREAM)
            ws.cell(r,3).alignment=_align(wrap=True,indent=1); ws.cell(r,3).border=_bot(); ws.row_dimensions[r].height=18; r+=1
        r+=1

    if user_context:
        r=hdr(ws,r,2,"USER-PROVIDED MANAGEMENT CONTEXT (used in AI commentary generation)",3)
        ws.merge_cells(f"B{r}:C{r}")
        c=ws.cell(r,2,user_context); c.font=_font(size=10,color=DARK,italic=True); c.fill=_f(CREAM)
        c.alignment=_align(wrap=True,v="top",indent=1); ws.row_dimensions[r].height=max(60,len(user_context)//4)
        r+=2

    r=hdr(ws,r,2,"FULL AI OUTPUT (unedited — for audit trail only)",3)
    ws.merge_cells(f"B{r}:C{r}")
    c=ws.cell(r,2,commentary.get("_raw",""))
    c.font=_font(size=9,color=MID,italic=True); c.fill=_f(WHITE)
    c.alignment=_align(wrap=True,v="top",indent=1)
    ws.row_dimensions[r].height=max(200,len(commentary.get("_raw",""))//4)


# ── HTML DASHBOARD ────────────────────────────────────────────────────────────

def build_html_dashboard(summary, commentary, y0, y1, terr):
    pl=summary[summary["Report"]=="Profit and Loss"]
    ns_y0=pl[pl["SubClass2"]=="Sales"]["Y0"].sum(); ns_y1=pl[pl["SubClass2"]=="Sales"]["Y1"].sum()
    cos_y0=pl[pl["SubClass2"]=="Cost of Sales"]["Y0"].sum(); cos_y1=pl[pl["SubClass2"]=="Cost of Sales"]["Y1"].sum()
    gp_y0,gp_y1=ns_y0+cos_y0,ns_y1+cos_y1
    op_y0=pl[pl["Class"]=="Operating account"]["Y0"].sum(); op_y1=pl[pl["Class"]=="Operating account"]["Y1"].sum()
    eb_y0,eb_y1=gp_y0+op_y0,gp_y1+op_y1

    m_y0=monthly_series(y0,210); m_y1=monthly_series(y1,210)
    t_y0=territory_rev(y0,210); t_y1=territory_rev(y1,210)
    terr_map=terr.set_index("Territory_key")["Country"].to_dict() if not terr.empty else {}
    terr_data=[{"country":terr_map.get(k,f"T{k}"),"y0":round(t_y0.get(k,0)),
                "y1":round(t_y1.get(k,0)),"var":round(t_y1.get(k,0)-t_y0.get(k,0))}
               for k in sorted(set(t_y0)|set(t_y1))]

    bridge=[{"label":"Net Revenue","y0":round(ns_y0),"y1":round(ns_y1),"type":"bridge"},
            {"label":"Cost of Sales","y0":round(cos_y0),"y1":round(cos_y1),"type":"bridge"},
            {"label":"Gross Profit","y0":round(gp_y0),"y1":round(gp_y1),"type":"total"}]
    for sub in ["Sales & Distribution","Marketing","Administration"]:
        v0=pl[pl["SubClass2"]==sub]["Y0"].sum(); v1=pl[pl["SubClass2"]==sub]["Y1"].sum()
        bridge.append({"label":sub,"y0":round(v0),"y1":round(v1),"type":"bridge"})
    bridge.append({"label":"EBITDA","y0":round(eb_y0),"y1":round(eb_y1),"type":"total"})

    pl_mat=summary[(summary["Report"]=="Profit and Loss") &
                   (~summary.index.isin(ONEOFF_KEYS)) & (summary.apply(is_material,axis=1))].copy()
    pl_mat=pl_mat.reindex(pl_mat["Var"].abs().sort_values(ascending=False).index)

    variances_js=[]
    for _,r in pl_mat.iterrows():
        d=get_drivers(y0,y1,r.name,2)
        drv=[f"{i}: Δ{v['Var']:+,.0f}" for i,v in d.iterrows()] if not d.empty else []
        variances_js.append({"account":r["Account"],"category":r["SubClass2"],
                             "y0":round(r["Y0"]),"y1":round(r["Y1"]),"var":round(r["Var"]),
                             "pct":round(r["VarPct"],1) if r["VarPct"] else 0,
                             "fav":bool(is_fav(r["Var"])),"drivers":drv})

    class_order={"Trading account":0,"Operating account":1,"Non-operating":2,"Interest & Tax":3}
    pl_sorted=pl.copy(); pl_sorted["_ord"]=pl_sorted["Class"].map(class_order).fillna(99)
    pl_sorted=pl_sorted.sort_values(["_ord","SubClass2","Account"])
    pl_rows=[{"account":r["Account"],"class":r["Class"],"sub":r["SubClass2"],
              "y0":round(r["Y0"]),"y1":round(r["Y1"]),"var":round(r["Var"]),
              "pct":round(r["VarPct"],1) if r["VarPct"] else 0,
              "mat":bool(is_material(r)),"fav":bool(is_fav(r["Var"]))}
             for _,r in pl_sorted.iterrows()]

    data_js = f"""const D={{
  y0:"{Y0_LABEL}",y1:"{Y1_LABEL}",
  gen:"{datetime.now().strftime('%d %B %Y')}",
  kpis:{{
    rev:{{a:{round(ns_y0)},b:{round(ns_y1)}}},
    gp:{{a:{round(gp_y0)},b:{round(gp_y1)}}},
    ebitda:{{a:{round(eb_y0)},b:{round(eb_y1)}}},
    gpm:{{a:{round(gp_y0/ns_y0*100,1)},b:{round(gp_y1/ns_y1*100,1)}}}
  }},
  mrev:{{a:{m_y0},b:{m_y1}}},
  terr:{json.dumps(terr_data)},
  bridge:{json.dumps(bridge)},
  vars:{json.dumps(variances_js)},
  pl:{json.dumps(pl_rows)},
  cmt:{{
    exec:{json.dumps(commentary.get("exec",""))},
    rev:{json.dumps(commentary.get("revenue",""))},
    opex:{json.dumps(commentary.get("opex",""))},
    mgmt:{json.dumps(commentary.get("mgmt",""))}
  }}
}};"""

    # NOTE: All JS uses single braces {{}}-free since we use a separate var
    # instead of f-string for the entire script block
    script = """
const fmt=n=>new Intl.NumberFormat('en-GB').format(Math.round(n));
const pct=n=>(n>=0?'+':'')+n.toFixed(1)+'%';
const MO=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const RED='#FF3C1F',FAV='#1B5E20',ADV='#B71C1C',MID='#6B6560';

function show(id,btn){
  document.querySelectorAll('.sec').forEach(s=>s.classList.remove('on'));
  document.querySelectorAll('.ntab').forEach(t=>t.classList.remove('on'));
  document.getElementById(id).classList.add('on');
  if(btn)btn.classList.add('on');
  if(id==='rev')buildRevCharts();
  if(id==='vars'){buildVarChart();renderVars();}
  if(id==='pl')buildPL();
  if(id==='cmt')fillCmt();
}

// KPIs
function buildKPIs(){
  const tiles=[['NET REVENUE','rev',false],['GROSS PROFIT','gp',false],['EBITDA','ebitda',false],['GP MARGIN %','gpm',true]];
  document.getElementById('kpis').innerHTML=tiles.map(([lbl,k,ip])=>{
    const d=D.kpis[k],v=d.b-d.a,fav=v>0;
    const vs=ip?d.b.toFixed(1)+'%':fmt(d.b);
    const bv=ip?`${fav?'▲':'▼'} ${Math.abs(v).toFixed(1)}pp`:`${fav?'▲':'▼'} ${fmt(Math.abs(v))}  (${pct(d.a?v/Math.abs(d.a)*100:0)})`;
    return `<div class="kpi ${fav?'fav':'adv'}">
      <div class="kl">${lbl}</div>
      <div class="kv">${vs}</div>
      <div class="kb ${fav?'fav':'adv'}">${bv}</div>
      <div class="ks">Prior: ${ip?d.a.toFixed(1)+'%':fmt(d.a)}</div>
    </div>`;
  }).join('');
}

// Charts
let C={};
function mkChart(id,cfg){if(C[id])C[id].destroy();C[id]=new Chart(document.getElementById(id).getContext('2d'),cfg);}

function buildRevChart(id){
  mkChart(id,{type:'line',data:{labels:MO,datasets:[
    {label:D.y1,data:D.mrev.b,borderColor:RED,backgroundColor:'rgba(255,60,31,.08)',
     borderWidth:2.5,fill:true,tension:.35,pointRadius:4,pointBackgroundColor:RED},
    {label:D.y0,data:D.mrev.a,borderColor:MID,borderDash:[4,4],borderWidth:1.5,
     tension:.35,pointRadius:3,pointBackgroundColor:MID}
  ]},options:{responsive:true,maintainAspectRatio:true,
    plugins:{legend:{position:'top'},tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${fmt(c.raw)}`}}},
    scales:{y:{ticks:{callback:v=>fmt(v)}}}}});
}

function buildBridge(){
  const nt=D.bridge.filter(s=>s.type!=='total');
  mkChart('bridgeChart',{type:'bar',data:{labels:nt.map(s=>s.label),
    datasets:[{label:'Movement',data:nt.map(s=>s.y1-s.y0),
      backgroundColor:nt.map(s=>(s.y1-s.y0)>0?'rgba(27,94,32,.7)':'rgba(183,28,28,.7)'),
      borderRadius:3}]},
    options:{responsive:true,maintainAspectRatio:true,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`Movement: ${fmt(c.raw)}`}}},
      scales:{y:{ticks:{callback:v=>fmt(v)}}}}});
}

function buildTerrChart(){
  mkChart('terrChart',{type:'bar',data:{labels:D.terr.map(t=>t.country),
    datasets:[
      {label:D.y0,data:D.terr.map(t=>t.y0),backgroundColor:'rgba(107,101,96,.25)',borderRadius:2},
      {label:D.y1,data:D.terr.map(t=>t.y1),backgroundColor:'rgba(255,60,31,.7)',borderRadius:2}
    ]},options:{responsive:true,maintainAspectRatio:true,
      plugins:{legend:{position:'top'},tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${fmt(c.raw)}`}}},
      scales:{y:{ticks:{callback:v=>fmt(v)}}}}});
}

function buildTerrVar(){
  mkChart('terrVar',{type:'bar',data:{labels:D.terr.map(t=>t.country),
    datasets:[{label:'YoY',data:D.terr.map(t=>t.var),
      backgroundColor:D.terr.map(t=>t.var>0?'rgba(27,94,32,.7)':'rgba(183,28,28,.7)'),
      borderRadius:3}]},
    options:{responsive:true,maintainAspectRatio:true,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`${fmt(c.raw)}`}}},
      scales:{y:{ticks:{callback:v=>fmt(v)}}}}});
}

function buildRevCharts(){buildRevChart('revChart2');buildTerrVar();}

let vf='all',vs='',vsort={k:'var',d:-1};
const mx=D.vars.length?Math.max(...D.vars.map(v=>Math.abs(v.var))):1;

function renderVars(){
  let rows=D.vars.filter(v=>{
    if(vf==='fav'&&!v.fav)return false;
    if(vf==='adv'&&v.fav)return false;
    if(vs&&!v.account.toLowerCase().includes(vs)&&!v.category.toLowerCase().includes(vs))return false;
    return true;
  });
  rows=[...rows].sort((a,b)=>{
    const av=a[vsort.k]||0,bv=b[vsort.k]||0;
    return typeof av==='string'?av.localeCompare(bv)*vsort.d:(av-bv)*vsort.d;
  });
  document.getElementById('vcnt').textContent=`Showing ${rows.length} of ${D.vars.length}`;
  const bw=v=>Math.round(Math.abs(v)/mx*80);
  document.getElementById('vtbody').innerHTML=rows.map(v=>`<tr>
    <td><strong>${v.account}</strong></td>
    <td style="color:var(--mid);font-size:12px">${v.category}</td>
    <td class="nr">${fmt(v.y0)}</td>
    <td class="nr">${fmt(v.y1)}</td>
    <td class="nr"><div style="display:flex;align-items:center;gap:8px;justify-content:flex-end">
      <strong>${fmt(v.var)}</strong>
      <div style="height:6px;border-radius:3px;width:${bw(v.var)}px;background:${v.fav?FAV:ADV}"></div>
    </div></td>
    <td class="nr" style="color:${v.fav?FAV:ADV};font-weight:600">${pct(v.pct)}</td>
    <td><span class="badge ${v.fav?'fav':'adv'}">${v.fav?'✓ Fav':'✗ Adv'}</span></td>
    <td>${v.drivers.map(d=>`<span class="tag">${d}</span>`).join('')}</td>
  </tr>`).join('');
}

function fv(f,btn){vf=f;document.querySelectorAll('.fb').forEach(b=>b.classList.remove('on'));btn.classList.add('on');renderVars();}
function sv(v){vs=v.toLowerCase();renderVars();}
function sortV(k,e){vsort.d=vsort.k===k?-vsort.d:-1;vsort.k=k;document.querySelectorAll('.vth').forEach(t=>t.classList.remove('s'));(e||event).target.classList.add('s');renderVars();}

let vchart=null;
function buildVarChart(){
  const top=D.vars.slice(0,15);
  if(vchart)vchart.destroy();
  vchart=new Chart(document.getElementById('varChart').getContext('2d'),{
    type:'bar',data:{labels:top.map(v=>v.account),
      datasets:[{label:'Variance',data:top.map(v=>v.var),
        backgroundColor:top.map(v=>v.fav?'rgba(27,94,32,.7)':'rgba(183,28,28,.7)'),borderRadius:3}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`${fmt(c.raw)}`}}},
      scales:{x:{ticks:{callback:v=>fmt(v)}}}}});
  document.getElementById('varChart').style.maxHeight='400px';
}

function buildPL(){
  let h='',lc=null,ls=null;
  const st={Rev:{a:0,b:0},GP:{a:0,b:0},EB:{a:0,b:0}};
  D.pl.forEach(r=>{
    if(r.class!==lc){lc=r.class;ls=null;h+=`<tr class="ch"><td colspan="6">${r.class.toUpperCase()}</td></tr>`;}
    if(r.sub!==ls){ls=r.sub;h+=`<tr class="sh"><td colspan="6" style="padding-left:20px">${r.sub}</td></tr>`;}
    const cl=r.mat?(r.fav?'mf':'ma'):'';
    h+=`<tr class="${cl}"><td style="padding-left:28px">${r.account}</td>
      <td class="nr">${fmt(r.y0)}</td><td class="nr">${fmt(r.y1)}</td>
      <td class="nr">${fmt(r.var)}</td>
      <td class="nr" style="color:${r.fav?FAV:ADV}">${pct(r.pct)}</td>
      <td>${r.mat?`<span class="badge ${r.fav?'fav':'adv'}">${r.fav?'✓':'✗'}</span>`:''}</td></tr>`;
    if(r.sub==='Sales'){st.Rev.a+=r.y0;st.Rev.b+=r.y1;}
    if(r.class==='Trading account'){st.GP.a+=r.y0;st.GP.b+=r.y1;}
    if(r.class==='Trading account'||r.class==='Operating account'){st.EB.a+=r.y0;st.EB.b+=r.y1;}
  });
  [['Net Revenue',st.Rev],['Gross Profit',st.GP],['EBITDA',st.EB]].forEach(([lb,s])=>{
    const v=s.b-s.a,fav=v>0;
    h+=`<tr class="st"><td>${lb}</td><td class="nr">${fmt(s.a)}</td><td class="nr">${fmt(s.b)}</td>
      <td class="nr">${fmt(v)}</td><td class="nr">${pct(s.a?v/Math.abs(s.a)*100:0)}</td>
      <td><span class="badge ${fav?'fav':'adv'}">${fav?'Fav':'Adv'}</span></td></tr>`;
  });
  document.getElementById('pltbody').innerHTML=h;
}

function fillCmt(){
  ['exec','rev','opex','mgmt'].forEach(k=>{
    const txt=D.cmt[k]||'No commentary generated.';
    ['c'+k,'c'+k+'2'].forEach(id=>{
      const el=document.getElementById(id);
      if(el)el.textContent=txt;
    });
  });
}

// Init
document.getElementById('meta').textContent=`Generated ${D.gen}  ·  alloutoftokens.com  ·  Finance × AI: In Practice  ·  v6.0`;
document.getElementById('exec0').textContent=D.cmt.exec||'';
buildKPIs();
buildRevChart('revChart');
buildBridge();
buildTerrChart();
renderVars();
fillCmt();
"""

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Variance Analysis · {Y0_LABEL} vs {Y1_LABEL}</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;900&family=DM+Sans:wght@300;400;500;600&family=Space+Mono:wght@400;700&display=swap" rel="stylesheet"/>
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
:root{{--black:#0D0D0D;--cream:#F5F0E8;--red:#FF3C1F;--dark:#1A1714;--mid:#6B6560;--light:#EDE8E0;--white:#fff;--fav-bg:#E8F5E9;--adv-bg:#FFF0ED;--fav:#1B5E20;--adv:#B71C1C;--r:4px;--sh:0 2px 16px rgba(13,13,13,.08)}}
html{{scroll-behavior:smooth}}body{{font-family:'DM Sans',sans-serif;background:var(--cream);color:var(--dark);font-size:15px;line-height:1.6}}
nav{{position:sticky;top:0;z-index:100;background:var(--black);display:flex;align-items:center;justify-content:space-between;padding:0 40px;height:56px;border-bottom:2px solid var(--red)}}
.nlogo{{font-family:'Space Mono',monospace;font-size:12px;font-weight:700;color:var(--white);letter-spacing:.1em;text-transform:uppercase}}.nlogo span{{color:var(--red)}}
.ntabs{{display:flex;gap:4px}}.ntab{{background:none;border:none;font-family:'DM Sans',sans-serif;font-size:12px;color:rgba(255,255,255,.5);padding:6px 14px;cursor:pointer;border-radius:3px;transition:all .2s}}.ntab:hover,.ntab.on{{color:var(--white);background:rgba(255,255,255,.08)}}.ntab.on{{border-bottom:2px solid var(--red)}}
main{{max-width:1400px;margin:0 auto;padding:40px}}.sec{{display:none}}.sec.on{{display:block}}
.legend{{display:flex;gap:20px;flex-wrap:wrap;margin-bottom:32px;padding:14px 20px;background:var(--white);border-radius:var(--r);border-left:3px solid var(--red)}}
.li{{display:flex;align-items:center;gap:8px;font-size:12px;color:var(--mid)}}.dot{{width:12px;height:12px;border-radius:2px;flex-shrink:0}}.dot.fav{{background:var(--fav-bg);border:1px solid var(--fav)}}.dot.adv{{background:var(--adv-bg);border:1px solid var(--adv)}}
.ph{{margin-bottom:40px}}.ph h1{{font-family:'Playfair Display',serif;font-size:clamp(28px,4vw,44px);font-weight:900;letter-spacing:-.02em;color:var(--black)}}.ph h1 em{{color:var(--red);font-style:italic}}
.pm{{font-family:'Space Mono',monospace;font-size:11px;color:var(--mid);margin-top:6px;letter-spacing:.06em}}
.kgrid{{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:40px}}
.kpi{{background:var(--white);border-radius:var(--r);padding:24px;box-shadow:var(--sh);position:relative;overflow:hidden}}.kpi::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px}}.kpi.fav::before{{background:var(--fav)}}.kpi.adv::before{{background:var(--adv)}}
.kl{{font-family:'Space Mono',monospace;font-size:10px;font-weight:700;letter-spacing:.14em;text-transform:uppercase;color:var(--mid);margin-bottom:10px}}
.kv{{font-family:'Playfair Display',serif;font-size:clamp(22px,2.5vw,32px);font-weight:700;color:var(--black);margin-bottom:6px}}
.kb{{display:inline-flex;align-items:center;gap:6px;font-size:12px;font-weight:600;padding:3px 10px;border-radius:20px}}.kb.fav{{background:var(--fav-bg);color:var(--fav)}}.kb.adv{{background:var(--adv-bg);color:var(--adv)}}
.ks{{font-size:11px;color:var(--mid);margin-top:6px}}
.cgrid{{display:grid;grid-template-columns:1fr 1fr;gap:24px;margin-bottom:40px}}
.cc{{background:var(--white);border-radius:var(--r);padding:28px;box-shadow:var(--sh)}}.fc{{grid-column:1/-1}}
.ct{{font-family:'Space Mono',monospace;font-size:11px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--red);margin-bottom:4px;display:flex;align-items:center;gap:8px}}.ct::before{{content:'';width:20px;height:2px;background:var(--red)}}
.cs{{font-size:13px;color:var(--mid);margin-bottom:20px}}canvas{{max-height:280px}}.fcc{{margin-bottom:24px}}.fcc canvas{{max-height:220px}}
.cmtgrid{{display:grid;grid-template-columns:1fr 1fr;gap:24px;margin-bottom:40px}}.cmtcard{{background:var(--white);border-radius:var(--r);padding:28px;box-shadow:var(--sh)}}.cmtcard.full{{grid-column:1/-1}}
.cmtlbl{{font-family:'Space Mono',monospace;font-size:10px;font-weight:700;letter-spacing:.14em;text-transform:uppercase;color:var(--red);margin-bottom:16px;display:flex;align-items:center;gap:8px}}.cmtlbl::before{{content:'';width:20px;height:2px;background:var(--red)}}
.cmttxt{{font-size:14px;line-height:1.8;color:var(--dark)}}
.ai{{display:inline-block;font-family:'Space Mono',monospace;font-size:9px;letter-spacing:.1em;background:var(--black);color:var(--white);padding:3px 8px;border-radius:2px;margin-bottom:12px}}
.vtw{{background:var(--white);border-radius:var(--r);padding:28px;box-shadow:var(--sh);margin-bottom:24px}}
.tctrl{{display:flex;align-items:center;gap:12px;margin-bottom:20px;flex-wrap:wrap}}
.fb{{background:var(--light);border:none;font-family:'DM Sans',sans-serif;font-size:12px;padding:6px 14px;border-radius:20px;cursor:pointer;transition:all .2s;color:var(--dark)}}.fb:hover,.fb.on{{background:var(--black);color:var(--white)}}
.si{{border:1px solid var(--light);background:var(--white);font-family:'DM Sans',sans-serif;font-size:13px;padding:6px 14px;border-radius:20px;color:var(--dark);width:200px;outline:none}}.si:focus{{border-color:var(--red)}}
table.vt{{width:100%;border-collapse:collapse;font-size:13px}}.vth{{text-align:left;padding:10px 12px;font-family:'Space Mono',monospace;font-size:10px;letter-spacing:.08em;text-transform:uppercase;color:var(--mid);border-bottom:2px solid var(--light);white-space:nowrap;cursor:pointer}}.vth:hover{{color:var(--red)}}.vth.s{{color:var(--red)}}
.vt td{{padding:12px 12px;border-bottom:1px solid var(--light);vertical-align:top}}.vt tr:hover td{{background:rgba(245,240,232,.5)}}
.badge{{display:inline-block;padding:2px 8px;border-radius:3px;font-size:11px;font-weight:600}}.badge.fav{{background:var(--fav-bg);color:var(--fav)}}.badge.adv{{background:var(--adv-bg);color:var(--adv)}}
.tag{{font-size:10px;background:var(--light);color:var(--mid);padding:2px 7px;border-radius:3px;font-family:'Space Mono',monospace;display:inline-block;margin:2px 2px 0 0}}
.pltw{{background:var(--white);border-radius:var(--r);padding:28px;box-shadow:var(--sh);margin-bottom:24px}}
table.plt{{width:100%;border-collapse:collapse;font-size:13px}}.plt th{{padding:10px 12px;font-family:'Space Mono',monospace;font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:var(--mid);border-bottom:2px solid var(--light)}}
.plt td{{padding:9px 12px;border-bottom:1px solid rgba(237,232,224,.6)}}
.plt tr.ch td{{background:var(--red);color:var(--white);font-weight:700;font-size:11px;letter-spacing:.06em;text-transform:uppercase}}
.plt tr.sh td{{background:var(--light);font-weight:600;font-size:12px;color:var(--dark)}}
.plt tr.mf td{{background:var(--fav-bg);color:var(--fav);font-weight:600}}
.plt tr.ma td{{background:var(--adv-bg);color:var(--adv);font-weight:600}}
.plt tr.st td{{background:var(--black);color:var(--white);font-weight:700;font-size:13px}}
.nr{{text-align:right;font-family:'Space Mono',monospace;font-size:12px}}
@media(max-width:900px){{.kgrid{{grid-template-columns:1fr 1fr}}.cgrid,.cmtgrid{{grid-template-columns:1fr}}nav{{padding:0 20px}}main{{padding:20px}}}}
@media(max-width:540px){{.kgrid{{grid-template-columns:1fr}}.ntabs{{display:none}}}}
.fade{{animation:fi .4s ease both}}@keyframes fi{{from{{opacity:0;transform:translateY(12px)}}to{{opacity:1;transform:none}}}}
</style>
</head>
<body>
<nav>
  <div class="nlogo">All Out <span>of</span> Tokens · Variance Analysis</div>
  <div class="ntabs">
    <button class="ntab on"  onclick="show('ov',this)">Overview</button>
    <button class="ntab"     onclick="show('rev',this)">Revenue</button>
    <button class="ntab"     onclick="show('vars',this)">Variances</button>
    <button class="ntab"     onclick="show('pl',this)">P&L Detail</button>
    <button class="ntab"     onclick="show('cmt',this)">Commentary</button>
  </div>
</nav>
<main>
<div class="legend">
  <div class="li"><div class="dot fav"></div> Favourable — revenue increased OR cost decreased vs prior year</div>
  <div class="li"><div class="dot adv"></div> Adverse — revenue decreased OR cost increased vs prior year</div>
  <div class="li" style="margin-left:auto;font-size:11px;color:var(--mid)">AI commentary · verify before use</div>
</div>

<!-- OVERVIEW -->
<section id="ov" class="sec on">
  <div class="ph fade"><h1>Variance Analysis · <em>{Y0_LABEL} vs {Y1_LABEL}</em></h1><div class="pm" id="meta"></div></div>
  <div class="kgrid" id="kpis"></div>
  <div class="cgrid">
    <div class="cc fade"><div class="ct">Monthly Revenue</div><div class="cs">Full year comparison · {Y0_LABEL} vs {Y1_LABEL}</div><canvas id="revChart"></canvas></div>
    <div class="cc fade"><div class="ct">P&L Bridge</div><div class="cs">Movement from prior year · green = fav · red = adv</div><canvas id="bridgeChart"></canvas></div>
  </div>
  <div class="cc fcc fade"><div class="ct">Revenue by Territory</div><div class="cs">{Y0_LABEL} vs {Y1_LABEL}</div><canvas id="terrChart"></canvas></div>
  <div class="cmtcard fade">
    <div class="cmtlbl">Executive Summary</div><div class="ai">AI-GENERATED · VERIFY BEFORE USE</div>
    <div class="cmttxt" id="exec0"></div>
  </div>
</section>

<!-- REVENUE -->
<section id="rev" class="sec">
  <div class="ph fade"><h1>Revenue <em>&amp; Gross Margin</em></h1><div class="pm">{Y0_LABEL} vs {Y1_LABEL}</div></div>
  <div class="cgrid">
    <div class="cc fade"><div class="ct">Monthly Revenue Trend</div><div class="cs">Month-by-month comparison</div><canvas id="revChart2"></canvas></div>
    <div class="cc fade"><div class="ct">Territory YoY Movement</div><div class="cs">Absolute Δ per market · green = growth</div><canvas id="terrVar"></canvas></div>
  </div>
  <div class="cmtgrid">
    <div class="cmtcard fade"><div class="cmtlbl">Revenue &amp; Gross Margin</div><div class="ai">AI-GENERATED · VERIFY BEFORE USE</div><div class="cmttxt" id="crev"></div></div>
    <div class="cmtcard fade"><div class="cmtlbl">Operating Expenses</div><div class="ai">AI-GENERATED · VERIFY BEFORE USE</div><div class="cmttxt" id="copex"></div></div>
  </div>
</section>

<!-- VARIANCES -->
<section id="vars" class="sec">
  <div class="ph fade"><h1>Material <em>Variances</em></h1><div class="pm">P&L accounts · ranked by absolute impact · one-off capital items excluded</div></div>
  <div class="vtw fade">
    <div class="tctrl">
      <button class="fb on" onclick="fv('all',this)">All</button>
      <button class="fb"    onclick="fv('fav',this)">Favourable</button>
      <button class="fb"    onclick="fv('adv',this)">Adverse</button>
      <input class="si" type="text" placeholder="Search accounts..." oninput="sv(this.value)"/>
      <span style="font-size:12px;color:var(--mid);margin-left:auto" id="vcnt"></span>
    </div>
    <table class="vt">
      <thead><tr>
        <th class="vth" onclick="sortV('account',event)">Account</th>
        <th class="vth" onclick="sortV('category',event)">Category</th>
        <th class="vth nr" onclick="sortV('y0',event)">{Y0_LABEL}</th>
        <th class="vth nr" onclick="sortV('y1',event)">{Y1_LABEL}</th>
        <th class="vth nr" onclick="sortV('var',event)">Variance</th>
        <th class="vth nr" onclick="sortV('pct',event)">Var %</th>
        <th class="vth">F/A</th>
        <th class="vth">Drivers</th>
      </tr></thead>
      <tbody id="vtbody"></tbody>
    </table>
  </div>
  <div class="cc fcc fade">
    <div class="ct">Top Variances · Absolute Impact</div>
    <div class="cs">Green = Favourable · Red = Adverse</div>
    <canvas id="varChart"></canvas>
  </div>
</section>

<!-- P&L DETAIL -->
<section id="pl" class="sec">
  <div class="ph fade"><h1>P&amp;L <em>Detail</em></h1><div class="pm">Full profit &amp; loss · material variances highlighted</div></div>
  <div class="pltw fade">
    <table class="plt">
      <thead><tr>
        <th style="width:36%">Account</th><th class="nr">{Y0_LABEL}</th>
        <th class="nr">{Y1_LABEL}</th><th class="nr">Variance</th>
        <th class="nr">Var %</th><th>F/A</th>
      </tr></thead>
      <tbody id="pltbody"></tbody>
    </table>
  </div>
</section>

<!-- COMMENTARY -->
<section id="cmt" class="sec">
  <div class="ph fade"><h1>Management <em>Commentary</em></h1><div class="pm">AI-generated · verify before use · {datetime.now().strftime('%d %B %Y')}</div></div>
  <div class="cmtgrid">
    <div class="cmtcard full fade"><div class="cmtlbl">Executive Summary</div><div class="ai">AI-GENERATED · VERIFY BEFORE USE</div><div class="cmttxt" id="cexec"></div></div>
    <div class="cmtcard fade"><div class="cmtlbl">Revenue &amp; Gross Margin</div><div class="ai">AI-GENERATED · VERIFY BEFORE USE</div><div class="cmttxt" id="crev2"></div></div>
    <div class="cmtcard fade"><div class="cmtlbl">Operating Expenses</div><div class="ai">AI-GENERATED · VERIFY BEFORE USE</div><div class="cmttxt" id="copex2"></div></div>
    <div class="cmtcard full fade" style="border-left:4px solid var(--red)"><div class="cmtlbl">Areas Requiring Management Input</div><div class="cmttxt" id="cmgmt" style="white-space:pre-line"></div></div>
  </div>
</section>
</main>
<script>{data_js}</script>
<script>{script}</script>
</body></html>"""
    return html


# ── ORCHESTRATION ─────────────────────────────────────────────────────────────

def preflight_check():
    """Check all requirements before running. Exit with clear messages if anything is missing."""
    errors = []
    if not os.path.exists(Y0_FILE):
        errors.append(f"  ✗ Missing: {Y0_FILE}  (rename your prior-year file to {Y0_FILE})")
    if not os.path.exists(Y1_FILE):
        errors.append(f"  ✗ Missing: {Y1_FILE}  (rename your current-year file to {Y1_FILE})")
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key or not api_key.startswith("sk-ant"):
        errors.append("  ✗ API key not set. Set environment variable: ANTHROPIC_API_KEY=your-key-here")
        errors.append("    Example (Mac/Linux): export ANTHROPIC_API_KEY=sk-ant-...")
        errors.append("    Example (Windows):   set ANTHROPIC_API_KEY=sk-ant-...")
    if errors:
        print("\n" + "="*60)
        print("  SETUP INCOMPLETE — fix these issues before running:")
        print()
        for e in errors: print(e)
        print()
        print("  See README.md for full setup instructions.")
        print("="*60 + "\n")
        sys.exit(1)


def main():
    preflight_check()

    print(f"\n{'='*60}")
    print(f"  Variance Commentary Generator v6.0")
    print(f"  {Y0_LABEL} vs {Y1_LABEL}")
    print(f"{'='*60}\n")

    print("[ 1/6 ] Loading GL data...")
    y0, y1, coa, terr, sa_map = load_data()
    print(f"        Y0: {len(y0):,} rows  |  Y1: {len(y1):,} rows")

    print("[ 2/6 ] Building variance summary...")
    summary = build_summary(y0, y1, coa)
    pl_mat = summary[
        (summary["Report"] == "Profit and Loss") &
        (~summary.index.isin(ONEOFF_KEYS)) &
        (summary.apply(is_material, axis=1))
    ]
    print(f"        Total accounts: {len(summary)}  |  Material P&L: {len(pl_mat)}")

    # Quick F/A sanity check
    print()
    print("        F/A sanity check:")
    for _, r in list(pl_mat.iterrows())[:5]:
        fav = is_fav(r["Var"])
        print(f"          {r['Account']:<28} Var={r['Var']:>10,.0f}  {'✓ Fav' if fav else '✗ Adv'}")

    user_context = load_user_context()
    if user_context:
        print(f"\n        Management context loaded from {CONTEXT_FILE} ({len(user_context)} chars)")

    print("\n[ 3/6 ] Calling Claude API for commentary...")
    commentary = generate_commentary(summary, y0, y1, user_context)

    # Verify commentary parsed correctly
    for k in ["exec", "revenue", "opex", "mgmt"]:
        txt = commentary.get(k, "")
        status = f"{len(txt)} chars" if txt else "EMPTY - will show placeholder"
        print(f"        {k}: {status}")

    print("\n[ 4/6 ] Building Excel workbook...")
    wb = Workbook()
    build_exec_tab(wb, summary, commentary, y0, y1)
    build_pl_tab(wb, summary, commentary, y0, y1)
    build_bridge_tab(wb, summary)
    build_material_tab(wb, summary, y0, y1)
    build_bs_tab(wb, summary)
    build_trail_tab(wb, summary, y0, y1, sa_map)
    build_opex_tab(wb, summary, y0, y1)
    build_stable_tab(wb, summary)
    build_methodology_tab(wb, summary, commentary, user_context)
    wb.save(XLS_OUTPUT)
    print(f"        Saved: {XLS_OUTPUT}")

    print("\n[ 5/6 ] Building HTML dashboard...")
    html = build_html_dashboard(summary, commentary, y0, y1, terr)
    with open(HTML_OUTPUT, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"        Saved: {HTML_OUTPUT}")

    print("\n[ 6/6 ] Done.\n")
    print(f"  Outputs:")
    print(f"    Excel:  {XLS_OUTPUT}")
    print(f"    HTML:   {HTML_OUTPUT}  ← open in browser for interactive dashboard")
    print(f"\n  ⚠  AI commentary is a FIRST DRAFT. Verify before sharing.\n")


if __name__ == "__main__":
    main()
