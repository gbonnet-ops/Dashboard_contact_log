#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Construit reconciliation_GL_vs_PnL.xlsx a partir de :
  - Grand_Livre_Consolide_2024-2025_vGB.csv  (charges par nature x entite)
  - P_L_consolide_par_unite_-_2025_06.xlsx   (onglet CB035-C, source de verite)
  - mapping_entites.csv / mapping_postes.csv  (mappings externalises, editables)
Tous les calculs du bridge sont FORMULES dans Excel (SUMIF/VLOOKUP), tracables
jusqu'aux postes GL d'origine (onglet GL_source).
Usage: python3 build_reconciliation.py
"""
import csv, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

GL_CSV  = sys.argv[1] if len(sys.argv)>1 else "/root/.claude/uploads/b7430bf7-92dd-59e5-af57-f052a95c7f18/a781960f-Grand_Livre_Consolide_20242025_vGB.csv"
PL_XLSX = sys.argv[2] if len(sys.argv)>2 else "/root/.claude/uploads/b7430bf7-92dd-59e5-af57-f052a95c7f18/7c169b89-PL_consolid__par_unit___2025.06.xlsx"
OUT     = "reconciliation_GL_vs_PnL.xlsx"

def num(s):
    s=s.strip()
    if s in ('-',''): return 0.0
    neg=s.startswith('(') and s.endswith(')')
    s=s.strip('()').replace(',','')
    return -float(s) if neg else float(s)

# ---- read GL ----
rows=list(csv.reader(open(GL_CSV,encoding='iso-8859-1')))
GL_ENT=rows[0][1:]                      # H2R ALLIANCE BM EST KCOLD RIVAGROUPE VALPOLIS
GL=[(r[0],[num(x) for x in r[1:1+len(GL_ENT)]]) for r in rows[1:]
    if r and r[0].strip() and r[0] not in ('Total','Consolid\xe9')]

# ---- read mappings ----
def readmap(path):
    rr=list(csv.reader(open(path,encoding='utf-8')))
    return rr[0], rr[1:]
ent_hdr, ent_rows = readmap("mapping_entites.csv")
pos_hdr, pos_rows = readmap("mapping_postes.csv")
ENT_MAP={r[0]:r[1] for r in ent_rows}                  # GL_ent -> PnL_ent (ORPHELIN for KCOLD)
PL_ENT=[ENT_MAP[g] for g in GL_ENT]                    # parallel to GL_ENT columns

# ---- read P&L target (CB035-C) ----
wb_pl=openpyxl.load_workbook(PL_XLSX,data_only=True); ws=wb_pl["CB035-C"]
PL_COLS={ws.cell(2,c).value:c for c in range(3,9)}     # entity name -> col idx (incl 'Total')
PL_LINES=["Achats consommés","Charges externes","Charges de personnel",
 "Autres charges d'exploitation","Impôts et taxes",
 "Variations nettes des amortissements et des dépréciations",
 "Opérations d'exploitation Intra-Groupe","Charges et produits financiers",
 "Opérations financières Intra-Groupe","Charges et produits exceptionnels",
 "Opérations exceptionnelles Intra-Groupe","Impôt sur les bénéfices"]
PL_ROWIDX={ws.cell(r,1).value:r for r in range(1,ws.max_row+1)}
def pl_val(line,entname):
    r=PL_ROWIDX.get(line); c=PL_COLS.get(entname)
    if not r or not c: return 0.0
    v=ws.cell(r,c).value
    return 0.0 if v in (None,'') else float(v)

# P&L target ordered by the 5 mapped P&L entities (KCOLD->ORPHELIN excluded)
PNL_ENT=[e for e in PL_ENT if e!="ORPHELIN"]           # 5 entities, GL order minus KCOLD

# ===================== BUILD WORKBOOK =====================
wb=openpyxl.Workbook()
BLUE=PatternFill("solid",fgColor="1F4E78"); LGREY=PatternFill("solid",fgColor="D9E1F2")
YEL=PatternFill("solid",fgColor="FFF2CC"); RED=PatternFill("solid",fgColor="F8CBAD")
GRN=PatternFill("solid",fgColor="C6E0B4")
WH=Font(color="FFFFFF",bold=True); B=Font(bold=True)
thin=Side(style="thin",color="BFBFBF"); BORD=Border(thin,thin,thin,thin)
EUR='#,##0;(#,##0)'; PCT='0.0%'
# Excel traite le texte litteral "#N/A" comme une valeur d'erreur -> on lui donne
# un libelle texte sur (meme poste, toujours tracable au GL). Applique aux 2 cotes
# du VLOOKUP pour que la cle corresponde.
def safe(lbl):
    return "Poste #N/A (non mappe GL)" if lbl=="#N/A" else lbl
def hdr(ws,row,vals,fill=BLUE,font=WH):
    for i,v in enumerate(vals,1):
        c=ws.cell(row,i,v); c.fill=fill; c.font=font; c.border=BORD
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)

# ---------- Sheet 0: Lisez-moi ----------
ws0=wb.active; ws0.title="Lisez-moi"
ws0.column_dimensions['A'].width=120
intro=[
 ("RAPPROCHEMENT GRAND LIVRE -> P&L CONSOLIDE",True),
 ("",False),
 ("Source de verite : P&L par destination (onglet CB035-C). Le GL est par nature x entite.",False),
 ("Le rapprochement porte sur les CHARGES + IS uniquement (le GL n'a pas de chiffre d'affaires).",False),
 ("",False),
 ("Onglets :",True),
 ("  Map_Entites   : table de mapping entite GL -> entite P&L (traceur = Charges de personnel)",False),
 ("  Map_Postes    : table de mapping poste GL -> ligne P&L (editable, recharger pour relancer)",False),
 ("  GL_source     : grand livre parse + ligne P&L mappee par VLOOKUP (col H) = base tracable",False),
 ("  GL_retraite   : GL agrege au format des lignes P&L, par entite (SUMIF sur GL_source)",False),
 ("  PnL_cible     : valeurs cibles lues dans CB035-C",False),
 ("  Bridge_entite : GL recalcule vs P&L cible, ecart EUR et %, par entite",False),
 ("  Bridge_conso  : bridge consolide (somme des 5 entites mappees)",False),
 ("  Ecarts        : ecarts > seuil avec explication (MAPPING vs DONNEES)",False),
 ("",False),
 ("Tous les montants du bridge sont des FORMULES Excel tracables jusqu'aux postes GL (GL_source).",False),
 ("KCOLD est l'entite ORPHELINE (aucune contrepartie P&L) : isolee, non incluse dans le consolide.",False),
]
for i,(t,bo) in enumerate(intro,1):
    c=ws0.cell(i,1,t); 
    if bo: c.font=B
ws0.cell(1,1).font=Font(bold=True,size=14,color="1F4E78")

# ---------- Sheet: Map_Entites ----------
me=wb.create_sheet("Map_Entites")
hdr(me,1,["entite_GL","entite_PnL","justification (traceur Charges de personnel)"])
for j,r in enumerate(ent_rows,2):
    for i,v in enumerate(r,1):
        cc=me.cell(j,i,v); cc.border=BORD
        if r[1]=="ORPHELIN": cc.fill=YEL
for w,col in zip((16,22,70),"ABC"): me.column_dimensions[col].width=w

# ---------- Sheet: Map_Postes ----------
mp=wb.create_sheet("Map_Postes")
hdr(mp,1,["poste_GL","ligne_PnL","note"])
for j,r in enumerate(pos_rows,2):
    for i,v in enumerate(r,1):
        cc=mp.cell(j,i,safe(v) if i==1 else v); cc.border=BORD
for w,col in zip((42,46,55),"ABC"): mp.column_dimensions[col].width=w
MP_FIRST,MP_LAST=2,1+len(pos_rows)

# ---------- Sheet: GL_source ----------
gs=wb.create_sheet("GL_source")
hdr(gs,1,["Poste_GL"]+GL_ENT+["Ligne_PnL (VLOOKUP)"])
GL_FIRST=2; n=len(GL)
for j,(p,vals) in enumerate(GL,2):
    gs.cell(j,1,safe(p)).border=BORD
    for i,v in enumerate(vals,2):
        cc=gs.cell(j,i,v); cc.number_format=EUR; cc.border=BORD
    # col H = mapped P&L line via VLOOKUP into Map_Postes
    hcol=len(GL_ENT)+2
    gs.cell(j,hcol,f"=VLOOKUP(A{j},Map_Postes!$A:$B,2,0)").border=BORD
GL_LAST=1+n
HCOL=len(GL_ENT)+2; HCOLL=get_column_letter(HCOL)
# total row (control)
tr=GL_LAST+1
gs.cell(tr,1,"Total (controle)").font=B
for i in range(2,2+len(GL_ENT)):
    cl=get_column_letter(i)
    cc=gs.cell(tr,i,f"=SUM({cl}{GL_FIRST}:{cl}{GL_LAST})"); cc.number_format=EUR; cc.font=B; cc.fill=LGREY
gs.column_dimensions['A'].width=42
for i in range(2,2+len(GL_ENT)): gs.column_dimensions[get_column_letter(i)].width=13
gs.column_dimensions[HCOLL].width=46
# map GL entity -> source column letter
GLCOL={g:get_column_letter(2+i) for i,g in enumerate(GL_ENT)}

# ---------- Sheet: GL_retraite ----------  (lines x PnL entities, SUMIF formulas)
gr=wb.create_sheet("GL_retraite")
cols=["Ligne P&L"]+PNL_ENT+["KCOLD (orphelin)","Total conso (5 ent.)"]
hdr(gr,1,cols)
# which GL column for each PnL entity:
PNLENT_GLCOL={}
for g in GL_ENT:
    pe=ENT_MAP[g]
    if pe!="ORPHELIN": PNLENT_GLCOL[pe]=GLCOL[g]
for j,line in enumerate(PL_LINES,2):
    gr.cell(j,1,line).border=BORD
    for k,pe in enumerate(PNL_ENT):
        col=PNLENT_GLCOL[pe]; cl=get_column_letter(2+k)
        f=f"=SUMIF(GL_source!${HCOLL}${GL_FIRST}:${HCOLL}${GL_LAST},$A{j},GL_source!${col}${GL_FIRST}:${col}${GL_LAST})"
        cc=gr.cell(j,2+k,f); cc.number_format=EUR; cc.border=BORD
    # KCOLD column
    kcol=GLCOL["KCOLD"]; kc=2+len(PNL_ENT)
    cc=gr.cell(j,kc,f"=SUMIF(GL_source!${HCOLL}${GL_FIRST}:${HCOLL}${GL_LAST},$A{j},GL_source!${kcol}${GL_FIRST}:${kcol}${GL_LAST})")
    cc.number_format=EUR; cc.border=BORD; cc.fill=YEL
    # Total conso = sum of 5 mapped entity cells
    tc=kc+1; first=get_column_letter(2); last=get_column_letter(1+len(PNL_ENT))
    cc=gr.cell(j,tc,f"=SUM({first}{j}:{last}{j})"); cc.number_format=EUR; cc.border=BORD; cc.font=B; cc.fill=LGREY
GR_FIRST=2; GR_LAST=1+len(PL_LINES)
# total row
tr=GR_LAST+1; gr.cell(tr,1,"TOTAL charges + IS").font=B
for i in range(2,2+len(PNL_ENT)+2):
    cl=get_column_letter(i); cc=gr.cell(tr,i,f"=SUM({cl}{GR_FIRST}:{cl}{GR_LAST})")
    cc.number_format=EUR; cc.font=B; cc.fill=LGREY
gr.column_dimensions['A'].width=52
for i in range(2,2+len(PNL_ENT)+2): gr.column_dimensions[get_column_letter(i)].width=15

# ---------- Sheet: PnL_cible ----------
pc=wb.create_sheet("PnL_cible")
hdr(pc,1,["Ligne P&L"]+PNL_ENT+["Total"])
for j,line in enumerate(PL_LINES,2):
    pc.cell(j,1,line).border=BORD
    for k,pe in enumerate(PNL_ENT):
        cc=pc.cell(j,2+k,pl_val(line,pe)); cc.number_format=EUR; cc.border=BORD
    tc=2+len(PNL_ENT); cc=pc.cell(j,tc,f"=SUM({get_column_letter(2)}{j}:{get_column_letter(1+len(PNL_ENT))}{j})")
    cc.number_format=EUR; cc.border=BORD; cc.font=B; cc.fill=LGREY
PC_FIRST=2; PC_LAST=1+len(PL_LINES)
# info rows: CA + autres produits (hors bridge charges)
inf=PC_LAST+2
pc.cell(inf,1,"INFO (produits - hors bridge charges, le GL n'a pas de CA) :").font=B
for o,line in enumerate(["Chiffre d'affaires","Autres produits d'exploitation"]):
    pc.cell(inf+1+o,1,line)
    for k,pe in enumerate(PNL_ENT):
        cc=pc.cell(inf+1+o,2+k,pl_val(line,pe)); cc.number_format=EUR
pc.column_dimensions['A'].width=52
for i in range(2,2+len(PNL_ENT)+1): pc.column_dimensions[get_column_letter(i)].width=15

# ---------- Sheet: Bridge_entite ----------
be=wb.create_sheet("Bridge_entite")
hdr(be,1,["Entité P&L","Ligne P&L","GL recalculé","P&L cible","Écart €","Écart %"])
row=2
for k,pe in enumerate(PNL_ENT):
    grcol=get_column_letter(2+k)   # column in GL_retraite & PnL_cible
    for j,line in enumerate(PL_LINES):
        glr=f"GL_retraite!{grcol}{GR_FIRST+j}"; plc=f"PnL_cible!{grcol}{PC_FIRST+j}"
        be.cell(row,1,pe).border=BORD
        be.cell(row,2,line).border=BORD
        c=be.cell(row,3,f"={glr}"); c.number_format=EUR; c.border=BORD
        c=be.cell(row,4,f"={plc}"); c.number_format=EUR; c.border=BORD
        c=be.cell(row,5,f"=C{row}-D{row}"); c.number_format=EUR; c.border=BORD
        c=be.cell(row,6,f'=IF(D{row}=0,"",E{row}/D{row})'); c.number_format=PCT; c.border=BORD
        row+=1
for w,col in zip((20,50,15,15,14,9),"ABCDEF"): be.column_dimensions[col].width=w

# ---------- Sheet: Bridge_conso ----------
bc=wb.create_sheet("Bridge_conso")
hdr(bc,1,["Ligne P&L","GL recalculé","P&L cible","Écart €","Écart %","Type d'écart"])
TOTCOL=get_column_letter(1+len(PNL_ENT)+2)   # 'Total conso' col in GL_retraite
PCTOT=get_column_letter(1+len(PNL_ENT)+1)    # 'Total' col in PnL_cible
expl={
 "Achats consommés":"DONNEES: reclassement par destination (P&L inclut des achats/sous-traitance non identifiables en nature dans le GL)",
 "Charges externes":"DONNEES: la nature 'charges externes' du GL > P&L par destination (reclassements + conso)",
 "Charges de personnel":"MAPPING ok (incl. #N/A). Residuel = ALLIANCE/GROUPE RIVALIS (-10 534, ecart de source)",
 "Autres charges d'exploitation":"MAPPING ok (ecart ~0)",
 "Impôts et taxes":"DONNEES: perimetre/reclassement partiel",
 "Variations nettes des amortissements et des dépréciations":"DONNEES: dotations GL > P&L (retraitements de consolidation, ecarts d'acquisition)",
 "Opérations d'exploitation Intra-Groupe":"ELIMINATION: GL = charge intra -4 314 875; P&L net a 0 (jambe produit absente du GL)",
 "Charges et produits financiers":"DONNEES: charges fin. GL > P&L (reclassement / conso)",
 "Opérations financières Intra-Groupe":"Pas de detail intra-groupe financier dans le GL (P&L net a 0)",
 "Charges et produits exceptionnels":"MAPPING ok ([obsolete] rattache); faible ecart",
 "Opérations exceptionnelles Intra-Groupe":"Aucun poste GL (P&L = 0)",
 "Impôt sur les bénéfices":"DONNEES: IS GL vs P&L (signe/credits d'impot, retraitement)",
}
for j,line in enumerate(PL_LINES):
    r=2+j
    bc.cell(r,1,line).border=BORD
    c=bc.cell(r,2,f"=GL_retraite!{TOTCOL}{GR_FIRST+j}"); c.number_format=EUR; c.border=BORD
    c=bc.cell(r,3,f"=PnL_cible!{PCTOT}{PC_FIRST+j}"); c.number_format=EUR; c.border=BORD
    c=bc.cell(r,4,f"=B{r}-C{r}"); c.number_format=EUR; c.border=BORD
    c=bc.cell(r,5,f'=IF(C{r}=0,"",D{r}/C{r})'); c.number_format=PCT; c.border=BORD
    bc.cell(r,6,expl[line]).border=BORD
tr=2+len(PL_LINES)
bc.cell(tr,1,"TOTAL charges + IS").font=B
for col in ("B","C","D"):
    c=bc.cell(tr,ord(col)-64,f"=SUM({col}2:{col}{tr-1})"); c.number_format=EUR; c.font=B; c.fill=LGREY
# control block
cb=tr+2
bc.cell(cb,1,"CONTROLES").font=B
bc.cell(cb+1,1,"GL retraite total 5 entites (conso)")
c=bc.cell(cb+1,2,f"=B{tr}"); c.number_format=EUR
bc.cell(cb+2,1,"KCOLD orphelin (isole, hors conso)")
kcoltot=get_column_letter(1+len(PNL_ENT)+1)
c=bc.cell(cb+2,2,f"=GL_retraite!{kcoltot}{GR_LAST+1}"); c.number_format=EUR
bc.cell(cb+3,1,"P&L cible total charges + IS")
c=bc.cell(cb+3,2,f"=C{tr}"); c.number_format=EUR
bc.cell(cb+4,1,"Ecart total GL vs P&L")
c=bc.cell(cb+4,2,f"=D{tr}"); c.number_format=EUR; c.font=B
bc.cell(cb+5,1,"dont elimination intra-groupe (op. expl. intra)")
_intra_row=2+PL_LINES.index("Opérations d'exploitation Intra-Groupe")
c=bc.cell(cb+5,2,f"=D{_intra_row}"); c.number_format=EUR
bc.cell(cb+6,1,"Consolide GL annonce (CSV) - NON reconcilie")
c=bc.cell(cb+6,2,-6191160); c.number_format=EUR; c.fill=RED
for w,col in zip((50,16,16,14,9,70),"ABCDEF"): bc.column_dimensions[col].width=w

# ---------- Sheet: Ecarts ----------
ec=wb.create_sheet("Ecarts")
hdr(ec,1,["Ligne P&L (conso)","Écart €","Écart %","Catégorie","Explication"])
# precompute numeric to know which exceed threshold (threshold 5 k€ or 1%)
# we recompute in python to order, but values displayed as references to Bridge_conso
import openpyxl as _o
# numeric recompute:
def num_line(line):
    glr=0.0
    for (p,vals) in GL:
        if M_LINE[p]==line:
            for g,v in zip(GL_ENT,vals):
                if ENT_MAP[g]!="ORPHELIN": glr+=v
    plc=sum(pl_val(line,pe) for pe in PNL_ENT)
    return glr,plc
M_LINE={r[0]:r[1] for r in pos_rows}
diffs=[]
for line in PL_LINES:
    glr,plc=num_line(line); d=glr-plc
    diffs.append((line,d,(d/plc if plc else 0)))
diffs.sort(key=lambda x:-abs(x[1]))
THRESH=5000
catmap={l:("ELIMINATION" if "Intra-Groupe" in l and "exploitation" in l else
           ("MAPPING" if l in ("Charges de personnel","Autres charges d'exploitation","Charges et produits exceptionnels") else "DONNEES"))
        for l in PL_LINES}
r=2
for line,d,pct in diffs:
    if abs(d)<THRESH: continue
    ec.cell(r,1,line).border=BORD
    c=ec.cell(r,2,round(d)); c.number_format=EUR; c.border=BORD
    c=ec.cell(r,3,pct); c.number_format=PCT; c.border=BORD
    cat=catmap[line]
    cc=ec.cell(r,4,cat); cc.border=BORD
    cc.fill=GRN if cat=="MAPPING" else (YEL if cat=="ELIMINATION" else RED)
    ec.cell(r,5,expl[line]).border=BORD
    r+=1
for w,col in zip((50,14,9,14,70),"ABCDE"): ec.column_dimensions[col].width=w

wb.save(OUT)
print("Saved",OUT)
