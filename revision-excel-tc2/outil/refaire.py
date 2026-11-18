"""Refait le sujet Excel TC2 à partir du fichier de départ (feuilles BDD, TVA, Exercice final).
Usage : python3 refaire.py sujet.xlsx resultat.xlsx
Formules écrites en anglais (Excel les affiche en français automatiquement)."""
import sys, copy, statistics
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import CellIsRule, IconSetRule, Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chartsheet import Chartsheet

src, out = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(src)
bdd = wb['BDD']; tva = wb['TVA']; ex = wb['Exercice final']
RATES = {tva.cell(r,1).value: tva.cell(r,3).value for r in range(3, tva.max_row+1) if tva.cell(r,1).value is not None}
EUR = '#,##0.00 "€"'; PCT = '0.00%'
GREEN = PatternFill('solid', fgColor='C6EFCE'); RED_F = Font(bold=True, color='C00000'); GREEN_F = Font(color='006100')
BOLD = Font(bold=True)

# ---------- 1. lire les clients ----------
hdr = [c.value for c in bdd[1]][:13]
rows = []
for r in range(2, bdd.max_row+1):
    v = [bdd.cell(r,c).value for c in range(1,14)]
    if v[0] is None: continue
    rows.append(v)
def key_name(v): return str(v[1]).strip().lower()
def calc(v):
    ttc, code = v[12], v[10]
    ht = ttc/(1+RATES[code]); n = ht*12
    for lim, t in [(10000,.05),(15000,.08),(25000,.10),(35000,.15),(50000,.20)]:
        if n < lim: return ht, n, n*t
    return ht, n, n*.25

def fill_sheet(ws, data, first_row=2):
    """écrit les clients + formules J, L, N, O ; renvoie la dernière ligne"""
    for i, h in enumerate(hdr + ['Montant annuel HT','Honoraires','Gestion des clients'], 1):
        ws.cell(1, i, h).font = BOLD
    r = first_row
    for v in data:
        for c in range(1,14):
            if c in (10,12): continue
            ws.cell(r, c, v[c-1])
        ws.cell(r,10, f'=IF(K{r}=1,M{r}/(1+TVA!$C$3),IF(K{r}=2,M{r}/(1+TVA!$C$4),M{r}/(1+TVA!$C$5)))')
        ws.cell(r,12, f'=IF(K{r}=1,J{r}*TVA!$C$3,IF(K{r}=2,J{r}*TVA!$C$4,J{r}*TVA!$C$5))')
        ws.cell(r,14, f'=J{r}*12')
        ws.cell(r,15, f'=IF(N{r}<10000,N{r}*5%,IF(N{r}<15000,N{r}*8%,IF(N{r}<25000,N{r}*10%,IF(N{r}<35000,N{r}*15%,IF(N{r}<50000,N{r}*20%,N{r}*25%)))))')
        for c in (10,12,13,14,15): ws.cell(r,c).number_format = EUR
        r += 1
    for col, w in zip('ABCDEFGHIJKLMNOP', [10,22,8,10,16,28,8,18,14,14,8,14,14,16,14,22]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'
    return r-1

def gestion(ws, last):
    dv = DataValidation(type='list', formula1='"Bon client,Client à rencontrer,Solvabilité à vérifier"', allow_blank=True)
    ws.add_data_validation(dv); dv.add(f'P2:P{last}')
    for txt, dxf in [('bon client', DifferentialStyle(fill=GREEN, font=GREEN_F)), ('solvabilité', DifferentialStyle(font=RED_F))]:
        rule = Rule(type='containsText', operator='containsText', text=txt, dxf=dxf)
        rule.formula = [f'NOT(ISERROR(SEARCH("{txt}",P2)))']
        ws.conditional_formatting.add(f'P2:P{last}', rule)

by_name = sorted(rows, key=key_name)

# ---------- FICHIER CLIENTS ----------
fc = wb.create_sheet('FICHIER CLIENTS', 0)
last = fill_sheet(fc, by_name); gestion(fc, last)

# ---------- Tarifs ----------
ta = wb.create_sheet('Tarifs', 1)
last_t = fill_sheet(ta, by_name); gestion(ta, last_t)
ta.conditional_formatting.add(f'M2:M{last_t}', CellIsRule(operator='lessThan', formula=['5000'], fill=PatternFill('solid', fgColor='FFC7CE'), font=Font(color='9C0006')))
ta.conditional_formatting.add(f'M2:M{last_t}', CellIsRule(operator='between', formula=['5000','10000'], fill=PatternFill('solid', fgColor='FFEB9C'), font=Font(color='9C5700')))

# ---------- LISTING (filtre TGE + au-dessus de la moyenne) ----------
li = wb.create_sheet('LISTING', 2)
last_l = fill_sheet(li, by_name); gestion(li, last_l)
annuels = [calc(v)[1] for v in by_name]; moy = statistics.mean(annuels)
li.auto_filter.ref = f'A1:P{last_l}'
li.auto_filter.add_filter_column(2, ['TGE'])
from openpyxl.worksheet.filters import DynamicFilter, FilterColumn
fcol = FilterColumn(colId=13, dynamicFilter=DynamicFilter(type='aboveAverage', val=moy))
li.auto_filter.filterColumn.append(fcol)
visibles = 0
for i, v in enumerate(by_name, start=2):
    keep = v[2] == 'TGE' and annuels[i-2] > moy
    li.row_dimensions[i].hidden = not keep; visibles += keep

# ---------- HONORAIRES (tri par type + sous-totaux) ----------
ho = wb.create_sheet('HONORAIRES', 3)
by_type = sorted(rows, key=lambda v: (v[2], key_name(v)))
for i, h in enumerate(hdr + ['Montant annuel HT','Honoraires','Gestion des clients'], 1): ho.cell(1,i,h).font = BOLD
r = 2; blocks = []
for t in sorted({v[2] for v in rows}):
    grp = [v for v in by_type if v[2] == t]
    start = r
    end = fill_sheet(ho, grp, first_row=r)
    for rr in range(start, end+1): ho.row_dimensions[rr].outlineLevel = 2
    r = end + 1
    for lab, col, fn in [(f'Total {t}', 'O', 9), (f'Total {t}', 'M', 9), (f'Moyenne {t}', 'N', 1)]:
        ho.cell(r, 3, lab).font = BOLD
        c = ho[f'{col}{r}']; c.value = f'=SUBTOTAL({fn},{col}{start}:{col}{end})'; c.font = BOLD; c.number_format = EUR
        ho.row_dimensions[r].outlineLevel = 1
        r += 1
    blocks.append((start, end))
for lab, col, fn in [('Total général','O',9), ('Total général','M',9), ('Moyenne','N',1)]:
    ho.cell(r,3,lab).font = BOLD
    c = ho[f'{col}{r}']; c.value = f'=SUBTOTAL({fn},{col}2:{col}{blocks[-1][1]})'; c.font = BOLD; c.number_format = EUR
    r += 1
ho.sheet_properties.outlinePr.summaryBelow = True

# ---------- TCD1 / TCD2 (tableaux de synthèse en formules) ----------
types = sorted({v[2] for v in rows})
t1 = wb.create_sheet('TCD1', 0)
t1['A1'] = 'Code TVA'; t1['B1'] = 3
t1.append([]); t1.append(['Type Société','Moyenne de Montant mensuel TTC','Max. de Montant mensuel TTC','Min. de Montant mensuel TTC'])
for c in t1[3]: c.font = BOLD
rng = lambda col: f"'FICHIER CLIENTS'!${col}$2:${col}${last}"
for t in types + ['Total général']:
    rr = t1.max_row + 1
    if t == 'Total général':
        cond = f"{rng('K')},$B$1"
    else:
        t1.cell(rr,1,t); cond = f"{rng('C')},A{rr},{rng('K')},$B$1"
    t1.cell(rr,1,t)
    t1.cell(rr,2,f"=_xlfn.AVERAGEIFS({rng('M')},{cond})" if False else f"=AVERAGEIFS({rng('M')},{cond})")
    t1.cell(rr,3,f"=_xlfn.MAXIFS({rng('M')},{cond})")
    t1.cell(rr,4,f"=_xlfn.MINIFS({rng('M')},{cond})")
    for c in (2,3,4): t1.cell(rr,c).number_format = EUR
t1[f'A{t1.max_row}'].font = BOLD
for col,w in zip('ABCD',[16,32,30,30]): t1.column_dimensions[col].width = w

t2 = wb.create_sheet('TCD2', 1)
t2['A1'] = 'Ville Société'; t2['B1'] = '(Tous)'
t2.append([]); t2.append(['Type Société','Somme de Montant mensuel TTC (% du total)'])
for c in t2[3]: c.font = BOLD
for t in types:
    rr = t2.max_row + 1
    t2.cell(rr,1,t); t2.cell(rr,2,f"=SUMIF({rng('C')},A{rr},{rng('M')})/SUM({rng('M')})").number_format = PCT
rr = t2.max_row + 1
t2.cell(rr,1,'Total général').font = BOLD; t2.cell(rr,2,f'=SUM(B4:B{rr-1})').number_format = PCT
t2.column_dimensions['A'].width = 16; t2.column_dimensions['B'].width = 40

# ---------- Exercice final ----------
first = 4; lastc = max(r for r in range(first, ex.max_row+1) if ex.cell(r,1).value)
tot = lastc + 1
for col, h in zip('NOPQ', ['CA TOTAL','Part du CA / ccial','Taux évolution S2/S1','Montant prime']):
    ex[f'{col}3'] = h; ex[f'{col}3'].font = BOLD
for r in range(first, lastc+1):
    ex[f'N{r}'] = f'=SUM(B{r}:M{r})'
    ex[f'O{r}'] = f'=N{r}/$N${tot}'; ex[f'O{r}'].number_format = PCT
    ex[f'P{r}'] = f'=(SUM(H{r}:M{r})-SUM(B{r}:G{r}))/SUM(B{r}:G{r})'; ex[f'P{r}'].number_format = PCT
    ex[f'Q{r}'] = f'=IF(N{r}>60000,N{r}*10%,500)'; ex[f'Q{r}'].number_format = EUR
ex[f'A{tot}'] = 'CA TOTAL'; ex[f'A{tot}'].font = BOLD
for col in 'BCDEFGHIJKLMN':
    ex[f'{col}{tot}'] = f'=SUM({col}{first}:{col}{lastc})'; ex[f'{col}{tot}'].font = BOLD
ex.conditional_formatting.add(f'O{first}:O{lastc}', IconSetRule('3TrafficLights1', 'num', [0, 0.10, 0.15]))
ex.conditional_formatting.add(f'P{first}:P{lastc}', IconSetRule('3Arrows', 'num', [-1, -0.0001, 0]))
for col, w in zip('NOPQ', [12,18,20,16]): ex.column_dimensions[col].width = w

# ---------- Graphiques ----------
g1 = BarChart(); g1.type = 'col'; g1.grouping = 'clustered'
g1.title = 'Répartition CA 2ème semestre'; g1.x_axis.title = 'Commerciaux'; g1.y_axis.title = 'CA en €'
g1.add_data(Reference(ex, min_col=8, max_col=13, min_row=3, max_row=lastc), titles_from_data=True)
g1.set_categories(Reference(ex, min_col=1, min_row=first, max_row=lastc))
g1.x_axis.delete = False; g1.y_axis.delete = False
cs1 = wb.create_chartsheet('Graphique1'); cs1.add_chart(g1)
g2 = PieChart(); g2.title = 'Part du CA total par commercial'
g2.add_data(Reference(ex, min_col=14, min_row=3, max_row=lastc), titles_from_data=True)
g2.set_categories(Reference(ex, min_col=1, min_row=first, max_row=lastc))
g2.dataLabels = DataLabelList(showPercent=True, showVal=False, showSerName=False, showCatName=False, showLegendKey=False)
cs2 = wb.create_chartsheet('Graphique2'); cs2.add_chart(g2)

# ordre des onglets comme le corrigé
order = ['TCD1','TCD2','FICHIER CLIENTS','HONORAIRES','LISTING','Tarifs','TVA','Graphique1','Graphique2','Exercice final']
del wb['BDD']
wb._sheets = [wb[n] for n in order]
wb.active = 2
wb.save(out)
print(f'OK : {len(rows)} clients, LISTING = {visibles} lignes visibles (moyenne annuel HT {moy:,.2f})')
