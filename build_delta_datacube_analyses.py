"""
Build Delta Datacube Analyses
==============================

Adds analysis sheets (formulas only, no hard-coded values) to the
260511_Delta_Datacube.xlsx workbook, referencing the `ResultatFinancement`
data sheet.

Usage
-----
    python build_delta_datacube_analyses.py [input_xlsx] [output_xlsx]

Defaults to:
    input  = 260511_Delta_Datacube.xlsx
    output = 260511_Delta_Datacube_with_Analyses.xlsx

All analysis sheets compute their KPIs from the data sheet via Excel
formulas (COUNTIFS, SUMIFS, AVERAGEIFS, SUMPRODUCT, LARGE, ...). When the
output is opened in Excel, formulas evaluate automatically.
"""

from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DATA_SHEET = "ResultatFinancement"

# ---------------------------------------------------------------------------
# Header detection
# ---------------------------------------------------------------------------

EXPECTED_HEADERS = [
    "Id", "CodePostal", "LibelleCommune", "ActivitePrincipaleEntreprise",
    "Secteur d'activité récolté dans Henrri", "forme juridique", "effectif",
    "SourceCreation", "SocieteFrontaleId", "DatePassageModeOfficiel",
    "DateCreation", "client actif", "TypeAbo", "TypeDuree",
    "DateAchatEnCours", "DateFinEnCours", "MRR en cours HT", "StatusOffre",
    "DatePremierAchat", "Nb utlisateurs du compte",
]

MONTH_COLS = [
    f"NbDocs_{y}_{m:02d}"
    for y in (2025, 2026)
    for m in range(1, 13)
]  # 24 month columns; only those present in data are used

REVENUE_COLS = ["CA_2025", "CA_2026"]


def find_header_row(ws) -> int:
    """Scan first 10 rows; return row index whose values match expected headers.

    Works on both read-only and normal worksheets.
    """
    for i, row in enumerate(ws.iter_rows(max_row=10, values_only=True), start=1):
        as_text = [str(v).strip() for v in row if v is not None]
        hits = sum(1 for h in EXPECTED_HEADERS if h in as_text)
        if hits >= 8:
            return i
    raise RuntimeError("Could not detect header row in ResultatFinancement.")


def build_column_map(ws, header_row: int) -> dict[str, str]:
    """Map header label -> column letter (e.g. 'Id' -> 'B').

    Works on both read-only and normal worksheets.
    """
    col_map: dict[str, str] = {}
    for i, row in enumerate(ws.iter_rows(min_row=header_row,
                                          max_row=header_row,
                                          values_only=False), start=1):
        for cell in row:
            if cell.value is None:
                continue
            col_map[str(cell.value).strip()] = get_column_letter(cell.column)
    return col_map


def collect_value_counts(input_path: Path, header_row: int,
                          cm: dict[str, str],
                          fields: list[str]) -> dict[str, list[tuple[str, int]]]:
    """Stream the data sheet once and count occurrences per requested field.

    Returns {field_label: [(value, count), ...]} sorted desc by count.
    """
    from openpyxl import load_workbook as _lw
    from openpyxl.utils import column_index_from_string
    print("Streaming data sheet to compute unique values ...")
    wb = _lw(input_path, read_only=True, data_only=True)
    ws = wb[DATA_SHEET]
    col_idx = {f: column_index_from_string(cm[f]) - 1
               for f in fields if f in cm}
    counts: dict[str, dict[str, int]] = {f: {} for f in fields if f in cm}
    n = 0
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i + 1 <= header_row:
            continue
        for f, c in col_idx.items():
            if c >= len(row):
                continue
            v = row[c]
            if v is None or v == "":
                continue
            key = str(v).strip()
            if not key:
                continue
            counts[f][key] = counts[f].get(key, 0) + 1
        n += 1
        if n % 50000 == 0:
            print(f"  scanned {n} rows ...", flush=True)
    wb.close()
    print(f"  total scanned: {n} rows.")
    return {
        f: sorted(d.items(), key=lambda x: -x[1])
        for f, d in counts.items()
    }


# ---------------------------------------------------------------------------
# Styling helpers
# ---------------------------------------------------------------------------

TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
TITLE_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="2E75B6")
LABEL_FONT = Font(name="Calibri", size=11, bold=True)
TOTAL_FONT = Font(name="Calibri", size=11, bold=True)
TOTAL_FILL = PatternFill("solid", fgColor="D9E1F2")
BORDER = Border(*(Side(style="thin", color="BFBFBF"),) * 4)

CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")


def write_title(ws, text: str, span: int = 8) -> None:
    ws["A1"] = text
    ws["A1"].font = TITLE_FONT
    ws["A1"].fill = TITLE_FILL
    ws["A1"].alignment = LEFT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    ws.row_dimensions[1].height = 24


def write_header_row(ws, row: int, headers: list[str]) -> None:
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = CENTER
        c.border = BORDER


def apply_number_formats(ws, start_row: int, end_row: int, formats: dict[int, str]):
    """formats: {column_index: number_format}."""
    for col, fmt in formats.items():
        for r in range(start_row, end_row + 1):
            ws.cell(row=r, column=col).number_format = fmt


# ---------------------------------------------------------------------------
# Sheet builders
# ---------------------------------------------------------------------------

def col_ref(letter: str) -> str:
    """Return whole-column absolute reference into the data sheet."""
    return f"'{DATA_SHEET}'!${letter}:${letter}"


def add_kpi_dashboard(wb, cm: dict[str, str], months_present: list[str]) -> None:
    ws = wb.create_sheet("01_KPIs", 0 if DATA_SHEET == wb.sheetnames[0] else 1)
    write_title(ws, "Delta — Dashboard KPIs", span=4)

    rows = []
    rows.append(("Volumétrie clients", None))
    rows.append(("Nombre total de clients (lignes)",
                 f"=COUNTA({col_ref(cm['Id'])})-1"))
    rows.append(("Dont 'client actif'",
                 f'=COUNTIF({col_ref(cm["client actif"])},"client actif")'))
    rows.append(("Dont 'ancien client payant'",
                 f'=COUNTIF({col_ref(cm["client actif"])},"ancien client payant")'))
    rows.append(("Clients avec MRR > 0",
                 f'=COUNTIF({col_ref(cm["MRR en cours HT"])},">0")'))
    rows.append(("Clients avec StatusOffre = 'En cours'",
                 f'=COUNTIF({col_ref(cm["StatusOffre"])},"En cours")'))

    rows.append(("", None))
    rows.append(("MRR / ARR", None))
    rows.append(("MRR total (€ HT)",
                 f'=SUM({col_ref(cm["MRR en cours HT"])})'))
    rows.append(("ARR total (€ HT, MRR×12)",
                 f'=SUM({col_ref(cm["MRR en cours HT"])})*12'))
    rows.append(("ARPU mensuel (MRR / clients payants)",
                 f'=IFERROR(SUM({col_ref(cm["MRR en cours HT"])})'
                 f'/COUNTIF({col_ref(cm["MRR en cours HT"])},">0"),0)'))
    rows.append(("MRR moyen tous clients confondus",
                 f'=AVERAGE({col_ref(cm["MRR en cours HT"])})'))

    rows.append(("", None))
    rows.append(("Chiffre d'affaires", None))
    rows.append(("CA 2025 total",
                 f"=SUM({col_ref(cm['CA_2025'])})"))
    rows.append(("CA 2026 (YTD)",
                 f"=SUM({col_ref(cm['CA_2026'])})"))
    rows.append(("Clients avec CA 2025 > 0",
                 f'=COUNTIF({col_ref(cm["CA_2025"])},">0")'))
    rows.append(("Clients avec CA 2026 > 0",
                 f'=COUNTIF({col_ref(cm["CA_2026"])},">0")'))
    rows.append(("Croissance CA (2026/2025-1)",
                 f"=IFERROR(SUM({col_ref(cm['CA_2026'])})"
                 f"/SUM({col_ref(cm['CA_2025'])})-1,0)"))

    rows.append(("", None))
    rows.append(("Engagement (NbDocs)", None))
    sum_2025 = "+".join(f"SUM({col_ref(cm[c])})" for c in months_present if "_2025_" in c)
    sum_2026 = "+".join(f"SUM({col_ref(cm[c])})" for c in months_present if "_2026_" in c)
    rows.append(("Documents émis 2025 (total)", f"={sum_2025}" if sum_2025 else "=0"))
    rows.append(("Documents émis 2026 (YTD)", f"={sum_2026}" if sum_2026 else "=0"))

    rows.append(("", None))
    rows.append(("Utilisateurs", None))
    rows.append(("Total utilisateurs (somme)",
                 f"=SUM({col_ref(cm['Nb utlisateurs du compte'])})"))
    rows.append(("Utilisateurs moyens / compte",
                 f"=AVERAGE({col_ref(cm['Nb utlisateurs du compte'])})"))

    # Write rows starting at row 3
    r = 3
    for label, formula in rows:
        c = ws.cell(row=r, column=1, value=label)
        if formula is None and label and not label.startswith(("Volumétrie",
                                                                "MRR", "Chiffre",
                                                                "Engagement", "Utilisateurs")):
            # blank separator
            pass
        if formula is None:
            # section header style
            if label:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="2E75B6")
                ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        else:
            ws.cell(row=r, column=2, value=formula)
            c.font = LABEL_FONT
            # number format
            fmt = "#,##0"
            if "€" in label or "MRR" in label or "ARR" in label or "ARPU" in label or "CA" in label:
                fmt = "#,##0 €"
            if "Croissance" in label or "%" in label:
                fmt = "0.0%"
            ws.cell(row=r, column=2).number_format = fmt
        r += 1

    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 22


def add_subscription_breakdown(wb, cm: dict[str, str],
                                uniques: dict[str, list[tuple[str, int]]]) -> None:
    ws = wb.create_sheet("02_MRR_Subscriptions")
    write_title(ws, "Souscriptions — décomposition", span=6)

    type_abo_col = cm["TypeAbo"]
    type_duree_col = cm["TypeDuree"]
    status_col = cm["StatusOffre"]
    mrr_col = cm["MRR en cours HT"]
    ca25_col = cm["CA_2025"]
    ca26_col = cm["CA_2026"]

    # --- TypeAbo block (static values pre-computed) ---
    ws["A3"] = "Par TypeAbo"
    ws["A3"].font = LABEL_FONT
    for i, h in enumerate(["TypeAbo", "# clients", "MRR (€)", "ARR (€)",
                            "ARPU (€)", "% du MRR"], start=1):
        c = ws.cell(row=4, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = CENTER
    type_abo_values = [v for v, _ in uniques.get("TypeAbo", [])]
    last_typeabo = 4
    for i, val in enumerate(type_abo_values):
        r = 5 + i
        last_typeabo = r
        ws.cell(row=r, column=1, value=val)
        ws.cell(row=r, column=2,
                value=f'=COUNTIF({col_ref(type_abo_col)},$A{r})')
        ws.cell(row=r, column=3,
                value=f'=SUMIF({col_ref(type_abo_col)},$A{r},'
                      f'{col_ref(mrr_col)})')
        ws.cell(row=r, column=4, value=f'=C{r}*12')
        ws.cell(row=r, column=5,
                value=f'=IFERROR(SUMIFS({col_ref(mrr_col)},'
                      f'{col_ref(type_abo_col)},$A{r},'
                      f'{col_ref(mrr_col)},">0")/'
                      f'COUNTIFS({col_ref(type_abo_col)},$A{r},'
                      f'{col_ref(mrr_col)},">0"),0)')
        ws.cell(row=r, column=6,
                value=f'=IFERROR(C{r}/SUM({col_ref(mrr_col)}),0)')
    if type_abo_values:
        apply_number_formats(ws, 5, last_typeabo,
                             {2: "#,##0", 3: "#,##0 €",
                              4: "#,##0 €", 5: "#,##0 €", 6: "0.0%"})

    # --- TypeDuree block (static values) ---
    ws["H3"] = "Par TypeDuree"
    ws["H3"].font = LABEL_FONT
    for i, h in enumerate(["TypeDuree", "# clients", "MRR (€)", "ARR (€)"], start=8):
        c = ws.cell(row=4, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = CENTER
    type_duree_values = [v for v, _ in uniques.get("TypeDuree", [])]
    last_duree = 4
    for i, val in enumerate(type_duree_values):
        r = 5 + i
        last_duree = r
        ws.cell(row=r, column=8, value=val)
        ws.cell(row=r, column=9,
                value=f'=COUNTIF({col_ref(type_duree_col)},$H{r})')
        ws.cell(row=r, column=10,
                value=f'=SUMIF({col_ref(type_duree_col)},$H{r},'
                      f'{col_ref(mrr_col)})')
        ws.cell(row=r, column=11, value=f'=J{r}*12')
    if type_duree_values:
        apply_number_formats(ws, 5, last_duree,
                             {9: "#,##0", 10: "#,##0 €", 11: "#,##0 €"})

    # --- StatusOffre block (static values) ---
    status_start = max(last_typeabo, last_duree) + 3
    ws.cell(row=status_start - 1, column=1,
            value="Par StatusOffre").font = LABEL_FONT
    for i, h in enumerate(["StatusOffre", "# clients", "MRR (€)",
                            "ARR (€)", "CA 2025 (€)", "CA 2026 (€)"], start=1):
        c = ws.cell(row=status_start, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = CENTER
    status_values = [v for v, _ in uniques.get("StatusOffre", [])]
    for i, val in enumerate(status_values):
        r = status_start + 1 + i
        ws.cell(row=r, column=1, value=val)
        ws.cell(row=r, column=2,
                value=f'=COUNTIF({col_ref(status_col)},$A{r})')
        ws.cell(row=r, column=3,
                value=f'=SUMIF({col_ref(status_col)},$A{r},'
                      f'{col_ref(mrr_col)})')
        ws.cell(row=r, column=4, value=f'=C{r}*12')
        ws.cell(row=r, column=5,
                value=f'=SUMIF({col_ref(status_col)},$A{r},'
                      f'{col_ref(ca25_col)})')
        ws.cell(row=r, column=6,
                value=f'=SUMIF({col_ref(status_col)},$A{r},'
                      f'{col_ref(ca26_col)})')
    if status_values:
        apply_number_formats(ws, status_start + 1,
                             status_start + len(status_values),
                             {2: "#,##0", 3: "#,##0 €",
                              4: "#,##0 €", 5: "#,##0 €", 6: "#,##0 €"})

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["H"].width = 16
    for col in ["B", "C", "D", "E", "F", "I", "J", "K"]:
        ws.column_dimensions[col].width = 16


def add_cohort_analysis(wb, cm: dict[str, str], months_present: list[str]) -> None:
    ws = wb.create_sheet("03_Cohorts")
    write_title(ws, "Cohortes — par année d'inscription (DateCreation)", span=8)

    date_col = cm["DateCreation"]
    mrr_col = cm["MRR en cours HT"]
    ca25_col = cm["CA_2025"]
    ca26_col = cm["CA_2026"]
    actif_col = cm["client actif"]

    # Cohort years to show
    years = list(range(2015, 2027))
    headers = ["Année cohorte", "# clients", "% clients actifs",
               "MRR total (€)", "MRR moyen (€)",
               "CA 2025 (€)", "CA 2026 (€)",
               "CA moyen 2025 (€)"]
    write_header_row(ws, 3, headers)

    # Use SUMPRODUCT with LEFT()=year as text matching, since DateCreation is text
    for i, y in enumerate(years):
        r = 4 + i
        ws.cell(row=r, column=1, value=y)
        year_match = (
            f'(LEFT({col_ref(date_col)},4)="{y}")'
        )
        ws.cell(row=r, column=2,
                value=f"=SUMPRODUCT(--{year_match})")
        # % actifs : (client actif = "client actif" AND cohort year)
        ws.cell(row=r, column=3,
                value=f'=IFERROR(SUMPRODUCT({year_match}*'
                      f'({col_ref(actif_col)}="client actif"))/B{r},0)')
        ws.cell(row=r, column=4,
                value=f'=SUMPRODUCT({year_match}*{col_ref(mrr_col)})')
        ws.cell(row=r, column=5,
                value=f"=IFERROR(D{r}/B{r},0)")
        ws.cell(row=r, column=6,
                value=f'=SUMPRODUCT({year_match}*{col_ref(ca25_col)})')
        ws.cell(row=r, column=7,
                value=f'=SUMPRODUCT({year_match}*{col_ref(ca26_col)})')
        ws.cell(row=r, column=8,
                value=f"=IFERROR(F{r}/B{r},0)")

    last_r = 4 + len(years) - 1
    apply_number_formats(ws, 4, last_r,
                         {2: "#,##0", 3: "0.0%",
                          4: "#,##0 €", 5: "#,##0 €",
                          6: "#,##0 €", 7: "#,##0 €", 8: "#,##0 €"})

    # Total row
    tot_r = last_r + 1
    ws.cell(row=tot_r, column=1, value="TOTAL").font = TOTAL_FONT
    for col in (2, 4, 6, 7):
        letter = get_column_letter(col)
        ws.cell(row=tot_r, column=col,
                value=f"=SUM({letter}4:{letter}{last_r})")
    ws.cell(row=tot_r, column=3,
            value=f"=IFERROR(SUMPRODUCT(C4:C{last_r}*B4:B{last_r})/B{tot_r},0)")
    ws.cell(row=tot_r, column=5,
            value=f"=IFERROR(D{tot_r}/B{tot_r},0)")
    ws.cell(row=tot_r, column=8,
            value=f"=IFERROR(F{tot_r}/B{tot_r},0)")
    for col in range(1, 9):
        ws.cell(row=tot_r, column=col).fill = TOTAL_FILL
        ws.cell(row=tot_r, column=col).font = TOTAL_FONT
    apply_number_formats(ws, tot_r, tot_r,
                         {2: "#,##0", 3: "0.0%",
                          4: "#,##0 €", 5: "#,##0 €",
                          6: "#,##0 €", 7: "#,##0 €", 8: "#,##0 €"})

    # Cohort × Month engagement matrix
    ws.cell(row=tot_r + 3, column=1,
            value="Documents émis par cohorte × mois (somme des NbDocs)").font = LABEL_FONT
    matrix_header_row = tot_r + 4
    ws.cell(row=matrix_header_row, column=1, value="Cohorte \\ Mois")
    for j, mc in enumerate(months_present, start=2):
        ws.cell(row=matrix_header_row, column=j, value=mc.replace("NbDocs_", ""))
    for i in range(matrix_header_row, matrix_header_row + 1):
        for c in ws[i]:
            c.font = HEADER_FONT
            c.fill = HEADER_FILL
            c.alignment = CENTER

    for i, y in enumerate(years):
        r = matrix_header_row + 1 + i
        ws.cell(row=r, column=1, value=y)
        year_match = f'(LEFT({col_ref(date_col)},4)="{y}")'
        for j, mc in enumerate(months_present, start=2):
            ws.cell(row=r, column=j,
                    value=f"=SUMPRODUCT({year_match}*{col_ref(cm[mc])})")
            ws.cell(row=r, column=j).number_format = "#,##0"

    ws.column_dimensions["A"].width = 18
    for col_idx in range(2, 9 + len(months_present)):
        ws.column_dimensions[get_column_letter(col_idx)].width = 14


def add_engagement_monthly(wb, cm: dict[str, str], months_present: list[str]) -> None:
    ws = wb.create_sheet("04_Engagement_Monthly")
    write_title(ws, "Engagement mensuel — NbDocs", span=7)

    headers = ["Mois", "Total docs", "# clients ≥1 doc", "# clients ≥10 docs",
               "# clients ≥50 docs", "Docs moyens / client actif",
               "Docs max (1 client)"]
    write_header_row(ws, 3, headers)

    for i, mc in enumerate(months_present):
        r = 4 + i
        c_letter = cm[mc]
        ws.cell(row=r, column=1, value=mc.replace("NbDocs_", ""))
        ws.cell(row=r, column=2, value=f"=SUM({col_ref(c_letter)})")
        ws.cell(row=r, column=3, value=f'=COUNTIF({col_ref(c_letter)},">0")')
        ws.cell(row=r, column=4, value=f'=COUNTIF({col_ref(c_letter)},">=10")')
        ws.cell(row=r, column=5, value=f'=COUNTIF({col_ref(c_letter)},">=50")')
        ws.cell(row=r, column=6, value=f"=IFERROR(B{r}/C{r},0)")
        ws.cell(row=r, column=7, value=f"=MAX({col_ref(c_letter)})")

    end_r = 3 + len(months_present)
    apply_number_formats(ws, 4, end_r,
                         {2: "#,##0", 3: "#,##0", 4: "#,##0",
                          5: "#,##0", 6: "#,##0.0", 7: "#,##0"})

    ws.column_dimensions["A"].width = 14
    for col in "BCDEFG":
        ws.column_dimensions[col].width = 20


def add_sector_segmentation(wb, cm: dict[str, str],
                              uniques: dict[str, list[tuple[str, int]]]) -> None:
    ws = wb.create_sheet("05_Segment_Sector")
    write_title(ws, "Segmentation — secteur, forme juridique, effectif", span=6)

    sector_col = cm["ActivitePrincipaleEntreprise"]
    secteur_henrri_col = cm["Secteur d'activité récolté dans Henrri"]
    forme_col = cm["forme juridique"]
    eff_col = cm["effectif"]
    mrr_col = cm["MRR en cours HT"]
    ca25_col = cm["CA_2025"]
    ca26_col = cm["CA_2026"]

    # NAF top 20 by count — static values
    ws["A3"] = "Top 20 codes NAF par # clients"
    ws["A3"].font = LABEL_FONT
    for i, h in enumerate(["Code NAF", "# clients", "MRR (€)",
                            "CA 2025 (€)", "CA 2026 (€)"], start=1):
        c = ws.cell(row=4, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = CENTER
    naf_top = [v for v, _ in uniques.get("ActivitePrincipaleEntreprise", [])[:20]]
    for i, val in enumerate(naf_top):
        r = 5 + i
        ws.cell(row=r, column=1, value=val)
        ws.cell(row=r, column=2,
                value=f'=COUNTIF({col_ref(sector_col)},$A{r})')
        ws.cell(row=r, column=3,
                value=f'=SUMIF({col_ref(sector_col)},$A{r},'
                      f'{col_ref(mrr_col)})')
        ws.cell(row=r, column=4,
                value=f'=SUMIF({col_ref(sector_col)},$A{r},'
                      f'{col_ref(ca25_col)})')
        ws.cell(row=r, column=5,
                value=f'=SUMIF({col_ref(sector_col)},$A{r},'
                      f'{col_ref(ca26_col)})')
    if naf_top:
        apply_number_formats(ws, 5, 4 + len(naf_top),
                             {2: "#,##0", 3: "#,##0 €",
                              4: "#,##0 €", 5: "#,##0 €"})

    # Secteur Henrri top 20 — static values
    ws["G3"] = "Top 20 Secteurs Henrri par # clients"
    ws["G3"].font = LABEL_FONT
    for i, h in enumerate(["Secteur", "# clients", "MRR (€)", "CA 2025 (€)"],
                          start=7):
        c = ws.cell(row=4, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = CENTER
    henrri_top = [v for v, _ in uniques.get(
        "Secteur d'activité récolté dans Henrri", [])[:20]]
    for i, val in enumerate(henrri_top):
        r = 5 + i
        ws.cell(row=r, column=7, value=val)
        ws.cell(row=r, column=8,
                value=f'=COUNTIF({col_ref(secteur_henrri_col)},$G{r})')
        ws.cell(row=r, column=9,
                value=f'=SUMIF({col_ref(secteur_henrri_col)},$G{r},'
                      f'{col_ref(mrr_col)})')
        ws.cell(row=r, column=10,
                value=f'=SUMIF({col_ref(secteur_henrri_col)},$G{r},'
                      f'{col_ref(ca25_col)})')
    if henrri_top:
        apply_number_formats(ws, 5, 4 + len(henrri_top),
                             {8: "#,##0", 9: "#,##0 €", 10: "#,##0 €"})

    # Forme juridique — all values, sorted by count
    forme_start = max(4 + len(naf_top), 4 + len(henrri_top)) + 4
    ws.cell(row=forme_start - 1, column=1,
            value="Par forme juridique").font = LABEL_FONT
    for i, h in enumerate(["Forme juridique", "# clients", "MRR (€)",
                            "CA 2025 (€)", "% MRR"], start=1):
        c = ws.cell(row=forme_start, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = CENTER
    forme_values = [v for v, _ in uniques.get("forme juridique", [])]
    for i, val in enumerate(forme_values):
        r = forme_start + 1 + i
        ws.cell(row=r, column=1, value=val)
        ws.cell(row=r, column=2,
                value=f'=COUNTIF({col_ref(forme_col)},$A{r})')
        ws.cell(row=r, column=3,
                value=f'=SUMIF({col_ref(forme_col)},$A{r},'
                      f'{col_ref(mrr_col)})')
        ws.cell(row=r, column=4,
                value=f'=SUMIF({col_ref(forme_col)},$A{r},'
                      f'{col_ref(ca25_col)})')
        ws.cell(row=r, column=5,
                value=f'=IFERROR(C{r}/SUM({col_ref(mrr_col)}),0)')
    if forme_values:
        apply_number_formats(ws, forme_start + 1,
                             forme_start + len(forme_values),
                             {2: "#,##0", 3: "#,##0 €",
                              4: "#,##0 €", 5: "0.0%"})

    # Effectif tranches — place to the right of the forme juridique block
    eff_header_row = forme_start
    eff_data_start = forme_start + 1
    ws.cell(row=eff_header_row - 1, column=7,
            value="Par tranche d'effectif").font = LABEL_FONT
    for i, h in enumerate(["Tranche effectif", "# clients", "MRR (€)",
                            "CA 2025 (€)"], start=7):
        c = ws.cell(row=eff_header_row, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = CENTER

    eff_bands = [
        ("0 - vide", '""'),
        ("1", '"1"'),
        ("2-5", None),
        ("6-10", None),
        ("11-50", None),
        ("51+", None),
    ]
    band_ranges = {
        "2-5": (2, 5),
        "6-10": (6, 10),
        "11-50": (11, 50),
        "51+": (51, 10**9),
    }

    r = eff_data_start
    for name, _ in eff_bands:
        ws.cell(row=r, column=7, value=name)
        if name in band_ranges:
            lo, hi = band_ranges[name]
            ws.cell(row=r, column=8,
                    value=f'=COUNTIFS({col_ref(eff_col)},">={lo}",'
                          f'{col_ref(eff_col)},"<={hi}")')
            ws.cell(row=r, column=9,
                    value=f'=SUMIFS({col_ref(mrr_col)},'
                          f'{col_ref(eff_col)},">={lo}",'
                          f'{col_ref(eff_col)},"<={hi}")')
            ws.cell(row=r, column=10,
                    value=f'=SUMIFS({col_ref(ca25_col)},'
                          f'{col_ref(eff_col)},">={lo}",'
                          f'{col_ref(eff_col)},"<={hi}")')
        elif name == "0 - vide":
            ws.cell(row=r, column=8,
                    value=f'=COUNTBLANK({col_ref(eff_col)})-1')
            ws.cell(row=r, column=9,
                    value=f'=SUMIFS({col_ref(mrr_col)},'
                          f'{col_ref(eff_col)},"")')
            ws.cell(row=r, column=10,
                    value=f'=SUMIFS({col_ref(ca25_col)},'
                          f'{col_ref(eff_col)},"")')
        elif name == "1":
            ws.cell(row=r, column=8,
                    value=f'=COUNTIF({col_ref(eff_col)},1)')
            ws.cell(row=r, column=9,
                    value=f'=SUMIF({col_ref(eff_col)},1,{col_ref(mrr_col)})')
            ws.cell(row=r, column=10,
                    value=f'=SUMIF({col_ref(eff_col)},1,{col_ref(ca25_col)})')
        r += 1
    apply_number_formats(ws, eff_data_start, r - 1,
                         {8: "#,##0", 9: "#,##0 €", 10: "#,##0 €"})

    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["G"].width = 18
    for col in "BCDEFHIJ":
        ws.column_dimensions[col].width = 16


def add_geo_segmentation(wb, cm: dict[str, str]) -> None:
    ws = wb.create_sheet("06_Segment_Geo")
    write_title(ws, "Segmentation géographique — par département", span=6)

    cp_col = cm["CodePostal"]
    mrr_col = cm["MRR en cours HT"]
    ca25_col = cm["CA_2025"]
    ca26_col = cm["CA_2026"]

    headers = ["Département (2 prem. chiffres)", "# clients",
               "MRR (€)", "CA 2025 (€)", "CA 2026 (€)", "% du CA 2025"]
    write_header_row(ws, 3, headers)

    # 01 .. 95 (skip 00, include 2A/2B as text via TEXT format), 97x DOM
    departments = [f"{i:02d}" for i in range(1, 96)] + \
                  ["971", "972", "973", "974", "976"]

    for i, dep in enumerate(departments):
        r = 4 + i
        ws.cell(row=r, column=1, value=dep)
        # Use LEFT() match on CodePostal text representation
        # Match on length: dep len 2 -> LEFT(CP,2)==dep ; dep len 3 -> LEFT(CP,3)==dep
        if len(dep) == 2:
            match = f'(LEFT(TEXT({col_ref(cp_col)},"00000"),2)="{dep}")'
        else:
            match = f'(LEFT(TEXT({col_ref(cp_col)},"00000"),3)="{dep}")'
        # Exclude blanks: CodePostal <> ""
        nonblank = f'({col_ref(cp_col)}<>"")'
        ws.cell(row=r, column=2,
                value=f'=SUMPRODUCT({match}*{nonblank})')
        ws.cell(row=r, column=3,
                value=f'=SUMPRODUCT({match}*{nonblank}*{col_ref(mrr_col)})')
        ws.cell(row=r, column=4,
                value=f'=SUMPRODUCT({match}*{nonblank}*{col_ref(ca25_col)})')
        ws.cell(row=r, column=5,
                value=f'=SUMPRODUCT({match}*{nonblank}*{col_ref(ca26_col)})')
        ws.cell(row=r, column=6,
                value=f'=IFERROR(D{r}/SUM({col_ref(ca25_col)}),0)')

    end_r = 3 + len(departments)
    apply_number_formats(ws, 4, end_r,
                         {2: "#,##0", 3: "#,##0 €", 4: "#,##0 €",
                          5: "#,##0 €", 6: "0.0%"})

    ws.column_dimensions["A"].width = 28
    for col in "BCDEF":
        ws.column_dimensions[col].width = 18


def add_acquisition(wb, cm: dict[str, str],
                    uniques: dict[str, list[tuple[str, int]]]) -> None:
    ws = wb.create_sheet("07_Acquisition")
    write_title(ws, "Acquisition — par SourceCreation et cohorte", span=6)

    source_col = cm["SourceCreation"]
    date_creation_col = cm["DateCreation"]
    date_paid_col = cm["DatePremierAchat"]
    mrr_col = cm["MRR en cours HT"]
    ca25_col = cm["CA_2025"]

    ws["A3"] = "Par SourceCreation (top 50)"
    ws["A3"].font = LABEL_FONT
    for i, h in enumerate(["SourceCreation", "# clients",
                            "# clients payants (MRR>0)",
                            "Taux conversion payant",
                            "MRR (€)", "CA 2025 (€)"], start=1):
        c = ws.cell(row=4, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = CENTER

    source_top = [v for v, _ in uniques.get("SourceCreation", [])[:50]]
    for i, val in enumerate(source_top):
        r = 5 + i
        ws.cell(row=r, column=1, value=val)
        ws.cell(row=r, column=2,
                value=f'=COUNTIF({col_ref(source_col)},$A{r})')
        ws.cell(row=r, column=3,
                value=f'=COUNTIFS({col_ref(source_col)},$A{r},'
                      f'{col_ref(mrr_col)},">0")')
        ws.cell(row=r, column=4, value=f"=IFERROR(C{r}/B{r},0)")
        ws.cell(row=r, column=5,
                value=f'=SUMIF({col_ref(source_col)},$A{r},{col_ref(mrr_col)})')
        ws.cell(row=r, column=6,
                value=f'=SUMIF({col_ref(source_col)},$A{r},{col_ref(ca25_col)})')
    if source_top:
        apply_number_formats(ws, 5, 4 + len(source_top),
                             {2: "#,##0", 3: "#,##0", 4: "0.0%",
                              5: "#,##0 €", 6: "#,##0 €"})

    # Conversion rate to paying by cohort year
    cohort_start = 4 + len(source_top) + 4
    ws.cell(row=cohort_start - 1, column=1,
            value="Conversion création → premier achat — par année de cohorte"
            ).font = LABEL_FONT
    for i, h in enumerate(["Année cohorte", "# créés", "# avec DatePremierAchat",
                            "Taux conversion"], start=1):
        c = ws.cell(row=cohort_start, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = CENTER

    years = list(range(2015, 2027))
    for i, y in enumerate(years):
        r = cohort_start + 1 + i
        ws.cell(row=r, column=1, value=y)
        year_match = f'(LEFT({col_ref(date_creation_col)},4)="{y}")'
        ws.cell(row=r, column=2,
                value=f"=SUMPRODUCT(--{year_match})")
        ws.cell(row=r, column=3,
                value=f'=SUMPRODUCT({year_match}*({col_ref(date_paid_col)}<>""))')
        ws.cell(row=r, column=4, value=f"=IFERROR(C{r}/B{r},0)")
    end_r = cohort_start + len(years)
    apply_number_formats(ws, cohort_start + 1, end_r,
                         {2: "#,##0", 3: "#,##0", 4: "0.0%"})

    ws.column_dimensions["A"].width = 36
    for col in "BCDEF":
        ws.column_dimensions[col].width = 22


def add_pareto(wb, cm: dict[str, str]) -> None:
    ws = wb.create_sheet("08_Pareto")
    write_title(ws, "Concentration — Pareto top clients (CA 2025)", span=6)

    id_col = cm["Id"]
    ville_col = cm["LibelleCommune"]
    naf_col = cm["ActivitePrincipaleEntreprise"]
    mrr_col = cm["MRR en cours HT"]
    ca25_col = cm["CA_2025"]
    ca26_col = cm["CA_2026"]

    # Summary block
    ws["A3"] = "Synthèse concentration CA 2025"
    ws["A3"].font = LABEL_FONT
    summary = [
        ("CA 2025 total", f"=SUM({col_ref(ca25_col)})", "#,##0 €"),
        ("Top 10 clients (CA 2025)",
         f"=SUMPRODUCT(LARGE({col_ref(ca25_col)},ROW(INDIRECT(\"1:10\"))))",
         "#,##0 €"),
        ("% top 10",
         f"=IFERROR(B5/B4,0)", "0.0%"),
        ("Top 50 clients (CA 2025)",
         f"=SUMPRODUCT(LARGE({col_ref(ca25_col)},ROW(INDIRECT(\"1:50\"))))",
         "#,##0 €"),
        ("% top 50",
         f"=IFERROR(B7/B4,0)", "0.0%"),
        ("Top 100 clients (CA 2025)",
         f"=SUMPRODUCT(LARGE({col_ref(ca25_col)},ROW(INDIRECT(\"1:100\"))))",
         "#,##0 €"),
        ("% top 100",
         f"=IFERROR(B9/B4,0)", "0.0%"),
        ("Top 500 clients (CA 2025)",
         f"=SUMPRODUCT(LARGE({col_ref(ca25_col)},ROW(INDIRECT(\"1:500\"))))",
         "#,##0 €"),
        ("% top 500",
         f"=IFERROR(B11/B4,0)", "0.0%"),
    ]
    for i, (lbl, formula, fmt) in enumerate(summary):
        r = 4 + i
        ws.cell(row=r, column=1, value=lbl).font = LABEL_FONT
        ws.cell(row=r, column=2, value=formula).number_format = fmt

    # Top 50 list with IDs
    ws["A15"] = "Top 50 clients par CA 2025"
    ws["A15"].font = LABEL_FONT
    write_header_row(ws, 16, ["Rang", "Id client", "Commune", "NAF",
                               "CA 2025 (€)", "CA 2026 (€)",
                               "MRR (€)", "% CA 2025 cumulé"])
    for i in range(1, 51):
        r = 16 + i
        ws.cell(row=r, column=1, value=i)
        # CA 2025 (use LARGE)
        ws.cell(row=r, column=5,
                value=f"=LARGE({col_ref(ca25_col)},{i})")
        # ID = INDEX/MATCH of LARGE rank — handle ties with MATCH on (Range=val, k-th)
        # Simpler approach: INDEX(Id, MATCH(LARGE(CA, i), CA, 0))
        # Ties: use a tiebreaker by offsetting via SMALL approach.
        ws.cell(row=r, column=2,
                value=f"=IFERROR(INDEX({col_ref(id_col)},"
                      f"MATCH(LARGE({col_ref(ca25_col)},{i}),"
                      f"{col_ref(ca25_col)},0)),\"\")")
        ws.cell(row=r, column=3,
                value=f"=IFERROR(INDEX({col_ref(ville_col)},"
                      f"MATCH(LARGE({col_ref(ca25_col)},{i}),"
                      f"{col_ref(ca25_col)},0)),\"\")")
        ws.cell(row=r, column=4,
                value=f"=IFERROR(INDEX({col_ref(naf_col)},"
                      f"MATCH(LARGE({col_ref(ca25_col)},{i}),"
                      f"{col_ref(ca25_col)},0)),\"\")")
        ws.cell(row=r, column=6,
                value=f"=IFERROR(INDEX({col_ref(ca26_col)},"
                      f"MATCH(LARGE({col_ref(ca25_col)},{i}),"
                      f"{col_ref(ca25_col)},0)),\"\")")
        ws.cell(row=r, column=7,
                value=f"=IFERROR(INDEX({col_ref(mrr_col)},"
                      f"MATCH(LARGE({col_ref(ca25_col)},{i}),"
                      f"{col_ref(ca25_col)},0)),\"\")")
        ws.cell(row=r, column=8,
                value=f"=IFERROR(SUM($E$17:$E{r})/$B$4,0)")

    apply_number_formats(ws, 17, 66,
                         {5: "#,##0 €", 6: "#,##0 €", 7: "#,##0 €", 8: "0.0%"})

    ws.column_dimensions["A"].width = 36
    for col in "BCDEFGH":
        ws.column_dimensions[col].width = 18


def add_churn(wb, cm: dict[str, str], months_present: list[str]) -> None:
    ws = wb.create_sheet("09_Churn_Activity")
    write_title(ws, "Churn & réactivation — basé sur NbDocs", span=6)

    # Build SUMPRODUCT formulas using all 2025 monthly columns and 2026 monthly columns
    cols_2025 = [cm[c] for c in months_present if "_2025_" in c]
    cols_2026 = [cm[c] for c in months_present if "_2026_" in c]

    def col_sum(col_letters: list[str]) -> str:
        # Build "(col1+col2+...) " expression for SUMPRODUCT row-wise sum
        return "(" + "+".join(col_ref(c) for c in col_letters) + ")"

    sum25 = col_sum(cols_2025) if cols_2025 else "0"
    sum26 = col_sum(cols_2026) if cols_2026 else "0"

    headers = ["Segment", "# clients", "CA 2025 (€)", "CA 2026 (€)", "MRR (€)"]
    write_header_row(ws, 3, headers)

    rows = [
        ("Actifs 2025 ET 2026 (docs>0 dans chacun)",
         f"=SUMPRODUCT(({sum25}>0)*({sum26}>0))"),
        ("Actifs 2025 SEULEMENT (churn présumé)",
         f"=SUMPRODUCT(({sum25}>0)*({sum26}=0))"),
        ("Actifs 2026 SEULEMENT (réactivation/nouveaux)",
         f"=SUMPRODUCT(({sum25}=0)*({sum26}>0))"),
        ("Inactifs (0 doc 2025 et 2026)",
         f"=SUMPRODUCT(({sum25}=0)*({sum26}=0))"),
    ]

    mrr_col = cm["MRR en cours HT"]
    ca25_col = cm["CA_2025"]
    ca26_col = cm["CA_2026"]

    conditions = [
        f"({sum25}>0)*({sum26}>0)",
        f"({sum25}>0)*({sum26}=0)",
        f"({sum25}=0)*({sum26}>0)",
        f"({sum25}=0)*({sum26}=0)",
    ]

    for i, ((label, count_formula), cond) in enumerate(zip(rows, conditions)):
        r = 4 + i
        ws.cell(row=r, column=1, value=label).font = LABEL_FONT
        ws.cell(row=r, column=2, value=count_formula)
        ws.cell(row=r, column=3,
                value=f"=SUMPRODUCT({cond}*{col_ref(ca25_col)})")
        ws.cell(row=r, column=4,
                value=f"=SUMPRODUCT({cond}*{col_ref(ca26_col)})")
        ws.cell(row=r, column=5,
                value=f"=SUMPRODUCT({cond}*{col_ref(mrr_col)})")

    apply_number_formats(ws, 4, 7,
                         {2: "#,##0", 3: "#,##0 €", 4: "#,##0 €", 5: "#,##0 €"})

    # Totals
    ws.cell(row=8, column=1, value="TOTAL").font = TOTAL_FONT
    for col in (2, 3, 4, 5):
        letter = get_column_letter(col)
        ws.cell(row=8, column=col, value=f"=SUM({letter}4:{letter}7)")
        ws.cell(row=8, column=col).fill = TOTAL_FILL
        ws.cell(row=8, column=col).font = TOTAL_FONT
    ws.cell(row=8, column=1).fill = TOTAL_FILL
    apply_number_formats(ws, 8, 8,
                         {2: "#,##0", 3: "#,##0 €", 4: "#,##0 €", 5: "#,##0 €"})

    # %
    ws.cell(row=10, column=1, value="% du total clients").font = LABEL_FONT
    for r in range(4, 8):
        ws.cell(row=r, column=6, value=f"=IFERROR(B{r}/$B$8,0)").number_format = "0.0%"
    ws.cell(row=3, column=6, value="% clients").font = HEADER_FONT
    ws.cell(row=3, column=6).fill = HEADER_FILL
    ws.cell(row=3, column=6).alignment = CENTER

    ws.column_dimensions["A"].width = 48
    for col in "BCDEF":
        ws.column_dimensions[col].width = 18


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    input_path = Path(sys.argv[1] if len(sys.argv) > 1
                      else "260511_Delta_Datacube.xlsx")
    output_path = Path(sys.argv[2] if len(sys.argv) > 2
                       else "260511_Delta_Datacube_with_Analyses.xlsx")

    if not input_path.exists():
        sys.exit(f"Input file not found: {input_path}")

    # Pass 1 (streaming, read-only): detect headers, columns, and collect uniques
    print(f"Inspecting {input_path} in read-only mode ...")
    wb_ro = load_workbook(input_path, read_only=True, data_only=True)
    if DATA_SHEET not in wb_ro.sheetnames:
        sys.exit(f"Sheet '{DATA_SHEET}' not found. Sheets: {wb_ro.sheetnames}")
    ws_data_ro = wb_ro[DATA_SHEET]
    header_row = find_header_row(ws_data_ro)
    print(f"Detected header row at row {header_row}")
    cm = build_column_map(ws_data_ro, header_row)
    print(f"Detected {len(cm)} columns")
    missing = [h for h in EXPECTED_HEADERS if h not in cm]
    if missing:
        print(f"WARNING: missing expected columns: {missing}")
    months_present = [m for m in MONTH_COLS if m in cm]
    revenues_present = [m for m in REVENUE_COLS if m in cm]
    print(f"Monthly columns present: {len(months_present)} "
          f"({months_present[0] if months_present else 'none'} ... "
          f"{months_present[-1] if months_present else 'none'})")
    print(f"Revenue columns present: {revenues_present}")
    wb_ro.close()

    fields_to_count = [
        "TypeAbo", "TypeDuree", "StatusOffre",
        "ActivitePrincipaleEntreprise",
        "Secteur d'activité récolté dans Henrri",
        "forme juridique", "SourceCreation",
    ]
    uniques = collect_value_counts(input_path, header_row, cm, fields_to_count)
    for f, lst in uniques.items():
        print(f"  {f}: {len(lst)} distinct values "
              f"(top: {lst[0] if lst else '-'})")

    # Pass 2: load workbook in normal mode and add the analysis sheets
    print(f"Loading {input_path} in normal mode (for editing) ...")
    wb = load_workbook(input_path)

    print("Building analysis sheets ...")
    add_kpi_dashboard(wb, cm, months_present)
    add_subscription_breakdown(wb, cm, uniques)
    add_cohort_analysis(wb, cm, months_present)
    add_engagement_monthly(wb, cm, months_present)
    add_sector_segmentation(wb, cm, uniques)
    add_geo_segmentation(wb, cm)
    add_acquisition(wb, cm, uniques)
    add_pareto(wb, cm)
    add_churn(wb, cm, months_present)

    print(f"Saving to {output_path} ...")
    wb.save(output_path)
    print("Done.")


if __name__ == "__main__":
    main()
