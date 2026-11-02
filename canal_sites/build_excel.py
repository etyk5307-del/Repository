"""Classeur de suivi des points de vente (3 feuilles) relié à la carte en ligne.

Usage : python canal_sites/build_excel.py   (après build_sites.py, qui produit dist/sites.pkl)
Feuilles : TABLEAU_DE_BORD (lien Netlify à coller, indicateurs et graphiques en formules),
SITES (registre complet avec liens carte / Google Maps / OpenStreetMap), QUALITE_INTEGRATION
(contrôle des coordonnées et procédure d'intégration du Mali et du Niger).
"""
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo

HERE = Path(__file__).resolve().parent
DIST = HERE / "dist"
OUT = DIST / "Sites_PVA_Burkina_Mali_Niger.xlsx"
MAXR = 5000  # lignes couvertes par les formules (place pour les sites du Mali et du Niger)
FONT, NAVY, BLUE, MUTED, LINE, PALE = "Arial", "121A26", "1F5FD1", "5A6577", "D8DEE7", "F4F6F9"
QC_COL = {"OK": "2E7D4F", "Corrigé": "C76A0A", "Corrigé, à confirmer": "C76A0A", "Coordonnées partagées": "B88A00",
          "Frontière, à vérifier": "B88A00", "Position générique": "C8102E", "Hors pays déclaré": "C8102E"}
thin = Side(style="thin", color=LINE)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def f(size=10, bold=False, color="121A26", italic=False, underline=None):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic, underline=underline)


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


def banner(ws, title, sub, width):
    ws.sheet_view.showGridLines = False
    for c in range(1, width + 1):
        for r in (1, 2):
            ws.cell(r, c).fill = fill(NAVY)
    ws["A1"], ws["A2"] = title, sub
    ws["A1"].font = Font(name=FONT, size=16, bold=True, color="FFFFFF")
    ws["A2"].font = Font(name=FONT, size=9, color="C9D3DC")
    for k in ("A1", "A2"):
        ws[k].alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height, ws.row_dimensions[2].height = 30, 18


def header(ws, row, labels, col=1, color=NAVY):
    for k, label in enumerate(labels):
        c = ws.cell(row, col + k, label)
        c.font, c.fill, c.border = Font(name=FONT, size=9, bold=True, color="FFFFFF"), fill(color), BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30


def cell(ws, r, c, v, bold=False, fmt=None, color="121A26", center=False, size=10):
    x = ws.cell(r, c, v)
    x.font, x.border = f(size, bold, color), BOX
    x.alignment = Alignment(horizontal="center" if center else None, vertical="center", wrap_text=True)
    if fmt:
        x.number_format = fmt
    return x


def section(ws, r, text, c=2, span=6):
    x = ws.cell(r, c, text)
    x.font = f(12, True, BLUE)
    for k in range(span):
        ws.cell(r, c + k).border = Border(bottom=Side(style="medium", color=BLUE))


COLS = [  # (en-tête, clé, largeur)
    ("N° distributeur", "id", 12), ("Nom", "nom", 30), ("Pays", "pays", 13), ("Région (géocodée)", "region", 15), ("Province / cercle / département", "province", 18),
    ("Région déclarée", "region_declaree", 14), ("Ville", "ville", 24), ("Adresse", "adresse", 34), ("Zone commerciale", "zone", 13), ("Secteur", "secteur", 18),
    ("Date d'activation", "activation", 12), ("Année", None, 8), ("Latitude", "lat", 11), ("Longitude", "lon", 11), ("Géolocalisation", "qc", 18), ("Détail du contrôle", "qc_detail", 40),
    ("Latitude d'origine", "lat_source", 11), ("Longitude d'origine", "lon_source", 11), ("Site voisin (km)", "voisin_km", 9), ("N° parent", "parent", 10),
    ("Téléphone", "telephone", 18), ("Email", "email", 26), ("Login principal", "login", 15), ("Mode de crédit", "credit", 9), ("Filiale", "filiale", 15),
    ("Carte", None, 10), ("Google Maps", None, 11), ("OpenStreetMap", None, 12),
]
L = {h: get_column_letter(i + 1) for i, (h, _, _) in enumerate(COLS)}


def sheet_sites(wb, s):
    ws = wb.create_sheet("SITES")
    header(ws, 1, [h for h, _, _ in COLS])
    for i, (_, _, w) in enumerate(COLS):
        ws.column_dimensions[get_column_letter(i + 1)].width = w
    for r, row in enumerate(s.itertuples(index=False), start=2):
        d = row._asdict()
        for i, (h, k, _) in enumerate(COLS):
            c = ws.cell(r, i + 1)
            if k == "activation":
                c.value = date.fromisoformat(d[k]) if d[k] else None
                c.number_format = "DD/MM/YYYY"
            elif k:
                v = d[k]
                c.value = None if (isinstance(v, float) and pd.isna(v)) else v
            c.font = f(9)
            c.border = Border(bottom=thin)
        la, lo, idc = f"{L['Latitude']}{r}", f"{L['Longitude']}{r}", f"{L['N° distributeur']}{r}"
        act = L["Date d'activation"]
        ws[f"{L['Année']}{r}"] = f'=IF({act}{r}="","",YEAR({act}{r}))'
        # coordonnées converties en texte à point décimal : un Excel en français écrirait « 12,39 » et casserait l'URL
        tl, tn = f'SUBSTITUTE({la}&"",",",".")', f'SUBSTITUTE({lo}&"",",",".")'
        ws[f"{L['Carte']}{r}"] = f'=HYPERLINK(URL_CARTE&"#site-"&{idc},"Carte ↗")'
        ws[f"{L['Google Maps']}{r}"] = f'=HYPERLINK("https://www.google.com/maps/search/?api=1&query="&{tl}&"%2C"&{tn},"Google ↗")'
        ws[f"{L['OpenStreetMap']}{r}"] = f'=HYPERLINK("https://www.openstreetmap.org/?mlat="&{tl}&"&mlon="&{tn}&"#map=17/"&{tl}&"/"&{tn},"OSM ↗")'
        for h in ("Carte", "Google Maps", "OpenStreetMap"):
            ws[f"{L[h]}{r}"].font = f(9, color="0563C1", underline="single")
        ws[f"{L['Année']}{r}"].font = f(9)
        for h in ("Latitude", "Longitude", "Latitude d'origine", "Longitude d'origine"):
            ws[f"{L[h]}{r}"].number_format = "0.000000"
    last = len(s) + 1
    t = Table(displayName="tblSites", ref=f"A1:{get_column_letter(len(COLS))}{last}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(t)
    ws.freeze_panes = "C2"
    q = f"{L['Géolocalisation']}2:{L['Géolocalisation']}{MAXR}"
    for k, col in QC_COL.items():
        ws.conditional_formatting.add(q, CellIsRule(operator="equal", formula=[f'"{k}"'], font=Font(name=FONT, bold=True, color=col)))
    lon0 = L["Longitude d'origine"]
    ws.conditional_formatting.add(f"{lon0}2:{lon0}{MAXR}", CellIsRule(operator="greaterThan", formula=["0"], fill=fill("FBE3D0")))
    ws["A1"].comment = Comment("Source : export des points de vente agréés. Régions et provinces « géocodées » : jointure spatiale des coordonnées corrigées. "
                               "Les colonnes de contact sont internes : elles ne sont pas publiées sur la carte en ligne.", "Cartographie")
    for nm, h in (("s_id", "N° distributeur"), ("s_pays", "Pays"), ("s_region", "Région (géocodée)"), ("s_prov", "Province / cercle / département"),
                  ("s_zone", "Zone commerciale"), ("s_annee", "Année"), ("s_qc", "Géolocalisation")):
        wb.defined_names[nm] = DefinedName(nm, attr_text=f"SITES!${L[h]}$2:${L[h]}${MAXR}")
    return last


def sheet_dash(wb, s):
    ws = wb.active
    ws.title = "TABLEAU_DE_BORD"
    banner(ws, "Points de vente agréés — Burkina Faso · Mali · Niger", f"Cartographie et suivi sécuritaire · généré le {date.today():%d/%m/%Y} · usage interne", 13)
    for c, w in zip("ABCDEFGHIJKLM", (2, 30, 14, 12, 3, 26, 12, 12, 3, 14, 12, 3, 3)):
        ws.column_dimensions[c].width = w
    section(ws, 4, "Accès à la carte en ligne", 2, 7)
    cell(ws, 5, 2, "Lien Netlify (à coller après déploiement)", bold=True)
    u = ws.cell(5, 3, "https://votre-site.netlify.app")
    u.font, u.fill, u.border = f(11, True, "0000FF"), fill("FFF2B3"), BOX
    ws.merge_cells("C5:H5")
    wb.defined_names["URL_CARTE"] = DefinedName("URL_CARTE", attr_text="TABLEAU_DE_BORD!$C$5")
    u.comment = Comment("Collez ici l'adresse donnée par Netlify (ex. https://pva-sahel.netlify.app). Tous les liens « Carte ↗ » du classeur s'y raccrochent.", "Cartographie")
    o = ws.cell(6, 3, '=HYPERLINK(URL_CARTE,"▶  Ouvrir la cartographie dynamique")')
    o.font = f(12, True, "0563C1", underline="single")
    ws.merge_cells("C6:H6")
    ws.cell(7, 3, "Chaque site de la feuille SITES a un lien « Carte ↗ » qui ouvre directement sa fiche sur la carte.").font = f(9, color=MUTED, italic=True)
    ws.merge_cells("C7:H7")

    section(ws, 9, "Indicateurs", 2, 11)
    kpis = [("Sites actifs", "=COUNTA(s_id)"), ("Burkina Faso", '=COUNTIFS(s_pays,"Burkina Faso")'), ("Mali", '=COUNTIFS(s_pays,"Mali")'), ("Niger", '=COUNTIFS(s_pays,"Niger")'),
            ("Positions fiables", '=COUNTIFS(s_qc,"OK")'), ("Corrigées ou à vérifier", '=COUNTA(s_id)-COUNTIFS(s_qc,"OK")')]
    for k, (lab, fml) in enumerate(kpis):
        r, c = 10 + (k // 3) * 3, [2, 6, 10][k % 3]
        a = ws.cell(r, c, lab.upper())
        a.font = Font(name=FONT, size=8, bold=True, color=MUTED)
        v = ws.cell(r + 1, c, fml)
        v.font = Font(name=FONT, size=22, bold=True, color=BLUE if k < 4 else ("2E7D4F" if k == 4 else "C76A0A"))
        v.number_format = "#,##0"
        for rr in (r, r + 1):
            ws.cell(rr, c).fill = fill(PALE)
            ws.cell(rr, c).border = Border(left=Side(style="thick", color=BLUE))
        ws.row_dimensions[r + 1].height = 32

    # Répartition par région (pays + région géocodée)
    regs = s.groupby(["pays", "region"]).size().reset_index().sort_values(["pays", 0], ascending=[True, False])
    r0 = 17
    section(ws, r0 - 1, "Sites par région", 2, 3)
    header(ws, r0, ["Région", "Pays", "Sites"], 2)
    for i, row in enumerate(regs.itertuples(index=False), start=r0 + 1):
        cell(ws, i, 2, row.region, bold=True)
        cell(ws, i, 3, row.pays)
        cell(ws, i, 4, f'=COUNTIFS(s_region,B{i},s_pays,C{i})', center=True, fmt="0")
    r1 = r0 + len(regs)
    for c in (2, 3):
        ws.cell(r1 + 1, c).border = BOX
    cell(ws, r1 + 1, 2, "Total", bold=True).fill = fill(PALE)
    cell(ws, r1 + 1, 4, f"=SUM(D{r0 + 1}:D{r1})", bold=True, center=True).fill = fill(PALE)
    ws.conditional_formatting.add(f"D{r0 + 1}:D{r1}", DataBarRule(start_type="num", start_value=0, end_type="max", color=BLUE, showValue=True))

    # Zones commerciales
    section(ws, r0 - 1, "Sites par zone commerciale", 6, 3)
    header(ws, r0, ["Zone", "Sites", "Part"], 6)
    zones = sorted(s.zone.replace("", "Non renseignée").unique())
    for i, z in enumerate(zones, start=r0 + 1):
        cell(ws, i, 6, z, bold=True)
        cell(ws, i, 7, f'=COUNTIFS(s_zone,F{i})', center=True, fmt="0")
        cell(ws, i, 8, f"=IF(SUM($G${r0 + 1}:$G${r0 + len(zones)})=0,0,G{i}/SUM($G${r0 + 1}:$G${r0 + len(zones)}))", center=True, fmt="0%")
    # Activations par année
    section(ws, r0 - 1, "Activations par année", 10, 2)
    header(ws, r0, ["Année", "Sites"], 10)
    years = sorted(int(y) for y in s.annee.dropna().unique())
    for i, y in enumerate(years, start=r0 + 1):
        cell(ws, i, 10, y, bold=True, center=True)
        cell(ws, i, 11, f"=COUNTIFS(s_annee,J{i})", center=True, fmt="0")

    ch = BarChart()
    ch.type, ch.style, ch.title = "bar", 2, "Sites par région"
    ch.add_data(Reference(ws, min_col=4, min_row=r0, max_row=r1), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=r0 + 1, max_row=r1))
    ch.series[0].graphicalProperties.solidFill = BLUE
    ch.x_axis.scaling.orientation = "maxMin"
    ch.legend = None
    ch.dataLabels = DataLabelList(showVal=True, showSerName=False, showCatName=False, showLegendKey=False)
    ch.width, ch.height = 15, 10
    ws.add_chart(ch, f"B{r1 + 4}")
    ch2 = BarChart()
    ch2.style, ch2.title = 2, "Activations par année"
    ch2.add_data(Reference(ws, min_col=11, min_row=r0, max_row=r0 + len(years)), titles_from_data=True)
    ch2.set_categories(Reference(ws, min_col=10, min_row=r0 + 1, max_row=r0 + len(years)))
    ch2.series[0].graphicalProperties.solidFill = "5B93FF"
    ch2.legend = None
    ch2.dataLabels = DataLabelList(showVal=True, showSerName=False, showCatName=False, showLegendKey=False)
    ch2.width, ch2.height = 12, 10
    ws.add_chart(ch2, f"G{r1 + 4}")
    ws.freeze_panes = "A4"


def sheet_quality(wb, s):
    ws = wb.create_sheet("QUALITE_INTEGRATION")
    banner(ws, "Contrôle des coordonnées et intégration Mali / Niger", "Corrections appliquées sur la carte ; coordonnées d'origine conservées dans SITES.", 8)
    for c, w in zip("ABCDEFGH", (2, 24, 12, 30, 22, 50, 20, 3)):
        ws.column_dimensions[c].width = w
    section(ws, 4, "Synthèse du contrôle", 2, 3)
    header(ws, 5, ["Résultat du contrôle", "Sites", "Signification"], 2)
    meaning = {"OK": "Position cohérente avec le pays déclaré.", "Corrigé": "Signe « − » manquant sur la longitude : corrigé automatiquement.",
               "Corrigé, à confirmer": "Chiffre déplacé probable : correction proposée, à valider sur place.",
               "Coordonnées partagées": "Même position exacte que d'autres sites : même lieu ou saisie copiée.",
               "Frontière, à vérifier": "À moins de 10 km de la frontière, hors du pays déclaré.",
               "Position générique": "Coordonnées par défaut partagées par des sites de villes différentes : à géolocaliser.",
               "Hors pays déclaré": "Position incohérente : à reprendre."}
    for i, (k, v) in enumerate(meaning.items(), start=6):
        cell(ws, i, 2, k, bold=True, color=QC_COL[k])
        cell(ws, i, 3, f'=COUNTIFS(s_qc,B{i})', center=True, fmt="0")
        cell(ws, i, 4, v, color=MUTED, size=9)
        ws.merge_cells(start_row=i, start_column=4, end_row=i, end_column=6)
        ws.row_dimensions[i].height = 28
    r = 6 + len(meaning) + 1
    section(ws, r, "Sites à corriger ou à confirmer", 2, 6)
    header(ws, r + 1, ["N°", "Nom", "Ville", "Contrôle", "Détail", "Fiche"], 2)
    bad = s[s.qc != "OK"].sort_values(["qc", "ville"])
    for i, row in enumerate(bad.itertuples(index=False), start=r + 2):
        cell(ws, i, 2, row.id, center=True)
        cell(ws, i, 3, row.nom)
        cell(ws, i, 4, row.ville)
        cell(ws, i, 5, row.qc, bold=True, color=QC_COL.get(row.qc, MUTED))
        cell(ws, i, 6, row.qc_detail, size=9, color=MUTED)
        x = cell(ws, i, 7, f'=HYPERLINK(URL_CARTE&"#site-"&B{i},"Carte ↗")')
        x.font = f(9, color="0563C1", underline="single")
        ws.row_dimensions[i].height = 30
    ws.column_dimensions["C"].width = 30
    r = r + 2 + len(bad) + 1
    section(ws, r, "Fédérer les trois pays : procédure d'intégration du Mali et du Niger", 2, 6)
    steps = [
        ("1. Export", "Extraire les points de vente actifs du Mali et du Niger au même format que l'export du Burkina (mêmes colonnes, dont Latitude, Longitude, Pays)."),
        ("2. Génération", "Lancer : python canal_sites/build_sites.py <dossier_geo> export_BF.xlsx export_ML.xlsx export_NE.xlsx, puis python canal_sites/build_excel.py."),
        ("3. Contrôle", "Le traitement vérifie chaque position contre le pays déclaré, corrige les erreurs de signe et rattache les sites aux cercles (Mali) et départements (Niger)."),
        ("4. Déploiement", "Glisser le nouveau dossier site_netlify sur le site Netlify existant (onglet Deploys) : le lien ne change pas."),
        ("5. Suivi sécuritaire", "Sur la carte, importer la base d'incidents (CSV ou Excel, colonnes lat / lon) : chaque site reçoit son exposition dans un rayon de 5 à 50 km."),
    ]
    for i, (a, b) in enumerate(steps, start=r + 1):
        cell(ws, i, 2, a, bold=True)
        cell(ws, i, 3, b)
        ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=7)
        ws.row_dimensions[i].height = 32
    ws.freeze_panes = "A4"


def main():
    s = pd.read_pickle(DIST / "sites.pkl")
    wb = Workbook()
    sheet_dash(wb, s)
    sheet_sites(wb, s)
    sheet_quality(wb, s)
    for ws in wb.worksheets:
        ws.sheet_view.zoomScale = 90
        ws.page_setup.orientation = "landscape"
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    wb["TABLEAU_DE_BORD"].sheet_properties.tabColor = NAVY
    wb["SITES"].sheet_properties.tabColor = BLUE
    wb["QUALITE_INTEGRATION"].sheet_properties.tabColor = "C76A0A"
    wb.calculation.fullCalcOnLoad = True
    wb.save(OUT)
    print(OUT)


if __name__ == "__main__":
    main()
