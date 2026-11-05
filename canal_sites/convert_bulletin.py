"""Convertit un bulletin mensuel d'incidents (feuille « BASE DE DONNEES INCIDENTS ») au format
du modèle d'import de l'application de sûreté des sites.

Géolocalisation, de la plus précise à la moins précise :
  1. coordonnées GPS du bulletin (DMS converties en degrés décimaux) ;
  2. même lieu géolocalisé dans un autre incident du bulletin ;
  3. chef-lieu de commune connu (Natural Earth, ou point de vente dont la province concorde),
     décalé si le bulletin indique une distance et une direction (« 15km NE Arbinda ») ;
  4. centre de la province (geoBoundaries ADM2).
Les sources (téléphones, noms d'informateurs) sont remplacées par un code ; la correspondance
est écrite dans un fichier séparé, à conserver en interne.

Usage : python canal_sites/convert_bulletin.py <bulletin.xlsx> <dossier_geo> <dossier_sortie>
"""
import json
import math
import re
import sys
import unicodedata
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from shapely.geometry import Point, shape

HERE = Path(__file__).resolve().parent
COLS = ["date", "lat", "lon", "type", "gravite", "tues", "blesses", "localite", "region", "auteur", "description", "source"]
EXTRA = ["id_incident", "heure", "pays", "province", "commune", "precision_geo", "methode_geo", "niveau_menace", "statut"]

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_sites import PROV_REGION_2025 as PROV_REGION  # noqa: E402

DIRS = {"N": 0, "NE": 45, "E": 90, "EST": 90, "SE": 135, "S": 180, "SUD": 180, "SO": 225, "SW": 225, "O": 270, "OUEST": 270, "W": 270, "NO": 315, "NW": 315, "NORD": 0}
TYPE_GRAVE = re.compile(r"attaque|eei|ied|kamikaze|explosif|projectile", re.I)
NOT_GRAVE = re.compile(r"d[ée]minage|destruction|soutien|dotation|arrestation|interpellation|orpaillage|surveillance|survol|alerte|projet|pr[ée]paration|menace|rassemblement|mouvement|interception|r[ée]union", re.I)


def n(s):
    s = unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode().upper()
    return re.sub(r"[^A-Z0-9 ]", " ", s).split()


def key(s):
    return " ".join(n(s))


def dms(s):
    m = re.findall(r"(\d+)°(\d+)′([\d.]+)″([NSEWO])", str(s))
    if len(m) != 2:
        return None
    out = []
    for d, mi, se, h in m:
        v = int(d) + int(mi) / 60 + float(se) / 3600
        out.append(-v if h in "SWO" else v)
    return round(out[0], 6), round(out[1], 6)


def offset(lat, lon, km, bearing):
    b = math.radians(bearing)
    return round(lat + km * math.cos(b) / 110.6, 5), round(lon + km * math.sin(b) / (111.32 * math.cos(math.radians(lat))), 5)


def build_gazetteer(geo):
    """Chefs-lieux fiables : Natural Earth + points de vente (médiane par ville) dont la province concorde."""
    alias = {"KOMONJDJARI": "KOMANDJARI"}
    provs = [(alias.get(key(f["properties"]["shapeName"]), key(f["properties"]["shapeName"])), shape(f["geometry"])) for f in json.loads((geo / "bfa_ADM2.geojson").read_text(encoding="utf-8"))["features"]]
    prov_of = lambda lat, lon: next((p for p, g in provs if g.contains(Point(lon, lat))), None)  # noqa: E731
    gaz = {}
    for f in json.loads((geo / "ne_10m_populated_places_simple.geojson").read_text(encoding="utf-8"))["features"]:
        p = f["properties"]
        if p["adm0_a3"] == "BFA":
            gaz[key(p["name"])] = (round(p["latitude"], 5), round(p["longitude"], 5), prov_of(p["latitude"], p["longitude"]), "Natural Earth")
    sites_pkl = HERE / "dist" / "sites.pkl"
    if sites_pkl.exists():
        s = pd.read_pickle(sites_pkl)
        s = s[s.qc.isin(["OK", "Corrigé", "Frontière, à vérifier"])].copy()
        s["base"] = s.ville.str.split("/").str[0].map(key)
        for b, g in s.groupby("base"):
            if b and b not in gaz:
                lat, lon = g.lat.median(), g.lon.median()
                gaz[b] = (round(lat, 5), round(lon, 5), prov_of(lat, lon), f"points de vente ({len(g)})")
    centroids = {p: tuple(round(c, 5) for c in reversed(g.representative_point().coords[0])) for p, g in provs}
    return gaz, centroids, prov_of


def convert(xlsx, geo):
    d = pd.read_excel(xlsx, sheet_name="BASE DE DONNEES INCIDENTS")
    d = d[d["ID Incident"].astype(str).str.startswith("INC")].reset_index(drop=True)
    gaz, centroids, prov_of = build_gazetteer(geo)
    # lieux déjà géolocalisés dans le bulletin (réutilisés pour les incidents sans GPS au même endroit)
    known = {}
    for _, r in d.iterrows():
        c = dms(r["Coordonnées GPS"])
        if c:
            for part in re.split(r"[/,]|->|\(|\)", str(r["Localité Principale"])):
                if key(part):
                    known.setdefault(key(part), c)
    src_codes, rows, checks = {}, [], []
    for _, r in d.iterrows():
        iid = r["ID Incident"]
        prov_txt = str(r["Province"])
        prov_keys = [key(p) for p in re.split(r"/", re.sub(r"\(.*?\)", "", prov_txt)) if key(p)]
        lieu, commune = str(r["Localité Principale"]), str(r["Département / Commune"])
        c, method, prec = dms(r["Coordonnées GPS"]), "", ""
        if c:
            method, prec = "GPS du bulletin", "GPS (≈ 10 m)"
        if not c:
            for part in re.split(r"[/,]|->|\(|\)", lieu):
                if key(part) in known:
                    c, method, prec = known[key(part)], f"GPS d'un autre incident au même lieu ({part.strip()})", "GPS (≈ 10 m)"
                    break
        if not c:  # « 15km NE Arbinda », « 7km Est de Sangha »
            m = re.search(r"(\d+)\s*km\s*(NE|NO|NW|SE|SO|SW|N|S|E|O|W|Est|Ouest|Nord|Sud)\s*(?:de\s+|d')?([A-Za-zÀ-ÿ' -]+)", lieu + " " + str(r["Nature & Description de l'Événement"]), re.I)
            if m and key(m.group(3).split()[0]) in gaz:
                g = gaz[key(m.group(3).split()[0])]
                if g[2] in prov_keys or not g[2]:
                    c = offset(g[0], g[1], int(m.group(1)), DIRS[m.group(2).upper()])
                    method, prec = f"{m.group(1)} km {m.group(2)} de {m.group(3).split()[0]} ({g[3]})", "Localité estimée (≈ 3 km)"
        if not c:
            parts = [p for p in re.split(r"[/,]|->|\(|\)", lieu) if key(p)]
            # seule la première localité citée vaut position « localité » ; les suivantes restent approximatives
            cands = [(p, k == 0) for k, p in enumerate(parts)] + [(p, False) for p in re.split(r"[/,]", commune) if key(p)]
            for part, from_lieu in cands:
                k = key(part)
                for kk in (k, k.replace(" SECTEUR 06", ""), " ".join(n(part)[:1])):
                    g = gaz.get(kk)
                    if g and (g[2] in prov_keys or not prov_keys):
                        c = (g[0], g[1])
                        method = f"{'Localité' if from_lieu else 'Chef-lieu de la commune'} « {part.strip()} » ({g[3]})"
                        prec = "Localité (≈ 2 km)" if from_lieu else "Chef-lieu de commune (≈ 5 à 30 km)"
                        break
                if c:
                    break
        if not c and prov_keys and prov_keys[0] in centroids:
            c = centroids[prov_keys[0]]
            method, prec = f"Centre de la province {prov_txt}", "Province (≈ 30 à 60 km)"
        lat, lon = c if c else (None, None)
        # contrôle : province calculée à partir des coordonnées
        p_calc = prov_of(lat, lon) if c else None
        pays = str(r["Pays"])
        if c and p_calc is None and "mali" not in pays.lower():
            checks.append((iid, "Coordonnées au Mali, juste derrière la frontière, alors que le pays déclaré est le Burkina Faso : position ou pays à confirmer."))
        if c and p_calc and prov_keys and p_calc not in prov_keys and method.startswith("GPS"):
            checks.append((iid, f"Coordonnées GPS situées dans la province {p_calc.title()} et non {prov_txt} : à vérifier."))
        # région : celle du bulletin si elle est précisée entre parenthèses, sinon déduite de la province
        rz = str(r["Région / Zone"])
        region = PROV_REGION.get(prov_keys[0] if prov_keys else "", "")
        if not region:
            mreg = re.search(r"\(([^)]+)\)", rz)
            region = mreg.group(1).split("/")[0].strip() if mreg else rz
        # date
        raw_date = str(r["Date Incident"])
        dm = re.match(r"(\d{4}-\d{2}-\d{2})", raw_date)
        date = dm.group(1) if dm else ""
        if "/" in raw_date:
            checks.append((iid, f"Date sur plusieurs jours ({raw_date}) : premier jour retenu ({date})."))
        # bilan
        mil, civ, bl = int(r["Pertes Mil./VDP"] or 0), int(r["Pertes Civiles"] or 0), int(r["Blessés"] or 0)
        typ = str(r["Type d'Incident"]).strip()
        bilan = str(r["Bilan Humain Synthétique"])
        if "VDP" in bilan and civ and not mil:
            checks.append((iid, f"Bilan « {bilan[:60]}… » : VDP comptés en pertes civiles dans le bulletin ; total des tués inchangé ({civ})."))
        grave = (mil + civ + bl > 0 or TYPE_GRAVE.search(typ)) and not NOT_GRAVE.search(typ) and "accident" not in bilan.lower()
        if "orpaillage" in typ.lower() and civ:
            checks.append((iid, "Décès accidentel (éboulement sur site d'orpaillage) : compté en tué mais classé non grave (pas un acte hostile)."))
        # auteur
        txt = f"{typ} {r['Nature & Description de l\'Événement']} {r['Statut / Recommandations']}"
        if re.search(r"op[ée]ration|d[ée]minage|arrestation|interpellation|dotation|d[ée]ploiement|riposte amie|retour de mission", typ, re.I):
            auteur = "FDS / VDP"
        elif "orpaillage" in typ.lower():
            auteur = ""
        elif "JNIM" in txt and "EIGS" in txt:
            auteur = "JNIM / EIGS"
        elif "JNIM" in txt:
            auteur = "JNIM"
        elif "EIGS" in txt:
            auteur = "EIGS"
        else:
            auteur = "GT (non identifié)"
        # source pseudonymisée
        src = str(r["Source / Unité"]).strip()
        code = src_codes.setdefault(src, f"SRC-{len(src_codes) + 1:02d}")
        # description sans noms de personnes
        desc = str(r["Nature & Description de l'Événement"]).strip()
        desc2 = re.sub(r"\((?:[A-Z][a-zé]+ [A-Z]\.,?\s*){2,}\)", "(identités au rapport)", desc)
        if desc2 != desc:
            checks.append((iid, "Noms de personnes retirés de la description."))
        deg, stat = str(r["Dégâts Matériels / Pillage"]), str(r["Statut / Recommandations"])
        if deg not in ("nan", "Néant", "Aucun", ""):
            desc2 += f" Dégâts : {deg}."
        if stat and stat != "nan":
            desc2 += f" Suivi : {stat}."
        if prec.startswith("Province"):
            checks.append((iid, f"Lieu « {lieu} » non géolocalisé : placé au centre de la province ({prov_txt}). Coordonnées GPS à demander à la source."))
        rows.append({
            "date": date, "lat": lat, "lon": lon, "type": typ, "gravite": "grave" if grave else "non", "tues": mil + civ, "blesses": bl,
            "localite": lieu, "region": region, "auteur": auteur, "description": desc2, "source": code,
            "id_incident": iid, "heure": "" if str(r["Heure"]) == "nan" else str(r["Heure"]), "pays": pays, "province": prov_txt, "commune": commune,
            "precision_geo": prec, "methode_geo": method, "niveau_menace": str(r["Niveau de Menace"]), "statut": stat,
        })
    return pd.DataFrame(rows), checks, src_codes


def write(df, checks, src_codes, out):
    out.mkdir(parents=True, exist_ok=True)
    head = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    fill = PatternFill("solid", start_color="0E5C5A", end_color="0E5C5A")
    extra_fill = PatternFill("solid", start_color="5D6668", end_color="5D6668")
    thin = Side(style="thin", color="DCDAD2")
    wb = Workbook()
    ws = wb.active
    ws.title = "Incidents"
    ws.append(COLS + EXTRA)
    for i, c in enumerate(ws[1]):
        c.font, c.fill = head, fill if i < len(COLS) else extra_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for _, r in df.iterrows():
        ws.append([r[c] for c in COLS + EXTRA])
    widths = {"date": 11, "lat": 10, "lon": 10, "type": 28, "gravite": 9, "tues": 6, "blesses": 8, "localite": 28, "region": 13, "auteur": 16, "description": 70, "source": 9,
              "id_incident": 13, "heure": 11, "pays": 14, "province": 16, "commune": 20, "precision_geo": 26, "methode_geo": 38, "niveau_menace": 13, "statut": 40}
    for i, c in enumerate(COLS + EXTRA, start=1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = widths[c]
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name="Arial", size=9)
            c.alignment = Alignment(vertical="top", wrap_text=c.column_letter in ("K", "H", "S", "U"))
            c.border = Border(bottom=thin)
        if row[4].value == "grave":
            row[4].font = Font(name="Arial", size=9, bold=True, color="C62A2A")
        if str(row[17].value).startswith("Province"):
            row[17].font = Font(name="Arial", size=9, bold=True, color="E07B1A")
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    cs = wb.create_sheet("Contrôle")
    cs.append(["ID incident", "Point de contrôle"])
    for c in cs[1]:
        c.font, c.fill = head, fill
    for iid, msg in checks:
        cs.append([iid, msg])
    cs.column_dimensions["A"].width, cs.column_dimensions["B"].width = 14, 120
    for row in cs.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name="Arial", size=9)
            c.alignment = Alignment(wrap_text=True, vertical="top")
    rs = wb.create_sheet("Lisez-moi")
    lines = [
        ("Fichier d'import des incidents d'octobre 2026", True),
        ("Importer la feuille « Incidents » dans l'onglet Sécurité de l'application (glisser le fichier). Les 12 premières colonnes (en vert) suivent le modèle d'import ; les colonnes grises sont des informations de traçabilité, ignorées par l'application.", False),
        ("gravite = « grave » pour les actes hostiles réalisés (attaques, y compris repoussées ; EEI ayant explosé ; attaques de drone) ou avec victimes ; « non » pour les mouvements, rassemblements, alertes, opérations amies, déminages, arrestations, pillages sans victime et accidents.", False),
        ("tues = pertes militaires/VDP + pertes civiles du bulletin ; blesses = colonne Blessés. region = région de 2025 déduite de la province (17 régions, noms du bulletin).", False),
        ("precision_geo indique la fiabilité de la position : GPS (≈ 10 m), localité estimée, chef-lieu de commune, ou centre de province. Les incidents au centre d'une province faussent les distances aux sites : compléter leurs coordonnées en priorité (voir feuille Contrôle).", False),
        ("source = code anonymisé. La correspondance code → source est dans un fichier séparé, à ne pas diffuser.", False),
    ]
    for t, b in lines:
        rs.append([t])
        rs.cell(rs.max_row, 1).font = Font(name="Arial", size=12 if b else 10, bold=b)
        rs.cell(rs.max_row, 1).alignment = Alignment(wrap_text=True, vertical="top")
    rs.column_dimensions["A"].width = 130
    p1 = out / "Import_incidents_octobre_2026.xlsx"
    wb.save(p1)
    wb2 = Workbook()
    w2 = wb2.active
    w2.title = "Sources (confidentiel)"
    w2.append(["Code", "Source / unité d'origine"])
    for k, v in src_codes.items():
        w2.append([v, k])
    w2.column_dimensions["A"].width, w2.column_dimensions["B"].width = 10, 40
    p2 = out / "Correspondance_sources_CONFIDENTIEL.xlsx"
    wb2.save(p2)
    return p1, p2


if __name__ == "__main__":
    df, checks, codes = convert(Path(sys.argv[1]), Path(sys.argv[2]))
    p1, p2 = write(df, checks, codes, Path(sys.argv[3]))
    print(p1, p2)
    print(df.precision_geo.value_counts().to_string())
    print("graves", (df.gravite == "grave").sum(), "tués", df.tues.sum(), "blessés", df.blesses.sum())
    print(df.region.value_counts().to_string())
    for c in checks:
        print(*c)
    print(df[["id_incident", "localite", "lat", "lon", "precision_geo", "methode_geo", "gravite", "auteur"]].to_string())
