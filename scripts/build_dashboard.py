"""Construit le tableau de bord décisionnel Mali à partir de la base Excel.

Étapes : lecture de DATA_CARTOGRAPHIE -> normalisation (acteurs, modes opératoires,
cibles, sources) -> jointure spatiale (pays, région, cercle) -> indicateurs ->
écriture de dist/ (page HTML autonome, fragment Artifact, GeoJSON et CSV enrichis).

Usage : python scripts/build_dashboard.py [chemin.xlsx]
"""
import json
import math
import re
import sys
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from shapely.geometry import Point, shape

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_XLSX = ROOT / "data" / "Base_Donnees_Securite_Mali_OpenMap_Dashboard.xlsx"
BAMAKO = (12.6392, -8.0029)

LOSS_COLS = {
    "Morts_Civils": "civ_k", "Blesses_Civils": "civ_w", "Morts_FDS_FAMA": "fama_k",
    "Morts_Allies_AfricaCorps": "ac_k", "Morts_Groupes_Armes": "ga_k", "Captures_Prisonniers": "pow",
    "Vehicules_Aéronefs_Détruits": "veh_d", "Vehicules_Aéronefs_Saisis": "veh_s",
}

# Mode opératoire normalisé pour chaque typologie présente dans la base.
MODE = {
    "Embuscade": "Embuscade", "Embuscade / Fuite de Convoi": "Embuscade", "Embuscade / Riposte": "Embuscade",
    "Embuscade sur Patrouille": "Embuscade", "Embuscade Mortelle sur Convoi": "Embuscade",
    "Embuscade & Affrontement Inter-groupe": "Embuscade",
    "Engin Explosif Improvisé (EEI)": "EEI", "EEI & Embuscade Dozo": "EEI", "Attaque à l'EEI Dirigé": "EEI",
    "Attaque de Poste / Localité": "Attaque de poste / camp", "Attaque Majeure de Camp / Base": "Attaque de poste / camp",
    "Attaque de Poste / Infiltration": "Attaque de poste / camp", "Attaque de Caserne": "Attaque de poste / camp",
    "Attaque de Poste / Infiltration Sud": "Attaque de poste / camp", "Attaque / Menace Sécuritaire": "Attaque de poste / camp",
    "Explosion / Attaque Ciblée": "Attaque ciblée / VIP", "Attaque de VIP / Chef de Milice": "Attaque ciblée / VIP",
    "Tirs de Mortier / Artillerie": "Tirs indirects",
    "Bombardement Aérien": "Frappe aérienne étatique", "Frappe Drone / Bavure Civile": "Frappe aérienne étatique",
    "Bombardement Aérien / Bavure Civile": "Frappe aérienne étatique", "Campagne de Frappes Aériennes": "Frappe aérienne étatique",
    "Frappe Aérienne / Représailles": "Frappe aérienne étatique",
    "Frappe Drone / Air": "Drone armé (FPV)", "Frappe Drone Kamikaze": "Drone armé (FPV)", "Frappe Drone FPV": "Drone armé (FPV)",
    "Défense Antiaérienne / Drone Abattu": "Lutte anti-drone",
    "Affrontement Inter-JNIM-EIS": "Affrontement JNIM–EIS", "Bataille Majeure Inter-terroriste": "Affrontement JNIM–EIS",
    "Sabotage / Incendie de Convoi": "Sabotage logistique",
    "Reconquête Stratégique": "Offensive terrestre",
    "Crash d'Aéronef / Hélicoptère": "Perte d'aéronef (accident)", "Crash Drone de Combat MALE": "Perte d'aéronef (accident)",
}

SOURCE_KIND = {
    "AzZalaqa": "partisane", "Amaq": "partisane", "EIS": "partisane", "JNIM": "partisane", "FLA": "partisane",
    "État-Major FAMa": "officielle", "Jeune Afrique": "média",
    "Naïm TOURÉ": "compilateur",
}
SOURCE_NOTE = {
    "AzZalaqa": "Organe de propagande du JNIM", "Amaq": "Agence de propagande de l'État islamique",
    "EIS": "Revendication EI-Sahel", "JNIM": "Revendication JNIM", "FLA": "Communication du FLA",
    "État-Major FAMa": "Communiqué officiel des FAMa", "Jeune Afrique": "Presse panafricaine",
    "Naïm TOURÉ": "Compilateur de la base",
}


# Graphie française des régions et cercles (geoBoundaries les publie sans accents).
NAMES_FR = {
    "Segou": "Ségou", "Koulikouro": "Koulikoro", "Tenenkou": "Ténenkou", "Djenne": "Djenné", "Niafunke": "Niafunké",
    "Dire": "Diré", "Menaka": "Ménaka", "Bafoulabe": "Bafoulabé", "Diema": "Diéma", "Kenieba": "Kéniéba",
    "Yelimane": "Yélimané", "Dioila": "Dioïla", "Abeibara": "Abeïbara", "Baroueli": "Barouéli",
}


def fr(name):
    return NAMES_FR.get(name, name)


def actor_family(s):
    s = str(s)
    if s.startswith("JNIM vs"):
        return "JNIM"
    if "Africa Corps" in s or s.startswith("FAMa"):
        return "FAMa / Africa Corps"
    if "JNIM" in s:
        return "JNIM"
    if "EIS" in s or "EI-Sahel" in s:
        return "EI-Sahel"
    if "FLA" in s or "Azawad" in s:
        return "FLA"
    return "Indéterminé / accident"


def target_family(s, mode):
    s = str(s)
    if re.search(r"Civil|Village|Maire|Starlink", s):
        return "Civils"
    if re.search(r"Dozo|Milice|Toloba", s):
        return "Milices pro-gouv."
    if "Transport commercial" in s:
        return "Logistique civile"
    if re.match(r"(Positions )?(JNIM|FLA)|JNIM / EIS", s) or mode == "Affrontement JNIM–EIS":
        return "Groupes armés"
    return "FAMa / Africa Corps"


def haversine(a, b):
    la1, lo1, la2, lo2 = map(math.radians, (*a, *b))
    h = math.sin((la2 - la1) / 2) ** 2 + math.cos(la1) * math.cos(la2) * math.sin((lo2 - lo1) / 2) ** 2
    return 2 * 6371 * math.asin(math.sqrt(h))


def decimals(v):
    t = repr(float(v)).rstrip("0").split(".")
    return len(t[1]) if len(t) > 1 else 0


def nearest(polys, pt):
    return min(polys, key=lambda p: p[1].distance(pt))[0]


def read_incidents(xlsx):
    raw = pd.read_excel(xlsx, sheet_name="DATA_CARTOGRAPHIE", header=None)
    hdr = next(i for i, r in raw.iterrows() if r.iloc[0] == "ID_Incident")
    df = pd.read_excel(xlsx, sheet_name="DATA_CARTOGRAPHIE", header=hdr)
    df = df[df["ID_Incident"].astype(str).str.startswith("INC")].copy()
    df["Date"] = pd.to_datetime(df["Date"]).dt.date
    for c in LOSS_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
    return df.sort_values(["Date", "ID_Incident"]).reset_index(drop=True)


def build(xlsx):
    geo = json.loads((ROOT / "data" / "geo" / "basemap.json").read_text(encoding="utf-8"))
    for layer in ("adm1", "adm2"):
        for f in geo[layer]["features"]:
            f["properties"]["name"] = fr(f["properties"]["name"])
    adm1 = [(f["properties"]["name"], shape(f["geometry"])) for f in geo["adm1"]["features"]]
    adm2 = [(f["properties"]["name"], shape(f["geometry"])) for f in geo["adm2"]["features"]]
    ctry = [(f["properties"]["name"], shape(f["geometry"])) for f in geo["countries"]["features"]]
    df = read_incidents(xlsx)

    d0, d1 = min(df["Date"]), max(df["Date"])
    split = d0 + timedelta(days=((d1 - d0).days + 1) // 2)

    incidents = []
    for _, r in df.iterrows():
        pt = Point(r["Longitude"], r["Latitude"])
        country = next((n for n, g in ctry if g.contains(pt)), nearest(ctry, pt))
        in_mali = country == "Mali"
        region = next((n for n, g in adm1 if g.contains(pt)), None) if in_mali else None
        cercle = next((n for n, g in adm2 if g.contains(pt)), None) if in_mali else None
        mode = MODE.get(r["Typologie_Incident"], "Autre")
        srcs = [s.strip() for s in str(r["Sources"]).split(",") if s.strip()]
        kinds = {s: SOURCE_KIND.get(s, "osint") for s in srcs}
        n_ind = sum(k in ("osint", "média") for k in kinds.values())
        n_part = sum(k in ("partisane", "officielle") for k in kinds.values())
        if n_ind >= 2:
            rel = "Corroboré"
        elif n_ind == 1:
            rel = "Partiellement corroboré"
        elif n_part:
            rel = "Source partisane seule"
        else:
            rel = "Source unique"
        loss = {v: int(r[k]) for k, v in LOSS_COLS.items()}
        killed = loss["civ_k"] + loss["fama_k"] + loss["ac_k"] + loss["ga_k"]
        prec = min(decimals(r["Latitude"]), decimals(r["Longitude"]))
        incidents.append({
            "id": r["ID_Incident"], "date": r["Date"].isoformat(), "period": 1 if r["Date"] < split else 2,
            "country_reported": r["Pays"], "country": country, "region_reported": r["Region_Secteur"],
            "region": region or f"Hors Mali ({country})", "cercle": cercle or "—", "place": r["Localite"],
            "lat": float(r["Latitude"]), "lon": float(r["Longitude"]),
            "precision": "Zone (±10 km)" if prec <= 1 else "Localité",
            "typology": r["Typologie_Incident"], "mode": mode,
            "actor_raw": r["Acteur_Instigateur"], "actor": actor_family(r["Acteur_Instigateur"]),
            "actor_presumed": "Probablement" in str(r["Acteur_Instigateur"]),
            "target_raw": r["Acteur_Cible"], "target": target_family(r["Acteur_Cible"], mode),
            "third": r["Acteur_Tiers"], "desc": r["Description_Details"],
            **loss, "killed": killed,
            "equipment": r["Armements_Equipements_Saisis"], "impact": r["Impact_Infrastructures"],
            "sources": [{"name": s, "kind": kinds[s], "note": SOURCE_NOTE.get(s, "Analyste / veille OSINT")} for s in srcs],
            "reliability": rel, "claim_only": n_ind == 0 and killed > 0,
            "dist_bamako": round(haversine(BAMAKO, (r["Latitude"], r["Longitude"]))),
            "axis": str(r["Localite"]).startswith("Axe") or "Axe" in str(r["Localite"]),
        })
    return geo, incidents, d0, d1, split


def write_outputs(geo, incidents, d0, d1, split):
    dist = ROOT / "dist"
    dist.mkdir(exist_ok=True)
    meta = {
        "start": d0.isoformat(), "end": d1.isoformat(), "split": split.isoformat(),
        "days": (d1 - d0).days + 1, "generated": date.today().isoformat(),
        "bamako": {"lat": BAMAKO[0], "lon": BAMAKO[1]},
    }
    payload = {"meta": meta, "incidents": incidents, "geo": geo}
    data_js = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))

    tpl = (ROOT / "dashboard" / "index.html").read_text(encoding="utf-8")
    fragment = tpl.replace("/*__DATA__*/null", data_js)
    (dist / "artifact_mali.html").write_text(fragment, encoding="utf-8")
    standalone = "<!doctype html>\n<html lang=\"fr\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1, viewport-fit=cover\">\n</head>\n<body>\n" + fragment + "\n</body>\n</html>\n"
    (dist / "tableau_de_bord_mali.html").write_text(standalone, encoding="utf-8")

    feats = [{"type": "Feature", "geometry": {"type": "Point", "coordinates": [i["lon"], i["lat"]]},
              "properties": {k: (", ".join(s["name"] for s in v) if k == "sources" else v) for k, v in i.items() if k not in ("lat", "lon")}}
             for i in incidents]
    (dist / "incidents_mali.geojson").write_text(json.dumps({"type": "FeatureCollection", "features": feats}, ensure_ascii=False, indent=1), encoding="utf-8")
    (dist / "incidents_mali.kml").write_text(kml(incidents), encoding="utf-8")
    flat = pd.DataFrame([{**i, "sources": ", ".join(s["name"] for s in i["sources"])} for i in incidents])
    flat.to_csv(dist / "incidents_mali_enrichi.csv", index=False, encoding="utf-8-sig")
    return meta


# Couleurs KML (aabbggrr) alignées sur la légende du tableau de bord.
KML_COLORS = {"JNIM": "ff264adc", "EI-Sahel": "ffd64f8a", "FLA": "ff129ad6", "FAMa / Africa Corps": "ffb56f2f", "Indéterminé / accident": "ff978a7d"}


def kml(incidents):
    """Couche KML pour Google Earth / Google My Maps : un dossier par acteur, une fiche par incident."""
    from xml.sax.saxutils import escape as x

    styles = "".join(
        f'<Style id="s{k}"><IconStyle><color>{c}</color><scale>1.1</scale><Icon><href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href></Icon></IconStyle><LabelStyle><scale>0.7</scale></LabelStyle></Style>'
        for k, c in enumerate(KML_COLORS.values()))
    keys = list(KML_COLORS)
    folders = []
    for a in keys:
        marks = []
        for i in (i for i in incidents if i["actor"] == a):
            desc = (f"<b>{x(i['date'])} · {x(i['region'])} / {x(i['cercle'])}</b><br/>{x(i['mode'])} — {x(i['typology'])}<br/>"
                    f"Instigateur : {x(i['actor_raw'])} · Cible : {x(i['target_raw'])}<br/><br/>{x(i['desc'])}<br/><br/>"
                    f"Tués : {i['killed']} (FDS {i['fama_k']}, Africa Corps {i['ac_k']}, groupes armés {i['ga_k']}, civils {i['civ_k']}) · Civils blessés : {i['civ_w']}<br/>"
                    f"Capturés : {i['pow']} · Véhicules détruits / saisis : {i['veh_d']} / {i['veh_s']}<br/>Fiabilité : {x(i['reliability'])} · Sources : {x(', '.join(s['name'] for s in i['sources']))}")
            marks.append(f"<Placemark><name>{x(i['id'])} · {x(i['place'])}</name><TimeStamp><when>{i['date']}</when></TimeStamp><styleUrl>#s{keys.index(a)}</styleUrl>"
                         f"<description><![CDATA[{desc}]]></description><Point><coordinates>{i['lon']},{i['lat']},0</coordinates></Point></Placemark>")
        if marks:
            folders.append(f"<Folder><name>{x(a)} ({len(marks)})</name>{''.join(marks)}</Folder>")
    return ('<?xml version="1.0" encoding="UTF-8"?>\n<kml xmlns="http://www.opengis.net/kml/2.2"><Document><name>Incidents sécuritaires Mali</name>'
            f"{styles}{''.join(folders)}</Document></kml>\n")


def summary(incidents, meta):
    df = pd.DataFrame(incidents)
    print(meta)
    print(df[["civ_k", "civ_w", "fama_k", "ac_k", "ga_k", "pow", "veh_d", "veh_s", "killed"]].sum().to_dict())
    for c in ["actor", "mode", "target", "region", "reliability", "precision", "country"]:
        print(c, df.groupby(c).agg(n=("id", "size"), tues=("killed", "sum")).sort_values("n", ascending=False).to_dict("index"))
    print(df.groupby("period").agg(n=("id", "size"), tues=("killed", "sum"), fama=("fama_k", "sum"), ac=("ac_k", "sum"), civ=("civ_k", "sum")).to_dict("index"))
    print(df[df.country_reported.str.contains("Burkina") | (df.country != "Mali")][["id", "country_reported", "country", "region"]].to_string())
    print(df[["id", "region_reported", "region", "cercle", "dist_bamako", "reliability"]].to_string())


if __name__ == "__main__":
    geo, inc, d0, d1, split = build(Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_XLSX)
    meta = write_outputs(geo, inc, d0, d1, split)
    summary(inc, meta)
