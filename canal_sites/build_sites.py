"""Cartographie des points de vente agréés (Burkina Faso, Mali, Niger).

Lit un ou plusieurs exports « Points de vente » (même format que export_75.xlsx),
contrôle et corrige les coordonnées, rattache chaque site à sa région et à sa
province / son cercle / son département par jointure spatiale, puis produit :
  - dist/site_netlify/      dossier à déposer sur Netlify Drop (index.html, sans données de contact)
  - dist/carte_sites.html   même carte en fichier autonome
  - dist/Sites_PVA_BF_ML_NE.xlsx  classeur de suivi (synthèse, sites, contrôle qualité)
  - dist/sites_pva.geojson / .kml  couches SIG

Usage : python canal_sites/build_sites.py <dossier_sources_geo> <export1.xlsx> [export2.xlsx ...]
Le dossier géo doit contenir : ne_10m_admin_0_countries, ne_10m_admin_1_states_provinces,
ne_10m_roads, rivers, ne_10m_lakes (Natural Earth) et bfa_/mli_/ner_ ADM1/ADM2 (geoBoundaries).
"""
import json
import re
import math
import shutil
import sys
from datetime import date, datetime
from pathlib import Path
from xml.sax.saxutils import escape

import pandas as pd
from shapely.geometry import Point, box, mapping, shape
from shapely.ops import unary_union

HERE = Path(__file__).resolve().parent
DIST = HERE / "dist"
COUNTRIES = {"BFA": "Burkina Faso", "MLI": "Mali", "NER": "Niger"}
BBOX = box(-20.0, 2.0, 26.0, 30.0)

MLI_FR = {"Segou": "Ségou", "Koulikouro": "Koulikoro"}

# Burkina : provinces -> 17 régions de la réforme de 2025 (graphie des bulletins de sûreté).
# Rattachements à confirmer : Kompienga (Goulmou), Oudalan (Soum), Nayala (Sourou).
PROV_REGION_2025 = {
    "BALE": "Bankui", "BANWA": "Bankui", "MOUHOUN": "Bankui", "SOUROU": "Sourou", "NAYALA": "Sourou", "KOSSI": "Koosin",
    "BOUGOURIBA": "Djôrô", "IOBA": "Djôrô", "NOUMBIEL": "Djôrô", "PONI": "Djôrô", "GOURMA": "Goulmou", "KOMPIENGA": "Goulmou",
    "GNAGNA": "Sirba", "KOMANDJARI": "Sirba", "KOMONJDJARI": "Sirba", "TAPOA": "Tapoa", "HOUET": "Guiriko", "KENEDOUGOU": "Guiriko",
    "TUY": "Guiriko", "KADIOGO": "Kadiogo", "BAM": "Koulsé", "NAMENTENGA": "Koulsé", "SANMATENGA": "Koulsé", "SENO": "Liptako",
    "YAGHA": "Liptako", "SOUM": "Soum", "OUDALAN": "Soum", "BOULGOU": "Nakambé", "KOULPELOGO": "Nakambé", "KOURITENGA": "Nakambé",
    "BOULKIEMDE": "Nando", "SANGUIE": "Nando", "SISSILI": "Nando", "ZIRO": "Nando", "BAZEGA": "Nazinon", "NAHOURI": "Nazinon",
    "ZOUNDWEOGO": "Nazinon", "GANZOURGOU": "Oubri", "KOURWEOGO": "Oubri", "OUBRITENGA": "Oubri", "COMOE": "Tannounyan",
    "LERABA": "Tannounyan", "LOROUM": "Yaadga", "PASSORE": "Yaadga", "YATENGA": "Yaadga", "ZONDOMA": "Yaadga",
}


def pkey(name):
    import unicodedata
    return " ".join(re.sub(r"[^A-Z0-9 ]", " ", unicodedata.normalize("NFD", str(name)).encode("ascii", "ignore").decode().upper()).split())


def bfa_regions_2025(src):
    """Régions du Burkina reconstituées par fusion des provinces (geoBoundaries ADM2)."""
    groups = {}
    for f in load(src, "bfa_ADM2.geojson")["features"]:
        reg = PROV_REGION_2025.get(pkey(f["properties"]["shapeName"]))
        if reg:
            groups.setdefault(reg, []).append(shape(f["geometry"]).buffer(0))
    return [(name, unary_union(gs).buffer(0.0005).buffer(-0.0005)) for name, gs in sorted(groups.items())]
GENERIC_POINT = (12.355807, -1.520662)  # coordonnées attribuées à plusieurs sites de villes différentes

# Colonnes publiées sur le site (aucune donnée de contact ni identifiant de connexion).
PUBLIC = ["id", "nom", "pays", "ville", "adresse", "lat", "lon", "zone", "secteur", "region", "province",
          "region_declaree", "activation", "annee", "parent", "qc", "qc_detail", "voisin_km"]


def hav(lat1, lon1, lat2, lon2):
    p1, p2 = math.radians(lat1), math.radians(lat2)
    h = math.sin((p2 - p1) / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(math.radians(lon2 - lon1) / 2) ** 2
    return 2 * 6371 * math.asin(math.sqrt(h))


def load(src, name):
    return json.loads((src / name).read_text(encoding="utf-8"))


def admin(src):
    """Polygones pays, régions et niveau 2 pour les trois pays."""
    a0 = {f["properties"]["ADM0_A3"]: shape(f["geometry"]) for f in load(src, "ne_10m_admin_0_countries.geojson")["features"]
          if f["properties"]["ADM0_A3"] in COUNTRIES}
    reg, sub = [], []
    for name, g in bfa_regions_2025(src):
        reg.append(("BFA", name, g))
    for f in load(src, "gb_ADM1.geojson")["features"]:
        reg.append(("MLI", MLI_FR.get(f["properties"]["shapeName"], f["properties"]["shapeName"]), shape(f["geometry"])))
    for f in load(src, "ne_10m_admin_1_states_provinces.geojson")["features"]:
        if f["properties"]["adm0_a3"] == "NER":
            reg.append(("NER", f["properties"]["name"], shape(f["geometry"])))
    for iso, fn in (("BFA", "bfa_ADM2.geojson"), ("MLI", "gb_ADM2.geojson"), ("NER", "ner_ADM2.geojson")):
        for f in load(src, fn)["features"]:
            sub.append((iso, MLI_FR.get(f["properties"]["shapeName"], f["properties"]["shapeName"]), shape(f["geometry"])))
    return a0, reg, sub


def read_exports(paths):
    frames = []
    for p in paths:
        d = pd.read_excel(p)
        d["fichier_source"] = Path(p).name
        frames.append(d)
    d = pd.concat(frames, ignore_index=True)
    return d.drop_duplicates("NumDist").reset_index(drop=True)


def qc_and_geocode(d, a0, reg, sub):
    union = {iso: g.buffer(0.02) for iso, g in a0.items()}
    rows = []
    for _, r in d.iterrows():
        lat0, lon0 = float(r["Latitude"]), float(r["Longitude"])
        lat, lon, qc, detail = lat0, lon0, "OK", ""
        declared = {"BURKINA FASO": "BFA", "MALI": "MLI", "NIGER": "NER"}.get(str(r["Pays"]).strip().upper())
        inside = lambda la, lo: next((iso for iso, g in union.items() if g.contains(Point(lo, la))), None)  # noqa: E731
        iso = inside(lat, lon)
        if declared and iso != declared:
            if lon > 0 and inside(lat, -lon) == declared:
                lon, iso = -lon, declared
                qc, detail = "Corrigé", "Signe de la longitude manquant : longitude rendue négative."
            elif str(r["Ville"]).strip().upper() == "BOULSA" and abs(lon0 + 5.7431) < 1e-3:
                lon, iso = -0.5743, declared
                qc, detail = "Corrigé, à confirmer", "Longitude -5,7431 (au Mali) ; chiffre déplacé probable : -0,5743, à Boulsa."
            elif a0[declared].distance(Point(lon, lat)) * 111 <= 10:
                iso = declared
                qc, detail = "Frontière, à vérifier", f"Point à moins de 10 km de la frontière, hors du pays déclaré ({COUNTRIES[declared]}) : position à confirmer."
            else:
                qc, detail = "Hors pays déclaré", f"Coordonnées hors du pays déclaré ({COUNTRIES[declared]}) : à reprendre."
        if (round(lat0, 6), round(lon0, 6)) == GENERIC_POINT:
            qc, detail = "Position générique", "Coordonnées communes à plusieurs sites de villes différentes (centre de Ouagadougou) : à géolocaliser."
        pt = Point(lon, lat)
        region = next((n for i, n, g in reg if i == iso and g.contains(pt)), None)
        prov = next((n for i, n, g in sub if i == iso and g.contains(pt)), None)
        if iso and region is None:  # point en bordure : polygone le plus proche
            region = min(((g.distance(pt), n) for i, n, g in reg if i == iso))[1]
        if iso and prov is None:
            prov = min(((g.distance(pt), n) for i, n, g in sub if i == iso))[1]
        act = str(r.get("Date d'activation") or "")
        try:
            act_iso = datetime.strptime(act, "%d/%m/%Y").date().isoformat()
        except ValueError:
            act_iso = ""
        rows.append({
            "id": int(r["NumDist"]), "nom": str(r["Nom"]).strip(), "pays": COUNTRIES.get(iso, str(r["Pays"]).title()), "iso": iso or "",
            "ville": str(r["Ville"]).strip() if pd.notna(r["Ville"]) else "", "adresse": str(r["Adresse"]).strip() if pd.notna(r["Adresse"]) else "",
            "lat": round(lat, 6), "lon": round(lon, 6), "lat_source": lat0, "lon_source": lon0,
            "zone": str(r["Zone 5"]) if pd.notna(r["Zone 5"]) else "", "secteur": str(r["Sectorisation"]) if pd.notna(r["Sectorisation"]) else "",
            "region": region or "", "province": prov or "", "region_declaree": str(r["Région"]).strip().title() if pd.notna(r["Région"]) else "",
            "activation": act_iso, "annee": int(act_iso[:4]) if act_iso else None, "parent": int(r["Numdist Parent"]) if pd.notna(r["Numdist Parent"]) else None,
            "statut": r["Statut"], "type": r["Type"], "canal": r["Canal de distribution"], "filiale": r["CANAL+"],
            "email": r["Email"] if pd.notna(r["Email"]) else "", "telephone": r["Téléphone"] if pd.notna(r["Téléphone"]) else "",
            "login": r["Login principal"] if pd.notna(r["Login principal"]) else "", "credit": r["Credit mode"] if pd.notna(r["Credit mode"]) else "",
            "qc": qc, "qc_detail": detail, "fichier_source": r["fichier_source"],
        })
    sites = pd.DataFrame(rows)
    # Sites partageant exactement les mêmes coordonnées (hors position générique déjà signalée)
    dup = sites[sites.duplicated(["lat", "lon"], keep=False) & (sites.qc == "OK")]
    for (la, lo), g in dup.groupby(["lat", "lon"]):
        others = ", ".join(str(x) for x in g.id)
        for i in g.index:
            sites.at[i, "qc"] = "Coordonnées partagées"
            sites.at[i, "qc_detail"] = f"Même position que les sites {others} : vérifier s'il s'agit d'un même lieu."
    # Voisin le plus proche (densité du réseau)
    la, lo = sites.lat.tolist(), sites.lon.tolist()
    nn = []
    for i in range(len(sites)):
        best = min((hav(la[i], lo[i], la[j], lo[j]) for j in range(len(sites)) if j != i and abs(la[i] - la[j]) < 1 and abs(lo[i] - lo[j]) < 1), default=None)
        nn.append(round(best, 2) if best is not None else None)
    sites["voisin_km"] = nn
    return sites


def basemap(src, a0):
    def fc(feats):
        return {"type": "FeatureCollection", "features": feats}

    def simp(g, tol):
        return json.loads(json.dumps(mapping(g.intersection(BBOX).simplify(tol, preserve_topology=True))), parse_float=lambda x: round(float(x), 3))

    near = {k: v.buffer(0.4) for k, v in a0.items()}
    zone = unary_union(list(near.values()))
    neigh = [f for f in load(src, "ne_10m_admin_0_countries.geojson")["features"] if shape(f["geometry"]).intersects(BBOX)]
    out = {
        "countries": fc([{"type": "Feature", "properties": {"iso": f["properties"]["ADM0_A3"], "name": f["properties"]["NAME_FR"]}, "geometry": simp(shape(f["geometry"]), 0.02)} for f in neigh]),
        "regions": fc([]), "subs": fc([]),
    }
    _, reg, sub = admin(src)
    out["regions"]["features"] = [{"type": "Feature", "properties": {"iso": i, "name": n, "label": [round(c, 3) for c in g.representative_point().coords[0]]}, "geometry": simp(g, 0.01)} for i, n, g in reg]
    out["subs"]["features"] = [{"type": "Feature", "properties": {"iso": i, "name": n}, "geometry": simp(g, 0.008)} for i, n, g in sub]
    roads = []
    for f in load(src, "ne_10m_roads.geojson")["features"]:
        if f["properties"].get("continent") != "Africa" or f["properties"].get("scalerank", 9) > 7:
            continue
        g = shape(f["geometry"])
        if g.intersects(zone):
            roads.append({"type": "Feature", "properties": {"major": f["properties"].get("scalerank", 9) <= 4}, "geometry": simp(g.intersection(zone), 0.01)})
    out["roads"] = fc(roads)
    out["rivers"] = fc([{"type": "Feature", "properties": {}, "geometry": simp(shape(f["geometry"]), 0.01)} for f in load(src, "rivers.geojson")["features"]
                        if f["properties"].get("featurecla", "").startswith("River") and shape(f["geometry"]).intersects(zone)])
    towns = []
    for f in load(src, "ne_10m_populated_places_simple.geojson")["features"]:
        p = f["properties"]
        if p["adm0_a3"] in COUNTRIES and (p["scalerank"] <= 6 or p["adm0cap"]):
            towns.append({"name": {"Timbuktu": "Tombouctou", "Bobo Dioulasso": "Bobo-Dioulasso"}.get(p["name"], p["name"]), "lat": round(p["latitude"], 4), "lon": round(p["longitude"], 4), "cap": bool(p["adm0cap"])})
    out["towns"] = towns
    return out


def mali_incidents():
    """Incidents de la base Mali déjà traitée dans ce dépôt (couche sécuritaire facultative)."""
    p = HERE.parent / "dist" / "incidents_mali.geojson"
    if not p.exists():
        return []
    out = []
    for f in json.loads(p.read_text(encoding="utf-8"))["features"]:
        q = f["properties"]
        out.append({"lon": f["geometry"]["coordinates"][0], "lat": f["geometry"]["coordinates"][1], "date": q["date"], "type": q["mode"],
                    "acteur": q["actor"], "description": f"{q['place']} — {q['desc']}", "tues": q["killed"], "source": "Base Mali 31/08–26/10/2026"})
    return out


def write_site(sites, geo, incidents):
    DIST.mkdir(exist_ok=True)
    pub = sites[PUBLIC].copy()
    pub["annee"] = pub["annee"].astype("Int64")
    records = json.loads(pub.to_json(orient="records", force_ascii=False))
    data = {"generated": date.today().isoformat(), "sites": records, "geo": geo, "incidents": incidents,
            "countries": [{"iso": k, "name": v, "n": int((sites.iso == k).sum())} for k, v in COUNTRIES.items()]}
    tpl = (HERE / "index.html").read_text(encoding="utf-8")
    frag = tpl.replace("/*__DATA__*/null", json.dumps(data, ensure_ascii=False, separators=(",", ":")))
    page = ('<!doctype html>\n<html lang="fr">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
            '<meta name="robots" content="noindex, nofollow">\n</head>\n<body>\n' + frag + "\n</body>\n</html>\n")
    (DIST / "carte_sites.html").write_text(page, encoding="utf-8")
    site = DIST / "site_netlify"
    shutil.rmtree(site, ignore_errors=True)
    site.mkdir(parents=True)
    (site / "index.html").write_text(page, encoding="utf-8")
    (site / "_headers").write_text("/*\n  X-Robots-Tag: noindex, nofollow\n  Referrer-Policy: no-referrer\n  X-Frame-Options: DENY\n", encoding="utf-8")
    (site / "robots.txt").write_text("User-agent: *\nDisallow: /\n", encoding="utf-8")
    shutil.make_archive(str(DIST / "Site_Netlify_Cartographie_PVA"), "zip", site)

    feats = [{"type": "Feature", "geometry": {"type": "Point", "coordinates": [r["lon"], r["lat"]]}, "properties": {k: r[k] for k in PUBLIC if k not in ("lat", "lon")}} for r in records]
    (DIST / "sites_pva.geojson").write_text(json.dumps({"type": "FeatureCollection", "features": feats}, ensure_ascii=False), encoding="utf-8")
    pm = "".join(f"<Placemark><name>{escape(str(r['id']))} · {escape(r['nom'])}</name><description>{escape(r['ville'])} · {escape(r['region'])} / {escape(r['province'])} · zone {escape(r['zone'])} · contrôle : {escape(r['qc'])}</description>"
                 f"<Point><coordinates>{r['lon']},{r['lat']},0</coordinates></Point></Placemark>" for r in records)
    (DIST / "sites_pva.kml").write_text(f'<?xml version="1.0" encoding="UTF-8"?>\n<kml xmlns="http://www.opengis.net/kml/2.2"><Document><name>Points de vente agréés</name>{pm}</Document></kml>\n', encoding="utf-8")


if __name__ == "__main__":
    src = Path(sys.argv[1])
    exports = sys.argv[2:]
    a0, reg, sub = admin(src)
    sites = qc_and_geocode(read_exports(exports), a0, reg, sub)
    geo = basemap(src, a0)
    write_site(sites, geo, mali_incidents())
    sites.to_pickle(DIST / "sites.pkl")
    print(sites.qc.value_counts().to_string())
    print(sites.groupby(["pays", "region"]).size().to_string())
    print("voisin médian (km)", sites.voisin_km.median())
