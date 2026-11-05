"""Application de sûreté des sites (Burkina Faso, Mali, Niger) pour Netlify.

Assemble dans une seule page : les sites contrôlés par build_sites.py (dist/sites.pkl),
les limites pays / régions / provinces-cercles-départements, un réseau routier compact
(calcul d'itinéraire hors ligne), un répertoire de localités et la base d'incidents Mali.

Usage : python canal_sites/build_app.py <dossier_sources_geo>
Sorties : dist/app_netlify/ (index.html, _headers, robots.txt) et dist/App_Netlify_Surete_Sites.zip
"""
import json
import math
import shutil
import sys
from datetime import date
from pathlib import Path

import pandas as pd
from shapely.geometry import box, mapping, shape
from shapely.ops import unary_union

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_sites import BFA_REGION, COUNTRIES, MLI_FR, PUBLIC, mali_incidents  # noqa: E402

HERE = Path(__file__).resolve().parent
DIST = HERE / "dist"
VIEW = box(-17.5, 3.5, 20.5, 27.5)  # emprise du fond (pays voisins compris)
FR_NAMES = {"MRT": "MAURITANIE", "DZA": "ALGÉRIE", "LBY": "LIBYE", "TCD": "TCHAD", "NGA": "NIGERIA", "BEN": "BÉNIN", "TGO": "TOGO",
            "GHA": "GHANA", "CIV": "CÔTE D'IVOIRE", "GIN": "GUINÉE", "SEN": "SÉNÉGAL", "CMR": "CAMEROUN", "LBR": "LIBERIA", "SLE": "SIERRA LEONE",
            "GNB": "GUINÉE-BISSAU", "GMB": "GAMBIE", "ESH": "SAHARA OCC.", "MAR": "MAROC", "BFA": "BURKINA FASO", "MLI": "MALI", "NER": "NIGER"}
LABEL_AT = {"MRT": (-10.5, 19.5), "DZA": (1.5, 25.0), "LBY": (17.0, 25.5), "TCD": (18.0, 15.5), "NGA": (8.0, 9.5), "BEN": (2.3, 9.6), "TGO": (1.05, 7.4),
            "GHA": (-1.7, 8.6), "CIV": (-6.2, 7.0), "GIN": (-10.8, 10.6), "SEN": (-14.8, 14.6), "MLI": (-3.5, 17.6), "NER": (9.0, 17.6)}


def load(src, name):
    return json.loads((src / name).read_text(encoding="utf-8"))


def rnd(geom, nd=3):
    def r(c):
        return [round(c[0], nd), round(c[1], nd)] if isinstance(c[0], (int, float)) else [r(x) for x in c]
    g = mapping(geom)
    return {"type": g["type"], "coordinates": r(g["coordinates"])}


def hav(a, b):
    lo1, la1, lo2, la2 = map(math.radians, (a[0], a[1], b[0], b[1]))
    h = math.sin((la2 - la1) / 2) ** 2 + math.cos(la1) * math.cos(la2) * math.sin((lo2 - lo1) / 2) ** 2
    return 2 * 6371 * math.asin(math.sqrt(h))


def geo_layers(src):
    countries, labels = [], []
    for f in load(src, "ne_10m_admin_0_countries.geojson")["features"]:
        g, iso = shape(f["geometry"]), f["properties"]["ADM0_A3"]
        if not g.intersects(VIEW):
            continue
        countries.append({"type": "Feature", "properties": {"iso": iso, "focus": iso in COUNTRIES}, "geometry": rnd(g.intersection(VIEW).simplify(0.012, preserve_topology=True))})
        if iso in LABEL_AT:
            labels.append({"iso": iso, "name": FR_NAMES[iso], "lon": LABEL_AT[iso][0], "lat": LABEL_AT[iso][1], "focus": iso in COUNTRIES})
    regions, subs = [], []
    for iso, fn, rename in (("BFA", "bfa_ADM1.geojson", BFA_REGION), ("MLI", "gb_ADM1.geojson", MLI_FR)):
        for f in load(src, fn)["features"]:
            g = shape(f["geometry"])
            regions.append({"type": "Feature", "properties": {"iso": iso, "name": rename.get(f["properties"]["shapeName"], f["properties"]["shapeName"])},
                            "geometry": rnd(g.simplify(0.006, preserve_topology=True))})
    for f in load(src, "ne_10m_admin_1_states_provinces.geojson")["features"]:
        if f["properties"]["adm0_a3"] == "NER":
            regions.append({"type": "Feature", "properties": {"iso": "NER", "name": f["properties"]["name"]}, "geometry": rnd(shape(f["geometry"]).simplify(0.01, preserve_topology=True))})
    for iso, fn in (("BFA", "bfa_ADM2.geojson"), ("MLI", "gb_ADM2.geojson"), ("NER", "ner_ADM2.geojson")):
        for f in load(src, fn)["features"]:
            subs.append({"type": "Feature", "properties": {"iso": iso, "name": MLI_FR.get(f["properties"]["shapeName"], f["properties"]["shapeName"])},
                         "geometry": rnd(shape(f["geometry"]).simplify(0.006, preserve_topology=True))})
    focus = unary_union([shape(f["geometry"]) for f in load(src, "ne_10m_admin_0_countries.geojson")["features"] if f["properties"]["ADM0_A3"] in COUNTRIES])
    zone = focus.buffer(0.3)
    rivers = [{"type": "Feature", "properties": {}, "geometry": rnd(shape(f["geometry"]).intersection(VIEW).simplify(0.01))}
              for f in load(src, "rivers.geojson")["features"] if f["properties"].get("featurecla", "").startswith("River") and shape(f["geometry"]).intersects(zone)]
    lakes = [{"type": "Feature", "properties": {}, "geometry": rnd(shape(f["geometry"]).simplify(0.008))}
             for f in load(src, "ne_10m_lakes.geojson")["features"] if shape(f["geometry"]).intersects(zone)]
    return countries, labels, regions, subs, rivers, lakes, zone


def road_network(src, zone):
    """Routes Natural Earth de la zone, simplifiées, et graphe compact pour le calcul d'itinéraire hors ligne."""
    feats, adj = [], {}
    key = lambda c: (round(c[0], 3), round(c[1], 3))  # noqa: E731
    ends = []
    for f in load(src, "ne_10m_roads.geojson")["features"]:
        p = f["properties"]
        if p.get("continent") != "Africa":
            continue
        g = shape(f["geometry"])
        if not g.intersects(zone):
            continue
        g = g.intersection(zone)
        major = p.get("scalerank", 9) <= 4 or p.get("type") == "Major Highway"
        feats.append({"type": "Feature", "properties": {"major": major}, "geometry": rnd(g.simplify(0.004))})
        for part in ([g] if g.geom_type == "LineString" else [x for x in getattr(g, "geoms", []) if x.geom_type == "LineString"]):
            cs = [key(c) for c in part.coords]
            ends += [cs[0], cs[-1]]
            for a, b in zip(cs, cs[1:]):
                if a != b:
                    adj.setdefault(a, {})[b] = adj.setdefault(b, {})[a] = hav(a, b)
    nodes = list(adj)
    for e in ends:  # raccorde les tronçons non jointifs (< 3 km)
        best = min(((hav(e, n), n) for n in nodes if n != e and n not in adj[e] and abs(n[0] - e[0]) < .05 and abs(n[1] - e[1]) < .05), default=None)
        if best and best[0] < 3:
            adj[e][best[1]] = adj[best[1]][e] = best[0]
    # ne garde que la plus grande composante connexe
    seen, comps = set(), []
    for n in nodes:
        if n in seen:
            continue
        comp, stack = [], [n]
        seen.add(n)
        while stack:
            u = stack.pop()
            comp.append(u)
            for v in adj[u]:
                if v not in seen:
                    seen.add(v)
                    stack.append(v)
        comps.append(comp)
    main = set(max(comps, key=len))
    idx = {n: i for i, n in enumerate(sorted(main))}
    edges = sorted({(min(idx[a], idx[b]), max(idx[a], idx[b]), round(w, 2)) for a in main for b, w in adj[a].items() if b in main})
    return feats, {"nodes": [list(n) for n in sorted(main)], "edges": [list(e) for e in edges]}


def gazetteer(src, sites):
    out = {}
    for f in load(src, "ne_10m_populated_places_simple.geojson")["features"]:
        p = f["properties"]
        if p["adm0_a3"] in COUNTRIES:
            name = {"Timbuktu": "Tombouctou", "Bobo Dioulasso": "Bobo-Dioulasso"}.get(p["name"], p["name"])
            out[(name.upper(), p["adm0_a3"])] = {"name": name, "iso": p["adm0_a3"], "lat": round(p["latitude"], 4), "lon": round(p["longitude"], 4),
                                                 "rank": 1 if p["adm0cap"] else (2 if p["scalerank"] <= 5 else 3), "src": "ne"}
    ok = sites[sites.qc == "OK"].copy()
    ok["base"] = ok.ville.str.split("/").str[0].str.strip().str.title()
    for (base, iso), g in ok.groupby(["base", "iso"]):
        k = (base.upper(), iso)
        if base and k not in out:
            out[k] = {"name": base, "iso": iso, "lat": round(g.lat.median(), 4), "lon": round(g.lon.median(), 4), "rank": 4, "src": "sites"}
    return sorted(out.values(), key=lambda t: (t["rank"], t["name"]))


def build(src):
    sites = pd.read_pickle(DIST / "sites.pkl")
    countries, labels, regions, subs, rivers, lakes, zone = geo_layers(src)
    roads, graph = road_network(src, zone)
    pub = sites[PUBLIC + ["iso"]].copy()
    pub["annee"] = pub["annee"].astype("Int64")
    data = {
        "generated": date.today().isoformat(),
        "sites": json.loads(pub.to_json(orient="records", force_ascii=False)),
        "incidents_mali": mali_incidents(),
        "geo": {"countries": {"type": "FeatureCollection", "features": countries}, "labels": labels,
                "regions": {"type": "FeatureCollection", "features": regions}, "subs": {"type": "FeatureCollection", "features": subs},
                "roads": {"type": "FeatureCollection", "features": roads}, "rivers": {"type": "FeatureCollection", "features": rivers},
                "lakes": {"type": "FeatureCollection", "features": lakes}},
        "graph": graph, "places": gazetteer(src, sites),
    }
    return data


def write(data):
    tpl = (HERE / "app.html").read_text(encoding="utf-8")
    body = tpl.replace("/*__DATA__*/null", json.dumps(data, ensure_ascii=False, separators=(",", ":")))
    page = ('<!doctype html>\n<html lang="fr">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
            '<meta name="robots" content="noindex, nofollow">\n</head>\n<body>\n' + body + "\n</body>\n</html>\n")
    out = DIST / "app_netlify"
    shutil.rmtree(out, ignore_errors=True)
    out.mkdir(parents=True)
    (out / "index.html").write_text(page, encoding="utf-8")
    (out / "_headers").write_text("/*\n  X-Robots-Tag: noindex, nofollow\n  Referrer-Policy: no-referrer\n  X-Frame-Options: DENY\n", encoding="utf-8")
    (out / "robots.txt").write_text("User-agent: *\nDisallow: /\n", encoding="utf-8")
    shutil.make_archive(str(DIST / "App_Netlify_Surete_Sites"), "zip", out)
    return out / "index.html"


if __name__ == "__main__":
    d = build(Path(sys.argv[1]))
    p = write(d)
    print(p, f"{p.stat().st_size / 1024:.0f} Ko", "sites", len(d["sites"]), "noeuds", len(d["graph"]["nodes"]), "arêtes", len(d["graph"]["edges"]), "localités", len(d["places"]))
