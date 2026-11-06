"""Prépare le fond cartographique embarqué du tableau de bord (frontières, régions,
cercles, fleuves, routes, localités) à partir de sources ouvertes.

Sources :
  - geoBoundaries gbOpen MLI ADM1 / ADM2 (CC BY 4.0)
  - Natural Earth 10m : pays, fleuves, lacs, routes, localités (domaine public)

Usage : python scripts/prepare_geo.py <dossier_sources>  -> data/geo/basemap.json
"""
import json
import sys
from pathlib import Path

from shapely.geometry import box, mapping, shape

ROOT = Path(__file__).resolve().parents[1]
BBOX = box(-17.5, 6.5, 8.5, 26.5)
NEIGHBOURS = {"MLI", "BFA", "NER", "DZA", "MRT", "SEN", "GIN", "CIV", "GHA", "TGO", "BEN", "SLE", "LBR", "GMB", "GNB", "NGA", "LBY", "MAR", "ESH", "TCD"}

# Localités clés absentes ou mal placées dans Natural Earth (coordonnées WGS84).
EXTRA_TOWNS = [
    ("Kidal", 18.4411, 1.4078, 2), ("Sévaré", 14.5333, -4.0960, 2), ("Douentza", 15.0019, -2.9497, 3),
    ("Niono", 14.2526, -5.9930, 3), ("Bandiagara", 14.3500, -3.6100, 3), ("Ansongo", 15.6597, 0.5022, 3),
    ("Ouagadougou", 12.3714, -1.5197, 1), ("Niamey", 13.5116, 2.1254, 1), ("Tinzaouatène", 19.9550, 2.9650, 3),
    ("Diabaly", 14.6870, -6.0160, 3), ("Ménaka", 15.9182, 2.4022, 2), ("Gao", 16.2717, -0.0447, 2),
]


def rnd(geom, nd):
    def r(c):
        if isinstance(c[0], (int, float)):
            return [round(c[0], nd), round(c[1], nd)]
        return [r(x) for x in c]
    g = mapping(geom)
    return {"type": g["type"], "coordinates": r(g["coordinates"])}


def load(p):
    return json.loads(Path(p).read_text(encoding="utf-8"))


def poly_layer(features, tol, nd, props):
    out = []
    for f in features:
        g = shape(f["geometry"])
        if not g.intersects(BBOX):
            continue
        g = g.intersection(BBOX).simplify(tol, preserve_topology=True)
        if g.is_empty:
            continue
        out.append({"type": "Feature", "properties": props(f["properties"]), "geometry": rnd(g, nd)})
    return {"type": "FeatureCollection", "features": out}


def main(src):
    src = Path(src)
    countries = poly_layer(
        [f for f in load(src / "ne_10m_admin_0_countries.geojson")["features"] if f["properties"]["ADM0_A3"] in NEIGHBOURS],
        0.03, 3, lambda p: {"iso": p["ADM0_A3"], "name": p["NAME_FR"] or p["NAME"]},
    )
    adm1 = poly_layer(load(src / "gb_ADM1.geojson")["features"], 0.008, 3, lambda p: {"name": p["shapeName"]})
    adm2 = poly_layer(load(src / "gb_ADM2.geojson")["features"], 0.008, 3, lambda p: {"name": p["shapeName"]})

    rivers = []
    for f in load(src / "rivers.geojson")["features"]:
        g = shape(f["geometry"])
        if f["properties"].get("featurecla", "").startswith("River") and g.intersects(BBOX):
            g = g.intersection(BBOX).simplify(0.01)
            if not g.is_empty:
                rivers.append({"type": "Feature", "properties": {"name": f["properties"].get("name") or "", "rank": f["properties"].get("scalerank")}, "geometry": rnd(g, 3)})
    lakes = poly_layer(load(src / "ne_10m_lakes.geojson")["features"], 0.005, 3, lambda p: {"name": p.get("name") or ""})

    mali = shape(next(f for f in load(src / "ne_10m_admin_0_countries.geojson")["features"] if f["properties"]["ADM0_A3"] == "MLI")["geometry"]).buffer(0.6)
    roads = []
    for f in load(src / "ne_10m_roads.geojson")["features"]:
        p = f["properties"]
        if p.get("continent") != "Africa" or p.get("scalerank", 99) > 7:
            continue
        g = shape(f["geometry"])
        if g.intersects(mali):
            g = g.intersection(mali).simplify(0.01)
            if not g.is_empty:
                roads.append({"type": "Feature", "properties": {"major": p.get("type") == "Major Highway" or p.get("scalerank", 9) <= 4}, "geometry": rnd(g, 3)})

    towns = {}
    for f in load(src / "ne_10m_populated_places_simple.geojson")["features"]:
        p = f["properties"]
        if p["adm0_a3"] == "MLI":
            name = {"Timbuktu": "Tombouctou"}.get(p["name"], p["name"])
            towns[name] = {"name": name, "lat": round(p["latitude"], 4), "lon": round(p["longitude"], 4), "rank": 1 if name == "Bamako" else (2 if p["scalerank"] <= 4 else 3)}
    for name, lat, lon, rank in EXTRA_TOWNS:
        towns[name] = {"name": name, "lat": lat, "lon": lon, "rank": rank}

    out = {
        "attribution": "geoBoundaries gbOpen (CC BY 4.0) · Natural Earth (domaine public)",
        "countries": countries, "adm1": adm1, "adm2": adm2,
        "rivers": {"type": "FeatureCollection", "features": rivers},
        "lakes": lakes,
        "roads": {"type": "FeatureCollection", "features": roads},
        "towns": sorted(towns.values(), key=lambda t: t["rank"]),
    }
    dest = ROOT / "data" / "geo" / "basemap.json"
    dest.write_text(json.dumps(out, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"{dest} : {dest.stat().st_size / 1024:.0f} Ko · {len(roads)} routes · {len(rivers)} fleuves · {len(towns)} localités")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "geo_sources")
