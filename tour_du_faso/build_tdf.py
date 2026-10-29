"""Cartographie des linéaires du 37e Tour du Faso (30 octobre – 8 novembre 2026).

Source : document « Parcours » de la Fédération burkinabè de cyclisme (FBC).
Les tracés sont reconstitués sur le réseau routier Natural Earth (plus court chemin) ;
quand le réseau est incomplet, le tronçon est tracé en ligne droite entre localités
et signalé « indicatif ». Chaque étape compare la longueur reconstituée à la distance
officielle. Les circuits urbains sont schématiques (boucle de même longueur).

Usage : python tour_du_faso/build_tdf.py <dossier_sources_geo>
  (dossier contenant ne_10m_roads.geojson, ne_10m_admin_0_countries.geojson,
   ne_10m_rivers_lake_centerlines.geojson -> rivers.geojson, bfa_ADM1.geojson, bfa_ADM2.geojson)
Sorties : tour_du_faso/dist/
"""
import heapq
import json
import math
import sys
from pathlib import Path
from xml.sax.saxutils import escape

from shapely.geometry import LineString, Point, box, mapping, shape

HERE = Path(__file__).resolve().parent
DIST = HERE / "dist"
BBOX = box(-6.2, 8.8, 3.0, 15.6)

# Anciennes régions (geoBoundaries) -> régions issues de la réforme de 2025 (noms du document FBC).
REGION_NEW = {
    "Centre": "Kadiogo", "Centre-Nord": "Koulsé", "Centre-Ouest": "Nando", "Boucle du Mouhoun": "Bankui",
    "Hauts-Bassins": "Guiriko", "Centre-Est": "Nakanbé", "Centre-Sud": "Nazinon", "Plateau Central": "Oubri",
    "Cascades": "Tannounyan", "Est": "Goulmou", "Nord": "Yaadga", "Sahel": "Liptako", "Sud-Ouest": "Djôrô",
}

# Localités (WGS84). « approx » : position non vérifiée sur une source cartographique, à confirmer.
TOWNS = {
    "Ouagadougou": (12.3714, -1.5197, False), "Kaya": (13.0900, -1.0900, False), "Koudougou": (12.2500, -2.3700, False),
    "Boromo": (11.7454, -2.9306, False), "Bobo-Dioulasso": (11.1771, -4.2979, False), "Pa": (11.5480, -3.2630, True),
    "Saponé": (12.0475, -1.6036, True), "Léo": (11.0940, -2.0980, False), "Pô": (11.1690, -1.1350, False),
    "Koupéla": (12.1770, -0.3560, False), "Tenkodogo": (11.7800, -0.3700, False), "Ziniaré": (12.5770, -1.2930, False),
    "Dapélogo": (12.6830, -1.6830, True), "Sapouy": (11.5540, -1.7740, False), "Zorgho": (12.2430, -0.6110, False),
}

# Étapes : waypoints (localités traversées par l'axe le plus probable), distance officielle, transferts, hébergement.
STAGES = [
    dict(n=1, date="2026-10-30", jour="Vendredi 30 octobre", start="Ouagadougou", end="Kaya", via=["Ouagadougou", "Ziniaré", "Kaya"], km=149.82,
         t_before=None, t_after=("Kaya", "Ouagadougou", 100), hotel="Ouagadougou", note="Présentation des équipes avant le départ."),
    dict(n=2, date="2026-10-31", jour="Samedi 31 octobre", start="Ouagadougou", end="Koudougou", via=["Ouagadougou", "Koudougou"], km=129.75,
         t_before=None, t_after=None, hotel="Koudougou"),
    dict(n=3, date="2026-11-01", jour="Dimanche 1er novembre", start="Koudougou", end="Boromo", via=["Koudougou", "Boromo"], km=158.90,
         t_before=None, t_after=("Boromo", "Bobo-Dioulasso", 185), hotel="Bobo-Dioulasso"),
    dict(n=4, date="2026-11-02", jour="Lundi 2 novembre", start="Bobo-Dioulasso", end="Bobo-Dioulasso", via=[], km=116.10,
         circuit=("Bobo-Dioulasso", 11.61, 10), t_before=None, t_after=None, hotel="Bobo-Dioulasso", note="Étape en circuit : 11,61 km × 10 tours."),
    dict(n=5, date="2026-11-03", jour="Mardi 3 novembre", start="Bobo-Dioulasso", end="Pa", via=["Bobo-Dioulasso", "Pa"], km=131.32,
         t_before=None, t_after=("Pa", "Ouagadougou", 239), hotel="Ouagadougou"),
    dict(n=6, date="2026-11-04", jour="Mercredi 4 novembre", start="Saponé", end="Léo", via=["Saponé", "Sapouy", "Léo"], km=138.11,
         t_before=("Ouagadougou", "Saponé", 27), t_after=("Léo", "Ouagadougou", 167), hotel="Ouagadougou"),
    dict(n=7, date="2026-11-05", jour="Jeudi 5 novembre", start="Ouagadougou", end="Pô", via=["Ouagadougou", "Pô"], km=144.02,
         t_before=None, t_after=("Pô", "Ouagadougou", 145), hotel="Ouagadougou"),
    dict(n=8, date="2026-11-06", jour="Vendredi 6 novembre", start="Ouagadougou", end="Koupéla", via=["Ouagadougou", "Koupéla"], km=153.05,
         t_before=None, t_after=("Koupéla", "Tenkodogo", 45), hotel="Tenkodogo"),
    dict(n=9, date="2026-11-07", jour="Samedi 7 novembre", start="Tenkodogo", end="Ziniaré", via=["Tenkodogo", "Koupéla", "Zorgho", "Ziniaré"], km=167.53,
         t_before=None, t_after=("Ziniaré", "Ouagadougou", 35), hotel="Ouagadougou"),
    dict(n=10, date="2026-11-08", jour="Dimanche 8 novembre", start="Dapélogo", end="Ouagadougou", via=["Dapélogo", "Ouagadougou"], km=111.37,
         circuit=("Ouagadougou", 14.25, 4), t_before=("Ouagadougou", "Dapélogo", 45), t_after=None, hotel="Ouagadougou",
         note="54,37 km en ligne puis circuit final de 14,25 km × 4 tours à Ouagadougou."),
]
OFFICIAL = {"course": 1399.97, "transfert_avant": 72, "transfert_apres": 916, "transfert": 988, "total": 2387.97}


def hav(a, b):
    """Distance (km) entre deux points (lon, lat)."""
    lo1, la1, lo2, la2 = map(math.radians, (a[0], a[1], b[0], b[1]))
    h = math.sin((la2 - la1) / 2) ** 2 + math.cos(la1) * math.cos(la2) * math.sin((lo2 - lo1) / 2) ** 2
    return 2 * 6371 * math.asin(math.sqrt(h))


def length(coords):
    return sum(hav(a, b) for a, b in zip(coords, coords[1:]))


class RoadGraph:
    def __init__(self, roads):
        self.adj, ends = {}, []
        key = lambda c: (round(c[0], 3), round(c[1], 3))  # noqa: E731
        for f in roads["features"]:
            g = shape(f["geometry"])
            if not g.intersects(BBOX):
                continue
            for part in ([g] if g.geom_type == "LineString" else list(g.geoms)):
                cs = [key(c) for c in part.coords]
                ends += [cs[0], cs[-1]]
                for a, b in zip(cs, cs[1:]):
                    d = hav(a, b)
                    self.adj.setdefault(a, {})[b] = d
                    self.adj.setdefault(b, {})[a] = d
        self.nodes = list(self.adj)
        for e in ends:  # raccorde les extrémités de tronçons non jointifs (< 3 km)
            best = min(((hav(e, n), n) for n in self.nodes if n != e and n not in self.adj[e]), default=None)
            if best and best[0] < 3:
                self.adj[e][best[1]] = self.adj[best[1]][e] = best[0]

    def route(self, a, b):
        snap = lambda p: min(self.nodes, key=lambda n: hav(n, p))  # noqa: E731
        s, t = snap(a), snap(b)
        dist, prev, pq = {s: 0}, {}, [(0, s)]
        while pq:
            d, u = heapq.heappop(pq)
            if u == t:
                break
            if d > dist[u]:
                continue
            for v, w in self.adj[u].items():
                if d + w < dist.get(v, 1e18):
                    dist[v], prev[v] = d + w, u
                    heapq.heappush(pq, (d + w, v))
        if t not in dist:
            return None
        path = [t]
        while path[-1] != s:
            path.append(prev[path[-1]])
        return [list(a)] + [list(p) for p in reversed(path)] + [list(b)]

    def leg(self, a, b):
        """Tronçon routier si le réseau le permet sans détour aberrant, sinon ligne droite (indicatif)."""
        crow = hav(a, b)
        p = self.route(a, b)
        if p and length(p) <= 1.35 * crow + 5:
            return p, True
        return [list(a), list(b)], False


def corridor(coords, km):
    """Polygone tampon de km autour d'une ligne (projection locale équirectangulaire, précise à l'échelle du pays)."""
    kx, ky = 111.32 * math.cos(math.radians(12)), 110.6
    g = LineString([(x * kx, y * ky) for x, y in coords]).buffer(km, 12).simplify(0.4)
    ext = [[round(x / kx, 4), round(y / ky, 4)] for x, y in g.exterior.coords]
    return [ext] + [[[round(x / kx, 4), round(y / ky, 4)] for x, y in r.coords] for r in g.interiors]


def ll(name):
    lat, lon, _ = TOWNS[name]
    return (lon, lat)


def circuit(name, km, n=72):
    """Boucle schématique de circonférence km, tangente au centre-ville."""
    lon0, lat0 = ll(name)
    r = km / (2 * math.pi)
    cx, cy = lon0 + r / (111.32 * math.cos(math.radians(lat0))), lat0
    return [[round(cx + r / (111.32 * math.cos(math.radians(lat0))) * math.cos(math.pi + 2 * math.pi * k / n), 5),
             round(cy + r / 111.32 * math.sin(math.pi + 2 * math.pi * k / n), 5)] for k in range(n + 1)]


def build(src):
    src = Path(src)
    load = lambda f: json.loads((src / f).read_text(encoding="utf-8"))  # noqa: E731
    graph = RoadGraph(load("ne_10m_roads.geojson"))
    adm1 = [(f["properties"]["shapeName"], shape(f["geometry"])) for f in load("bfa_ADM1.geojson")["features"]]
    countries = [f for f in load("ne_10m_admin_0_countries.geojson")["features"]
                 if f["properties"]["ADM0_A3"] in {"BFA", "MLI", "NER", "BEN", "TGO", "GHA", "CIV"}]
    bfa = shape(next(f for f in countries if f["properties"]["ADM0_A3"] == "BFA")["geometry"])
    neighbours = {f["properties"]["NAME_FR"]: shape(f["geometry"]) for f in countries if f["properties"]["ADM0_A3"] != "BFA"}

    def analyse(coords):
        line = LineString(coords)
        regions = [REGION_NEW.get(n, n) for n, g in adm1 if g.intersects(line)]
        dmin = min(((round(g.distance(line) * 111), n) for n, g in neighbours.items()))
        return regions, dmin

    stages = []
    for s in STAGES:
        line_off = s["km"] - (s["circuit"][1] * s["circuit"][2] if s.get("circuit") else 0)
        # 1) itinéraire routier complet, retenu s'il reste cohérent avec la distance officielle ;
        # 2) sinon, tronçon par tronçon (route si pas de détour aberrant, ligne droite sinon).
        coords, routed = [], []
        legs = [graph.route(ll(a), ll(b)) for a, b in zip(s["via"], s["via"][1:])]
        if legs and all(legs) and sum(length(p) for p in legs) <= 1.10 * line_off + 5:
            for p in legs:
                coords += p if not coords else p[1:]
            routed = [True] * len(legs)
        else:
            for a, b in zip(s["via"], s["via"][1:]):
                p, ok = graph.leg(ll(a), ll(b))
                coords += p if not coords else p[1:]
                routed.append(ok)
        line_km = length(coords) if coords else 0
        loop = None
        if s.get("circuit"):
            town, lap, laps = s["circuit"]
            loop = circuit(town, lap)
        km_est = line_km + (s["circuit"][1] * s["circuit"][2] if s.get("circuit") else 0)
        geom = coords or loop
        regions, (dmin, nb) = analyse(geom)
        ratio = km_est / s["km"]
        if not coords:
            quality = "Circuit schématique"
        elif all(routed) and abs(ratio - 1) <= 0.08:
            quality = "Axe confirmé"
        elif all(routed):
            quality = "Axe probable, détours non documentés"
        else:
            quality = "Tracé indicatif"

        def transfer(t):
            if not t:
                return None
            p, ok = graph.leg(ll(t[0]), ll(t[1]))
            return {"from": t[0], "to": t[1], "km": t[2], "coords": [[round(x, 4), round(y, 4)] for x, y in p], "routed": ok}

        stages.append({
            "n": s["n"], "date": s["date"], "jour": s["jour"], "start": s["start"], "end": s["end"], "km": s["km"],
            "km_est": round(km_est, 1), "ecart": round(100 * (ratio - 1)), "quality": quality,
            "coords": [[round(x, 4), round(y, 4)] for x, y in coords], "loop": loop,
            "circuit": {"town": s["circuit"][0], "lap": s["circuit"][1], "laps": s["circuit"][2]} if s.get("circuit") else None,
            "via": s["via"], "regions": regions, "border_km": dmin, "border": nb,
            "t_before": transfer(s["t_before"]), "t_after": transfer(s["t_after"]), "hotel": s["hotel"], "note": s.get("note", ""),
            "duration_h": round(s["km"] / 40, 1),
            "corridors": {str(w): corridor(geom if not (coords and loop) else coords + loop, w) for w in (5, 10, 20)},
        })

    def fc(features):
        return {"type": "FeatureCollection", "features": features}

    def simp(g, tol):
        return json.loads(json.dumps(mapping(g.intersection(BBOX).simplify(tol, preserve_topology=True))), parse_float=lambda x: round(float(x), 3))

    zone_roads, zone_rivers = bfa.buffer(0.3), bfa.buffer(0.2)
    near = lambda f, zone: (lambda g: g.intersects(BBOX) and g.intersects(zone))(shape(f["geometry"]))  # noqa: E731
    basemap = {
        "countries": fc([{"type": "Feature", "properties": {"iso": f["properties"]["ADM0_A3"], "name": f["properties"]["NAME_FR"]},
                          "geometry": simp(shape(f["geometry"]), 0.02)} for f in countries]),
        "regions": fc([{"type": "Feature", "properties": {"old": n, "name": REGION_NEW.get(n, n), "label": [round(c, 3) for c in g.representative_point().coords[0]]},
                         "geometry": simp(g, 0.006)} for n, g in adm1]),
        "provinces": fc([{"type": "Feature", "properties": {"name": f["properties"]["shapeName"]}, "geometry": simp(shape(f["geometry"]), 0.006)}
                         for f in load("bfa_ADM2.geojson")["features"]]),
        "roads": fc([{"type": "Feature", "properties": {"major": f["properties"].get("scalerank", 9) <= 4},
                      "geometry": simp(shape(f["geometry"]), 0.005)} for f in load("ne_10m_roads.geojson")["features"]
                     if near(f, zone_roads)]),
        "rivers": fc([{"type": "Feature", "properties": {"name": f["properties"].get("name") or ""}, "geometry": simp(shape(f["geometry"]), 0.008)}
                      for f in load("rivers.geojson")["features"]
                      if f["properties"].get("featurecla", "").startswith("River") and near(f, zone_rivers)]),
    }
    towns = [{"name": k, "lat": v[0], "lon": v[1], "approx": v[2]} for k, v in TOWNS.items()]
    nights = {}
    for s in stages:
        nights[s["hotel"]] = nights.get(s["hotel"], 0) + 1
    nights["Ouagadougou"] = nights.get("Ouagadougou", 0) + 1  # nuit du 29 octobre (enregistrement des équipes)
    return {"stages": stages, "towns": towns, "nights": nights, "official": OFFICIAL, "basemap": basemap}


def write(data):
    DIST.mkdir(exist_ok=True)
    tpl = (HERE / "index.html").read_text(encoding="utf-8")
    frag = tpl.replace("/*__DATA__*/null", json.dumps(data, ensure_ascii=False, separators=(",", ":")))
    (DIST / "tdf_artifact.html").write_text(frag, encoding="utf-8")
    (DIST / "carte_tour_du_faso_2026.html").write_text(
        '<!doctype html>\n<html lang="fr">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n</head>\n<body>\n'
        + frag + "\n</body>\n</html>\n", encoding="utf-8")
    feats = []
    for s in data["stages"]:
        props = {k: s[k] for k in ("n", "date", "start", "end", "km", "km_est", "ecart", "quality", "hotel", "border_km", "border")}
        props["regions"] = ", ".join(s["regions"])
        if s["coords"]:
            feats.append({"type": "Feature", "properties": {**props, "type": "étape"}, "geometry": {"type": "LineString", "coordinates": s["coords"]}})
        if s["loop"]:
            feats.append({"type": "Feature", "properties": {**props, "type": "circuit (schématique)"}, "geometry": {"type": "LineString", "coordinates": s["loop"]}})
        for k in ("t_before", "t_after"):
            if s[k]:
                feats.append({"type": "Feature", "properties": {"n": s["n"], "type": "transfert", "from": s[k]["from"], "to": s[k]["to"], "km": s[k]["km"]},
                              "geometry": {"type": "LineString", "coordinates": s[k]["coords"]}})
    for t in data["towns"]:
        feats.append({"type": "Feature", "properties": {"type": "localité", "name": t["name"], "position_approx": t["approx"]},
                      "geometry": {"type": "Point", "coordinates": [t["lon"], t["lat"]]}})
    (DIST / "tour_du_faso_2026.geojson").write_text(json.dumps({"type": "FeatureCollection", "features": feats}, ensure_ascii=False, indent=1), encoding="utf-8")

    colors = ["ff1c7ce3", "ff3c9a1f", "ffb4781f", "ff2c2cd6", "ffbd67a9", "ff4b568c", "ffc277e3", "ff7f7f7f", "ff22bdbc", "ffcfbe17"]
    pm = []
    for s in data["stages"]:
        geom = s["coords"] or s["loop"]
        desc = f"{s['jour']} · {s['km']} km officiels · tracé {s['quality'].lower()} ({s['km_est']} km) · hébergement {s['hotel']}"
        pm.append(f"<Placemark><name>Étape {s['n']} · {escape(s['start'])} – {escape(s['end'])}</name><description>{escape(desc)}</description>"
                  f"<TimeStamp><when>{s['date']}</when></TimeStamp><Style><LineStyle><color>{colors[s['n'] - 1]}</color><width>4</width></LineStyle></Style>"
                  f"<LineString><tessellate>1</tessellate><coordinates>{' '.join(f'{x},{y},0' for x, y in geom)}</coordinates></LineString></Placemark>")
    (DIST / "tour_du_faso_2026.kml").write_text('<?xml version="1.0" encoding="UTF-8"?>\n<kml xmlns="http://www.opengis.net/kml/2.2"><Document><name>37e Tour du Faso 2026</name>'
                                               + "".join(pm) + "</Document></kml>\n", encoding="utf-8")
    trk = "".join(f"<trk><name>Étape {s['n']} {escape(s['start'])} - {escape(s['end'])}</name><trkseg>"
                  + "".join(f'<trkpt lat="{y}" lon="{x}"/>' for x, y in (s["coords"] or s["loop"])) + "</trkseg></trk>" for s in data["stages"])
    (DIST / "tour_du_faso_2026.gpx").write_text('<?xml version="1.0" encoding="UTF-8"?>\n<gpx version="1.1" creator="Observatoire" xmlns="http://www.topografix.com/GPX/1/1">'
                                               + trk + "</gpx>\n", encoding="utf-8")


if __name__ == "__main__":
    d = build(sys.argv[1] if len(sys.argv) > 1 else "geo_sources")
    write(d)
    for s in d["stages"]:
        print(s["n"], s["start"], s["end"], s["km"], s["km_est"], f"{s['ecart']:+d}%", s["quality"], s["regions"], s["border_km"], s["border"])
    print(d["nights"])
