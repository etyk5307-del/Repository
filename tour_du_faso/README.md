# Tour du Faso 2026 — cartographie des linéaires

Carte interactive des 10 étapes et des transferts du 37e Tour du Faso (30 octobre – 8 novembre 2026),
d'après le document « Parcours » de la Fédération burkinabè de cyclisme.

- `dist/carte_tour_du_faso_2026.html` : carte autonome. Elle comprend la frise des jours, les corridors
  de 5, 10 et 20 km, une lecture animée du peloton et l'import d'incidents (CSV ou GeoJSON) pour
  compter les incidents par étape.
- `dist/tour_du_faso_2026.kml` (Google Earth), `.gpx` (GPS / smartphone) et `.geojson` (QGIS).

Les tracés sont reconstitués sur le réseau Natural Earth : ils ne remplacent pas le roadbook officiel.
Chaque étape indique son écart à la distance officielle et un niveau de fiabilité (axe confirmé, axe
probable, tracé indicatif). La position de Pa, Saponé et Dapélogo est approximative.

Régénérer : `python tour_du_faso/build_tdf.py <dossier_sources_geo>` (voir l'en-tête du script).
