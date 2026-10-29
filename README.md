# Observatoire sécuritaire Mali

Tableau de bord décisionnel et cartographie dynamique des incidents sécuritaires au Mali
(base « Base_Donnees_Securite_Mali_OpenMap_Dashboard.xlsx », 31 août – 26 octobre 2026).

## Ouvrir le tableau de bord

Ouvrir `dist/tableau_de_bord_mali.html` dans un navigateur récent, avec un accès Internet
pour charger deck.gl, ECharts et les polices. Les données et le fond cartographique sont embarqués.

Fonctions :
- filtres croisés : période, acteur, mode opératoire, région, fiabilité, victimes civiles ;
- indicateurs clés avec comparaison P1 (31/08–27/09) / P2 (28/09–26/10) ;
- carte WebGL : incidents, chaleur, hexagones 3D, indice de menace par région, cercles, routes, hydrographie, rayons autour de Bamako, lecture chronologique animée ;
- dossier détaillé de chaque incident avec ses sources typées ;
- tendances : rythme hebdomadaire, bilan humain, flux acteur → mode → cible, pression sur Bamako, matrice région × semaine, fiabilité ;
- évaluation stratégique : jugements clés, scénarios, indicateurs d'alerte, recommandations, limites des données ;
- registre des incidents, avec recherche et tri.

## Régénérer après mise à jour de la base

```bash
pip install -r requirements.txt
python scripts/build_dashboard.py data/Base_Donnees_Securite_Mali_OpenMap_Dashboard.xlsx
```

La feuille `DATA_CARTOGRAPHIE` doit garder ses en-têtes de colonnes. Le script produit dans `dist/` :

| Fichier | Usage |
|---|---|
| `tableau_de_bord_mali.html` | Tableau de bord autonome |
| `artifact_mali.html` | Même page, au format fragment pour publication Artifact |
| `incidents_mali.geojson` | Couche de points enrichie pour QGIS / uMap / OpenStreetMap |
| `incidents_mali_enrichi.csv` | Table enrichie : région et cercle géocodés, mode normalisé, fiabilité, distance à Bamako |

Le fond `data/geo/basemap.json` provient de geoBoundaries (CC BY 4.0) et de Natural Earth.
Pour le reconstruire : `python scripts/prepare_geo.py <dossier_sources>` (voir l'en-tête du script).
