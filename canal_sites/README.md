# Cartographie des points de vente agréés (Burkina Faso, Mali, Niger)

Carte dynamique des sites et classeur de suivi, prêts pour un déploiement Netlify.

```bash
python canal_sites/build_sites.py <dossier_geo> export_BF.xlsx [export_ML.xlsx export_NE.xlsx]
python canal_sites/build_excel.py
```

Sorties dans `canal_sites/dist/` (exclu du dépôt : les exports contiennent des données de contact) :
- `site_netlify/` et `Site_Netlify_Cartographie_PVA.zip` : dossier à glisser sur https://app.netlify.com/drop.
  La page ne publie ni téléphones, ni emails, ni identifiants, et elle est marquée « noindex ».
- `Sites_PVA_Burkina_Mali_Niger.xlsx` : TABLEAU_DE_BORD (cellule du lien Netlify), SITES, QUALITE_INTEGRATION.
- `sites_pva.geojson` et `sites_pva.kml` : couches SIG.

Contrôles appliqués aux coordonnées : position comparée au pays déclaré, signe de longitude manquant
corrigé, positions génériques ou partagées signalées. Les coordonnées d'origine restent dans le classeur.
