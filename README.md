# Récupération de la dette publique Eurostat

Application PyQt6 qui automatise la récupération des ratios de dette publique (% du PIB)
(Eurostat, [sdg_17_40](https://ec.europa.eu/eurostat/databrowser/view/sdg_17_40/default/table?lang=en&category=t_gov.t_gov_gfs10.t_gov_dd))
puis la génération du mail de synthèse envoyé au cabinet : France, zone euro, Union
européenne, État le plus/le moins endetté, avec leurs évolutions.

## Démarrage rapide

1. `pip install -r requirements.txt`
2. Double-cliquer sur `lancer_generer_synthese.bat` (ou `python generer_synthese.py` depuis un terminal)

*(L'utilisation de l'interface nécessite Python et la création de mail Windows + Outlook)*
