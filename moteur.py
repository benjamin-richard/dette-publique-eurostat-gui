import eurostat

from datetime import datetime as dt

## NOM DE LA BASE D'INTÉRÊT (permet la réutilisation pour d'autres bases Eurostat)
code_bdd = 'sdg_17_40'
unit = 'PC_GDP'

#### RÉCUPÉRATION DES DONNÉES BRUTES ET COMPLÉMENTS D'IDENTIFICATION
brut = eurostat.get_data_df(code_bdd)

# Récupération dataframe en français permettant de lier le code_geo au nom du territoire
labels_geo = (
    eurostat
    .get_dic(
        code_bdd,
        par="geo",
        frmt="df",
        lang="fr",
        full=False
    )
)

complet = (
    brut
    .loc[
        lambda df: df['unit'] == unit
    ]
    .rename(columns={'geo\\TIME_PERIOD': 'code_geo'})
    .assign(
        territoire_geo=lambda df: df['code_geo'].map(labels_geo.set_index('val').squeeze())
    )
)

# Création d'une fonction et d'un dictionnaire permettant d'associer un article
# à chaque territoire
def article(mot):
    voyelles = "aeiouy"

    mot = mot.strip().lower()

    if mot in ['chypre', 'malte']:
        return ""
    elif mot[0] in voyelles:
        return "l'"
    elif mot.split()[0].endswith("e"):
        return "la"
    else:
        return "le"


articles_geo = {
    label: article(label)
    for label in labels_geo["descr"]
}

#### LISTING DES DONNÉES MOBILISABLES PAR L'UTILISATEUR
## Dates et années

# Récupération des informations temporelles associées à la base d'intérêt
infos_bdd = eurostat.get_toc_df()

dates = (
    infos_bdd
    .loc[
        lambda df: df["code"].str.upper() == code_bdd.upper(),
        ['last update of data', 'data start', 'data end']
    ]
    .squeeze()
)

(date_maj, premiere_annee, derniere_annee) = (
    dt.strptime(dates.iloc[0],"%Y-%m-%dT%H:%M:%S%z").strftime("%d-%m-%Y"),
    dates.iloc[1],
    dates.iloc[2]
)

date_telechargement = dt.now().strftime("%d-%m-%Y à %H:%M")

# Le parcours des colonnes permet de ne considérer que les années pour lesquelles
# des données existent (à priori, toutes celles situées entre la premiere_annee et la derniere_annee)
annees_dispo = (
    sorted(
        [
        int(col)
        for col in complet.columns
        if str(col).isdigit()
    ],
    reverse=True
    )
)

