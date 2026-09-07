import sys

import pandas as pd

# Permet de lancer/d'utiliser les applications Microsoft Office
import win32com.client

from PyQt6.QtCore import Qt, QMimeData
from PyQt6.QtWidgets import (
    QApplication,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

# Récupère tous les éléments produits dans moteur et utilisables pour l'interface
from moteur import (
    complet,
    articles_geo,
    annees_dispo,
    date_maj,
    date_telechargement,
    premiere_annee,
    derniere_annee,
)

# Mise en forme simple appliquée au tableau HTML de sortie. 
STYLE_TABLEAU = """
<style>
body { font-family: Calibri, sans-serif; font-size: 11pt; }
pre { font-family: Calibri, sans-serif; font-size: 11pt; }
th, td { border: 1px solid #ccc; padding: 6px 10px; text-align: center; }
</style>
"""

class FenetrePrincipale(QMainWindow):
    # Crée et organise les différents blocs de l'interface
    def __init__(self):
        super().__init__()

        self.setWindowTitle(
            "Synthèse de dette publique – Eurostat"
        )

        widget_principal = QWidget()
        self.setCentralWidget(widget_principal)

        self.layout_principal = QVBoxLayout(widget_principal)

        # Appelle chaque bloc de l'interface et la dynamise
        self.creer_entete()
        self.creer_selecteurs()
        self.creer_bouton_annees()
        self.creer_zone_resultat()
        self.creer_boutons_resultat()

        self.mettre_a_jour_apercu()

    # Premiers blocs : titre, données et texte introductifs
    def creer_entete(self):
        titre = QLabel(
            "<h1>Synthèse de dette publique Eurostat</h1>"
        )

        introduction = QGroupBox("Infos générales")
        layout_introduction = QVBoxLayout(introduction)

        informations = QLabel(
            f"<b>Période disponible :</b> "
            f"{premiere_annee}-{derniere_annee}<br>"
            f"<b>Dernière mise à jour Eurostat :</b> "
            f"{date_maj}<br>"
            f"<b>Date de téléchargement :</b> "
            f"{date_telechargement}"
        )

        layout_introduction.addWidget(informations)

        instructions = QGroupBox("Instructions")
        layout_instructions = QVBoxLayout(instructions)

        param_defaut = QLabel(
            f"Le tableau et le mail portent toujours sur la France, la Zone euro "
            f"et l’Union européenne.<br>"
            f"Par défaut, l’outil présélectionne les années n, n-1 et n-3."
        )

        instruction = QLabel(
            f"Vous pouvez modifier manuellement la sélection des années "
            f"(elle doit contenir exactement 3 éléments).<br>"
            f"Pour l’année la plus récente, l’outil précise également "
            f"l’État membre de l’UE27 le plus endetté et le moins endetté "
            f"(hors agrégats)."
        )

        layout_instructions.addWidget(param_defaut)
        layout_instructions.addWidget(instruction)

        self.layout_principal.addWidget(titre)
        self.layout_principal.addWidget(introduction)
        self.layout_principal.addWidget(instructions)

    # Création d'un bloc cliquable d'années
    def creer_selecteurs(self):
        # Sélecteur des années - seul élément que l'utilisateur choisit,
        # les territoires du tableau et du mail étant fixes.
        groupe_annees = QGroupBox("Années d’intérêt")
        layout_annees = QVBoxLayout(groupe_annees)

        self.liste_annees = QListWidget()
        layout_annees.addWidget(self.liste_annees)

        annees_par_defaut = {annees_dispo[0], annees_dispo[1], annees_dispo[3]}
        self.annees_par_defaut = annees_par_defaut

        for annee in annees_dispo:
            item = QListWidgetItem(str(annee))

            item.setFlags(
                item.flags()
                | Qt.ItemFlag.ItemIsUserCheckable
            )

            if annee in annees_par_defaut:
                item.setCheckState(Qt.CheckState.Checked)
            else:
                item.setCheckState(Qt.CheckState.Unchecked)

            self.liste_annees.addItem(item)

        self.layout_principal.addWidget(groupe_annees)

        # Connecté une fois la liste peuplée, pour ne pas se déclencher
        # pendant sa construction : chaque coche/décoche met l'aperçu à jour.
        self.liste_annees.itemChanged.connect(
            lambda _item: self.mettre_a_jour_apercu()
        )

    # Création d'un bouton dont la fonction/l'action associée 
    # permet de réinitialiser la sélection (n, n-1 et n-3)
    def creer_bouton_annees(self):
        layout_boutons = QHBoxLayout()

        bouton_reinitialiser = QPushButton(
            "Réinitialiser la sélection par défaut"
        )

        bouton_reinitialiser.clicked.connect(
            self.reinitialiser_selection
        )

        layout_boutons.addWidget(
            bouton_reinitialiser
        )

        self.layout_principal.addLayout(
            layout_boutons
        )

    # Création d'un espace aperçu
    def creer_zone_resultat(self):
        groupe_resultat = QGroupBox("Aperçu du résultat")
        layout_resultat = QVBoxLayout(groupe_resultat)

        self.resultat = QTextEdit()
        self.resultat.setReadOnly(True)
        self.resultat.setPlaceholderText(
            "La sélection apparaîtra ici."
        )

        layout_resultat.addWidget(self.resultat)

        self.layout_principal.addWidget(groupe_resultat)


    # Créations d'un bouton dont la fonction/l'action associée
    # permet de copier les résultats ou d'en faire un mail
    def creer_boutons_resultat(self):
        layout_boutons = QHBoxLayout()

        bouton_copier = QPushButton(
            "Copier"
        )
        bouton_creer_mail = QPushButton(
            "Créer le mail"
        )

        bouton_copier.clicked.connect(
            self.copier_apercu
        )
        bouton_creer_mail.clicked.connect(
            self.creer_mail
        )

        layout_boutons.addWidget(
            bouton_copier
        )
        layout_boutons.addWidget(
            bouton_creer_mail
        )

        self.layout_principal.addLayout(
            layout_boutons
        )


    # Récupère les années sélectionnées par l'utilisateur
    def valeurs_cochees(self, liste):
        return [
            liste.item(index).text()
            for index in range(liste.count())
            if liste.item(index).checkState()
            == Qt.CheckState.Checked
        ]

    # Action/fonction permettant la réinitialisation des années sélectionnées
    def reinitialiser_selection(self):
        valeurs_par_defaut = {
            str(annee) for annee in self.annees_par_defaut
        }

        for index in range(self.liste_annees.count()):
            item = self.liste_annees.item(index)
            item.setCheckState(
                Qt.CheckState.Checked
                if item.text() in valeurs_par_defaut
                else Qt.CheckState.Unchecked
            )

        self.statusBar().showMessage(
            "Sélection réinitialisée aux valeurs par défaut.", 4000
        )

    # Fonction automatique de mise à jour de l'aperçu dès
    # modification de la sélection des années : l'utilisateur
    # doit nécessairement sélectionner 3 années
    def mettre_a_jour_apercu(self):
        annees_selectionnees = [
            int(annee)
            for annee in self.valeurs_cochees(
                self.liste_annees
            )
        ]

        NOMBRE_ANNEES_SELECTION = 3

        if len(annees_selectionnees) != NOMBRE_ANNEES_SELECTION:
            self.resultat.setPlainText(
                f"Sélectionnez exactement {NOMBRE_ANNEES_SELECTION} années "
                f"(actuellement : {len(annees_selectionnees)})."
            )
            return

        texte_tableau, corps_mail = self.construire_synthese(
            annees_selectionnees
        )

        # Conservés pour le bouton "Créer le mail", afin de ne pas
        # recalculer la synthèse à ce moment-là.
        self.dernier_tableau_html = texte_tableau
        self.dernier_corps_mail = corps_mail

        # Application de la légère mise en forme spécifiée en début de fichier
        # sur les résultats
        texte_html = (
            STYLE_TABLEAU
            + f"<pre>{corps_mail}</pre><br>"
            + f"{texte_tableau}"
        )

        self.resultat.setHtml(texte_html)

    # Action/fonction permettant de créer un mail/brouillon avec les éléments de
    # l'aperçu + ajoute un objet
    def creer_mail(self):
        """
        Crée un brouillon Outlook avec le message et le tableau déjà
        présents dans le corps du mail (aucun collage manuel nécessaire).
        """
        if not hasattr(self, "dernier_corps_mail"):
            self.statusBar().showMessage(
                "Sélectionnez d’abord exactement 3 années.", 4000
            )
            return

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject = (
            f"[Pour info] Publication par Eurostat des ratios de dette "
            f"publique (données mises à jour le {date_maj})"
        )
        mail.HTMLBody = (
            STYLE_TABLEAU
            + f"<pre>{self.dernier_corps_mail}</pre><br>"
            + self.dernier_tableau_html
        )
        mail.Display()

        self.statusBar().showMessage(
            "Brouillon Outlook créé avec le message et le tableau.", 5000
        )

    # Action/fonction permettant de copier l'aperçu dans le presse-papiers
    def copier_apercu(self):
        if not hasattr(self, "dernier_corps_mail"):
            self.statusBar().showMessage(
                "Sélectionnez d’abord exactement 3 années.", 4000
            )
            return

        mime = QMimeData()
        mime.setHtml(
            STYLE_TABLEAU
            + f"<pre>{self.dernier_corps_mail}</pre><br>"
            + self.dernier_tableau_html
        )
        QApplication.clipboard().setMimeData(mime)

        self.statusBar().showMessage("Aperçu copié dans le presse-papiers.", 4000)

    # Fonction construisant les éléments de synthèse : texte + tableau
    def construire_synthese(self, annees_selectionnees):

        annee_r1, annee_r2, annee_r3 = sorted(
            annees_selectionnees, reverse=True
        )


        extreme_max, extreme_min = (
            complet.set_index(['code_geo', 'territoire_geo'])[f"{annee_r1}"].idxmax(),
            complet.set_index(['code_geo', 'territoire_geo'])[f"{annee_r1}"].idxmin()
        )

        COLONNES_PAR_ANNEE = {
            int(colonne): colonne
            for colonne in complet.columns
            if str(colonne).isdigit()
        }

        def valeur_territoire(nom_territoire, annee):
            valeurs = complet.loc[
                complet["territoire_geo"] == nom_territoire, COLONNES_PAR_ANNEE[annee]
            ].dropna()

            return valeurs.iloc[0]

        # dico avec nom : [valeur_ar2, valeur_ar1, delta_ar1_ar2, delta_ar1_ar3]
        tous_elements = {}

        ZONE_EURO = 'Zone euro - 21 pays (à partir de 2026)'
        UE = 'Union européenne - 27 pays (à partir de 2020)'

        for t in ['France', ZONE_EURO, UE, extreme_max[1], extreme_min[1]]:

            tous_elements[t] = [
                f'{valeur_territoire(t, annee_r2):.1f}'.replace(".",","),
                f'{valeur_territoire(t, annee_r1):.1f}'.replace(".",","),
                f'{(valeur_territoire(t, annee_r1) - valeur_territoire(t, annee_r2)):+.1f}'.replace(".",","),
                f'{(valeur_territoire(t, annee_r1) - valeur_territoire(t, annee_r3)):+.1f}'.replace(".",",")
                ]


        mail = (
        f"Bonjour à tous,\n"
        f"\n"
        f"Eurostat a publié ce jour les ratios de dette publique des États "
        f"membres ({annee_r1} étant l’année la plus récente utilisée).\n"
        f"\n"
        f"La dette publique française s’élève à {tous_elements['France'][1]} % du PIB, en "
        f"{'hausse' if tous_elements['France'][2][0]=='+' else 'baisse'} de {tous_elements['France'][2][1:]} points par "
        f"rapport à {annee_r2}. Au niveau de la zone euro, le ratio de dette "
        f"publique des États membres s’élève à {tous_elements[ZONE_EURO][1]} %.\n"
        f"\n"
        f"{articles_geo[extreme_min[1]].title()}{extreme_min[1]} est l’État membre avec le plus faible ratio de "
        f"dette publique ({tous_elements[extreme_min[1]][1]} %) alors que {articles_geo[extreme_max[1]]}{extreme_max[1]} voit sa "
        f"dette atteindre {tous_elements[extreme_max[1]][1]} %, ce qui constitue le ratio le plus "
        f"élevé de l’Union européenne."
        )


        tableau = pd.DataFrame(
            {
                ('Ratio dette/PIB', 'par territoire'): [
                    'France',
                    'Zone euro', 
                    'Union européenne',
                    f'État avec la plus forte dette en {annee_r1} : {extreme_max[1]}',
                    f'État avec la plus faible dette en {annee_r1} : {extreme_min[1]}'
                ],
                (str(annee_r2), '% PIB'): [tous_elements[t][0] for t in tous_elements],
                (str(annee_r1), '% PIB'): [tous_elements[t][1] for t in tous_elements],
                (f'Variation {annee_r1}/{annee_r2}', 'points de PIB'): [tous_elements[t][2] for t in tous_elements],
                (f'Variation {annee_r1}/{annee_r3}', 'points de PIB'): [tous_elements[t][3] for t in tous_elements]
            }
        )

        return (
            tableau.style.hide(axis="index").set_table_attributes('cellspacing="0"').to_html(),
            mail
        )


if __name__ == "__main__":
    app = QApplication(sys.argv)

    fenetre = FenetrePrincipale()
    fenetre.show()

    sys.exit(app.exec())
