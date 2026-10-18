# -*- coding: utf-8 -*-
"""
FPPN - Construction du rapport actuariel et statistique (PowerPoint, charte Adenka, Timesaver)
==============================================================================================

Entrée  : resultats.json (analyses_fppn.py) ; Timesaver Adenka ; outils du skill adenka-powerpoint.
Sortie  : le deck .pptx, avec graphiques et tableaux natifs, listes des figures et des tableaux,
          table des signes et abréviations, table des références et sources.

Utilisation : python rapport_fppn.py <resultats.json> <dossier_skill_adenka_powerpoint> <sortie.pptx>
"""
import copy
import io
import json
import os
import sys

RES, SKILL, SORTIE = sys.argv[1], sys.argv[2], sys.argv[3]
sys.path.insert(0, os.path.join(SKILL, "scripts"))
from pptx_tools import selection, remplir, finaliser  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.util import Cm, Pt  # noqa: E402
from pptx.dml.color import RGBColor  # noqa: E402
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR  # noqa: E402
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402

R = json.load(open(RES, encoding="utf-8"))
TIMESAVER = os.path.join(SKILL, "assets", "Timesaver_Adenka.pptx")
DATE_RAPPORT = "15 octobre 2026"


# ---------------------------------------------------------------------------
# Mise en forme des nombres (français)
# ---------------------------------------------------------------------------
def nb(x, dec=0):
    s = f"{x:,.{dec}f}".replace(",", " ").replace(".", ",")
    return s


def M(x, dec=0):
    """millions de FCFA"""
    return nb(x / 1e6, dec)


def Md(x, dec=1):
    return nb(x / 1e9, dec)


def pc(x, dec=0):
    return nb(100 * x, dec) + " %"


def an(liste, cle, annee):
    return next(r for r in liste if r.get("exercice", r.get("annee")) == annee)


SIN = R["sinistralite"]
BIL = R["bilan"]
s21, s23, s25, s26 = (an(SIN, "exercice", a) for a in (2021, 2023, 2025, 2026))
b20, b25, b26 = (an(BIL, "exercice", a) for a in (2020, 2025, 2026))
GP = R["grenier_par"]
GE = R["grenier_encours"]
EFF = R["effectifs_mensuels"]
CG = R["cgrae"]
T936 = R["taux_936"]

# ---------------------------------------------------------------------------
# Plan du rapport : chaque élément décrit une diapositive
# ---------------------------------------------------------------------------
SL = []          # diapositives
FIG, TAB = [], []  # (numéro, titre, n° de diapositive)


def contenu(entete, **k):
    SL.append({"type": "contenu", "entete": entete, **k})


def separateur(num, titre, sous_titre):
    SL.append({"type": "separateur", "num": num, "titre": titre, "sous_titre": sous_titre})


def fig(titre):
    FIG.append([len(FIG) + 1, titre, len(SL) + 1])
    return f"Figure {len(FIG)} : {titre}"


def tab(titre):
    TAB.append([len(TAB) + 1, titre, len(SL) + 1])
    return f"Tableau {len(TAB)} : {titre}"


SRC_PREC = "Source : fichiers de précomptes de la Solde, de la CGRAE et du FPPN reçus le 15/10/2026 ; traitements Adenka"
SRC_COMPTA = "Source : balances des comptes du FPPN (Sage 100), exercices 2020 à 2025 et 01/01 au 31/08/2026 (provisoire) ; traitements Adenka"
SRC_GRENIER = "Source : exports du logiciel de prêts du Grenier du Policier (balances âgées, production) reçus le 15/10/2026 ; traitements Adenka"

SL.append({"type": "couverture"})
SL.append({"type": "sommaire"})

# --- Résumé exécutif --------------------------------------------------------
eff0, eff1 = EFF[0], EFF[-1]
contenu({"titre": "Résumé exécutif",
         "sous_titre": "Le FPPN dégage des excédents mais reste fragile et peu documenté",
         "synthese": f"Fonds propres négatifs, épargne couverte à {pc(b26['couverture_epargne'])} par des actifs dédiés et provisions techniques figées depuis 2020"},
        tableau={"titre": tab("Principaux constats au 15/10/2026"),
                 "entetes": ["Domaine", "Constat", "Chiffre clé"],
                 "largeurs": [5.2, 18.6, 7.4],
                 "lignes": [
                     ["Données", "Les précomptes ne sont disponibles qu'à partir de 2023 ou 2024 selon le poste. Aucune base de prestations n'a été reçue.", "4,45 millions de lignes contrôlées"],
                     ["Cotisations", "Le précompte total est bien de 5 % du salaire de base, mais réparti 3,77 % maladie et 1,23 % décès.", "5,00 % = 3,77 % + 1,23 %"],
                     ["Maladie", "Les prestations maladie comptabilisées augmentent plus vite que les cotisations.", f"S/C {pc(s21['sc_maladie'])} en 2021, {pc(s25['sc_maladie'])} en 2025, {pc(s26['sc_maladie'])} sur 8 mois 2026"],
                     ["Provisions", "Les provisions pour prestations à payer n'ont pas changé depuis 2020.", f"{M(b26['provisions_techniques'])} M FCFA chaque année"],
                     ["Épargne", "L'épargne des mutualistes (PEP, EFC) dépasse largement les actifs qui lui sont dédiés.", f"{Md(b26['epargne_totale'])} Md FCFA pour {Md(b26['actifs_dedies_epargne'])} Md d'actifs"],
                     ["Fonds propres", "Les fonds propres sont négatifs sur toute la période, les excédents réduisent l'écart.", f"{M(b20['fonds_propres'])} M (2020), {M(b25['fonds_propres'])} M (2025)"],
                     ["Grenier", "Le portefeuille de prêts croît vite, financé par le FPPN, les retards ont culminé début 2026.", f"Encours {M(GP['encours'])} M FCFA au 31/08/2026"],
                     ["Actifs", "Une part importante de l'actif est faite de créances et d'investissements peu liquides.", f"Hôpital de la Police {M(b26['creance_hopital_police'])} M, CMP {M(b26['cmp_investissements'])} M"]]},
        source="Source : analyses Adenka sur les documents reçus le 15/10/2026 (S/C : prestations / cotisations ; M : millions ; Md : milliards)")

contenu({"titre": "Recommandations prioritaires",
         "sous_titre": "Cinq actions conditionnent la fiabilité des engagements et la solvabilité",
         "synthese": "Les trois premières actions relèvent de la transmission de données, les deux dernières, des décisions de gestion"},
        tableau={"titre": tab("Actions recommandées"),
                 "entetes": ["N°", "Action", "Objet", "Priorité"],
                 "largeurs": [1.4, 11.0, 15.0, 3.8],
                 "lignes": [
                     ["1", "Transmettre les bases de prestations décès et santé 2020 à 2026", "Calculer les provisions (PSAP), la sinistralité réelle par année de survenance et les délais", "Urgente"],
                     ["2", "Transmettre le fichier des mutualistes avec date de naissance et sexe", "Mesurer la mortalité, projeter l'équilibre technique, calculer les engagements", "Urgente"],
                     ["3", "Transmettre les comptes individuels PEP et EFC (soldes, intérêts, retraits)", "Évaluer la dette d'épargne envers chaque mutualiste et la rapprocher de la comptabilité", "Urgente"],
                     ["4", "Recalculer les provisions techniques à chaque clôture", "Remplacer des montants figés depuis 2020 par une estimation fondée sur les dossiers", "Haute"],
                     ["5", "Adosser l'épargne à des placements liquides et suivis", "Couvrir l'épargne exigible, encadrer le financement du Grenier et du CMP", "Haute"]]},
        source="Source : analyses Adenka ; numéros de la liste des documents (PBC) en annexe")

# --- Partie 1 : contexte et démarche --------------------------------------
separateur(1, "Contexte et démarche", "Mission, périmètre, référentiels et outils")
contenu({"titre": "Mission et périmètre",
         "sous_titre": "Le volet actuariel couvre les engagements, les provisions et les placements",
         "synthese": "Période d'analyse : du 1er janvier 2020 au 31 août 2026, données arrêtées à l'envoi du FPPN reçu le 15/10/2026"},
        tableau={"titre": tab("Périmètre de la mission actuarielle et statistique"),
                 "entetes": ["Axe", "Question traitée", "État au 15/10/2026"],
                 "largeurs": [6.4, 15.6, 9.2],
                 "lignes": [
                     ["Contrôle des données de base", "Les données sont-elles complètes et cohérentes avec la comptabilité ?", "Réalisé sur les données reçues"],
                     ["Engagements envers les adhérents", "Combien le FPPN doit-il aux mutualistes et ayants droit ?", "Partiel : épargne seulement"],
                     ["Suffisance des provisions", "Les provisions couvrent-elles les prestations dues ?", "Constat sur les montants, calcul en attente des dossiers"],
                     ["Adossement actif / passif", "Les actifs sont-ils disponibles quand les dettes arrivent à échéance ?", "Réalisé sur les balances"],
                     ["Adéquation des placements", "Les placements sont-ils sûrs, liquides, rentables ?", "Réalisé sur les balances"],
                     ["Indicateurs AIRMS", "Les ratios des mutuelles sociales sont-ils respectés ?", "En attente des seuils du règlement"],
                     ["Équilibre technique", "Les cotisations futures couvrent-elles les prestations futures ?", "En attente des âges et des dossiers"]]},
        source="Source : offre technique Adenka et liste des documents demandés (PBC)")

contenu({"titre": "Démarche et outils",
         "sous_titre": "Chaque chiffre du rapport est tracé jusqu'au fichier du FPPN qui le porte",
         "synthese": "Extraction sans perte, base contrôlée, analyses reproductibles : 4 287 fichiers inventoriés, 724 fichiers mensuels retenus"},
        etapes=[["01", "Inventaire et extraction", "4 287 fichiers lus, feuilles Excel converties sans perte, empreinte SHA-256 de chaque fichier."],
                ["02", "Base d'analyse", "Lecture par en-tête ou par contenu, lignes de total écartées, une source retenue par poste et par mois."],
                ["03", "Contrôles", "Rapprochement aller / retour, totaux imprimés du Grenier, équilibre des balances comptables."],
                ["04", "Analyses et rapport", "Indicateurs, graphiques natifs, classeur des calculs, données externes validées par Adenka."]],
        source="Source : scripts Python Adenka (extraction_fppn.py, base_fppn.py, grenier_fppn.py, compta_fppn.py, analyses_fppn.py)")

# --- Partie 2 : données ----------------------------------------------------
separateur(2, "Données et qualité", "Couverture de la période, contrôles et anomalies")
cov = {(c["poste"], c["sens"]): c["mois_recus"] for c in R["couverture"]}
lib_flux = [("Comptabilité (balances)", 80), ("PEP 998 retour", cov[("998", "retour")]),
            ("EFC 870 retour", cov[("870", "retour")]), ("Grenier 9X4 retour", cov[("9X4", "retour")]),
            ("Flotte 9W3 retour", cov[("9W3", "retour")]), ("Maladie 936", cov[("936", "retour")]),
            ("Retraités CGRAE retour", cov[("CGRAE", "retour")]), ("Décès 8Z6", cov[("8Z6", "retour")]),
            ("MEDOC 8Z7 retour", cov[("8Z7", "retour")]), ("Prestations décès et santé", 0)]
contenu({"titre": "Couverture de la période",
         "sous_titre": "Les données mensuelles ne couvrent pas les années 2020 à 2022",
         "synthese": "Sur 80 mois, seuls la comptabilité et le PEP sont complets, aucune donnée de prestations n'a été reçue"},
        graphique={"titre": fig("Mois reçus par flux de données, sur 80 mois"), "type": "barres_h",
                   "categories": [x[0] for x in lib_flux][::-1],
                   "series": [{"nom": "Mois reçus", "valeurs": [x[1] for x in lib_flux][::-1], "couleur": "112A40"}],
                   "format": "0", "unite": "En nombre de mois (janvier 2020 à août 2026)"},
        texte=["**La comptabilité couvre toute la période en totaux annuels.** Les balances 2020 incluent le cumul de janvier à août, seul le détail des écritures de ces huit mois manque.",
               "**Les précomptes mensuels commencent en 2023 ou 2024.** Maladie et décès débutent en janvier 2024, retraités en février 2024, EFC en janvier 2023.",
               "**Le PEP est complet grâce aux fichiers texte de la Solde.** Ils donnent le détail par mutualiste et par mois de 2020 à 2022.",
               "**Aucune base de prestations n'a été transmise.** Les analyses de sinistres et de provisions par dossier restent en attente."],
        source=SRC_PREC)

rt = R["lignes_total_ecartees"]
contenu({"titre": "Contrôles d'exhaustivité",
         "sous_titre": "La base est exhaustive et les anomalies des fichiers sources sont corrigées",
         "synthese": f"{rt['n']} lignes de total écartées ({Md(rt['montant'])} Md FCFA de double compte évité) et {R['montants_corriges_8Z6']} montants doublés corrigés"},
        tableau={"titre": tab("Contrôles réalisés et résultats"),
                 "entetes": ["Contrôle", "Résultat", "Conséquence"],
                 "largeurs": [10.2, 11.2, 9.8],
                 "lignes": [
                     ["Archive transmise et fichiers copiés", "Empreintes SHA-256 identiques (218 fichiers sur 218)", "Aucune perte au transfert"],
                     ["Feuilles Excel converties", "Lignes et cellules identiques à la source (491 sur 491)", "Aucune perte à la conversion"],
                     ["Balances âgées du Grenier", "Nombre de prêts, encours et montants débloqués égaux aux totaux imprimés (37 sur 37)", "Données de prêts fiables"],
                     ["Balances comptables", "Débit = crédit, ouverture + mouvements = solde (7 exercices)", "Comptabilité cohérente"],
                     ["Lignes de total dans les fichiers de la Solde", f"{rt['n']} lignes écartées", "Évite de compter deux fois les montants"],
                     ["Montants doublés (décès, avril 2024)", f"{R['montants_corriges_8Z6']} montants saisis deux fois (ex. 4792447924 pour 47924)", "Corrigés et signalés au FPPN"],
                     ["Fichiers vides ou endommagés", "Décès et maladie de décembre 2024 vides, 7 fichiers comptables endommagés", "À renvoyer par le FPPN"]]},
        source="Source : journal de contrôle de l'extraction et traitements Adenka")

ar = R["aller_retour"]
nom_poste = {"870": "EFC 870", "998": "PEP 998", "9X4": "Grenier 9X4", "9W3": "Flotte 9W3", "CGRAE": "Retraités CGRAE", "8Z7": "MEDOC 8Z7"}
contenu({"titre": "Rapprochement aller et retour",
         "sous_titre": "La Solde précompte environ 98 % des montants demandés par le FPPN",
         "synthese": "Sur les mois communs, l'écart entre montants demandés et précomptés reste faible, mais certains mois s'écartent fortement"},
        graphique={"titre": fig("Montant précompté rapporté au montant demandé, mois communs"), "type": "barres",
                   "categories": [nom_poste[a["poste"]] for a in ar],
                   "series": [{"nom": "Taux moyen", "valeurs": [round(100 * a["taux_recouvrement"], 1) for a in ar], "couleur": "112A40"},
                              {"nom": "Mois le plus bas", "valeurs": [round(100 * a["taux_min_mensuel"], 1) for a in ar], "couleur": "EFAC62"}],
                   "format": "0.0", "unite": "En % du montant demandé (fichier aller)"},
        texte=["**Le taux de précompte moyen est compris entre 98 % et 100 %.** Les montants demandés par le FPPN sont donc presque entièrement retenus sur les salaires et les pensions.",
               f"**Certains mois s'écartent nettement.** Le mois le plus bas descend à {pc(min(a['taux_min_mensuel'] for a in ar))} pour le Grenier et à {pc([a for a in ar if a['poste']=='998'][0]['taux_min_mensuel'])} pour le PEP.",
               "**Ces écarts traduisent des rejets ou des décalages.** Ils devront être expliqués par le FPPN, mois par mois, à partir des listes de rejets de la Solde."],
        source=SRC_PREC)

# --- Partie 3 : effectifs et cotisations ---------------------------------
separateur(3, "Effectifs et cotisations", "Cotisants, retraités, taux de cotisation")
lab = [e["libelle"] for e in EFF]
contenu({"titre": "Effectifs cotisants",
         "sous_titre": "Les cotisants actifs augmentent d'un tiers en trente mois",
         "synthese": f"{nb(eff0['cotisants'])} cotisants en {eff0['libelle']}, {nb(eff1['cotisants'])} en {eff1['libelle']}, les recrues entrent chaque mois de juillet"},
        graphique={"titre": fig("Cotisants actifs précomptés (poste maladie 936), par mois"), "type": "lignes",
                   "categories": lab, "series": [{"nom": "Cotisants actifs", "valeurs": [e["cotisants"] for e in EFF], "couleur": "112A40"}],
                   "format": "# ##0", "etiquettes": False, "axe_min": 20000, "legende": None,
                   "unite": "En nombre de mutualistes"},
        texte=[f"**Les effectifs cotisants progressent de {pc(eff1['cotisants'] / eff0['cotisants'] - 1)}.** Ils passent de {nb(eff0['cotisants'])} en janvier 2024 à {nb(eff1['cotisants'])} en août 2026.",
               "**La hausse se fait par paliers en juillet.** Juillet 2024 compte 4 926 entrées et juillet 2025 4 127 : ce sont les promotions de recrues.",
               f"**Les sorties sont faibles et concentrées.** {nb(R['flux_annuels'][0]['sorties'])} sorties en 2024 et {nb(R['flux_annuels'][1]['sorties'])} en 2025, leurs motifs (retraite, décès, radiation) ne sont pas connus.",
               "**Le fichier ne donne ni âge ni sexe.** Sans ces données, la mortalité et les départs à la retraite ne peuvent pas être projetés."],
        source=SRC_PREC + " (décembre 2024 : fichier vide)")

an_ = R["anciennete"]
contenu({"titre": "Structure des effectifs",
         "sous_titre": "Une population jeune en service, recrutée en majorité depuis moins de dix ans",
         "synthese": f"En {an_['mois']}, ancienneté médiane de {nb(an_['mediane'], 1)} ans et salaire de base médian de {nb(an_['salaire_base_median'])} FCFA"},
        graphique={"titre": fig(f"Agents par ancienneté de service, {an_['mois']}"), "type": "barres",
                   "categories": an_["categories"], "series": [{"nom": "Agents", "valeurs": an_["valeurs"], "couleur": "112A40"}],
                   "format": "# ##0", "unite": "En nombre d'agents"},
        texte=[f"**{pc(sum(an_['valeurs'][:2]) / sum(an_['valeurs']))} des agents ont moins de dix ans de service.** Les fortes promotions récentes rajeunissent la population cotisante.",
               f"**Une seconde vague a 15 à 25 ans de service.** Elle représente {pc((an_['valeurs'][3] + an_['valeurs'][4]) / sum(an_['valeurs']))} des agents et partira à la retraite dans les dix à quinze prochaines années.",
               "**Ces départs feront passer des cotisants au statut de retraités CGRAE.** Le coût de leur prise en charge santé devra être projeté dès que les âges seront connus."],
        source=SRC_PREC + " (fichiers maladie 936 détaillés : grade, date de recrutement, salaire de base)")

contenu({"titre": "Taux de cotisation",
         "sous_titre": "La Solde précompte 5 % du salaire de base, mais pas selon la répartition annoncée",
         "synthese": f"Sur {nb(T936['lignes'])} lignes, le précompte maladie vaut 3,77 % et le précompte décès 1,23 % du salaire de base"},
        graphique={"titre": fig("Répartition du taux de 5 % : annoncé et observé"), "type": "barres_empilees",
                   "categories": ["Taux annoncé", "Taux précompté (Solde)", "Répartition comptable 2025"],
                   "series": [{"nom": "Maladie", "valeurs": [3.5, 3.77, round(5 * s25['cot_maladie'] / (s25['cot_maladie'] + s25['cot_deces']), 2)], "couleur": "112A40"},
                              {"nom": "Décès", "valeurs": [1.5, 1.23, round(5 * s25['cot_deces'] / (s25['cot_maladie'] + s25['cot_deces']), 2)], "couleur": "EFAC62"}],
                   "format": "0.00", "unite": "En % du salaire de base"},
        texte=["**Le total de 5 % du salaire de base est respecté.** Il se vérifie sur la quasi-totalité des agents des mois détaillés.",
               "**La répartition entre maladie et décès diffère du barème annoncé.** La Solde retient 3,77 % au poste maladie et 1,23 % au poste décès, au lieu de 3,5 % et 1,5 %.",
               "**La comptabilité répartit les cotisations à peu près à 70 / 30.** Le résultat technique par garantie dépend donc d'une clé qui n'est pas celle des précomptes.",
               "**Le barème officiel n'a pas été reçu.** Il doit trancher entre les deux répartitions."],
        source=SRC_PREC + " ; " + SRC_COMPTA.replace("Source : ", ""))

cot25 = an(R["cotisations"], "exercice", 2025)
contenu({"titre": "Cotisations comptabilisées",
         "sous_titre": "Les cotisations comptabilisées progressent de 46 % entre 2020 et 2025",
         "synthese": f"Cotisations maladie et décès : {M(SIN[0]['cot_maladie'] + SIN[0]['cot_deces'])} M FCFA en 2020, {M(s25['cot_maladie'] + s25['cot_deces'])} M en 2025"},
        graphique={"titre": fig("Cotisations comptabilisées par garantie"), "type": "barres_empilees",
                   "categories": [str(s["exercice"]) if s["exercice"] < 2026 else "2026 (8 mois)" for s in SIN],
                   "series": [{"nom": "Maladie", "valeurs": [round(s["cot_maladie"] / 1e6) for s in SIN], "couleur": "112A40"},
                              {"nom": "Décès", "valeurs": [round(s["cot_deces"] / 1e6) for s in SIN], "couleur": "EFAC62"}],
                   "format": "# ##0", "unite": "En millions de FCFA"},
        texte=[f"**Les cotisations suivent la hausse des effectifs et des salaires.** Elles passent de {M(SIN[0]['cot_maladie'] + SIN[0]['cot_deces'])} M FCFA en 2020 à {M(s25['cot_maladie'] + s25['cot_deces'])} M en 2025.",
               f"**En 2025, la comptabilité et les précomptes concordent en maladie à 7 % près.** Comptabilité : {M(s25['cot_maladie'])} M, précomptes maladie et retraités : {M(cot25['prec_936'] + cot25['prec_cgrae'])} M.",
               f"**Fin 2024 et fin 2025, l'État doit des cotisations au FPPN.** Le compte de cotisations à recevoir du Trésor atteint {M(an(BIL, 'exercice', 2024)['cotisations_a_recevoir_etat'])} M puis {M(b25['cotisations_a_recevoir_etat'])} M FCFA, il est soldé en 2026."],
        source=SRC_COMPTA)

# --- Partie 4 : sinistralité et résultat ----------------------------------
separateur(4, "Sinistralité et résultat", "Prestations, frais de gestion et résultat comptable")
cats = [str(s["exercice"]) if s["exercice"] < 2026 else "2026 (8 mois)" for s in SIN]
contenu({"titre": "Sinistralité comptable",
         "sous_titre": "Les prestations maladie absorbent désormais presque toutes les cotisations maladie",
         "synthese": f"Prestations / cotisations maladie : {pc(s21['sc_maladie'])} en 2021, {pc(s25['sc_maladie'])} en 2025 et {pc(s26['sc_maladie'])} sur huit mois de 2026"},
        graphique={"titre": fig("Prestations comptabilisées rapportées aux cotisations (S/C)"), "type": "lignes",
                   "categories": cats,
                   "series": [{"nom": "Maladie", "valeurs": [round(100 * s["sc_maladie"], 1) for s in SIN], "couleur": "112A40"},
                              {"nom": "Décès (prime d'assurance incluse en 2020)", "valeurs": [round(100 * s["sc_deces_avec_prime"], 1) for s in SIN], "couleur": "EFAC62"}],
                   "format": "0", "unite": "En % des cotisations de la garantie"},
        texte=["**Le ratio maladie a été multiplié par douze en quatre ans.** Il passe de 6 % en 2021 à 74 % en 2025, puis 97 % sur les huit premiers mois de 2026.",
               "**Les ratios très bas de 2020 à 2022 sont suspects.** Ils laissent penser que des factures de ces années n'ont pas été comptabilisées à temps, ce qui est cohérent avec l'apurement d'arriérés en 2026.",
               f"**En décès, le ratio est stable autour de 20 % depuis 2022.** En 2020, le FPPN payait une prime d'assurance de {M(SIN[0]['prime_assurance_deces'])} M FCFA (82 % des cotisations décès), en 2021, les obsèques atteignent 55 %.",
               "**Seule une base des prestations par date de soins permettra de mesurer le coût réel de chaque année.**"],
        source=SRC_COMPTA)

nat = R["prestations_maladie_nature"]
contenu({"titre": "Prestations maladie",
         "sous_titre": "Plus de la moitié des prestations maladie est enregistrée sans détail de nature",
         "synthese": f"En 2025, le compte « charges de prestations maladie » porte {M(nat[5]['2025'])} M FCFA sur {M(s25['prest_maladie'])} M de prestations"},
        graphique={"titre": fig("Prestations maladie comptabilisées par nature"), "type": "barres_empilees",
                   "categories": cats,
                   "series": [{"nom": n["nature"], "valeurs": [round(n[str(e)] / 1e6) for e in range(2020, 2027)]} for n in nat],
                   "format": "# ##0", "etiquettes": False, "legende": "droite", "unite": "En millions de FCFA"},
        texte=[f"**Les prestations maladie passent de {M(s21['prest_maladie'])} M FCFA en 2021 à {M(s25['prest_maladie'])} M en 2025.** La hausse touche toutes les natures de soins.",
               "**L'imagerie, les analyses et les consultations progressent le plus.** Elles représentent ensemble 37 % des prestations de 2025.",
               "**Un compte global de prestations maladie concentre plus de la moitié des montants.** Il empêche de suivre la consommation par type de soins et par prestataire."],
        source=SRC_COMPTA)

contenu({"titre": "Frais de gestion et résultat",
         "sous_titre": "Les excédents se réduisent et deviennent un déficit en 2026",
         "synthese": f"Résultat net : {M(SIN[0]['resultat'])} M FCFA en 2020, {M(s25['resultat'])} M en 2025, {M(s26['resultat'])} M sur huit mois de 2026 (provisoire)"},
        graphique={"titre": fig("Frais de gestion et résultat net"), "type": "barres",
                   "categories": cats,
                   "series": [{"nom": "Frais de gestion", "valeurs": [round(s["frais_gestion"] / 1e6) for s in SIN], "couleur": "BBBCBC"},
                              {"nom": "Résultat net", "valeurs": [round(s["resultat"] / 1e6) for s in SIN], "couleur": "112A40"}],
                   "format": "# ##0", "unite": "En millions de FCFA"},
        texte=[f"**Les frais de gestion pèsent entre {pc(min(s['taux_frais'] for s in SIN[1:6]))} et {pc(max(s['taux_frais'] for s in SIN[1:6]))} des cotisations depuis 2021.** Ils se situent autour d'un milliard de FCFA par an.",
               "**Le résultat net baisse depuis 2021.** La hausse des prestations maladie réduit l'excédent de 2 084 M FCFA en 2021 à 438 M en 2025.",
               "**Les huit premiers mois de 2026 sont déficitaires.** Les amortissements et provisions de fin d'année n'étant pas encore passés, le déficit annuel sera plus lourd."],
        source=SRC_COMPTA + " ; frais de gestion : comptes 61 à 66 hors prime d'assurance décès")

# --- Partie 5 : épargne ----------------------------------------------------
separateur(5, "Épargne des mutualistes", "PEP et EFC : dette envers les mutualistes et actifs dédiés")
yrs = [str(b["exercice"]) if b["exercice"] < 2026 else "août 2026" for b in BIL]
contenu({"titre": "Adossement de l'épargne",
         "sous_titre": "L'épargne des mutualistes n'est couverte qu'à 16 % par des actifs dédiés",
         "synthese": f"Au 31/08/2026 : {Md(b26['epargne_totale'])} Md FCFA d'épargne PEP et EFC pour {Md(b26['actifs_dedies_epargne'])} Md de banques, dépôts à terme et fonds dédiés"},
        graphique={"titre": fig("Épargne au passif et actifs dédiés"), "type": "barres",
                   "categories": yrs,
                   "series": [{"nom": "Épargne PEP et EFC (passif)", "valeurs": [round(b["epargne_totale"] / 1e6) for b in BIL], "couleur": "112A40"},
                              {"nom": "Actifs dédiés (banques, DAT, FCP)", "valeurs": [round(b["actifs_dedies_epargne"] / 1e6) for b in BIL], "couleur": "EFAC62"}],
                   "format": "# ##0", "unite": "En millions de FCFA"},
        texte=[f"**La dette d'épargne augmente de {pc(b26['epargne_totale'] / b20['epargne_totale'] - 1)} depuis 2020.** L'EFC triple ({M(b20['epargne_efc'])} M à {M(b26['epargne_efc'])} M FCFA), le PEP baisse puis remonte.",
               f"**Les actifs dédiés n'en couvrent qu'une faible part.** Le taux de couverture varie entre {pc(min(b['couverture_epargne'] for b in BIL))} et {pc(max(b['couverture_epargne'] for b in BIL))} selon les années.",
               "**L'épargne finance donc d'autres emplois.** Terrains, investissements au CMP, créance sur l'Hôpital de la Police et prêts du Grenier absorbent les ressources.",
               "**Des retraits importants exposent le FPPN à un manque de liquidités.** Les conditions de retrait des produits doivent être connues pour mesurer ce risque."],
        source=SRC_COMPTA + " (comptes 41195100, 41195200, 5211020, 5221020, 5211030, 5221030, 5330)")

pv = [p for p in R["pep_versements"] if p["annee"] >= 2020]
ev = {e["annee"]: e for e in R["efc_versements"]}
contenu({"titre": "Versements d'épargne",
         "sous_titre": "Les versements au PEP ont été multipliés par quatre depuis 2022",
         "synthese": f"En 2025 : {M(pv[5]['versements'])} M FCFA versés au PEP par {nb(pv[5]['epargnants'])} épargnants et {M(ev[2025]['versements'])} M à l'EFC (onze mois)"},
        graphique={"titre": fig("Versements précomptés par produit d'épargne"), "type": "barres",
                   "categories": [str(p["annee"]) if p["annee"] < 2026 else "2026 (8 mois)" for p in pv],
                   "series": [{"nom": "PEP 998", "valeurs": [round(p["versements"] / 1e6) for p in pv], "couleur": "112A40"},
                              {"nom": "EFC 870", "valeurs": [round(ev[p["annee"]]["versements"] / 1e6) if p["annee"] in ev else None for p in pv], "couleur": "EFAC62"}],
                   "format": "# ##0", "unite": "En millions de FCFA (EFC : données à partir de 2023)"},
        texte=[f"**Le nombre d'épargnants PEP passe de {nb(pv[0]['epargnants'])} en 2020 à {nb(pv[-1]['epargnants'])} en 2026.** La collecte s'accélère à partir de 2023.",
               f"**Les comptes d'épargne augmentent moins vite que les versements.** En 2025, le PEP reçoit {M(pv[5]['versements'])} M FCFA mais le compte ne progresse que de {M(b25['epargne_pep'] - an(BIL, 'exercice', 2024)['epargne_pep'])} M : la différence correspond à des sorties (échéances, retraits).",
               "**Le poste 998 porte aussi d'autres références (PTI, RPEP, EP).** Leur nature doit être précisée : épargne, remboursement de prêts ou autre produit.",
               "**Sans comptes individuels, la dette réelle envers chaque épargnant ne peut pas être vérifiée.**"],
        source=SRC_PREC + " (PEP 2020 à 2022 : fichiers texte de la Solde)")

# --- Partie 6 : Grenier -----------------------------------------------------
separateur(6, "Prêts du Grenier", "Encours, production, retards et précomptes")
ge_fa = [g for g in GE if g["arrete"].startswith("31/12") or g["arrete"] == "31/08/2026"]
contenu({"titre": "Encours du Grenier",
         "sous_titre": "L'encours de prêts a été multiplié par 132 depuis fin 2021",
         "synthese": f"Au 31/08/2026 : {nb(GP['prets'])} prêts, {M(GP['encours'])} M FCFA d'encours, dont {M(b26['grenier_financement'])} M financés par le FPPN"},
        graphique={"titre": fig("Encours des prêts et financement apporté par le FPPN"), "type": "barres",
                   "categories": [g["arrete"] for g in ge_fa],
                   "series": [{"nom": "Encours des prêts", "valeurs": [round(g["encours"] / 1e6) for g in ge_fa], "couleur": "112A40"},
                              {"nom": "Financement FPPN (compte 47111102)", "valeurs": [round(g["financement_fppn_47111102"] / 1e6) if g["financement_fppn_47111102"] is not None else None for g in ge_fa], "couleur": "EFAC62"}],
                   "format": "# ##0", "unite": "En millions de FCFA"},
        texte=[f"**L'activité de prêt démarre en 2021 et croît très vite.** L'encours passe de 7,9 M FCFA (17 prêts) fin 2021 à {M(GP['encours'])} M au 31/08/2026.",
               f"**Le FPPN finance l'essentiel de cet encours.** Sa créance sur le Grenier atteint {M(b26['grenier_financement'])} M FCFA, soit {pc(b26['grenier_financement'] / GP['encours'])} de l'encours.",
               f"**Les prêts sont courts et chers.** Durée médiane de {nb(GP['duree_mediane'])} mois, taux médian de {nb(GP['taux_median'])} %, encours réparti sur {nb(GP['emprunteurs'])} emprunteurs (les dix premiers : {pc(GP['top10'], 1)})."],
        source=SRC_GRENIER + " ; " + SRC_COMPTA.replace("Source : ", ""))

gpr = R["grenier_production"]
contenu({"titre": "Production de prêts",
         "sous_titre": "Les crédits scolaires et habitat portent la croissance de la production",
         "synthese": f"Montants débloqués : {M(gpr['total'][1])} M FCFA en 2022, {M(gpr['total'][4])} M en 2025 ({nb(gpr['nombre'][4])} prêts)"},
        graphique={"titre": fig("Montants débloqués par type de crédit"), "type": "barres_empilees",
                   "categories": [str(a) if a < 2026 else "2026 (8 mois)" for a in gpr["annees"]],
                   "series": [{"nom": t.capitalize(), "valeurs": [round(v / 1e6) for v in gpr["valeurs"][t]]} for t in gpr["types"]],
                   "format": "# ##0", "etiquettes": False, "legende": "droite", "unite": "En millions de FCFA"},
        texte=["**La production annuelle a été multipliée par six entre 2022 et 2025.** Elle atteint 1 609 M FCFA en 2025.",
               "**Crédits scolaires et habitat dominent.** En 2025, ils représentent 66 % des montants débloqués, le crédit flash, plus petit, est le plus nombreux.",
               "**Le crédit habitat allonge les risques.** Les montants unitaires sont plus importants, alors que la politique de crédit de 2021 transmise est en ancien format et n'a pas pu être lue."],
        source=SRC_GRENIER + " (2026 : janvier à août, portefeuilles CRF, CRO et CRS)")

ph = R["grenier_par_historique"]
contenu({"titre": "Retards de remboursement",
         "sous_titre": "Les retards ont culminé début 2026 avant de revenir à un niveau faible",
         "synthese": f"Encours en retard de plus de 30 jours : {pc(ph[5]['par30'])} fin janvier 2026, {pc(GP['par30'], 1)} au 31/08/2026"},
        graphique={"titre": fig("Part de l'encours en retard de plus de 30 et de plus de 90 jours"), "type": "lignes",
                   "categories": [p["arrete"] for p in ph if p["arrete"] != "30/09/2026"],
                   "series": [{"nom": "Plus de 30 jours", "valeurs": [round(100 * p["par30"], 1) for p in ph if p["arrete"] != "30/09/2026"], "couleur": "112A40"},
                              {"nom": "Plus de 90 jours", "valeurs": [round(100 * p["par90"], 1) for p in ph if p["arrete"] != "30/09/2026"], "couleur": "EFAC62"}],
                   "format": "0", "etiquettes": False, "unite": "En % de l'encours"},
        texte=["**Fin janvier 2026, 62 % de l'encours avait plus de 30 jours de retard.** La plupart de ces retards étaient compris entre 31 et 90 jours.",
               "**Ces retards se sont résorbés en quelques mois.** La part en retard de plus de 30 jours revient à 1,8 % de l'encours au 31/08/2026.",
               "**Le profil fait penser à des échéances non précomptées en fin d'année 2025.** La cause doit être confirmée par le Grenier et la Solde.",
               "**Les prêts en retard de plus de 180 jours (126 prêts) doivent être provisionnés.**"],
        source=SRC_GRENIER)

ech = [e for e in R["grenier_echeances"] if e["precompte_9x4"]]
contenu({"titre": "Échéances et précomptes",
         "sous_titre": "Les précomptes de la Solde ne couvrent que 76 % à 85 % des échéances dues",
         "synthese": "De janvier à août 2026, les échéances des prêts en cours dépassent chaque mois les précomptes 9X4 de 25 à 45 M FCFA"},
        graphique={"titre": fig("Échéances des prêts en cours et précomptes 9X4, 2026"), "type": "barres",
                   "categories": [e["mois"] for e in ech],
                   "series": [{"nom": "Échéances (balance âgée)", "valeurs": [round(e["echeances_balance"] / 1e6, 1) for e in ech], "couleur": "112A40"},
                              {"nom": "Précomptes 9X4 (Solde)", "valeurs": [round(e["precompte_9x4"] / 1e6, 1) for e in ech], "couleur": "EFAC62"}],
                   "format": "0", "unite": "En millions de FCFA"},
        texte=["**Une part des échéances n'est pas retenue sur les salaires.** L'écart mensuel représente 15 % à 24 % des échéances dues.",
               "**Cet écart peut avoir plusieurs causes.** Remboursements payés autrement (caisse, banque), décalage du premier précompte, ou rejets de la Solde.",
               "**L'identifiant des emprunteurs complique le rapprochement.** L'export du logiciel de prêts perd la lettre-clé du mécano (321100C devient 321100).",
               "**Un rapprochement prêt par prêt est nécessaire** avec les encaissements bancaires du Grenier."],
        source=SRC_GRENIER + " ; " + SRC_PREC.replace("Source : ", ""))

# --- Partie 7 : retraités et produits annexes -----------------------------
separateur(7, "Retraités et annexes", "CGRAE, flotte automobile, MEDOC")
contenu({"titre": "Retraités CGRAE",
         "sous_titre": "Les retraités cotisants sont passés de 16 à 610 en deux ans",
         "synthese": f"Précompte moyen de {nb(round(sum(c['moyen'] for c in CG) / len(CG)))} FCFA par mois sur la pension, données disponibles depuis février 2024"},
        graphique={"titre": fig("Retraités précomptés par la CGRAE, par mois"), "type": "lignes",
                   "categories": [c["libelle"] for c in CG], "series": [{"nom": "Retraités", "valeurs": [c["retraites"] for c in CG], "couleur": "112A40"}],
                   "format": "0", "etiquettes": False, "legende": None, "unite": "En nombre de retraités"},
        texte=["**Le précompte sur pension démarre en 2024 et monte en charge.** 610 retraités sont précomptés en août 2026.",
               "**Les retraités consomment davantage de soins que les actifs.** Leur cotisation (environ 11 900 FCFA par mois) doit être comparée à leur coût santé dès réception des factures.",
               "**Plusieurs anomalies de gestion sont visibles.** Précomptes en double ou indus remboursés, numéros de décision erronés, précomptes non effectués.",
               "**La convention avec la CGRAE n'a pas été reçue.**"],
        source=SRC_PREC + " (fichiers retour de la CGRAE)")

ann = R["produits_annexes"]
contenu({"titre": "Produits annexes",
         "sous_titre": "Deux produits sont précomptés sans documentation transmise",
         "synthese": "Flotte automobile (9W3) : 75 à 112 adhérents par an, MEDOC (8Z7) : 5 000 FCFA par mois pour 118 à 145 adhérents"},
        tableau={"titre": tab("Produits annexes et points à clarifier"),
                 "entetes": ["Produit", "Ce que montrent les données", "À obtenir du FPPN"],
                 "largeurs": [5.0, 13.6, 12.6],
                 "lignes": [
                     ["Flotte automobile GNA (9W3)", "Précomptes mensuels de 2023 à 2026, environ 10 000 FCFA par adhérent, montants reversés à l'assureur non connus", "Convention avec l'assureur, primes reversées, sinistres, rôle et rémunération du FPPN"],
                     ["MEDOC (8Z7)", "Précompte de 5 000 FCFA par mois depuis novembre 2025, 121 contrats individuels signés en novembre et décembre 2025, retours vides en février et avril 2026", "Nature du produit, garanties, bénéficiaires, lien avec la santé, comptabilisation (compte 70100001)"],
                     ["Poste 998, références autres que PEP", "Références PTI, RPEP, EP, TI sur le même poste que l'épargne PEP", "Nature de chaque référence (prêt, épargne, autre)"]]},
        source=SRC_PREC)

# --- Partie 8 : bilan -----------------------------------------------------
separateur(8, "Bilan et placements", "Provisions, structure de l'actif, fonds propres")
contenu({"titre": "Provisions techniques",
         "sous_titre": "Les provisions pour prestations à payer sont figées depuis 2020",
         "synthese": f"{M(b20['provision_maladie'])} M FCFA en maladie et {M(b20['provision_obseques'])} M en obsèques, sans changement de 2020 à 2026"},
        graphique={"titre": fig("Provisions pour prestations à payer et prestations de l'année"), "type": "barres",
                   "categories": cats,
                   "series": [{"nom": "Provisions maladie et obsèques", "valeurs": [round(b["provisions_techniques"] / 1e6) for b in BIL], "couleur": "112A40"},
                              {"nom": "Prestations maladie et décès de l'année", "valeurs": [round((s["prest_maladie"] + s["prest_deces"]) / 1e6) for s in SIN], "couleur": "EFAC62"}],
                   "format": "# ##0", "unite": "En millions de FCFA"},
        texte=["**Le montant des provisions est identique à chaque clôture.** Seul le numéro de compte change en 2024 (passage du compte 19 au compte 158).",
               f"**Pendant ce temps, les prestations ont été multipliées par quatre.** Elles passent de {M(s21['prest_maladie'] + s21['prest_deces'])} M FCFA en 2021 à {M(s25['prest_maladie'] + s25['prest_deces'])} M en 2025.",
               "**Une provision qui ne bouge pas n'est pas une estimation.** Elle ne reflète ni les dossiers en attente, ni les factures non reçues, ni les sinistres non déclarés.",
               "**Le calcul (PSAP, Chain Ladder) suppose les bases de prestations demandées.**"],
        source=SRC_COMPTA + " (comptes 19011000, 19012000, 15801000, 15802000)")

contenu({"titre": "Structure de l'actif",
         "sous_titre": "L'actif repose sur des créances et des investissements peu liquides",
         "synthese": f"Au 31/08/2026 : CMP {M(b26['cmp_investissements'])} M, terrains {M(b26['terrains'])} M, Hôpital de la Police {M(b26['creance_hopital_police'])} M, trésorerie {M(b26['tresorerie_totale'])} M FCFA"},
        graphique={"titre": fig("Principaux postes de l'actif"), "type": "barres_empilees",
                   "categories": yrs,
                   "series": [{"nom": "Investissements CMP", "valeurs": [round(b["cmp_investissements"] / 1e6) for b in BIL], "couleur": "112A40"},
                              {"nom": "Terrains", "valeurs": [round(b["terrains"] / 1e6) for b in BIL], "couleur": "53565A"},
                              {"nom": "Créance Hôpital de la Police", "valeurs": [round(b["creance_hopital_police"] / 1e6) for b in BIL], "couleur": "BBBCBC"},
                              {"nom": "Financement du Grenier", "valeurs": [round(b["grenier_financement"] / 1e6) for b in BIL], "couleur": "EFAC62"},
                              {"nom": "Trésorerie", "valeurs": [round(b["tresorerie_totale"] / 1e6) for b in BIL], "couleur": "060606"}],
                   "format": "# ##0", "etiquettes": False, "legende": "droite", "unite": "En millions de FCFA"},
        texte=[f"**Les investissements au CMP sont passés de {M(b20['cmp_investissements'])} M à {M(b26['cmp_investissements'])} M FCFA.** Ils représentent le premier emploi des ressources.",
               f"**La créance sur l'Hôpital de la Police ({M(b26['creance_hopital_police'])} M FCFA) n'a pas bougé depuis 2020.** Sans remboursement ni provision, sa valeur est incertaine.",
               f"**La trésorerie reste faible au regard des dettes exigibles.** {M(b26['tresorerie_totale'])} M FCFA au 31/08/2026, pour {Md(b26['epargne_totale'])} Md d'épargne et {M(b26['fournisseurs_cliniques'])} M de dettes envers les cliniques."],
        source=SRC_COMPTA)

contenu({"titre": "Fonds propres",
         "sous_titre": "Les fonds propres restent négatifs malgré six années d'excédents",
         "synthese": f"Fonds propres : {M(b20['fonds_propres'])} M FCFA fin 2020, {M(b25['fonds_propres'])} M fin 2025, {M(b26['fonds_propres'])} M au 31/08/2026 (provisoire)"},
        graphique={"titre": fig("Fonds propres et résultat de l'exercice"), "type": "barres",
                   "categories": yrs,
                   "series": [{"nom": "Fonds propres (résultat inclus)", "valeurs": [round(b["fonds_propres"] / 1e6) for b in BIL], "couleur": "112A40"},
                              {"nom": "Résultat de l'exercice", "valeurs": [round(b["resultat_exercice"] / 1e6) for b in BIL], "couleur": "EFAC62"}],
                   "format": "# ##0", "unite": "En millions de FCFA"},
        texte=[f"**Les pertes passées pèsent encore sur le bilan.** Le compte de pertes atteint {M(-b20['pertes_cumulees'])} M FCFA en 2020 et {M(-b25['pertes_cumulees'])} M en 2025.",
               "**Les excédents de 2020 à 2025 réduisent l'écart sans le combler.** Les fonds propres remontent de 6,7 Md FCFA à 0,4 Md FCFA négatifs.",
               "**Le déficit de 2026 inverse la tendance.** Les fonds propres reculent à nouveau sur les huit premiers mois.",
               "**Des fonds propres négatifs signifient que les dettes dépassent les actifs.** Le respect des ratios du règlement UEMOA devra être vérifié sur le texte intégral."],
        source=SRC_COMPTA + " (comptes 10 à 13 et résultat de l'exercice)")

rp = R["rendement_placements"]
contenu({"titre": "Rendement des placements",
         "sous_titre": "Le problème des placements est leur faible volume, pas leur rendement",
         "synthese": "Rendement moyen apparent des dépôts à terme et fonds de 2021 à 2025 : 6,5 % par an, pour un taux directeur BCEAO de 2 % à 3,5 %"},
        graphique={"titre": fig("Taux directeur BCEAO"), "type": "lignes",
                   "categories": [t[0] for t in R["externe"]["taux_directeur"]],
                   "series": [{"nom": "Taux directeur", "valeurs": [t[1] for t in R["externe"]["taux_directeur"]], "couleur": "112A40"}],
                   "format": "0.00", "legende": None, "unite": "En %, date d'effet"},
        texte=["**Les produits financiers sont faibles en montant.** Ils totalisent 174 M FCFA de 2021 à 2025 pour un encours moyen de placements d'environ 540 M FCFA.",
               "**Le rendement moyen apparent est de 6,5 % par an.** Il dépasse le taux directeur, mais il est irrégulier : la plus grande partie des intérêts est enregistrée en 2025, à la clôture d'un dépôt à terme.",
               "**Le vrai sujet est le volume placé.** Environ 0,5 Md FCFA de placements face à 15,7 Md FCFA d'épargne à rendre aux mutualistes."],
        source="Source : BCEAO, communiqués du Comité de politique monétaire (données validées par Adenka) ; " + SRC_COMPTA.replace("Source : ", ""))

# --- Partie 9 : analyses en attente, manques et risques ----------------------
separateur(9, "Analyses en attente", "Documents à obtenir, risques, points à clarifier")
contenu({"titre": "Analyses en attente",
         "sous_titre": "Les analyses actuarielles centrales dépendent de trois bases non reçues",
         "synthese": "Provisions, mortalité, équilibre technique et contrôles santé seront réalisés dès réception des documents listés"},
        tableau={"titre": tab("Analyses en attente et documents nécessaires"),
                 "entetes": ["Analyse", "Objet", "Documents nécessaires (n° de la PBC)"],
                 "largeurs": [7.2, 12.4, 11.6],
                 "lignes": [
                     ["Provisions pour sinistres (PSAP, Chain Ladder)", "Montant restant à payer, y compris sinistres non déclarés", "Bases des prestations décès et santé, dossiers non payés (13, 14, 28, 29, 40)"],
                     ["Sinistralité par année de survenance", "Coût réel des prestations de chaque année", "Bases des prestations (13, 28, 40)"],
                     ["Mortalité des policiers (table CIMA H)", "Comparer la mortalité observée à la table réglementaire", "Mutualistes avec date de naissance et sexe (9), décès (28), table CIMA"],
                     ["Délais et dossiers en attente", "Délais de déclaration, d'instruction et de paiement", "Bases des prestations avec dates (28, 40)"],
                     ["Contrôles santé (S1 à S7)", "Doublons, bénéficiaires hors fichier, soins après sortie", "Base des factures santé (40), cartes actives (42)"],
                     ["Engagements d'épargne", "Dette envers chaque épargnant", "Comptes individuels PEP et EFC, règlement des produits (16, 17)"],
                     ["Ratios AIRMS et équilibre technique", "Solvabilité réglementaire, projection des cotisations et prestations", "Texte intégral du règlement UEMOA, âges et barèmes (1, 2, 3)"]]},
        source="Source : liste des documents demandés (PBC) pointée au 18/10/2026")

contenu({"titre": "Risques et points à clarifier",
         "sous_titre": "Sept points demandent une réponse du FPPN avant la prochaine version",
         "synthese": "Chaque point est relié à un constat chiffré du rapport et à une demande précise"},
        tableau={"titre": tab("Points à clarifier avec le FPPN"),
                 "entetes": ["N°", "Point", "Question au FPPN"],
                 "largeurs": [1.4, 11.0, 18.8],
                 "lignes": [
                     ["1", "Répartition du taux de 5 %", "Quel barème s'applique : 3,5 % / 1,5 % ou 3,77 % / 1,23 % ? Pourquoi la comptabilité suit-elle une autre clé que la Solde ?"],
                     ["2", "Provisions techniques figées", "Comment les provisions maladie et obsèques ont-elles été calculées ? Pourquoi n'ont-elles pas changé depuis 2020 ?"],
                     ["3", "Couverture de l'épargne", "Quelles sont les conditions de retrait du PEP et de l'EFC ? Comment le FPPN ferait-il face à des retraits importants ?"],
                     ["4", "Créance sur l'Hôpital de la Police", "Quelle est la nature de la créance de 3 408 M FCFA ? Un remboursement est-il prévu ?"],
                     ["5", "Retards du Grenier début 2026", "Pourquoi 62 % de l'encours était-il en retard fin janvier 2026 ? Comment sont réglées les échéances non précomptées ?"],
                     ["6", "Sinistralité 2020 à 2022", "Des factures de ces années ont-elles été comptabilisées plus tard ? Pour quels montants ?"],
                     ["7", "Données manquantes ou endommagées", "Transmettre les précomptes 2020 à 2023, les fichiers vides de décembre 2024 et les grands livres endommagés"]]},
        source="Source : analyses Adenka sur les documents reçus le 15/10/2026")

# --- Annexes ----------------------------------------------------------------
SL.append({"type": "separateur", "num": "A", "titre": "Annexes", "sous_titre": "Méthodes, outils, listes et références", "annexe": True})
SL.append({"type": "annexe", "titre": "Méthodes et formules", "tableau": {"titre": tab("Indicateurs, formules et sources"),
           "entetes": ["Indicateur", "Formule", "Source des données"], "largeurs": [7.0, 13.2, 11.0],
           "lignes": [["Ratio S/C", "Prestations comptabilisées de l'année / cotisations comptabilisées de l'année", "Prestations : comptes 601, 6013, 6018, 60800007, 6021, 6024. Cotisations : comptes 7010, 7021"],
                      ["Taux de cotisation", "Précompte du mois / salaire de base du mois, par agent", "Fichiers maladie 936 détaillés, décès 8Z6"],
                      ["Taux de précompte", "Montant retour (Solde, CGRAE) / montant aller (FPPN), mois communs", "Fichiers aller et retour"],
                      ["Entrées et sorties", "Matricules présents au mois m et absents au mois m-1 (et inversement)", "Fichiers maladie 936, mois consécutifs"],
                      ["Couverture de l'épargne", "(Banques, DAT et FCP dédiés) / (comptes PEP + EFC)", "Balances des comptes"],
                      ["Encours en retard (> 30 j)", "Encours des prêts en retard de plus de 30 jours / encours total", "Balances âgées du Grenier"],
                      ["Fonds propres", "Réserves + report à nouveau + pertes cumulées + résultat de l'exercice", "Comptes 10 à 13 et classes 6, 7, 8"],
                      ["Rendement apparent", "Produits financiers (77) / encours moyen des placements (52 à 53)", "Balances des comptes"]]},
           "source": "Source : Adenka"})
SL.append({"type": "annexe", "titre": "Outils et traçabilité", "tableau": {"titre": tab("Outils et bibliothèques utilisés"),
           "entetes": ["Outil", "Usage", "Version"], "largeurs": [7.0, 18.2, 6.0],
           "lignes": [["Python", "Langage des traitements, scripts conservés et rejouables", "3.13"],
                      ["pandas, numpy", "Tables de données, calculs, agrégations", "3.0 et 2.5"],
                      ["python-calamine, openpyxl", "Lecture des fichiers Excel sans perte", "dernière"],
                      ["lxml", "Lecture des exports au format page web du Grenier", "dernière"],
                      ["pyarrow", "Stockage de la base d'analyse (Parquet)", "26.0"],
                      ["pymupdf, pdftotext", "Lecture et récupération des PDF comptables", "dernière"],
                      ["python-pptx", "Rapport PowerPoint, graphiques et tableaux natifs", "1.0"],
                      ["Excel", "Classeur des calculs joint au rapport", "Microsoft 365"]]},
           "source": "Source : Adenka ; scripts du dépôt de la mission (dossier scripts)"})
SL.append({"type": "liste_figures"})
SL.append({"type": "liste_tableaux"})
SL.append({"type": "annexe", "titre": "Signes et abréviations", "tableau": {"titre": None,
           "entetes": ["Sigle", "Signification", "Sigle", "Signification"], "largeurs": [3.4, 12.2, 3.4, 12.2],
           "lignes": [["AIRMS", "Autorité de régulation de la mutualité sociale", "M / Md", "Millions / milliards de FCFA"],
                      ["BCEAO", "Banque centrale des États de l'Afrique de l'Ouest", "PAR", "Part de l'encours en retard"],
                      ["CGRAE", "Caisse générale de retraite des agents de l'État", "PBC", "Liste des documents demandés au client"],
                      ["CIMA", "Conférence interafricaine des marchés d'assurances", "PEP", "Plan d'épargne (poste 998)"],
                      ["CMP", "Centre médical de la Police (investissements)", "PSAP", "Provisions pour sinistres à payer"],
                      ["CRF, CRO, CRS", "Sous-portefeuilles du logiciel de prêts (sens à préciser)", "S/C", "Prestations / cotisations"],
                      ["DAT", "Dépôt à terme", "UEMOA", "Union économique et monétaire ouest-africaine"],
                      ["EFC", "Entraide fin de carrière (poste 870)", "936", "Poste de précompte maladie"],
                      ["FCP", "Fonds commun de placement", "8Z6", "Poste de précompte décès"],
                      ["FOMUS-CI", "Fonds de garantie des mutuelles sociales", "9X4", "Poste de remboursement des prêts du Grenier"],
                      ["FPPN", "Fonds de prévoyance de la Police nationale", "9W3", "Poste flotte automobile (GNA)"],
                      ["GNA", "Assureur de la flotte automobile", "8Z7", "Poste MEDOC"]]},
           "source": None})
SL.append({"type": "annexe", "titre": "Références et sources", "tableau": {"titre": None,
           "entetes": ["Référence", "Contenu utilisé", "Lien ou emplacement"], "largeurs": [8.6, 10.6, 12.0],
           "lignes": [["Envoi du FPPN du 15/10/2026", "Précomptes, balances comptables, exports du Grenier, pièces CGRAE", "Dossier 04-Documents reçus du client (inventaire du 16/10/2026)"],
                      ["Grands livres 2021, 2022, 2024, 2025 (PDF)", "Détail des écritures (2022 partiellement lisible)", "Envoi complémentaire du 18/10/2026"],
                      ["Règlement n° 07/2009/CM/UEMOA", "Cadre de la mutualité sociale", "flagship.social-protection.org (Media.action?id=9421)"],
                      ["Ordonnance n° 2024-1043", "FOMUS-CI", "civlii.laws.africa (ord. 2024/1043)"],
                      ["Code CIMA, article 338", "Taux technique maximal de 3,5 %", "cima-afrique.org (bulletin officiel, 15e édition)"],
                      ["BCEAO, communiqués du CPM 2020 à 2025", "Taux directeurs", "bceao.int (communiqués de presse)"],
                      ["Inflation 2020 à 2024", "Contexte des prix (non utilisé dans les ratios)", "INS, economie-ivoirienne.ci, Banque mondiale (FRED)"],
                      ["Classeur des sources externes validé", "Liste et validation des données externes", "FPPN_Sources_externes_a_valider_20261018.xlsx"],
                      ["Classeur des calculs", "Tous les tableaux et séries du rapport", "FPPN_Calculs_rapport.xlsx"]]},
           "source": None})
SL.append({"type": "merci"})


# ---------------------------------------------------------------------------
# Construction
# ---------------------------------------------------------------------------
GAB = {"couverture": 1, "sommaire": 5, "separateur": 2, "contenu": 5, "annexe": 5,
       "liste_figures": 5, "liste_tableaux": 5, "merci": 110}
tmp = SORTIE.replace(".pptx", "_tmp.pptx")
selection(TIMESAVER, tmp, [GAB[s["type"]] for s in SL])

C = {}
for i, s in enumerate(SL, 1):
    k = str(i)
    if s["type"] == "couverture":
        C[k] = {"Google Shape;138;p25": "Rapport actuariel\net statistique\ndu FPPN", "Text Placeholder 4": f"Rapport d'étape au {DATE_RAPPORT}"}
    elif s["type"] == "merci":
        pass
    elif s["type"] == "separateur":
        C[k] = {"Rectangle 24": str(s["num"]), "Rectangle 9": s["titre"], "Rectangle 5": s["sous_titre"]}
    elif s["type"] in ("sommaire", "liste_figures", "liste_tableaux", "annexe"):
        C[k] = {"@vider": True}
    elif s["type"] == "contenu":
        d = {"@vider": True, "@entete": s["entete"]}
        if "graphique" in s:
            g = dict(s["graphique"])
            d["@texte1"] = {"texte": g.pop("titre"), "style": "titre_graphique", "zone": "titre_graphique"}
            if g.get("unite"):
                d["@texte3"] = {"texte": g.pop("unite"), "style": "unite", "x": 1.4, "y": 4.45, "largeur": 18.2, "hauteur": 0.45}
            g["zone"] = "graphique_gauche"
            g["y"] = 4.95
            g["hauteur"] = 11.9
            d["@graphique"] = g
        if "texte" in s:
            d["@texte2"] = {"texte": "\n".join(s["texte"]), "style": "corps", "zone": "texte_droite"}
        if s.get("source"):
            d["@source"] = s["source"]
        C[k] = d
remplir(tmp, C, tmp)

prs = Presentation(tmp)
W, H = prs.slide_width, prs.slide_height
NOIR, GRIS, BLEU, ORANGE = "060606", "F2F2F2", "112A40", "EFAC62"
ts = Presentation(TIMESAVER)
icones = [sh for sh in ts.slides[110].shapes if sh.shape_type == 13 and sh.width < Cm(3)]


def icone_blanche(blob):
    """Pictogramme Adinkra en blanc : même forme (canal alpha), remplissage blanc."""
    from PIL import Image
    im = Image.open(io.BytesIO(blob)).convert("RGBA")
    blanc = Image.new("RGBA", im.size, (255, 255, 255, 0))
    alpha = im.split()[3]
    if alpha.getextrema() == (255, 255):           # pas de transparence : seuil sur la luminosité
        alpha = im.convert("L").point(lambda v: 255 if v < 128 else 0)
    blanc.putalpha(alpha)
    out = io.BytesIO()
    blanc.save(out, "PNG")
    return out.getvalue()


def txt(slide, x, y, l, h, texte, taille=14, gras=False, couleur=NOIR, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(l), Cm(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Cm(0.15)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = texte
    r.font.name = "Helvetica"
    r.font.size = Pt(taille)
    r.font.bold = gras
    r.font.color.rgb = RGBColor.from_string(couleur)
    return tb


def tableau(slide, spec, y0=3.85):
    """Tableau natif Adenka : en-tête noir texte blanc, lignes alternées gris clair, Helvetica."""
    y = y0
    if spec.get("titre"):
        txt(slide, 1.4, y, 31.2, 0.6, spec["titre"], 12, False, ORANGE, PP_ALIGN.CENTER)
        y += 0.75
    lignes = spec["lignes"]
    larg = spec["largeurs"]
    n = len(lignes) + 1
    haut_dispo = 16.95 - y
    total_car = max(len(" ".join(l)) for l in lignes)
    taille = 12 if n <= 9 and total_car < 200 else (11 if n <= 11 else 10)
    shape = slide.shapes.add_table(n, len(larg), Cm(1.4), Cm(y), Cm(sum(larg)), Cm(min(haut_dispo, 0.9 * n)))
    t = shape.table
    tblPr = t._tbl.tblPr
    for attr in ("bandRow", "firstRow"):
        tblPr.set(attr, "0")
    st = tblPr.find(qn("a:tableStyleId"))
    if st is not None:
        st.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"   # « Aucun style, quadrillage »
    for j, w in enumerate(larg):
        t.columns[j].width = Cm(w)
    for i in range(n):
        vals = spec["entetes"] if i == 0 else lignes[i - 1]
        for j, v in enumerate(vals):
            c = t.cell(i, j)
            c.margin_left = c.margin_right = Cm(0.15)
            c.margin_top = c.margin_bottom = Cm(0.06)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = c.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = ""
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", "0")
            pPr.set("indent", "0")
            pPr.set("lvl", "0")
            for b in pPr.findall(qn("a:buNone")):
                pPr.remove(b)
            pPr.insert(0, pPr.makeelement(qn("a:buNone"), {}))
            r = p.add_run()
            r.text = str(v)
            r.font.name = "Helvetica"
            r.font.size = Pt(taille)
            r.font.bold = (i == 0) or (j == 0 and spec.get("premiere_colonne_grasse", True))
            r.font.color.rgb = RGBColor(255, 255, 255) if i == 0 else RGBColor.from_string(NOIR)
            c.fill.solid()
            c.fill.fore_color.rgb = RGBColor.from_string(NOIR if i == 0 else (GRIS if i % 2 == 0 else "FFFFFF"))
    return shape


def etapes(slide, items):
    y = 4.3
    for i, (num, titre, corps) in enumerate(items):
        x = 1.6 + (i % 2) * 15.8
        yy = y + (i // 2) * 6.2
        b = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(x), Cm(yy), Cm(2.2), Cm(1.6))
        b.fill.solid()
        b.fill.fore_color.rgb = RGBColor.from_string(BLEU)
        b.line.fill.background()
        tf = b.text_frame
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        r = tf.paragraphs[0].add_run()
        r.text = num
        r.font.name = "Helvetica"
        r.font.size = Pt(20)
        r.font.color.rgb = RGBColor(255, 255, 255)
        txt(slide, x + 2.7, yy - 0.1, 12.2, 0.9, titre, 16, True, BLEU)
        txt(slide, x + 2.7, yy + 0.8, 12.2, 4.4, corps, 14)


def sommaire(slide, entrees):
    for sh in ts.slides[4].shapes:
        pass
    txt(slide, 2.7, 0.85, 20, 1.1, "SOMMAIRE", 24, True, BLEU)
    pas = 1.37 if len(entrees) <= 9 else 12.6 / len(entrees)
    taille = 16 if len(entrees) <= 9 else 14
    y0 = 4.02
    centres = [y0 + i * pas + 0.505 for i in range(len(entrees))]
    ligne = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(2.59), Cm(centres[0]), Cm(2.59), Cm(centres[-1]))
    ligne.line.color.rgb = RGBColor.from_string(BLEU)
    ligne.line.width = Pt(1)
    for i, (intitule, num) in enumerate(entrees):
        cy = centres[i]
        h = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(2.59), Cm(cy), Cm(4.85), Cm(cy))
        h.line.color.rgb = RGBColor.from_string(BLEU)
        h.line.width = Pt(1)
        dd = min(1.2, pas - 0.12)
        dsq = slide.shapes.add_shape(MSO_SHAPE.OVAL, Cm(4.85 + (1.2 - dd) / 2), Cm(cy - dd / 2), Cm(dd), Cm(dd))
        dsq.fill.solid()
        dsq.fill.fore_color.rgb = RGBColor.from_string(NOIR)
        dsq.line.fill.background()
        if icones:
            ic = icones[i % len(icones)]
            ti = dd * 0.6
            slide.shapes.add_picture(io.BytesIO(icone_blanche(ic.image.blob)), Cm(4.85 + 0.6 - ti / 2), Cm(cy - ti / 2), Cm(ti), Cm(ti))
        for x, l, t_, al in ((6.79, 22.98, intitule, PP_ALIGN.LEFT), (29.88, 1.53, str(num), PP_ALIGN.CENTER)):
            cadre = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(x), Cm(cy - 0.505), Cm(l), Cm(1.01))
            cadre.fill.background()
            cadre.line.color.rgb = RGBColor.from_string(NOIR)
            cadre.line.width = Pt(1.5)
            tf = cadre.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = Cm(0.3)
            p = tf.paragraphs[0]
            p.alignment = al
            r = p.add_run()
            r.text = t_
            r.font.name = "Helvetica"
            r.font.size = Pt(taille)
            r.font.color.rgb = RGBColor.from_string(NOIR)


# Entrées du sommaire : résumé exécutif puis chaque partie, annexes en dernier
entrees = [("Résumé exécutif", 3)]
for i, s in enumerate(SL, 1):
    if s["type"] == "separateur":
        entrees.append((s["titre"][0].upper() + s["titre"][1:], i))

from pptx.enum.chart import XL_TICK_LABEL_POSITION  # noqa: E402
slides = list(prs.slides)
for i, s in enumerate(SL, 1):
    sl = slides[i - 1]
    if s["type"] == "separateur":
        for sh in sl.shapes:
            if sh.name in ("Rectangle 9", "Rectangle 5"):
                sh.width = Cm(10.6)
    if s["type"] == "couverture":
        for sh in sl.shapes:
            if sh.name == "Google Shape;138;p25":
                sh.width = Cm(12.8)
            if sh.name == "Text Placeholder 4":
                sh.left, sh.width = Cm(25.2), Cm(7.3)
    for sh in sl.shapes:
        if getattr(sh, "has_chart", False) and sh.has_chart:
            ch = sh.chart
            try:
                negatif = any(v is not None and v < 0 for pl in ch.plots for se in pl.series for v in se.values)
                if negatif:
                    ch.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
            except Exception:
                pass
for i, s in enumerate(SL, 1):
    sl = slides[i - 1]
    if s["type"] == "sommaire":
        sommaire(sl, entrees)
    elif s["type"] == "contenu":
        if "tableau" in s:
            tableau(sl, s["tableau"])
        if "etapes" in s:
            etapes(sl, s["etapes"])
    elif s["type"] == "annexe":
        txt(sl, 1.52, 0.4, 31.24, 1.1, s["titre"], 20, True, NOIR)
        tableau(sl, s["tableau"], 2.0)
        if s.get("source"):
            txt(sl, 1.02, 17.15, 31.75, 0.51, s["source"], 7.5, False, "808080")
    elif s["type"] in ("liste_figures", "liste_tableaux"):
        liste = FIG if s["type"] == "liste_figures" else TAB
        nom = "Figure" if s["type"] == "liste_figures" else "Tableau"
        txt(sl, 1.52, 0.4, 31.24, 1.1, "Liste des figures" if nom == "Figure" else "Liste des tableaux", 20, True, NOIR)
        moitie = (len(liste) + 1) // 2
        for col, part in enumerate((liste[:moitie], liste[moitie:])):
            if not part:
                continue
            tableau(sl, {"titre": None, "entetes": [nom, "Titre", "Diapo."],
                         "largeurs": [1.9, 11.6, 1.6],
                         "lignes": [[str(n), t, str(d)] for n, t, d in part]}, 2.0) if col == 0 else None
            if col == 1:
                shp = tableau(sl, {"titre": None, "entetes": [nom, "Titre", "Diapo."], "largeurs": [1.9, 11.6, 1.6],
                                   "lignes": [[str(n), t, str(d)] for n, t, d in part]}, 2.0)
                shp.left = Cm(17.4)
prs.save(SORTIE)
finaliser(SORTIE)
os.remove(tmp)
json.dump({"figures": FIG, "tableaux": TAB, "diapositives": len(SL)},
          open(SORTIE.replace(".pptx", "_index.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(f"{len(SL)} diapositives ; {len(FIG)} figures ; {len(TAB)} tableaux -> {SORTIE}")
