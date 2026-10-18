# -*- coding: utf-8 -*-
"""
FPPN - Rapport actuariel et statistique, version 2 (PowerPoint, charte Adenka, Timesaver)
=========================================================================================

Entrées : resultats.json et complements.json (analyses), dossier des images de code, skill adenka-powerpoint.
Sorties : deck .pptx (graphiques et tableaux natifs, schémas modifiables, notes du présentateur) et
          fichier de traçabilité (chaque chiffre écrit : diapositive, valeur exacte, texte affiché).

Règles de présentation : pourcentages à deux décimales ; montants exacts en FCFA, sans décimale, dans
les textes et les tableaux ; graphiques : données exactes en FCFA, affichées en millions.
Utilisation : python rapport_v2_fppn.py <resultats.json> <complements.json> <dossier_images> <skill> <sortie.pptx>
"""
import io
import json
import os
import sys

RES, COMP, IMG, SKILL, SORTIE = sys.argv[1:6]
sys.path.insert(0, os.path.join(SKILL, "scripts"))
from pptx_tools import selection, remplir, finaliser  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.chart.data import CategoryChartData  # noqa: E402
from pptx.dml.color import RGBColor  # noqa: E402
from pptx.enum.chart import XL_TICK_LABEL_POSITION  # noqa: E402
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE  # noqa: E402
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402
from pptx.util import Cm, Pt  # noqa: E402

R = json.load(open(RES, encoding="utf-8"))
C = json.load(open(COMP, encoding="utf-8"))
TIMESAVER = os.path.join(SKILL, "assets", "Timesaver_Adenka.pptx")
NOIR, GRIS, BLEU, ORANGE, ROUGE, GRIS_F = "060606", "F2F2F2", "112A40", "EFAC62", "C00000", "53565A"

# ---------------------------------------------------------------------------
# Formats et traçabilité
# ---------------------------------------------------------------------------
TRACE = []
SL = []
COURANTE = None  # diapositive en cours de post-traitement (sinon : la prochaine à créer)
FIG, TAB = [], []


def nb(x, dec=0):
    return f"{x:,.{dec}f}".replace(",", " ").replace(".", ",")


def V(label, valeur, texte):
    """Enregistre le chiffre écrit (diapositive courante) pour le contrôle automatique."""
    TRACE.append({"diapositive": COURANTE or len(SL) + 1, "indicateur": label, "valeur_exacte": valeur, "texte": texte})
    return texte


def pc(label, x):
    return V(label, x, nb(100 * x, 2) + " %")


def pcv(label, x):
    return V(label, x, nb(x, 2) + " %")


def F(label, x):
    return V(label, x, nb(round(x)) + " FCFA")


def N(label, x, dec=0):
    return V(label, x, nb(x, dec))


def fig(titre):
    FIG.append([len(FIG) + 1, titre, len(SL) + 1])
    return f"Figure {len(FIG)} : {titre}"


def tab(titre):
    TAB.append([len(TAB) + 1, titre, len(SL) + 1])
    return f"Tableau {len(TAB)} : {titre}"


def an(liste, a):
    return next(r for r in liste if r.get("exercice", r.get("annee")) == a)


SIN, BIL = R["sinistralite"], R["bilan"]
s = {a: an(SIN, a) for a in range(2020, 2027)}
b = {a: an(BIL, a) for a in range(2020, 2027)}
GP, GE, EFF, CG, T936 = R["grenier_par"], R["grenier_encours"], R["effectifs_mensuels"], R["cgrae"], R["taux_936"]
EX = C["extraits"]
TC = {t["indicateur"]: t for t in C["tcam"]}
TD_ = {t["serie"][:20]: t for t in C["tendances"]}
MORT = {m["table"]: m for m in C["mortalite"]}
IC = C["ic_grenier"]
SC_ = C["sensibilite_cle"]
INF = {int(k): v for k, v in C["inflation"]["valeurs"].items()}
cats_ex = [str(a) if a < 2026 else "2026 (8 mois)" for a in range(2020, 2027)]
yrs_bil = [str(a) if a < 2026 else "31/08/2026" for a in range(2020, 2027)]
FMT_M = "# ##0,,"           # données exactes en FCFA, affichage en millions

SRC_PREC = "Source : fichiers de précomptes de la Solde, de la CGRAE et du FPPN reçus le 15/10/2026 ; traitements Adenka"
SRC_COMPTA = "Source : balances des comptes du FPPN (Sage 100), exercices 2020 à 2025 et du 01/01 au 31/08/2026 (provisoire) ; traitements Adenka"
SRC_GRENIER = "Source : exports du logiciel de prêts du Grenier du Policier (balances âgées, production) reçus le 15/10/2026 ; traitements Adenka"
SRC_CIMA = "Source : tables de mortalité CIMA H, CIMA F, TD CIMA et TV CIMA (fichiers fournis le 18/10/2026) ; table ronde CIMA FANAF 2011 ; calculs Adenka"
SRC_BM = f"Source : Banque mondiale, FP.CPI.TOTL.ZG, mise à jour du {C['inflation']['mise_a_jour'][8:10]}/{C['inflation']['mise_a_jour'][5:7]}/{C['inflation']['mise_a_jour'][:4]} ; " + SRC_COMPTA[9:]


def contenu(entete, notes, **k):
    SL.append({"type": "contenu", "entete": entete, "notes": notes, **k})


def separateur(num, titre, sous_titre, notes=""):
    SL.append({"type": "separateur", "num": num, "titre": titre, "sous_titre": sous_titre, "notes": notes})


def gab(num, entete, remplissage, notes, source=None):
    SL.append({"type": "gab", "gab": num, "entete": entete, "remplir": remplissage, "notes": notes, "source": source})


def retenir(theme, points, notes):
    contenu({"titre": f"À retenir : {theme}", "sous_titre": points[0][0], "synthese": points[0][1]}, notes,
            cartes=points[1:] if len(points) > 1 else points)


# ===========================================================================
# CONTENU
# ===========================================================================
SL.append({"type": "couverture", "notes": "Rapport d'étape au 15/10/2026, version 2. Données arrêtées à l'envoi du FPPN reçu le 15/10/2026, complétées le 18/10/2026 (grands livres PDF, inflation Banque mondiale, tables CIMA)."})
SL.append({"type": "sommaire", "notes": "Le rapport suit l'ordre : résumé exécutif, compréhension du FPPN, démarche, données, analyses par thème, tables de mortalité, méthodes à venir, outils, analyses en attente, annexes."})

# ---------------------------------------------------------------- RÉSUMÉ EXÉCUTIF
separateur("R", "Résumé exécutif", "Constats, chiffres clés, risques et actions",
           "Cette partie peut être présentée seule : elle reprend les messages essentiels du rapport, avec les chiffres et les renvois aux parties détaillées.")
contenu({"titre": "Résumé exécutif", "sous_titre": "Le FPPN dégage des excédents mais reste fragile et peu documenté",
         "synthese": f"Fonds propres négatifs, épargne couverte à {pc('Couverture épargne 31/08/2026', b[2026]['couverture_epargne'])} par des actifs dédiés, provisions figées depuis 2020"},
        "Questions probables. Pourquoi « fragile » malgré des excédents ? Parce que les excédents annuels ne suffisent pas encore à effacer les pertes passées (fonds propres négatifs) et que l'épargne due aux mutualistes n'est pas couverte par des actifs disponibles. Pourquoi « peu documenté » ? Aucune base de prestations, pas de date de naissance des mutualistes, pas de barème officiel transmis. Chaque ligne renvoie à une partie détaillée du rapport.",
        tableau={"titre": tab("Principaux constats au 15/10/2026"), "entetes": ["Domaine", "Constat", "Chiffre clé"],
                 "largeurs": [4.6, 16.4, 10.2], "lignes": [
                     ["Données", "Les précomptes mensuels ne sont disponibles qu'à partir de 2023 ou 2024 selon le poste. Aucune base de prestations n'a été reçue.", N("Lignes de précomptes contrôlées", 4454416) + " lignes de précomptes contrôlées"],
                     ["Cotisations", "Le précompte total est de 5 % du salaire de base, réparti 3,77 % maladie et 1,23 % décès au lieu de 3,5 % et 1,5 %.", pcv("Taux total observé", 5.0) + " = " + pcv("Taux maladie observé", 100 * T936["mediane"]) + " + 1,23 %"],
                     ["Maladie", "Les prestations maladie comptabilisées augmentent bien plus vite que les cotisations maladie.", "S/C " + pc("S/C maladie 2021", s[2021]["sc_maladie"]) + " (2021), " + pc("S/C maladie 2025", s[2025]["sc_maladie"]) + " (2025)"],
                     ["Provisions", "Les provisions pour prestations à payer ont le même montant de 2020 à 2026.", F("Provisions techniques", b[2026]["provisions_techniques"])],
                     ["Épargne", "L'épargne des mutualistes (PEP, EFC) dépasse largement les actifs qui lui sont dédiés.", F("Épargne totale 31/08/2026", b[2026]["epargne_totale"])],
                     ["Fonds propres", "Les fonds propres sont négatifs sur toute la période ; les excédents réduisent l'écart.", F("Fonds propres 2025", b[2025]["fonds_propres"]) + " (2025)"],
                     ["Grenier", "Le portefeuille de prêts croît très vite et le FPPN le finance en grande partie.", F("Encours Grenier 31/08/2026", GP["encours"]) + " d'encours"],
                     ["Actifs", "Une part importante de l'actif est faite de créances et d'investissements peu liquides.", "CMP : " + F("CMP 31/08/2026", b[2026]["cmp_investissements"])]]},
        source="Source : analyses Adenka sur les documents reçus le 15/10/2026 (S/C : prestations / cotisations)")

gab(49, {"titre": "Quatre indicateurs clés", "sous_titre": "Quatre ratios résument la situation technique et financière du FPPN",
         "synthese": "Sinistralité maladie, couverture de l'épargne, précompte des prêts et taux de cotisation, mesurés sur les données reçues"},
    {"titres": ["Sinistralité maladie 2025", "Couverture de l'épargne", "Précompte des prêts", "Taux de cotisation"],
     "valeurs": [pc("S/C maladie 2025 (gab)", s[2025]["sc_maladie"]), pc("Couverture épargne (gab)", b[2026]["couverture_epargne"]),
                 pc("Précomptes / échéances août 2026", [e for e in R["grenier_echeances"] if e["mois"] == "août 26"][0]["precompte_9x4"] / [e for e in R["grenier_echeances"] if e["mois"] == "août 26"][0]["echeances_balance"]),
                 pcv("Taux total (gab)", 5.0)],
     "commentaires": ["Prestations maladie comptabilisées rapportées aux cotisations maladie de 2025. Ce ratio était de " + pc("S/C 2021 (gab)", s[2021]["sc_maladie"]) + " en 2021.",
                      "Banques, dépôts à terme et fonds dédiés à l'épargne, rapportés à l'épargne due aux mutualistes au 31/08/2026.",
                      "Précomptes Grenier (poste 9X4) rapportés aux échéances des prêts en cours, août 2026.",
                      "Précompte total sur le salaire de base : 3,77 % maladie et 1,23 % décès, sur " + N("Lignes taux", T936["lignes"]) + " lignes."]},
    "Comment lire : chaque ligne donne un indicateur, sa valeur et sa définition. Les quatre valeurs sont calculées sur les fichiers du FPPN et détaillées dans les parties 5, 7, 8 et 4. Le précompte des prêts mesure la part des échéances effectivement retenue sur salaire ; l'écart peut être payé autrement (caisse, banque), ce que le Grenier doit confirmer.",
    source=SRC_PREC)

contenu({"titre": "Fonctionnement du FPPN", "sous_titre": "Les cotisations et l'épargne financent prestations, prêts et investissements",
         "synthese": f"En 2025 : {F('Cotisations 2025', s[2025]['cot_maladie'] + s[2025]['cot_deces'])} de cotisations, {F('Prestations 2025', s[2025]['prest_maladie'] + s[2025]['prest_deces'])} de prestations"},
        "Ce schéma résume les flux d'argent. Les mutualistes cotisent par précompte sur salaire (Solde) ou sur pension (CGRAE). Le FPPN paie les prestations maladie (cliniques, pharmacies) et décès (familles), gère l'épargne (PEP, EFC), finance les prêts du Grenier et des investissements (CMP, terrains). Les montants sont ceux de la comptabilité 2025 et du bilan au 31/08/2026.",
        schema="fonctionnement", source=SRC_COMPTA)

gab(23, {"titre": "Six constats majeurs", "sous_titre": "Les fragilités portent sur les données, les provisions, l'épargne et les actifs",
         "synthese": "Chaque constat est démontré dans la partie indiquée, à partir des fichiers du FPPN"},
    {"blocs": [("Données incomplètes", "Pas de base de prestations, pas de date de naissance. Les analyses actuarielles centrales sont en attente (parties 3 et 13)."),
               ("Sinistralité en hausse", "S/C maladie de " + pc("S/C 2021 (6c)", s[2021]["sc_maladie"]) + " en 2021 à " + pc("S/C 2025 (6c)", s[2025]["sc_maladie"]) + " en 2025 (partie 5)."),
               ("Provisions figées", "Même montant de 2020 à 2026, alors que les prestations ont été multipliées par " + N("Multiplicateur prestations", (s[2025]["prest_maladie"] + s[2025]["prest_deces"]) / (s[2021]["prest_maladie"] + s[2021]["prest_deces"]), 2) + " (partie 9)."),
               ("Épargne non adossée", pc("Couverture (6c)", b[2026]["couverture_epargne"]) + " seulement de l'épargne est couverte par des actifs dédiés (partie 6)."),
               ("Fonds propres négatifs", "De " + F("FP 2020 (6c)", b[2020]["fonds_propres"]) + " en 2020 à " + F("FP 2025 (6c)", b[2025]["fonds_propres"]) + " en 2025 (partie 9)."),
               ("Actifs peu liquides", "CMP, terrains, créance sur l'Hôpital de la Police : l'essentiel de l'actif n'est pas mobilisable rapidement (partie 9).")]},
    "Les six constats sont classés du plus général (données) au plus financier (actifs). Les chiffres sont exacts et proviennent des balances comptables et des précomptes. Le constat « actifs peu liquides » repose sur la nature des comptes (investissements, terrains, créances) ; leur valeur réelle n'a pas été expertisée.",
    source="Source : analyses Adenka")

contenu({"titre": "Graphiques clés (1/2)", "sous_titre": "La sinistralité maladie monte pendant que l'épargne reste peu couverte",
         "synthese": f"S/C maladie 2025 : {pc('S/C 2025 (gc)', s[2025]['sc_maladie'])} ; couverture de l'épargne au 31/08/2026 : {pc('Couverture (gc)', b[2026]['couverture_epargne'])}"},
        "Graphique de gauche : rapport entre prestations et cotisations maladie comptabilisées, par exercice. Graphique de droite : épargne due aux mutualistes (passif) et actifs qui lui sont dédiés. Les deux graphiques sont détaillés dans les parties 5 et 6.",
        graphiques=[{"titre": fig("Ratio S/C maladie (prestations / cotisations)"), "type": "lignes", "categories": cats_ex,
                     "series": [{"nom": "S/C maladie", "valeurs": [round(100 * s[a]["sc_maladie"], 2) for a in range(2020, 2027)], "couleur": BLEU}],
                     "format": "0.00", "legende": None, "unite": "En %"},
                    {"titre": fig("Épargne due et actifs dédiés"), "type": "barres", "categories": yrs_bil,
                     "series": [{"nom": "Épargne PEP et EFC", "valeurs": [round(b[a]["epargne_totale"]) for a in range(2020, 2027)], "couleur": BLEU},
                                {"nom": "Actifs dédiés", "valeurs": [round(b[a]["actifs_dedies_epargne"]) for a in range(2020, 2027)], "couleur": ORANGE}],
                     "format": FMT_M, "unite": "En millions de FCFA"}],
        source=SRC_COMPTA)

contenu({"titre": "Graphiques clés (2/2)", "sous_titre": "Les fonds propres restent négatifs et l'encours de prêts explose",
         "synthese": f"Fonds propres 2025 : {F('FP 2025 (gc)', b[2025]['fonds_propres'])} ; encours du Grenier au 31/08/2026 : {F('Encours (gc)', GP['encours'])}"},
        "Gauche : fonds propres (résultat compris) à chaque clôture. Droite : encours des prêts du Grenier en fin d'année et au 31/08/2026. Détails dans les parties 9 et 7.",
        graphiques=[{"titre": fig("Fonds propres"), "type": "barres", "categories": yrs_bil,
                     "series": [{"nom": "Fonds propres", "valeurs": [round(b[a]["fonds_propres"]) for a in range(2020, 2027)], "couleur": BLEU}],
                     "format": FMT_M, "legende": None, "unite": "En millions de FCFA"},
                    {"titre": fig("Encours des prêts du Grenier"), "type": "barres",
                     "categories": [g["arrete"] for g in GE if g["arrete"].startswith("31/12") or g["arrete"] == "31/08/2026"],
                     "series": [{"nom": "Encours", "valeurs": [round(g["encours"]) for g in GE if g["arrete"].startswith("31/12") or g["arrete"] == "31/08/2026"], "couleur": ORANGE}],
                     "format": FMT_M, "legende": None, "unite": "En millions de FCFA"}],
        source=SRC_COMPTA + " ; " + SRC_GRENIER[9:])

gab(107, {"titre": "Quatre taux à surveiller", "sous_titre": "Une faible part des engagements est couverte par des ressources disponibles",
          "synthese": "Couverture de l'épargne, sinistralité, trésorerie rapportée à l'épargne et prêts en retard de plus de 30 jours"},
    {"anneaux": [(b[2026]["couverture_epargne"], "Épargne couverte par des actifs dédiés (31/08/2026)", "Couverture anneau"),
                 (s[2025]["sc_maladie"], "Sinistralité maladie (S/C 2025)", "S/C anneau"),
                 (C["retrait_max_couvert"]["par_tresorerie"], "Trésorerie totale / épargne due (31/08/2026)", "Trésorerie/épargne anneau"),
                 (IC["p30"], "Prêts en retard de plus de 30 jours (en nombre, 31/08/2026)", "Prêts retard anneau")]},
    "Chaque anneau donne un pourcentage. La trésorerie totale (toutes banques et caisses) ne représente qu'une petite part de l'épargne due. Les prêts en retard sont mesurés ici en nombre de prêts ; en montant d'encours, la part est plus faible (voir partie 7).",
    source=SRC_COMPTA + " ; " + SRC_GRENIER[9:])

gab(83, {"titre": "Carte des risques", "sous_titre": "Quatre causes identifiées produisent quatre risques pour le FPPN",
         "synthese": "Les causes relèvent de la gestion des données et des choix d'emploi des ressources"},
    {"causes": ["Pas de base de prestations ni de dates de naissance", "Provisions jamais recalculées", "Épargne utilisée pour d'autres emplois", "Écart entre échéances et précomptes"],
     "effets": ["Coût réel des sinistres inconnu", "Engagements mal mesurés au bilan", "Manque de liquidités en cas de retraits", "Retards et pertes sur les prêts"]},
    "Lecture : chaque colonne relie une cause (en haut) à son effet (en bas). Les causes sont démontrées par les données ; les effets sont des risques, c'est-à-dire des conséquences possibles, à confirmer par les analyses en attente.",
    source="Source : analyses Adenka")

contenu({"titre": "Recommandations prioritaires", "sous_titre": "Cinq actions conditionnent la fiabilité des engagements et la solvabilité",
         "synthese": "Les trois premières actions relèvent de la transmission de données, les deux dernières des décisions de gestion"},
        "Ordre de priorité : sans les données (actions 1 à 3), les provisions et les engagements ne peuvent pas être calculés. Les actions 4 et 5 relèvent de la direction du FPPN et du conseil d'administration. Les numéros entre parenthèses renvoient à la liste des documents demandés (PBC).",
        tableau={"titre": tab("Actions recommandées"), "entetes": ["N°", "Action", "Objet", "Priorité"], "largeurs": [1.2, 11.4, 15.0, 3.6],
                 "lignes": [["1", "Transmettre les bases de prestations décès et santé 2020 à 2026 (13, 28, 40)", "Calculer les provisions, la sinistralité par année de survenance et les délais", "Urgente"],
                            ["2", "Transmettre le fichier des mutualistes avec date de naissance et sexe (9)", "Mesurer la mortalité, projeter l'équilibre technique, calculer les engagements", "Urgente"],
                            ["3", "Transmettre les comptes individuels PEP et EFC (16, 17)", "Évaluer la dette d'épargne envers chaque mutualiste", "Urgente"],
                            ["4", "Recalculer les provisions techniques à chaque clôture", "Remplacer des montants figés depuis 2020 par une estimation fondée sur les dossiers", "Haute"],
                            ["5", "Adosser l'épargne à des placements liquides", "Couvrir l'épargne exigible, encadrer le financement du Grenier et du CMP", "Haute"]]},
        source="Source : analyses Adenka ; numéros de la liste des documents demandés (PBC)")

gab(39, {"titre": "Feuille de route proposée", "sous_titre": "Quatre étapes pour passer du constat à l'évaluation des engagements",
         "synthese": "Chaque étape dépend de la précédente ; la première est la transmission des données manquantes"},
    {"etapes": [("Collecter", "Bases de prestations, mutualistes avec date de naissance, comptes d'épargne, barèmes et convention CGRAE."),
                ("Calculer", "Provisions (PSAP, Chain Ladder), mortalité observée comparée à CIMA H, engagements d'épargne."),
                ("Vérifier", "Rapprochement avec la comptabilité, ratios du règlement UEMOA, échanges avec les services du FPPN."),
                ("Projeter", "Équilibre technique sur 5 à 10 ans, sensibilités, recommandations chiffrées.")]},
    "La feuille de route ne fixe pas de dates : elles dépendent du délai de transmission des données par le FPPN. Chaque étape produit une mise à jour de ce rapport.",
    source="Source : Adenka")

# ---------------------------------------------------------------- PARTIE 1 : COMPRENDRE LE FPPN
separateur(1, "Comprendre le FPPN", "Cadre, garanties, circuits et notions de base",
           "Cette partie donne au lecteur les clés pour comprendre la suite, sans connaissance préalable en actuariat.")
contenu({"titre": "Cadre juridique", "sous_titre": "Le FPPN est une mutuelle sociale régie par le droit de l'UEMOA",
         "synthese": "Le règlement n° 07/2009/CM/UEMOA fixe les règles ; l'AIRMS contrôle ; le Code CIMA sert de référence technique"},
        "Une mutuelle sociale n'est pas une compagnie d'assurance : elle est régie par le règlement UEMOA de 2009, pas par le Code CIMA. Le Code CIMA est utilisé ici seulement comme référence technique (tables de mortalité, taux d'actualisation), faute de règles techniques équivalentes dans le règlement. Le FOMUS-CI garantit les mutuelles sociales ivoiriennes.",
        tableau={"titre": tab("Textes et organismes de référence"), "entetes": ["Texte ou organisme", "Rôle pour le FPPN", "Usage dans ce rapport"], "largeurs": [8.6, 12.0, 10.6],
                 "lignes": [["Règlement n° 07/2009/CM/UEMOA (26/06/2009)", "Règles de fonctionnement des mutuelles sociales : réserves, placements, gouvernance", "Cadre de référence ; seuils des ratios à relever dans le texte intégral"],
                            ["AIRMS (Autorité de régulation de la mutualité sociale)", "Contrôle des mutuelles sociales en Côte d'Ivoire", "Indicateurs à comparer (en attente)"],
                            ["Ordonnance n° 2024-1043 (FOMUS-CI)", "Fonds de garantie des mutuelles sociales", "Contexte de la garantie"],
                            ["Code CIMA, article 338", "Tables de mortalité et taux d'intérêt maximal (3,5 %) des assureurs vie", "Référence technique : tables CIMA H et F, taux de 3,5 %"],
                            ["Plan comptable des mutuelles sociales (PCMS)", "Plan de comptes utilisé par le FPPN depuis 2021", "Lecture des balances comptables"]]},
        source="Source : textes cités (classeur des sources externes validées)")

pos = {(c["poste"], c["sens"]): c for c in R["couverture"]}
contenu({"titre": "Garanties et produits", "sous_titre": "Huit postes de précompte portent l'ensemble des produits du FPPN",
         "synthese": "Maladie et décès sont les garanties de base ; épargne, prêts, flotte et MEDOC sont des produits complémentaires"},
        "Chaque produit a un code de poste sur la fiche de paie (Solde) : 936 maladie, 8Z6 décès, 870 EFC, 998 PEP, 9X4 Grenier, 9W3 flotte, 8Z7 MEDOC. Les retraités sont précomptés par la CGRAE sur leur pension. Les montants moyens sont observés dans les fichiers de précomptes, pas tirés d'un barème (non reçu).",
        tableau={"titre": tab("Produits du FPPN et montants observés"), "entetes": ["Poste", "Produit", "Nature", "Montant moyen observé par mois"], "largeurs": [2.6, 7.6, 12.0, 9.0],
                 "lignes": [["936", "Maladie", "Garantie santé des actifs, 3,77 % du salaire de base", F("Moyen 936 août 2026", [e for e in EFF if e["annee"] == 2026 and e["mois"] == 8][0]["montant"] / [e for e in EFF if e["annee"] == 2026 and e["mois"] == 8][0]["cotisants"]) + " (août 2026)"],
                            ["8Z6", "Décès", "Capital décès, frais funéraires, allocations ; 1,23 % du salaire de base", "voir partie 4"],
                            ["870", "Entraide fin de carrière (EFC)", "Épargne versée en fin de carrière", F("Moyen EFC", an(R["efc_versements"], 2025)["versements"] / sum(1 for _ in range(1)) / an(R["efc_versements"], 2025)["epargnants"] / an(R["efc_versements"], 2025)["mois"]) + " par épargnant (2025)"],
                            ["998", "Plan d'épargne (PEP)", "Épargne volontaire", "voir partie 6"],
                            ["9X4", "Grenier du Policier", "Remboursement de prêts aux policiers", "voir partie 7"],
                            ["9W3", "Flotte automobile (GNA)", "Assurance des véhicules des adhérents", "environ 10 000 FCFA"],
                            ["8Z7", "MEDOC", "Produit santé complémentaire (nature à préciser)", "5 000 FCFA"],
                            ["CGRAE", "Retraités", "Garantie santé des retraités, précompte sur pension", F("Moyen CGRAE", sum(c["moyen"] for c in CG) / len(CG))]]},
        source=SRC_PREC)

contenu({"titre": "Circuit des précomptes", "sous_titre": "Les cotisations sont retenues à la source, puis rapprochées par le FPPN",
         "synthese": "Chaque mois, le FPPN envoie un fichier « aller » ; la Solde ou la CGRAE renvoie un fichier « retour » des sommes retenues"},
        "Le fichier aller dit ce que le FPPN demande de retenir ; le fichier retour dit ce qui a été réellement retenu. L'écart entre les deux révèle des rejets (agent sorti, montant erroné, plafond atteint). Ce rapprochement est la base du contrôle des cotisations (partie 3).",
        schema="precomptes", source="Source : fichiers aller et retour reçus le 15/10/2026 ; Adenka")

contenu({"titre": "Flux financiers 2025", "sous_titre": "Les ressources de l'année couvrent les prestations, les frais et dégagent un excédent",
         "synthese": f"Produits {F('Produits 2025', s[2025]['produits'])}, charges {F('Charges 2025', s[2025]['charges'])}, résultat {F('Résultat 2025', s[2025]['resultat'])}"},
        "Le résultat de l'exercice est la différence entre produits et charges (y compris les éléments exceptionnels, dits hors activités ordinaires). En 2025, il reste positif, mais il diminue depuis 2021 (partie 5).",
        graphique={"titre": fig("Produits, charges et résultat 2025"), "type": "barres",
                   "categories": ["Cotisations maladie", "Cotisations décès", "Prestations maladie", "Prestations décès", "Frais de gestion", "Résultat"],
                   "series": [{"nom": "2025", "valeurs": [round(s[2025]["cot_maladie"]), round(s[2025]["cot_deces"]), round(s[2025]["prest_maladie"]), round(s[2025]["prest_deces"]), round(s[2025]["frais_gestion"]), round(s[2025]["resultat"])], "couleur": BLEU}],
                   "format": FMT_M, "legende": None, "unite": "En millions de FCFA"},
        texte=[f"**Les cotisations sont la ressource principale.** {F('Cot mal 2025', s[2025]['cot_maladie'])} en maladie et {F('Cot dec 2025', s[2025]['cot_deces'])} en décès.",
               f"**Les prestations maladie absorbent la plus grande part.** {F('Prest mal 2025', s[2025]['prest_maladie'])}, contre {F('Prest dec 2025', s[2025]['prest_deces'])} en décès.",
               f"**Les frais de gestion pèsent {pc('Taux frais 2025', s[2025]['taux_frais'])} des cotisations.** Ils couvrent le personnel, les indemnités des organes, le fonctionnement et les amortissements."],
        source=SRC_COMPTA)

contenu({"titre": "Notions de base", "sous_titre": "Sept notions suffisent pour lire l'ensemble du rapport",
         "synthese": "Chaque notion est définie ici une fois, puis utilisée avec le même sens dans tout le rapport"},
        "Ces définitions sont volontairement simples. Les formules exactes figurent dans les diapositives « méthode » de chaque partie et dans l'annexe des méthodes.",
        tableau={"titre": tab("Notions utilisées dans le rapport"), "entetes": ["Notion", "Définition simple", "Exemple dans le rapport"], "largeurs": [5.2, 15.4, 10.6],
                 "lignes": [["Cotisation", "Somme versée par le mutualiste pour être couvert", "5 % du salaire de base"],
                            ["Prestation", "Somme payée par le FPPN quand le risque survient (soins, décès)", "Factures des cliniques, capital décès"],
                            ["Sinistralité (S/C)", "Prestations divisées par cotisations : part des cotisations consommée par les prestations", "S/C maladie 2025"],
                            ["Provision technique", "Somme mise de côté au bilan pour payer les prestations déjà nées mais pas encore payées", "Provision pour prestations à payer"],
                            ["Engagement", "Ce que le FPPN devra payer plus tard à ses mutualistes", "Épargne PEP et EFC à restituer"],
                            ["Adossement", "Correspondance entre les dettes et les actifs disponibles pour les payer", "Épargne et placements dédiés"],
                            ["Fonds propres", "Ce qui reste quand on retire toutes les dettes de tous les actifs", "Négatifs de 2020 à 2026"]]},
        source="Source : Adenka")

# ---------------------------------------------------------------- PARTIE 2 : DÉMARCHE
separateur(2, "Contexte et démarche", "Mission, périmètre, chaîne de traitement")
contenu({"titre": "Mission et périmètre", "sous_titre": "Le volet actuariel couvre les engagements, les provisions et les placements",
         "synthese": "Période d'analyse du 1er janvier 2020 au 31 août 2026 ; données arrêtées à l'envoi du FPPN reçu le 15/10/2026"},
        "Les sept axes viennent de l'offre technique d'Adenka. La colonne « état » dit honnêtement ce qui est fait et ce qui attend des données.",
        tableau={"titre": tab("Périmètre de la mission actuarielle et statistique"), "entetes": ["Axe", "Question traitée", "État au 15/10/2026"], "largeurs": [6.4, 15.2, 9.6],
                 "lignes": [["Contrôle des données de base", "Les données sont-elles complètes et cohérentes avec la comptabilité ?", "Réalisé sur les données reçues"],
                            ["Engagements envers les adhérents", "Combien le FPPN doit-il aux mutualistes et ayants droit ?", "Partiel : épargne seulement"],
                            ["Suffisance des provisions", "Les provisions couvrent-elles les prestations dues ?", "Constat sur les montants, calcul en attente"],
                            ["Adossement actif / passif", "Les actifs sont-ils disponibles quand les dettes arrivent à échéance ?", "Réalisé sur les balances"],
                            ["Adéquation des placements", "Les placements sont-ils sûrs, liquides, rentables ?", "Réalisé sur les balances"],
                            ["Indicateurs AIRMS", "Les ratios des mutuelles sociales sont-ils respectés ?", "En attente des seuils du règlement"],
                            ["Équilibre technique", "Les cotisations futures couvrent-elles les prestations futures ?", "En attente des âges et des dossiers"]]},
        source="Source : offre technique Adenka et liste des documents demandés (PBC)")
contenu({"titre": "Chaîne de traitement", "sous_titre": "Chaque chiffre est tracé du fichier du FPPN jusqu'au rapport",
         "synthese": "Cinq étapes, chacune avec son contrôle : inventaire, extraction, base, analyses, rapport"},
        "Ce schéma est modifiable dans PowerPoint et sera complété au fil des entretiens. Chaque flèche correspond à un script Python conservé ; le contrôle indiqué sous chaque étape est automatique.",
        schema="chaine", source="Source : scripts Python Adenka (dossier scripts de la mission)")

# ---------------------------------------------------------------- PARTIE 3 : DONNÉES
separateur(3, "Données et qualité", "Couverture, contrôles, anomalies détectées")
contenu({"titre": "Méthode : contrôle des données", "sous_titre": "Trois contrôles garantissent que la base reproduit fidèlement les fichiers",
         "synthese": "Exhaustivité, cohérence interne et rapprochement entre sources indépendantes"},
        "Exhaustivité : rien n'est perdu entre le fichier reçu et la base (empreintes SHA-256, nombre de lignes et de cellules). Cohérence : les totaux imprimés par les logiciels sont retrouvés. Rapprochement : deux sources qui décrivent le même flux (aller et retour, précomptes et comptabilité) doivent concorder. Limite : ces contrôles prouvent la fidélité aux fichiers, pas l'exactitude des fichiers eux-mêmes.",
        methode=[("Exhaustivité", "Empreinte SHA-256 de chaque fichier ; lignes et cellules comptées avant et après conversion.", "218 fichiers copiés, 491 feuilles converties"),
                 ("Cohérence interne", "Totaux imprimés retrouvés (Grenier) ; débit = crédit (comptabilité).", "37 balances âgées, 7 exercices"),
                 ("Rapprochement", "Taux = montant retour / montant aller, par poste et par mois commun.", "6 postes comparés")],
        source="Source : Adenka")
cov = {(c["poste"], c["sens"]): c["mois_recus"] for c in R["couverture"]}
lib_flux = [("Comptabilité (balances)", 80), ("PEP 998 retour", cov[("998", "retour")]), ("EFC 870 retour", cov[("870", "retour")]),
            ("Grenier 9X4 retour", cov[("9X4", "retour")]), ("Flotte 9W3 retour", cov[("9W3", "retour")]), ("Maladie 936", cov[("936", "retour")]),
            ("Retraités CGRAE retour", cov[("CGRAE", "retour")]), ("Décès 8Z6", cov[("8Z6", "retour")]), ("MEDOC 8Z7 retour", cov[("8Z7", "retour")]),
            ("Prestations décès et santé", 0)]
contenu({"titre": "Couverture de la période", "sous_titre": "Les données mensuelles ne couvrent pas les années 2020 à 2022",
         "synthese": "Sur 80 mois, seuls la comptabilité et le PEP sont complets ; aucune donnée de prestations n'a été reçue"},
        "Un mois est compté comme reçu quand au moins un fichier lisible le couvre. Les balances comptables couvrent l'année entière en totaux annuels ; le détail des écritures de janvier à août 2020 manque.",
        graphique={"titre": fig("Mois reçus par flux de données, sur 80 mois"), "type": "barres_h", "categories": [x[0] for x in lib_flux][::-1],
                   "series": [{"nom": "Mois reçus", "valeurs": [x[1] for x in lib_flux][::-1], "couleur": BLEU}], "format": "0",
                   "unite": "En nombre de mois (janvier 2020 à août 2026)"},
        texte=["**La comptabilité couvre toute la période en totaux annuels.** Seul le détail des écritures de janvier à août 2020 manque.",
               "**Les précomptes mensuels commencent en 2023 ou 2024.** Maladie et décès en janvier 2024, retraités en février 2024, EFC en janvier 2023.",
               "**Le PEP est complet grâce aux fichiers texte de la Solde**, qui donnent le détail par mutualiste et par mois de 2020 à 2022.",
               "**Aucune base de prestations n'a été transmise.**"],
        source=SRC_PREC)
rt = R["lignes_total_ecartees"]
lt = EX["ligne_total"]
contenu({"titre": "Calcul pas à pas : lignes de total", "sous_titre": "Une ligne de total cachée aurait doublé les cotisations maladie du mois",
         "synthese": f"Fichier de février 2025 : la dernière ligne porte {F('Total février 2025', float(lt['lignes'][-1][5]))}, soit la somme des {N('Lignes fichier février 2025', lt['nb_lignes_fichier'] - 1)} lignes"},
        "Lecture : (1) la colonne matricule de la dernière ligne contient un nombre de lignes, pas un matricule ; (2) le montant est la somme de toutes les lignes au-dessus ; (3) compter cette ligne aurait doublé le total du mois. Règle appliquée : une ligne dont le montant égale la somme des autres (à 1 % près) est écartée. Au total, " + N("Lignes total écartées (notes)", rt["n"]) + " lignes de ce type ont été écartées. Noms et matricules masqués.",
        pas_a_pas={"entetes": lt["entete"], "lignes": lt["lignes"], "largeurs": [2.0, 2.6, 1.6, 3.4, 3.6, 3.6, 3.4, 2.4],
                   "rouge": [(len(lt["lignes"]) - 1, 3), (len(lt["lignes"]) - 1, 5)],
                   "explications": ["Le « matricule » est en fait le nombre de lignes du fichier.",
                                    f"Le montant {F('Total février (pas)', float(lt['lignes'][-1][5]))} est la somme de toutes les lignes.",
                                    f"Règle : ligne écartée. {N('Lignes total écartées', rt['n'])} lignes de ce type écartées au total ({F('Montant total écarté', rt['montant'])} de double compte évité)."],
                   "fichier": lt["fichier"]},
        source="Source : " + lt["fichier"] + " (identifiant d'extraction " + lt["id"] + ") ; noms et matricules masqués")
md = EX["montants_doubles"]
contenu({"titre": "Calcul pas à pas : montants doublés", "sous_titre": "Des montants saisis deux fois à la suite gonflaient les précomptes décès",
         "synthese": f"Avril 2024 : {N('Montants corrigés', md['total_corriges'])} montants de plus d'un million de FCFA sont des montants écrits deux fois"},
        "Lecture : (1) le montant source répète deux fois la même suite de chiffres ; (2) un précompte décès de plusieurs milliards de FCFA par agent est impossible (médiane du poste : environ 2 300 FCFA) ; (3) le montant retenu est la première moitié. Correction appliquée seulement au poste décès, au-delà d'un million de FCFA, et signalée au FPPN.",
        pas_a_pas={"entetes": ["Matricule", "Nom", "Montant dans le fichier", "Montant corrigé"], "lignes": md["lignes"], "largeurs": [4.0, 4.0, 8.0, 6.0],
                   "rouge": [(0, 2), (0, 3)],
                   "explications": ["Le montant source répète deux fois « 47924 ».", "Un précompte décès de plus d'un million par agent est impossible.",
                                    "Montant retenu : la première moitié ; ligne signalée au FPPN."], "fichier": md["fichier"]},
        source="Source : " + md["fichier"] + " ; noms et matricules masqués")
ar = R["aller_retour"]
nomp = {"870": "EFC 870", "998": "PEP 998", "9X4": "Grenier 9X4", "9W3": "Flotte 9W3", "CGRAE": "Retraités CGRAE", "8Z7": "MEDOC 8Z7"}
contenu({"titre": "Rapprochement aller et retour", "sous_titre": "La Solde précompte en moyenne 97 à 100 % des montants demandés",
         "synthese": "Sur les mois communs, l'écart moyen est faible, mais certains mois s'écartent fortement"},
        "Taux de précompte = montant retour / montant aller sur les mois où les deux fichiers existent. Un taux supérieur à 100 % signifie que le fichier aller était incomplet (cas du PEP en juillet 2024).",
        graphique={"titre": fig("Montant précompté rapporté au montant demandé, mois communs"), "type": "barres", "categories": [nomp[a["poste"]] for a in ar],
                   "series": [{"nom": "Taux moyen", "valeurs": [round(100 * a["taux_recouvrement"], 2) for a in ar], "couleur": BLEU},
                              {"nom": "Mois le plus bas", "valeurs": [round(100 * a["taux_min_mensuel"], 2) for a in ar], "couleur": ORANGE}],
                   "format": "0.00", "unite": "En % du montant demandé"},
        texte=[f"**Le taux moyen va de {pc('Taux min moyen', min(a['taux_recouvrement'] for a in ar))} à {pc('Taux max moyen', max(a['taux_recouvrement'] for a in ar))}.** Les montants demandés sont presque entièrement retenus.",
               f"**Certains mois s'écartent nettement.** Mois le plus bas : {pc('Min Grenier', [a for a in ar if a['poste'] == '9X4'][0]['taux_min_mensuel'])} pour le Grenier, {pc('Min PEP', [a for a in ar if a['poste'] == '998'][0]['taux_min_mensuel'])} pour le PEP.",
               "**Ces écarts traduisent des rejets ou des fichiers incomplets** et doivent être expliqués mois par mois par le FPPN."],
        source=SRC_PREC)
apm, apx = EX["aller_retour_pep"]["min"], EX["aller_retour_pep"]["max"]
contenu({"titre": "Calcul pas à pas : écarts du PEP", "sous_titre": "Les deux mois extrêmes du PEP s'expliquent par le contenu des fichiers",
         "synthese": f"Janvier 2025 : {pc('Taux PEP janv 2025', apm['taux'])} ; juillet 2024 : {pc('Taux PEP juil 2024', apx['taux'])}"},
        "Janvier 2025 : le fichier retour compte moins de lignes que le fichier aller, des précomptes n'ont pas été faits. Juillet 2024 : le fichier retour compte plus de lignes que le fichier aller, le fichier aller n'incluait pas les nouveaux épargnants (le fichier aller d'août 2024 en compte nettement plus).",
        pas_a_pas={"entetes": ["Mois", "Fichier aller", "Lignes aller", "Montant aller", "Lignes retour", "Montant retour", "Taux"],
                   "lignes": [["janv. 2025", apm["fichier_a"], N("Lignes aller janv25", apm["lignes_a"]), F("Montant aller janv25", apm["montant_a"]), N("Lignes retour janv25", apm["lignes_r"]), F("Montant retour janv25", apm["montant_r"]), pc("Taux janv25 (t)", apm["taux"])],
                              ["juil. 2024", apx["fichier_a"], N("Lignes aller juil24", apx["lignes_a"]), F("Montant aller juil24", apx["montant_a"]), N("Lignes retour juil24", apx["lignes_r"]), F("Montant retour juil24", apx["montant_r"]), pc("Taux juil24 (t)", apx["taux"])]],
                   "largeurs": [2.4, 6.6, 2.6, 5.4, 2.6, 5.4, 3.0], "rouge": [(0, 4), (1, 4)],
                   "explications": ["Janvier 2025 : " + N("Écart lignes janv25", apm["lignes_a"] - apm["lignes_r"]) + " épargnants demandés mais non précomptés.",
                                    "Juillet 2024 : " + N("Écart lignes juil24", apx["lignes_r"] - apx["lignes_a"]) + " lignes de plus au retour qu'à l'aller.",
                                    "Conclusion : le taux mensuel dépend de la qualité des deux fichiers ; le FPPN doit tenir un état des rejets."],
                   "fichier": "FPPN-SOLDE\\PEP\\FICHIER ALLER et FICHIER RETOUR"},
        source=SRC_PREC)
contenu({"titre": "Fichiers vides ou endommagés", "sous_titre": "Onze fichiers sont vides, illisibles ou mal classés et doivent être renvoyés",
         "synthese": "Les fichiers comptables endommagés ont tous un double lisible, sauf le grand livre des comptes 2022"},
        "Un fichier vide (0 octet) ne contient aucune donnée ; un fichier endommagé a une partie de ses données détruite (souvent lors d'une copie ou d'un transfert). Aucun de ces défauts ne vient du traitement d'Adenka : les empreintes prouvent que les fichiers sont identiques à ceux reçus.",
        tableau={"titre": tab("Fichiers à renvoyer ou à corriger"), "entetes": ["Fichier", "Défaut", "Conséquence"], "largeurs": [13.0, 8.6, 9.6],
                 "lignes": [["Décès 8Z6 et maladie 936, décembre 2024", "Vides (0 octet), identiques", "Aucune donnée pour décembre 2024"],
                            ["MEDOC 8Z7, retours de février et avril 2026", "« aucune ligne sélectionnée »", "Précomptes MEDOC inconnus ces deux mois"],
                            ["Balances des tiers 2021 et 2025 (Excel)", "Endommagées", "PDF lisibles utilisés"],
                            ["Grands livres des comptes 2021, 2024 (Excel)", "Endommagés", "PDF lisibles utilisés"],
                            ["Grand livre des comptes 2022 (Excel et PDF)", "Endommagés tous les deux", "417 pages sur 657 récupérées ; balance 2022 intacte"],
                            ["Grands livres des tiers 2023 et 2025 (Excel)", "Endommagés", "PDF lisibles utilisés"],
                            ["Fichiers aller dans les dossiers retour (CGRAE juin 2024, GNA janvier 2026)", "Mal classés", "Retours de ces mois manquants"]]},
        source="Source : journal de contrôle de l'extraction ; note interne des anomalies")
retenir("données", [("Les données reçues sont fidèlement reprises, mais incomplètes", "Base contrôlée à 100 % ; manquent les prestations, les âges et les années 2020 à 2022 en mensuel"),
                    ("Fidélité", "Empreintes identiques, lignes et cellules identiques, totaux des logiciels retrouvés."),
                    ("Anomalies corrigées", f"{N('Total écartés (ret)', rt['n'])} lignes de total écartées, {N('Corrigés (ret)', md['total_corriges'])} montants doublés corrigés."),
                    ("Manques", "Bases de prestations, dates de naissance, précomptes 2020 à 2022, fichiers vides de décembre 2024.")],
        "Message pour le client : la qualité de l'analyse dépend désormais de la transmission des données manquantes, pas du traitement.")

# ---------------------------------------------------------------- PARTIE 4 : EFFECTIFS ET COTISATIONS
separateur(4, "Effectifs et cotisations", "Cotisants, retraités, taux, cotisations en termes réels")
contenu({"titre": "Méthode : effectifs et taux", "sous_titre": "Effectifs, flux et taux sont mesurés agent par agent, chaque mois",
         "synthese": "Un cotisant est un matricule précompté au poste maladie 936 dans le mois"},
        "Effectif du mois m = nombre de matricules distincts du fichier maladie 936 du mois m. Entrée = matricule présent en m et absent en m-1 ; sortie = l'inverse (mois consécutifs seulement). Taux de cotisation = précompte du mois / salaire de base du mois, pour chaque agent des mois détaillés. Limite : les motifs de sortie (retraite, décès, radiation) ne sont pas connus.",
        methode=[("Effectif", "E(m) = nombre de matricules distincts au poste 936 le mois m", "31 mois disponibles"),
                 ("Flux", "Entrées(m) = |M(m) \\ M(m-1)| ; Sorties(m) = |M(m-1) \\ M(m)|", "Mois consécutifs"),
                 ("Taux de cotisation", "t = précompte / salaire de base, par agent", f"{N('Lignes taux (méthode)', T936['lignes'])} lignes")],
        source="Source : Adenka")
e0 = [e for e in EFF if e["annee"] == 2024 and e["mois"] == 1][0]
e1 = [e for e in EFF if e["annee"] == 2026 and e["mois"] == 8][0]
tdc = C["tendances"][0]
contenu({"titre": "Effectifs cotisants", "sous_titre": "Les cotisants actifs augmentent d'un tiers en trente et un mois",
         "synthese": f"{N('Cotisants janv 2024', e0['cotisants'])} cotisants en janvier 2024, {N('Cotisants août 2026', e1['cotisants'])} en août 2026"},
        f"La droite de tendance estimée par moindres carrés donne une hausse de {N('Pente cotisants', tdc['pente'], 2)} cotisants par mois (intervalle de confiance à 95 % : {N('IC bas pente', tdc['ic95_bas'], 2)} à {N('IC haut pente', tdc['ic95_haut'], 2)}). La probabilité d'observer une telle pente si la vraie pente était nulle est inférieure à 0,01 % (p-valeur {tdc['p_valeur']:.2e}). La hausse se fait par paliers en juillet, ce qui rend la droite seulement approximative (R² = {nb(tdc['r2'], 2)}).",
        graphique={"titre": fig("Cotisants actifs précomptés (poste 936), par mois"), "type": "lignes", "categories": [e["libelle"] for e in EFF],
                   "series": [{"nom": "Cotisants", "valeurs": [e["cotisants"] for e in EFF], "couleur": BLEU}], "format": "# ##0", "etiquettes": False,
                   "axe_min": 20000, "legende": None, "unite": "En nombre de mutualistes"},
        texte=[f"**Hausse de {pc('Hausse cotisants', e1['cotisants'] / e0['cotisants'] - 1)} en 31 mois.** Taux de croissance annuel moyen : {pc('TCAM cotisants', TC['Cotisants actifs (maladie 936)']['tcam'])}.",
               f"**La hausse se fait par paliers en juillet.** {N('Entrées juil 2024', EX['entrees_juillet']['2024']['entrees'])} entrées en juillet 2024 et {N('Entrées juil 2025', EX['entrees_juillet']['2025']['entrees'])} en juillet 2025 : ce sont les promotions de recrues.",
               f"**La tendance est statistiquement significative** : {N('Pente (texte)', tdc['pente'], 2)} cotisants de plus par mois en moyenne.",
               "**Le fichier ne donne ni âge ni sexe**, ce qui empêche de projeter la mortalité et les départs à la retraite."],
        source=SRC_PREC + " (décembre 2024 : fichier vide)")
fl = R["flux_mensuels"]
contenu({"titre": "Entrées et sorties", "sous_titre": "Les entrées massives de juillet s'opposent à des sorties faibles et régulières",
         "synthese": f"Entrées : {N('Entrées 2024', R['flux_annuels'][0]['entrees'])} en 2024, {N('Entrées 2025', R['flux_annuels'][1]['entrees'])} en 2025 ; sorties : {N('Sorties 2024', R['flux_annuels'][0]['sorties'])} et {N('Sorties 2025', R['flux_annuels'][1]['sorties'])}"},
        "Les mois sans mois précédent disponible (janvier 2024, janvier 2025 après le fichier vide de décembre 2024) n'ont pas de flux calculé. Les sorties de novembre 2024 et janvier 2026 sont plus élevées : leur motif doit être demandé au FPPN (départs à la retraite de fin d'année probables, non vérifiés).",
        graphique={"titre": fig("Entrées et sorties de cotisants, par mois"), "type": "barres", "categories": [f["mois"] for f in fl],
                   "series": [{"nom": "Entrées", "valeurs": [f["entrees"] for f in fl], "couleur": BLEU}, {"nom": "Sorties", "valeurs": [f["sorties"] for f in fl], "couleur": ORANGE}],
                   "format": "# ##0", "etiquettes": False, "unite": "En nombre de mutualistes"},
        texte=["**Les entrées sont concentrées en juillet**, mois d'intégration des recrues.",
               "**Les sorties restent faibles**, sauf en novembre 2024 et janvier 2026.",
               "**Le motif de sortie est indispensable** pour distinguer retraites, décès et radiations, et donc pour calculer des taux de mortalité."],
        source=SRC_PREC)
anc = R["anciennete"]
contenu({"titre": "Structure des effectifs", "sous_titre": "Une population jeune en service, recrutée en majorité depuis moins de dix ans",
         "synthese": f"En {anc['mois']}, ancienneté médiane de {N('Ancienneté médiane', anc['mediane'], 2)} ans, salaire de base médian de {F('Salaire base médian', anc['salaire_base_median'])}"},
        "L'ancienneté est calculée à partir de la date de recrutement du fichier détaillé de janvier 2025. Elle ne donne pas l'âge : deux agents de même ancienneté peuvent avoir des âges différents. Pour cette raison, aucune hypothèse d'âge n'a été faite.",
        graphique={"titre": fig(f"Agents par ancienneté de service, {anc['mois']}"), "type": "barres", "categories": anc["categories"],
                   "series": [{"nom": "Agents", "valeurs": anc["valeurs"], "couleur": BLEU}], "format": "# ##0", "unite": "En nombre d'agents"},
        texte=[f"**{pc('Part moins de 10 ans', sum(anc['valeurs'][:2]) / sum(anc['valeurs']))} des agents ont moins de dix ans de service.**",
               f"**Une seconde vague a 15 à 25 ans de service** : {pc('Part 15-25 ans', (anc['valeurs'][3] + anc['valeurs'][4]) / sum(anc['valeurs']))} des agents.",
               "**Cette vague partira à la retraite dans les années à venir** et deviendra une population de retraités CGRAE, plus consommatrice de soins."],
        source=SRC_PREC + " (fichiers maladie 936 détaillés)")
tx = EX["taux_cotisation"]
contenu({"titre": "Calcul pas à pas : taux de cotisation", "sous_titre": "Pour chaque agent, maladie et décès font 5,00 % du salaire de base",
         "synthese": f"La même répartition, 3,77 % et 1,23 %, se retrouve sur {N('Lignes taux (pas)', T936['lignes'])} lignes des mois détaillés"},
        "Calcul pour le premier agent : 11 354 / 301 159 = 3,77 % (maladie) ; 3 704 / 301 159 = 1,23 % (décès) ; total 5,00 %. Le barème annoncé en entretien est 3,5 % et 1,5 %. Le total est donc bon mais la répartition diffère. Matricules masqués.",
        pas_a_pas={"entetes": ["Matricule", "Grade", "Salaire de base", "Maladie 936", "Décès 8Z6", "Maladie %", "Décès %", "Total %"],
                   "lignes": [r[:5] + [r[5].replace(".", ",") + " %", r[6].replace(".", ",") + " %", r[7].replace(".", ",") + " %"] for r in tx["lignes"]],
                   "largeurs": [3.2, 2.2, 4.2, 3.8, 3.6, 3.4, 3.2, 3.2], "rouge": [(0, 5), (0, 6), (0, 7)],
                   "explications": ["11 354 / 301 159 = 3,77 % au lieu de 3,5 %.", "3 704 / 301 159 = 1,23 % au lieu de 1,5 %.", "Le total de 5,00 % est conforme ; la clé maladie / décès ne l'est pas."],
                   "fichier": "FPPN-SOLDE\\EFFECTIF POLICE (MALADIE)\\2024\\042024_effectif_police_936_AVRIL 2024.xlsx et FPPN_ALLOC_DECES_8Z6_AVRIL_2024.xlsx"},
        source=SRC_PREC + " ; matricules masqués")
tot25 = s[2025]["cot_maladie"] + s[2025]["cot_deces"]
contenu({"titre": "Taux de cotisation", "sous_titre": "Trois clés de répartition coexistent pour le même taux de 5 %",
         "synthese": f"Barème annoncé 3,5 / 1,5 ; Solde 3,77 / 1,23 ; comptabilité 2025 {nb(5 * s[2025]['cot_maladie'] / tot25, 2)} / {nb(5 * s[2025]['cot_deces'] / tot25, 2)}"},
        "La clé comptable est obtenue en rapportant les cotisations maladie comptabilisées au total des cotisations, puis en multipliant par 5 %. Elle est proche du barème annoncé, pas de la clé de la Solde.",
        graphique={"titre": fig("Répartition du taux de 5 % selon la source"), "type": "barres_empilees",
                   "categories": ["Barème annoncé", "Précompte Solde", "Comptabilité 2025"],
                   "series": [{"nom": "Maladie", "valeurs": [3.5, round(100 * T936["mediane"], 2), round(5 * s[2025]["cot_maladie"] / tot25, 2)], "couleur": BLEU},
                              {"nom": "Décès", "valeurs": [1.5, round(5 - 100 * T936["mediane"], 2), round(5 * s[2025]["cot_deces"] / tot25, 2)], "couleur": ORANGE}],
                   "format": "0.00", "unite": "En % du salaire de base"},
        texte=["**Le total de 5 % est respecté partout.**",
               "**La Solde retient 3,77 % au poste maladie et 1,23 % au poste décès.**",
               f"**La comptabilité répartit à {pc('Clé comptable maladie', SC_['cle_comptable_maladie'])} pour la maladie**, proche des 70 % du barème annoncé.",
               "**Le barème officiel n'a pas été reçu** et doit trancher."],
        source=SRC_PREC + " ; " + SRC_COMPTA[9:])
contenu({"titre": "Sensibilité : clé de répartition", "sous_titre": "La clé retenue change la sinistralité de chaque garantie de plusieurs points",
         "synthese": f"S/C maladie 2025 : {pc('S/C clé compta', SC_['sc_maladie_compta'])} avec la clé comptable, {pc('S/C clé Solde', SC_['sc_maladie_cle_solde'])} avec la clé de la Solde"},
        "Calcul : S/C maladie = prestations maladie / (cotisations totales × part maladie). Avec la clé de la Solde (75,40 %), la part des cotisations affectée à la maladie augmente, donc le ratio maladie baisse et le ratio décès monte. Les prestations ne changent pas : seule l'affectation des cotisations change.",
        graphique={"titre": fig("Ratio S/C 2025 selon la clé de répartition"), "type": "barres", "categories": ["Clé comptable", "Clé de la Solde (3,77 / 1,23)", "Barème annoncé (3,5 / 1,5)"],
                   "series": [{"nom": "S/C maladie", "valeurs": [round(100 * SC_[k], 2) for k in ("sc_maladie_compta", "sc_maladie_cle_solde", "sc_maladie_cle_bareme")], "couleur": BLEU},
                              {"nom": "S/C décès", "valeurs": [round(100 * SC_[k], 2) for k in ("sc_deces_compta", "sc_deces_cle_solde", "sc_deces_cle_bareme")], "couleur": ORANGE}],
                   "format": "0.00", "unite": "En %"},
        texte=[f"**Avec la clé de la Solde, le S/C maladie 2025 passe à {pc('S/C solde (t)', SC_['sc_maladie_cle_solde'])}** et le S/C décès à {pc('S/C décès solde', SC_['sc_deces_cle_solde'])}.",
               f"**Avec le barème annoncé, le S/C maladie vaut {pc('S/C barème', SC_['sc_maladie_cle_bareme'])}.**",
               "**Le résultat par garantie dépend donc d'une décision de gestion**, qui doit être écrite dans le barème et appliquée de la même façon par la Solde et la comptabilité."],
        source=SRC_COMPTA)
ip = C["indice_prix_2020_100"]
contenu({"titre": "Cotisations en termes réels", "sous_titre": "Hors inflation, les cotisations progressent de 23 % entre 2020 et 2025",
         "synthese": f"Hausse nominale {pc('Croissance nominale cot', C['cot_croissance_nominale_2020_2025'])}, hausse des prix {pc('Croissance prix', C['prix_croissance_2020_2025'])}, hausse réelle {pc('Croissance réelle cot', C['cot_croissance_reelle_2020_2025'])}"},
        "Méthode : indice des prix base 100 en 2020, construit en cumulant l'inflation annuelle de la Banque mondiale (2021 : 4,16 % ; 2022 : 5,23 % ; 2023 : 4,37 % ; 2024 : 3,45 % ; 2025 : 0,13 %). Cotisations réelles = cotisations nominales × 100 / indice. Elles mesurent le pouvoir d'achat des cotisations en francs de 2020.",
        graphique={"titre": fig("Cotisations nominales et réelles (francs de 2020)"), "type": "lignes", "categories": [str(a) for a in range(2020, 2026)],
                   "series": [{"nom": "Nominales", "valeurs": [round(r["cotisations_nominales"]) for r in C["reel"]], "couleur": BLEU},
                              {"nom": "Réelles (francs de 2020)", "valeurs": [round(r["cotisations_reelles_2020"]) for r in C["reel"]], "couleur": ORANGE}],
                   "format": FMT_M, "unite": "En millions de FCFA"},
        texte=[f"**L'inflation explique une partie de la hausse.** Les prix ont augmenté de {pc('Prix (t)', C['prix_croissance_2020_2025'])} entre 2020 et 2025.",
               f"**La hausse réelle reste nette : {pc('Réelle (t)', C['cot_croissance_reelle_2020_2025'])}.** Elle vient surtout de la hausse des effectifs.",
               f"**L'inflation de 2025 est très faible ({pcv('Inflation 2025', INF[2025])})**, après {pcv('Inflation 2022', INF[2022])} en 2022."],
        source=SRC_BM)
cot25 = an(R["cotisations"], 2025)
contenu({"titre": "Cotisations comptabilisées", "sous_titre": "Les cotisations comptabilisées progressent de 45,76 % entre 2020 et 2025",
         "synthese": f"{F('Cot 2020', s[2020]['cot_maladie'] + s[2020]['cot_deces'])} en 2020, {F('Cot 2025 (cc)', tot25)} en 2025 ; taux de croissance annuel moyen {pc('TCAM cot', TC['Cotisations maladie et décès']['tcam'])}"},
        "Rapprochement 2025 : cotisations maladie comptabilisées comparées aux précomptes maladie (12 mois) et retraités CGRAE (12 mois). Les cotisations à recevoir du Trésor (compte 41111000) mesurent ce que l'État doit encore au FPPN à la clôture.",
        graphique={"titre": fig("Cotisations comptabilisées par garantie"), "type": "barres_empilees", "categories": cats_ex,
                   "series": [{"nom": "Maladie", "valeurs": [round(s[a]["cot_maladie"]) for a in range(2020, 2027)], "couleur": BLEU},
                              {"nom": "Décès", "valeurs": [round(s[a]["cot_deces"]) for a in range(2020, 2027)], "couleur": ORANGE}],
                   "format": FMT_M, "unite": "En millions de FCFA"},
        texte=[f"**En 2025, comptabilité et précomptes maladie diffèrent de {pc('Écart maladie 2025', 1 - s[2025]['cot_maladie'] / (cot25['prec_936'] + cot25['prec_cgrae']))}.** Comptabilité {F('Compta mal 2025', s[2025]['cot_maladie'])} ; précomptes {F('Précomptes mal 2025', cot25['prec_936'] + cot25['prec_cgrae'])}.",
               "**L'écart vient de la clé de répartition** : la comptabilité affecte au décès une part plus grande que la Solde.",
               f"**Fin 2025, l'État devait {F('Créance Trésor 2025', b[2025]['cotisations_a_recevoir_etat'])} de cotisations** ; le compte est soldé au 31/08/2026."],
        source=SRC_COMPTA)
retenir("effectifs et cotisations", [("Les cotisations augmentent avec les effectifs, mais leur répartition doit être clarifiée", "Effectifs en hausse d'un tiers, taux total de 5 % respecté, trois clés de répartition différentes"),
                                      ("Effectifs", f"{N('Cotisants (ret)', e1['cotisants'])} cotisants en août 2026, hausse significative par paliers en juillet."),
                                      ("Taux", "5,00 % au total, mais 3,77 % / 1,23 % à la Solde contre 3,5 % / 1,5 % annoncés."),
                                      ("Données à obtenir", "Barème officiel, dates de naissance, motifs de sortie.")],
        "La question de la clé de répartition est la plus importante de cette partie : elle change le résultat de chaque garantie.")

# ---------------------------------------------------------------- PARTIE 5 : SINISTRALITÉ
separateur(5, "Sinistralité et résultat", "Prestations, frais de gestion, résultat")
contenu({"titre": "Méthode : sinistralité comptable", "sous_titre": "Le ratio S/C rapporte les prestations de l'année aux cotisations de l'année",
         "synthese": "S/C = Σ prestations comptabilisées / Σ cotisations comptabilisées, par garantie et par exercice"},
        "Le ratio comptable mélange les soins de l'année et les factures d'années antérieures payées ou comptabilisées dans l'année. Le vrai ratio (par année de survenance des soins) demande la base des prestations avec la date des soins. Les comptes retenus sont listés dans l'annexe des méthodes. En 2020, la prime d'assurance décès (compte 65180000) est ajoutée au décès car la garantie était alors confiée à un assureur.",
        methode=[("Formule", "S/C(garantie, n) = Prestations(garantie, n) / Cotisations(garantie, n)", "Exercices 2020 à 2026"),
                 ("Prestations maladie", "Comptes 601, 6013, 6018, 60800007", "12 comptes en 2025"),
                 ("Prestations décès", "Comptes 6021, 6024 (+ 65180000 en 2020)", "Obsèques, frais funéraires")],
        source="Source : Adenka")
sc25 = EX["sc_2025"]
contenu({"titre": "Calcul pas à pas : S/C maladie 2025", "sous_titre": "Douze comptes de prestations divisés par un compte de cotisations",
         "synthese": f"{F('Prest mal 2025 (pas)', s[2025]['prest_maladie'])} / {F('Cot mal 2025 (pas)', s[2025]['cot_maladie'])} = {pc('S/C 2025 (pas)', s[2025]['sc_maladie'])}"},
        "Le compte 60120000 « Charges de prestations maladie » porte à lui seul plus de la moitié du total : il regroupe des prestations sans détail de nature, ce qui empêche l'analyse par type de soins. Montants exacts lus dans la balance 2025.",
        pas_a_pas={"entetes": ["Compte", "Intitulé", "Montant 2025"], "lignes": [[c[0], c[1], nb(float(c[2])) + " FCFA"] for c in sc25["prestations"]] + [["Total", "Prestations maladie", F("Total prest (pas)", s[2025]["prest_maladie"])], [sc25["cotisations"][0][0], sc25["cotisations"][0][1], nb(-float(sc25["cotisations"][0][2])) + " FCFA"]],
                   "largeurs": [3.4, 11.0, 6.0], "rouge": [(7, 2), (12, 2), (13, 2)], "petit": True,
                   "explications": ["Compte global : plus de la moitié des prestations, sans détail.", "Somme des 12 comptes de prestations maladie.", "Cotisations maladie acquises ; S/C = total / cotisations."],
                   "fichier": "AUDIT ADENKA\\DAFC\\2025\\BALANCE FPPN 2025.xlsx"},
        source="Source : balance des comptes 2025 (Sage 100) ; traitements Adenka")
contenu({"titre": "Sinistralité comptable", "sous_titre": "Les prestations maladie absorbent presque toutes les cotisations maladie",
         "synthese": f"S/C maladie : {pc('S/C 2021 (s)', s[2021]['sc_maladie'])} en 2021, {pc('S/C 2025 (s)', s[2025]['sc_maladie'])} en 2025, {pc('S/C 2026 (s)', s[2026]['sc_maladie'])} sur huit mois de 2026"},
        f"Test de tendance (moindres carrés sur 2020 à 2025) : le S/C maladie progresse de {N('Pente S/C', C['tendances'][1]['pente'], 2)} points par an (IC 95 % : {N('IC bas S/C', C['tendances'][1]['ic95_bas'], 2)} à {N('IC haut S/C', C['tendances'][1]['ic95_haut'], 2)} ; p-valeur {nb(C['tendances'][1]['p_valeur'], 4)}). Avec six points seulement, la tendance est significative au seuil de 5 % mais son rythme est imprécis.",
        graphique={"titre": fig("Prestations comptabilisées / cotisations (S/C)"), "type": "lignes", "categories": cats_ex,
                   "series": [{"nom": "Maladie", "valeurs": [round(100 * s[a]["sc_maladie"], 2) for a in range(2020, 2027)], "couleur": BLEU},
                              {"nom": "Décès (prime d'assurance incluse en 2020)", "valeurs": [round(100 * s[a]["sc_deces_avec_prime"], 2) for a in range(2020, 2027)], "couleur": ORANGE}],
                   "format": "0.00", "unite": "En % des cotisations de la garantie"},
        texte=[f"**Le ratio maladie a été multiplié par {N('Mult S/C', s[2025]['sc_maladie'] / s[2021]['sc_maladie'], 2)} entre 2021 et 2025.**",
               "**Les ratios très bas de 2020 à 2022 sont suspects** : des factures de ces années ont probablement été comptabilisées plus tard. À vérifier sur la base des prestations.",
               f"**En décès, le ratio oscille entre {pc('Min S/C décès', min(s[a]['sc_deces'] for a in range(2022, 2027)))} et {pc('Max S/C décès', max(s[a]['sc_deces'] for a in range(2022, 2027)))} depuis 2022.**",
               f"**En 2020, la prime d'assurance décès ({F('Prime 2020', s[2020]['prime_assurance_deces'])}) représentait {pc('S/C décès 2020 prime', s[2020]['sc_deces_avec_prime'])} des cotisations décès.**"],
        source=SRC_COMPTA)
tlp = C["tendances"][2]
contenu({"titre": "Croissance des prestations", "sous_titre": "Les prestations maladie croissent beaucoup plus vite que les cotisations",
         "synthese": f"Taux de croissance annuel moyen 2021 à 2025 : prestations maladie {pc('TCAM prest mal', TC['Prestations maladie']['tcam'])}, cotisations {pc('TCAM cot (cr)', TC['Cotisations maladie et décès']['tcam'])}"},
        f"Le taux de croissance annuel moyen (TCAM) est le taux constant qui mène de la valeur de départ à la valeur d'arrivée : TCAM = (V_fin / V_début)^(1/n) - 1. Une régression du logarithme des prestations sur le temps donne une croissance annuelle moyenne de {pc('Croissance log', tlp['croissance_annuelle'])} (IC 95 % : {pc('IC bas log', tlp['croissance_ic95_bas'])} à {pc('IC haut log', tlp['croissance_ic95_haut'])}), estimation plus robuste aux valeurs extrêmes que le TCAM.",
        tableau={"titre": tab("Taux de croissance annuels moyens"), "entetes": ["Indicateur", "Période", "Valeur de départ", "Valeur d'arrivée", "Multiplié par", "TCAM"], "largeurs": [7.6, 4.6, 5.6, 5.6, 3.6, 4.2],
                 "lignes": [[t["indicateur"], f"{t['debut']} à {t['fin']}",
                             (F(f"V0 {t['indicateur']}", t["v0"]) if t["v0"] > 100000 else N(f"V0 {t['indicateur']}", t["v0"])),
                             (F(f"V1 {t['indicateur']}", t["v1"]) if t["v1"] > 100000 else N(f"V1 {t['indicateur']}", t["v1"])),
                             N(f"Mult {t['indicateur']}", t["multiplicateur"], 2), pc(f"TCAM {t['indicateur']}", t["tcam"])] for t in C["tcam"]]},
        source=SRC_COMPTA + " ; " + SRC_PREC[9:] + " ; " + SRC_GRENIER[9:])
nat = R["prestations_maladie_nature"]
contenu({"titre": "Prestations maladie par nature", "sous_titre": "Plus de la moitié des prestations maladie est enregistrée sans détail de nature",
         "synthese": f"En 2025, le compte global porte {F('Compte global 2025', nat[5]['2025'])} sur {F('Prest mal 2025 (n)', s[2025]['prest_maladie'])}"},
        "Lecture : chaque couleur est une nature de soins. La catégorie « autres prestations maladie » correspond aux comptes 6012x, dont le compte global 60120000. Imagerie, analyses et consultations sont les natures détaillées qui progressent le plus.",
        graphique={"titre": fig("Prestations maladie comptabilisées par nature"), "type": "barres_empilees", "categories": cats_ex,
                   "series": [{"nom": n["nature"], "valeurs": [round(n[str(e)]) for e in range(2020, 2027)]} for n in nat],
                   "format": FMT_M, "etiquettes": False, "legende": "droite", "unite": "En millions de FCFA"},
        texte=[f"**Les prestations maladie passent de {F('Prest 2021', s[2021]['prest_maladie'])} en 2021 à {F('Prest 2025 (nat)', s[2025]['prest_maladie'])} en 2025.**",
               f"**Imagerie, analyses et consultations représentent {pc('Part img+ana+cons', (nat[4]['2025'] + nat[3]['2025'] + nat[1]['2025']) / s[2025]['prest_maladie'])} des prestations de 2025.**",
               "**Un compte global empêche le suivi par type de soins et par prestataire.**"],
        source=SRC_COMPTA)
contenu({"titre": "Frais de gestion et résultat", "sous_titre": "Les excédents se réduisent et deviennent un déficit en 2026",
         "synthese": f"Résultat net : {F('Résultat 2020', s[2020]['resultat'])} en 2020, {F('Résultat 2025 (fr)', s[2025]['resultat'])} en 2025, {F('Résultat 2026', s[2026]['resultat'])} sur huit mois de 2026"},
        "Frais de gestion = comptes 61 à 66, hors prime d'assurance décès. Résultat = produits - charges + résultat hors activités ordinaires (classe 8). Les huit mois de 2026 n'intègrent pas les écritures de fin d'année (amortissements, provisions).",
        graphique={"titre": fig("Frais de gestion et résultat net"), "type": "barres", "categories": cats_ex,
                   "series": [{"nom": "Frais de gestion", "valeurs": [round(s[a]["frais_gestion"]) for a in range(2020, 2027)], "couleur": "BBBCBC"},
                              {"nom": "Résultat net", "valeurs": [round(s[a]["resultat"]) for a in range(2020, 2027)], "couleur": BLEU}],
                   "format": FMT_M, "unite": "En millions de FCFA"},
        texte=[f"**Les frais pèsent entre {pc('Min frais', min(s[a]['taux_frais'] for a in range(2021, 2026)))} et {pc('Max frais', max(s[a]['taux_frais'] for a in range(2021, 2026)))} des cotisations depuis 2021.**",
               f"**Le résultat baisse depuis 2021** : de {F('Résultat 2021', s[2021]['resultat'])} à {F('Résultat 2025 (t)', s[2025]['resultat'])} en 2025.",
               "**Les huit premiers mois de 2026 sont déficitaires**, avant écritures de fin d'année."],
        source=SRC_COMPTA)
retenir("sinistralité", [("La sinistralité maladie approche le niveau des cotisations", "Ratio S/C maladie multiplié par plus de douze depuis 2021, résultat en baisse, déficit sur 2026"),
                         ("Tendance", "Hausse significative du S/C maladie, environ 12,91 points par an de 2020 à 2025."),
                         ("Limite", "Ratio comptable : le coût réel par année de soins demande la base des prestations."),
                         ("Action", "Détailler le compte global de prestations par nature et par prestataire.")],
        "La hausse rapide du ratio maladie est le principal risque technique : à ce rythme, les cotisations maladie ne couvriront plus les prestations.")

# ---------------------------------------------------------------- PARTIE 6 : ÉPARGNE
separateur(6, "Épargne des mutualistes", "Dette d'épargne, actifs dédiés, sensibilité aux retraits")
contenu({"titre": "Méthode : adossement de l'épargne", "sous_titre": "La couverture compare l'épargne due aux actifs qui lui sont réservés",
         "synthese": "Couverture = (banques, dépôts à terme et fonds dédiés) / (comptes PEP et EFC)"},
        "L'épargne PEP et EFC appartient aux mutualistes : c'est une dette du FPPN. Un bon adossement suppose des actifs liquides et réservés d'un montant au moins égal à la part exigible. Limite : les conditions de retrait ne sont pas connues ; la part réellement exigible à court terme ne peut donc pas être calculée.",
        methode=[("Passif d'épargne", "Comptes 41195100 (PEP) et 41195200 (EFC)", "Soldes créditeurs"),
                 ("Actifs dédiés", "Comptes 5211020x, 5221020x (EFC), 5211030x, 5221030x, 526120 (PEP), 5330 (FCP)", "Soldes débiteurs"),
                 ("Couverture", "Actifs dédiés / passif d'épargne", "À chaque clôture")],
        source="Source : Adenka")
ep = EX["epargne_2026"]
contenu({"titre": "Calcul pas à pas : couverture de l'épargne", "sous_titre": "Deux comptes de dette face à quatre comptes d'actifs dédiés",
         "synthese": f"{F('Actifs dédiés (pas)', b[2026]['actifs_dedies_epargne'])} / {F('Épargne (pas)', b[2026]['epargne_totale'])} = {pc('Couverture (pas)', b[2026]['couverture_epargne'])} au 31/08/2026"},
        "Montants lus dans la balance au 31/08/2026 (provisoire). Les soldes créditeurs du passif sont affichés en valeur positive. Le fonds commun de placement est compté comme actif dédié bien que sa destination ne soit pas précisée : la couverture réelle est donc au plus de ce niveau.",
        pas_a_pas={"entetes": ["Compte", "Intitulé", "Solde au 31/08/2026"],
                   "lignes": [[c[0], c[1], nb(-float(c[2])) + " FCFA"] for c in ep["passif"]] + [["", "Total épargne due", F("Épargne totale (pas)", b[2026]["epargne_totale"])]] + [[c[0], c[1], nb(float(c[2])) + " FCFA"] for c in ep["actifs"]] + [["", "Total actifs dédiés", F("Total actifs (pas)", b[2026]["actifs_dedies_epargne"])]],
                   "largeurs": [3.4, 11.4, 6.6], "rouge": [(2, 2), (7, 2)],
                   "explications": ["Dette envers les épargnants PEP et EFC.", "Actifs réservés : banques, dépôt à terme, fonds commun de placement.", f"Couverture = {pc('Couverture (expl)', b[2026]['couverture_epargne'])}."],
                   "fichier": "AUDIT ADENKA\\DAFC\\2026\\BALANCE FPPN DU 01-01-2026 AU 31-08-2026.xlsx"},
        source="Source : balance au 31/08/2026 (impression provisoire)")
contenu({"titre": "Adossement de l'épargne", "sous_titre": "L'épargne des mutualistes n'est couverte qu'à 16,28 % par des actifs dédiés",
         "synthese": f"Couverture : entre {pc('Min couverture', min(b[a]['couverture_epargne'] for a in range(2020, 2027)))} et {pc('Max couverture', max(b[a]['couverture_epargne'] for a in range(2020, 2027)))} selon les années"},
        "L'épargne non couverte finance d'autres emplois : investissements au CMP, terrains, créance sur l'Hôpital de la Police, prêts du Grenier. Ces emplois ne sont pas mobilisables rapidement.",
        graphique={"titre": fig("Épargne au passif et actifs dédiés"), "type": "barres", "categories": yrs_bil,
                   "series": [{"nom": "Épargne PEP et EFC", "valeurs": [round(b[a]["epargne_totale"]) for a in range(2020, 2027)], "couleur": BLEU},
                              {"nom": "Actifs dédiés", "valeurs": [round(b[a]["actifs_dedies_epargne"]) for a in range(2020, 2027)], "couleur": ORANGE}],
                   "format": FMT_M, "unite": "En millions de FCFA"},
        texte=[f"**La dette d'épargne augmente de {pc('Hausse épargne', b[2026]['epargne_totale'] / b[2020]['epargne_totale'] - 1)} depuis 2020.**",
               f"**L'EFC triple** : {F('EFC 2020', b[2020]['epargne_efc'])} en 2020, {F('EFC 2026', b[2026]['epargne_efc'])} au 31/08/2026.",
               "**Les actifs dédiés restent faibles**, et l'écart finance des emplois peu liquides."],
        source=SRC_COMPTA)
pv = [p for p in R["pep_versements"] if p["annee"] >= 2020]
ev = {e["annee"]: e for e in R["efc_versements"]}
contenu({"titre": "Versements d'épargne", "sous_titre": "Les versements au PEP ont été multipliés par quatre depuis 2022",
         "synthese": f"PEP 2025 : {F('PEP 2025', pv[5]['versements'])} versés par {N('Épargnants PEP 2025', pv[5]['epargnants'])} épargnants"},
        "Les versements sont les précomptes retour de la Solde. La différence entre versements de l'année et variation du compte d'épargne mesure les sorties de l'année (échéances, retraits, transferts), sans pouvoir les détailler faute de comptes individuels.",
        graphique={"titre": fig("Versements précomptés par produit d'épargne"), "type": "barres", "categories": [str(p["annee"]) if p["annee"] < 2026 else "2026 (8 mois)" for p in pv],
                   "series": [{"nom": "PEP 998", "valeurs": [round(p["versements"]) for p in pv], "couleur": BLEU},
                              {"nom": "EFC 870", "valeurs": [round(ev[p["annee"]]["versements"]) if p["annee"] in ev else None for p in pv], "couleur": ORANGE}],
                   "format": FMT_M, "unite": "En millions de FCFA (EFC à partir de 2023)"},
        texte=[f"**Épargnants PEP : {N('PEP 2020', pv[0]['epargnants'])} en 2020, {N('PEP 2026', pv[-1]['epargnants'])} en 2026.**",
               f"**En 2025, le PEP reçoit {F('PEP 2025 (t)', pv[5]['versements'])} mais le compte n'augmente que de {F('Var PEP 2025', b[2025]['epargne_pep'] - b[2024]['epargne_pep'])}** : la différence correspond à des sorties.",
               "**Sans comptes individuels, la dette réelle envers chaque épargnant n'est pas vérifiable.**"],
        source=SRC_PREC)
rr = C["sensibilite_retraits"]
contenu({"titre": "Sensibilité : retraits d'épargne", "sous_titre": "Un retrait de 20 % de l'épargne dépasserait la trésorerie totale du FPPN",
         "synthese": f"Les actifs dédiés couvrent un retrait de {pc('Retrait couvert actifs', C['retrait_max_couvert']['par_actifs_dedies'])} de l'épargne ; la trésorerie totale, {pc('Retrait couvert tréso', C['retrait_max_couvert']['par_tresorerie'])}"},
        "Scénarios théoriques, pas des prévisions : on suppose qu'une part de l'épargne est demandée en une fois et on la compare aux actifs dédiés puis à toute la trésorerie (banques et caisses). Les conditions réelles de retrait (délais, pénalités, échéances) ne sont pas connues.",
        graphique={"titre": fig("Besoin de liquidités selon le taux de retrait"), "type": "barres", "categories": [nb(100 * r["taux_retrait"], 2) + " %" for r in rr],
                   "series": [{"nom": "Montant à rembourser", "valeurs": [round(r["besoin"]) for r in rr], "couleur": BLEU},
                              {"nom": "Actifs dédiés", "valeurs": [round(r["actifs_dedies"]) for r in rr], "couleur": ORANGE},
                              {"nom": "Trésorerie totale", "valeurs": [round(r["tresorerie_totale"]) for r in rr], "couleur": "BBBCBC"}],
                   "format": FMT_M, "unite": "En millions de FCFA, au 31/08/2026"},
        texte=[f"**Un retrait de 10 % ({F('Retrait 10%', rr[1]['besoin'])}) est couvert par les actifs dédiés.**",
               f"**Au-delà de {pc('Seuil tréso', C['retrait_max_couvert']['par_tresorerie'])}, toute la trésorerie ne suffit plus.**",
               "**Le FPPN devrait alors vendre ou mobiliser des actifs peu liquides.**"],
        source=SRC_COMPTA)
contenu({"titre": "Poste 998 : autres références", "sous_titre": "Le poste de l'épargne PEP porte aussi d'autres références à expliquer",
         "synthese": "PTI, RPEP, EP et TI apparaissent à côté des références PEP sur la période 2020 à 2026"},
        "La référence est le code qui accompagne chaque précompte (exemple : PEP0522 = PEP souscrit en mai 2022). Les références PTI ont des montants mensuels élevés et peu de mutualistes : elles pourraient correspondre à des remboursements de prêts, ce que le FPPN doit confirmer.",
        tableau={"titre": tab("Références du poste 998, 2020 à août 2026"), "entetes": ["Référence", "Lignes", "Montant total", "Mutualistes"], "largeurs": [6.0, 6.0, 10.0, 6.0],
                 "lignes": [[r["produit"], N(f"Lignes {r['produit']}", r["lignes"]), F(f"Montant {r['produit']}", r["montant"]), N(f"Mut {r['produit']}", r["mutualistes"])] for r in R["poste_998_references"]]},
        source=SRC_PREC)
retenir("épargne", [("L'épargne due aux mutualistes est très peu couverte par des actifs disponibles", "Couverture de 16,28 % au 31/08/2026 ; un retrait de 21,34 % épuiserait toute la trésorerie"),
                    ("Dette", F("Épargne (ret)", b[2026]["epargne_totale"]) + " au 31/08/2026."),
                    ("Risque", "Liquidité : impossibilité de rembourser rapidement en cas de retraits importants."),
                    ("Données à obtenir", "Règlement des produits, comptes individuels, conditions de retrait.")],
        "L'adossement de l'épargne est le principal risque financier relevé.")

# ---------------------------------------------------------------- PARTIE 7 : GRENIER
separateur(7, "Prêts du Grenier", "Encours, production, retards, précomptes")
contenu({"titre": "Méthode : risque de crédit", "sous_titre": "Le risque des prêts se mesure par la part de l'encours en retard",
         "synthese": "Encours en retard de plus de 30 jours / encours total, à chaque date de balance âgée"},
        "L'encours est le capital restant dû. Le retard (en jours) est celui imprimé par le logiciel de prêts. On mesure la part en nombre de prêts et en montant d'encours, avec un intervalle de confiance de Wilson pour la proportion de prêts en retard. Limite : les retards dépendent de la date des précomptes de la Solde.",
        methode=[("Part en retard (montant)", "Σ encours des prêts en retard > 30 j / Σ encours", "37 balances âgées"),
                 ("Part en retard (nombre)", "nombre de prêts en retard > 30 j / nombre de prêts, IC de Wilson à 95 %", "31/08/2026"),
                 ("Précompte des échéances", "précomptes 9X4 du mois / échéances des prêts en cours", "Mois de 2026")],
        source="Source : Adenka")
ge_fa = [g for g in GE if g["arrete"].startswith("31/12") or g["arrete"] == "31/08/2026"]
contenu({"titre": "Encours du Grenier", "sous_titre": "L'encours de prêts a été multiplié par 132 depuis fin 2021",
         "synthese": f"Au 31/08/2026 : {N('Prêts', GP['prets'])} prêts, {F('Encours (gr)', GP['encours'])}, dont {F('Financement FPPN', b[2026]['grenier_financement'])} financés par le FPPN"},
        f"Taux de croissance annuel moyen de l'encours de fin 2022 à fin 2025 : {pc('TCAM encours', TC['Encours des prêts du Grenier']['tcam'])}. La créance du FPPN sur le Grenier (compte 47111102) mesure le financement apporté.",
        graphique={"titre": fig("Encours des prêts et financement du FPPN"), "type": "barres", "categories": [g["arrete"] for g in ge_fa],
                   "series": [{"nom": "Encours", "valeurs": [round(g["encours"]) for g in ge_fa], "couleur": BLEU},
                              {"nom": "Financement FPPN", "valeurs": [round(g["financement_fppn_47111102"]) if g["financement_fppn_47111102"] is not None else None for g in ge_fa], "couleur": ORANGE}],
                   "format": FMT_M, "unite": "En millions de FCFA"},
        texte=[f"**Encours : {F('Encours 2021', ge_fa[0]['encours'])} fin 2021, {F('Encours 2026 (t)', GP['encours'])} au 31/08/2026**, soit {N('Mult encours', GP['encours'] / ge_fa[0]['encours'], 2)} fois plus.",
               f"**Le FPPN finance {pc('Part financement', b[2026]['grenier_financement'] / GP['encours'])} de l'encours.**",
               f"**Prêts courts et chers** : durée médiane {N('Durée médiane', GP['duree_mediane'])} mois, taux médian {pcv('Taux médian', GP['taux_median'])}."],
        source=SRC_GRENIER + " ; " + SRC_COMPTA[9:])
gpr = R["grenier_production"]
contenu({"titre": "Production de prêts", "sous_titre": "Les crédits scolaires et habitat portent la croissance de la production",
         "synthese": f"Montants débloqués : {F('Prod 2022', gpr['total'][1])} en 2022, {F('Prod 2025', gpr['total'][4])} en 2025 ({N('Nb prêts 2025', gpr['nombre'][4])} prêts)"},
        "Production = montants débloqués sur la période (fichiers annuels de 2021 à 2025, mensuels par sous-portefeuille en 2026). La politique de crédit de 2021 est au format Word ancien et doit être renvoyée lisible pour vérifier taux, durées et plafonds.",
        graphique={"titre": fig("Montants débloqués par type de crédit"), "type": "barres_empilees", "categories": [str(a) if a < 2026 else "2026 (8 mois)" for a in gpr["annees"]],
                   "series": [{"nom": t.capitalize(), "valeurs": [round(v) for v in gpr["valeurs"][t]]} for t in gpr["types"]],
                   "format": FMT_M, "etiquettes": False, "legende": "droite", "unite": "En millions de FCFA"},
        texte=[f"**La production a été multipliée par {N('Mult prod', gpr['total'][4] / gpr['total'][1], 2)} entre 2022 et 2025.**",
               f"**Scolaire et habitat font {pc('Part scol+hab', (gpr['valeurs']['CRÉDIT SCOLAIRE'][4] + gpr['valeurs']['CREDIT HABITAT'][4]) / gpr['total'][4])} des montants de 2025.**",
               "**Le crédit habitat, aux montants unitaires plus importants, allonge le risque.**"],
        source=SRC_GRENIER)
ph = [p for p in R["grenier_par_historique"] if p["arrete"] != "30/09/2026"]
contenu({"titre": "Retards de remboursement", "sous_titre": "Les retards ont culminé début 2026 avant de revenir à un niveau faible",
         "synthese": f"Encours en retard de plus de 30 jours : {pc('PAR30 janv 2026', [p for p in ph if p['arrete'] == '31/01/2026'][0]['par30'])} fin janvier 2026, {pc('PAR30 août 2026', GP['par30'])} au 31/08/2026"},
        f"Au 31/08/2026, {N('Prêts retard 30', IC['retard30'])} prêts sur {N('Prêts (ic)', IC['prets'])} ont plus de 30 jours de retard, soit {pc('Part prêts retard', IC['p30'])} des prêts (IC de Wilson à 95 % : {pc('IC bas retard', IC['p30_bas'])} à {pc('IC haut retard', IC['p30_haut'])}). En montant, la part est plus faible ({pc('PAR30 (n)', GP['par30'])}) : les prêts en retard sont de petits encours.",
        graphique={"titre": fig("Part de l'encours en retard de plus de 30 et 90 jours"), "type": "lignes", "categories": [p["arrete"] for p in ph],
                   "series": [{"nom": "Plus de 30 jours", "valeurs": [round(100 * p["par30"], 2) for p in ph], "couleur": BLEU},
                              {"nom": "Plus de 90 jours", "valeurs": [round(100 * p["par90"], 2) for p in ph], "couleur": ORANGE}],
                   "format": "0.00", "etiquettes": False, "unite": "En % de l'encours"},
        texte=["**Fin janvier 2026, la plupart de l'encours avait entre 31 et 90 jours de retard.**",
               "**Ces retards se sont résorbés en quelques mois**, ce qui fait penser à des échéances non précomptées en fin d'année 2025.",
               f"**{N('Prêts > 180 j', IC['retard180'])} prêts ont plus de 180 jours de retard** et doivent être provisionnés."],
        source=SRC_GRENIER)
pr_ = EX["prets_retard"]
contenu({"titre": "Calcul pas à pas : prêts en retard", "sous_titre": "Les retards les plus anciens dépassent trois ans et révèlent des mécanos anormaux",
         "synthese": "Les quatre prêts les plus en retard au 31/08/2026 ont entre 1 105 et 1 308 jours de retard"},
        "Lecture : (1) le retard est imprimé en jours par le logiciel ; (2) trois de ces prêts ont un « mécano » de 2 ou 3 chiffres (142, 87, 166), alors qu'un mécano en compte 6 : ce sont des anomalies de saisie ; (3) ces prêts anciens ont encore un encours et doivent être provisionnés ou passés en perte. Noms masqués.",
        pas_a_pas={"entetes": ["Mécano", "Nom", "Montant débloqué", "Encours", "Retard (jours)"], "lignes": [[r[0], r[1], nb(float(r[2])) + " FCFA", nb(float(r[3])) + " FCFA", r[4]] for r in pr_["lignes"]],
                   "largeurs": [3.4, 4.0, 6.0, 5.6, 4.6], "rouge": [(0, 0), (0, 4), (1, 0)],
                   "explications": ["Mécanos de 2 ou 3 chiffres : saisie anormale.", "Plus de trois ans de retard.", "Encours encore inscrit : à provisionner."],
                   "fichier": "AUDIT ADENKA\\SERVICE GRENIER DU POLICIER\\FICHIER ALLER\\GRENIER AOÛT 2026\\BalanceAgeeDetaille AOUT 2026.xls"},
        source=SRC_GRENIER + " ; noms masqués")
ech = [e for e in R["grenier_echeances"] if e["precompte_9x4"]]
contenu({"titre": "Échéances et précomptes", "sous_titre": "Les précomptes ne couvrent que 76 à 85 % des échéances des prêts en cours",
         "synthese": f"De janvier à août 2026, l'écart mensuel va de {F('Écart min', min(e['echeances_balance'] - e['precompte_9x4'] for e in ech))} à {F('Écart max', max(e['echeances_balance'] - e['precompte_9x4'] for e in ech))}"},
        "Échéances = somme des montants d'échéance des prêts ayant un encours dans la balance âgée du mois. Précomptes = précomptes 9X4 retour du même mois. L'écart peut être payé hors Solde (caisse, banque) ou correspondre à des rejets.",
        graphique={"titre": fig("Échéances des prêts en cours et précomptes 9X4, 2026"), "type": "barres", "categories": [e["mois"] for e in ech],
                   "series": [{"nom": "Échéances", "valeurs": [round(e["echeances_balance"]) for e in ech], "couleur": BLEU},
                              {"nom": "Précomptes 9X4", "valeurs": [round(e["precompte_9x4"]) for e in ech], "couleur": ORANGE}],
                   "format": FMT_M, "unite": "En millions de FCFA"},
        texte=[f"**Part précomptée : de {pc('Min précompte/éch', min(e['precompte_9x4'] / e['echeances_balance'] for e in ech))} à {pc('Max précompte/éch', max(e['precompte_9x4'] / e['echeances_balance'] for e in ech))}.**",
               "**L'export du logiciel perd la lettre-clé du mécano** (321100C devient 321100), ce qui complique le rapprochement prêt par prêt.",
               "**Un rapprochement avec les encaissements bancaires du Grenier est nécessaire.**"],
        source=SRC_GRENIER + " ; " + SRC_PREC[9:])
sp = C["sensibilite_pertes_grenier"]
contenu({"titre": "Sensibilité : pertes sur prêts", "sous_titre": "Même une perte totale sur les prêts très en retard reste limitée en montant",
         "synthese": f"Encours en retard de plus de 180 jours : {F('Encours 180', sp['encours_180j'])} ; de plus de 90 jours : {F('Encours 90', sp['encours_90j'])}"},
        "Trois taux de perte théoriques (25 %, 50 %, 100 %) appliqués à l'encours en retard. Ce ne sont pas des estimations de perte, mais des ordres de grandeur pour fixer une provision. La politique de provisionnement doit être fixée par le FPPN.",
        tableau={"titre": tab("Provision selon le taux de perte"), "entetes": ["Taux de perte", "Sur l'encours > 180 jours", "Sur l'encours > 90 jours"], "largeurs": [7.0, 10.0, 10.0],
                 "lignes": [[pc(f"Perte {p['perte']}", p["perte"]), F(f"Prov180 {p['perte']}", p["provision_180j"]), F(f"Prov90 {p['perte']}", p["provision_90j"])] for p in sp["scenarios"]]},
        source=SRC_GRENIER)
retenir("Grenier", [("Le Grenier croît très vite grâce au financement du FPPN", "Encours multiplié par 132, financé à 87,67 %, retards maîtrisés mais précomptes incomplets"),
                    ("Croissance", f"{F('Encours (ret)', GP['encours'])} d'encours au 31/08/2026."),
                    ("Risque", "Précomptes inférieurs aux échéances ; 126 prêts en retard de plus de 180 jours."),
                    ("Données à obtenir", "Politique de crédit lisible, encaissements hors Solde, sens des sous-portefeuilles CRF, CRO, CRS.")],
        "Le risque principal n'est pas la perte sur prêts, limitée, mais la mobilisation de l'épargne des mutualistes pour financer cette croissance.")

# ---------------------------------------------------------------- PARTIE 8 : RETRAITÉS ET ANNEXES
separateur(8, "Retraités et annexes", "CGRAE, flotte automobile, MEDOC")
contenu({"titre": "Retraités CGRAE", "sous_titre": "Les retraités cotisants sont passés de 16 à 610 en deux ans et demi",
         "synthese": f"Précompte moyen de {F('Moyen CGRAE (r)', sum(c['moyen'] for c in CG) / len(CG))} par mois sur la pension ; données depuis février 2024"},
        "Les retraités cotisent sur leur pension par la CGRAE. Leur coût santé, plus élevé que celui des actifs, devra être comparé à leur cotisation dès réception des factures. La convention avec la CGRAE n'a pas été reçue.",
        graphique={"titre": fig("Retraités précomptés par la CGRAE, par mois"), "type": "lignes", "categories": [c["libelle"] for c in CG],
                   "series": [{"nom": "Retraités", "valeurs": [c["retraites"] for c in CG], "couleur": BLEU}], "format": "0", "etiquettes": False, "legende": None,
                   "unite": "En nombre de retraités"},
        texte=[f"**{N('Retraités août 2026', CG[-1]['retraites'])} retraités sont précomptés en août 2026.**",
               "**Plusieurs anomalies de gestion apparaissent** : précomptes en double ou indus remboursés, numéros de décision erronés, précomptes non effectués.",
               "**La convention avec la CGRAE doit être transmise.**"],
        source=SRC_PREC)
contenu({"titre": "Produits annexes", "sous_titre": "Deux produits sont précomptés sans documentation transmise",
         "synthese": "Flotte automobile (9W3) et MEDOC (8Z7) : nature, garanties et comptabilisation à préciser"},
        "Ces produits ne relèvent pas directement des garanties de base, mais ils transitent par le FPPN et engagent sa responsabilité vis-à-vis des adhérents.",
        tableau={"titre": tab("Produits annexes et points à clarifier"), "entetes": ["Produit", "Ce que montrent les données", "À obtenir du FPPN"], "largeurs": [5.0, 13.6, 12.6],
                 "lignes": [["Flotte automobile GNA (9W3)", "Précomptes mensuels de 2023 à 2026, environ 10 000 FCFA par adhérent", "Convention avec l'assureur, primes reversées, sinistres, rôle du FPPN"],
                            ["MEDOC (8Z7)", "5 000 FCFA par mois depuis novembre 2025 ; 121 contrats signés en novembre et décembre 2025 ; retours vides en février et avril 2026", "Nature du produit, garanties, bénéficiaires, comptabilisation"],
                            ["Identifiants « 183FPPN »", "34 identifiants sans mécano dans les fichiers aller PEP et Grenier de 2025 et 2026", "Identité et mécano de ces personnes"]]},
        source=SRC_PREC)

# ---------------------------------------------------------------- PARTIE 9 : BILAN
separateur(9, "Bilan et placements", "Provisions, structure de l'actif, fonds propres, rendement")
contenu({"titre": "Méthode : lecture du bilan", "sous_titre": "Le bilan est reconstitué compte par compte à partir des balances",
         "synthese": "Provisions, épargne, actifs et fonds propres sont calculés à chaque clôture, de 2020 au 31/08/2026"},
        "Le bilan n'a pas été transmis : il est reconstitué à partir des soldes des balances. Fonds propres = réserves + report à nouveau + pertes reportées + résultat de l'exercice. Le 31/08/2026 est provisoire et ne comprend pas les écritures de fin d'année.",
        methode=[("Provisions techniques", "Comptes 19011000, 19012000 (jusqu'en 2023) puis 15801000, 15802000", "Prestations à payer"),
                 ("Fonds propres", "Comptes 10, 11, 12 + résultat de l'exercice", "À chaque clôture"),
                 ("Mois couverts", "Provisions / prestations mensuelles moyennes de l'année", "Indicateur de cohérence")],
        source="Source : Adenka")
pvx = EX["provisions"]
contenu({"titre": "Calcul pas à pas : provisions", "sous_titre": "Les deux comptes de provisions portent exactement le même montant sept ans de suite",
         "synthese": f"Maladie {F('Prov mal', b[2020]['provision_maladie'])} et obsèques {F('Prov obs', b[2020]['provision_obseques'])} à chaque clôture"},
        "Lecture : (1) même montant au franc près de 2020 à 2026 ; (2) seul le numéro de compte change en 2024 (reclassement du compte 19 vers le compte 158) ; (3) une estimation recalculée chaque année varierait avec les dossiers en attente. La provision n'est donc pas une estimation actuelle.",
        pas_a_pas={"entetes": ["Clôture", "Compte maladie", "Provision maladie", "Compte obsèques", "Provision obsèques"],
                   "lignes": [[a if a != "2026" else "31/08/2026", pvx[a][0][0], nb(-float(pvx[a][0][2])) + " FCFA", pvx[a][1][0], nb(-float(pvx[a][1][2])) + " FCFA"] for a in sorted(pvx)],
                   "largeurs": [3.6, 4.0, 6.8, 4.0, 6.8], "rouge": [(0, 2), (6, 2), (4, 1)],
                   "explications": ["Même montant en 2020…", "… et au 31/08/2026.", "Seul le compte change en 2024 (19 vers 158)."],
                   "fichier": "AUDIT ADENKA\\DAFC\\2020 à 2026 (balances des comptes)"},
        source=SRC_COMPTA)
pm_ = C["provisions_mois"]
contenu({"titre": "Provisions techniques", "sous_titre": "La provision ne couvre plus que 21 mois de prestations, contre 115 en 2021",
         "synthese": f"Provisions / prestations mensuelles : {N('Mois 2021', pm_[1]['mois_couverts'], 2)} mois en 2021, {N('Mois 2025', pm_[5]['mois_couverts'], 2)} en 2025, {N('Mois 2026', pm_[6]['mois_couverts'], 2)} en 2026"},
        "Le nombre de mois couverts n'est pas une norme : c'est un indicateur de cohérence. Une provision figée perd son sens quand les prestations changent : elle était très supérieure aux besoins en 2021 et diminue en proportion depuis. Seul un calcul fondé sur les dossiers dira si elle est suffisante.",
        graphique={"titre": fig("Mois de prestations couverts par les provisions"), "type": "barres", "categories": cats_ex,
                   "series": [{"nom": "Mois couverts", "valeurs": [round(p["mois_couverts"], 2) for p in pm_], "couleur": BLEU}], "format": "0.00", "legende": None,
                   "unite": "En nombre de mois de prestations"},
        texte=[f"**La provision vaut {F('Prov tot (p)', b[2026]['provisions_techniques'])} chaque année.**",
               f"**Les prestations mensuelles passent de {F('Prest mens 2021', pm_[1]['prestations_mensuelles'])} en 2021 à {F('Prest mens 2026', pm_[6]['prestations_mensuelles'])} en 2026.**",
               "**Le calcul (PSAP, Chain Ladder) suppose les bases de prestations demandées** (partie 11)."],
        source=SRC_COMPTA)
contenu({"titre": "Structure de l'actif", "sous_titre": "L'actif repose sur des créances et des investissements peu liquides",
         "synthese": f"Au 31/08/2026 : CMP {F('CMP (a)', b[2026]['cmp_investissements'])}, terrains {F('Terrains', b[2026]['terrains'])}, trésorerie {F('Trésorerie', b[2026]['tresorerie_totale'])}"},
        "Les investissements au CMP (comptes 47111101, 47111104, 47111105) et la créance sur l'Hôpital de la Police (47111100) sont des créances sur des structures liées : leur recouvrement et leur valeur réelle doivent être documentés.",
        graphique={"titre": fig("Principaux postes de l'actif"), "type": "barres_empilees", "categories": yrs_bil,
                   "series": [{"nom": "Investissements CMP", "valeurs": [round(b[a]["cmp_investissements"]) for a in range(2020, 2027)], "couleur": BLEU},
                              {"nom": "Terrains", "valeurs": [round(b[a]["terrains"]) for a in range(2020, 2027)], "couleur": GRIS_F},
                              {"nom": "Créance Hôpital de la Police", "valeurs": [round(b[a]["creance_hopital_police"]) for a in range(2020, 2027)], "couleur": "BBBCBC"},
                              {"nom": "Financement du Grenier", "valeurs": [round(b[a]["grenier_financement"]) for a in range(2020, 2027)], "couleur": ORANGE},
                              {"nom": "Trésorerie", "valeurs": [round(b[a]["tresorerie_totale"]) for a in range(2020, 2027)], "couleur": NOIR}],
                   "format": FMT_M, "etiquettes": False, "legende": "droite", "unite": "En millions de FCFA"},
        texte=[f"**CMP : {F('CMP 2020', b[2020]['cmp_investissements'])} en 2020, {F('CMP 2026', b[2026]['cmp_investissements'])} au 31/08/2026.**",
               f"**La créance sur l'Hôpital de la Police ({F('Hôpital', b[2026]['creance_hopital_police'])}) est inchangée depuis 2020.**",
               f"**Dettes envers les cliniques : {F('Fournisseurs cliniques', b[2026]['fournisseurs_cliniques'])} au 31/08/2026.**"],
        source=SRC_COMPTA)
fp = EX["fonds_propres_2025"]
contenu({"titre": "Calcul pas à pas : fonds propres 2025", "sous_titre": "Les pertes reportées dépassent les réserves, d'où des fonds propres négatifs",
         "synthese": f"Fonds propres 2025 = {F('FP hors résultat 2025', b[2025]['fonds_propres_hors_resultat'])} + résultat {F('Résultat 2025 (fp)', b[2025]['resultat_exercice'])} = {F('FP 2025 (pas)', b[2025]['fonds_propres'])}"},
        "Les comptes de capitaux propres ont un solde créditeur (positif pour le FPPN), sauf le report à nouveau débiteur (pertes passées), affiché ici en négatif. Libellés exacts de la balance 2025.",
        pas_a_pas={"entetes": ["Compte", "Intitulé (balance 2025)", "Montant"], "lignes": [[c[0], c[1], nb(-float(c[2])) + " FCFA"] for c in fp["comptes"]] + [["", "Sous-total capitaux propres", F("Sous-total FP", b[2025]["fonds_propres_hors_resultat"])], ["", "Résultat de l'exercice 2025", F("Résultat (pas FP)", b[2025]["resultat_exercice"])], ["", "Fonds propres", F("FP (pas total)", b[2025]["fonds_propres"])]],
                   "largeurs": [3.4, 11.4, 6.6], "rouge": [(3, 2), (6, 2)],
                   "explications": ["Pertes reportées : plus que les réserves.", "Fonds propres négatifs malgré le résultat positif.", "Les excédents futurs devront absorber ce solde."],
                   "fichier": "AUDIT ADENKA\\DAFC\\2025\\BALANCE FPPN 2025.xlsx"},
        source="Source : balance des comptes 2025")
contenu({"titre": "Fonds propres", "sous_titre": "Les fonds propres restent négatifs malgré six années d'excédents",
         "synthese": f"{F('FP 2020 (f)', b[2020]['fonds_propres'])} fin 2020, {F('FP 2025 (f)', b[2025]['fonds_propres'])} fin 2025, {F('FP 2026 (f)', b[2026]['fonds_propres'])} au 31/08/2026"},
        "Des fonds propres négatifs signifient que les dettes dépassent les actifs comptables. Le respect des ratios de solvabilité du règlement UEMOA sera vérifié dès que les seuils auront été relevés dans le texte intégral.",
        graphique={"titre": fig("Fonds propres et résultat de l'exercice"), "type": "barres", "categories": yrs_bil,
                   "series": [{"nom": "Fonds propres", "valeurs": [round(b[a]["fonds_propres"]) for a in range(2020, 2027)], "couleur": BLEU},
                              {"nom": "Résultat de l'exercice", "valeurs": [round(b[a]["resultat_exercice"]) for a in range(2020, 2027)], "couleur": ORANGE}],
                   "format": FMT_M, "unite": "En millions de FCFA"},
        texte=["**Les excédents de 2020 à 2025 réduisent l'écart sans le combler.**",
               "**Le déficit des huit premiers mois de 2026 inverse la tendance.**",
               "**Le redressement dépend de la maîtrise de la sinistralité maladie.**"],
        source=SRC_COMPTA)
rp = R["rendement_placements"]
contenu({"titre": "Rendement des placements", "sous_titre": "Le problème des placements est leur faible volume, pas leur rendement",
         "synthese": f"Rendement moyen apparent 2021 à 2025 : {pc('Rendement moyen', sum(r['produits_financiers'] for r in rp) / sum(r['placements_moyens'] for r in rp))} par an ; taux directeur BCEAO entre 2,00 % et 3,50 %"},
        "Rendement apparent = produits financiers (classe 77) / encours moyen des placements (52 à 53). Il est irrégulier : l'essentiel des intérêts est comptabilisé en 2025 à la clôture d'un dépôt à terme. Taux directeurs : communiqués du Comité de politique monétaire de la BCEAO (validés).",
        graphique={"titre": fig("Taux directeur de la BCEAO"), "type": "lignes", "categories": [t[0] for t in R["externe"]["taux_directeur"]],
                   "series": [{"nom": "Taux directeur", "valeurs": [t[1] for t in R["externe"]["taux_directeur"]], "couleur": BLEU}], "format": "0.00", "legende": None,
                   "unite": "En %, date d'effet"},
        texte=[f"**Produits financiers 2021 à 2025 : {F('Produits fin', sum(r['produits_financiers'] for r in rp))}**, pour un encours moyen de placements d'environ {F('Placements moyens', sum(r['placements_moyens'] for r in rp) / len(rp))}.",
               "**Le rendement dépasse le taux directeur, mais il est irrégulier.**",
               f"**Le vrai sujet est le volume placé, face à {F('Épargne (rend)', b[2026]['epargne_totale'])} d'épargne à restituer.**"],
        source="Source : BCEAO, communiqués du Comité de politique monétaire ; " + SRC_COMPTA[9:])
retenir("bilan", [("Le bilan combine provisions figées, fonds propres négatifs et actifs peu liquides", "Trois fragilités cumulées qui limitent la capacité du FPPN à faire face à un choc"),
                  ("Provisions", "Montant identique de 2020 à 2026, à recalculer à partir des dossiers."),
                  ("Fonds propres", "Négatifs à chaque clôture, redressement interrompu en 2026."),
                  ("Actifs", "CMP, terrains et créances dominent ; trésorerie faible face à l'épargne due.")],
        "Ces constats appellent une décision de gouvernance : politique de provisionnement, politique de placement, encadrement des investissements dans les structures liées.")

# ---------------------------------------------------------------- PARTIE 10 : MORTALITÉ
separateur(10, "Tables de mortalité", "Rôle, choix des tables, indicateurs de référence")
contenu({"titre": "Pourquoi des tables de mortalité", "sous_titre": "Les tables servent à mesurer et à évaluer tout engagement lié à la vie humaine",
         "synthese": "Elles préparent deux analyses de la mission : la mortalité des policiers et les engagements viagers"},
        "Une table de mortalité donne, pour chaque âge, la probabilité de mourir dans l'année (qx) et le nombre de survivants sur 1 000 000 de naissances (Lx). Pour le FPPN, elle sert à : (1) comparer la mortalité réelle des policiers à une référence, ce qui dit si la cotisation décès est bien calibrée ; (2) calculer la valeur des engagements qui durent toute la vie (prise en charge santé des retraités, rentes). Ces deux analyses demandent les dates de naissance (PBC, items 9 et 28).",
        methode=[("qx", "Probabilité de décéder entre l'âge x et x + 1", "Lue dans la table"),
                 ("Lx", "Nombre de survivants à l'âge x sur 1 000 000", "Lue dans la table"),
                 ("Usage", "Comparer la mortalité observée ; évaluer les engagements viagers", "En attente des dates de naissance")],
        source=SRC_CIMA)
contenu({"titre": "Choix des tables", "sous_titre": "Le principe de prudence fixe la table selon le risque porté",
         "synthese": "CIMA H pour le risque décès, CIMA F pour le risque de survie, TD et TV en comparaison historique"},
        "Principe de prudence : on choisit la table qui défavorise l'organisme. Pour un risque décès (le FPPN paie au décès), une table qui prévoit plus de décès est prudente : CIMA H. Pour un risque de survie (le FPPN paie tant que la personne vit), une table qui prévoit des vies plus longues est prudente : CIMA F. Les fichiers fournis portent eux-mêmes cet usage dans le nom de leur feuille. Les tables TD et TV sont les anciennes tables réglementaires remplacées en 2012 (Code CIMA, article 338 modifié).",
        tableau={"titre": tab("Tables de mortalité CIMA et usage retenu"), "entetes": ["Table", "Feuille du fichier", "Risque", "Usage dans la mission"], "largeurs": [3.8, 6.0, 7.4, 14.0],
                 "lignes": [["CIMA H", MORT["CIMA H"]["feuille"], "Décès (paiement au décès)", "Calibrage de la cotisation décès ; comparaison à la mortalité observée"],
                            ["CIMA F", MORT["CIMA F"]["feuille"], "Survie (paiement tant que la personne vit)", "Engagements viagers : prise en charge des retraités, rentes"],
                            ["TD CIMA", MORT["TD CIMA"]["feuille"], "Décès (ancienne table)", "Comparaison historique"],
                            ["TV CIMA", MORT["TV CIMA"]["feuille"], "Survie (ancienne table)", "Comparaison historique"]]},
        source=SRC_CIMA)
cs = C["courbes_survie"]
contenu({"titre": "Courbes de survie", "sous_titre": "Les tables de survie prévoient des vies plus longues que les tables de décès",
         "synthese": f"Survivants à 60 ans sur 100 naissances : {N('Survie60 H', 100 * MORT['CIMA H']['survie_60'], 2)} (CIMA H), {N('Survie60 F', 100 * MORT['CIMA F']['survie_60'], 2)} (CIMA F)"},
        "Une courbe de survie donne la part des personnes encore en vie à chaque âge. L'écart entre les courbes mesure l'écart de prudence entre une table de décès et une table de survie.",
        graphique={"titre": fig("Survivants pour 100 naissances, par âge"), "type": "lignes", "categories": [str(a) for a in cs["ages"]],
                   "series": [{"nom": n, "valeurs": [round(v, 2) for v in cs[n]], "couleur": c} for n, c in (("CIMA H", BLEU), ("CIMA F", ORANGE), ("TD CIMA", GRIS_F), ("TV CIMA", "BBBCBC"))],
                   "format": "0.00", "etiquettes": False, "unite": "En % des naissances"},
        texte=[f"**Espérance de vie à la naissance : {N('e0 H', MORT['CIMA H']['e0'], 2)} ans (CIMA H), {N('e0 F', MORT['CIMA F']['e0'], 2)} ans (CIMA F).**",
               f"**Les anciennes tables prévoient davantage de décès** : {N('e0 TD', MORT['TD CIMA']['e0'], 2)} ans (TD), {N('e0 TV', MORT['TV CIMA']['e0'], 2)} ans (TV).",
               "**Le choix de la table change donc fortement les montants calculés.**"],
        source=SRC_CIMA)
cq = C["courbes_qx"]
contenu({"titre": "Probabilités de décès", "sous_titre": "La probabilité de décès double environ tous les dix ans aux âges actifs",
         "synthese": f"Probabilité de décès à 40 ans : {N('q40 H', 1000 * MORT['CIMA H']['q40'], 2)} pour mille (CIMA H), {N('q40 F', 1000 * MORT['CIMA F']['q40'], 2)} pour mille (CIMA F)"},
        "qx en pour mille aux âges de 20 à 80 ans. Ce sont les âges de la population du FPPN (actifs puis retraités). La comparaison avec la mortalité observée se fera par le rapport décès observés / décès attendus (voir partie 11).",
        graphique={"titre": fig("Probabilité de décès dans l'année (qx), par âge"), "type": "lignes", "categories": [str(a) for a in cq["ages"]],
                   "series": [{"nom": n, "valeurs": [round(v, 2) for v in cq[n]], "couleur": c} for n, c in (("CIMA H", BLEU), ("CIMA F", ORANGE), ("TD CIMA", GRIS_F), ("TV CIMA", "BBBCBC"))],
                   "format": "0.00", "etiquettes": False, "unite": "En pour mille"},
        texte=[f"**À 55 ans : {N('q55 H', 1000 * MORT['CIMA H']['q55'], 2)} pour mille (CIMA H) et {N('q55 F', 1000 * MORT['CIMA F']['q55'], 2)} pour mille (CIMA F).**",
               f"**Entre 20 et 60 ans, CIMA H prévoit {pc('H vs TD', -C['mortalite_comparaison']['deces_20_60_H_vs_TD'])} de décès de moins que TD**, et CIMA F {pc('F vs TV', -C['mortalite_comparaison']['deces_20_60_F_vs_TV'])} de moins que TV.",
               "**La table ronde CIMA FANAF de 2011 indique 17 % pour CIMA H et environ 4 % pour CIMA F, sur l'ensemble des âges.**"],
        source=SRC_CIMA)
contenu({"titre": "Indicateurs de référence", "sous_titre": "Une rente viagère vaut entre 12 et 18 années de paiement selon l'âge et la table",
         "synthese": f"Valeur d'une rente de 1 FCFA par an à 60 ans, au taux de 3,5 % : {N('a60 F', MORT['CIMA F']['a60'], 2)} (CIMA F), {N('a60 H', MORT['CIMA H']['a60'], 2)} (CIMA H)"},
        "Rente viagère : valeur aujourd'hui de 1 FCFA payé chaque année, au début de l'année, tant que la personne vit, actualisé au taux de 3,5 % (plafond du Code CIMA). Formule : ä_x = Σ (L_(x+k) / L_x) × v^k, avec v = 1 / 1,035. Exemple : si le FPPN prenait en charge 100 000 FCFA par an de soins pour un retraité de 60 ans, l'engagement vaudrait environ 100 000 × ä_60. Ce sont des valeurs de référence, pas des engagements du FPPN.",
        tableau={"titre": tab("Espérance de vie et valeur de la rente viagère (taux de 3,5 %)"), "entetes": ["Table", "e55 (ans)", "e60 (ans)", "ä55", "ä60", "ä65"], "largeurs": [5.0, 5.0, 5.0, 5.0, 5.0, 5.0],
                 "lignes": [[t, N(f"e55 {t}", MORT[t]["e55"], 2), N(f"e60 {t}", MORT[t]["e60"], 2), N(f"a55 {t}", MORT[t]["a55"], 2), N(f"a60 {t}", MORT[t]["a60"], 2), N(f"a65 {t}", MORT[t]["a65"], 2)] for t in ("CIMA H", "CIMA F", "TD CIMA", "TV CIMA")]},
        source=SRC_CIMA)
retenir("tables de mortalité", [("Les tables CIMA sont prêtes ; il manque les dates de naissance pour les appliquer", "Choix prudent : CIMA H pour le décès, CIMA F pour la survie"),
                                ("Rôle", "Mesurer la mortalité des policiers et évaluer les engagements viagers."),
                                ("Impact du choix", "Une rente à 60 ans vaut 16,84 années de paiement avec CIMA F, 14,26 avec CIMA H."),
                                ("Condition", "Fichier des mutualistes avec date de naissance et sexe (PBC, item 9) et base des décès (item 28).")],
        "Les indicateurs présentés sont des références techniques. Ils deviendront des engagements chiffrés du FPPN dès réception des données.")

# ---------------------------------------------------------------- PARTIE 11 : MÉTHODES À VENIR
separateur(11, "Méthodes actuarielles", "À venir : provisions, mortalité, ratios, projection")
contenu({"titre": "Provisions : méthode PSAP", "sous_titre": "La provision pour sinistres à payer additionne dossiers connus et sinistres non déclarés",
         "synthese": "PSAP = dossiers déclarés non payés + sinistres survenus non encore déclarés (IBNR)"},
        "Les dossiers connus non payés se comptent dossier par dossier. Les sinistres non déclarés (tardifs) s'estiment à partir des délais observés dans le passé : c'est le rôle de la méthode Chain Ladder (diapositive suivante). Données nécessaires : bases des prestations avec date de survenance, de déclaration et de paiement ; liste des dossiers non payés à chaque clôture.",
        methode=[("Dossiers connus", "Somme des montants dus des dossiers déclarés non payés", "PBC 14, 29"),
                 ("Tardifs (IBNR)", "Estimés par Chain Ladder sur les délais passés", "PBC 13, 28, 40"),
                 ("PSAP", "Dossiers connus + tardifs, par garantie et par clôture", "À comparer aux provisions au bilan")],
        attente=True, source="Source : Adenka")
contenu({"titre": "Méthode Chain Ladder", "sous_titre": "Le triangle des paiements révèle le rythme de règlement des sinistres",
         "synthese": "Chaque ligne est une année de survenance, chaque colonne une année de développement"},
        "Principe : on range les paiements cumulés dans un triangle (année de survenance × délai). Le facteur de développement f_j = Σ C(i, j+1) / Σ C(i, j) mesure la progression moyenne d'un délai au suivant. On complète le triangle avec ces facteurs ; la charge ultime moins les paiements déjà faits donne la provision. Hypothèse : le rythme de règlement passé se répète. Aucun chiffre n'est présenté : les données ne sont pas encore reçues.",
        triangle=True, attente=True, source="Source : Adenka (méthode actuarielle standard)")
contenu({"titre": "Mortalité observée", "sous_titre": "Le rapport décès observés / décès attendus mesure l'écart à la table",
         "synthese": "Ratio A/E = décès observés / Σ qx(table) × exposition au risque, par âge"},
        "Exposition : nombre d'années-personnes passées à chaque âge pendant la période (calculée à partir des dates de naissance, d'entrée et de sortie). Décès attendus : exposition × qx de la table CIMA H. Un ratio A/E supérieur à 100 % signifie une mortalité plus forte que la table ; un intervalle de confiance (loi de Poisson) dira si l'écart est significatif.",
        methode=[("Exposition", "E_x = années-personnes passées à l'âge x", "Dates de naissance, entrée, sortie"),
                 ("Décès attendus", "Σ E_x × qx (CIMA H)", "Table fournie"),
                 ("Ratio A/E", "Décès observés / décès attendus, avec IC de Poisson", "Base des décès (PBC 28)")],
        attente=True, source="Source : Adenka")
contenu({"titre": "Ratios AIRMS et solvabilité", "sous_titre": "Les ratios réglementaires comparent réserves, engagements et placements",
         "synthese": "Les seuils seront relevés dans le texte intégral du règlement n° 07/2009/CM/UEMOA avant tout calcul"},
        "Les familles de ratios habituelles pour une mutuelle sociale sont : couverture des engagements par des actifs admis, niveau des réserves, part des frais de gestion, liquidité. Les seuils exacts n'ont pas pu être lus dans le texte et ne sont donc pas cités. Les montants nécessaires (provisions, épargne, placements, fonds propres) sont déjà disponibles dans les balances.",
        methode=[("Couverture des engagements", "Actifs admis / (provisions + épargne)", "Seuil à relever"),
                 ("Réserves", "Fonds propres / cotisations ou engagements", "Seuil à relever"),
                 ("Frais de gestion", "Frais / cotisations", "Seuil à relever")],
        attente=True, source="Source : règlement n° 07/2009/CM/UEMOA (texte intégral à exploiter)")
contenu({"titre": "Projection de l'équilibre technique", "sous_titre": "La projection teste si les cotisations futures couvriront les prestations futures",
         "synthese": "Population, salaires, consommation de soins et décès sont projetés année par année sur 5 à 10 ans"},
        "Méthode : on fait vieillir la population d'une année à l'autre (sorties par retraite et décès avec CIMA H, entrées de recrues), on projette les salaires et donc les cotisations, puis les prestations (coût moyen par âge et par statut). L'écart cotisations - prestations - frais donne l'équilibre de chaque année ; des sensibilités testent les hypothèses. Données nécessaires : âges, coût des soins par âge, barèmes, hypothèses de recrutement.",
        methode=[("Population", "N(x+1, t+1) = N(x, t) × (1 - qx - taux de sortie) + entrées", "Âges, motifs de sortie"),
                 ("Cotisations", "Σ salaires projetés × 5 %", "Barème officiel"),
                 ("Prestations", "Σ effectifs par âge × coût moyen par âge", "Base des prestations")],
        attente=True, source="Source : Adenka")
contenu({"titre": "Travaux en attente des données", "sous_titre": "Ces méthodes seront appliquées dès réception des données listées",
         "synthese": "Voir la liste des documents demandés (PBC pointée) et la partie 13 du rapport"},
        "Aucun chiffre n'est présenté pour ces méthodes : les appliquer sans les données conduirait à des résultats sans valeur. Chaque méthode est reliée aux numéros de la liste des documents demandés.",
        tableau={"titre": tab("Méthodes en attente et données nécessaires"), "entetes": ["Méthode", "Données nécessaires", "PBC"], "largeurs": [8.0, 18.0, 5.2],
                 "lignes": [["PSAP et Chain Ladder", "Bases des prestations avec dates ; dossiers non payés à chaque clôture", "13, 14, 28, 29, 40"],
                            ["Mortalité observée (A/E)", "Mutualistes avec date de naissance, sexe, entrée, sortie ; base des décès", "9, 10, 11, 28"],
                            ["Ratios AIRMS", "Texte intégral du règlement UEMOA ; liste des placements", "18, 23"],
                            ["Projection de l'équilibre", "Âges, barèmes, coût des soins par âge, hypothèses de recrutement", "1, 2, 3, 9, 40"],
                            ["Engagements d'épargne", "Comptes individuels PEP et EFC ; règlement des produits", "16, 17, 18, 19"]]},
        source="Source : liste des documents demandés (PBC) pointée au 18/10/2026")

# ---------------------------------------------------------------- PARTIE 12 : OUTILS
separateur(12, "Outils et traçabilité", "Chaîne de traitement, extraits de code, contrôle des chiffres")
contenu({"titre": "Environnement de travail", "sous_titre": "Les traitements sont écrits en Python et exécutés dans VS Code",
         "synthese": "Huit scripts conservés, rejouables à chaque nouvel envoi du FPPN"},
        "Les images des diapositives suivantes sont des extraits réels des scripts de la mission, mis en forme comme dans l'éditeur VS Code (ce ne sont pas des captures d'écran). Les numéros de ligne correspondent aux scripts du dossier de la mission.",
        tableau={"titre": tab("Scripts de la mission"), "entetes": ["Script", "Rôle", "Contrôle intégré"], "largeurs": [7.6, 13.0, 10.6],
                 "lignes": [["inventaire_fppn.py", "Inventaire et résumé des 4 287 fichiers reçus", "Doublons, fichiers illisibles"],
                            ["extraction_fppn.py", "Extraction sans perte, archive découpée pour l'envoi", "Empreintes SHA-256, lignes et cellules"],
                            ["base_fppn.py", "Base d'analyse : lecture par en-tête ou par contenu", "Lignes de total, montants doublés"],
                            ["grenier_fppn.py", "Lecture des exports du logiciel de prêts", "Totaux imprimés retrouvés"],
                            ["compta_fppn.py", "Lecture des balances Sage", "Débit = crédit, totaux"],
                            ["analyses_fppn.py, analyses_complements_fppn.py", "Indicateurs, tests, sensibilités, tables CIMA", "Classeur des calculs"],
                            ["rapport_v2_fppn.py", "Rapport, graphiques et tableaux natifs", "Contrôle automatique des chiffres"]]},
        source="Source : dépôt de la mission (dossier scripts)")
for nomimg, titre, sous, synth, expl in [
        ("code_extraction", "Code : reconnaître le vrai format", "Le script lit les premiers octets pour reconnaître le vrai format d'un fichier", "C'est ainsi que les 70 fichiers « .xls » du Grenier ont été reconnus comme des pages web lisibles",
         ["**Un fichier Excel moderne commence par « PK »**, un ancien par d'autres octets.", "**Une page web contient « <html » ou « <table ».**", "**Ce test évite de déclarer illisible un fichier simplement mal nommé.**"]),
        ("code_total", "Code : écarter les lignes de total", "Une ligne égale à la somme des autres est écartée automatiquement", "Règle appliquée feuille par feuille, avec une tolérance de 1 %",
         ["**On calcule la somme de la feuille.**", "**Une ligne dont le montant égale la somme des autres est un total**, pas un mutualiste.", "**Chaque ligne écartée est enregistrée** avec son fichier et son montant."]),
        ("code_double", "Code : corriger les montants doublés", "Les montants écrits deux fois à la suite sont corrigés et signalés", "Correction limitée au poste décès, au-delà d'un million de FCFA",
         ["**Le montant est découpé en deux moitiés.**", "**Si les deux moitiés sont identiques**, la première est retenue.", "**Le montant d'origine est conservé** dans une colonne à part."]),
        ("code_grenier", "Code : lire les montants du Grenier", "Les montants imprimés avec espaces et « XOF » sont convertis sans perte", "Les totaux recalculés sont égaux aux totaux imprimés dans les 37 balances âgées",
         ["**Les espaces, y compris insécables, sont retirés.**", "**La mention « XOF » est retirée.**", "**Le nombre obtenu est comparé au total imprimé** par le logiciel."]),
        ("code_sc", "Code : calculer la sinistralité", "Le ratio S/C est calculé compte par compte, pour chaque exercice", "Chaque compte utilisé est cité dans l'annexe des méthodes",
         ["**Cotisations : comptes 7010 et 7021.**", "**Prestations : comptes 601, 6013, 6018, 60800007, 6021, 6024.**", "**Le résultat est écrit dans le classeur des calculs.**"]),
        ("code_stats", "Code : intervalle de confiance", "L'intervalle de Wilson encadre la proportion de prêts en retard", "Méthode adaptée aux proportions, même faibles",
         ["**p = nombre de prêts en retard / nombre de prêts.**", "**L'intervalle est centré et élargi** pour rester entre 0 et 1.", "**z = 1,96 pour un niveau de confiance de 95 %.**"])]:
    contenu({"titre": titre, "sous_titre": sous, "synthese": synth},
            "Extrait réel du script, mis en forme comme dans VS Code (thème sombre, numéros de ligne). Le code complet est dans le dossier scripts de la mission ; il peut être relu et rejoué par un tiers.",
            image={"fichier": os.path.join(IMG, nomimg + ".png"), "legende": fig(f"Extrait du script, mis en forme comme dans VS Code ({titre.split(': ')[1]})")}, texte=expl,
            source="Source : scripts de la mission (extrait mis en forme, pas une capture d'écran)")
contenu({"titre": "Contrôle automatique des chiffres", "sous_titre": "Chaque chiffre du rapport est recalculé et comparé au texte affiché",
         "synthese": "Le contrôle relit le PowerPoint produit et vérifie que chaque chiffre enregistré figure sur sa diapositive"},
        "Principe : à la construction, chaque chiffre est écrit à partir de la valeur calculée et enregistré (diapositive, valeur exacte, texte affiché). Après la construction, un script relit le fichier PowerPoint et vérifie la présence de chaque texte sur sa diapositive. Le résultat figure dans l'onglet « Traçabilité » du classeur des calculs.",
        methode=[("Enregistrement", "Chaque chiffre est produit par une fonction qui garde sa valeur exacte", "Diapositive, indicateur, valeur"),
                 ("Relecture", "Le PowerPoint produit est relu, textes et tableaux compris", "Après construction"),
                 ("Résultat", "Chiffres retrouvés / chiffres enregistrés", "Onglet Traçabilité")],
        source="Source : Adenka")

# ---------------------------------------------------------------- PARTIE 13 : ANALYSES EN ATTENTE ET RISQUES
separateur(13, "Analyses en attente", "Documents à obtenir, risques, points à clarifier")
contenu({"titre": "Analyses en attente", "sous_titre": "Les analyses actuarielles centrales dépendent de trois bases non reçues",
         "synthese": "Provisions, mortalité, équilibre technique et contrôles santé seront réalisés dès réception des documents"},
        "Liste détaillée et à jour dans la PBC pointée.",
        tableau={"titre": tab("Analyses en attente et documents nécessaires"), "entetes": ["Analyse", "Objet", "Documents nécessaires (PBC)"], "largeurs": [7.2, 12.4, 11.6],
                 "lignes": [["Provisions (PSAP, Chain Ladder)", "Montant restant à payer, y compris non déclarés", "Prestations décès et santé, dossiers non payés (13, 14, 28, 29, 40)"],
                            ["Sinistralité par année de survenance", "Coût réel de chaque année", "Bases des prestations (13, 28, 40)"],
                            ["Mortalité observée (A/E)", "Comparaison à CIMA H", "Mutualistes avec date de naissance (9) ; décès (28)"],
                            ["Délais et dossiers en attente", "Délais de déclaration et de paiement", "Bases des prestations avec dates (28, 40)"],
                            ["Contrôles santé (S1 à S7)", "Doublons, bénéficiaires hors fichier", "Factures santé (40) ; cartes actives (42)"],
                            ["Engagements d'épargne", "Dette envers chaque épargnant", "Comptes individuels (16, 17)"],
                            ["Ratios AIRMS, équilibre technique", "Solvabilité, projection", "Règlement UEMOA, âges, barèmes (1, 2, 3)"]]},
        source="Source : PBC pointée au 18/10/2026")
contenu({"titre": "Points à clarifier", "sous_titre": "Huit points demandent une réponse du FPPN avant la prochaine version",
         "synthese": "Chaque point est relié à un constat chiffré du rapport"},
        "Ces questions peuvent être posées lors des prochains entretiens ; les réponses seront intégrées à la version suivante du rapport.",
        tableau={"titre": tab("Points à clarifier avec le FPPN"), "entetes": ["N°", "Point", "Question au FPPN"], "largeurs": [1.2, 9.0, 21.0],
                 "lignes": [["1", "Répartition du taux de 5 %", "Quel barème s'applique : 3,5 / 1,5 ou 3,77 / 1,23 ? Pourquoi la comptabilité suit-elle une autre clé ?"],
                            ["2", "Provisions figées", "Comment ont-elles été calculées ? Pourquoi n'ont-elles pas changé depuis 2020 ?"],
                            ["3", "Couverture de l'épargne", "Quelles sont les conditions de retrait du PEP et de l'EFC ?"],
                            ["4", "Créance sur l'Hôpital de la Police", "Quelle est sa nature ? Un remboursement est-il prévu ?"],
                            ["5", "Retards du Grenier début 2026", "Pourquoi la plupart de l'encours était-il en retard fin janvier 2026 ?"],
                            ["6", "Écart échéances / précomptes", "Comment sont réglées les échéances non précomptées ?"],
                            ["7", "Sinistralité 2020 à 2022", "Des factures de ces années ont-elles été comptabilisées plus tard ?"],
                            ["8", "Données manquantes", "Précomptes 2020 à 2023, fichiers vides de décembre 2024, grands livres endommagés"]]},
        source="Source : analyses Adenka")

# ---------------------------------------------------------------- ANNEXES
SL.append({"type": "separateur", "num": "A", "titre": "Annexes", "sous_titre": "Méthodes, listes, abréviations, références", "notes": ""})
SL.append({"type": "annexe", "titre": "Méthodes et formules", "notes": "Formules exactes utilisées dans le rapport.", "tableau": {"titre": tab("Indicateurs, formules et sources"),
           "entetes": ["Indicateur", "Formule", "Source des données"], "largeurs": [6.4, 14.0, 10.8],
           "lignes": [["Ratio S/C", "Prestations comptabilisées / cotisations comptabilisées, par garantie et par exercice", "Comptes 601, 6013, 6018, 60800007, 6021, 6024 ; 7010, 7021"],
                      ["Taux de cotisation", "Précompte du mois / salaire de base du mois, par agent", "Fichiers 936 détaillés, 8Z6"],
                      ["Taux de précompte", "Montant retour / montant aller, mois communs", "Fichiers aller et retour"],
                      ["TCAM", "(V_fin / V_début)^(1/n) - 1", "Selon l'indicateur"],
                      ["Tendance", "Régression linéaire par moindres carrés ; test de Student sur la pente ; IC à 95 %", "Séries mensuelles ou annuelles"],
                      ["Intervalle de Wilson", "(p + z²/2n ± z √(p(1-p)/n + z²/4n²)) / (1 + z²/n), z = 1,96", "Balance âgée du 31/08/2026"],
                      ["Valeur réelle", "Valeur nominale × 100 / indice des prix (base 100 en 2020)", "Banque mondiale, FP.CPI.TOTL.ZG"],
                      ["Couverture de l'épargne", "Actifs dédiés / (comptes PEP + EFC)", "Balances des comptes"],
                      ["Rente viagère ä_x", "Σ (L_(x+k) / L_x) × (1,035)^(-k)", "Tables CIMA"],
                      ["Espérance de vie e_x", "Σ L_(x+k) / L_x pour k ≥ 1, + 0,5", "Tables CIMA"]]},
           "source": "Source : Adenka"})
SL.append({"type": "annexe", "titre": "Outils et bibliothèques", "notes": "", "tableau": {"titre": tab("Outils et bibliothèques utilisés"),
           "entetes": ["Outil", "Usage", "Version"], "largeurs": [7.0, 18.2, 6.0],
           "lignes": [["Python", "Langage des traitements ; scripts conservés et rejouables", "3.13"],
                      ["Visual Studio Code", "Éditeur et exécution des scripts", "dernière"],
                      ["pandas, numpy", "Tables de données, calculs, agrégations", "3.0 et 2.5"],
                      ["statsmodels", "Régressions et tests statistiques", "0.15"],
                      ["python-calamine, openpyxl, lxml", "Lecture sans perte des fichiers Excel et des pages web", "dernière"],
                      ["pyarrow", "Stockage de la base d'analyse (Parquet)", "26.0"],
                      ["pymupdf, pdftotext", "Lecture et récupération des PDF comptables", "dernière"],
                      ["pygments, Pillow", "Mise en forme des extraits de code", "dernière"],
                      ["python-pptx", "Rapport, graphiques et tableaux natifs", "1.0"]]},
           "source": "Source : Adenka"})
SL.append({"type": "liste_figures", "notes": ""})
SL.append({"type": "liste_tableaux", "notes": ""})
SL.append({"type": "annexe", "titre": "Signes et abréviations", "notes": "", "tableau": {"titre": None,
           "entetes": ["Sigle", "Signification", "Sigle", "Signification"], "largeurs": [3.4, 12.2, 3.4, 12.2],
           "lignes": [["A/E", "Décès observés / décès attendus", "IC", "Intervalle de confiance"],
                      ["AIRMS", "Autorité de régulation de la mutualité sociale", "M / Md", "Millions / milliards de FCFA"],
                      ["ä_x", "Valeur d'une rente viagère de 1 par an à l'âge x", "PAR", "Part de l'encours en retard"],
                      ["BCEAO", "Banque centrale des États de l'Afrique de l'Ouest", "PBC", "Liste des documents demandés au client"],
                      ["CGRAE", "Caisse générale de retraite des agents de l'État", "PEP", "Plan d'épargne (poste 998)"],
                      ["CIMA", "Conférence interafricaine des marchés d'assurances", "PSAP", "Provision pour sinistres à payer"],
                      ["CMP", "Centre médical de la Police", "qx", "Probabilité de décès entre x et x + 1"],
                      ["DAT", "Dépôt à terme", "S/C", "Prestations / cotisations"],
                      ["e_x", "Espérance de vie à l'âge x", "TCAM", "Taux de croissance annuel moyen"],
                      ["EFC", "Entraide fin de carrière (poste 870)", "UEMOA", "Union économique et monétaire ouest-africaine"],
                      ["FCP", "Fonds commun de placement", "936, 8Z6", "Postes maladie, décès"],
                      ["IBNR", "Sinistres survenus non encore déclarés", "9X4, 9W3, 8Z7", "Postes Grenier, flotte, MEDOC"]]},
           "source": None})
SL.append({"type": "annexe", "titre": "Références et sources", "notes": "", "tableau": {"titre": None,
           "entetes": ["Référence", "Contenu utilisé", "Lien ou emplacement"], "largeurs": [8.6, 10.6, 12.0],
           "lignes": [["Envoi du FPPN du 15/10/2026", "Précomptes, balances, exports du Grenier", "Dossier 04-Documents reçus du client"],
                      ["Grands livres 2021, 2022, 2024, 2025 (PDF)", "Détail des écritures", "Envoi complémentaire du 18/10/2026"],
                      ["Règlement n° 07/2009/CM/UEMOA", "Cadre de la mutualité sociale", "flagship.social-protection.org (Media.action?id=9421)"],
                      ["Ordonnance n° 2024-1043", "FOMUS-CI", "civlii.laws.africa"],
                      ["Code CIMA, article 338", "Tables et taux technique de 3,5 %", "cima-afrique.org"],
                      ["Tables CIMA H, CIMA F, TD, TV", "Mortalité de référence", "Fichiers fournis le 18/10/2026"],
                      ["Table ronde CIMA FANAF (2011)", "Comparaison des tables", "fanaf.org"],
                      ["BCEAO, communiqués du CPM 2020 à 2025", "Taux directeurs", "bceao.int"],
                      ["Banque mondiale, FP.CPI.TOTL.ZG", "Inflation 2020 à 2025", "donnees.banquemondiale.org"],
                      ["Classeur des calculs v2", "Tous les tableaux, séries et la traçabilité", "FPPN_Calculs_rapport_20261015_v2.xlsx"]]},
           "source": None})
SL.append({"type": "merci", "notes": ""})

# ===========================================================================
# CONSTRUCTION
# ===========================================================================
GAB = {"couverture": 1, "sommaire": 5, "separateur": 2, "contenu": 5, "annexe": 5, "liste_figures": 5, "liste_tableaux": 5, "merci": 110}
tmp = SORTIE.replace(".pptx", "_tmp.pptx")
selection(TIMESAVER, tmp, [x["gab"] if x["type"] == "gab" else GAB[x["type"]] for x in SL])

CJ = {}
for i, x in enumerate(SL, 1):
    k = str(i)
    if x["type"] == "couverture":
        CJ[k] = {"Google Shape;138;p25": "Rapport actuariel\net statistique\ndu FPPN", "Text Placeholder 4": "Rapport d'étape au 15 octobre 2026, version 2"}
    elif x["type"] == "separateur":
        CJ[k] = {"Rectangle 24": str(x["num"]), "Rectangle 9": x["titre"], "Rectangle 5": x["sous_titre"]}
    elif x["type"] in ("sommaire", "liste_figures", "liste_tableaux", "annexe"):
        CJ[k] = {"@vider": True}
    elif x["type"] == "gab":
        CJ[k] = {"@entete": x["entete"]}
        if x.get("source"):
            CJ[k]["@source"] = x["source"]
    elif x["type"] == "contenu":
        d = {"@vider": True, "@entete": x["entete"]}
        if "graphique" in x:
            g = dict(x["graphique"])
            d["@texte1"] = {"texte": g.pop("titre"), "style": "titre_graphique", "zone": "titre_graphique"}
            if g.get("unite"):
                d["@texte3"] = {"texte": g.pop("unite"), "style": "unite", "x": 1.4, "y": 4.45, "largeur": 18.2, "hauteur": 0.45}
            g.update({"zone": "graphique_gauche", "y": 4.95, "hauteur": 11.9})
            d["@graphique"] = g
        if "graphiques" in x:
            for j, gg in enumerate(x["graphiques"]):
                g = dict(gg)
                xx = 1.4 + j * 16.0
                d[f"@texte1{j}"] = {"texte": g.pop("titre"), "style": "titre_graphique", "x": xx, "y": 3.85, "largeur": 15.2, "hauteur": 0.7}
                if g.get("unite"):
                    d[f"@texte3{j}"] = {"texte": g.pop("unite"), "style": "unite", "x": xx, "y": 4.45, "largeur": 15.2, "hauteur": 0.45}
                g.update({"x": xx, "y": 4.95, "largeur": 15.2, "hauteur": 11.9})
                d[f"@graphique{j}"] = g
        if "texte" in x:
            d["@texte2"] = {"texte": "\n".join(x["texte"]), "style": "corps", "zone": "texte_droite"}
        if x.get("source"):
            d["@source"] = x["source"]
        CJ[k] = d
remplir(tmp, CJ, tmp)

prs = Presentation(tmp)
ts = Presentation(TIMESAVER)
icones = [sh for sh in ts.slides[110].shapes if sh.shape_type == 13 and sh.width < Cm(3)]


def icone_blanche(blob):
    from PIL import Image
    im = Image.open(io.BytesIO(blob)).convert("RGBA")
    blanc = Image.new("RGBA", im.size, (255, 255, 255, 0))
    alpha = im.split()[3]
    if alpha.getextrema() == (255, 255):
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
    for i, ligne in enumerate(str(texte).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = ligne
        r.font.name = "Helvetica"
        r.font.size = Pt(taille)
        r.font.bold = gras
        r.font.color.rgb = RGBColor.from_string(couleur)
    return tb


def boite(slide, x, y, l, h, texte, fond=BLEU, couleur="FFFFFF", taille=12, gras=True, forme=MSO_SHAPE.ROUNDED_RECTANGLE, contour=None):
    sh = slide.shapes.add_shape(forme, Cm(x), Cm(y), Cm(l), Cm(h))
    if fond:
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(fond)
    else:
        sh.fill.background()
    if contour:
        sh.line.color.rgb = RGBColor.from_string(contour)
        sh.line.width = Pt(1.5)
    else:
        sh.line.fill.background()
    tf = sh.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = Cm(0.15)
    for i, ligne in enumerate(str(texte).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = ligne
        r.font.name = "Helvetica"
        r.font.size = Pt(taille if i == 0 else max(taille - 2, 9))
        r.font.bold = gras and i == 0
        r.font.color.rgb = RGBColor.from_string(couleur)
    return sh


def fleche(slide, x1, y1, x2, y2, couleur=GRIS_F, epaisseur=2.0):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x1), Cm(y1), Cm(x2), Cm(y2))
    c.line.color.rgb = RGBColor.from_string(couleur)
    c.line.width = Pt(epaisseur)
    ln = c.line._get_or_add_ln()
    tail = ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"})
    ln.append(tail)
    return c


def tableau(slide, spec, y0=3.85, x0=1.4, hauteur_ligne=None, taille=None):
    y = y0
    if spec.get("titre"):
        txt(slide, x0, y, sum(spec["largeurs"]), 0.6, spec["titre"], 12, False, ORANGE, PP_ALIGN.CENTER)
        y += 0.75
    lignes, larg = spec["lignes"], spec["largeurs"]
    n = len(lignes) + 1
    if taille is None:
        total = max(len(" ".join(l)) for l in lignes)
        taille = 12 if n <= 9 and total < 200 else (11 if n <= 11 else 10)
    hl = hauteur_ligne or 0.9
    shape = slide.shapes.add_table(n, len(larg), Cm(x0), Cm(y), Cm(sum(larg)), Cm(hl * n))
    t = shape.table
    tblPr = t._tbl.tblPr
    for attr in ("bandRow", "firstRow"):
        tblPr.set(attr, "0")
    st = tblPr.find(qn("a:tableStyleId"))
    if st is not None:
        st.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"
    for j, w in enumerate(larg):
        t.columns[j].width = Cm(w)
    for i in range(n):
        if hauteur_ligne:
            t.rows[i].height = Cm(hauteur_ligne)
        vals = spec["entetes"] if i == 0 else lignes[i - 1]
        for j, v in enumerate(vals):
            c = t.cell(i, j)
            c.margin_left = c.margin_right = Cm(0.12)
            c.margin_top = c.margin_bottom = Cm(0.04)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = c.text_frame.paragraphs[0]
            p.text = ""
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", "0")
            pPr.set("indent", "0")
            pPr.set("lvl", "0")
            for bn in pPr.findall(qn("a:buNone")):
                pPr.remove(bn)
            pPr.insert(0, pPr.makeelement(qn("a:buNone"), {}))
            r = p.add_run()
            r.text = str(v)
            r.font.name = "Helvetica"
            r.font.size = Pt(taille)
            r.font.bold = (i == 0) or (j == 0)
            r.font.color.rgb = RGBColor(255, 255, 255) if i == 0 else RGBColor.from_string(NOIR)
            c.fill.solid()
            c.fill.fore_color.rgb = RGBColor.from_string(NOIR if i == 0 else (GRIS if i % 2 == 0 else "FFFFFF"))
    return shape, y


def pastille(slide, x, y, num):
    return boite(slide, x, y, 0.75, 0.75, str(num), ROUGE, "FFFFFF", 11, True, MSO_SHAPE.OVAL)


def pas_a_pas(slide, spec):
    lignes = spec["lignes"]
    hl = 0.62 if len(lignes) <= 10 else 0.5
    taille = 10 if not spec.get("petit") else 9
    shp, y = tableau(slide, {"titre": None, "entetes": spec["entetes"], "lignes": lignes, "largeurs": spec["largeurs"]}, 4.3, 1.4, hl, taille)
    txt(slide, 1.4, 3.75, sum(spec["largeurs"]), 0.5, "Extrait : " + spec["fichier"], 9, False, GRIS_F)
    larg = spec["largeurs"]
    for k, (ri, cj) in enumerate(spec["rouge"], 1):
        xx = 1.4 + sum(larg[:cj])
        yy = y + hl * (ri + 1)
        cadre = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(xx - 0.05), Cm(yy - 0.03), Cm(larg[cj] + 0.1), Cm(hl + 0.06))
        cadre.fill.background()
        cadre.line.color.rgb = RGBColor.from_string(ROUGE)
        cadre.line.width = Pt(2.25)
        pastille(slide, xx + larg[cj] - 0.5, yy - 0.45, k)
    xr = 1.4 + sum(larg) + 0.6
    lr = 33.0 - xr - 0.4
    if lr < 6:
        xr, lr = 1.4, sum(larg)
        ytxt = y + hl * (len(lignes) + 1) + 0.4
        for k, e in enumerate(spec["explications"], 1):
            pastille(slide, xr, ytxt + (k - 1) * 1.0, k)
            txt(slide, xr + 1.0, ytxt + (k - 1) * 1.0 - 0.05, lr - 1.0, 0.9, e, 12)
    else:
        for k, e in enumerate(spec["explications"], 1):
            yy = 4.4 + (k - 1) * 3.6
            pastille(slide, xr, yy, k)
            txt(slide, xr + 1.0, yy - 0.1, lr - 1.0, 3.3, e, 13)


def methode(slide, items, attente=False):
    y = 4.2
    for k, (titre, formule, donnees) in enumerate(items):
        yy = y + k * 3.7
        boite(slide, 1.4, yy, 6.6, 3.0, titre, BLEU, "FFFFFF", 15)
        cadre = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(8.3), Cm(yy), Cm(16.4), Cm(3.0))
        cadre.fill.solid()
        cadre.fill.fore_color.rgb = RGBColor.from_string(GRIS)
        cadre.line.fill.background()
        txt(slide, 8.6, yy + 0.25, 15.8, 2.6, formule, 14, False, NOIR, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
        boite(slide, 25.0, yy, 7.6, 3.0, donnees, None, BLEU, 12, False, MSO_SHAPE.ROUNDED_RECTANGLE, BLEU)
    if attente:
        boite(slide, 1.4, 15.45, 31.2, 1.25, "Travail en attente des données : voir la liste des données manquantes (PBC pointée et partie 13)", ORANGE, NOIR, 14)


def cartes(slide, items):
    n = len(items)
    l = (31.2 - (n - 1) * 0.6) / n
    for k, (titre, corps) in enumerate(items):
        x = 1.4 + k * (l + 0.6)
        boite(slide, x, 4.3, l, 1.6, titre, BLEU, "FFFFFF", 16)
        cadre = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(x), Cm(6.2), Cm(l), Cm(9.8))
        cadre.fill.background()
        cadre.line.color.rgb = RGBColor.from_string(BLEU)
        cadre.line.width = Pt(1.5)
        txt(slide, x + 0.4, 6.6, l - 0.8, 9.0, corps, 16)
        if icones:
            ic = icones[(k * 7) % len(icones)]
            slide.shapes.add_picture(io.BytesIO(ic.image.blob), Cm(x + l / 2 - 0.8), Cm(13.6), Cm(1.6), Cm(1.6))


def triangle(slide):
    ent = ["Survenance", "Délai 1", "Délai 2", "Délai 3", "Délai 4"]
    lignes = [["Année n-3", "C(1,1)", "C(1,2)", "C(1,3)", "C(1,4)"], ["Année n-2", "C(2,1)", "C(2,2)", "C(2,3)", "?"],
              ["Année n-1", "C(3,1)", "C(3,2)", "?", "?"], ["Année n", "C(4,1)", "?", "?", "?"]]
    shp, y = tableau(slide, {"titre": None, "entetes": ent, "lignes": lignes, "largeurs": [4.6, 3.6, 3.6, 3.6, 3.6]}, 4.3, 1.4, 1.2, 14)
    txt(slide, 1.4, 10.6, 19.0, 4.0, "C(i, j) : paiements cumulés de l'année de survenance i, au délai j.\n« ? » : montants futurs à estimer.\nFacteur de développement : f(j) = Σ C(i, j+1) / Σ C(i, j).\nCharge ultime = dernier montant connu × produit des facteurs suivants.\nProvision = charge ultime - paiements déjà faits.", 13)
    boite(slide, 21.6, 4.3, 11.0, 5.2, "Hypothèse\nLe rythme de règlement passé se répète dans le futur.", GRIS, NOIR, 15, True, MSO_SHAPE.ROUNDED_RECTANGLE)
    boite(slide, 21.6, 10.0, 11.0, 5.0, "Données\nPaiements par année de survenance et par délai, issus des bases de prestations (PBC 13, 28, 40).", GRIS, NOIR, 15, True, MSO_SHAPE.ROUNDED_RECTANGLE)
    boite(slide, 1.4, 15.45, 31.2, 1.25, "Travail en attente des données : voir la liste des données manquantes (PBC pointée et partie 13)", ORANGE, NOIR, 14)


def schema(slide, nom):
    if nom == "fonctionnement":
        boite(slide, 1.4, 7.6, 6.0, 3.2, "Mutualistes\nactifs et retraités", BLEU, "FFFFFF", 16)
        boite(slide, 9.0, 5.0, 6.2, 2.4, "Solde (actifs)", GRIS, NOIR, 14, contour=BLEU)
        boite(slide, 9.0, 11.0, 6.2, 2.4, "CGRAE (retraités)", GRIS, NOIR, 14, contour=BLEU)
        boite(slide, 16.8, 7.2, 6.6, 4.0, "FPPN\nCotisations " + F("Cot (schéma)", s[2025]["cot_maladie"] + s[2025]["cot_deces"]), ORANGE, NOIR, 16)
        sorties = [("Prestations maladie\n" + F("Prest mal (schéma)", s[2025]["prest_maladie"]), 3.9),
                   ("Prestations décès\n" + F("Prest déc (schéma)", s[2025]["prest_deces"]), 6.8),
                   ("Épargne PEP et EFC\n" + F("Épargne (schéma)", b[2026]["epargne_totale"]), 9.7),
                   ("Prêts du Grenier\n" + F("Grenier (schéma)", b[2026]["grenier_financement"]), 12.6),
                   ("CMP et terrains\n" + F("CMP+terrains (schéma)", b[2026]["cmp_investissements"] + b[2026]["terrains"]), 15.5)]
        for t_, yy in sorties:
            boite(slide, 25.2, yy - 1.1, 7.4, 2.3, t_, BLEU, "FFFFFF", 12)
            fleche(slide, 23.4, 9.2, 25.2, yy, GRIS_F)
        fleche(slide, 7.4, 8.6, 9.0, 6.2, GRIS_F)
        fleche(slide, 7.4, 9.8, 9.0, 12.2, GRIS_F)
        fleche(slide, 15.2, 6.2, 16.8, 8.4, GRIS_F)
        fleche(slide, 15.2, 12.2, 16.8, 10.0, GRIS_F)
        txt(slide, 1.4, 14.2, 15.0, 2.4, "Flux 2025 (comptabilité) ; épargne, Grenier, CMP et terrains : soldes au 31/08/2026.\nPrécompte : retenue à la source sur salaire ou pension.", 11, False, GRIS_F)
    elif nom == "precomptes":
        etapes_ = [("1. FPPN", "Prépare le fichier aller : qui précompter et combien"), ("2. Solde / CGRAE", "Retient les montants sur salaire ou pension"),
                   ("3. Fichier retour", "Liste des montants réellement retenus"), ("4. FPPN", "Rapproche aller et retour, comptabilise, relance les rejets")]
        for k, (t_, c_) in enumerate(etapes_):
            x = 1.4 + k * 8.0
            boite(slide, x, 5.2, 6.6, 2.4, t_, BLEU if k % 2 == 0 else ORANGE, "FFFFFF" if k % 2 == 0 else NOIR, 16)
            txt(slide, x, 8.0, 6.6, 3.0, c_, 14, False, NOIR, PP_ALIGN.CENTER)
            if k < 3:
                fleche(slide, x + 6.6, 6.4, x + 8.0, 6.4, GRIS_F, 3)
        boite(slide, 1.4, 12.0, 31.2, 2.6, "Contrôle Adenka : taux de précompte = montant retour / montant aller, poste par poste et mois par mois (partie 3)", GRIS, NOIR, 15, contour=BLEU)
    elif nom == "chaine":
        etapes_ = [("Inventaire", "4 287 fichiers", "Doublons, illisibles"), ("Extraction", "CSV sans perte", "Empreintes SHA-256"),
                   ("Base", "4,45 millions de lignes", "Totaux, montants doublés"), ("Analyses", "Indicateurs, tests", "Classeur des calculs"),
                   ("Rapport", "Graphiques natifs", "Contrôle des chiffres")]
        for k, (t_, c_, ctl) in enumerate(etapes_):
            x = 1.4 + k * 6.35
            boite(slide, x, 5.0, 5.4, 2.6, t_ + "\n" + c_, BLEU, "FFFFFF", 15)
            boite(slide, x, 9.2, 5.4, 2.4, "Contrôle\n" + ctl, None, BLEU, 13, True, MSO_SHAPE.ROUNDED_RECTANGLE, ORANGE)
            fleche(slide, x + 2.7, 7.6, x + 2.7, 9.2, ORANGE, 1.5)
            if k < 4:
                fleche(slide, x + 5.4, 6.3, x + 6.35, 6.3, GRIS_F, 3)
        txt(slide, 1.4, 12.6, 31.2, 3.0, "Schéma modifiable, à compléter au fil des entretiens avec les services du FPPN (sources supplémentaires, nouveaux contrôles).", 13, False, GRIS_F)


def remplir_gab(slide, num, d):
    formes = list(slide.shapes)

    def par_pos(pred):
        return sorted([sh for sh in formes if pred(sh)], key=lambda sh: (round(sh.top / 360000), sh.left))

    def ecrire(sh, lignes_):
        tf = sh.text_frame
        paras = [p for p in tf.paragraphs]
        for k, ligne in enumerate(lignes_):
            if k < len(paras):
                runs = paras[k].runs
                if runs:
                    runs[0].text = ligne
                    for r in runs[1:]:
                        r._r.getparent().remove(r._r)
                else:
                    paras[k].text = ligne
        for p in paras[len(lignes_):]:
            p._p.getparent().remove(p._p)
    if num == 49:
        for nm, t_ in zip(("Freeform 7", "Freeform 9", "Freeform 11", "Freeform 13"), d["titres"]):
            for sh in formes:
                if sh.name == nm:
                    ecrire(sh, [t_])
        for sh, v in zip(par_pos(lambda sh: sh.name == "TextBox 31"), d["valeurs"]):
            ecrire(sh, [v])
            sh.text_frame.word_wrap = False
            sh.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
            for r in sh.text_frame.paragraphs[0].runs:
                r.font.size = Pt(21)
        for sh, v in zip(par_pos(lambda sh: sh.name == "TextBox 96"), d["commentaires"]):
            ecrire(sh, [v])
    elif num == 23:
        for sh, (t_, c_) in zip(par_pos(lambda sh: sh.name == "TextBox 49"), d["blocs"]):
            ecrire(sh, [t_, c_])
    elif num == 83:
        noms = {"CAUSE": "CAUSES", "EFFECT": "EFFETS"}
        for sh in formes:
            if sh.has_text_frame:
                t0 = sh.text_frame.text.strip()
                if t0 in noms:
                    ecrire(sh, [noms[t0]])
                elif t0.startswith("CAUSE 0"):
                    ecrire(sh, ["CAUSE " + t0[-1]])
                elif t0.startswith("EFFECT 0"):
                    ecrire(sh, ["EFFET " + t0[-1]])
        dummies = [sh for sh in formes if sh.has_text_frame and sh.text_frame.text.startswith("This is dummy")]
        cols = sorted(dummies, key=lambda sh: sh.left)
        paires = [sorted(cols[i:i + 2], key=lambda sh: sh.top) for i in range(0, len(cols), 2)]
        for (haut, bas), c_, e_ in zip(paires, d["causes"], d["effets"]):
            ecrire(haut, [c_])
            ecrire(bas, [e_])
    elif num == 39:
        nums = {sh.text_frame.text.strip(): sh for sh in formes if sh.has_text_frame and sh.text_frame.text.strip() in ("01", "02", "03", "04")}
        rects = [sh for sh in formes if sh.name in ("Rectangle 18", "Rectangle 19", "Rectangle 21", "Rectangle 22")]
        for k, (t_, c_) in enumerate(d["etapes"], 1):
            nsh = nums[f"0{k}"]
            r = min(rects, key=lambda sh: abs(sh.top - nsh.top) + abs(sh.left - nsh.left))
            ecrire(r, [t_, c_])
    elif num == 107:
        charts = sorted([sh for sh in formes if getattr(sh, "has_chart", False) and sh.has_chart], key=lambda sh: sh.left)
        for sh, (v, lib, cle) in zip(charts, d["anneaux"]):
            cd = CategoryChartData()
            cd.categories = ["Valeur", "Reste"]
            cd.add_series("Série", (round(100 * v, 2), round(100 - 100 * v, 2)))
            sh.chart.replace_data(cd)
            sh.chart.has_title = False
            delta = sh.top - Cm(5.2)
            for o in formes:
                if o is not sh and sh.left <= o.left + o.width / 2 <= sh.left + sh.width and sh.top <= o.top + o.height / 2 <= sh.top + sh.height:
                    o.top = o.top - delta
            sh.top = Cm(5.2)
            pts = sh.chart.plots[0].series[0].points
            for k_, coul in enumerate((ORANGE, BLEU)):
                pts[k_].format.fill.solid()
                pts[k_].format.fill.fore_color.rgb = RGBColor.from_string(coul)
            txt(slide, sh.left / 360000 - 0.3, sh.top / 360000 + sh.height / 360000 + 0.1, sh.width / 360000 + 0.6, 0.9, pc(cle, v), 20, True, BLEU, PP_ALIGN.CENTER)
            txt(slide, sh.left / 360000 - 0.3, sh.top / 360000 + sh.height / 360000 + 1.0, sh.width / 360000 + 0.6, 2.0, lib, 12, False, NOIR, PP_ALIGN.CENTER)


def sommaire(slide, entrees):
    txt(slide, 2.7, 0.85, 20, 1.1, "SOMMAIRE", 24, True, BLEU)
    pas = 1.37 if len(entrees) <= 9 else 13.2 / len(entrees)
    taille = 16 if len(entrees) <= 9 else (13 if len(entrees) <= 14 else 12)
    centres = [3.6 + i * pas + pas / 2 for i in range(len(entrees))]
    hb = min(1.01, pas - 0.12)
    ligne = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(2.59), Cm(centres[0]), Cm(2.59), Cm(centres[-1]))
    ligne.line.color.rgb = RGBColor.from_string(BLEU)
    ligne.line.width = Pt(1)
    for i, (intitule, num) in enumerate(entrees):
        cy = centres[i]
        h = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(2.59), Cm(cy), Cm(4.85), Cm(cy))
        h.line.color.rgb = RGBColor.from_string(BLEU)
        h.line.width = Pt(1)
        dd = min(1.2, pas - 0.1)
        dsq = slide.shapes.add_shape(MSO_SHAPE.OVAL, Cm(4.85 + (1.2 - dd) / 2), Cm(cy - dd / 2), Cm(dd), Cm(dd))
        dsq.fill.solid()
        dsq.fill.fore_color.rgb = RGBColor.from_string(NOIR)
        dsq.line.fill.background()
        if icones:
            ti = dd * 0.6
            slide.shapes.add_picture(io.BytesIO(icone_blanche(icones[i % len(icones)].image.blob)), Cm(4.85 + 0.6 - ti / 2), Cm(cy - ti / 2), Cm(ti), Cm(ti))
        for xx, l, t_, al in ((6.79, 22.98, intitule, PP_ALIGN.LEFT), (29.88, 1.53, str(num), PP_ALIGN.CENTER)):
            cadre = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(xx), Cm(cy - hb / 2), Cm(l), Cm(hb))
            cadre.fill.background()
            cadre.line.color.rgb = RGBColor.from_string(NOIR)
            cadre.line.width = Pt(1.5)
            tf = cadre.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = Cm(0.3)
            tf.margin_top = tf.margin_bottom = Cm(0)
            p = tf.paragraphs[0]
            p.alignment = al
            r = p.add_run()
            r.text = t_
            r.font.name = "Helvetica"
            r.font.size = Pt(taille)
            r.font.color.rgb = RGBColor.from_string(NOIR)


entrees = [("Résumé exécutif", 3)]
for i, x in enumerate(SL, 1):
    if x["type"] == "separateur" and x["num"] != "R":
        entrees.append((x["titre"], i))

slides = list(prs.slides)
for i, x in enumerate(SL, 1):
    sl = slides[i - 1]
    COURANTE = i
    if x["type"] == "separateur":
        for sh in sl.shapes:
            if sh.name in ("Rectangle 9", "Rectangle 5"):
                sh.width = Cm(10.6)
            if sh.name == "Rectangle 24" and len(str(x["num"])) > 1:
                sh.text_frame.word_wrap = False
                for p_ in sh.text_frame.paragraphs:
                    for r in p_.runs:
                        r.font.size = Pt(r.font.size.pt * 0.7 if r.font.size else 40)
    if x["type"] == "couverture":
        for sh in sl.shapes:
            if sh.name == "Google Shape;138;p25":
                sh.width = Cm(12.8)
            if sh.name == "Text Placeholder 4":
                sh.left, sh.width = Cm(21.5), Cm(11.0)
    for sh in sl.shapes:
        if getattr(sh, "has_chart", False) and sh.has_chart and x["type"] != "gab":
            try:
                if any(v is not None and v < 0 for pl in sh.chart.plots for se in pl.series for v in se.values):
                    sh.chart.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
            except Exception:
                pass
    if x["type"] == "sommaire":
        sommaire(sl, entrees)
    elif x["type"] == "gab":
        remplir_gab(sl, x["gab"], x["remplir"])
    elif x["type"] == "contenu":
        if "tableau" in x:
            tableau(sl, x["tableau"])
        if "pas_a_pas" in x:
            pas_a_pas(sl, x["pas_a_pas"])
        if "methode" in x:
            methode(sl, x["methode"], x.get("attente", False))
        if x.get("triangle"):
            triangle(sl)
        if "cartes" in x:
            cartes(sl, x["cartes"])
        if "schema" in x:
            schema(sl, x["schema"])
        if "image" in x:
            from PIL import Image
            im = Image.open(x["image"]["fichier"])
            larg = 19.0
            haut = larg * im.height / im.width
            if haut > 11.6:
                haut = 11.6
                larg = haut * im.width / im.height
            sl.shapes.add_picture(x["image"]["fichier"], Cm(1.4), Cm(4.7), Cm(larg), Cm(haut))
            txt(sl, 1.4, 3.85, 19.0, 0.7, x["image"]["legende"], 12, False, ORANGE, PP_ALIGN.CENTER)
    elif x["type"] == "annexe":
        txt(sl, 1.52, 0.4, 31.24, 1.1, x["titre"], 20, True, NOIR)
        tableau(sl, x["tableau"], 2.0)
        if x.get("source"):
            txt(sl, 1.02, 17.15, 31.75, 0.51, x["source"], 7.5, False, "808080")
    elif x["type"] in ("liste_figures", "liste_tableaux"):
        liste = FIG if x["type"] == "liste_figures" else TAB
        nom = "Figure" if x["type"] == "liste_figures" else "Tableau"
        txt(sl, 1.52, 0.4, 31.24, 1.1, "Liste des figures" if nom == "Figure" else "Liste des tableaux", 20, True, NOIR)
        moitie = (len(liste) + 1) // 2
        for col, part in enumerate((liste[:moitie], liste[moitie:])):
            if part:
                tableau(sl, {"titre": None, "entetes": [nom, "Titre", "Diapo."], "largeurs": [1.7, 12.0, 1.5],
                             "lignes": [[str(n_), t_, str(d_)] for n_, t_, d_ in part]}, 2.0, 1.4 + col * 16.0, 0.5 if len(part) > 14 else None, 9 if len(part) > 14 else 10)
    if x.get("notes"):
        sl.notes_slide.notes_text_frame.text = x["notes"]


def sans_point_virgule(tf):
    """Charte Adenka : pas de point-virgule dans les phrases ; on coupe en deux phrases."""
    for p_ in tf.paragraphs:
        for r in p_.runs:
            while " ; " in r.text:
                avant, apres = r.text.split(" ; ", 1)
                r.text = avant + ". " + apres[:1].upper() + apres[1:]


for sl in prs.slides:
    for sh in sl.shapes:
        if sh.has_text_frame:
            sans_point_virgule(sh.text_frame)
        if getattr(sh, "has_table", False) and sh.has_table:
            for row in sh.table.rows:
                for c in row.cells:
                    sans_point_virgule(c.text_frame)
prs.save(SORTIE)
finaliser(SORTIE)
os.remove(tmp)

# ---------------------------------------------------------------- contrôle automatique des chiffres
prs2 = Presentation(SORTIE)
textes = []
for sl in prs2.slides:
    morceaux = []
    for sh in sl.shapes:
        if sh.has_text_frame:
            morceaux.append(sh.text_frame.text)
        if getattr(sh, "has_table", False) and sh.has_table:
            for row in sh.table.rows:
                for c in row.cells:
                    morceaux.append(c.text)
    morceaux.append(sl.notes_slide.notes_text_frame.text if sl.has_notes_slide else "")
    textes.append("\n".join(morceaux))
for t_ in TRACE:
    t_["retrouve"] = t_["texte"] in textes[t_["diapositive"] - 1] if t_["diapositive"] <= len(textes) else False
json.dump({"figures": FIG, "tableaux": TAB, "diapositives": len(SL), "trace": TRACE},
          open(SORTIE.replace(".pptx", "_index.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
ok = sum(t_["retrouve"] for t_ in TRACE)
print(f"{len(SL)} diapositives ; {len(FIG)} figures ; {len(TAB)} tableaux ; chiffres retrouvés {ok}/{len(TRACE)}")
for t_ in TRACE:
    if not t_["retrouve"]:
        print("  NON RETROUVÉ", t_["diapositive"], t_["indicateur"], t_["texte"])
