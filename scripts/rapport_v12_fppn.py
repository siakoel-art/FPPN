# -*- coding: utf-8 -*-
"""
FPPN - Rapport actuariel et statistique, version 12 (PowerPoint, charte Adenka, Timesaver)
=========================================================================================

Entrées : resultats.json, complements.json, ratios.json (analyses), dossier des images de code, skill adenka-powerpoint.
Sorties : deck .pptx (graphiques et tableaux natifs, schémas modifiables, notes du présentateur) et
          fichier de traçabilité (chaque chiffre écrit : diapositive, valeur exacte, texte affiché).

Nouveautés de la version 10 (réponse aux quatre demandes, revue actuarielle complète, adossement actif-passif) :
- résumé exécutif ouvert par la réponse aux quatre demandes de la mission et un verdict chiffré par demande ;
- provisions : méthode courte maladie, allocations décès (états de fin d'année et 41591000), risques courts,
  intérêts d'épargne dus, dépréciation du Grenier selon sa politique, santé des retraités ;
- adossement actif-passif : méthodes justifiées, écoulement de l'épargne, écarts de liquidité, résistance, gouvernance ;
- projection agrégée de l'équilibre technique avec avertissement ; annexe des méthodes non retenues ;
- validité exacte par diapositive ; renvois numérotés entre diapositives ; mentions internes retirées.

Nouveautés de la version 9 (documents reçus le 22/10/2026, périmètre et validité des analyses) :
- périmètre, domaine et période de validité à trois niveaux : diapositive de synthèse (partie 2), ligne « Valable pour »
  sous chaque tableau et graphique, tableau détaillé en annexe ; colonne « Valable pour » dans le classeur ;
- contrôles croisés (partie 3) : états financiers contre balances, effectifs du FPPN contre fichiers 8Z6, rapport
  8Z6 / 936 par mois, rapprochements bancaires ;
- méthode de sinistralité du FPPN (communiquée le 22/10/2026) contre méthode réglementaire, sinistralité décès des deux
  sources, mortalité brute des membres (partie 5) ;
- taux servis sur l'épargne selon les documents, absence d'intérêts comptabilisés, sensibilité de l'EFC au taux,
  épargne présentée dans le rapport de gestion 2025 (partie 6) ; convention IPS-CGRAE (partie 8) ;
- provision du personnel recalculée agent par agent à partir des DISA, séparée entre Direction Générale, CMP et Grenier,
  avec sensibilités (partie 9) ; corrections (mesure du taux de 1,23 %, part de financement du Grenier, taux servi).

Nouveautés de la version 8 (analyses de sensibilité, statuts et règlement intérieur du 15/06/2023) :
- partie 10 : analyses de sensibilité et leviers de conformité de 2020 au 31/08/2026 (méthode et références, ratios
  structurels, cinq facteurs récurrents, montants à combler, effet de chaque levier, montants d'équilibre, conflits
  entre leviers, liquidité, réserve et fonds propres, faisabilité conjointe des normes de flux, inflation médicale
  ANStat, rendement comparé aux taux BCEAO) ; chiffres lus dans le classeur à formules recalculé (onglets 13a à 13i) ;
- partie 1 : dispositions des statuts et du règlement intérieur utiles à la mission ;
- barème de cotisation officiel (statuts, art. 17), barème de l'indemnité de départ (décret n° 2017-210, convention
  de 1977, art. 40), circulaire CIMA n° 00230/2005, PBC complémentaire « textes et personnel » (items 100 à 109).

Nouveautés de la version 7 (réponse aux quatre tâches actuarielles de la mission) :
- résumé exécutif : réponse aux quatre questions de la mission (établi, non établi, sens d'évolution, PBC) ;
- conclusions par tâche en partie 9 : engagements, suffisance des provisions, adossement, adéquation des placements ;
- adossement présenté comme une borne favorable, règles de valorisation des actifs tirées des textes (art. 17, 19,
  20, 21 et 27 du règlement d'exécution n° 004/2022 ; Décision n° 05/2022, section II ; PCMS), quatre scénarios
  de couverture et sorties d'épargne observées.

Nouveautés de la version 6 (entretiens du 19/10/2026 avec M. Koffi, chef comptable) :
- placements (DAT annuel, FCP, décision de placement en logigramme), méthode du FPPN pour les allocations décès
  (logigramme, question du compte utilisé), GNA apporteur d'affaires, compte 10620000, taux servi de 3,5 % ;
- retouches de l'utilisateur sur la version 5 reprises : couverture (titre, logo du FPPN, date), paragraphe
  d'implication sans libellé, positions des formes (assets/corrections_mise_en_page_v5.json).

Nouveautés de la version 5 :
- revue des provisions produit par produit (règlement d'exécution n° 004/2022, art. 7 à 11) : PSAP et dettes envers
  les cliniques, méthode CIMA des sinistres tardifs, dépréciation des prêts du Grenier (barème BCEAO des SFD),
  flotte automobile (risque porté par GNA), départ à la retraite du personnel (compte 196) ;
- « Choix de la formule » : formule du texte, méthode du FPPN (en rouge si non communiquée) et écart, pour la
  sinistralité et les frais de gestion, de 2020 au 31/08/2026 ;
- tableau des formules en annexe (registre formules_fppn.py), complet dans l'onglet 11_formules du classeur.

Nouveautés de la version 4 (entretien du 19/10/2026 avec le responsable des cotisations) :
- chaînes de traitement et circuits dessinés en logigrammes (formes PowerPoint modifiables : début, étape,
  décision, document, alerte, flèches attachées) : chaîne de traitement Adenka, circuit des précomptes en
  couloirs, parcours des nouvelles recrues ;
- circuit des fichiers par produit et règles des produits confrontées aux fichiers de précomptes ;
- déclarations de l'entretien signalées « déclaration du responsable des cotisations, entretien du 19/10/2026,
  à confirmer par document ».

Nouveautés de la version 3 :
- partie « Cadre réglementaire et conformité prudentielle » (Décision n° 05/2022/COM/UEMOA, règlement
  d'exécution n° 004/2022/COM/UEMOA, règlement n° 07/2009/CM/UEMOA, PCMS, arrêté n° 00115/MEPS/CAB) ;
- sigles de chaque diapositive définis en bas de page ;
- analyse en trois temps sous chaque tableau, graphique ou illustration : constat en gras, preuves et
  explication, puis un paragraphe d'implication pour le FPPN et la mission ;
- gabarits Timesaver pour les méthodes, la chaîne de traitement et les diapositives « À retenir » ;
- extraits de code mis en forme comme dans VS Code sous Windows.

Règles de présentation : pourcentages à deux décimales ; montants exacts en FCFA, sans décimale, dans
les textes et les tableaux ; graphiques : données exactes en FCFA, affichées en millions.
Utilisation : python rapport_v12_fppn.py <resultats.json> <complements.json> <ratios.json> <dossier_images> <skill> <sortie.pptx> <dossier_base>
              <classeur_sensibilites_recalcule.xlsx> <carte_sensibilites.json> <donnees_v9.json> <donnees_v10.json> <donnees_v11.json> <correspondance_pbc.json>
"""
import io
import json
import math
import os
import re
import sys

RES, COMP, RAT, IMG, SKILL, SORTIE, BASE, SENS, CARTE8, V9J, V10J, V11J, PBCJ = sys.argv[1:14]
sys.path.insert(0, os.path.join(SKILL, "scripts"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl as _ox  # noqa: E402
import pandas as pd  # noqa: E402
import logigramme_fppn as LG  # noqa: E402
import formules_fppn as FORM  # noqa: E402
from pptx_tools import selection, remplir, finaliser  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.chart.data import CategoryChartData  # noqa: E402
from pptx.dml.color import RGBColor  # noqa: E402
from pptx.enum.chart import XL_TICK_LABEL_POSITION  # noqa: E402
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE  # noqa: E402
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402
from pptx.util import Cm, Emu, Pt  # noqa: E402

R = json.load(open(RES, encoding="utf-8"))
C = json.load(open(COMP, encoding="utf-8"))
Q = json.load(open(RAT, encoding="utf-8"))
PV = json.load(open(os.path.join(os.path.dirname(RES), "provisions.json"), encoding="utf-8"))
V9 = json.load(open(V9J, encoding="utf-8"))
V11 = json.load(open(V11J, encoding="utf-8"))
PBCM = {int(k): v for k, v in json.load(open(PBCJ, encoding="utf-8"))["mapping"].items()}
# v11 : encours des placements lu dans les grands livres (moyenne pondérée par les jours), à la place de la moyenne des soldes de clôture
R["rendement_placements"] = V11["rendement_gl"]
_p9 = pd.read_parquet(os.path.join(BASE, "precomptes.parquet"))
GNA_P = _p9[_p9.retenu & (_p9.sens == "retour") & (_p9.poste == "9W3")].groupby("annee").montant.sum().to_dict()
del _p9
FLOW = "« FLOW_CHART_COTISATION_ET_PRODUITS_DEPARGNE_1.pptx » (reçu le 23/10/2026)"
SRC_ENT14 = "Source : entretien du 19/10/2026 (14h) avec M. Koffi, chef comptable"
SRC_ENT16 = "Source : entretien du 19/10/2026 (16h) avec M. Koffi, chef comptable"

_bal = pd.read_parquet(os.path.join(BASE, "compta_balances.parquet"))
_bal = _bal[_bal.nature == "comptes"].sort_values("periode_au").groupby(["exercice", "compte"]).tail(1)


def _mv(ex, motif, col):
    g = _bal[(_bal.exercice == ex) & _bal.compte.str.match(motif)]
    return float(g[col].sum())


def _so(ex, motif):
    g = _bal[(_bal.exercice == ex) & _bal.compte.str.match(motif)]
    return float((g.sol_d - g.sol_c).sum())


TIMESAVER = os.path.join(SKILL, "assets", "Timesaver_Adenka.pptx")
NOIR, GRIS, BLEU, ORANGE, ROUGE, GRIS_F, GRIS_C = "060606", "F2F2F2", "112A40", "EFAC62", "C00000", "53565A", "BBBCBC"
BAS_CONTENU = 16.3          # aucun contenu sous cette ordonnée (cm) : place des sigles et de la source

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
MORT = {m["table"]: m for m in C["mortalite"]}
IC = C["ic_grenier"]
SC_ = C["sensibilite_cle"]
INF = {int(k): v for k, v in C["inflation"]["valeurs"].items()}
RT = {x["code"]: x for x in Q["ratios"]}
for x in Q["ratios"]:
    for cle in ("valeurs", "valeurs_retraitees", "statuts", "statuts_retraites", "num", "den"):
        if cle in x:
            x[cle] = {(int(k) if str(k).isdigit() else k): v for k, v in x[cle].items()}
AG = {int(k): v for k, v in Q["agregats"].items()}
# v12 : le taux brut de rentabilité des placements (RL6) suit l'encours lu dans les grands livres, comme la partie 9
RT["RL6"]["valeurs"] = {x["exercice"]: x["rendement"] for x in V11["rendement_gl"]}
RT["RL6"]["formule"] = "Produits financiers / encours moyen des DAT et du FCP, pondéré par les jours (grands livres)"
MS = {m["exercice"]: m for m in Q["marge_solvabilite"]}
AFF = {a["exercice_affectation"]: a for a in Q["affectation"]}
EXR = {e["exercice"]: e for e in Q["excedent_par_risque"]}
MOB = Q["mobilisation"]
cats_ex = [str(a) if a < 2026 else "2026 (8 mois)" for a in range(2020, 2027)]
yrs_bil = [str(a) if a < 2026 else "31/08/2026" for a in range(2020, 2027)]
FMT_M = "# ##0,,"           # données exactes en FCFA, affichage en millions

SRC_PREC = "Source : fichiers de précomptes de la Solde, de la CGRAE et du FPPN reçus le 15/10/2026 (dossiers « FPPN-SOLDE », « FPPN-CGRAE », « FPPN-GNA », noms exacts en annexe « Fichiers sources ») ; traitements Adenka"
SRC_COMPTA = "Source : balances des comptes du FPPN de 2020 au 31/08/2026 (provisoire), fichiers « BALANCE … » listés en annexe « Fichiers sources » ; traitements Adenka"
SRC_GRENIER = "Source : exports du logiciel de prêts du Grenier du Policier « BalanceAgeeDetaille … .xls » et « production_de_la_periode … .xls » reçus le 15/10/2026 ; traitements Adenka"
SRC_CIMA = "Source : tables de mortalité CIMA H, CIMA F, TD CIMA et TV CIMA (fichiers fournis le 18/10/2026) ; table ronde CIMA FANAF 2011 ; calculs Adenka"
SRC_BM = f"Source : Banque mondiale, FP.CPI.TOTL.ZG, mise à jour du {C['inflation']['mise_a_jour'][8:10]}/{C['inflation']['mise_a_jour'][5:7]}/{C['inflation']['mise_a_jour'][:4]} ; " + SRC_COMPTA[9:]
DEC = "Décision n° 05/2022/COM/UEMOA du 1er juin 2022, annexe"
REX = "règlement d'exécution n° 004/2022/COM/UEMOA du 1er juin 2022"
R07 = "règlement n° 07/2009/CM/UEMOA du 26 juin 2009"
SRC_REG = "Source : " + DEC + " ; " + REX + " ; " + R07 + " ; balances du FPPN ; calculs Adenka"
ENT19 = "déclaration du responsable des cotisations, entretien du 19/10/2026"
SRC_ENT19 = "Source : entretien du 19/10/2026 avec le responsable des cotisations"

# Règles des produits : confrontation des déclarations du 19/10/2026 aux fichiers retour (janvier 2020 à août 2026)
_pre = pd.read_parquet(os.path.join(BASE, "precomptes.parquet"))
_pre = _pre[_pre.retenu & (_pre.sens == "retour") & (_pre.annee >= 2020) & ((_pre.annee < 2026) | (_pre.mois <= 8))]
_efc = _pre[_pre.poste == "870"]
EFC10 = {a: (g.montant == 10000).mean() for a, g in _efc.groupby("annee")}
EFC6 = {a: (g.montant == 6000).mean() for a, g in _efc.groupby("annee")}
_pep = _pre[_pre.poste == "998"]
PEP_M5, PEP_LT10, PEP_MAX = (_pep.montant % 5000 == 0).mean(), (_pep.montant < 10000).mean(), _pep.montant.max()
BAREME = "barème fixé par l'article 17 des statuts du FPPN"

# ---------------------------------------------------------------------------
# Sigles : définis en bas de chaque diapositive où ils apparaissent
# ---------------------------------------------------------------------------
SIGLES = [
    ("FPPN", r"\bFPPN\b", "Fonds de prévoyance de la Police nationale"),
    ("AIRMS", r"\bAIRMS\b", "Agence ivoirienne de régulation de la mutualité sociale"),
    ("UEMOA", r"\bUEMOA\b", "Union économique et monétaire ouest-africaine"),
    ("PCMS", r"\bPCMS\b", "plan comptable des mutuelles sociales"),
    ("CIMA", r"\bCIMA\b", "Conférence interafricaine des marchés d'assurances"),
    ("CGRAE", r"\bCGRAE\b", "Caisse générale de retraite des agents de l'État"),
    ("BCEAO", r"\bBCEAO\b", "Banque centrale des États de l'Afrique de l'Ouest"),
    ("PEP", r"\bPEP\b", "Plan Épargne Plus"),
    ("EFC", r"\bEFC\b", "entraide fin de carrière"),
    ("CMP", r"\bCMP\b", "Clinique médicale du policier"),
    ("FCP", r"\bFCP\b", "fonds commun de placement"),
    ("DAT", r"\bDAT\b", "dépôt à terme"),
    ("DAF", r"\bDAF\b", "directeur administratif et financier"),
    ("CNPS", r"\bCNPS\b", "Caisse nationale de prévoyance sociale"),
    ("IS", r"\bIS\b", "impôt sur les salaires (part patronale)"),
    ("FDFP", r"\bFDFP\b", "Fonds de développement de la formation professionnelle"),
    ("SGI", r"\bSGI\b", "société de gestion et d'intermédiation"),
    ("S/C", r"S/C", "prestations / cotisations"),
    ("TCAM", r"\bTCAM\b", "taux de croissance annuel moyen"),
    ("IC", r"\bIC\b", "intervalle de confiance"),
    ("PAR30", r"\bPAR ?30\b", "part de l'encours en retard de plus de 30 jours"),
    ("PSAP", r"\bPSAP\b", "provision pour sinistres à payer"),
    ("PIDR", r"\bPIDR\b", "provision pour indemnités de départ à la retraite"),
    ("SFD", r"\bSFD\b", "systèmes financiers décentralisés"),
    ("SYSCOHADA", r"SYSCOHADA", "système comptable de l'OHADA"),
    ("IBNR", r"\bIBNR\b", "sinistres survenus non encore déclarés"),
    ("A/E", r"\bA/E\b", "décès observés / décès attendus"),
    ("PBC", r"\bPBC\b", "liste des documents demandés au client"),
    ("FNG", r"\bFNG\b", "fonds national de garantie"),
    ("FOMUS-CI", r"FOMUS-CI", "fonds de garantie des mutuelles sociales de Côte d'Ivoire"),
    ("MEPS", r"\bMEPS\b", "ministère de l'Emploi et de la Protection sociale"),
    ("GNA", r"\bGNA\b", "assureur de la flotte automobile (compte 47111106)"),
    ("MEDOC", r"\bMEDOC\b", "produit santé complémentaire du poste 8Z7"),
    ("TAFIRE", r"\bTAFIRE\b", "tableau financier des ressources et des emplois"),
    ("CAC", r"\bCAC\b", "commissaire aux comptes"),
    ("HAO", r"\bHAO\b", "hors activités ordinaires"),
    ("CMU", r"\bCMU\b", "couverture maladie universelle"),
    ("CSV", r"\bCSV\b", "fichier texte de tableau"),
    ("OCR", r"\bOCR\b", "reconnaissance optique de caractères"),
    ("FANAF", r"\bFANAF\b", "Fédération des sociétés d'assurances de droit national africaines"),
    ("CPM", r"\bCPM\b", "Comité de politique monétaire"),
    ("SCI", r"\bSCI\b", "société civile immobilière"),
    ("AG", r"\bAG\b", "assemblée générale"),
    ("R²", r"R²", "coefficient de détermination"),
    ("qx", r"\bqx\b", "probabilité de décès entre x et x + 1 ans"),
    ("Lx", r"\bLx\b", "survivants à l'âge x"),
    ("e_x", r"\be_?x\b|\be5[05]\b|\be60\b|\be0\b", "espérance de vie à l'âge x"),
    ("ä_x", r"ä_?x|ä5[05]|ä6[05]", "valeur d'une rente viagère de 1 par an à l'âge x"),
    ("TD, TV", r"\bT[DV] CIMA\b|\(T[DV]\)", "anciennes tables CIMA de décès et de survie"),
    ("936", r"(?:poste |postes |maladie |Maladie |\(|^)936\b", "poste maladie"),
    ("8Z6", r"\b8Z6\b", "poste décès"),
    ("870", r"(?:poste |EFC |\(|^)870\b", "poste EFC"),
    ("998", r"(?:poste |PEP |\(|^)998\b", "poste PEP"),
    ("9X4", r"\b9X4\b", "poste Grenier"),
    ("9W3", r"\b9W3\b", "poste flotte automobile"),
    ("8Z7", r"\b8Z7\b", "poste MEDOC"),
    ("ANStat", r"\bANStat\b", "Agence nationale de la statistique de Côte d'Ivoire"),
    ("IHPC", r"\bIHPC\b", "indice harmonisé des prix à la consommation"),
    ("UMOA", r"\bUMOA\b", "Union monétaire ouest-africaine"),
    ("AMF-UMOA", r"AMF-UMOA", "Autorité des marchés financiers de l'UMOA"),
    ("ISAP", r"\bISAP\b", "norme internationale de pratique actuarielle de l'Association Actuarielle Internationale"),
    ("IAS, IFRS", r"\bIAS\b|\bIFRS\b", "normes comptables internationales"),
    ("SMIG", r"\bSMIG\b", "salaire minimum interprofessionnel garanti"),
    ("DISA", r"\bDISA\b", "déclaration individuelle des salaires annuels à la CNPS"),
    ("IPS-CGRAE", r"IPS-CGRAE", "Institution de prévoyance sociale - Caisse générale de retraite des agents de l'État"),
    ("DG", r"\bDG\b", "Direction Générale du FPPN"),
    ("DEPA", r"\bDEPA\b", "Direction de l'École de Police d'Abidjan"),
    ("S/P", r"\bS/P\b", "sinistres sur primes (cotisations)"),
]


def sigles_du_texte(texte):
    trouves = []
    for nom, motif, defi in SIGLES:
        if re.search(motif, texte, flags=re.M):
            trouves.append(f"{nom} : {defi}")
    return trouves


PARTIE = None
VALIDITE = {  # niveau 2 : ligne « Valable pour » par défaut de chaque partie (remplacée par une mention propre si nécessaire)
    "R": "chaque chiffre dans le périmètre de sa partie d'origine, voir l'annexe « Périmètre et validité des analyses »",
    1: "textes du FPPN (statuts et règlement intérieur du 15/06/2023) et circuits décrits en entretien, sans test de leur application",
    2: "mission au 15/10/2026, documents reçus jusqu'au 22/10/2026",
    3: "fichiers de précomptes reçus le 15/10/2026, mois effectivement reçus pour chaque poste (diapositive « Couverture de la période »)",
    4: "policiers en activité précomptés par la Solde, janvier 2024 à septembre 2026 ; taux mesurés sur les mois détaillés cités",
    5: "comptes 2020 à 2025 et du 01/01 au 31/08/2026 (provisoire) ; prestations comptabilisées, sans sinistres tardifs",
    6: "soldes et mouvements comptables de 2020 au 31/08/2026 ; épargne sans échéancier individuel ni intérêts comptabilisés",
    7: "exports du logiciel de prêts du Grenier, arrêtés du 31/12/2021 au 30/09/2026, non rapprochés des contrats",
    8: "retraités précomptés par la CGRAE de février 2024 à septembre 2026 ; produits annexes sur les mois reçus",
    9: "balances de 2020 au 31/08/2026, valeurs comptables non réévaluées ; 2026 provisoire (huit mois)",
    10: "ratios de la Décision n° 05/2022 recalculés sur les balances ; statut opposable de 2022 à 2025, indicatif en 2020, 2021 et 2026",
    11: "tables réglementaires CIMA ; non comparées à la mortalité par âge du FPPN (dates de naissance non reçues)",
    "9.2": "bilans au 31/12/2025 et au 31/08/2026 (provisoire) ; sorties d'épargne observées de 2021 à 2025",
    "9.3": "balances de 2020 au 31/08/2026 ; documents sur les actifs reçus le 22/10/2026",
    12: "comptes 2023 à 2025, garanties maladie et décès prises ensemble",
    13: "traitements des fichiers reçus jusqu'au 22/10/2026",
    14: "demandes et constats au 22/10/2026",
}


def contenu(entete, notes, **k):
    if any(c in k for c in ("tableau", "graphique", "graphiques")):
        k.setdefault("valable", VALIDITE.get(PARTIE))
    k["_partie"] = PARTIE
    SL.append({"type": "contenu", "entete": entete, "notes": notes, **k})


def separateur(num, titre, sous_titre, notes=""):
    global PARTIE
    PARTIE = num
    SL.append({"type": "separateur", "num": num, "titre": titre, "sous_titre": sous_titre, "notes": notes})


def gab(num, entete, remplissage, notes, source=None, analyse=None):
    SL.append({"type": "gab", "gab": num, "entete": entete, "remplir": remplissage, "notes": notes, "source": source,
               "analyse": analyse})


def retenir(theme, constat, preuves, implication, action, notes):
    """Diapositive « À retenir » sur le gabarit Timesaver n° 13 (4 cartes) : constat, preuves, implication, action."""
    gab(13, {"titre": f"À retenir : {theme}", "sous_titre": constat[0], "synthese": constat[1]},
        {"cartes": [("Constat", constat[2]), ("Preuves", preuves), ("Implication", implication), ("Action", action)]},
        notes, source="Source : analyses Adenka")


def methode_gab(entete, items, notes, source="Source : Adenka", attente=False):
    """Méthode sur le gabarit Timesaver n° 9 (arborescence à 4 blocs) : (titre, texte) x 4."""
    gab(9, entete, {"blocs": items, "attente": attente}, notes, source=source)

# ===========================================================================
# CONTENU
# ===========================================================================
# ---- v7 : données des conclusions par tâche (adossement, engagements, provisions, placements)
mo = {m["cloture"]: m for m in MOB}
ENG = {k: mo[k]["passif_moins_1_an"] + mo[k]["epargne_echeance_a_etablir"] for k in mo}
EXA = {k: {"hop": _so(ex, r"^47111100$"), "cmpi": _so(ex, r"^47111104$"), "cmpc": _so(ex, r"^47111105$"), "terr": _so(ex, r"^22200000$")}
       for k, ex in (("31/12/2025", 2025), ("31/08/2026", 2026))}
SCEN = {}
for k in mo:
    a = mo[k]
    s2 = a["total_actif"] - EXA[k]["hop"] - EXA[k]["cmpc"]
    SCEN[k] = (a["total_actif"], s2, s2 - EXA[k]["cmpi"] - EXA[k]["terr"], a["immediat"] + a["moins_1_an"])
SORT = [(ex, _mv(ex, r"^41195100$", "mvt_d"), _mv(ex, r"^41195200$", "mvt_d"), _mv(ex, r"^41195100$", "mvt_c"), _mv(ex, r"^41195200$", "mvt_c")) for ex in range(2020, 2027)]
EP26 = {"pep": -_so(2026, r"^41195100$"), "efc": -_so(2026, r"^41195200$")}
PRc = {r["exercice"]: r for r in PV["psap_rapprochement"]}
RDM = sum(r["produits_financiers"] for r in R["rendement_placements"]) / sum(r["placements_moyens"] for r in R["rendement_placements"])
SRC_TXT = "Source : " + REX + ", art. 17, 19, 20, 21 et 27 ; " + DEC + ", section II, § 2.1 ; PCMS (règlement n° 03/2012/CM/UEMOA)"

NB_NC = sum(1 for x in Q["ratios"] if x["statut_global"] == "Non conforme")
NB_C = sum(1 for x in Q["ratios"] if x["statut_global"] == "Conforme")
NB_ATT = sum(1 for x in Q["ratios"] if x["statut_global"] in ("En attente", "À apprécier"))
sig = lambda c, a=2025: RT[c]["valeurs"].get(a)  # noqa: E731

def _hv(x, largeur=31.2):
    """Hauteur (cm) réservée à la ligne « Valable pour » : une ligne par tranche de largeur, 8,5 points."""
    if not x.get("valable"):
        return 0
    return 0.40 * math.ceil(len("Valable pour : " + x["valable"] + ".") * 0.158 / largeur) + 0.15


def _snap():
    return (len(TRACE), len(FIG), len(TAB))


def _omis(n0, *a, **k):
    """Diapositive retirée en v10 : on efface les chiffres, figures et tableaux enregistrés pendant l'évaluation de ses arguments."""
    del TRACE[n0[0]:], FIG[n0[1]:], TAB[n0[2]:]


# ---- v10 : données de la revue actuarielle complète et renvois entre diapositives
V10 = json.load(open(V10J, encoding="utf-8"))


def ref(titre):
    """Renvoi vers une diapositive, résolu en numéro une fois toutes les diapositives créées."""
    return "§§" + titre + "§§"


_E10 = {e["exercice"]: e for e in V10["epargne"]}


def interets_dus(t_pep, t_efc, jusqu_a=2026):
    """Intérêts dus non comptabilisés du 01/01/2020 à la fin de l'exercice indiqué (non capitalisés, encours moyen)."""
    tot = 0.0
    for a in range(2020, jusqu_a + 1):
        e = _E10[a]
        tot += ((e["pep_debut"] + e["pep_fin"]) / 2 * t_pep + (e["efc_debut"] + e["efc_fin"]) / 2 * t_efc) * e["mois"] / 12
    return tot


EP_DUE26 = _E10[2026]["pep_fin"] + _E10[2026]["efc_fin"]
EP_DUE25 = _E10[2025]["pep_fin"] + _E10[2025]["efc_fin"]
INT55 = interets_dus(0.035, 0.055)
INT35 = interets_dus(0.035, 0.035)
INT55_25 = interets_dus(0.035, 0.055, 2025)
DST = {x["exercice"]: x for x in V10["deces_stock"]}
DETTE_DEC26 = (DST[2026]["ouv_41591000"] + DST[2026]["mvt_c_41591000"] - DST[2026]["mvt_d_41591000"]) + (DST[2026]["ouv_41592000"] + DST[2026]["mvt_c_41592000"] - DST[2026]["mvt_d_41592000"])
PROV_OBS = PV["psap_deces"]["2026"]
PROV_MAL = PV["psap_maladie"]["2026"]
TARD_DEC = sum(x["montant"] for x in V10["deces_tardifs"])
_PM = {x["exercice"]: x for x in V10["psap_maladie_delais"]}
TARDM = {x["cloture"]: x for x in V10["psap_maladie_tardives"]}


def psap_m_elargie(cl):
    x, pa = TARDM[cl], _PM[cl + 1]
    return x["references_annee_precedente"] / (x["references_identifiees"] / x["factures_sinistres"]) * (pa["prestations_maladie"] / pa["factures_enregistrees"])


DELAI = {a: _PM[a]["dettes_cliniques"] / (_PM[a]["factures_enregistrees"] / _PM[a]["mois"]) for a in range(2020, 2027)}
GRD = {}
for _g in V10["grenier_depreciation"]:
    GRD[_g["arrete"]] = GRD.get(_g["arrete"], 0.0) + _g["encours"] * _g["taux"]
SORT10 = V10["epargne_sorties"]
T_MOY = sum(x["total_taux"] for x in SORT10) / len(SORT10)
T_MAX = max(x["total_taux"] for x in SORT10)
MOB10 = {m["cloture"]: m for m in V10["mobilisation"]}


def couverture(cl, taux, choc=0.0):
    m = MOB10[cl]
    actif = m["immediat"] + m["moins_1_an"]
    besoin = m["passif_moins_1_an"] + m["epargne_echeance_a_etablir"] * taux + choc
    return actif, besoin, actif / besoin


PIC = V10["pic_prestations"]
CHOC_MAL = s[2025]["prest_maladie"] * PIC
# écarts de liquidité année par année au 31/08/2026 (écoulement statique, sans nouveaux versements)
_m = MOB10["31/08/2026"]
_E0 = _m["epargne_echeance_a_etablir"]
GAP = []
_cum = 0.0
for k in range(1, 6):
    act = (_m["immediat"] + _m["moins_1_an"]) if k == 1 else _m["1_a_5_ans"] / 4
    sor = _E0 * (1 - T_MOY) ** (k - 1) * T_MOY
    pas_ = (_m["passif_moins_1_an"] if k == 1 else 0.0) + sor
    _cum += act - pas_
    GAP.append((f"Année {k}", act, _m["passif_moins_1_an"] if k == 1 else 0.0, sor, act - pas_, _cum))
_res_ep = _E0 * (1 - T_MOY) ** 5
GAP.append(("Au-delà de 5 ans", _m["plus_5_ans_ou_incertain"], 0.0, _res_ep, _m["plus_5_ans_ou_incertain"] - _res_ep, _cum + _m["plus_5_ans_ou_incertain"] - _res_ep))
PDV = PV["pidr"]
PIDR25 = V9["pidr_synthese"][-1]
SRC_LOT4 = "« PROV ALLOCATION DECES 2021.xlsx », « PROV ALLOCATION DECES 2022.xlsx », « ALLOC DECES 2023 PAYER ET A PROVISIONNER .xlsx », « PROV ALLOCATION DECES 2024.xlsx », « PROV ALLOCATIONS DECES ET FRAIS FUNERAIRES 2025.xlsx »"
SRC_LOT4C = "états de fin d'année des allocations décès 2021 à 2025 (« PROV ALLOCATION DECES 2021.xlsx » et suivants, voir l'annexe « Fichiers sources »)"
SRC_GLT = "grands livres des tiers 2023 à 2025 (« GRAND LIVRE DES TIERS FPPN 2023.xlsx » et suivants)"
SRC_RI = "règlement intérieur du FPPN édité le 15/06/2023"
SRC_ST = "statuts du FPPN édités le 15/06/2023"
RDN = V10["retraites_naissance"]

SL.append({"type": "couverture", "notes": "Rapport d'étape (réponse aux quatre demandes, revue actuarielle complète, adossement actif-passif ; documents du 22/10/2026, périmètre et validité, provision du personnel recalculée, contrôles croisés ; analyses de sensibilité et leviers de conformité ; statuts et règlement intérieur ; réponse aux quatre tâches actuarielles ; entretiens placements et provisions du 19/10/2026 ; revue des provisions, choix des formules et tableau des formules ; entretien cotisations du 19/10/2026 intégré). Données arrêtées à l'envoi du FPPN reçu le 15/10/2026, complétées le 18/10/2026 (grands livres PDF, inflation Banque mondiale, tables CIMA, documentation publique de l'AIRMS)."})
SL.append({"type": "sommaire", "notes": "Le rapport suit l'ordre : résumé exécutif (réponse aux quatre demandes), compréhension du FPPN, démarche, données, analyses par thème, adossement actif-passif et placements, conformité prudentielle, tables de mortalité, projection de l'équilibre technique, outils, analyses en attente, annexes."})

# ---------------------------------------------------------------- RÉSUMÉ EXÉCUTIF
separateur("R", "Résumé exécutif", "Constats, chiffres clés, risques et actions",
           "Cette partie peut être présentée seule : elle reprend les messages essentiels du rapport, avec les chiffres et les renvois aux parties détaillées.")
# ---- v10 : réponse aux quatre demandes de la mission (premières diapositives du résumé exécutif)
SRC_DEM = "Source : demande de M. Wabi (associé responsable) : engagements, suffisance des provisions, adossement actif-passif, adéquation des placements, de 2020 au 31/08/2026 ; analyses détaillées citées dans la colonne « Détail »"
V_ENG, V_PROV, V_ADO, V_PLA = "Partiellement démontré", "Partiellement démontré", "Insuffisant", "Insuffisant"
contenu({"titre": "Réponse aux quatre demandes de la mission", "sous_titre": "Engagements et provisions partiellement démontrés ; adossement et placements insuffisants",
         "synthese": "Chaque verdict est détaillé dans la diapositive suivante, avec tous les chiffres qui le fondent et le renvoi à leur démonstration"},
        "Échelle des verdicts : « Démontré » (les documents reçus permettent de conclure et la conclusion est favorable), « Partiellement démontré » (une partie est mesurée, une partie ne l'est pas), « Insuffisant » (les mesures montrent un manque), « Non démontrable » (les documents reçus ne permettent pas de conclure).",
        tableau={"titre": tab("Verdict sur chacune des quatre demandes"), "entetes": ["Demande", "Verdict", "En une phrase", "Détail"], "largeurs": [6.6, 4.4, 15.4, 4.8], "taille": 9, "statut_col": 1,
                 "lignes": [["1. Engagements envers les adhérents et ayants droit", V_ENG, "L'épargne, les allocations décès, la maladie et la provision du personnel sont mesurés ; les intérêts dus sur l'épargne ne sont pas comptabilisés ; la santé des retraités n'est pas mesurable.", ref("Demande 1 : engagements envers les adhérents")],
                            ["2. Suffisance des provisions au regard des prestations servies", V_PROV, "Les provisions maladie et décès dépassent largement les besoins mesurés mais ne résultent d'aucun calcul ; les intérêts d'épargne et la perte probable sur les prêts ne sont pas provisionnés.", ref("Demande 2 : suffisance des provisions")],
                            ["3. Adossement actif-passif", V_ADO, "L'actif mobilisable dans l'année couvre environ un tiers de ce que le FPPN devra payer dans l'année, épargne comprise.", ref("Demande 3 : adossement actif-passif")],
                            ["4. Adéquation des placements aux engagements", V_PLA, "Les placements financiers sont trop petits, peu rémunérateurs et sans politique ; ils ne couvrent ni les intérêts dus ni les sorties d'épargne.", ref("Demande 4 : adéquation des placements")]]},
        disposition="dessous",
        analyse=["**Le FPPN connaît ses dettes de court terme, mais pas la valeur de ses engagements d'épargne, faute de taux arrêté, d'échéancier et d'intérêts comptabilisés.** C'est la cause commune des verdicts 1, 3 et 4."],
        implication="Les quatre diapositives suivantes donnent, pour chaque demande, tous les chiffres qui fondent le verdict, ses limites et le renvoi vers la démonstration.",
        valable="période du 01/01/2020 au 31/08/2026 (mission à partir du 01/09/2020), documents du FPPN reçus jusqu'au 22/10/2026",
        source=SRC_DEM)
contenu({"titre": "Demande 1 : engagements envers les adhérents", "sous_titre": f"Verdict : {V_ENG}. Épargne due au 31/08/2026 de {F('V1 épargne 2026', EP_DUE26)}, hors intérêts dus",
         "synthese": f"Avec les intérêts dus depuis le 01/01/2020 (PEP 3,5 %, EFC 5,5 %), l'engagement d'épargne atteint {F('V1 engagement réel', EP_DUE26 + INT55)}"},
        "Un engagement est ce que le FPPN devra payer plus tard au titre de droits déjà acquis par les adhérents et leurs ayants droit. Le tableau reprend chaque engagement, sa valeur et la diapositive qui la démontre.",
        tableau={"titre": tab("Engagements envers les adhérents : chiffres qui fondent le verdict"), "entetes": ["Engagement", "Valeur", "Ce que le chiffre montre", "Détail"], "largeurs": [7.6, 5.6, 13.2, 4.8], "taille": 8,
                 "lignes": [["Épargne PEP et EFC due au 31/08/2026 (comptes 41195100 et 41195200)", F("V1 épargne (t)", EP_DUE26), "Somme des cotisations d'épargne nettes des sorties, sans intérêts", ref("Épargne : intérêts dus non comptabilisés")],
                            ["Intérêts dus du 01/01/2020 au 31/08/2026, PEP 3,5 % et EFC 5,5 %", F("V1 intérêts 5,5", INT55), "Intérêts promis par le manuel d'épargne, absents des comptes", ref("Épargne : intérêts dus non comptabilisés")],
                            ["Mêmes intérêts si l'EFC est servie à 3,5 %", F("V1 intérêts 3,5", INT35), "Borne basse : taux indiqué par le chef comptable", ref("Épargne : intérêts dus non comptabilisés")],
                            ["Taux servi sur l'EFC selon les documents", "3,5 %, 5,5 % ou 14 %", "Aucun taux arrêté par une décision signée", ref("EFC et PEP : trois taux dans trois documents")],
                            ["Allocations décès et frais funéraires dus au 31/08/2026 (41591000 et 41592000)", F("V1 dette décès", DETTE_DEC26), "Liste du FPPN des décès non payés, valorisés au barème", ref("Allocations décès : dette réelle et provision obsèques")],
                            ["Décès connus après la clôture (états 2021 et 2023)", F("V1 décès tardifs", TARD_DEC), "Ordre de grandeur de ce qui manque chaque 31 décembre", ref("Allocations décès : décès connus après la clôture")],
                            ["Soins facturés l'année suivante (clôture 2024, estimation élargie)", F("V1 maladie tardive", psap_m_elargie(2024)), "Dette maladie non encore facturée à la clôture", ref("Provision maladie : montant nécessaire")],
                            ["Provision du personnel recalculée au 31/12/2025", F("V1 PIDR", PIDR25["dbo"]), f"Compte 196 du FPPN : {F('V1 PIDR 196', PIDR25['provision_196'])}", ref("Provision du personnel par entité de 2021 à 2025")],
                            ["Santé des policiers retraités", N("V1 retraités", CG[-1]["retraites"]) + " retraités cotisants (" + CG[-1]["libelle"] + ")", f"Non mesurable : soins des retraités non isolés ; date de naissance connue pour {N('V1 ret dn', RDN['nb'])} retraités seulement", ref("Santé des retraités : ce que l'on peut mesurer")]]},
        disposition="dessous",
        analyse=[f"**Le premier engagement est l'épargne : {F('V1 épargne (a)', EP_DUE26)} au 31/08/2026, plus {F('V1 intérêts (a)', INT55)} d'intérêts dus non comptabilisés au taux du manuel.** Les autres engagements mesurés sont des dettes de court terme."],
        implication="Le verdict deviendra « Démontré » quand le FPPN aura arrêté le taux servi, transmis les relevés individuels d'épargne et les âges des retraités (PBC, items 17, 99 et 114).",
        valable="épargne et dettes au 31/08/2026 (balance provisoire) ; intérêts dus calculés du 01/01/2020 au 31/08/2026 seulement ; provision du personnel au 31/12/2025",
        source=SRC_DEM)
contenu({"titre": "Demande 2 : suffisance des provisions", "sous_titre": f"Verdict : {V_PROV}. Provisions maladie et décès supérieures aux besoins, mais non calculées",
         "synthese": f"Provisions maladie et obsèques : {F('V2 prov tech', PROV_MAL + PROV_OBS)} au bilan, inchangées depuis 2020 ; intérêts d'épargne dus non provisionnés : {F('V2 intérêts', INT55)}"},
        "Une provision est suffisante si elle couvre ce que le FPPN devra payer pour des prestations déjà nées. Le tableau compare chaque provision du bilan au besoin mesuré sur les documents reçus.",
        tableau={"titre": tab("Provisions du bilan et besoins mesurés : chiffres qui fondent le verdict"), "entetes": ["Provision", "Au bilan", "Besoin mesuré", "Ce que le chiffre montre", "Détail"], "largeurs": [6.0, 4.8, 4.8, 10.8, 4.8], "taille": 8,
                 "lignes": [["Maladie (15801000)", F("V2 prov mal", PROV_MAL), F("V2 besoin mal", psap_m_elargie(2024)), "Besoin au 31/12/2024 : soins facturés l'année suivante (estimation élargie) ; les factures reçues sont déjà au passif (40110000)", ref("Provision maladie : montant nécessaire")],
                            ["Obsèques (15802000)", F("V2 prov obs", PROV_OBS), F("V2 dette déc", DETTE_DEC26), "La dette calculée par le FPPN est déjà inscrite aux comptes 41591000 et 41592000 : la provision fait double emploi", ref("Allocations décès : dette réelle et provision obsèques")],
                            ["Délai moyen de paiement des cliniques", "-", N("V2 délai 2025", DELAI[2025], 2) + " mois (2025)", f"{N('V2 délai 2026', DELAI[2026], 2)} mois au 31/08/2026 ; la règle de l'art. 52 est de 3 mois", ref("Provision maladie : délais de paiement des cliniques")],
                            ["Risques courts (statuts, art. 88)", "Aucune", F("V2 risques courts", 0), "Les cotisations mensuelles couvrent leur mois : aucune période à couvrir", ref("Provision pour risques courts")],
                            ["Intérêts dus sur l'épargne", "Aucune", F("V2 int (t)", INT55), "PEP 3,5 % et EFC 5,5 % du 01/01/2020 au 31/08/2026", ref("Épargne : intérêts dus non comptabilisés")],
                            ["Pertes probables sur les prêts du Grenier", "Aucune", F("V2 Grenier", GRD["2026-08-31"]), "Taux de la politique de crédit du Grenier (40 %, 80 %, 100 %)", ref("Prêts du Grenier : dépréciation selon la politique de crédit")],
                            ["Départ à la retraite du personnel (196)", F("V2 PIDR 196", PIDR25["provision_196"]), F("V2 PIDR", PIDR25["dbo"]), "Compte 196 supérieur au recalcul au 31/12/2025", ref("Provision du personnel par entité de 2021 à 2025")]]},
        disposition="dessous",
        analyse=[f"**Les provisions maladie et obsèques dépassent les besoins mesurés de {F('V2 excédent', PROV_MAL + PROV_OBS - psap_m_elargie(2024) - DETTE_DEC26)}, alors que les intérêts d'épargne ({F('V2 int (a)', INT55)}) et les pertes sur prêts ne sont pas provisionnés.** Les provisions ne sont pas insuffisantes en montant global ; elles sont mal affectées et non calculées, contrairement à l'article 89 des statuts."],
        implication="Les provisions maladie et obsèques doivent être recalculées chaque année (statuts, art. 89) ; l'excédent ne peut être réaffecté à l'épargne qu'après décision des organes et avis de l'AIRMS.",
        valable="provisions au bilan du 31/08/2026 ; besoin maladie mesuré au 31/12/2024 sur les factures de cliniques du journal SIN ; dette décès au 31/08/2026",
        source=SRC_DEM)
_a26, _b26, _c26 = couverture("31/08/2026", T_MOY)
_, _, _c26s = couverture("31/08/2026", T_MAX, CHOC_MAL)
contenu({"titre": "Demande 3 : adossement actif-passif", "sous_titre": f"Verdict : {V_ADO}. L'actif mobilisable dans l'année couvre {pc('V3 couverture', _c26)} des paiements de l'année",
         "synthese": f"Au 31/08/2026 : actif à moins d'un an {F('V3 actif', _a26)} ; paiements attendus dans l'année, épargne comprise, {F('V3 besoin', _b26)}"},
        "L'adossement vérifie que l'argent disponible arrive au moment où le FPPN doit payer. Sans échéancier de l'épargne, les sorties sont estimées avec les taux de sortie observés de 2021 à 2025.",
        tableau={"titre": tab("Adossement actif-passif : chiffres qui fondent le verdict"), "entetes": ["Mesure", "Valeur", "Ce que le chiffre montre", "Détail"], "largeurs": [8.4, 5.2, 12.8, 4.8], "taille": 8,
                 "lignes": [["Épargne couverte par les actifs qui lui sont réservés (31/08/2026)", pc("V3 RL1", b[2026]["couverture_epargne"]), "Banques et DAT de l'épargne et FCP rapportés à l'épargne due", ref("Adossement de l'épargne")],
                            ["Taux de sortie annuel de l'épargne, moyenne 2021 à 2025", pc("V3 taux moyen", T_MOY), f"Maximum observé : {pc('V3 taux max', T_MAX)}", ref("Adossement : à quel rythme l'épargne sort-elle ?")],
                            ["Couverture des paiements de l'année, scénario central", pc("V3 couv (t)", _c26), "Actif à moins d'un an / (dettes à moins d'un an + sorties d'épargne prévues)", ref("Adossement : tests de résistance")],
                            ["Couverture, test de résistance sévère", pc("V3 couv sévère", _c26s), "Sorties au taux maximal et hausse des prestations maladie la plus forte de 2024 et 2025", ref("Adossement : tests de résistance")],
                            ["Ratio de liquidité au 31/08/2026 (norme au moins 1)", N("V3 LIQ", RT["LIQ"]["valeurs"][2026], 2), "Trésorerie / dettes à court terme", ref("Calcul pas à pas : liquidité 2025")],
                            ["Ratio d'équilibre stable au 31/08/2026 (norme au plus 1)", N("V3 EQS", RT["EQS"]["valeurs"][2026], 2), "Emplois longs / ressources longues", ref("Calcul pas à pas : équilibre stable 2025")],
                            ["Réserve de trésorerie des risques longs (norme un trimestre)", N("V3 RL4", RT["RL4"]["valeurs"][2026], 2), "Aucune réserve constituée", ref("Calcul pas à pas : réserves des risques longs 2025")]]},
        disposition="dessous",
        analyse=["**L'épargne, dette exigible à tout moment, finance des actifs longs : terrains, clinique, créance sur l'hôpital et prêts du Grenier.** Même au taux de sortie moyen observé, l'actif mobilisable dans l'année ne couvre pas les paiements de l'année."],
        implication="L'adossement ne peut devenir suffisant sans un échéancier de l'épargne, une réserve de liquidité et une politique actif-passif votée par le conseil d'administration.",
        valable="bilans au 31/12/2025 et au 31/08/2026 (provisoire) ; sorties d'épargne observées de 2021 à 2025 ; actifs classés par nature comptable, sans échéancier contractuel",
        source=SRC_DEM)
_pf25 = next(x["produits_financiers"] for x in R["rendement_placements"] if x["exercice"] == 2025)
_pm25 = next(x["placements_moyens"] for x in R["rendement_placements"] if x["exercice"] == 2025)
contenu({"titre": "Demande 4 : adéquation des placements", "sous_titre": f"Verdict : {V_PLA}. Placements financiers de {pc('V4 part st', _pm25 / EP_DUE25)} de l'épargne due en 2025",
         "synthese": f"Rendement apparent moyen de 2021 à 2025 : {pc('V4 RDM', RDM)} ; produits financiers 2025 : {F('V4 pf 2025', _pf25)}, inférieurs aux intérêts d'une année sur l'EFC au taux le plus bas"},
        "Un portefeuille est adéquat s'il permet de payer les engagements à leur échéance, avec sécurité, liquidité, rendement et diversification (règlement d'exécution n° 004/2022, art. 17).",
        tableau={"titre": tab("Adéquation des placements : chiffres qui fondent le verdict"), "entetes": ["Mesure", "Valeur", "Ce que le chiffre montre", "Détail"], "largeurs": [8.4, 5.2, 12.8, 4.8], "taille": 8,
                 "lignes": [["Placements financiers moyens 2025 (DAT et FCP)", F("V4 pm (t)", _pm25), f"{pc('V4 part placements', _pm25 / EP_DUE25)} de l'épargne due au 31/12/2025", ref("Placements du FPPN")],
                            ["Produits financiers 2025", F("V4 pf (t)", _pf25), "Comprend le reversement des intérêts d'un DAT arrivé à maturité (rapport de gestion 2025)", ref("Rendement des placements")],
                            ["Intérêts d'une année sur l'EFC à 3,5 % (2025)", F("V4 int EFC", _E10[2025]["efc_fin"] * 0.035), "Supérieurs aux produits financiers de chaque année de 2021 à 2025", ref("Sensibilité : intérêts annuels dus sur l'EFC")],
                            ["Terrains, clinique (CMP) et créance sur l'hôpital au 31/08/2026", F("V4 immobilisé", b[2026]["terrains"] + b[2026]["cmp_investissements"] + b[2026]["creance_hopital_police"]), "Actifs peu liquides, sans évaluation récente", ref("Structure de l'actif")],
                            ["Financement des prêts du Grenier au 31/08/2026", F("V4 Grenier", b[2026]["grenier_financement"]), "Épargne prêtée aux membres, prêts d'une durée médiane de 10 mois", ref("Encours du Grenier")],
                            ["Politique de placement votée par le conseil d'administration", "Aucune", "Le directeur général décide seul (entretien du 19/10/2026)", ref("Décision de placement")]]},
        disposition="dessous",
        analyse=["**Le problème des placements n'est pas d'abord leur rendement, mais leur volume et leur liquidité** : l'essentiel de l'épargne est immobilisé dans des actifs qui ne produisent ni revenu financier ni liquidité."],
        implication="Une politique de placement doit fixer une part minimale d'actifs liquides au regard des sorties d'épargne observées et un rendement minimal au regard du taux servi, une fois ce taux arrêté.",
        valable="placements et produits financiers de 2021 à 2025, actifs au 31/08/2026 ; épargne due au 31/12/2025",
        source=SRC_DEM)
contenu({"titre": "Résumé exécutif", "sous_titre": "Le FPPN dégage des excédents mais reste fragile, peu documenté et hors normes prudentielles",
         "synthese": f"Fonds propres négatifs, épargne couverte à {pc('Couverture épargne 31/08/2026', b[2026]['couverture_epargne'])} par des actifs dédiés, {NB_NC} ratios prudentiels non conformes sur 36"},
        "Questions probables. Pourquoi « fragile » malgré des excédents ? Parce que les excédents annuels ne suffisent pas encore à effacer les pertes passées (fonds propres négatifs) et que l'épargne due aux mutualistes n'est pas couverte par des actifs disponibles. Pourquoi « peu documenté » ? Aucune base de prestations, pas de date de naissance des mutualistes, pas de barème officiel transmis. Pourquoi « hors normes » ? Les ratios de la Décision n° 05/2022/COM/UEMOA, appliqués aux balances, ne sont pas respectés pour la réserve obligatoire, les fonds propres, la liquidité et les frais. Chaque ligne renvoie à une partie détaillée du rapport.",
        tableau={"titre": tab("Principaux constats au 23/10/2026"), "entetes": ["Domaine", "Constat", "Chiffre clé"],
                 "largeurs": [4.2, 17.2, 9.8], "lignes": [
                     ["Données", "Les précomptes mensuels ne commencent qu'en 2023 ou 2024 selon le poste, et aucune base de prestations n'a été reçue.", N("Lignes de précomptes contrôlées", 4454416) + " lignes de précomptes contrôlées"],
                     ["Cotisations", "Le précompte total vaut 5 % du salaire de base, mais la Solde le répartit en 3,77 % maladie et 1,23 % décès, au lieu du " + BAREME + " (3,5 % et 1,5 %).", pcv("Taux total observé", 5.0) + " = " + pcv("Taux maladie observé", 100 * T936["mediane"]) + " + 1,23 %"],
                     ["Maladie", "Les prestations maladie comptabilisées augmentent bien plus vite que les cotisations maladie.", "S/C " + pc("S/C maladie 2021", s[2021]["sc_maladie"]) + " (2021), " + pc("S/C maladie 2025", s[2025]["sc_maladie"]) + " (2025), " + pc("S/C maladie 2026 (rés)", s[2026]["sc_maladie"]) + " (31/08/2026)"],
                     ["Provisions", "Les provisions pour prestations à payer gardent le même montant de 2020 à 2026, alors que les prestations ont fortement augmenté.", F("Provisions techniques", b[2026]["provisions_techniques"])],
                     ["Épargne", "L'épargne due aux mutualistes (PEP, EFC) dépasse largement les actifs qui lui sont réservés.", F("Épargne totale 31/08/2026", b[2026]["epargne_totale"])],
                     ["Fonds propres", "Les fonds propres restent négatifs à chaque clôture, même si les excédents réduisent l'écart.", F("Fonds propres 2025", b[2025]["fonds_propres"]) + " (2025), " + F("Fonds propres 2026 (rés)", b[2026]["fonds_propres"]) + " (31/08/2026)"],
                     ["Conformité", "Les ratios de la Décision n° 05/2022 ne sont pas respectés pour la réserve obligatoire, les fonds propres, la liquidité et les frais de fonctionnement.", f"{NB_NC} non conformes, {NB_C} conformes, {NB_ATT} en attente"],
                     ["Grenier", "Le portefeuille de prêts croît très vite, et le FPPN le finance en grande partie.", F("Encours Grenier 31/08/2026", GP["encours"]) + " d'encours"]]},
        disposition="dessous",
        analyse=["**Les fragilités sont surtout financières et documentaires, pas d'exploitation.** Le FPPN encaisse plus qu'il ne dépense chaque année depuis 2020, sauf sur les huit premiers mois de 2026."],
        implication="Les trois chantiers prioritaires sont la transmission des données manquantes, le recalcul des provisions et un plan de mise en conformité prudentielle à présenter à l'AIRMS.",
        source="Source : analyses Adenka sur les documents reçus du 15/10/2026 au 23/10/2026 et les textes de l'UEMOA")

# ---- v12 : boules de couleur (calcul) ; le dessin est dans le moteur des tableaux
BOULES = "🔴🟠🟡🟢⚪"
STATUT_BOULE = {"Conforme": "🟢", "Non conforme": "🔴", "Au-dessus": "🟠", "En attente": "🟡", "À apprécier": "⚪"}


LEG_CONF = [("🔴", "Non conforme", "norme non respectée"), ("🟠", "Non démontré", "non démontré par les documents reçus"), ("🟡", "En attente", "document du FPPN attendu"),
            ("⚪", "À apprécier", "norme non chiffrée"), ("🟢", "Conforme", "norme respectée")]


def boule_statut(st):
    """Boule d'un statut de ratio (« Non conforme (indicatif) » compte comme « Non conforme »)."""
    return STATUT_BOULE.get(str(st).split(" (")[0].strip(), "")


def bv(texte, boule):
    """Case de matrice : boule devant la valeur (la valeur reste lisible)."""
    return f"{boule} {texte}" if boule else texte


def boule_fourchette(v, bas, haut):
    """Repère en fourchette (sinistralité par garantie) : 🟢 dans la fourchette, 🟠 en dessous, 🔴 au-dessus."""
    if v is None:
        return ""
    return "🟢" if bas <= v <= haut else ("🟠" if v < bas else "🔴")


def boule_seuil(v, seuil, sens="max"):
    """Seuil simple : sens « max » = la valeur doit rester au plus égale au seuil, « min » = au moins égale."""
    if v is None:
        return ""
    ok = v <= seuil if sens == "max" else v >= seuil
    return "🟢" if ok else "🔴"



# ---- v12 : définition et lecture de chaque ratio (formule et norme : Décision n° 05/2022 ; lecture : Adenka)
def val_ratio(x, a, retraite=False):
    v = (x.get("valeurs_retraitees") if retraite else x["valeurs"]).get(a)
    if v is None:
        return "n.d."
    if abs(v) < 5e-7:
        v = 0.0
    lab = f"{x['code']} {a}{' retraité' if retraite else ''}"
    if x["unite"] == "%":
        return pc(lab, v)
    if x["unite"] == "mois":
        return N(lab, v, 2) + " mois"
    if x["unite"] == "jours":
        return N(lab, v, 0) + " jours"
    if x["unite"] == "FCFA":
        return F(lab, v)
    return N(lab, v, 2)


DEFS = {  # code : (ce que le ratio mesure, comment le lire)
    "SIN": ("La part des cotisations de l'année consommée par les prestations, provisions comprises, toutes garanties confondues.",
            "Sous 70 %, les cotisations sont peu consommées. Au-dessus de 80 %, il ne reste pas assez pour payer les frais et constituer des réserves."),
    "MSE": ("La capacité des fonds propres libres à absorber une perte imprévue, rapportée au minimum exigé par le texte.",
            "Au moins 1 : la marge est suffisante. Sous 1 : les fonds propres libres sont trop faibles pour le volume d'activité."),
    "ROB": ("Le nombre de mois de prestations que la réserve obligatoire permettrait de payer.",
            "Moins de 6 mois : la réserve ne couvre pas un choc de prestations. Plus de 9 mois : réserve au-delà de la norme."),
    "FEP": ("La part des dettes, provisions techniques comprises, couverte par les fonds propres et assimilés.",
            "Sous 1 : les dettes dépassent les fonds propres et assimilés. Le FPPN dépend alors de ses créanciers et de l'épargne des adhérents."),
    "LIQ": ("La capacité à payer les dettes à court terme avec la trésorerie disponible.",
            "Sous 1 : la trésorerie ne suffit pas à payer les dettes exigibles à court terme."),
    "EQS": ("Le financement des emplois longs (immobilisations, prêts) par des ressources longues.",
            "Au-dessus de 1 : une partie des emplois longs est financée par des ressources de court terme, ce qui crée un risque de liquidité."),
    "AUT": ("La capacité des cotisations à payer l'ensemble des charges de l'année.",
            "Sous 1 : les cotisations ne suffisent pas à payer les charges de l'année."),
    "EXN": ("La part des cotisations conservée en excédent à la fin de l'exercice.",
            "Au-dessus de 5 % : excédent supérieur à la norme d'une mutuelle sociale. Négatif : déficit de l'exercice."),
    "TFP": ("La capacité de l'ensemble des produits, hors subventions, à payer l'ensemble des charges.",
            "Sous 100 % : les produits ne couvrent pas les charges et l'exercice est déficitaire."),
    "RL1": ("La part de l'épargne due (PEP et EFC) couverte par les actifs qui lui sont réservés : banques et dépôts à terme de l'épargne, FCP.",
            "Sous 1 : une partie de l'épargne des adhérents n'est garantie par aucun actif réservé."),
    "RL2": ("La capacité à payer les dettes à court terme avec les créances à court terme et la trésorerie.",
            "Sous 1 : les créances et la trésorerie ne suffisent pas à payer les dettes à court terme."),
    "RL3": ("La capacité à payer immédiatement les dettes à court terme avec la seule trésorerie.",
            "Sous 1 : un retrait massif ne pourrait pas être payé avec l'argent disponible."),
    "RL4": ("La réserve de trésorerie des risques longs, en trimestres de sorties d'épargne et de frais.",
            "0 : aucune réserve de trésorerie n'est constituée."),
    "RL5": ("La couverture par les fonds propres des sorties d'épargne des deux derniers exercices.",
            "Négatif : les fonds propres sont négatifs et ne couvrent aucune sortie."),
    "RL6": ("Le rendement des sommes réellement placées (dépôts à terme et FCP).",
            "À comparer au taux du marché et au taux servi aux épargnants : un rendement inférieur au taux servi creuse la dette d'épargne."),
    "RL7": ("L'argent laissé sur des comptes courants qui ne rapportent aucun intérêt.",
            "Plus le solde est élevé, plus le manque à gagner est important. La norme demande un solde proche de zéro."),
    "CHF": ("La part des cotisations absorbée par les charges de fonctionnement.",
            "Au-dessus de 20 % : les frais prennent une part des cotisations qui manque aux prestations et aux réserves."),
    "CHB": ("La part de l'ensemble des produits absorbée par les charges de fonctionnement brutes.",
            "Au-dessus de 15 % : les frais de gestion sont élevés au regard des ressources."),
    "CUP": ("Les frais de fonctionnement engagés pour 1 FCFA de prestations.",
            "Au-dessus de 15 % : la gestion coûte cher par rapport aux prestations servies."),
    "CHP": ("La part des charges de fonctionnement consacrée au personnel.",
            "Au-dessus de 85 % : les salaires laissent peu de place aux autres dépenses de fonctionnement."),
    "REC": ("La part des cotisations appelées effectivement encaissée.",
            "Sous 90 % : une partie des cotisations dues n'est pas encaissée."),
    "FID": ("La part des adhérents d'une année encore adhérents l'année suivante.",
            "Sous 80 % : les départs sont nombreux et la mutualisation s'affaiblit."),
    "DPP": ("Le délai moyen entre la prestation et le paiement du prestataire de soins.",
            "À comparer au délai contractuel des conventions (non transmis) et au délai de trois mois du règlement intérieur (art. 52)."),
    "FOR": ("La part des produits investie dans la formation du personnel.",
            "Au plus 3 % selon la norme."),
    "CRB": ("La croissance du nombre de bénéficiaires d'une année à l'autre.",
            "À comparer à l'objectif du plan de développement stratégique (non transmis)."),
}


def _defs_diapo(k, codes):
    """Diapositive de définitions avant chaque tableau de bord : formule, ce que le ratio mesure, lecture, norme."""
    noms = {1: "solvabilité et réserves", 2: "risques longs", 3: "gestion"}
    contenu({"titre": f"Définition et lecture des ratios ({k}/3) : {noms[k]}", "sous_titre": "Ce que chaque ratio mesure et comment lire sa valeur, avant le tableau de bord",
             "synthese": "Formule et norme de la Décision n° 05/2022 ; la lecture explique ce que la valeur dit de la situation du FPPN"},
            "Formules et normes de l'annexe de la Décision n° 05/2022/COM/UEMOA. La colonne « Comment le lire » est rédigée par Adenka pour un lecteur non spécialiste.",
            tableau={"titre": tab(f"Définition et lecture des ratios, {noms[k]}"), "entetes": ["Ratio", "Formule", "Ce que le ratio mesure", "Comment le lire", "Norme"],
                     "largeurs": [4.8, 7.0, 7.2, 8.8, 3.4], "taille": 8, "sans_code": True,
                     "lignes": [[RT[c]["nom"].replace(" (risques longs)", ""), RT[c]["formule"], DEFS[c][0], DEFS[c][1],
                                 RT[c]["norme"].replace(" (tableau 1) ; 9 mois au § 1.2.2.5.4", "").replace(" (marge disponible au moins égale à la marge exigée)", "")] for c in codes]},
            valable="formules et normes en vigueur depuis le 01/06/2022",
            source="Source : " + DEC + " ; lecture Adenka")


# ---- v12 : synthèse des constats au format fait, norme ou référence, écart chiffré, effet pour le FPPN, criticité
_GR26 = GRD.get("2026-08-31", 0.0)
_PLG = V11["placements_gl"]
_RG = {(x["exercice"], x["garantie"]): x for x in V11["ratios_garantie"]}
_PP = V11["pidr_plafond"]
_BESOIN = psap_m_elargie(2024) + TARD_DEC
_ACT, _PAI, _COUV = couverture("31/08/2026", T_MOY)
SYN1 = [
    ["❑ Provisions\nmaladie et décès",
     f"Les provisions maladie (**{F('SY prov mal', PROV_MAL)}**) et obsèques (**{F('SY prov obs', PROV_OBS)}**) ont le même montant depuis le 01/09/2020. Seul mouvement : un changement de compte en 2024.",
     "Statuts, art. 89 : provisions recalculées à chaque clôture. Règlement d'exécution n° 004/2022, art. 10 : calcul dossier par dossier",
     f"Besoin mesuré : **{F('SY besoin', _BESOIN)}**, soit le montant que les provisions devaient couvrir à la clôture. Il comprend les soins de 2024 facturés en 2025, estimés à {F('SY besoin mal', psap_m_elargie(2024))} à partir des factures de 2024 repérées dans le journal SIN de 2025, et {F('SY besoin dec', TARD_DEC)} d'allocations de six décès de 2021 enregistrés en 2022. Les provisions dépassent ce besoin de **{F('SY exces', PROV_MAL + PROV_OBS - _BESOIN)}**",
     f"Résultat et fonds propres sous-évalués d'autant. Les allocations décès dues ({F('SY dette dec', DETTE_DEC26)} au 31/08/2026) sont déjà en dette aux comptes 41591000 et 41592000 : la provision obsèques les compte une seconde fois",
     "Critique"],
    ["❑ Épargne\nPEP et EFC",
     "Aucun intérêt n'est comptabilisé sur l'épargne du 01/01/2020 au 31/08/2026. Trois taux figurent dans trois documents (3,5 %, 5,5 % et 14 %), aucun n'est arrêté par une décision signée.",
     "Manuel des produits d'épargne et cahier des charges EFC-PEP de 2021 : épargne rémunérée. Règlement d'exécution n° 004/2022, art. 11 : provision des risques longs",
     f"Intérêts dus non comptabilisés : **{F('SY int', INT55)}** au 31/08/2026, aux taux de 3,5 % (PEP) et 5,5 % (EFC), soit **{pc('SY int part', INT55 / b[2026]['epargne_totale'])}** de l'épargne inscrite au bilan ({F('SY ep', b[2026]['epargne_totale'])})",
     "La dette envers les épargnants est sous-évaluée d'autant. Le FPPN ne peut pas remettre à chaque épargnant un relevé de ses droits",
     "Critique"],
    ["❑ Adossement\nactif-passif",
     f"Au 31/08/2026, les actifs disponibles à moins d'un an valent **{F('SY act', _ACT)}**. Les paiements attendus dans l'année, sorties d'épargne au taux observé de {pc('SY tsort', T_MOY)} par an comprises, valent **{F('SY pai', _PAI)}**.",
     "Règlement d'exécution n° 004/2022 : placements sûrs et liquides (art. 17), réserve de trésorerie d'un trimestre pour les risques longs (art. 19)",
     f"Couverture des paiements de l'année : **{pc('SY couv', _COUV)}**. Manque cumulé à cinq ans : **{F('SY gap', -GAP[4][5])}**. Réserve de trésorerie constituée : 0 FCFA",
     "Pour payer des retraits supérieurs à la trésorerie, le FPPN devrait vendre ou mobiliser des actifs longs : terrains, clinique, créance sur l'Hôpital de la Police",
     "Critique"],
    ["❑ Conformité\nprudentielle",
     f"**{NB_NC}** des 36 indicateurs de la Décision n° 05/2022 sont non conformes de 2022 à 2025, dont la réserve obligatoire ({val_ratio(RT['ROB'], 2025)} en 2025), la liquidité ({val_ratio(RT['LIQ'], 2025)}) et les fonds propres et assimilés ({val_ratio(RT['FEP'], 2025)}).",
     "Décision n° 05/2022 : réserve obligatoire de 6 à 9 mois de prestations, liquidité et fonds propres et assimilés au moins égaux à 1",
     f"En 2025 : il manque {N('SY rob ec', 6 - RT['ROB']['valeurs'][2025], 2)} mois de réserve obligatoire, {N('SY liq ec', 1 - RT['LIQ']['valeurs'][2025], 2)} point de liquidité et {N('SY fep ec', 1 - RT['FEP']['valeurs'][2025], 2)} point de fonds propres et assimilés",
     "Deux exercices de suite non conformes : retrait d'agrément possible (règlement n° 07/2009, art. 84) et sanctions de l'arrêté n° 00115/MEPS/CAB du 12/12/2023",
     "Critique"],
    ["❑ Fonds propres",
     f"Les fonds propres, résultat compris, sont négatifs à chaque clôture depuis 2020 : **{F('SY fp', b[2025]['fonds_propres'])}** au 31/12/2025 et **{F('SY fp26', b[2026]['fonds_propres'])}** au 31/08/2026.",
     "Décision n° 05/2022 : fonds propres et assimilés au moins égaux aux dettes",
     "Les pertes reportées dépassent l'ensemble des réserves",
     "Aucun coussin pour absorber un choc de prestations, une perte sur les prêts du Grenier ou une baisse de valeur des actifs",
     "Élevé"],
]


def _syn_diapo(_k, _lig, _suite):
    contenu({"titre": f"Synthèse des constats ({_k}/2)", "sous_titre": "Cinq constats critiques portent sur les engagements, la liquidité et la conformité" if _k == 1 else "Les autres constats touchent la sinistralité, les placements, les données et la gouvernance",
             "synthese": "Chaque constat donne le fait chiffré, la norme qui s'applique, l'écart mesuré et son effet pour le FPPN"},
            "Tableau de synthèse : le fait constaté et depuis quand, la norme ou la référence qui s'applique, l'écart chiffré, l'effet concret pour le FPPN et le niveau de criticité. Chiffres identiques à ceux du corps du rapport.",
            tableau={"titre": tab(f"Synthèse des constats ({_k}/2) : {_suite}"), "entetes": ["Section", "Constat", "Norme ou référence", "Écart chiffré", "Effet pour le FPPN", "Criticité"],
                     "largeurs": [2.8, 7.6, 5.6, 7.0, 6.2, 2.0], "taille": 7.5, "lignes": _lig},
            disposition="dessous",
            valable="chiffres du corps du rapport, de 2020 au 31/08/2026",
            source="Source : analyses Adenka, parties 3 à 10 du rapport. Provisions : balances de 2020 au 31/08/2026, journal SIN des grands livres des tiers 2023 à 2025, états de fin d'année des allocations décès. Épargne : manuel et cahier des charges des produits d'épargne")


_syn_diapo(1, SYN1, "provisions, épargne, adossement, conformité, fonds propres")
SYN2 = [
    ["❑ Sinistralité\nmaladie",
     f"La sinistralité maladie passe de **{pc('SY sc21', s[2021]['sc_maladie'])}** en 2021 à **{pc('SY sc25', s[2025]['sc_maladie'])}** en 2025 et **{pc('SY sc26', s[2026]['sc_maladie'])}** sur les huit mois de 2026.",
     "Repère : fourchette de 70 % à 80 % du taux réglementaire, qui est global (Décision n° 05/2022, § 1.2.1.2.5)",
     f"Ratio combiné maladie, frais compris : **{pc('SY rcm', _RG[(2025, 'Maladie')]['ratio_combine'])}** en 2025. Résultat technique maladie 2025 : **{F('SY rtm', _RG[(2025, 'Maladie')]['resultat_technique'])}**",
     "La cotisation maladie de 3,5 % du salaire de base ne couvre plus le coût complet de la garantie. Le décès, excédentaire, finance la maladie",
     "Élevé"],
    ["❑ Placements",
     f"En 2025, l'encours moyen réellement placé (dépôts à terme et FCP, lu jour par jour dans les grands livres) vaut **{F('SY pl', _PLG['2025']['placements_moyens'])}**.",
     "Règlement d'exécution n° 004/2022, art. 17 : placements sûrs, rentables, liquides et diversifiés, selon une politique votée",
     f"Soit **{pc('SY pl part', _PLG['2025']['placements_moyens'] / EP_DUE25)}** de l'épargne due au 31/12/2025 ({F('SY ep25', EP_DUE25)}). Rendement moyen de 2021 à 2025 : **{pc('SY rdm', V11['rendement_moyen_2021_2025'])}** par an, sous le taux de 5,5 % annoncé sur l'EFC",
     "Les produits financiers ne peuvent pas payer les intérêts dus sur l'épargne. [[Le directeur général décide seul des placements]]",
     "Élevé"],
    ["❑ Données",
     "Ni la base des prestations datée, ni les dates de naissance des mutualistes ne sont reçues. Les fichiers maladie et décès de décembre 2024 sont vides (0 octet).",
     f"PBC du 23/10/2026, n° {PBCM[9]}, {PBCM[62]} et {PBCM[71]}. Règlement d'exécution n° 004/2022, art. 7 à 10 : provisions calculées dossier par dossier",
     f"Date de naissance absente des fichiers de précomptes maladie et décès. Connue pour {N('SY rdn', RDN['nb'])} retraités seulement, dans les fichiers de la CGRAE de février à avril 2024",
     "Mortalité observée, engagement santé des retraités et projection par âge impossibles à calculer",
     "Élevé"],
    ["❑ Gouvernance",
     "[[Le directeur général décide seul des placements]]. Le projet de Cocody n'a ni budget ni plan de financement transmis. La désignation du directeur général n'est documentée que pour 2020 à 2023.",
     "Statuts, art. 85 : délibération du conseil d'administration pour chaque placement",
     "Aucune délibération de placement reçue",
     "L'épargne des adhérents est employée sans décision des organes",
     "Élevé"],
    ["❑ Prêts du Grenier",
     f"L'encours des prêts vaut **{F('SY gre', GP['encours'])}** au 31/08/2026. Aucune dépréciation n'est passée depuis 2022.",
     "Politique de crédit du Grenier : dépréciation de 40 %, 80 % ou 100 % de l'encours selon l'ancienneté du retard",
     f"Dépréciation nécessaire au 31/08/2026 : **{F('SY dep', _GR26)}**",
     "Actif et fonds propres surévalués de ce montant. L'épargne des adhérents finance l'essentiel du crédit",
     "Moyen"],
    ["❑ Personnel\ndu FPPN",
     f"Le compte 196 porte **{F('SY 196', _PP['provision_196'])}** au 31/12/2025 pour les indemnités de départ à la retraite du personnel du FPPN.",
     "Décret n° 2017-210, art. 4 et 5 : indemnité de départ. Convention collective, art. 40 : plafond de 25 fois le SMIG annuel",
     f"Recalcul agent par agent : **{F('SY pidr', _PP['dbo_sans_plafond'])}** sans plafond, **{F('SY pidr cap', _PP['dbo_avec_plafond'])}** avec le plafond. Le compte 196 dépasse le recalcul de {F('SY pidr ec', _PP['provision_196'] - _PP['dbo_sans_plafond'])}",
     "Provision suffisante. La méthode du FPPN n'est pas documentée",
     "Faible"],
]
_syn_diapo(2, SYN2, "sinistralité, placements, données, gouvernance, Grenier, personnel")

_omis(_snap(), {"titre": "Réponse aux quatre questions de la mission", "sous_titre": "Deux constats sont établis, deux conclusions attendent les données du FPPN",
         "synthese": "Le non-adossement et l'inadéquation des placements sont démontrés ; les engagements et la suffisance des provisions ne peuvent pas encore être chiffrés"},
        "Les quatre tâches actuarielles de la mission, sur la période du 01/01/2020 au 31/08/2026. Pour chacune : ce qui peut être conclu, ce qui ne peut pas l'être et pourquoi, le sens dans lequel le résultat évoluera, et les documents qui débloquent la réponse. Le détail figure dans les conclusions de la partie 9.",
        tableau={"titre": tab("Réponse aux quatre questions de la mission au 19/10/2026"), "entetes": ["Tâche", "Ce qui est établi", "Ce qui ne l'est pas, et pourquoi", "Sens d'évolution", "Documents (PBC)"], "largeurs": [4.4, 8.0, 8.4, 4.4, 6.0], "taille": 8.5,
                 "lignes": [["1. Engagements envers les adhérents", "Épargne due : " + F("Épargne RE4", EP26["pep"] + EP26["efc"]) + " au 31/08/2026", "Provisions non estimées ; provisions mathématiques et sinistres tardifs absents ; données individuelles non reçues", "Hausse", "57, 71, 72, 79, 92"],
                            ["2. Suffisance des provisions", "Les provisions sont des montants hérités, sans calcul", "Suffisance non démontrable sans recalcul dossier par dossier et méthode CIMA", "Inconnu", "71, 72, 92, 93"],
                            ["3. Adossement actif-passif", "Non-adossement : couverture des engagements connus de " + pc("RE sc1", SCEN["31/08/2026"][0] / ENG["31/08/2026"]) + " au mieux", "Ampleur réelle et flux futurs : échéancier de l'épargne et évaluations d'actifs manquants", "Aggravation", "57, 84, 86 à 88, 99"],
                            ["4. Adéquation des placements", "Conditions de liquidité et de diversification non remplies", "Sécurité et rendement à apprécier : décision ministérielle et relevés non reçus", "Stable", "54, 80 à 82"]]},
        disposition="dessous",
        analyse=["**Les deux constats établis sont robustes : les données manquantes ne peuvent que les aggraver.** Les deux autres conclusions dépendent des bases de prestations, des âges des mutualistes et des règlements de l'EFC et du PEP."],
        implication="La transmission des documents listés conditionne la version définitive du rapport. Les items 57, 71, 72 et 92 sont prioritaires : ils débloquent à la fois les engagements, les provisions et l'adossement.",
        source="Source : dossier de préparation de la mission (13/10/2026) ; analyses Adenka, parties 6 et 9 ; PBC pointée au 19/10/2026")

# ---- v13 : matrice de criticité des constats (criticité × horizon de la recommandation qui y répond)
ROUGE_M, JAUNE_M, VERT_M, ORANGE_M = "C00000", "FFD966", "70AD47", "EFAC62"
_CRIT_L = ["Critique", "Élevé", "Moyen", "Faible"]
_HOR_C = ["Court terme (moins de 6 mois)", "Moyen terme (6 à 18 mois)", "Long terme (plus de 18 mois)"]
_PLACE = {("Critique", 0): "Provisions maladie et décès (rec. 1)\nÉpargne PEP et EFC (rec. 2)\nConformité prudentielle (rec. 4)",
          ("Critique", 1): "Adossement actif-passif (rec. 5)",
          ("Élevé", 0): "Fonds propres (rec. 4)\nDonnées (rec. 3)",
          ("Élevé", 1): "Sinistralité maladie (rec. 6)\nPlacements (rec. 5)\nGouvernance (rec. 7)",
          ("Moyen", 1): "Prêts du Grenier (rec. 7)",
          ("Faible", 2): "Personnel du FPPN (rec. 9)"}
_ZONES_C = {"Critique": [ROUGE_M, ROUGE_M, ORANGE_M], "Élevé": [ROUGE_M, ORANGE_M, JAUNE_M], "Moyen": [ORANGE_M, JAUNE_M, VERT_M], "Faible": [JAUNE_M, VERT_M, VERT_M]}
_BOULE_CRIT = {"Critique": "🔴", "Élevé": "🟠", "Moyen": "🟡", "Faible": "🟢"}
_MAT_C, _FONDS_C = [], {}
for i_, cr_ in enumerate(_CRIT_L):
    _MAT_C.append([_BOULE_CRIT[cr_] + " " + cr_] + [_PLACE.get((cr_, k_), "") for k_ in range(3)])
    for k_ in range(3):
        _FONDS_C[(i_, k_ + 1)] = _ZONES_C[cr_][k_]
contenu({"titre": "Matrice de criticité des constats", "sous_titre": "Trois des quatre constats critiques se traitent dans les six mois",
         "synthese": "Criticité du constat en lignes, horizon de la recommandation qui y répond en colonnes"},
        "Chaque constat de la synthèse est placé selon sa criticité et l'horizon de la recommandation qui le traite. La couleur de la case donne l'ordre de traitement.",
        tableau={"titre": tab("Matrice de criticité des constats et horizon des recommandations"), "entetes": ["Criticité \\ Horizon"] + _HOR_C,
                 "largeurs": [5.2, 8.67, 8.67, 8.66], "taille": 9.5, "sans_code": True, "hauteur_ligne": 1.6, "fonds": _FONDS_C, "centre": True,
                 "legende": [("🔴", "Priorité 1", "à engager immédiatement"), ("🟠", "Priorité 2", "à engager dans l'année"), ("🟡", "Priorité 3", "à planifier"), ("🟢", "Priorité 4", "à surveiller")],
                 "lignes": _MAT_C},
        disposition="dessous",
        analyse=["**Les provisions, l'épargne et la conformité prudentielle sont critiques et se traitent à court terme : ce sont les recommandations 1, 2 et 4.** L'adossement, critique lui aussi, demande une réserve de liquidité et une politique de placement qui se construisent sur 6 à 18 mois (recommandation 5)."],
        valable="constats de la synthèse et recommandations du rapport",
        source="Source : synthèse des constats et recommandations du rapport (analyses Adenka)")

# ---- v12 : indicateurs clés de 2020 au 31/08/2026, valeur et boule dans chaque case
_RGm = {(x["exercice"], x["garantie"]): x for x in V11["ratios_garantie"]}
_AN = list(range(2020, 2027))
_ENT_AN = [str(a) if a < 2026 else "31/08/2026" for a in _AN]
_GEa = {int(g["arrete"][-4:]): g["encours"] for g in GE if g["arrete"].startswith("31/12") or g["arrete"] == "31/08/2026"}


def _rt(c, a, en_pc=False):
    """Valeur d'un ratio de la Décision et boule de son statut de l'année."""
    x = RT[c]
    v = x["valeurs"].get(a)
    if v is None:
        return "n.d."
    t_ = pc(f"{c} {a} (%)", v) if en_pc else val_ratio(x, a)
    return bv(t_, boule_statut(x["statuts"].get(a, "")))


def _gar(g, a, cle):
    """Ratio d'une garantie : sinistralité (repère 70 % à 80 %), ratio combiné (seuil 100 %), résultat technique (seuil 0)."""
    x = _RGm.get((a, g))
    if not x:
        return "n.d."
    v = x[cle]
    if cle == "sinistralite":
        return bv(pc(f"RG sin {a} {g}", v), boule_fourchette(v, 0.70, 0.80))
    if cle == "ratio_combine":
        return bv(pc(f"RG rc {a} {g}", v), boule_seuil(v, 1.0, "max"))
    return bv(F(f"RG rt {a} {g}", v), boule_seuil(v, 0.0, "min"))


LEG_MAT = [("🔴", "Non respecté", "norme ou seuil dépassé"), ("🟢", "Respecté", "norme ou seuil respecté"),
           ("🟠", "Sous la fourchette", "sinistralité d'une garantie sous 70 %"), ("⚪", "À apprécier", "norme non chiffrée")]
_IND = [
    ["Cotisations acquises (comptes 70)", "Sans norme"] + [F(f"IC cot {a}", RT["SIN"]["den"][a]) for a in _AN],
    ["Prestations, provisions comprises (comptes 60, 6901, 7901)", "Sans norme"] + [F(f"IC pre {a}", RT["SIN"]["num"][a]) for a in _AN],
    ["Taux de sinistralité réglementaire, toutes garanties", "70 % à 80 %"] + [_rt("SIN", a) for a in _AN],
    ["Sinistralité maladie", "Repère 70 % à 80 %"] + [_gar("Maladie", a, "sinistralite") for a in _AN],
    ["Sinistralité décès", "Repère 70 % à 80 %"] + [_gar("Décès", a, "sinistralite") for a in _AN],
    ["Ratio combiné maladie (prestations et frais)", "Au plus 100 %"] + [_gar("Maladie", a, "ratio_combine") for a in _AN],
    ["Taux des charges de fonctionnement", "Au plus 20 %"] + [_rt("CHF", a) for a in _AN],
    ["Taux d'excédent net", "Au plus 5 %"] + [_rt("EXN", a) for a in _AN],
    ["Épargne due (PEP et EFC)", "Sans norme"] + [F(f"IC ep {a}", b[a]["epargne_totale"]) for a in _AN],
    ["Épargne couverte par des actifs réservés", "Au moins 100 %"] + [_rt("RL1", a, True) for a in _AN],
    ["Fonds propres, résultat compris", "Sans norme"] + [F(f"IC fp {a}", b[a]["fonds_propres"]) for a in _AN],
    ["Fonds propres et assimilés / dettes", "Au moins 1"] + [_rt("FEP", a) for a in _AN],
    ["Ratio de liquidité", "Au moins 1"] + [_rt("LIQ", a) for a in _AN],
    ["Encours des prêts du Grenier", "Sans norme"] + [F(f"IC gre {a}", _GEa[a]) if a in _GEa else "n.d." for a in _AN],
]
contenu({"titre": "Indicateurs clés de 2020 au 31/08/2026", "sous_titre": "La sinistralité monte, l'épargne reste peu couverte et les normes de solidité ne sont jamais atteintes",
         "synthese": f"Sinistralité réglementaire : {pc('IC sin 2021', RT['SIN']['valeurs'][2021])} en 2021, {pc('IC sin 2025', RT['SIN']['valeurs'][2025])} en 2025, {pc('IC sin 2026', RT['SIN']['valeurs'][2026])} sur huit mois de 2026"},
        "Chaque case donne la valeur de l'année et une boule qui la compare à la norme de la Décision n° 05/2022 ou au repère indiqué. 2020 et 2021 sont antérieurs à l'entrée en vigueur de la Décision (01/06/2022) : ils sont indicatifs. 2026 couvre huit mois, sur la balance provisoire.",
        tableau={"titre": tab("Indicateurs clés, de 2020 au 31/08/2026"), "entetes": ["Indicateur", "Norme ou repère"] + _ENT_AN,
                 "largeurs": [6.4, 3.0] + [3.114] * 7, "taille": 7.5, "sans_code": True, "legende": LEG_MAT[:3], "lignes": _IND},
        disposition="dessous",
        analyse=[f"**La sinistralité réglementaire passe de {pc('IC sin 2021 (a)', RT['SIN']['valeurs'][2021])} en 2021 à {pc('IC sin 2025 (a)', RT['SIN']['valeurs'][2025])} en 2025, et la maladie dépasse 80 % sur les huit mois de 2026 ({pc('IC sin mal 2026', _RGm[(2026, 'Maladie')]['sinistralite'])}).** "
                 "Les ratios de solidité (couverture de l'épargne, fonds propres et assimilés, liquidité) sont non conformes chaque année de 2020 au 31/08/2026."],
        implication="Les écarts se répètent chaque année : ils sont structurels. Le plan de mise en conformité doit traiter leurs causes (provisions, épargne, liquidité), pas une année isolée.",
        valable="comptes de 2020 à 2025 et du 01/01 au 31/08/2026 (balance provisoire) ; 2020 et 2021 indicatifs ; encours du Grenier aux 31/12 et au 31/08/2026",
        source=SRC_COMPTA + " ; " + DEC + " ; " + SRC_GRENIER[9:])
_DEF_IND = [
    ["Cotisations acquises", "Les cotisations de l'année revenant au FPPN (comptes 70, variation des provisions sur cotisations comprise).", "Elles mesurent la ressource technique de l'année."],
    ["Prestations, provisions comprises", "Les prestations de l'année (comptes 60) et la variation des provisions techniques (6901, 7901).", "Elles mesurent le coût des garanties de l'année."],
    ["Taux de sinistralité réglementaire", DEFS["SIN"][0], DEFS["SIN"][1] + " Seul ce taux global a une norme dans la Décision."],
    ["Sinistralité maladie ou décès", "Les prestations d'une garantie divisées par ses cotisations.", "Elle montre quelle garantie consomme ses cotisations. La fourchette de 70 % à 80 % sert de repère."],
    ["Ratio combiné maladie", "Les prestations maladie et les charges de fonctionnement affectées, divisées par les cotisations maladie.", "Au-dessus de 100 %, la garantie coûte plus qu'elle ne rapporte, frais compris."],
    ["Taux des charges de fonctionnement", DEFS["CHF"][0], DEFS["CHF"][1]],
    ["Taux d'excédent net", DEFS["EXN"][0], DEFS["EXN"][1]],
    ["Épargne couverte par des actifs réservés", DEFS["RL1"][0], DEFS["RL1"][1]],
    ["Fonds propres, résultat compris", "Les ressources propres du FPPN (comptes 10 à 14) : réserves, report à nouveau et résultat.", "Négatifs : les pertes passées dépassent les réserves, le FPPN n'a pas de coussin propre."],
    ["Fonds propres et assimilés / dettes", DEFS["FEP"][0] + " Les provisions techniques comptent dans les fonds assimilés.", DEFS["FEP"][1]],
    ["Ratio de liquidité", DEFS["LIQ"][0], DEFS["LIQ"][1]],
    ["Encours des prêts du Grenier", "Le capital restant dû par les policiers emprunteurs.", "Il mesure le montant de l'épargne et des fonds du FPPN engagé dans le crédit."],
]
contenu({"titre": "Définition et lecture des indicateurs clés", "sous_titre": "Ce que chaque indicateur mesure et comment lire sa valeur",
         "synthese": "Les normes viennent de la Décision n° 05/2022 ; les repères des garanties sont propres à l'analyse d'Adenka"},
        "Définitions des indicateurs de la diapositive précédente et des ratios du résumé exécutif. Formules exactes des ratios réglementaires : partie 10.",
        tableau={"titre": tab("Définition et lecture des indicateurs clés"), "entetes": ["Indicateur", "Ce qu'il mesure", "Comment le lire"],
                 "largeurs": [6.4, 12.0, 12.8], "taille": 8, "sans_code": True, "lignes": _DEF_IND},
        valable="formules et normes en vigueur depuis le 01/06/2022",
        source="Source : " + DEC + " ; lecture Adenka")

contenu({"titre": "Fonctionnement du FPPN", "sous_titre": "Les cotisations et l'épargne financent prestations, prêts et investissements",
         "synthese": f"En 2025 : {F('Cotisations 2025', s[2025]['cot_maladie'] + s[2025]['cot_deces'])} de cotisations, {F('Prestations 2025', s[2025]['prest_maladie'] + s[2025]['prest_deces'])} de prestations"},
        "Ce schéma résume les flux d'argent. Les mutualistes cotisent par précompte sur salaire (Solde) ou sur pension (CGRAE). Le FPPN paie les prestations maladie (cliniques, pharmacies) et décès (familles), gère l'épargne (PEP, EFC), finance les prêts du Grenier et des investissements (CMP, terrains). Les montants sont ceux de la comptabilité 2025 et du bilan au 31/08/2026.",
        schema="fonctionnement",
        analyse=["**L'épargne des mutualistes est la ressource la plus importante du FPPN, devant les cotisations.** Elle atteint " + F("Épargne (schéma texte)", b[2026]["epargne_totale"]) + " au 31/08/2026, soit plus de quatre fois les cotisations de 2025."],
        implication="Le FPPN emploie cette épargne en partie dans des actifs peu liquides (CMP, terrains, prêts) : sa capacité à rembourser les épargnants est le premier enjeu de la mission.",
        source=SRC_COMPTA)

gab(23, {"titre": "Six constats majeurs", "sous_titre": "Les fragilités portent sur les données, les provisions, l'épargne et les normes",
         "synthese": "Chaque constat est démontré dans la partie indiquée, à partir des fichiers du FPPN et des textes de l'UEMOA"},
    {"blocs": [("Données incomplètes", "Le FPPN n'a transmis ni base de prestations, ni dates de naissance. Les analyses actuarielles centrales restent en attente (parties 3 et 14)."),
               ("Sinistralité en hausse", "Le S/C maladie passe de " + pc("S/C 2021 (6c)", s[2021]["sc_maladie"]) + " en 2021 à " + pc("S/C 2025 (6c)", s[2025]["sc_maladie"]) + " en 2025 (partie 5)."),
               ("Provisions figées", "Les provisions gardent le même montant depuis 2020, alors que les prestations ont été multipliées par " + N("Multiplicateur prestations", (s[2025]["prest_maladie"] + s[2025]["prest_deces"]) / (s[2021]["prest_maladie"] + s[2021]["prest_deces"]), 2) + " (partie 9)."),
               ("Épargne non adossée", "Des actifs réservés couvrent seulement " + pc("Couverture (6c)", b[2026]["couverture_epargne"]) + " de l'épargne due (partie 6)."),
               ("Fonds propres négatifs", "Les fonds propres valent " + F("FP 2025 (6c)", b[2025]["fonds_propres"]) + " fin 2025, et l'actif repose sur des emplois peu liquides (partie 9)."),
               ("Normes non respectées", f"{NB_NC} ratios de la Décision n° 05/2022 sur 36 ne sont pas respectés, dont la réserve obligatoire et la liquidité (partie 10).")]},
    "Les six constats sont classés du plus général (données) au plus réglementaire (normes). Les chiffres sont exacts et proviennent des balances comptables et des précomptes. Le constat « actifs peu liquides » repose sur la nature des comptes (investissements, terrains, créances) ; leur valeur réelle n'a pas été expertisée.",
    source="Source : analyses Adenka")

contenu({"titre": "Graphiques clés (1/2)", "sous_titre": "La sinistralité maladie monte pendant que l'épargne reste peu couverte",
         "synthese": f"S/C maladie 2025 : {pc('S/C 2025 (gc)', s[2025]['sc_maladie'])} ; couverture de l'épargne au 31/08/2026 : {pc('Couverture (gc)', b[2026]['couverture_epargne'])}"},
        "Graphique de gauche : rapport entre prestations et cotisations maladie comptabilisées, par exercice. Graphique de droite : épargne due aux mutualistes (passif) et actifs qui lui sont réservés. Les deux graphiques sont détaillés dans les parties 5 et 6.",
        graphiques=[{"titre": fig("Ratio S/C maladie (prestations / cotisations)"), "type": "lignes", "categories": cats_ex,
                     "series": [{"nom": "S/C maladie", "valeurs": [round(100 * s[a]["sc_maladie"], 2) for a in range(2020, 2027)], "couleur": BLEU}],
                     "format": "0.00", "legende": None, "unite": "En %"},
                    {"titre": fig("Épargne due et actifs réservés"), "type": "barres", "categories": yrs_bil,
                     "series": [{"nom": "Épargne PEP et EFC", "valeurs": [round(b[a]["epargne_totale"]) for a in range(2020, 2027)], "couleur": BLEU},
                                {"nom": "Actifs réservés", "valeurs": [round(b[a]["actifs_dedies_epargne"]) for a in range(2020, 2027)], "couleur": ORANGE}],
                     "format": FMT_M, "unite": "En millions de FCFA"}],
        bandeau="**Les deux courbes vont dans le mauvais sens pour la solvabilité.** Les prestations consomment une part croissante des cotisations, et l'épargne grossit plus vite que les actifs qui la garantissent. Implication : le FPPN dispose de moins en moins de marge pour faire face à un choc de prestations ou de retraits.",
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
        bandeau="**Le redressement des fonds propres s'interrompt en 2026, au moment où l'encours des prêts atteint son plus haut niveau.** Implication : le FPPN finance une activité de crédit en croissance rapide sans coussin de fonds propres pour absorber d'éventuelles pertes.",
        source=SRC_COMPTA + " ; " + SRC_GRENIER[9:])

cles = ["SIN", "MSE", "ROB", "FEP", "LIQ", "EQS", "CHF", "RL1"]
STATUT_TXT = {"Conforme": "Conforme", "Non conforme": "Non conforme", "En attente": "En attente", "À apprécier": "À apprécier"}


# ---- v12 : conformité prudentielle en bref, matrice de 2020 au 31/08/2026 (lecture du FPPN, puis lecture corrigée selon le PCMS)
cles = ["SIN", "MSE", "ROB", "FEP", "LIQ", "EQS", "CHF", "RL1"]


def _nom_court(c):
    return RT[c]["nom"].replace(" (risques longs)", "").replace("Ratio de solvabilité de la branche épargne", "Solvabilité de l'épargne")


def _norme_courte(c):
    return RT[c]["norme"].replace("Au moins 1 (marge disponible au moins égale à la marge exigée)", "Au moins 1").replace(" (tableau 1) ; 9 mois au § 1.2.2.5.4", "")


def _ligne_bref(c, corr=False):
    x = RT[c]
    sts = x.get("statuts_retraites" if corr else "statuts", {})
    vals = []
    for a in _AN:
        v_ = val_ratio(x, a, corr)
        vals.append(v_ if v_ == "n.d." else bv(v_, boule_statut(sts.get(a, ""))))
    if corr:
        glob = "Non conforme" if any(str(sts.get(a, "")).startswith("Non conforme") for a in range(2022, 2026)) else "Conforme"
        nom = _nom_court(c) + ", lecture corrigée selon le PCMS"
    else:
        glob, nom = x["statut_global"], _nom_court(c)
    return [nom, _norme_courte(c)] + vals + [glob]


_LB = [_ligne_bref(c) for c in cles] + [_ligne_bref(c, True) for c in ("MSE", "LIQ", "EQS")]
contenu({"titre": "Conformité prudentielle en bref", "sous_titre": "La réserve obligatoire, les fonds propres, la liquidité et la couverture de l'épargne ne sont respectées aucune année",
         "synthese": f"Sur 36 indicateurs de la Décision n° 05/2022 : {NB_NC} non conformes, {NB_C} conformes, {NB_ATT} en attente de données ou à apprécier"},
        "Le statut global est jugé sur les exercices 2022 à 2025, la Décision étant entrée en vigueur le 1er juin 2022. 2020 et 2021 sont indicatifs, comme les huit mois de 2026. Détail dans la partie 10.",
        tableau={"titre": tab("Huit ratios prudentiels clés, de 2020 au 31/08/2026"), "entetes": ["Ratio", "Norme"] + _ENT_AN + ["Statut 2022 à 2025"],
                 "largeurs": [7.0, 3.0] + [2.6] * 7 + [3.0], "statut_col": 9, "taille": 7, "lignes": _LB},
        disposition="dessous",
        encadre=["**Lecture du FPPN** : comptes tels qu'ils sont tenus. **Lecture corrigée selon le PCMS** : mêmes montants, avec l'épargne des adhérents (41195100, 41195200) en ressource longue et les avances durables (47111100 à 47111105, 47111201, 485) en emplois longs. Pour la marge, pertes reportées déduites et réévaluation (10620000) exclue."],
        encadre_titre="Deux lectures des comptes",
        analyse=["**Le classement de l'épargne change le verdict sur la liquidité, pas sur la solidité.** Corrigée selon le PCMS, la liquidité devient conforme en 2025 (" + val_ratio(RT["LIQ"], 2025, True) + " contre " + val_ratio(RT["LIQ"], 2025) + "). La réserve obligatoire, les fonds propres et la couverture de l'épargne restent non conformes."],
        valable="ratios de 2020 au 31/08/2026 (balance provisoire), statut global jugé de 2022 à 2025",
        source=SRC_REG)

# ---- v13 : matrice de conformité (domaines × textes applicables)
_TXT_C = ["Décision n° 05/2022", "Règlement d'exécution n° 004/2022", "Règlement n° 07/2009", "Statuts du FPPN", "Règlement intérieur"]
_MAT_CF = [
    ["Provisions", "—", "🔴 Art. 7 à 10 : provisions calculées dossier par dossier", "—", "🔴 Art. 89 : recalcul à chaque clôture\n🟢 Art. 88 : risques courts, aucune provision nécessaire", "—"],
    ["Cotisations et recouvrement", "🟢 Recouvrement des retraités CGRAE au moins égal à 90 %", "—", "—", "🔴 Art. 17 : 3,5 % et 1,5 %, la Solde retient 3,77 % et 1,23 %", "—"],
    ["Prestations", "⚪ Délai de paiement des prestataires", "—", "—", "—", "🟢 Art. 52 : paiement des cliniques en trois mois, 2,62 mois en 2025"],
    ["Épargne (risques longs)", "🔴 Cinq ratios des risques longs non conformes\n⚪ Deux à apprécier", "🔴 Art. 11 à 14 et 19 : provision mathématique, comptes certifiés, étude actuarielle, réserve de trésorerie", "—", "—", "—"],
    ["Placements et liquidité", "🔴 Liquidité\n⚪ Rentabilité des placements", "🔴 Art. 17 : placements sûrs, rentables, liquides, diversifiés", "—", "🟠 Art. 85 : délibération du conseil d'administration non reçue", "—"],
    ["Solvabilité et réserves", "🔴 Réserve obligatoire, fonds propres et assimilés\n🟢 Marge de solvabilité (lecture du FPPN)", "—", "🔴 Art. 84 : non-conformité sur deux exercices", "—", "—"],
    ["Gestion", "🔴 Trois ratios de frais\n🟢 Charges de personnel, formation", "—", "—", "—", "—"],
    ["Gouvernance et contrôle", "🟡 Réunions des organes, contrôles a posteriori", "🔴 Art. 27 : comptabilité séparée de la clinique", "🟠 Art. 63 : adhésion au fonds de garantie", "—", "—"],
]
contenu({"titre": "Matrice de conformité", "sous_titre": "Les exigences des textes sont surtout non respectées pour les provisions, l'épargne et la solidité financière",
         "synthese": "Domaines en lignes, textes applicables en colonnes ; « — » : le texte ne fixe pas de règle sur ce domaine"},
        "Chaque case donne le statut et l'exigence du texte. Les statuts reprennent ceux du corps du rapport : tableaux de bord (partie 10), grille de conformité, provisions (partie 9), cotisations (partie 4).",
        tableau={"titre": tab("Matrice de conformité : domaines et textes applicables"), "entetes": ["Domaine"] + _TXT_C,
                 "largeurs": [4.2, 5.6, 6.4, 4.4, 5.6, 5.0], "taille": 7.5, "sans_code": True, "legende": LEG_CONF, "lignes": _MAT_CF},
        disposition="dessous",
        valable="textes applicables au 31/08/2026 ; statuts établis sur les documents reçus jusqu'au 23/10/2026",
        source="Source : " + DEC + " ; " + REX + " ; " + R07 + " ; statuts et règlement intérieur du FPPN édités le 15/06/2023 ; analyses Adenka")

gab(107, {"titre": "Quatre taux à surveiller", "sous_titre": "Une faible part des engagements est couverte par des ressources disponibles",
          "synthese": "Couverture de l'épargne, sinistralité, trésorerie rapportée à l'épargne et prêts en retard de plus de 30 jours"},
    {"anneaux": [(b[2026]["couverture_epargne"], "Épargne couverte par des actifs réservés (31/08/2026)", "Couverture anneau"),
                 (s[2025]["sc_maladie"], "Sinistralité maladie (S/C 2025)", "S/C anneau"),
                 (C["retrait_max_couvert"]["par_tresorerie"], "Trésorerie totale / épargne due (31/08/2026)", "Trésorerie/épargne anneau"),
                 (IC["p30"], "Prêts en retard de plus de 30 jours (en nombre, 31/08/2026)", "Prêts retard anneau")]},
    "Chaque anneau donne un pourcentage. La trésorerie totale (toutes banques et caisses) ne représente qu'une petite part de l'épargne due. Les prêts en retard sont mesurés ici en nombre de prêts ; en montant d'encours, la part est plus faible (voir partie 7).",
    source=SRC_COMPTA + " ; " + SRC_GRENIER[9:],
    analyse="**Trois anneaux sur quatre montrent une couverture faible des engagements.** La trésorerie de toutes les banques ne rembourserait qu'environ un cinquième de l'épargne due. Implication : en cas de retraits importants, le FPPN doit céder ou mobiliser des actifs peu liquides.")

gab(83, {"titre": "Carte des risques", "sous_titre": "Quatre causes identifiées produisent quatre risques pour le FPPN",
         "synthese": "Les causes sont établies par les données ; le manque de liquidité et les pertes sur prêts sont chiffrés en partie 9"},
    {"causes": ["Pas de base de prestations ni de dates de naissance", "Provisions jamais recalculées", "Épargne utilisée pour d'autres emplois", "Écart entre échéances et précomptes"],
     "effets": ["Coût réel des sinistres inconnu", "Engagements mal mesurés au bilan", "Manque de liquidités en cas de retraits", "Retards et pertes sur les prêts"],
     "lecture": "Chaque colonne relie une cause (en haut) à l'effet qu'elle produit (en bas). Les quatre causes sont établies par les données. Deux effets sont chiffrés en partie 9 : le manque de liquidités (adossement) et les pertes sur les prêts (dépréciation du Grenier). Les deux autres, le coût réel des sinistres et la mesure des engagements par âge, attendent la base des prestations datée et les dates de naissance des mutualistes."},
    "Lecture : chaque colonne relie une cause (en haut) à son effet (en bas). Les causes sont démontrées par les données. Le manque de liquidité et la dépréciation des prêts sont chiffrés (partie 9). Le coût réel des sinistres et la mesure des engagements par âge attendent la base des prestations et les dates de naissance.",
    source="Source : analyses Adenka")

# ---- v13 : matrices du rapport (risques, criticité des constats, conformité, données et analyses)
ROUGE_M, JAUNE_M, VERT_M, ORANGE_M = "C00000", "FFD966", "70AD47", "EFAC62"


def _grav(montant):
    """Gravité d'un risque chiffré, selon son impact en FCFA."""
    m_ = abs(montant)
    return 1 if m_ < 50e6 else 2 if m_ < 500e6 else 3 if m_ < 2e9 else 4 if m_ <= 5e9 else 5


def _zone(p_, g_):
    sc = p_ * g_
    return ROUGE_M if sc >= 15 else JAUNE_M if sc >= 8 else VERT_M


_INT_AN = (INT55_25 - interets_dus(0.035, 0.055, 2024)) - next(x["produits_financiers"] for x in V11["rendement_gl"] if x["exercice"] == 2025)


def _risques():
    return [  # (code, risque, probabilité, justification, gravité, justification)
        ("R1", "Manque de liquidités face aux sorties d'épargne", 4, f"Sorties observées de {pc('RQ ts', T_MOY)} de l'épargne par an. L'actif à moins d'un an couvre {pc('RQ couv', couverture('31/08/2026', T_MOY)[2])} des paiements de l'année",
         _grav(GAP[4][5]), f"Manque cumulé de {F('RQ gap', -GAP[4][5])} à cinq ans"),
        ("R2", "Retrait d'agrément ou sanctions de l'AIRMS", 4, f"{NB_NC} indicateurs non conformes deux exercices de suite : la condition de l'art. 84 du règlement n° 07/2009 est remplie, la sanction n'est pas automatique",
         5, "Arrêt de l'activité du FPPN"),
        ("R3", "Dette d'épargne sous-évaluée", 5, "Constaté : aucun intérêt comptabilisé de 2020 au 31/08/2026",
         _grav(INT55), f"Intérêts dus non comptabilisés : {F('RQ int', INT55)}"),
        ("R4", "Provisions sans lien avec les risques", 5, "Constaté : montants identiques depuis le 01/09/2020",
         _grav(PROV_MAL + PROV_OBS - _BESOIN), f"Provisions supérieures au besoin mesuré de {F('RQ exc', PROV_MAL + PROV_OBS - _BESOIN)}"),
        ("R5", "Déficit de la garantie maladie", 5, f"Constaté : ratio combiné de {pc('RQ rc25', _RGm[(2025, 'Maladie')]['ratio_combine'])} en 2025 et {pc('RQ rc26', _RGm[(2026, 'Maladie')]['ratio_combine'])} sur huit mois de 2026",
         _grav(_RGm[(2026, "Maladie")]["resultat_technique"]), f"Résultat technique maladie de {F('RQ rt26', _RGm[(2026, 'Maladie')]['resultat_technique'])} sur huit mois de 2026"),
        ("R6", "Absence de coussin de fonds propres", 5, "Constaté : fonds propres négatifs à chaque clôture depuis 2020",
         _grav(b[2026]["fonds_propres"]), f"Fonds propres de {F('RQ fp', b[2026]['fonds_propres'])} au 31/08/2026"),
        ("R7", "Pertes sur les prêts du Grenier", 5, "Constaté : prêts en retard au 31/08/2026, aucune dépréciation passée",
         _grav(_GR26), f"Dépréciation nécessaire de {F('RQ dep', _GR26)}"),
        ("R8", "Produits financiers inférieurs aux intérêts dus", 5, f"Constaté : rendement moyen de {pc('RQ rdm', V11['rendement_moyen_2021_2025'])} par an de 2021 à 2025",
         _grav(_INT_AN), f"Intérêts dus de 2025 supérieurs aux produits financiers de 2025 de {F('RQ int an', _INT_AN)}"),
        ("R9", "Engagements impossibles à mesurer faute de données", 5, "Constaté : ni base des prestations datée, ni dates de naissance des mutualistes",
         3, "Non chiffrable : bloque les provisions dossier par dossier, la mortalité, la santé des retraités et la projection"),
        ("R10", "Placements décidés sans délibération des organes", 4, "[[Le directeur général décide seul des placements]]. Aucune délibération reçue",
         _grav(500000000), f"Placement dans le FCP de {F('RQ fcp', 500000000)} en 2025"),
    ]


# les justifications chiffrées s'affichent sur la diapositive qui suit la matrice
COURANTE = len(SL) + 2
RISQUES = _risques()
COURANTE = None
_ETIQ_P = {5: "5 - Très probable", 4: "4 - Probable", 3: "3 - Possible", 2: "2 - Peu probable", 1: "1 - Très peu probable"}
_ETIQ_G = ["1 - Négligeable", "2 - Mineure", "3 - Modérée", "4 - Majeure", "5 - Catastrophique"]
_MAT_R, _FONDS_R = [], {}
for i_, p_ in enumerate((5, 4, 3, 2, 1)):
    ligne_ = [_ETIQ_P[p_]]
    for g_ in range(1, 6):
        ligne_.append(", ".join(c_ for c_, _, pp_, _, gg_, _ in RISQUES if pp_ == p_ and gg_ == g_))
        _FONDS_R[(i_, g_)] = _zone(p_, g_)
    _MAT_R.append(ligne_)
LEG_RISQ = [("🔴", "Élevé", "score de 15 à 25, traitement immédiat"), ("🟡", "Moyen", "score de 8 à 12, plan d'action"), ("🟢", "Faible", "score de 1 à 6, surveillance")]
contenu({"titre": "Matrice des risques", "sous_titre": f"{sum(1 for r in RISQUES if r[2] * r[4] >= 15)} risques sur {len(RISQUES)} se situent dans la zone rouge",
         "synthese": "Probabilité en lignes, gravité en colonnes ; chaque risque est justifié sur la diapositive suivante"},
        "Score = probabilité × gravité. Zone rouge : score de 15 à 25. Zone jaune : de 8 à 12. Zone verte : de 1 à 6. La gravité d'un risque chiffré suit son impact en FCFA ; la probabilité suit ce qui est constaté dans les comptes et les fichiers.",
        tableau={"titre": tab("Matrice des risques : probabilité et gravité"), "entetes": ["Probabilité \\ Gravité"] + _ETIQ_G,
                 "largeurs": [5.2] * 6, "taille": 10, "sans_code": True, "hauteur_ligne": 1.15, "fonds": _FONDS_R, "centre": True, "legende": LEG_RISQ, "lignes": _MAT_R},
        disposition="dessous",
        encadre=["**Gravité** (impact en FCFA) : 1 négligeable, moins de 50 000 000 FCFA. 2 mineure, de 50 000 000 à 500 000 000 FCFA. 3 modérée, de 500 000 000 à 2 000 000 000 FCFA. 4 majeure, de 2 000 000 000 à 5 000 000 000 FCFA. 5 catastrophique, plus de 5 000 000 000 FCFA ou arrêt de l'activité.",
                 "**Probabilité** : 5 très probable, déjà constaté dans les comptes ou les fichiers. 4 probable, conditions réunies mais risque non encore réalisé. 3 possible. 2 peu probable. 1 très peu probable."],
        encadre_titre="Échelles",
        valable="comptes et fichiers de 2020 au 31/08/2026 ; impacts chiffrés dans le corps du rapport",
        source="Source : analyses Adenka, parties 5 à 10 du rapport")
contenu({"titre": "Matrice des risques : justification", "sous_titre": "Chaque note de probabilité et de gravité repose sur un chiffre du rapport",
         "synthese": "Les risques constatés sont notés 5 en probabilité ; la gravité suit l'impact chiffré en FCFA"},
        "Justification de la place de chaque risque dans la matrice. Niveau : 🔴 élevé (score de 15 à 25), 🟡 moyen (8 à 12), 🟢 faible (1 à 6).",
        tableau={"titre": tab("Risques, probabilité, gravité et score"), "entetes": ["N°", "Risque", "Probabilité et justification", "Gravité et justification", "Score"],
                 "largeurs": [1.2, 6.4, 10.2, 10.2, 3.2], "taille": 7.5, "sans_code": True, "legende": LEG_RISQ,
                 "lignes": [[c_, r_, f"{p_}. {jp_}", f"{g_}. {jg_}", bv(str(p_ * g_), "🔴" if p_ * g_ >= 15 else "🟡" if p_ * g_ >= 8 else "🟢")] for c_, r_, p_, jp_, g_, jg_ in RISQUES]},
        disposition="dessous",
        valable="comptes et fichiers de 2020 au 31/08/2026",
        source="Source : analyses Adenka ; chiffres des parties 5 à 10 du rapport")

# ---- v11 : recommandations au format du rapport de diagnostic : recommandation, actions concrètes, priorité et horizon
REC1 = [
    ["1. Calculer les provisions maladie et décès à chaque clôture",
     "• Le service comptable établit chaque année une note de calcul des provisions : méthode, fichiers utilisés, montant de chaque provision, comparaison au bilan\n"
     "• Provision maladie : soins de l'année facturés après la clôture, mesurés sur le journal SIN de janvier à mars\n"
     "• Provision décès : décès survenus et déclarés après la clôture. Les dossiers connus restent en dette au compte 41591000\n"
     "• Le directeur général signe la note. Elle est présentée à l'Assemblée générale avec les comptes (statuts, art. 85 et 89)\n"
     "• Les montants hérités (" + F("RC hérités", PROV_MAL + PROV_OBS) + ") sont repris après décision des organes et information de l'AIRMS",
     "Court terme (moins de 6 mois)"],
    ["2. Arrêter le taux servi sur l'épargne et tenir les comptes individuels",
     "• Le conseil d'administration fixe par délibération le taux servi sur le PEP et sur l'EFC\n"
     "• Le FPPN tient pour chaque épargnant un relevé annuel : solde d'ouverture, versements, intérêts, retraits, solde de clôture\n"
     "• Les intérêts dus (" + F("RC int", INT55) + " au 31/08/2026) sont comptabilisés",
     "Court terme (moins de 6 mois)"],
    ["3. Transmettre les données qui manquent aux calculs actuariels",
     f"• Base des prestations datée : date des soins, du dépôt et du paiement de chaque facture (PBC n° {PBCM[71]})\n"
     f"• Fichier des mutualistes avec date de naissance et sexe, fichier des ayants droit (PBC n° {PBCM[9]} et {PBCM[66]})\n"
     f"• Fichiers retour de décembre 2024 et mois manquants du décès (PBC n° {PBCM[62]})",
     "Court terme (moins de 6 mois)"],
    ["4. Présenter à l'AIRMS un plan de mise en conformité prudentielle",
     "• Objectifs chiffrés par ratio non conforme, à partir des montants à combler de la partie 10\n"
     "• Calendrier sur trois exercices, validé par le conseil d'administration\n"
     "• Suivi semestriel des ratios par le service comptable",
     "Court terme (moins de 6 mois)"],
    ["5. Constituer une réserve de liquidité et adopter une politique de placement",
     "• Réserve égale aux sorties d'épargne d'une année au taux observé (" + pc("RC tmoy", T_MOY) + " de l'épargne)\n"
     "• Politique de placement votée par le conseil d'administration : part liquide minimale, rendement minimal au regard du taux servi, limites sur l'immobilier\n"
     "• Placement à terme d'une partie du solde des comptes courants non rémunérés",
     "Moyen terme (6 à 18 mois)"],
]
REC2 = [
    ["6. Rééquilibrer la garantie maladie",
     "• Tarification fondée sur le coût complet de la garantie, frais compris\n"
     "• Revue du barème des prestations et des conventions avec les cliniques\n"
     "• Contrôle a posteriori d'au moins 5 % des factures (Décision n° 05/2022)",
     "Moyen terme (6 à 18 mois)"],
    ["7. Encadrer les investissements dans la clinique, le Grenier et le projet de Cocody",
     "• Comptabilité séparée de la clinique (règlement d'exécution n° 004/2022, art. 27)\n"
     "• Budget, financement et effet sur la liquidité du projet de Cocody présentés aux organes avant tout engagement\n"
     "• Dépréciation annuelle des prêts du Grenier selon sa politique de crédit",
     "Moyen terme (6 à 18 mois)"],
    ["8. Faire réaliser l'étude actuarielle des risques longs et certifier les comptes",
     "• Étude actuarielle de l'EFC et du PEP tous les cinq ans (règlement d'exécution n° 004/2022, art. 14)\n"
     "• Certification des comptes par un commissaire aux comptes et états financiers par risque long (art. 12 et 13)",
     "Long terme (plus de 18 mois)"],
    ["9. Mettre en place un dispositif actuariel permanent",
     "• Rapprochement mensuel des fichiers aller et retour, état des rejets\n"
     "• États de fin d'année des prestations à payer, tableau de bord des ratios de la Décision n° 05/2022\n"
     "• Procédure écrite pour chaque calcul, avec un responsable et un contrôle (partie 13)",
     "Long terme (plus de 18 mois)"],
]
for _k, _lig in enumerate((REC1, REC2), 1):
    contenu({"titre": f"Recommandations ({_k}/2)", "sous_titre": "Quatre actions à engager dans les six mois conditionnent toutes les autres" if _k == 1 else "Les actions de moyen et long terme corrigent l'emploi des ressources et la gouvernance",
             "synthese": "Chaque recommandation répond à un constat de la synthèse et précise qui agit, comment et dans quel délai"},
            "Format du rapport de diagnostic : recommandation, actions concrètes, priorité et horizon. Les numéros de la PBC renvoient à la liste des documents demandés du 23/10/2026.",
            tableau={"titre": tab(f"Recommandations et actions ({_k}/2)"), "entetes": ["Recommandation", "Actions concrètes", "Priorité et horizon"],
                     "largeurs": [7.2, 19.8, 4.2], "taille": 8.5, "lignes": _lig},
            disposition="dessous",
            valable="constats du rapport au 31/08/2026 ; horizons comptés à partir de la remise du rapport",
            source="Source : analyses Adenka ; statuts du FPPN, art. 85 et 89 ; " + REX + ", art. 12 à 14, 17 et 27 ; " + DEC)

gab(39, {"titre": "Feuille de route proposée", "sous_titre": "Quatre étapes pour passer du constat à l'évaluation des engagements",
         "synthese": "Chaque étape dépend de la précédente ; la première est la transmission des données manquantes"},
    {"etapes": [("Collecter", "Le FPPN transmet les bases de prestations, les mutualistes avec date de naissance, les comptes d'épargne et les barèmes."),
                ("Calculer", "Adenka précise la provision maladie (factures datées), la provision décès (dossiers du 41591000), la mortalité observée et les engagements par épargnant."),
                ("Vérifier", "Adenka rapproche les résultats de la comptabilité et des ratios de l'UEMOA, puis les discute avec le FPPN."),
                ("Projeter", "Adenka projette l'équilibre technique sur 5 à 10 ans et chiffre les recommandations.")]},
    "La feuille de route ne fixe pas de dates : elles dépendent du délai de transmission des données par le FPPN. Chaque étape produit une mise à jour de ce rapport.",
    source="Source : Adenka")

# ---------------------------------------------------------------- PARTIE 1 : COMPRENDRE LE FPPN
separateur(1, "Comprendre le FPPN", "Cadre, garanties, circuits et notions de base",
           "Cette partie donne au lecteur les clés pour comprendre la suite, sans connaissance préalable en actuariat.")
contenu({"titre": "Cadre juridique", "sous_titre": "Le FPPN est une mutuelle sociale agréée, régie par le droit de l'UEMOA",
         "synthese": "Le règlement n° 07/2009/CM/UEMOA fixe les règles, l'AIRMS contrôle, le Code CIMA sert de référence technique"},
        "Une mutuelle sociale n'est pas une compagnie d'assurance : elle est régie par le règlement UEMOA de 2009, pas par le Code CIMA. Le Code CIMA sert ici seulement de référence technique (tables de mortalité, taux d'actualisation), faute de règle équivalente dans les textes de la mutualité. Le répertoire de l'AIRMS mis à jour le 15/10/2026 cite le FPPN au n° 05 des mutuelles agréées et immatriculées ; le numéro et la date de son agrément restent à obtenir.",
        tableau={"titre": tab("Textes et organismes de référence"), "entetes": ["Texte ou organisme", "Rôle pour le FPPN", "Usage dans ce rapport"], "largeurs": [8.0, 11.4, 11.8],
                 "lignes": [["Règlement n° 07/2009/CM/UEMOA du 26/06/2009", "Il fixe l'agrément, les obligations, les sanctions et le contrôle des mutuelles sociales.", "Cadre général (partie 10)"],
                            ["Règlement d'exécution n° 004/2022/COM/UEMOA du 01/06/2022", "Il fixe les règles prudentielles : provisions, réserves, risques longs, placements.", "Exigences prudentielles (partie 10)"],
                            ["Décision n° 05/2022/COM/UEMOA du 01/06/2022", "Elle fixe les indicateurs et ratios prudentiels et leurs normes.", "Tableau de bord des ratios (partie 10)"],
                            ["Règlement n° 03/2012/CM/UEMOA et PCMS", "Ils imposent le plan comptable des mutuelles sociales depuis l'exercice 2014.", "Lecture et contrôle des comptes"],
                            ["AIRMS", "L'Agence ivoirienne de régulation de la mutualité sociale agrée et contrôle les mutuelles.", "Destinataire des états et des ratios"],
                            ["Code CIMA, article 338", "Il fixe les tables de mortalité et le taux d'intérêt maximal (3,5 %) des assureurs vie.", "Référence technique (partie 11)"]]},
        disposition="dessous",
        analyse=["**Le FPPN doit respecter des normes chiffrées depuis le 1er juin 2022.** Avant cette date, les textes ne fixaient pas de ratios prudentiels précis pour les mutuelles."],
        implication="Les ratios de la Décision n° 05/2022 s'appliquent pleinement aux exercices 2022 à 2025 de la période d'analyse : la partie 10 les calcule un par un.",
        source="Source : textes cités (documentation publique de l'AIRMS fournie le 18/10/2026) ; répertoire AIRMS des mutuelles agréées, mise à jour du 15/10/2026")

# ---- v8 : statuts et règlement intérieur du FPPN (édités le 15/06/2023), lus en entier le 19/10/2026
SRC_STAT = "Source : statuts et règlement intérieur du FPPN édités le 15/06/2023 (textes adoptés par l'Assemblée générale extraordinaire du 30/06/2021)"
contenu({"titre": "Statuts et règlement intérieur : cotisations, provisions, prestations", "sous_titre": "Les textes du FPPN fixent déjà des règles que les comptes ne suivent pas",
         "synthese": "Barème de cotisation, calcul des provisions et dotation de la réserve légale sont écrits dans les statuts"},
        "Lecture intégrale des statuts (43 pages) et du règlement intérieur (45 pages) transmis le 19/10/2026. Le tableau ne retient que les dispositions utiles aux tâches actuarielles. Les deux textes renvoient, pour les montants des prestations, à un « cahier des charges du régime des prestations sociales » et à un « tableau des allocations », reçus le 22/10/2026 (« CAHIER DE CHARGE  MALADIES ET INVAL-DECES 2021.pdf », « CARTOGRAPHIE DES RISQUES COUVERTS ET DES PRESTATIONS OFFERTES.docx »).",
        tableau={"titre": tab("Dispositions des textes du FPPN utiles à la mission (1/2)"), "entetes": ["Disposition", "Article", "Constat dans les données", "Effet sur la mission"], "largeurs": [9.0, 4.4, 8.6, 9.2], "taille": 8.5,
                 "lignes": [["Cotisation obligatoire de 5 % de la solde de base : 3,5 % maladie-maternité et 1,5 % invalidité-décès", "Statuts, art. 17 ; règlement intérieur, art. 33", "La Solde retient 3,77 % et 1,23 % (partie 4)", "Le barème officiel est établi. [[La Solde applique une autre répartition, sans document de notification reçu]] (PBC, item 56)."],
                            ["Provision pour prestations à payer calculée exercice par exercice, dossier par dossier, plus les sinistres survenus non déclarés ; provision pour risques courts", "Statuts, art. 86 à 89 ; règlement intérieur, art. 79 à 82", "Provisions identiques de 2020 au 31/08/2026 (partie 9)", "Le FPPN ne suit pas ses propres statuts, en plus du règlement d'exécution n° 004/2022 (art. 10)."],
                            ["Réserve légale d'au moins 20 % à 50 % des excédents nets de l'exercice précédent", "Statuts, art. 90", "Dotation 2023 : " + F("S8 dotation 2023", AFF[2023]["dotation_reserve_legale"]) + ", soit " + pc("S8 taux dotation 2023", AFF[2023]["taux_dotation_legale"]) + " de l'excédent 2022", "Écart aux statuts en 2023 ; dotations conformes en 2024 (" + pc("S8 taux dotation 2024", AFF[2024]["taux_dotation_legale"]) + ") et 2025 (" + pc("S8 taux dotation 2025", AFF[2025]["taux_dotation_legale"]) + ")."],
                            ["Couverture maladie limitée aux soins liés à une hospitalisation ; exclusions dont l'optique, les prothèses dentaires et les maladies de longue durée", "Règlement intérieur, art. 41 et 45", "Sinistralité inférieure à 70 % de 2020 à 2025 (partie 5)", "Explique en partie la faible sinistralité (facteur F5, partie 10)."],
                            ["Factures payées sous 3 mois ; remboursements demandés sous 3 mois et payés sous 6 mois ; forclusion des allocations décès à 3 ans (membre) et 12 mois (ayant droit)", "Règlement intérieur, art. 52, 57, 58 et 60", "Aucune base de prestations reçue", "Horizon des sinistres tardifs : moins d'un an en maladie, trois ans en décès, dans les quatre années de la circulaire CIMA n° 00230/2005."]]},
        disposition="dessous",
        analyse=["**Les statuts imposent eux-mêmes le calcul des provisions et la dotation de la réserve légale.** Les écarts relevés en parties 9 et 10 sont donc aussi des écarts aux textes internes du FPPN."],
        implication="Les délais de forclusion bornent la durée de développement des sinistres : ils serviront à construire les cadences de la méthode CIMA dès réception des bases de prestations (PBC, item 71).",
        source=SRC_STAT + " ; circulaire n° 00230/CIMA/CRCA/PDT/2005 ; balances du FPPN ; calculs Adenka")
contenu({"titre": "Statuts et règlement intérieur : épargne, personnel, organes", "sous_titre": "Trois dispositions changent le périmètre des engagements à évaluer",
         "synthese": "L'EFC est un engagement de carrière, la provision du personnel ne vise que les civils, et la situation des retraités est contradictoire"},
        "Suite du tableau précédent. Le personnel civil de la Direction Générale perçoit un salaire ; les fonctionnaires de police en service à la Direction Générale perçoivent des indemnités (règlement intérieur, art. 31). L'indemnité de départ du décret n° 2017-210 vise les salariés ; la retraite des fonctionnaires relève de l'État (décret n° 2012-652 : 60 ou 65 ans selon le grade).",
        tableau={"titre": tab("Dispositions des textes du FPPN utiles à la mission (2/2)"), "entetes": ["Disposition", "Article", "Constat dans les données", "Effet sur la mission"], "largeurs": [9.0, 4.4, 8.6, 9.2], "taille": 8.5,
                 "lignes": [["Plan Épargne Plus : souscription d'au moins 3 ans, terme choisi par le membre, 10 000 FCFA par mois au minimum ; il ouvre droit à des prêts", "Règlement intérieur, art. 65 à 69", "Prestations et rémunération renvoyées au cahier des charges", "Les sorties du PEP dépendent des termes choisis : l'échéancier individuel est nécessaire (PBC, item 102)."],
                            ["Entraide de Fin de Carrière : risques retraite et décès, 10 000 FCFA par mois sur toute la carrière", "Règlement intérieur, art. 70 et 71", "Épargne EFC au bilan, sans échéancier", "Engagement long, exigible au départ à la retraite : sa valeur demande le cahier des charges (PBC, item 100)."],
                            ["Personnel de la Direction Générale : civils salariés et fonctionnaires de police payés par des indemnités", "Règlement intérieur, art. 31 ; statuts, art. 66", "Provision pour départ à la retraite au compte 196", "La provision vise a priori le seul personnel civil, au barème du décret n° 2017-210 (PBC, items 103 et 104)."],
                            ["Policier retraité membre participant, mais carte restituée au départ à la retraite ; seule la cotisation du policier en activité est prévue", "Statuts, art. 7 et 80 ; règlement intérieur, art. 38", "Garantie santé des retraités non documentée", "Contradiction à lever avant tout calcul de provision mathématique (PBC, item 108)."],
                            ["Placements autorisés par le Conseil d'administration ; méthodes comptables immuables sauf approbation de l'Assemblée générale", "Statuts, art. 75 et 85", "Aucune délibération transmise ; provisions reclassées en 2024", "Délibérations et procès-verbaux demandés (PBC, items 105 et 106)."]]},
        disposition="dessous",
        analyse=["**Le périmètre des engagements dépend de documents que les statuts citent mais que le FPPN n'a pas transmis.** Le cahier des charges fixe les prestations de l'EFC et du PEP ; le fichier du personnel doit distinguer civils et policiers."],
        implication="Ces dix nouvelles demandes forment la PBC complémentaire « textes et personnel » (items 100 à 109), jointe à la PBC générale.",
        source=SRC_STAT + " ; décret n° 2017-210 du 30/03/2017 ; décret n° 2012-652")

contenu({"titre": "Garanties et produits", "sous_titre": "Huit postes de précompte portent l'ensemble des produits du FPPN",
         "synthese": "Maladie et décès sont les garanties de base ; épargne, prêts, flotte et MEDOC sont des produits complémentaires"},
        "Chaque produit a un code de poste sur la fiche de paie (Solde) : 936 maladie, 8Z6 décès, 870 EFC, 998 PEP, 9X4 Grenier, 9W3 flotte, 8Z7 MEDOC. Les retraités sont précomptés par la CGRAE sur leur pension. Les montants moyens sont observés dans les fichiers de précomptes. La cotisation des retraités suit le barème de la CGRAE reçu le 23/10/2026.",
        tableau={"titre": tab("Produits du FPPN et montants observés"), "entetes": ["Poste", "Produit", "Nature", "Montant moyen observé"], "largeurs": [2.2, 5.4, 7.8, 5.2],
                 "lignes": [["936", "Maladie", "Garantie santé des actifs, 3,5 % de la solde de base (statuts, art. 17), 3,77 % retenu par la Solde", F("Moyen 936 août 2026", [e for e in EFF if e["annee"] == 2026 and e["mois"] == 8][0]["montant"] / [e for e in EFF if e["annee"] == 2026 and e["mois"] == 8][0]["cotisants"]) + " par mois (août 2026)"],
                            ["8Z6", "Décès", "Capital décès, frais funéraires, allocations, 1,5 % (statuts), 1,23 % retenu par la Solde", "voir partie 4"],
                            ["870", "EFC", "Épargne versée en fin de carrière", F("Moyen EFC", an(R["efc_versements"], 2025)["versements"] / an(R["efc_versements"], 2025)["epargnants"] / an(R["efc_versements"], 2025)["mois"]) + " par mois (2025)"],
                            ["998", "PEP", "Épargne volontaire", "voir partie 6"],
                            ["9X4", "Grenier du Policier", "Remboursement de prêts", "voir partie 7"],
                            ["9W3", "Flotte automobile", "Assurance des véhicules, assureur GNA", "environ 10 000 FCFA"],
                            ["8Z7", "MEDOC", "[[Complément pharmacie, le policier paie 20 % (déclaration du 19/10/2026)]]", "5 000 FCFA"],
                            ["CGRAE", "Retraités", "Garantie santé, précompte sur pension : 2,5 % de (pension de base - MUGEFCI) + 5 000 FCFA (sous-officier), 8 000 FCFA (officier), 10 000 FCFA (commissaire de 2e classe à principal) ou 15 000 FCFA (commissaire divisionnaire à général)", F("Moyen CGRAE", sum(c["moyen"] for c in CG) / len(CG))]]},
        disposition="droite",
        analyse=["**Le FPPN couvre à la fois des risques courts et des risques longs.** La maladie se règle dans l'année ; l'EFC et le PEP sont versés des années après les cotisations.",
                 "**Le barème tarifaire du FPPN fixe, pour chaque acte de la clinique (CMP), la part du FPPN à 80 % du tarif et celle du mutualiste à 20 %, ainsi qu'un plafond annuel de 1 500 000 FCFA par bénéficiaire et de 3 000 000 FCFA par famille.** Il donne aussi les tarifs proposés par le FPPN aux structures privées. [[Le document porte la mention « tarifications proposées » et n'est pas daté : sa date d'application reste à établir.]]",
                 "**MEDOC est le seul produit sans aucun document.** L'EFC et le PEP sont décrits par le cahier des charges de 2021 et le manuel des produits d'épargne, sans décision d'adoption signée (partie 6)."],
        implication="Les risques longs (EFC, PEP) relèvent de règles prudentielles plus strictes : étude actuarielle tous les cinq ans, certification des comptes, réserve de trésorerie (partie 10).",
        source=SRC_PREC + " ; « Bar_me_tarif_du_FPPN.pdf » et « BAREME_CGRAE.xlsx » (reçus le 23/10/2026)")

contenu({"titre": "Circuit des précomptes", "sous_titre": "La maladie et le décès suivent un circuit direct que le FPPN ne maîtrise pas",
         "synthese": "Les produits complémentaires, les retraités et la flotte passent par une plateforme de partage, la maladie et le décès non"},
        "Logigramme en couloirs : chaque couloir est un acteur. Rectangle : étape ; losange : décision ; forme ondulée : fichier ; cadre rouge : point de contrôle manquant. Le circuit est celui décrit par le responsable des cotisations le 19/10/2026 ; la procédure écrite est demandée au FPPN. Les formes sont modifiables dans PowerPoint.",
        logigramme="circuit",
        analyse=["**Deux circuits coexistent, et le logigramme du FPPN les confirme.** Pour la maladie et le décès, le service du personnel crée le fichier de précompte et calcule les cotisations, le Trésor prélève, et le FPPN reçoit le fichier retour, sur papier puis en numérique. Sans carte de santé FPPN, l'adhérent n'a pas de prestation. Pour l'EFC et le PEP, le service informatique du FPPN intègre la souscription dans la plateforme Easy money et transmet le fichier aller. Pour la CGRAE, le prélèvement porte sur la pension et la CGRAE vire les fonds selon son barème (" + ENT19 + " ; " + FLOW + ")."],
        implication="Pour la maladie et le décès, le FPPN ne dispose pas de ce qui aurait dû être précompté : le taux de recouvrement exigé par la Décision n° 05/2022 (au moins 90 %) ne peut pas être calculé sans les fichiers de la direction du personnel. Pour les autres produits, le FPPN détient les deux fichiers mais ne les rapproche pas.",
        source=SRC_ENT19 + " ; " + FLOW + " ; fichiers aller et retour reçus le 15/10/2026")
contenu({"titre": "Circuit des fichiers par produit", "sous_titre": "Le FPPN ne produit pas les fichiers aller de la maladie et du décès",
         "synthese": "Quatre circuits, dont trois passent par la plateforme de partage"},
        "Le tableau reprend les circuits décrits le 19/10/2026. Le fichier aller contient le mécano, le nom, le code du poste, une référence et le montant ; la Solde retrouve le mutualiste par son mécano.",
        tableau={"titre": tab("Circuit des fichiers de précomptes par produit"), "entetes": ["Produits", "Fichier aller", "Destinataire", "Fichier retour", "Plateforme de partage"], "largeurs": [7.0, 8.6, 3.4, 8.0, 4.2],
                 "lignes": [["Maladie (936) et décès (8Z6)", "Transmis directement par la direction du personnel de la Police", "Solde", "Reçu par le FPPN : « …_effectif_police_936_….xlsx » et « FPPN_ALLOC_DECES_8Z6_….xlsx »", "Non"],
                            ["EFC (870), PEP (998), Grenier (9X4), MEDOC (8Z7)", "Produit et déposé par le responsable des cotisations", "Solde", "Déposé par la Solde", "Oui"],
                            ["Retraités", "Produit et déposé par le responsable des cotisations", "CGRAE", "Déposé par la CGRAE", "Oui"],
                            ["Flotte automobile (9W3)", "Produit et déposé par le responsable des cotisations", "GNA", "Déposé par GNA", "Oui"]]},
        disposition="dessous",
        analyse=["**Les fichiers aller maladie et décès sont absents de tous les envois reçus.** [[Le FPPN ne les produit pas]] (" + ENT19 + ")."],
        implication="Les fichiers aller maladie et décès sont demandés à la direction du personnel de la Police ou à la Solde (PBC, item 61). Les retours vides ou manquants (décembre 2024, décès de mai à octobre 2024 et d'avril 2025) sont demandés au FPPN (PBC, item 62).",
        source=SRC_ENT19)

# ---- v13 : circuits des cotisations par produit, d'après les logigrammes du FPPN, complétés des contrôles absents (cadre rouge)
_SRC_FLOW = "Source : « FLOW_CHART_COTISATION_ET_PRODUITS_DEPARGNE_1.pptx » (FPPN, reçu le 23/10/2026) ; contrôles absents : analyses Adenka"
contenu({"titre": "Circuit EFC et PEP : du bulletin à la comptabilisation", "sous_titre": "Le FPPN produit le fichier de l'épargne, mais ne rapproche pas l'aller et le retour",
         "synthese": "Logigramme du FPPN redessiné ; les cadres rouges signalent les contrôles absents constatés par Adenka"},
        "Couloirs repris du logigramme du FPPN. Forme arrondie bleue : début ou fin. Rectangle : étape. Losange : décision. Forme ondulée : document. Cadre rouge : contrôle absent.",
        logigramme="epargne",
        analyse=["**Le circuit s'arrête à la comptabilisation des précomptes.** Aucune étape ne rapproche le fichier aller du fichier retour, aucune ne calcule les intérêts dus ni ne produit de relevé par épargnant (" + F("CE int", INT55) + " d'intérêts dus au 31/08/2026)."],
        source=_SRC_FLOW + " ; " + SRC_ENT19[9:])
contenu({"titre": "Circuit de la flotte automobile (GNA)", "sous_titre": "Le FPPN encaisse une commission de 20 % dont l'assiette n'est pas documentée",
         "synthese": "Logigramme du FPPN redessiné ; les cadres rouges signalent les contrôles absents constatés par Adenka"},
        "Couloirs repris du logigramme du FPPN. Cadre rouge : contrôle absent ou fichier manquant.",
        logigramme="gna",
        analyse=["**Le FPPN collecte les primes pour le GNA et gère les trop-perçus, mais la convention d'apporteur d'affaires n'a pas été transmise.** L'assiette de la commission de 20 % ne peut pas être vérifiée, et le fichier retour de janvier 2026 manque."],
        source=_SRC_FLOW)
contenu({"titre": "Circuit des retraités (CGRAE)", "sous_titre": "La cotisation suit le barème de la CGRAE, dans la limite de la quotité cessible",
         "synthese": "Logigramme du FPPN redessiné ; les cadres rouges signalent les contrôles absents constatés par Adenka"},
        "Couloirs repris du logigramme du FPPN. Barème de la CGRAE : 2,5 % de (pension de base - MUGEFCI) + 5 000 à 15 000 FCFA selon le grade.",
        logigramme="cgrae",
        analyse=["**Seuls les retraités qui disposent d'une quotité cessible sont précomptés : la liste des autres souscripteurs n'a pas été transmise.** Les fichiers retour de janvier et de juin 2024 manquent."],
        source=_SRC_FLOW + " ; « BAREME_CGRAE.xlsx » ; CONVENTION DE PARTENARIAT ENTRE IPS-CGRAE ET LE FPPN.pdf")
contenu({"titre": "Circuit de la maladie et du décès", "sous_titre": "Le fichier aller est produit hors du FPPN : le recouvrement ne peut pas être mesuré",
         "synthese": "Logigramme du FPPN redessiné ; les cadres rouges signalent les contrôles absents constatés par Adenka"},
        "Couloirs repris du logigramme du FPPN. Sans carte de santé FPPN, l'adhérent n'a pas de prestation.",
        logigramme="maladie_deces",
        analyse=["**Le service du personnel calcule et transmet le fichier aller sans le communiquer au FPPN.** Le FPPN ne peut ni vérifier les taux retenus (3,77 % et 1,23 % au lieu de 3,5 % et 1,5 %), ni mesurer le taux de recouvrement exigé par la Décision n° 05/2022 (au moins 90 %)."],
        source=_SRC_FLOW + " ; statuts du FPPN édités le 15/06/2023, art. 17 ; fichiers de précomptes 936 et 8Z6")

contenu({"titre": "Flux financiers 2025", "sous_titre": "Les ressources de l'année couvrent les prestations, les frais et dégagent un excédent",
         "synthese": f"Produits {F('Produits 2025', s[2025]['produits'])}, charges {F('Charges 2025', s[2025]['charges'])}, résultat {F('Résultat 2025', s[2025]['resultat'])}"},
        "Le résultat de l'exercice est la différence entre produits et charges (y compris les éléments exceptionnels, dits hors activités ordinaires). En 2025, il reste positif, mais il diminue depuis 2021 (partie 5).",
        graphique={"titre": fig("Produits, charges et résultat 2025"), "type": "barres",
                   "categories": ["Cotisations maladie", "Cotisations décès", "Prestations maladie", "Prestations décès", "Frais de gestion", "Résultat"],
                   "series": [{"nom": "2025", "valeurs": [round(s[2025]["cot_maladie"]), round(s[2025]["cot_deces"]), round(s[2025]["prest_maladie"]), round(s[2025]["prest_deces"]), round(s[2025]["frais_gestion"]), round(s[2025]["resultat"])], "couleur": BLEU}],
                   "format": FMT_M, "legende": None, "unite": "En millions de FCFA"},
        analyse=[f"**Les cotisations sont la ressource principale du FPPN.** Elles atteignent {F('Cot mal 2025', s[2025]['cot_maladie'])} en maladie et {F('Cot dec 2025', s[2025]['cot_deces'])} en décès.",
                 f"**Les prestations maladie absorbent la plus grande part des ressources.** Elles valent {F('Prest mal 2025', s[2025]['prest_maladie'])}, contre {F('Prest dec 2025', s[2025]['prest_deces'])} en décès.",
                 f"**Les frais de gestion pèsent {pc('Taux frais 2025', s[2025]['taux_frais'])} des cotisations.** Ils couvrent le personnel, les indemnités des organes, le fonctionnement et les amortissements."],
        implication="La garantie décès dégage un excédent qui compense en partie la garantie maladie : le résultat global masque un déséquilibre de la maladie, analysé par risque dans la partie 10.",
        source=SRC_COMPTA)

contenu({"titre": "Notions de base", "sous_titre": "Huit notions suffisent pour lire l'ensemble du rapport",
         "synthese": "Chaque notion est définie ici une fois, puis utilisée avec le même sens dans tout le rapport"},
        "Ces définitions sont volontairement simples. Les formules exactes figurent dans les diapositives « méthode » de chaque partie et dans l'annexe des méthodes.",
        tableau={"titre": tab("Notions utilisées dans le rapport"), "entetes": ["Notion", "Définition simple", "Exemple dans le rapport"], "largeurs": [4.8, 16.2, 10.2],
                 "lignes": [["Cotisation", "Somme que le mutualiste verse pour être couvert.", "5 % du salaire de base"],
                            ["Prestation", "Somme que le FPPN paie quand le risque survient (soins, décès).", "Factures des cliniques, capital décès"],
                            ["Sinistralité (S/C)", "Part des cotisations consommée par les prestations.", "S/C maladie 2025"],
                            ["Provision technique", "Somme mise de côté au bilan pour payer des prestations déjà nées mais pas encore payées.", "Provision pour prestations à payer"],
                            ["Engagement", "Ce que le FPPN devra payer plus tard à ses mutualistes.", "Épargne PEP et EFC à restituer"],
                            ["Adossement", "Correspondance entre les dettes et les actifs disponibles pour les payer.", "Épargne et placements réservés"],
                            ["Fonds propres", "Ce qui reste quand on retire toutes les dettes de tous les actifs.", "Négatifs de 2020 à 2026"],
                            ["Ratio prudentiel", "Rapport chiffré qu'une norme impose de respecter pour protéger les adhérents.", "Réserve obligatoire de 6 à 9 mois"]]},
        disposition="dessous",
        analyse=["**Ces huit notions portent toute l'analyse.** Les quatre premières décrivent l'activité de l'année ; les quatre dernières décrivent la capacité du FPPN à tenir ses promesses dans la durée."],
        source="Source : Adenka ; " + DEC)

# ---------------------------------------------------------------- PARTIE 2 : DÉMARCHE
separateur(2, "Contexte et démarche", "Mission, périmètre, chaîne de traitement")
contenu({"titre": "Mission et périmètre", "sous_titre": "Le volet actuariel répond à quatre questions sur les engagements et les placements",
         "synthese": "Période d'analyse du 1er janvier 2020 au 31 août 2026 ; données arrêtées à l'envoi du FPPN reçu le 15/10/2026"},
        "Les quatre questions viennent du dossier de préparation de la mission ; les trois axes complémentaires (données, conformité, équilibre) en découlent. La colonne « état » dit honnêtement ce qui est fait et ce qui attend des données.",
        tableau={"titre": tab("Périmètre de la mission actuarielle et statistique"), "entetes": ["Axe", "Question traitée", "État au 19/10/2026"], "largeurs": [6.2, 15.0, 10.0],
                 "lignes": [["1. Engagements", "Combien le FPPN doit-il à ses adhérents et ayants droit ?", "Partiel : épargne mesurée, autres engagements en attente (partie 9)"],
                            ["2. Suffisance des provisions", "Les provisions au bilan suffisent-elles à payer les prestations dues ?", "Non démontrable : provisions héritées, recalcul en attente"],
                            ["3. Adossement actif-passif", "Les actifs arrivent-ils en caisse au bon moment pour payer ?", "Non-adossement démontré (borne favorable)"],
                            ["4. Adéquation des placements", "Les placements sont-ils sûrs, liquides et rentables ?", "Liquidité et diversification insuffisantes"],
                            ["Contrôle des données", "Les données sont-elles complètes et cohérentes avec la comptabilité ?", "Réalisé sur les données reçues"],
                            ["Conformité prudentielle", "Les ratios de la Décision n° 05/2022 sont-ils respectés ?", "Réalisé : 36 indicateurs examinés"],
                            ["Équilibre technique", "Les cotisations futures couvrent-elles les prestations futures ?", "En attente des âges et des dossiers"]]},
        disposition="dessous",
        analyse=["**Deux des quatre questions reçoivent une réponse établie : le FPPN n'est pas adossé, et ses placements ne sont ni assez liquides ni assez diversifiés.** Les engagements et la suffisance des provisions attendent les données individuelles (réponse détaillée en résumé exécutif et en partie 9)."],
        implication="Le calcul des engagements (question 1) conditionne la réponse définitive aux questions 2 et 3 : il reste la priorité dès réception des fichiers (PBC, items 57, 71, 72 et 92).",
        source="Source : dossier de préparation de la mission (13/10/2026) et liste des documents demandés (PBC)")
# ---- v9 : périmètre, domaine et période de validité (niveau 1 : diapositive de synthèse)
SRC_22 = "Source : documents du FPPN reçus le 22/10/2026"
contenu({"titre": "Périmètre et validité des analyses", "sous_titre": "Chaque résultat ne vaut que pour les données, la population et la période qui le fondent",
         "synthese": "Trois niveaux de lecture : cette diapositive, une ligne « Valable pour » sous chaque tableau et graphique, et un tableau détaillé en annexe"},
        "Un résultat mesuré sur quelques mois, une partie de la population ou un seul document ne peut pas être étendu au reste sans vérification. Le tableau donne, pour chaque famille d'analyses, ce qui a été mesuré et ce à quoi le résultat ne doit pas être généralisé.",
        tableau={"titre": tab("Périmètre, domaine et période de validité des analyses"), "entetes": ["Analyses", "Données utilisées", "Période", "Ne pas généraliser à"], "largeurs": [6.0, 10.4, 5.6, 9.2], "taille": 8.5,
                 "lignes": [["Effectifs, taux de cotisation (parties 3 et 4)", "Fichiers de précomptes de la Solde et de la CGRAE", "01/2024 à 09/2026 selon le poste", "2020 à 2023, civils de la Direction Générale"],
                            ["Taux de 3,77 % (maladie)", "Salaire de base du fichier 936, " + N("Lignes taux (validité)", T936["lignes"]) + " lignes", "8 mois : 04 à 10/2024 et 01/2025", "Mois sans salaire de base dans les fichiers"],
                            ["Taux de 1,23 % (décès)", "8Z6 rapporté au salaire de base, puis rapport 8Z6 / 936", "Mesure directe 04/2024 et 01/2025, rapport de 03/2024 à 09/2026", "Janvier et février 2024 (autre clé)"],
                            ["Sinistralité, résultat, bilan, ratios (parties 5, 9, 10)", "Balances 2020 au 31/08/2026, états financiers", "2020 à 2025, 2026 provisoire", "Sinistres tardifs, valeurs de marché"],
                            ["Épargne (partie 6)", "Soldes et mouvements comptables, grand livre 2024", "2020 au 31/08/2026", "Engagement individuel (aucun échéancier)"],
                            ["Mortalité brute (partie 5)", "Étude des ratios du FPPN : décès et adhérents", "2015 à 2025", "Mortalité par âge (dates de naissance absentes)"],
                            ["Provision du personnel (partie 9)", "DISA 2021, 2022, 2023 et 2025 (CNPS)", "31/12 des années citées", "2024 (DISA scannée), policiers détachés"],
                            ["Contrôles croisés (partie 3)", "États financiers, rapprochements bancaires, étude du FPPN", "2020 au 31/08/2026", "Comptes non cités dans le contrôle"]]},
        disposition="dessous",
        analyse=["**Les documents reçus le 22/10/2026 permettent de vérifier plusieurs résultats par une seconde source.** Les états financiers, les rapprochements bancaires et l'étude des ratios du FPPN confirment les balances et les fichiers de la Solde sur les périodes indiquées."],
        implication="Tout chiffre repris hors du rapport doit être accompagné de sa ligne « Valable pour ». L'annexe « Périmètre et validité des analyses » détaille chaque diapositive.",
        valable="les familles d'analyses du rapport, sur les documents reçus jusqu'au 22/10/2026",
        source="Source : Adenka, à partir des fichiers cités dans chaque partie")
contenu({"titre": "Documents reçus le 22/10/2026", "sous_titre": "124 fichiers reçus intacts, qui complètent les engagements, la méthode du FPPN et les contrôles",
         "synthese": "Les nouveaux documents servent à cinq usages : provisions, engagements d'épargne, provision du personnel, actifs immobilisés et contre-vérification des chiffres"},
        "La taille de chaque fichier reçu a été comparée à celle relevée sur le poste du FPPN avant l'envoi : aucun fichier n'a été altéré pendant le transfert.",
        tableau={"titre": tab("Documents reçus le 22/10/2026 et usage dans le rapport"), "entetes": ["Famille", "Documents", "Usage dans le rapport"], "largeurs": [4.2, 15.0, 12.0], "taille": 8,
                 "lignes": [["Règles et produits", "Cahiers des charges maladie-décès et EFC-PEP 2021, manuel des produits d'épargne, cartographie des risques, politique de crédit du Grenier, conventions IPS-CGRAE, Transvie-CMP et DEPA", "Taux servis, barème décès, garantie santé des retraités, dépréciation des prêts"],
                            ["Comptes", "États financiers de 2020 au 31/08/2026, grands livres généraux et des tiers 2023 à 2025, balances, rapprochements bancaires de 2020 à 2026", "Contrôles croisés, délais de paiement des cliniques, factures arrivées en retard, paiements d'allocations décès"],
                            ["Engagements", "États de fin d'année des allocations décès et frais funéraires 2021 à 2025, étude et interprétation des ratios maladie et décès", "Dette décès, décès connus après la clôture, méthode de sinistralité du FPPN, mortalité brute"],
                            ["Personnel", "DISA et ETAT 301 de 2020 à 2025, DASC", "Provision du personnel agent par agent"],
                            ["Actifs et gouvernance", "Arrêté d'autorisation du CMP, dérogation d'urbanisme du projet de Cocody, attestation de désignation du directeur général, titres de propriété, baux, rapports de gestion, budgets, organigrammes", "Clinique, projet de Cocody, gouvernance des placements, épargne présentée en Assemblée générale"]]},
        disposition="dessous",
        analyse=["**Les états financiers et les rapprochements bancaires confirment les balances utilisées dans le rapport**, à deux écarts près, présentés en partie 3."],
        implication="Les documents encore manquants (base des prestations datée, relevés individuels d'épargne, dates de naissance des mutualistes) restent dans la PBC et limitent les analyses présentées.",
        valable="fichiers reçus le 22/10/2026 dans les dossiers transmis par le FPPN ; les documents reçus le 15/10/2026 et le 18/10/2026 sont décrits en partie 2 et en partie 3",
        source="Source : documents du FPPN reçus le 22/10/2026")
contenu({"titre": "Chaîne de traitement", "sous_titre": "Chaque chiffre est tracé du fichier du FPPN jusqu'au rapport",
         "synthese": "Quatre décisions de contrôle arrêtent ou corrigent le traitement ; le détail de chaque étape figure dans la partie 13"},
        "Logigramme modifiable dans PowerPoint. Ovale : début ou fin ; rectangle : traitement ; losange : contrôle avec sa décision ; forme ondulée : document produit. Chaque étape correspond à un script Python conservé dans le dossier de la mission ; chaque contrôle est automatique.",
        logigramme="chaine",
        analyse=["**Aucun chiffre n'entre dans le rapport sans avoir passé quatre contrôles.** Un fichier illisible part dans la liste des fichiers à renvoyer, une conversion inexacte est refaite, une ligne de total ou un montant doublé est écarté et journalisé, et un chiffre affiché qui ne correspond pas à sa valeur exacte est corrigé."],
        implication="Le FPPN peut retrouver, pour chaque chiffre du rapport, le fichier d'origine, les lignes écartées et le calcul : les écarts relevés peuvent être discutés sur pièces.",
        source="Source : scripts Python Adenka (dossier scripts de la mission)")

# ---------------------------------------------------------------- PARTIE 3 : DONNÉES
separateur(3, "Données et qualité", "Couverture, contrôles, anomalies détectées")
methode_gab({"titre": "Méthode : contrôle des données", "sous_titre": "Trois contrôles garantissent que la base reproduit fidèlement les fichiers",
             "synthese": "Exhaustivité, cohérence interne et rapprochement entre sources indépendantes, puis une limite assumée"},
            [("Exhaustivité", "Adenka calcule l'empreinte SHA-256 de chaque fichier et compte lignes et cellules avant et après conversion : 218 fichiers copiés et 491 feuilles converties sont identiques."),
             ("Cohérence interne", "Les totaux imprimés par les logiciels sont retrouvés dans les 37 balances âgées du Grenier, et les débits égalent les crédits dans les 7 exercices comptables."),
             ("Rapprochement", "Deux sources qui décrivent le même flux doivent concorder : le taux de précompte rapporte le montant retour au montant aller, poste par poste et mois par mois."),
             ("Limite", "Ces contrôles prouvent que la base est fidèle aux fichiers reçus. Ils ne prouvent pas que les fichiers du FPPN sont eux-mêmes exacts.")],
            "Exhaustivité : rien n'est perdu entre le fichier reçu et la base. Cohérence : les totaux imprimés par les logiciels sont retrouvés. Rapprochement : deux sources qui décrivent le même flux doivent concorder. Limite : la fidélité aux fichiers n'est pas l'exactitude des fichiers.")
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
        analyse=["**Seules la comptabilité et l'épargne PEP couvrent toute la période d'analyse.** La comptabilité donne des totaux annuels ; les fichiers texte de la Solde donnent le PEP par mutualiste et par mois de 2020 à 2022.",
                 "**Les précomptes mensuels commencent en 2023 ou en 2024.** Les fichiers maladie et décès débutent en janvier 2024, ceux des retraités en février 2024 et ceux de l'EFC en janvier 2023.",
                 "**Aucune base de prestations n'a été transmise.** Le coût des soins et des décès n'est connu que par les comptes, sans date de survenance ni bénéficiaire."],
        implication="Les analyses mensuelles portent sur 2024 à 2026, et les années 2020 à 2023 ne peuvent être analysées qu'en totaux comptables. Les conclusions sur l'évolution de la sinistralité restent donc fragiles tant que les bases de prestations ne sont pas reçues.",
        source=SRC_PREC)
rt = R["lignes_total_ecartees"]
lt = EX["ligne_total"]
contenu({"titre": "Calcul pas à pas : lignes de total", "sous_titre": "Une ligne de total cachée aurait doublé les cotisations maladie du mois",
         "synthese": f"Fichier de février 2025 : la dernière ligne porte {F('Total février 2025', float(lt['lignes'][-1][5]))}, soit la somme des {N('Lignes fichier février 2025', lt['nb_lignes_fichier'] - 1)} lignes"},
        "Lecture : (1) la colonne matricule de la dernière ligne contient un nombre de lignes, pas un matricule ; (2) le montant est la somme de toutes les lignes au-dessus ; (3) compter cette ligne aurait doublé le total du mois. Règle appliquée : une ligne dont le montant égale la somme des autres (à 1 % près) est écartée. Au total, " + N("Lignes total écartées (notes)", rt["n"]) + " lignes de ce type ont été écartées. Noms et matricules masqués.",
        pas_a_pas={"entetes": lt["entete"], "lignes": lt["lignes"], "largeurs": [2.0, 2.6, 1.6, 3.4, 3.6, 3.6, 3.4, 2.4],
                   "rouge": [(len(lt["lignes"]) - 1, 3), (len(lt["lignes"]) - 1, 5)],
                   "explications": ["La colonne « matricule » de la dernière ligne contient le nombre de lignes du fichier.",
                                    f"Le montant de {F('Total février (pas)', float(lt['lignes'][-1][5]))} est la somme de toutes les lignes au-dessus.",
                                    f"Adenka écarte cette ligne et les {N('Lignes total écartées', rt['n'])} lignes du même type, ce qui évite {F('Montant total écarté', rt['montant'])} de double compte."],
                   "fichier": lt["fichier"]},
        analyse=["**Une simple somme des fichiers aurait gonflé les cotisations.** Les lignes de total ont la même forme que les lignes des mutualistes et ne se repèrent que par le calcul."],
        implication="Le FPPN doit retirer ces lignes de total de ses fichiers, ou les isoler, avant tout traitement automatique.",
        source="Source : " + lt["fichier"] + " (identifiant d'extraction " + lt["id"] + ") ; noms et matricules masqués")
md = EX["montants_doubles"]
contenu({"titre": "Calcul pas à pas : montants doublés", "sous_titre": "Des montants saisis deux fois à la suite gonflaient les précomptes décès",
         "synthese": f"Avril 2024 : {N('Montants corrigés', md['total_corriges'])} montants de plus d'un million de FCFA sont des montants écrits deux fois"},
        "Lecture : (1) le montant source répète deux fois la même suite de chiffres ; (2) un précompte décès de plusieurs milliards de FCFA par agent est impossible (médiane du poste : environ 2 300 FCFA) ; (3) le montant retenu est la première moitié. Correction appliquée seulement au poste décès, au-delà d'un million de FCFA, et signalée au FPPN.",
        pas_a_pas={"entetes": ["Matricule", "Nom", "Montant dans le fichier", "Montant corrigé"], "lignes": md["lignes"], "largeurs": [4.0, 4.0, 8.0, 6.0],
                   "rouge": [(0, 2), (0, 3)],
                   "explications": ["Le montant du fichier répète deux fois la suite « 47924 ».", "Un précompte décès de plus d'un million de FCFA par agent est impossible.",
                                    "Adenka retient la première moitié du montant et signale la ligne au FPPN."], "fichier": md["fichier"]},
        analyse=["**L'erreur vient de la saisie du fichier source, pas d'un calcul.** Les deux moitiés du montant sont identiques, ce qui signe une double frappe."],
        implication="La correction est limitée au poste décès et à un seuil d'un million de FCFA pour ne pas modifier de montants réels. Le FPPN doit confirmer chaque ligne signalée.",
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
        analyse=[f"**La Solde et la CGRAE retiennent presque tout ce que le FPPN demande.** Le taux moyen va de {pc('Taux min moyen', min(a['taux_recouvrement'] for a in ar))} à {pc('Taux max moyen', max(a['taux_recouvrement'] for a in ar))} selon le poste.",
                 f"**Certains mois s'écartent nettement de la moyenne.** Le mois le plus bas atteint {pc('Min Grenier', [a for a in ar if a['poste'] == '9X4'][0]['taux_min_mensuel'])} pour le Grenier et {pc('Min PEP', [a for a in ar if a['poste'] == '998'][0]['taux_min_mensuel'])} pour le PEP.",
                 "**Ces écarts traduisent des rejets ou des fichiers incomplets.** Le FPPN ne rapproche pas les fichiers aller et retour et ne reçoit les motifs de rejet que depuis environ deux mois (" + ENT19 + ")."],
        implication="Pour les retraités, le taux mesuré sert d'indicateur de recouvrement des cotisations : il respecte la norme d'au moins 90 % de la Décision n° 05/2022. Pour les cotisations maladie et décès des actifs, le calcul attend les fichiers aller, que la direction du personnel de la Police transmet directement à la Solde.",
        source=SRC_PREC)
apm, apx = EX["aller_retour_pep"]["min"], EX["aller_retour_pep"]["max"]
contenu({"titre": "Calcul pas à pas : écarts du PEP", "sous_titre": "Les deux mois extrêmes du PEP s'expliquent par le contenu des fichiers",
         "synthese": f"Janvier 2025 : {pc('Taux PEP janv 2025', apm['taux'])} ; juillet 2024 : {pc('Taux PEP juil 2024', apx['taux'])}"},
        "Janvier 2025 : le fichier retour compte moins de lignes que le fichier aller, des précomptes n'ont pas été faits. Juillet 2024 : le fichier retour compte plus de lignes que le fichier aller, le fichier aller n'incluait pas les nouveaux épargnants (le fichier aller d'août 2024 en compte nettement plus).",
        pas_a_pas={"entetes": ["Mois", "Fichier aller", "Lignes aller", "Montant aller", "Lignes retour", "Montant retour", "Taux"],
                   "lignes": [["janv. 2025", apm["fichier_a"], N("Lignes aller janv25", apm["lignes_a"]), F("Montant aller janv25", apm["montant_a"]), N("Lignes retour janv25", apm["lignes_r"]), F("Montant retour janv25", apm["montant_r"]), pc("Taux janv25 (t)", apm["taux"])],
                              ["juil. 2024", apx["fichier_a"], N("Lignes aller juil24", apx["lignes_a"]), F("Montant aller juil24", apx["montant_a"]), N("Lignes retour juil24", apx["lignes_r"]), F("Montant retour juil24", apx["montant_r"]), pc("Taux juil24 (t)", apx["taux"])]],
                   "largeurs": [2.4, 6.6, 2.6, 5.4, 2.6, 5.4, 3.0], "rouge": [(0, 4), (1, 4)],
                   "explications": ["En janvier 2025, " + N("Écart lignes janv25", apm["lignes_a"] - apm["lignes_r"]) + " épargnants demandés n'ont pas été précomptés.",
                                    "En juillet 2024, le fichier retour compte " + N("Écart lignes juil24", apx["lignes_r"] - apx["lignes_a"]) + " lignes de plus que le fichier aller.",
                                    "Le taux mensuel dépend donc de la qualité des deux fichiers, et le FPPN doit tenir un état des rejets."],
                   "fichier": "FPPN-SOLDE\\PEP\\FICHIER ALLER et FICHIER RETOUR"},
        analyse=["**Un taux supérieur à 100 % ne signifie pas un trop-perçu.** Il révèle un fichier aller incomplet, qui n'inclut pas les nouveaux épargnants."],
        implication="Sans état des rejets, le FPPN ne peut pas prouver que chaque épargnant a été précompté du bon montant : c'est une condition pour reconstituer les comptes individuels.",
        source=SRC_PREC)
contenu({"titre": "Fichiers vides ou endommagés", "sous_titre": "Quatre défauts restent ouverts : décembre 2024, MEDOC et deux fichiers mal classés",
         "synthese": "Tous les fichiers comptables endommagés ont été renvoyés lisibles le 22/10/2026, dont le grand livre des comptes 2022"},
        "Un fichier vide (0 octet) ne contient aucune donnée ; un fichier endommagé a une partie de ses données détruite (souvent lors d'une copie ou d'un transfert). Aucun de ces défauts ne vient du traitement d'Adenka : les empreintes prouvent que les fichiers sont identiques à ceux reçus.",
        tableau={"titre": tab("Fichiers à renvoyer ou à corriger"), "entetes": ["Fichier", "Défaut", "Conséquence"], "largeurs": [11.4, 6.2, 13.6],
                 "lignes": [["« FPPN_ALLOC_DECES_8Z6_DECEMBRE_2024.xlsx » et « 122024_effectif_police_936_DECEMBRE 2024.xlsx »", "Vides (0 octet)", "Le mois de décembre 2024 n'a aucune donnée. Aucun envoi suivant ne les contient."],
                            ["MEDOC 8Z7, retours de février et avril 2026", "Aucune ligne", "Les précomptes MEDOC de ces deux mois sont inconnus."],
                            ["Balances des tiers 2021 et 2025, grands livres des comptes 2021, 2022 et 2024, grands livres des tiers 2021, 2023, 2024 et 2025 (Excel)", "Endommagés à la première réception", "Renvoyés lisibles le 22/10/2026 (dossier « 02-Nouveaux_GL et Balances »). « GRAND LIVRE DES COMPTES 2022.xlsx » couvre l'année entière (192 comptes)."],
                            ["Fichiers aller rangés dans les dossiers retour", "Mal classés", "Les retours de CGRAE juin 2024 et GNA janvier 2026 manquent."]]},
        disposition="dessous",
        analyse=["**Les défauts portent sur les fichiers reçus et non sur leur traitement.** Les empreintes SHA-256 prouvent que les fichiers analysés sont identiques à ceux transmis."],
        implication="Le renvoi des fichiers de décembre 2024 complètera l'année 2024, sans changer les totaux annuels déjà utilisés.",
        source="Source : journal de contrôle de l'extraction ; note interne des anomalies")
# ---- v11 : qualité des bases variable par variable (complétude et cohérence)
QB = json.load(open(os.path.join(os.path.dirname(V11J), "qualite.json"), encoding="utf-8"))
_SEL = ["Matricule", "Nom", "Montant", "Date de naissance"]
_bases = []
for x in QB["completude"]:
    if x["base"] not in _bases:
        _bases.append(x["base"])
_lq = []
for bse in _bases:
    lig = [x for x in QB["completude"] if x["base"] == bse]
    n_ = lig[0]["lignes"]
    if bse.startswith("Prêts"):
        vals = {x["variable"]: x["manquantes"] for x in lig}
        _lq.append([bse, N(f"QB n {bse}", n_), "Mécano : " + N(f"QB mec {bse}", vals["Mécano de l'emprunteur"]) + " manquants", "Dates, taux, durée : complets",
                    "Aucune incohérence de dates ni d'encours sur " + N(f"QB co {bse}", sum(1 for c in QB["coherence"] if c["base"] == bse)) + " contrôles", "Sans objet"])
        continue
    vals = {x["variable"]: x["manquantes"] for x in lig}
    dbl = next(c["anomalies"] for c in QB["coherence"] if c["base"] == bse and c["controle"].startswith("Même matricule"))
    neg = next(c["anomalies"] for c in QB["coherence"] if c["base"] == bse and c["controle"].startswith("Montant nul"))
    _lq.append([bse, N(f"QB n {bse}", n_), "Matricule : " + N(f"QB mat {bse}", vals["Matricule"]) + " manquants. Montant : " + N(f"QB mt {bse}", vals["Montant"]),
                "Date de naissance : " + pc(f"QB dn {bse}", vals["Date de naissance"] / n_) + " manquante",
                N(f"QB dbl {bse}", dbl) + " lignes d'un même matricule dans un mois. " + N(f"QB neg {bse}", neg) + " montants nuls ou négatifs",
                "Non" if vals["Date de naissance"] / n_ > 0.5 else "Oui"])
contenu({"titre": "Qualité des bases, variable par variable", "sous_titre": "Les montants et les identifiants sont complets, les dates de naissance absentes",
         "synthese": "Les bases permettent les calculs comptables et statistiques, pas les calculs par âge"},
        "Contrôles de complétude (valeur manquante) et de cohérence (doublons dans un mois, montants nuls ou négatifs, dates incohérentes) sur les lignes retenues de janvier 2020 à août 2026. Calculs par âge : possibles si la date de naissance est présente sur plus de la moitié des lignes.",
        tableau={"titre": tab("Complétude et cohérence des bases de données"), "entetes": ["Base", "Lignes", "Identifiants et montants", "Date de naissance", "Cohérence", "Utilisable pour les calculs par âge ?"],
                 "largeurs": [5.6, 2.8, 6.6, 4.8, 8.4, 3.0], "taille": 8.5, "lignes": _lq},
        disposition="dessous",
        analyse=["**Les lignes multiples d'un même matricule dans un mois ne sont pas des erreurs pour le PEP et le Grenier** : un épargnant peut avoir plusieurs plans et un emprunteur plusieurs prêts. Pour la maladie et le décès, les effectifs comptent les matricules distincts : ces lignes ne gonflent pas les effectifs.",
                 "**Seuls les fichiers CGRAE de février à avril 2024 portent la date de naissance** (105 retraités). Les grades et salaires ne figurent que dans les fichiers détaillés de la maladie d'avril à octobre 2024 et de janvier 2025 (209 670 lignes, complètes)."],
        implication="La qualité des montants autorise les analyses comptables du rapport. L'absence de date de naissance empêche les seuls calculs par âge (mortalité, engagements viagers, projection).",
        valable="lignes retenues des fichiers de précomptes de janvier 2020 à août 2026 ; balance âgée du Grenier au 31/08/2026",
        source=SRC_PREC + " ; " + SRC_GRENIER[9:] + " ; script qualite_bases_fppn.py")

retenir("données", ("Les données reçues sont fidèlement reprises, mais incomplètes", "Base contrôlée à 100 %, sans prestations, sans âges et sans précomptes mensuels de 2020 à 2022",
                    "La base reproduit les fichiers du FPPN, à 26 lignes près sans matricule valide (737 890 FCFA, 0,0017 % des précomptes), mais il manque les données nécessaires aux calculs actuariels."),
        f"Les empreintes, les lignes et les totaux des logiciels sont identiques. Adenka a écarté {N('Total écartés (ret)', rt['n'])} lignes de total et corrigé {N('Corrigés (ret)', md['total_corriges'])} montants doublés.",
        "La qualité des analyses dépend désormais de la transmission des données manquantes, et non plus du traitement des fichiers.",
        "Le FPPN doit transmettre les bases de prestations, les dates de naissance et les fichiers vides de décembre 2024.",
        "Message pour le client : la qualité de l'analyse dépend désormais de la transmission des données manquantes, pas du traitement.")

# ---------------------------------------------------------------- PARTIE 4 : EFFECTIFS ET COTISATIONS
# ---- v9 : contrôles croisés (partie 3)
EFB = V9["ef_contre_balances"]
_e23 = next(x for x in EFB if x["exercice"] == 2023)
_e26 = next(x for x in EFB if x["exercice"] == 2026)
_nok = [x["exercice"] for x in EFB if x["ef_brut"] != x["balance_brut"] or x["ef_resultat"] != x["balance_resultat"]]
contenu({"titre": "Contrôle croisé : états financiers et balances", "sous_titre": f"Les états financiers concordent avec les balances {len(EFB) - len(_nok)} années sur {len(EFB)}",
         "synthese": "Total de l'actif, résultat et trésorerie identiques au franc près, sauf en 2023 (deux versions des comptes) et au 31/08/2026 (résultat)"},
        "Les montants des états financiers ont été relevés sur les pages « Bilan Actif » et « Bilan Passif » des PDF scannés, puis comparés aux balances utilisées dans tout le rapport. La trésorerie-actif des états financiers est comparée à la somme des soldes débiteurs de la classe 5, car les découverts figurent au passif.",
        tableau={"titre": tab("États financiers et balances : actif brut, résultat, trésorerie"), "entetes": ["Exercice", "Actif brut (états financiers)", "Écart avec la balance", "Résultat (états financiers)", "Écart avec la balance", "Trésorerie : écart"],
                 "largeurs": [3.6, 6.0, 4.8, 6.0, 4.8, 4.2], "taille": 9,
                 "lignes": [[str(x["exercice"]) if x["exercice"] < 2026 else "31/08/2026", F(f"EF brut {x['exercice']}", x["ef_brut"]), F(f"EF écart brut {x['exercice']}", x["ef_brut"] - x["balance_brut"]),
                             F(f"EF résultat {x['exercice']}", x["ef_resultat"]), F(f"EF écart rés {x['exercice']}", x["ef_resultat"] - x["balance_resultat"]),
                             F(f"EF écart trés {x['exercice']}", x["ef_tresorerie"] - x["balance_classe5"])] for x in EFB]},
        disposition="dessous",
        analyse=[f"**En 2023, les états financiers scannés portent un actif brut inférieur de {F('EF écart 2023 (t)', _e23['balance_brut'] - _e23['ef_brut'])} et un excédent de {F('EF rés 2023 (t)', _e23['ef_resultat'])}, contre {F('Bal rés 2023 (t)', _e23['balance_resultat'])} dans la balance.** Le rapport de gestion 2023 « corrigé et finalisé » et la colonne 2023 des états financiers 2024 reprennent les montants de la balance : le scan est une version antérieure à la correction.",
                 f"**Au 31/08/2026, le résultat des états financiers ({F('EF rés 2026 (t)', _e26['ef_resultat'])}) diffère de celui de la balance provisoire ({F('Bal rés 2026 (t)', _e26['balance_resultat'])}).** L'actif brut et la trésorerie sont pourtant identiques."],
        implication="Les balances utilisées dans le rapport sont confirmées par les comptes publiés. Le FPPN doit indiquer quelle version des comptes 2023 a été approuvée et expliquer l'écart de résultat au 31/08/2026 (PBC, item 110).",
        valable="exercices 2020 à 2025 et situation au 31/08/2026, totaux du bilan et résultat ; les rubriques détaillées n'ont pas été comparées",
        source="Source : ETATS FINANCIERS 2020.pdf à ETATS FINANCIERS 2025.pdf, ETATS FINANCIERS AU 31-08-2026.pdf, BILAN_DE_L'EXERCICE_CLOS_31.12.2023_(corrigé_et_finalisé)[1].pdf ; " + SRC_COMPTA[9:])
_ad = V9["controle_adherents_8z6"]
_rm = {(x["annee"], x["mois"]): x for x in V9["rapport_8z6_936_mois"]}
_ra = V9["rapport_8z6_936_annee"]
contenu({"titre": "Contrôle croisé : effectifs du FPPN et fichiers de la Solde", "sous_titre": "Les effectifs déclarés par le FPPN sont ceux des fichiers 8Z6, à dix agents près en 2026",
         "synthese": "Le nombre d'adhérents du FPPN égale le nombre de lignes 8Z6 en novembre 2024 et décembre 2025 ; la clé 3,77 / 1,23 s'applique depuis mars 2024"},
        "Deux sources indépendantes sont comparées : l'étude des ratios du FPPN (nombre d'adhérents par période) et les fichiers retour 8Z6 de la Solde (une ligne par policier précompté). La seconde partie du tableau rapporte, agent par agent, le précompte décès (8Z6) au précompte maladie (936) du même mois : la clé 3,77 / 1,23 donne un rapport de 0,3263.",
        tableau={"titre": tab("Adhérents du FPPN, lignes 8Z6 et rapport 8Z6 / 936"), "entetes": ["Période", "Adhérents (FPPN)", "Lignes 8Z6 (Solde)", "Écart", "Rapport médian 8Z6 / 936", "Lignes à la clé 3,77 / 1,23"],
                 "largeurs": [4.0, 5.0, 5.0, 3.4, 6.0, 6.0], "taille": 9,
                 "lignes": [[f"{x['mois']:02d}/{x['annee']}", N(f"Adh FPPN {x['annee']}-{x['mois']}", x["adherents_fppn"]), N(f"Lignes 8Z6 {x['annee']}-{x['mois']}", x["lignes_8z6"]),
                             N(f"Écart adh {x['annee']}-{x['mois']}", x["adherents_fppn"] - x["lignes_8z6"]),
                             nb(_rm[(x["annee"], x["mois"])]["mediane"], 4) if (x["annee"], x["mois"]) in _rm else "-",
                             pc(f"Part clé {x['annee']}-{x['mois']}", _rm[(x["annee"], x["mois"])]["part_cle"]) if (x["annee"], x["mois"]) in _rm else "-"] for x in _ad]
                          + [[f"{m:02d}/2024", "-", N(f"Lignes 8Z6 2024-{m}", _rm[(2024, m)]["lignes"]), "-", nb(_rm[(2024, m)]["mediane"], 4), pc(f"Part clé 2024-{m}", _rm[(2024, m)]["part_cle"])] for m in (1, 2, 3)]},
        disposition="dessous",
        analyse=[f"**La répartition 3,77 / 1,23 s'applique à {pc('Part clé 2025 (t)', next(x['part_cle'] for x in _ra if x['annee'] == 2025))} des lignes de 2025 et à {pc('Part clé 2026 (t)', next(x['part_cle'] for x in _ra if x['annee'] == 2026))} des lignes de 2026.** En janvier et février 2024, le rapport médian vaut {nb(_rm[(2024, 1)]['mediane'], 4)} et {nb(_rm[(2024, 2)]['mediane'], 4)} : une autre clé s'appliquait alors.",
                 "**De janvier à juin 2026, le FPPN compte chaque mois dix adhérents de moins que les fichiers 8Z6.** Les cotisations décès déclarées par le FPPN dépassent aussi la somme des lignes 8Z6 du mois."],
        implication="Les effectifs utilisés par le FPPN pour ses ratios sont fiables à l'unité près. La clé appliquée avant mars 2024 et l'origine des dix agents d'écart sont à demander (PBC, item 111).",
        valable="mois cités, policiers précomptés par la Solde (8Z6 et 936) ; les mois de 2024 sans fichier 8Z6 ne sont pas couverts",
        source="Source : ETUDE DES RATIO MALADIE DECES 2.xlsx, feuille Tableau, lignes « NB D'ADHERENTS » et « COTISATIONS PAYEES » ; " + SRC_PREC[9:])
_rb = V9["rapprochements"]
_r25 = next(x for x in _rb if x["exercice"] == 2025)
contenu({"titre": "Contrôle croisé : rapprochements bancaires", "sous_titre": "Les soldes bancaires de la comptabilité sont ceux de la balance ; des chèques restent en suspens depuis 2021",
         "synthese": f"Au 31/12/2025, {N('Chèques suspens 2025', _r25['suspens_plus_un_an'])} chèques émis depuis plus d'un an ne sont pas encaissés, pour {F('Montant suspens 2025', _r25['montant_suspens_plus_un_an'])}"},
        "Chaque état de rapprochement compare, à une date, le solde du compte dans la comptabilité du FPPN et le solde du relevé de la banque, puis liste les opérations qui expliquent l'écart. Les chèques émis par le FPPN mais pas encore présentés à la banque figurent dans cette liste.",
        tableau={"titre": tab("Rapprochements bancaires au 31 décembre"), "entetes": ["Exercice", "Comptes rapprochés", "Soldes égaux à la balance", "Solde comptable", "Solde bancaire", "Chèques en suspens depuis plus d'un an", "Montant"],
                 "largeurs": [2.8, 3.4, 3.8, 5.2, 5.2, 5.4, 4.6], "taille": 9,
                 "lignes": [[str(x["exercice"]), N(f"ERB comptes {x['exercice']}", x["comptes"]), N(f"ERB concord {x['exercice']}", x["concordants"]), F(f"ERB compta {x['exercice']}", x["solde_compta"]),
                             F(f"ERB banque {x['exercice']}", x["solde_banque"]), N(f"ERB suspens {x['exercice']}", x["suspens_plus_un_an"]), F(f"ERB montant {x['exercice']}", x["montant_suspens_plus_un_an"])] for x in _rb]},
        disposition="dessous",
        analyse=["**Chaque solde comptable des rapprochements est égal au solde du compte de banque dans la balance, pour toutes les années.** Les balances sont donc cohérentes avec le suivi bancaire du FPPN.",
                 f"**Le nombre de chèques en suspens depuis plus d'un an passe de {N('Suspens 2022 (t)', next(x['suspens_plus_un_an'] for x in _rb if x['exercice'] == 2022))} fin 2022 à {N('Suspens 2025 (t)', _r25['suspens_plus_un_an'])} fin 2025**, le plus ancien datant du {_r25['plus_ancien'][8:10]}/{_r25['plus_ancien'][5:7]}/{_r25['plus_ancien'][:4]}. Une partie concerne le compte BOA de l'EFC, donc des règlements à des mutualistes."],
        implication="Les chèques non encaissés restent une dette envers les bénéficiaires : leur montant est faible, mais leur ancienneté doit être expliquée et les chèques périmés régularisés (PBC, item 91).",
        valable="comptes bancaires présents dans les états de rapprochement (BOA, Banque Atlantique, Orabank), au 31/12 de 2020 à 2025 ; les DAT, le FCP et la caisse ne sont pas couverts",
        source="Source : ETAT DE RAPPROCHEMENT BANCAIRE FPPN 2020.xlsx à ETAT DE RAPPROCHEMENT BANCAIRE FPPN 2026.xlsx ; " + SRC_COMPTA[9:])
separateur(4, "Effectifs et cotisations", "Cotisants, retraités, taux, cotisations en termes réels")
methode_gab({"titre": "Méthode : effectifs et taux", "sous_titre": "Effectifs, flux et taux sont mesurés agent par agent, chaque mois",
             "synthese": "Un cotisant est un matricule précompté au poste maladie 936 dans le mois"},
            [("Effectif", "L'effectif du mois est le nombre de matricules distincts du fichier maladie 936 de ce mois. Trente et un mois sont disponibles."),
             ("Flux", "Une entrée est un matricule présent un mois et absent le mois précédent ; une sortie est l'inverse. Seuls les mois consécutifs sont comparés."),
             ("Taux de cotisation", f"Le taux rapporte le précompte du mois au salaire de base du mois, pour chaque agent : {N('Lignes taux (méthode)', T936['lignes'])} lignes sont calculées."),
             ("Limite", "Les fichiers ne donnent pas le motif de sortie (retraite, décès, radiation), ni l'âge, ni le sexe des agents.")],
            "Effectif du mois m = nombre de matricules distincts du fichier maladie 936 du mois m. Entrée = matricule présent en m et absent en m-1 ; sortie = l'inverse. Taux de cotisation = précompte du mois / salaire de base du mois.")
e0 = [e for e in EFF if e["annee"] == 2024 and e["mois"] == 1][0]
e1 = [e for e in EFF if e["annee"] == 2026 and e["mois"] == 8][0]
tdc = C["tendances"][0]
contenu({"titre": "Effectifs cotisants", "sous_titre": "Les cotisants actifs augmentent d'un tiers en trente et un mois",
         "synthese": f"{N('Cotisants janv 2024', e0['cotisants'])} cotisants en janvier 2024, {N('Cotisants août 2026', e1['cotisants'])} en août 2026"},
        f"La droite de tendance estimée par moindres carrés donne une hausse de {N('Pente cotisants', tdc['pente'], 2)} cotisants par mois (intervalle de confiance à 95 % : {N('IC bas pente', tdc['ic95_bas'], 2)} à {N('IC haut pente', tdc['ic95_haut'], 2)}). La probabilité d'observer une telle pente si la vraie pente était nulle est inférieure à 0,01 % (p-valeur {tdc['p_valeur']:.2e}). La hausse se fait par paliers en juillet, ce qui rend la droite seulement approximative (R² = {nb(tdc['r2'], 2)}).",
        graphique={"titre": fig("Cotisants actifs précomptés (poste 936), par mois"), "type": "lignes", "categories": [e["libelle"] for e in EFF],
                   "series": [{"nom": "Cotisants", "valeurs": [e["cotisants"] for e in EFF], "couleur": BLEU}], "format": "# ##0", "etiquettes": False,
                   "axe_min": 20000, "legende": None, "unite": "En nombre de mutualistes"},
        analyse=[f"**Les cotisants actifs augmentent de {pc('Hausse cotisants', e1['cotisants'] / e0['cotisants'] - 1)} en 31 mois.** Le taux de croissance annuel moyen vaut {pc('TCAM cotisants', TC['Cotisants actifs (maladie 936)']['tcam'])}.",
                 f"**La hausse se fait par paliers en juillet, mois d'intégration des promotions de recrues.** Les fichiers comptent {N('Entrées juil 2024', EX['entrees_juillet']['2024']['entrees'])} entrées en juillet 2024 et {N('Entrées juil 2025', EX['entrees_juillet']['2025']['entrees'])} en juillet 2025.",
                 f"**La tendance est statistiquement significative.** Elle correspond à {N('Pente (texte)', tdc['pente'], 2)} cotisants de plus par mois en moyenne."],
        implication="Les recrues cotisent dès leur entrée, alors que leurs besoins de soins et d'épargne viendront plus tard : la croissance actuelle améliore les comptes de l'année, mais elle prépare des engagements plus importants à long terme, que seule une projection par âge permettra de chiffrer.",
        source=SRC_PREC + " (décembre 2024 : fichier vide)")
fl = R["flux_mensuels"]
contenu({"titre": "Entrées et sorties", "sous_titre": "Les entrées massives de juillet s'opposent à des sorties faibles et régulières",
         "synthese": f"Entrées : {N('Entrées 2024', R['flux_annuels'][0]['entrees'])} en 2024, {N('Entrées 2025', R['flux_annuels'][1]['entrees'])} en 2025 ; sorties : {N('Sorties 2024', R['flux_annuels'][0]['sorties'])} et {N('Sorties 2025', R['flux_annuels'][1]['sorties'])}"},
        "Les mois sans mois précédent disponible (janvier 2024, janvier 2025 après le fichier vide de décembre 2024) n'ont pas de flux calculé. Les sorties de novembre 2024 et janvier 2026 sont plus nombreuses : leur motif doit être demandé au FPPN (départs à la retraite de fin d'année probables, non vérifiés).",
        graphique={"titre": fig("Entrées et sorties de cotisants, par mois"), "type": "barres", "categories": [f["mois"] for f in fl],
                   "series": [{"nom": "Entrées", "valeurs": [f["entrees"] for f in fl], "couleur": BLEU}, {"nom": "Sorties", "valeurs": [f["sorties"] for f in fl], "couleur": ORANGE}],
                   "format": "# ##0", "etiquettes": False, "unite": "En nombre de mutualistes"},
        analyse=["**Les entrées sont concentrées en juillet, mois d'intégration des recrues.** Les autres mois, les entrées et les sorties restent faibles.",
                 "**Les sorties ne dépassent la moyenne qu'en novembre 2024 et en janvier 2026.** Leur motif n'est pas indiqué dans les fichiers.",
                 f"**La fidélisation des adhérents est forte.** Parmi les matricules précomptés en 2024, {pc('Fidélisation 2025 (flux)', RT['FID']['valeurs'][2025])} le sont encore en 2025, au-dessus de la norme d'au moins 80 % de la Décision n° 05/2022."],
        implication="Le motif de chaque sortie est indispensable pour séparer retraites, décès et radiations, et donc pour mesurer la mortalité des policiers et projeter les départs.",
        source=SRC_PREC)
anc = R["anciennete"]
contenu({"titre": "Structure des effectifs", "sous_titre": "Une population jeune en service, recrutée en majorité depuis moins de dix ans",
         "synthese": f"En {anc['mois']}, ancienneté médiane de {N('Ancienneté médiane', anc['mediane'], 2)} ans, salaire de base médian de {F('Salaire base médian', anc['salaire_base_median'])}"},
        "L'ancienneté est calculée à partir de la date de recrutement du fichier détaillé de janvier 2025. Elle ne donne pas l'âge : deux agents de même ancienneté peuvent avoir des âges différents. Pour cette raison, aucune hypothèse d'âge n'a été faite.",
        graphique={"titre": fig(f"Agents par ancienneté de service, {anc['mois']}"), "type": "barres", "categories": anc["categories"],
                   "series": [{"nom": "Agents", "valeurs": anc["valeurs"], "couleur": BLEU}], "format": "# ##0", "unite": "En nombre d'agents"},
        analyse=[f"**{pc('Part moins de 10 ans', sum(anc['valeurs'][:2]) / sum(anc['valeurs']))} des agents ont moins de dix ans de service.** Les recrutements récents dominent la population cotisante.",
                 f"**Une seconde vague a 15 à 25 ans de service.** Elle regroupe {pc('Part 15-25 ans', (anc['valeurs'][3] + anc['valeurs'][4]) / sum(anc['valeurs']))} des agents.",
                 "**Cette seconde vague partira à la retraite dans les années à venir.** Elle deviendra une population de retraités précomptés par la CGRAE, qui consomme davantage de soins."],
        implication="Les départs de cette vague déclencheront le versement de l'EFC et une hausse de la sinistralité des retraités : c'est le premier effet que la projection de l'équilibre devra chiffrer dès réception des âges.",
        source=SRC_PREC + " (fichiers maladie 936 détaillés)")
tx = EX["taux_cotisation"]
contenu({"titre": "Calcul pas à pas : taux de cotisation", "sous_titre": "Pour chaque agent, maladie et décès font 5,00 % du salaire de base",
         "synthese": f"Le taux maladie de 3,77 % se retrouve sur {N('Lignes taux (pas)', T936['lignes'])} lignes de huit mois ; le taux décès de 1,23 % est mesuré en avril 2024 et janvier 2025, puis confirmé par le rapport 8Z6 / 936 (partie 3)"},
        "Calcul pour le premier agent : 11 354 / 301 159 = 3,77 % (maladie) ; 3 704 / 301 159 = 1,23 % (décès) ; total 5,00 %. Le " + BAREME + " est de 3,5 % et 1,5 %. Le total est donc bon mais la répartition diffère. Matricules masqués.",
        pas_a_pas={"entetes": ["Matricule", "Grade", "Salaire de base", "Maladie 936", "Décès 8Z6", "Maladie %", "Décès %", "Total %"],
                   "lignes": [r[:5] + [r[5].replace(".", ",") + " %", r[6].replace(".", ",") + " %", r[7].replace(".", ",") + " %"] for r in tx["lignes"]],
                   "largeurs": [3.2, 2.2, 4.2, 3.8, 3.6, 3.4, 3.2, 3.2], "rouge": [(0, 5), (0, 6), (0, 7)],
                   "explications": ["Le précompte maladie vaut 11 354 / 301 159 = 3,77 % du salaire de base, au lieu de 3,5 %.", "Le précompte décès vaut 3 704 / 301 159 = 1,23 % du salaire de base, au lieu de 1,5 %.", "Le total de 5,00 % est respecté, mais la répartition entre maladie et décès diffère du barème indiqué."],
                   "fichier": "FPPN-SOLDE\\EFFECTIF POLICE (MALADIE)\\2024\\042024_effectif_police_936_AVRIL 2024.xlsx et FPPN_ALLOC_DECES_8Z6_AVRIL_2024.xlsx"},
        analyse=["**La Solde applique la même clé à tous les agents.** L'écart avec le barème n'est donc pas une erreur ponctuelle, mais une règle de calcul."],
        implication="Le barème officiel est fixé par les statuts (art. 17) et le règlement intérieur (art. 33) : la répartition de la Solde doit être corrigée ou expliquée, faute de quoi le résultat de chaque garantie dépend de la source retenue.",
        source=SRC_PREC + " ; matricules masqués")
tot25 = s[2025]["cot_maladie"] + s[2025]["cot_deces"]
contenu({"titre": "Taux de cotisation", "sous_titre": "Trois clés de répartition coexistent pour le même taux de 5 %",
         "synthese": f"Barème des statuts 3,5 / 1,5 ; Solde 3,77 / 1,23 ; comptabilité 2025 {nb(5 * s[2025]['cot_maladie'] / tot25, 2)} / {nb(5 * s[2025]['cot_deces'] / tot25, 2)}"},
        "La clé comptable est la part des cotisations maladie (compte 70100000) dans le total des cotisations maladie et décès (comptes 70100000 et 70210000) de 2025, multipliée par 5 %. Elle est proche du barème indiqué en entretien, pas de la clé de la Solde.",
        graphique={"titre": fig("Répartition du taux de 5 % selon la source"), "type": "barres_empilees",
                   "categories": ["Barème indiqué en entretien", "Précompte Solde", "Comptabilité 2025"],
                   "series": [{"nom": "Maladie", "valeurs": [3.5, round(100 * T936["mediane"], 2), round(5 * s[2025]["cot_maladie"] / tot25, 2)], "couleur": BLEU},
                              {"nom": "Décès", "valeurs": [1.5, round(5 - 100 * T936["mediane"], 2), round(5 * s[2025]["cot_deces"] / tot25, 2)], "couleur": ORANGE}],
                   "format": "0.00", "unite": "En % du salaire de base"},
        analyse=["**Le total de 5 % est respecté dans les trois sources.** Seule la répartition entre maladie et décès change.",
                 "**La Solde retient 3,77 % au poste maladie et 1,23 % au poste décès.** Les statuts (art. 17) et le règlement intérieur (art. 33) fixent 3,5 % et 1,5 %, comme l'avait indiqué le responsable des sinistres et prestations le 16/10/2026. Le responsable des cotisations renvoie la répartition à la comptabilité (entretien du 19/10/2026).",
                 f"**La comptabilité affecte {pc('Clé comptable maladie', SC_['cle_comptable_maladie'])} des cotisations à la maladie**, proche des 70 % du barème indiqué en entretien. Elle ne suit donc pas la Solde."],
        implication="La Décision n° 05/2022 impose de suivre l'excédent par risque : sans clé unique, le FPPN ne peut pas calculer la réserve obligatoire exigée pour chaque risque (au moins 30 % de l'excédent maladie, 50 % des autres).",
        source=SRC_PREC + " ; " + SRC_COMPTA[9:])
contenu({"titre": "Sensibilité : clé de répartition", "sous_titre": "La clé retenue change la sinistralité de chaque garantie de plusieurs points",
         "synthese": f"S/C maladie 2025 : {pc('S/C clé compta', SC_['sc_maladie_compta'])} avec la clé comptable, {pc('S/C clé Solde', SC_['sc_maladie_cle_solde'])} avec la clé de la Solde"},
        "Calcul : S/C maladie = prestations maladie / (cotisations totales × part maladie). Avec la clé de la Solde (75,40 %), la part des cotisations affectée à la maladie augmente, donc le ratio maladie baisse et le ratio décès monte. Les prestations ne changent pas : seule l'affectation des cotisations change.",
        graphique={"titre": fig("Ratio S/C 2025 selon la clé de répartition"), "type": "barres", "categories": ["Clé comptable", "Clé de la Solde (3,77 / 1,23)", "Barème des statuts (3,5 / 1,5)"],
                   "series": [{"nom": "S/C maladie", "valeurs": [round(100 * SC_[k], 2) for k in ("sc_maladie_compta", "sc_maladie_cle_solde", "sc_maladie_cle_bareme")], "couleur": BLEU},
                              {"nom": "S/C décès", "valeurs": [round(100 * SC_[k], 2) for k in ("sc_deces_compta", "sc_deces_cle_solde", "sc_deces_cle_bareme")], "couleur": ORANGE}],
                   "format": "0.00", "unite": "En %"},
        analyse=[f"**Avec la clé de la Solde, le S/C maladie 2025 passe à {pc('S/C solde (t)', SC_['sc_maladie_cle_solde'])}** et le S/C décès monte à {pc('S/C décès solde', SC_['sc_deces_cle_solde'])}.",
                 f"**Avec le barème indiqué en entretien, le S/C maladie vaut {pc('S/C barème', SC_['sc_maladie_cle_bareme'])}.** Les prestations restent les mêmes : seule l'affectation des cotisations change.",
                 "**Le résultat par garantie dépend donc d'une décision de gestion.** Cette décision doit être écrite dans le barème et appliquée de la même façon par la Solde et par la comptabilité."],
        implication="Selon la clé, la maladie se situe dans la fourchette réglementaire de 70 % à 80 % ou en dessous : la conclusion sur l'équilibre de la garantie maladie ne sera ferme qu'avec le barème officiel.",
        source=SRC_COMPTA)
contenu({"titre": "Cotisations en termes réels", "sous_titre": "Hors inflation, les cotisations progressent de 23 % entre 2020 et 2025",
         "synthese": f"Hausse nominale {pc('Croissance nominale cot', C['cot_croissance_nominale_2020_2025'])}, hausse des prix {pc('Croissance prix', C['prix_croissance_2020_2025'])}, hausse réelle {pc('Croissance réelle cot', C['cot_croissance_reelle_2020_2025'])}"},
        "Méthode : indice des prix base 100 en 2020, construit en cumulant l'inflation annuelle de la Banque mondiale (2021 : 4,16 % ; 2022 : 5,23 % ; 2023 : 4,37 % ; 2024 : 3,45 % ; 2025 : 0,13 %). Cotisations réelles = cotisations nominales × 100 / indice. Elles mesurent le pouvoir d'achat des cotisations en francs de 2020.",
        graphique={"titre": fig("Cotisations nominales et réelles (francs de 2020)"), "type": "lignes", "categories": [str(a) for a in range(2020, 2026)],
                   "series": [{"nom": "Nominales", "valeurs": [round(r["cotisations_nominales"]) for r in C["reel"]], "couleur": BLEU},
                              {"nom": "Réelles (francs de 2020)", "valeurs": [round(r["cotisations_reelles_2020"]) for r in C["reel"]], "couleur": ORANGE}],
                   "format": FMT_M, "unite": "En millions de FCFA"},
        analyse=[f"**L'inflation explique une partie de la hausse des cotisations.** Les prix ont augmenté de {pc('Prix (t)', C['prix_croissance_2020_2025'])} entre 2020 et 2025.",
                 f"**La hausse réelle reste nette : {pc('Réelle (t)', C['cot_croissance_reelle_2020_2025'])}.** Elle vient surtout de la hausse des effectifs, le taux de 5 % étant resté le même.",
                 f"**L'inflation de 2025 est très faible ({pcv('Inflation 2025', INF[2025])})**, après {pcv('Inflation 2022', INF[2022])} en 2022."],
        implication="Le coût des soins suit les prix de la santé : une cotisation fixée en pourcentage du salaire ne protège le FPPN que si les salaires suivent ces prix. La projection devra tester cette hypothèse.",
        source=SRC_BM)
cot25 = an(R["cotisations"], 2025)
contenu({"titre": "Cotisations comptabilisées", "sous_titre": "Les cotisations comptabilisées progressent de 45,76 % entre 2020 et 2025",
         "synthese": f"{F('Cot 2020', s[2020]['cot_maladie'] + s[2020]['cot_deces'])} en 2020, {F('Cot 2025 (cc)', tot25)} en 2025 ; taux de croissance annuel moyen {pc('TCAM cot', TC['Cotisations maladie et décès']['tcam'])}"},
        "Rapprochement 2025 : cotisations maladie comptabilisées comparées aux précomptes maladie (12 mois) et retraités CGRAE (12 mois). Les cotisations à recevoir du Trésor (compte 41111000) mesurent ce que l'État doit encore au FPPN à la clôture.",
        graphique={"titre": fig("Cotisations comptabilisées par garantie"), "type": "barres_empilees", "categories": cats_ex,
                   "series": [{"nom": "Maladie", "valeurs": [round(s[a]["cot_maladie"]) for a in range(2020, 2027)], "couleur": BLEU},
                              {"nom": "Décès", "valeurs": [round(s[a]["cot_deces"]) for a in range(2020, 2027)], "couleur": ORANGE}],
                   "format": FMT_M, "unite": "En millions de FCFA"},
        analyse=[f"**En 2025, comptabilité et précomptes maladie diffèrent de {pc('Écart maladie 2025', 1 - s[2025]['cot_maladie'] / (cot25['prec_936'] + cot25['prec_cgrae']))}.** La comptabilité enregistre {F('Compta mal 2025', s[2025]['cot_maladie'])}, contre {F('Précomptes mal 2025', cot25['prec_936'] + cot25['prec_cgrae'])} de précomptes.",
                 "**L'écart vient de la clé de répartition.** La comptabilité affecte au décès une part des cotisations plus grande que la Solde.",
                 f"**Fin 2025, l'État devait {F('Créance Trésor 2025', b[2025]['cotisations_a_recevoir_etat'])} de cotisations.** Le compte est soldé au 31/08/2026."],
        implication="Les cotisations totales concordent, mais leur partage entre maladie et décès diffère selon la source. Tant que cet écart n'est pas réglé, la sinistralité et l'excédent de chaque risque restent incertains.",
        source=SRC_COMPTA)
contenu({"titre": "Règles des produits", "sous_titre": "Les fichiers confirment les montants de l'EFC et du PEP ; les textes reçus le 22/10/2026 ne sont pas adoptés formellement",
         "synthese": "Maladie et décès en pourcentage du salaire de base, EFC et PEP en montants fixes, sans règle écrite transmise"},
        "Colonne « règle déclarée » : déclaration du responsable des cotisations, entretien du 19/10/2026. Colonne « fichiers » : fichiers retour de janvier 2020 à août 2026 reçus le 15/10/2026, part des lignes de chaque montant.",
        tableau={"titre": tab("Règles des produits déclarées et confrontées aux fichiers de précomptes"), "entetes": ["Produit", "Règle déclarée (19/10/2026)", "Ce que montrent les fichiers", "Pièce demandée"], "largeurs": [3.6, 10.0, 10.0, 7.6], "taille": 9,
                 "lignes": [["Maladie (936) et décès (8Z6)", "Pourcentage du salaire de base. La répartition entre maladie et décès relève de la comptabilité.", "La Solde retient 3,77 % et 1,23 % du salaire de base, soit 5,00 % au total.", "Barème officiel et document de notification à la Solde (PBC, item 56)."],
                            ["EFC (870)", "6 000 FCFA par mois avant 2021, 10 000 FCFA pour les souscriptions depuis 2021. Les anciens souscripteurs restent à 6 000 FCFA et peuvent passer à 10 000 FCFA.", "10 000 FCFA : " + pc("EFC 10000 2023", EFC10[2023]) + " des lignes en 2023 et " + pc("EFC 10000 2026", EFC10[2026]) + " en 2026. 6 000 FCFA : " + pc("EFC 6000 2023", EFC6[2023]) + " puis " + pc("EFC 6000 2026", EFC6[2026]) + ".", "Décision de 2021 (PBC, item 58)."],
                            ["PEP (998)", "À partir de 10 000 FCFA, par paliers de 5 000 FCFA. Durée d'au moins trois ans (règlement intérieur, art. 66).", pc("PEP multiples", PEP_M5) + " des montants sont des multiples de 5 000 FCFA et " + pc("PEP sous 10000", PEP_LT10) + " sont inférieurs à 10 000 FCFA. Montant maximal : " + F("PEP max", PEP_MAX) + ".", "Durée des plans et conditions de retrait (PBC, item 57)."],
                            ["EFC et PEP", "[[Présentés aux élèves policiers comme obligatoires dès l'école, sans texte.]]", "Non mesurable dans les fichiers.", "Texte qui fonde l'obligation (aucun reçu)."],
                            ["Prestations", "[[Au moins trois mois de cotisation avant de pouvoir bénéficier des prestations.]]", "Non mesurable sans base des prestations.", "Texte et produits concernés (PBC, item 59). L'entretien du 16/10/2026 indiquait l'absence de délai de carence."]]},
        disposition="dessous",
        analyse=["**Les fichiers confirment les déclarations sur l'EFC et le PEP.** Le cahier des charges EFC et PEP de 2021 et le manuel des produits d'épargne, reçus le 22/10/2026, décrivent ces produits sans décision d'adoption signée (partie 6). La part de l'EFC à 10 000 FCFA augmente chaque année, ce qui correspond à l'arrivée de nouveaux souscripteurs depuis 2021."],
        implication="Les engagements d'épargne se calculeront sur les montants réellement précomptés. Le règlement de chaque produit reste indispensable pour connaître les échéances et les conditions de retrait.",
        source=SRC_ENT19 + " ; " + SRC_PREC[9:])
contenu({"titre": "Nouvelles recrues", "sous_titre": "Les recrues souscrivent avant d'avoir un mécano, ce qui retarde le premier précompte",
         "synthese": "Le mécano est reporté sur les fiches de souscription en masse, après son attribution"},
        "Logigramme du parcours d'une recrue, de l'école de police au premier précompte, tel que décrit par le responsable des cotisations le 19/10/2026. Cadre rouge : étape à risque. Les formes sont modifiables dans PowerPoint.",
        logigramme="recrues",
        analyse=["[[**Les nouvelles recrues souscrivent avant d'avoir un mécano.** Les élèves policiers n'ont qu'un matricule. Le FPPN attend l'attribution des mécanos, puis reporte chaque mécano sur les fiches de souscription avant de les inscrire au fichier aller. Le premier précompte peut être décalé d'un mois]] (" + ENT19 + ").",
                 "**Le report en masse des mécanos est une étape à risque.** Il peut produire des erreurs d'appariement, et donc des précomptes sur la mauvaise personne. Il crée un retard entre la souscription et le premier versement."],
        implication="Une table de correspondance entre matricule et mécano, fournie par la direction du personnel de la Police, éviterait le report manuel et sécuriserait le premier précompte (PBC, item 67).",
        source=SRC_ENT19)
retenir("effectifs et cotisations", ("Les cotisations augmentent avec les effectifs, mais leur répartition doit être clarifiée", "Effectifs en hausse d'un tiers, taux total de 5 % respecté, trois clés de répartition différentes",
                                     "Le FPPN encaisse davantage chaque année grâce aux recrues, mais il applique trois clés différentes pour partager les cotisations entre maladie et décès."),
        f"Le FPPN compte {N('Cotisants (ret)', e1['cotisants'])} cotisants en août 2026. La Solde retient 3,77 % et 1,23 %, alors que les statuts (art. 17) prévoient 3,5 % et 1,5 %.",
        "Le résultat de chaque garantie et la réserve obligatoire par risque dépendent d'une clé qui n'est fixée par aucun document transmis.",
        "Le FPPN doit transmettre le barème officiel, les dates de naissance et les motifs de sortie des agents.",
        "La question de la clé de répartition est la plus importante de cette partie : elle change le résultat de chaque garantie.")

# ---------------------------------------------------------------- PARTIE 5 : SINISTRALITÉ
separateur(5, "Sinistralité et résultat", "Prestations, frais de gestion, résultat")
methode_gab({"titre": "Méthode : sinistralité comptable", "sous_titre": "Le ratio S/C rapporte les prestations de l'année aux cotisations de l'année",
             "synthese": "S/C = prestations comptabilisées / cotisations comptabilisées, par garantie et par exercice"},
            [("Formule", "Le S/C d'une garantie rapporte ses prestations comptabilisées de l'année à ses cotisations comptabilisées de la même année, de 2020 à 2026."),
             ("Prestations maladie", "Adenka retient les comptes 601, 6013, 6018 et 60800007, soit douze comptes de prestations maladie en 2025."),
             ("Prestations décès", "Adenka retient les comptes 6021 et 6024. En 2020, la prime d'assurance décès (65180000) s'y ajoute, car un assureur portait alors la garantie."),
             ("Norme et limite", "La Décision n° 05/2022 fixe une sinistralité entre 70 % et 80 %. Le ratio comptable mélange les soins de l'année et des factures d'années antérieures.")],
            "Le ratio comptable mélange les soins de l'année et les factures d'années antérieures payées ou comptabilisées dans l'année. Le vrai ratio (par année de survenance des soins) demande la base des prestations avec la date des soins.")
sc25 = EX["sc_2025"]
contenu({"titre": "Calcul pas à pas : S/C maladie 2025", "sous_titre": "Douze comptes de prestations divisés par un compte de cotisations",
         "synthese": f"{F('Prest mal 2025 (pas)', s[2025]['prest_maladie'])} / {F('Cot mal 2025 (pas)', s[2025]['cot_maladie'])} = {pc('S/C 2025 (pas)', s[2025]['sc_maladie'])}"},
        "Le compte 60120000 « Charges de prestations maladie » porte à lui seul plus de la moitié du total : il regroupe des prestations sans détail de nature, ce qui empêche l'analyse par type de soins. Montants exacts lus dans la balance 2025.",
        pas_a_pas={"entetes": ["Compte", "Intitulé", "Montant 2025"], "lignes": [[c[0], c[1], nb(float(c[2])) + " FCFA"] for c in sc25["prestations"]] + [["Total", "Prestations maladie", F("Total prest (pas)", s[2025]["prest_maladie"])], [sc25["cotisations"][0][0], sc25["cotisations"][0][1], nb(-float(sc25["cotisations"][0][2])) + " FCFA"]],
                   "largeurs": [3.4, 11.0, 6.0], "rouge": [(7, 2), (12, 2), (13, 2)], "petit": True,
                   "explications": ["Le compte global 60120000 porte plus de la moitié des prestations, sans détail de nature.", "Les douze comptes de prestations maladie totalisent ce montant.", "Les cotisations maladie acquises forment le dénominateur du ratio."],
                   "fichier": "AUDIT ADENKA\\DAFC\\2025\\BALANCE FPPN 2025.xlsx"},
        analyse=["**Appliquée à la seule maladie, la fourchette réglementaire de 70 % à 80 % est respectée en 2025.** Le compte global empêche pourtant de savoir quels soins font monter le ratio."],
        implication="Le FPPN doit ventiler le compte 60120000 par nature de soins et par prestataire, ce qui permettra de cibler les mesures de maîtrise des dépenses.",
        source="Source : balance des comptes 2025 (Sage 100) ; traitements Adenka")
contenu({"titre": "Sinistralité comptable", "sous_titre": "Les prestations maladie absorbent presque toutes les cotisations maladie",
         "synthese": f"S/C maladie : {pc('S/C 2021 (s)', s[2021]['sc_maladie'])} en 2021, {pc('S/C 2025 (s)', s[2025]['sc_maladie'])} en 2025, {pc('S/C 2026 (s)', s[2026]['sc_maladie'])} sur huit mois de 2026"},
        f"Test de tendance (moindres carrés sur 2020 à 2025) : le S/C maladie progresse de {N('Pente S/C', C['tendances'][1]['pente'], 2)} points par an (IC 95 % : {N('IC bas S/C', C['tendances'][1]['ic95_bas'], 2)} à {N('IC haut S/C', C['tendances'][1]['ic95_haut'], 2)} ; p-valeur {nb(C['tendances'][1]['p_valeur'], 4)}). Avec six points seulement, la tendance est significative au seuil de 5 % mais son rythme est imprécis.",
        graphique={"titre": fig("Prestations comptabilisées / cotisations (S/C)"), "type": "lignes", "categories": cats_ex,
                   "series": [{"nom": "Maladie", "valeurs": [round(100 * s[a]["sc_maladie"], 2) for a in range(2020, 2027)], "couleur": BLEU},
                              {"nom": "Décès (prime d'assurance incluse en 2020)", "valeurs": [round(100 * s[a]["sc_deces_avec_prime"], 2) for a in range(2020, 2027)], "couleur": ORANGE}],
                   "format": "0.00", "unite": "En % des cotisations de la garantie"},
        analyse=[f"**Le ratio maladie a été multiplié par {N('Mult S/C', s[2025]['sc_maladie'] / s[2021]['sc_maladie'], 2)} entre 2021 et 2025.** Il dépasse la fourchette réglementaire sur les huit premiers mois de 2026.",
                 "**Les ratios de 2020 à 2022 sont bas et s'expliquent par les textes et les comptes.** Le règlement intérieur limite la couverture aux soins liés à une hospitalisation (art. 41 et 45). Sur 2023 à 2025, seules années mesurables, moins de 1 % des factures d'une année sont enregistrées l'année suivante (" + ref("Provision maladie : factures arrivées en retard") + ").",
                 f"**En décès, le ratio oscille entre {pc('Min S/C décès', min(s[a]['sc_deces'] for a in range(2022, 2027)))} et {pc('Max S/C décès', max(s[a]['sc_deces'] for a in range(2022, 2027)))} depuis 2022.** La garantie décès reste largement excédentaire."],
        implication="À ce rythme, les prestations maladie dépasseront les cotisations maladie : sans mesure de maîtrise ou révision du barème, la garantie maladie deviendra déficitaire et pèsera sur la réserve obligatoire.",
        source=SRC_COMPTA)
# ---- v5 : choix de la formule de sinistralité (partie 5)
contenu({"titre": "Choix de la formule : sinistralité", "sous_titre": "Le texte fixe un taux global, le rapport ajoute un taux par garantie",
         "synthese": "La méthode du FPPN, communiquée le 22/10/2026, ignore les provisions et applique la norme globale à chaque garantie"},
        "Un taux de sinistralité répond à une question simple : sur 100 FCFA de cotisations, combien repartent en prestations ? Les formules diffèrent par ce qu'elles mettent au numérateur (quelles prestations, avec ou sans provisions) et au dénominateur (quelles cotisations). Le texte de référence est la Décision n° 05/2022, § 1.2.1.2.5.",
        tableau={"titre": tab("Formules du taux de sinistralité comparées"), "entetes": ["Méthode", "Formule", "Ce qu'elle mesure", "Norme", "Statut"], "largeurs": [4.2, 8.4, 8.6, 4.4, 5.6], "taille": 9,
                 "lignes": [["Texte (Décision n° 05/2022)", "(60 + 6901 - 7901) / (70 + 7906 - 6906)", "Toutes les prestations de l'année, variation des provisions comprise, sur toutes les cotisations acquises.", "70 % à 80 %", "Conforme au texte"],
                            ["Méthode du FPPN", "Prestations payées / cotisations payées, par garantie et global, sans provisions", "Communiquée le 22/10/2026 (étude des ratios) : 12,95 % en décès, 99,23 % en maladie, 78,49 % au total au 30/06/2026.", "70 % à 80 % appliqués à chaque garantie", "Divergent du texte"],
                            ["Par garantie (rapport)", "Maladie : (601 + 6013 + 6018 + 60800007) / 7010. Décès : (6021 + 6024) / 7021", "La part des cotisations d'une garantie consommée par ses propres prestations.", "Aucune (norme globale)", "Divergent, justifié"]]},
        disposition="dessous",
        analyse=["**Le taux par garantie n'est pas un choix libre : il découle du texte.** La Décision n° 05/2022 (§ 1.1.1) impose de doter la réserve obligatoire d'au moins 30 % de l'excédent maladie et 50 % de l'excédent des autres risques. Il faut donc connaître le résultat de chaque garantie, et donc son taux de sinistralité.",
                 "**Les deux formules du texte donnent ici le même résultat.** Le FPPN n'a passé aucune dotation ni reprise de provision technique (comptes 6901, 7901, 7906 et 6906 absents) : la variation des provisions vaut zéro chaque année."],
        implication="Seul le taux global est comparé à la norme de 70 % à 80 %. Le taux par garantie sert à calculer l'excédent de chaque risque. La méthode du FPPN est détaillée et chiffrée dans la diapositive « Méthode de sinistralité du FPPN ».",
        source="Source : " + DEC + ", § 1.1.1 et § 1.2.1.2.5 ; " + SRC_COMPTA[9:])
PS = {r["exercice"]: r for r in PV["sinistralite"]}
contenu({"titre": "Sinistralité selon la formule", "sous_titre": "Formule du texte et taux par garantie racontent la même hausse, à des niveaux différents",
         "synthese": f"2025 : {pc('Variante texte 2025', PS[2025]['taux_texte'])} selon le texte, {pc('Variante maladie 2025', PS[2025]['taux_maladie'])} en maladie, {pc('Variante décès 2025', PS[2025]['taux_deces'])} en décès"},
        "Période : les taux sont calculés de 2020 au 31/08/2026 pour montrer la tendance. La conformité n'est jugée que sur 2022 à 2025 : la Décision n° 05/2022 s'applique depuis le 1er juin 2022, et 2026 ne couvre que huit mois, sur une balance provisoire, sans annualisation. 2020, 2021 et 2026 sont donc indicatifs.",
        graphique={"titre": fig("Taux de sinistralité selon la formule, 2020 au 31/08/2026"), "type": "lignes", "categories": cats_ex,
                   "series": [{"nom": "Texte (global)", "valeurs": [round(100 * PS[a]["taux_texte"], 2) for a in range(2020, 2027)], "couleur": NOIR},
                              {"nom": "Maladie", "valeurs": [round(100 * PS[a]["taux_maladie"], 2) for a in range(2020, 2027)], "couleur": BLEU},
                              {"nom": "Décès", "valeurs": [round(100 * PS[a]["taux_deces"], 2) for a in range(2020, 2027)], "couleur": ORANGE},
                              {"nom": "Texte sans bons de prise en charge", "valeurs": [round(100 * PS[a]["taux_sans_bons"], 2) for a in range(2020, 2027)], "couleur": GRIS_C}],
                   "format": "0.00", "etiquettes": False, "unite": "En % des cotisations"},
        analyse=[f"**Le taux global du texte reste sous la fourchette de 70 % à 80 % de 2022 à 2025** : il passe de {pc('Variante texte 2022', PS[2022]['taux_texte'])} en 2022 à {pc('Variante texte 2025 (b)', PS[2025]['taux_texte'])} en 2025. Il atteint {pc('Variante texte 2026', PS[2026]['taux_texte'])} sur huit mois de 2026.",
                 "**Le taux global est plus bas que le taux maladie, car il mélange une garantie maladie chargée et une garantie décès excédentaire.** La moyenne masque le déséquilibre de la maladie.",
                 f"**Retirer les bons de prise en charge (60800007, 60800008) ou les cotisations MEDOC (7022) change peu le taux** : {pc('Variante sans bons 2025', PS[2025]['taux_sans_bons'])} sans les bons et {pc('Variante sans MEDOC 2025', PS[2025]['taux_sans_medoc'])} sans MEDOC en 2025."],
        implication="La conclusion ne dépend pas du choix entre ces variantes : la maladie se rapproche du niveau de ses cotisations, quelle que soit la formule. En revanche, un taux global sous la norme ne prouve pas que chaque garantie est équilibrée.",
        source="Source : " + DEC + ", § 1.2.1.2.5 ; " + SRC_COMPTA[9:])
# ---- v9 : méthode de sinistralité du FPPN (communiquée le 22/10/2026) et mortalité brute (partie 5)
SRC_RATIO = "Source : ETUDE DES RATIO MALADIE DECES 2.xlsx (feuilles Tableau, Notes et Historique 2015-2021) et INTERPRETATION DES RATIOS MALADIE MATERNITE ET DECES INVALIDITE.pdf, reçus le 22/10/2026"
_sin26 = RT["SIN"]["valeurs"].get(2026)
contenu({"titre": "Méthode de sinistralité du FPPN", "sous_titre": "Le FPPN rapporte les prestations payées aux cotisations payées, sans provisions, garantie par garantie",
         "synthese": "Au 30/06/2026, le FPPN calcule 12,95 % en décès, 99,23 % en maladie et 78,49 % pour l'ensemble ; sa propre note écrit aussi 19,66 % et 107 %"},
        "La méthode du FPPN figure dans la feuille Notes de son étude des ratios. Elle est comparée à la formule de la Décision n° 05/2022. Les deux méthodes répondent à la même question (part des cotisations consommée par les prestations), mais pas avec les mêmes montants.",
        tableau={"titre": tab("Méthode du FPPN et méthode réglementaire au 30/06/2026"), "entetes": ["Élément", "Méthode du FPPN", "Méthode réglementaire (Décision n° 05/2022)"], "largeurs": [5.2, 13.0, 13.0], "taille": 8.5,
                 "lignes": [["Numérateur", "Prestations payées : allocations décès et frais funéraires au barème ; prestations maladie de l'état des hospitalisations et consultations", "Charges de prestations (comptes 60) et variation des provisions techniques (6901 - 7901)"],
                            ["Dénominateur", "Cotisations payées : policiers en activité, civils de la Direction Générale, retraités", "Cotisations acquises (comptes 70) et variation des provisions sur cotisations (7906 - 6906)"],
                            ["Décès", "55 600 000 / 429 215 718 = 12,95 % (« TSD = 13 % »), la même note écrit aussi « TS = 19, 66 % »", "Pas de taux par garantie dans le texte ; taux par garantie du rapport : partie 5"],
                            ["Maladie", "1 345 906 033 / 1 356 344 706 = 99,23 % (« TSM »), la même note écrit aussi « TS = 107 % »", "Pas de taux par garantie dans le texte"],
                            ["Ensemble", "1 401 506 033 / 1 785 560 424 = 78,49 % (« TSP »)", "Taux global au 31/08/2026 : " + (pc("SIN 2026 (méthode)", _sin26) if _sin26 is not None else "non calculé") + " (balance provisoire, huit mois)"],
                            ["Norme", "Le FPPN compare chaque taux à 70 % à 80 %", "70 % à 80 % pour le seul taux global"]]},
        disposition="dessous",
        analyse=["**La méthode du FPPN ignore les provisions et compare chaque garantie à la norme globale.** La norme de 70 % à 80 % de la Décision s'applique au seul taux global, provisions comprises.",
                 "**La note du FPPN contient deux taux différents pour le décès et pour la maladie.** Les quotients écrits donnent 12,95 % et 99,23 % ; les mentions 19,66 % et 107 % ne correspondent à aucun calcul affiché."],
        implication="La méthode du FPPN est désormais connue : le registre des formules (onglet 11_formules) la reprend. Elle convient au pilotage par garantie, pas à la conformité, qui exige la formule du texte avec les provisions.",
        valable="période du 01/01 au 30/06/2026 pour les taux du FPPN, du 01/01 au 31/08/2026 pour le taux réglementaire",
        source=SRC_RATIO + " ; " + DEC + ", § 1.2.1.2.5 ; " + SRC_COMPTA[9:])
_sc = V9["sinistralite_comparee"]
contenu({"titre": "Sinistralité décès : le FPPN et les comptes", "sous_titre": "De 2022 à 2025, les cotisations et les sinistres du FPPN sont ceux de la comptabilité",
         "synthese": "Le rapport sinistres / cotisations du FPPN égale le taux décès par garantie du rapport de 2022 à 2025 ; 2020 et 2021 diffèrent"},
        "Le tableau compare, année par année, les montants de l'étude du FPPN (feuille Tableau) et ceux des balances : cotisations décès (compte 7021) et prestations décès (comptes 6021 et 6024). En 2020, un assureur portait la garantie : la prime d'assurance décès (compte 65180000) remplace les prestations.",
        tableau={"titre": tab("Sinistralité décès selon le FPPN et selon les comptes"), "entetes": ["Année", "Cotisations (FPPN)", "Sinistres (FPPN)", "S/P du FPPN", "Cotisations décès (comptes)", "Prestations décès (comptes)", "S/C (comptes)"],
                 "largeurs": [2.6, 5.0, 4.8, 3.4, 5.4, 5.4, 3.4], "taille": 9,
                 "lignes": [[str(x["annee"]), F(f"FPPN cot déc {x['annee']}", x["fppn_cotisations"]), F(f"FPPN sin déc {x['annee']}", x["fppn_sinistres"]), pc(f"FPPN S/P {x['annee']}", x["fppn_sp"]),
                             F(f"Compta cot déc {x['annee']}", x["compta_cot_deces"]), F(f"Compta prest déc {x['annee']}", x["compta_prest_deces"] + x["compta_prime_deces"]),
                             pc(f"Compta S/C déc {x['annee']}", x["compta_sc_deces_avec_prime"])] for x in _sc]},
        disposition="dessous",
        analyse=["**De 2022 à 2025, les deux sources donnent les mêmes cotisations, les mêmes sinistres et donc le même taux.** Le taux décès par garantie du rapport est ainsi confirmé par le FPPN.",
                 f"**En 2021, le FPPN retient {F('FPPN sin 2021 (t)', next(x['fppn_sinistres'] for x in _sc if x['annee'] == 2021))} de sinistres, contre {F('Compta prest 2021 (t)', next(x['compta_prest_deces'] for x in _sc if x['annee'] == 2021))} dans les comptes.** En 2020, ses cotisations et ses sinistres diffèrent aussi des comptes, l'assureur portant alors le risque."],
        implication="Pour 2022 à 2025, la garantie décès consomme moins d'un quart de ses cotisations selon les deux sources. L'écart de 2021 est demandé au FPPN (questions au FPPN, partie 14).",
        valable="garantie décès-invalidité, exercices 2020 à 2025 ; prestations 2020 des comptes = obsèques " + F("D20 obs", V11["deces_2020"]["obseques"]) + " + prime d'assurance décès " + F("D20 prime", V11["deces_2020"]["prime_assurance"]) + " (compte 65180000)",
        source=SRC_RATIO + " ; " + SRC_COMPTA[9:])
_mb = V9["mortalite_brute"]
_m25 = next(x for x in _mb if x["annee"] == 2025)
contenu({"titre": "Mortalité brute des membres participants", "sous_titre": "Entre trois et cinq décès pour mille adhérents chaque année depuis 2021",
         "synthese": f"En 2025 : {N('Décès adh 2025', _m25['deces_adherents'])} décès pour {N('Adh 2025', _m25['adherents'])} adhérents, soit {N('Taux brut 2025', 1000 * _m25['taux_brut'], 2)} pour mille"},
        "Le taux brut rapporte les décès de membres participants de l'année au nombre d'adhérents retenu par le FPPN. Il ne tient pas compte de l'âge : faute de dates de naissance des mutualistes, il ne peut pas être comparé aux décès attendus par la table CIMA H (rapport A/E de la partie 12).",
        tableau={"titre": tab("Décès et taux brut de mortalité des membres participants"), "entetes": ["Année", "Adhérents", "Décès de membres", "Décès de conjoints", "Décès d'enfants", "Taux brut (pour mille)"],
                 "largeurs": [3.0, 5.0, 5.0, 5.0, 5.0, 5.4], "taille": 9,
                 "lignes": [[str(x["annee"]), N(f"Adh {x['annee']}", x["adherents"]), N(f"Déc adh {x['annee']}", x["deces_adherents"]), N(f"Déc conj {x['annee']}", x["deces_conjoints"]),
                             N(f"Déc enf {x['annee']}", x["deces_enfants"]), N(f"Taux brut {x['annee']}", 1000 * x["taux_brut"], 2)] for x in _mb]},
        disposition="dessous",
        analyse=[f"**Le taux brut varie de {N('Taux brut min', 1000 * min(x['taux_brut'] for x in _mb if x['annee'] >= 2021), 2)} à {N('Taux brut max', 1000 * max(x['taux_brut'] for x in _mb if x['annee'] >= 2021), 2)} pour mille de 2021 à 2025.** Le FPPN reporte le même nombre d'adhérents trois années de suite de 2016 à 2018, puis en 2021 et 2022, ce qui fragilise les taux de ces années.",
                 f"**Les décès d'ayants droit (conjoints et enfants) passent de {N('Déc AD 2020', next(x['deces_conjoints'] + x['deces_enfants'] for x in _mb if x['annee'] == 2020))} en 2020 à {N('Déc AD 2025', _m25['deces_conjoints'] + _m25['deces_enfants'])} en 2025.** Le nombre d'ayants droit couverts n'est pas connu : aucun taux ne peut être calculé pour eux."],
        implication="Le taux brut sert d'ordre de grandeur. La comparaison à la table CIMA H, nécessaire pour juger la cotisation décès, demande les dates de naissance des membres et la base des décès (PBC, items 9 à 11 et 28).",
        valable="membres participants recensés par le FPPN, années 2015 à 2025 ; ne vaut ni par âge, ni par grade, ni pour les ayants droit",
        source=SRC_RATIO)
contenu({"titre": "Nature des risques", "sous_titre": "Chaque garantie relève d'une catégorie de risque qui fixe ses provisions",
         "synthese": "Les textes de la mutualité distinguent risques courts et longs, la CIMA distingue vie et non-vie"},
        "Risque court : couvert pour une durée d'un an au plus, réglé rapidement (règlement d'exécution n° 004/2022). Risque long : engagement sur plusieurs années (épargne, rentes). La classification CIMA (vie, non-vie) sert de référence technique pour les méthodes de provision, faute de méthode dans les textes UEMOA.",
        tableau={"titre": tab("Nature des risques par garantie et provisions attendues"), "entetes": ["Garantie", "UEMOA", "CIMA (référence)", "Provision attendue", "Effet sur la sinistralité"], "largeurs": [5.2, 3.6, 4.4, 9.0, 9.0], "taille": 9,
                 "lignes": [["Maladie des actifs", "Risque court", "Non-vie", "PSAP dossier par dossier et sinistres tardifs (art. 10)", "Le taux doit intégrer la variation de la PSAP."],
                            ["Décès (capital, frais funéraires, allocations)", "Risque court si annuel", "Vie (décès temporaire)", "PSAP des dossiers déclarés non payés et tardifs", "Le taux doit intégrer la variation de la PSAP obsèques."],
                            ["EFC, PEP", "Risque long", "Vie (épargne)", "Dette d'épargne, provision mathématique si un capital est garanti (art. 11)", "Hors sinistralité : versements et retraits d'épargne."],
                            ["Santé des retraités", "Risque long si garantie à vie", "Vie (rente de soins)", "Provision mathématique (art. 11)", "Le taux annuel sous-estime l'engagement futur."],
                            ["Flotte automobile (GNA)", "Hors champ du FPPN", "Non-vie (porté par GNA)", "Aucune provision technique au FPPN", "Hors sinistralité du FPPN."],
                            ["MEDOC", "Risque court", "Non-vie", "PSAP si le FPPN porte le risque", "Cotisations 7022 incluses dans le taux du texte."]]},
        disposition="dessous",
        analyse=["**Le taux de sinistralité ne mesure bien que les risques courts.** Pour les risques longs, l'engagement se mesure par une provision mathématique, calculée avec une table de mortalité et un taux d'intérêt technique (règlement d'exécution n° 004/2022, art. 11)."],
        implication="Les provisions du FPPN doivent être revues garantie par garantie : seules la maladie et le décès ont une provision, figée depuis 2020 (partie 9).",
        source="Source : " + REX + ", art. 8 à 15 ; Code CIMA (référence technique) ; Adenka")

tlp = C["tendances"][2]
contenu({"titre": "Croissance des prestations", "sous_titre": "Les prestations maladie croissent beaucoup plus vite que les cotisations",
         "synthese": f"Taux de croissance annuel moyen 2021 à 2025 : prestations maladie {pc('TCAM prest mal', TC['Prestations maladie']['tcam'])}, cotisations {pc('TCAM cot (cr)', TC['Cotisations maladie et décès']['tcam'])}"},
        f"Le taux de croissance annuel moyen (TCAM) est le taux constant qui mène de la valeur de départ à la valeur d'arrivée : TCAM = (V_fin / V_début)^(1/n) - 1. Une régression du logarithme des prestations sur le temps donne une croissance annuelle moyenne de {pc('Croissance log', tlp['croissance_annuelle'])} (IC 95 % : {pc('IC bas log', tlp['croissance_ic95_bas'])} à {pc('IC haut log', tlp['croissance_ic95_haut'])}), estimation plus robuste aux valeurs extrêmes que le TCAM.",
        tableau={"titre": tab("Taux de croissance annuels moyens"), "entetes": ["Indicateur", "Période", "Valeur de départ", "Valeur d'arrivée", "Multiplié par", "TCAM"], "largeurs": [7.6, 4.6, 5.6, 5.6, 3.6, 4.2],
                 "lignes": [[t["indicateur"], f"{t['debut']} à {t['fin']}",
                             (F(f"V0 {t['indicateur']}", t["v0"]) if t["v0"] > 100000 else N(f"V0 {t['indicateur']}", t["v0"])),
                             (F(f"V1 {t['indicateur']}", t["v1"]) if t["v1"] > 100000 else N(f"V1 {t['indicateur']}", t["v1"])),
                             N(f"Mult {t['indicateur']}", t["multiplicateur"], 2), pc(f"TCAM {t['indicateur']}", t["tcam"])] for t in C["tcam"]]},
        disposition="dessous",
        analyse=["**Les prestations maladie et l'encours du Grenier croissent bien plus vite que les cotisations et les effectifs.** Les ressources ne suivent pas le rythme des emplois."],
        implication="Un écart durable entre la croissance des prestations et celle des cotisations conduit mécaniquement à un déficit technique : la projection devra en fixer l'horizon.",
        source="Source : balances des comptes du FPPN de 2020 au 31/08/2026 ; fichiers de précomptes et exports du Grenier reçus le 15/10/2026 (noms exacts en annexe « Fichiers sources ») ; traitements Adenka")
nat = R["prestations_maladie_nature"]
contenu({"titre": "Prestations maladie par nature", "sous_titre": "Plus de la moitié des prestations maladie est enregistrée sans détail de nature",
         "synthese": f"En 2025, le compte global porte {F('Compte global 2025', nat[5]['2025'])} sur {F('Prest mal 2025 (n)', s[2025]['prest_maladie'])}"},
        "Lecture : chaque couleur est une nature de soins. La catégorie « autres prestations maladie » correspond aux comptes 6012x, dont le compte global 60120000. Imagerie, analyses et consultations sont les natures détaillées qui progressent le plus.",
        graphique={"titre": fig("Prestations maladie comptabilisées par nature"), "type": "barres_empilees", "categories": cats_ex,
                   "series": [{"nom": n["nature"], "valeurs": [round(n[str(e)]) for e in range(2020, 2027)]} for n in nat],
                   "format": FMT_M, "etiquettes": False, "legende": "droite", "unite": "En millions de FCFA"},
        analyse=[f"**Les prestations maladie passent de {F('Prest 2021', s[2021]['prest_maladie'])} en 2021 à {F('Prest 2025 (nat)', s[2025]['prest_maladie'])} en 2025.**",
                 f"**Imagerie, analyses et consultations représentent {pc('Part img+ana+cons', (nat[4]['2025'] + nat[3]['2025'] + nat[1]['2025']) / s[2025]['prest_maladie'])} des prestations de 2025.** Ce sont les natures détaillées qui progressent le plus.",
                 "**Le compte global empêche le suivi par type de soins et par prestataire.**"],
        implication="Sans détail par nature et par prestataire, le FPPN ne peut ni négocier ses tarifs, ni repérer les abus, ni calculer le taux de prestations indues exigé par la Décision n° 05/2022.",
        source=SRC_COMPTA)
contenu({"titre": "Frais de gestion et résultat", "sous_titre": "Les excédents se réduisent et deviennent un déficit en 2026",
         "synthese": f"Résultat net : {F('Résultat 2020', s[2020]['resultat'])} en 2020, {F('Résultat 2025 (fr)', s[2025]['resultat'])} en 2025, {F('Résultat 2026', s[2026]['resultat'])} sur huit mois de 2026"},
        "Frais de gestion = comptes 61 à 66, hors prime d'assurance décès. Résultat = produits - charges + résultat hors activités ordinaires (classe 8). Les huit mois de 2026 n'intègrent pas les écritures de fin d'année (amortissements, provisions).",
        graphique={"titre": fig("Frais de gestion et résultat net"), "type": "barres", "categories": cats_ex,
                   "series": [{"nom": "Frais de gestion", "valeurs": [round(s[a]["frais_gestion"]) for a in range(2020, 2027)], "couleur": GRIS_C},
                              {"nom": "Résultat net", "valeurs": [round(s[a]["resultat"]) for a in range(2020, 2027)], "couleur": BLEU}],
                   "format": FMT_M, "unite": "En millions de FCFA"},
        analyse=[f"**Les frais de gestion pèsent entre {pc('Min frais', min(s[a]['taux_frais'] for a in range(2021, 2026)))} et {pc('Max frais', max(s[a]['taux_frais'] for a in range(2021, 2026)))} des cotisations depuis 2021.** La Décision n° 05/2022 limite les charges de fonctionnement à 20 % des cotisations.",
                 f"**Le résultat baisse depuis 2021.** Il passe de {F('Résultat 2021', s[2021]['resultat'])} à {F('Résultat 2025 (t)', s[2025]['resultat'])} en 2025.",
                 "**Les huit premiers mois de 2026 sont déficitaires**, avant même les écritures de fin d'année."],
        implication="Les frais pèsent presque deux fois la norme réglementaire. Leur réduction est le levier le plus direct pour restaurer l'excédent pendant que la sinistralité monte.",
        source=SRC_COMPTA)
contenu({"titre": "Choix de la formule : frais de gestion", "sous_titre": "La formule du texte et celle du rapport ne retiennent pas les mêmes comptes",
         "synthese": f"2025 : {pc('Frais rapport 2025', s[2025]['taux_frais'])} selon le rapport, {pc('Frais texte 2025', RT['CHF']['valeurs'][2025])} selon le texte, pour une norme d'au plus 20 %"},
        "Formule du rapport : frais de gestion (comptes 61 à 66, hors prime d'assurance décès 6518) / cotisations maladie et décès (7010 + 7021). Formule du texte (Décision n° 05/2022, tableau 1) : charges de fonctionnement (612, 613, 618, 62 à 68) / cotisations acquises (70 + 7906 - 6906). La méthode du FPPN n'a pas été communiquée.",
        graphique={"titre": fig("Taux de frais selon la formule"), "type": "barres", "categories": cats_ex,
                   "series": [{"nom": "Formule du rapport", "valeurs": [round(100 * s[a]["taux_frais"], 2) for a in range(2020, 2027)], "couleur": GRIS_C},
                              {"nom": "Formule du texte", "valeurs": [round(100 * RT["CHF"]["valeurs"][a], 2) for a in range(2020, 2027)], "couleur": BLEU}],
                   "format": "0.00", "unite": "En % des cotisations"},
        analyse=["**Les deux formules diffèrent sur trois points.** Le texte exclut les comptes 611 et 614 à 617, et inclut les comptes 67 et 68. Il divise par toutes les cotisations acquises, MEDOC compris. Le rapport retenait les comptes 61 à 66 et les seules cotisations maladie et décès.",
                 "**Seule la formule du texte est comparée à la norme de 20 %.** La formule du rapport reste utile pour suivre les frais sur la même base que la sinistralité par garantie."],
        implication="Quelle que soit la formule, les frais dépassent la norme de 20 % des cotisations : la conclusion de la partie 5 ne change pas.",
        source="Source : " + DEC + ", tableau 1 ; " + SRC_COMPTA[9:])

retenir("sinistralité", ("La sinistralité maladie approche le niveau des cotisations", "Ratio S/C maladie multiplié par plus de douze depuis 2021, résultat en baisse, déficit sur 2026",
                         "Les prestations maladie absorbent une part croissante des cotisations, et le résultat du FPPN baisse depuis 2021."),
        f"La hausse du S/C maladie est statistiquement significative, d'environ {N('Pente S/C (ret)', C['tendances'][1]['pente'], 2)} points par an de 2020 à 2025. Les frais dépassent la norme de 20 % des cotisations.",
        "Sans maîtrise des prestations et des frais, le FPPN deviendra déficitaire et ne pourra plus doter sa réserve obligatoire.",
        "Le FPPN doit détailler le compte global de prestations par nature et par prestataire, et transmettre la base des prestations.",
        "La hausse rapide du ratio maladie est le principal risque technique : à ce rythme, les cotisations maladie ne couvriront plus les prestations.")

# ---------------------------------------------------------------- PARTIE 6 : ÉPARGNE
separateur(6, "Épargne des mutualistes", "Dette d'épargne, actifs réservés, sensibilité aux retraits")
methode_gab({"titre": "Méthode : adossement de l'épargne", "sous_titre": "La couverture compare l'épargne due aux actifs qui lui sont réservés",
             "synthese": "Couverture = (banques, dépôts à terme et fonds réservés) / (comptes PEP et EFC)"},
            [("Passif d'épargne", "L'épargne due correspond aux soldes créditeurs des comptes 41195100 (PEP) et 41195200 (EFC). C'est une dette du FPPN envers les mutualistes."),
             ("Actifs réservés", "Adenka retient les comptes bancaires et dépôts à terme dont l'intitulé désigne l'EFC ou le PEP (5211020x, 5221020x, 5211030x, 5221030x, 526120) et le fonds commun de placement (5330)."),
             ("Couverture", "La couverture rapporte les actifs réservés à l'épargne due, à chaque clôture de 2020 au 31/08/2026."),
             ("Limite", "Le manuel des produits d'épargne ne fixe pas d'échéancier individuel. Le FCP est compté comme actif réservé, aucun document ne l'affectant à un autre emploi.")],
            "L'épargne PEP et EFC appartient aux mutualistes : c'est une dette du FPPN. Un bon adossement suppose des actifs liquides et réservés d'un montant au moins égal à la part exigible.")
ep = EX["epargne_2026"]
contenu({"titre": "Calcul pas à pas : couverture de l'épargne", "sous_titre": "Deux comptes de dette face à quatre comptes d'actifs réservés",
         "synthese": f"{F('Actifs dédiés (pas)', b[2026]['actifs_dedies_epargne'])} / {F('Épargne (pas)', b[2026]['epargne_totale'])} = {pc('Couverture (pas)', b[2026]['couverture_epargne'])} au 31/08/2026"},
        "Montants lus dans la balance au 31/08/2026 (provisoire). Les soldes créditeurs du passif sont affichés en valeur positive. Le fonds commun de placement est compté comme actif réservé bien que sa destination ne soit pas précisée : la couverture réelle est donc au plus de ce niveau.",
        pas_a_pas={"entetes": ["Compte", "Intitulé", "Solde au 31/08/2026"],
                   "lignes": [[c[0], c[1], nb(-float(c[2])) + " FCFA"] for c in ep["passif"]] + [["", "Total épargne due", F("Épargne totale (pas)", b[2026]["epargne_totale"])]] + [[c[0], c[1], nb(float(c[2])) + " FCFA"] for c in ep["actifs"]] + [["", "Total actifs réservés", F("Total actifs (pas)", b[2026]["actifs_dedies_epargne"])]],
                   "largeurs": [3.4, 11.4, 6.6], "rouge": [(2, 2), (7, 2)],
                   "explications": ["Les deux comptes 41195 enregistrent la dette envers les épargnants PEP et EFC.", "Les banques, le dépôt à terme et le fonds commun de placement forment les actifs réservés.", f"La couverture vaut donc {pc('Couverture (expl)', b[2026]['couverture_epargne'])} au 31/08/2026."],
                   "fichier": "AUDIT ADENKA\\DAFC\\2026\\BALANCE FPPN DU 01-01-2026 AU 31-08-2026.xlsx"},
        analyse=["**Ce ratio est aussi le ratio de solvabilité des risques longs de la Décision n° 05/2022** (actifs admis / engagements réglementés), dont la norme est d'au moins 1."],
        implication="Avec une couverture inférieure à un cinquième, le FPPN ne respecte pas cette norme : la partie 10 la reprend dans le tableau de bord sans refaire le calcul.",
        source="Source : balance au 31/08/2026 (impression provisoire)")
contenu({"titre": "Adossement de l'épargne", "sous_titre": "L'épargne des mutualistes n'est couverte qu'à 16,28 % par des actifs réservés",
         "synthese": f"Couverture : entre {pc('Min couverture', min(b[a]['couverture_epargne'] for a in range(2020, 2027)))} et {pc('Max couverture', max(b[a]['couverture_epargne'] for a in range(2020, 2027)))} selon les années"},
        "L'épargne non couverte finance d'autres emplois : investissements au CMP, terrains, créance sur l'Hôpital de la Police, prêts du Grenier. Ces emplois ne sont pas mobilisables rapidement.",
        graphique={"titre": fig("Épargne au passif et actifs réservés"), "type": "barres", "categories": yrs_bil,
                   "series": [{"nom": "Épargne PEP et EFC", "valeurs": [round(b[a]["epargne_totale"]) for a in range(2020, 2027)], "couleur": BLEU},
                              {"nom": "Actifs réservés", "valeurs": [round(b[a]["actifs_dedies_epargne"]) for a in range(2020, 2027)], "couleur": ORANGE}],
                   "format": FMT_M, "unite": "En millions de FCFA"},
        analyse=[f"**La dette d'épargne augmente de {pc('Hausse épargne', b[2026]['epargne_totale'] / b[2020]['epargne_totale'] - 1)} depuis 2020.** Les actifs réservés ne suivent pas cette hausse.",
                 f"**L'épargne EFC triple.** Elle passe de {F('EFC 2020', b[2020]['epargne_efc'])} en 2020 à {F('EFC 2026', b[2026]['epargne_efc'])} au 31/08/2026.",
                 "**L'épargne non couverte finance des emplois peu liquides** : investissements au CMP, terrains, créance sur l'Hôpital de la Police et prêts du Grenier."],
        implication="Le règlement d'exécution n° 004/2022 impose le cantonnement des patrimoines quand une mutuelle couvre à la fois des risques courts et longs (art. 15). L'épargne n'est séparée qu'en partie des autres ressources du FPPN.",
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
        analyse=[f"**Les épargnants PEP passent de {N('PEP 2020', pv[0]['epargnants'])} en 2020 à {N('PEP 2026', pv[-1]['epargnants'])} en 2026.** Le produit attire un nombre croissant de policiers.",
                 f"**En 2025, le PEP reçoit {F('PEP 2025 (t)', pv[5]['versements'])} mais le compte n'augmente que de {F('Var PEP 2025', b[2025]['epargne_pep'] - b[2024]['epargne_pep'])}.** La différence correspond à des sorties d'épargne.",
                 f"**Les sorties d'épargne atteignent {F('Sorties épargne 2025', AG[2025]['sorties_epargne'])} en 2025** (débits des comptes 41195100 et 41195200)."],
        implication="Ces sorties sont les « engagements » d'un trimestre que la réserve de trésorerie des risques longs doit couvrir (art. 19 du règlement d'exécution) : leur suivi mensuel est nécessaire pour fixer le niveau de cette réserve.",
        source=SRC_PREC + " ; " + SRC_COMPTA[9:])
rr = C["sensibilite_retraits"]
contenu({"titre": "Sensibilité : retraits d'épargne", "sous_titre": "Un retrait de plus d'un cinquième de l'épargne épuiserait toute la trésorerie du FPPN",
         "synthese": f"Les actifs réservés couvrent un retrait de {pc('Retrait couvert actifs', C['retrait_max_couvert']['par_actifs_dedies'])} de l'épargne ; la trésorerie totale, {pc('Retrait couvert tréso', C['retrait_max_couvert']['par_tresorerie'])}"},
        "Scénarios théoriques, pas des prévisions : on suppose qu'une part de l'épargne est demandée en une fois et on la compare aux actifs réservés puis à toute la trésorerie (banques et caisses). Les conditions réelles de retrait (délais, pénalités, échéances) ne sont pas connues.",
        graphique={"titre": fig("Besoin de liquidités selon le taux de retrait"), "type": "barres", "categories": [nb(100 * r["taux_retrait"], 2) + " %" for r in rr],
                   "series": [{"nom": "Montant à rembourser", "valeurs": [round(r["besoin"]) for r in rr], "couleur": BLEU},
                              {"nom": "Actifs réservés", "valeurs": [round(r["actifs_dedies"]) for r in rr], "couleur": ORANGE},
                              {"nom": "Trésorerie totale", "valeurs": [round(r["tresorerie_totale"]) for r in rr], "couleur": GRIS_C}],
                   "format": FMT_M, "unite": "En millions de FCFA, au 31/08/2026"},
        analyse=[f"**Les actifs réservés couvrent un retrait de 10 % de l'épargne ({F('Retrait 10%', rr[1]['besoin'])}).**",
                 f"**Au-delà de {pc('Seuil tréso', C['retrait_max_couvert']['par_tresorerie'])} de retraits, toute la trésorerie du FPPN ne suffit plus.**",
                 "**Le FPPN doit alors céder ou mobiliser des actifs peu liquides**, comme les terrains ou les avances au CMP."],
        implication="Le risque de liquidité dépend des conditions de retrait du PEP et de l'EFC : le règlement de ces produits et les âges des épargnants permettront de passer de scénarios théoriques à un échéancier réel.",
        source=SRC_COMPTA)
contenu({"titre": "Poste 998 : autres références", "sous_titre": "Le poste de l'épargne PEP porte aussi d'autres références à expliquer",
         "synthese": "PTI, RPEP, EP et TI apparaissent à côté des références PEP sur la période 2020 à 2026"},
        "La référence est le code qui accompagne chaque précompte (exemple : PEP0522 = PEP souscrit en mai 2022). Les références PTI ont des montants mensuels importants et peu de mutualistes : elles pourraient correspondre à des remboursements de prêts, ce que le FPPN doit confirmer.",
        tableau={"titre": tab("Références du poste 998, 2020 à août 2026"), "entetes": ["Référence", "Lignes", "Montant total", "Mutualistes"], "largeurs": [4.0, 4.0, 7.4, 4.0],
                 "lignes": [[r["produit"], N(f"Lignes {r['produit']}", r["lignes"]), F(f"Montant {r['produit']}", r["montant"]), N(f"Mut {r['produit']}", r["mutualistes"])] for r in R["poste_998_references"]]},
        disposition="droite",
        analyse=["**Le poste 998 ne porte pas que de l'épargne PEP.** D'autres références, dont PTI, apparaissent avec des montants importants et peu de mutualistes.",
                 "**Leur nature n'est documentée nulle part.** Il peut s'agir de remboursements de prêts PEP (comptes 27110000 et 27120000), ce que le FPPN doit confirmer."],
        implication="Tant que ces références ne sont pas expliquées, la dette d'épargne PEP ne peut pas être reconstituée à partir des précomptes.",
        source=SRC_PREC)
retenir("épargne", ("L'épargne due aux mutualistes est très peu couverte par des actifs disponibles", "Couverture de 16,28 % au 31/08/2026 ; un retrait de 21,34 % épuiserait toute la trésorerie",
                    "La dette d'épargne augmente plus vite que les actifs réservés, et l'écart finance des emplois peu liquides."),
        F("Épargne (ret)", b[2026]["epargne_totale"]) + " d'épargne sont dus au 31/08/2026, pour " + F("Actifs réservés (ret)", b[2026]["actifs_dedies_epargne"]) + " d'actifs réservés.",
        "Le FPPN ne peut pas rembourser rapidement une vague de retraits, et il ne respecte pas le ratio de solvabilité des risques longs.",
        "Le FPPN doit transmettre le règlement des produits, les comptes individuels et les conditions de retrait, puis cantonner l'épargne.",
        "L'adossement de l'épargne est le principal risque financier relevé.")

# ---------------------------------------------------------------- PARTIE 7 : GRENIER
# ---- v9 : règles et taux de l'épargne, intérêts servis, épargne présentée en Assemblée générale (partie 6)
SRC_EP = "Source : CAHIER DE CHARGEEFC & PEP 2021.pdf ; MANUEL DE PRODUIT D'EPARGNE FPPN VF VF.docx ; CARTOGRAPHIE DES RISQUES COUVERTS ET DES PRESTATIONS OFFERTES.docx (reçus le 22/10/2026) ; entretien du chef comptable du 19/10/2026"
contenu({"titre": "EFC et PEP : trois taux dans trois documents", "sous_titre": "Le taux servi sur l'EFC vaut 3,5 %, 5,5 % ou 14 % selon le document",
         "synthese": "Aucun document ne fixe un taux unique, signé et daté : l'engagement d'épargne ne peut pas être évalué sans ce taux"},
        "Le taux servi est l'intérêt que le FPPN doit ajouter à l'épargne de chaque mutualiste. Il fixe la valeur de l'engagement : à épargne égale, un taux plus élevé fait payer davantage au départ du mutualiste.",
        tableau={"titre": tab("Taux servis sur l'épargne selon les documents du FPPN"), "entetes": ["Document (fichier exact)", "Produit", "Taux écrit", "Portée"], "largeurs": [10.4, 3.0, 7.2, 10.6], "taille": 8.5,
                 "lignes": [["CAHIER DE CHARGEEFC & PEP 2021.pdf, « B- Taux d'intérêt applicable »", "EFC", "14 % d'intérêt simple, sans période indiquée", "Cotisations au-delà de 2024 et souscriptions postérieures au cahier des charges ; bonus de 300 000 FCFA si à jour"],
                            ["MANUEL DE PRODUIT D'EPARGNE FPPN VF VF.docx", "EFC", "5,5 % l'an", "Rémunération de l'EFC ; certificat d'adoption non signé"],
                            ["MANUEL DE PRODUIT D'EPARGNE FPPN VF VF.docx", "PEP", "3,5 % tous les ans", "Rémunération du PEP"],
                            ["CARTOGRAPHIE DES RISQUES COUVERTS ET DES PRESTATIONS OFFERTES.docx", "EFC", "5,5 % à partir de 2024", "Taux d'intérêt de l'EFC"],
                            ["Entretien du chef comptable, 19/10/2026", "Épargne", "3,5 %", "Taux servi déclaré, sans document"]]},
        disposition="dessous",
        analyse=["**Pour l'EFC, trois taux coexistent : 3,5 %, 5,5 % l'an et 14 % d'intérêt simple.** Le cahier des charges n'indique pas si les 14 % s'appliquent une fois ou chaque année.",
                 "**Aucun des trois documents n'est accompagné d'une décision d'adoption signée.** Le manuel porte un certificat d'adoption non signé."],
        implication="Le taux applicable doit être arrêté par une décision des organes du FPPN, puis appliqué de manière identique aux comptes, aux relevés individuels et au calcul de l'engagement (PBC, item 99).",
        valable="documents du FPPN reçus le 22/10/2026 et entretien du 19/10/2026 ; le taux réellement appliqué aux mutualistes n'est pas connu",
        source=SRC_EP)
_gl = V9["gl2024_epargne"]
_nat = {(x["compte"], x["nature"]): x for x in _gl}
_lig_gl = [[("PEP (41195100)" if x["compte"] == "41195100" else "EFC (41195200)"), x["nature"], N(f"GL écr déb {x['compte']} {x['nature']}", x["ecritures_debit"]),
            F(f"GL déb {x['compte']} {x['nature']}", x["debit"]), N(f"GL écr cré {x['compte']} {x['nature']}", x["ecritures_credit"]), F(f"GL cré {x['compte']} {x['nature']}", x["credit"])] for x in _gl]
contenu({"titre": "Intérêts servis : aucune écriture dans les comptes", "sous_titre": "Les comptes d'épargne ne reçoivent que des cotisations et ne versent que des remboursements",
         "synthese": "En 2024, aucune écriture d'intérêts n'est passée sur les comptes PEP et EFC, et aucune balance de 2020 à 2026 ne porte de charge d'intérêts servis aux épargnants"},
        "Le grand livre 2024 détaille chaque écriture des deux comptes d'épargne. Les écritures sont classées par nature d'après leur libellé : précomptes et cotisations, règlements EFC, remboursements PEP, rachats, autres. Une écriture d'intérêts servis porterait un libellé « intérêts » ou une contrepartie en charge (classe 67).",
        tableau={"titre": tab("Écritures 2024 des comptes d'épargne par nature"), "entetes": ["Compte", "Nature", "Écritures au débit", "Débit", "Écritures au crédit", "Crédit"], "largeurs": [4.4, 6.0, 3.6, 6.0, 3.6, 6.0], "taille": 9,
                 "lignes": _lig_gl},
        disposition="dessous",
        analyse=["**Aucune écriture d'intérêts ne figure sur les comptes d'épargne en 2024, et aucun compte de charge d'intérêts servis n'existe dans les balances de 2020 au 31/08/2026.** En 2024, le solde comptable de l'épargne n'évolue donc que par les cotisations, les remboursements et les rachats.",
                 f"**En 2024, {N('Règlements EFC 2024', _nat[('41195200', 'Règlements EFC')]['ecritures_debit'])} règlements EFC ont été payés pour {F('Montant règlements EFC 2024', _nat[('41195200', 'Règlements EFC')]['debit'])}.** Le grand livre ne dit pas si ces règlements comprennent des intérêts calculés hors comptabilité."],
        implication="Si un taux est dû aux mutualistes, la dette d'épargne inscrite au bilan est inférieure à l'engagement réel du montant des intérêts acquis et non comptabilisés. Ce montant ne peut être chiffré qu'avec le taux applicable et l'historique individuel des versements.",
        valable="grand livre 2024 pour le détail des écritures ; balances de 2020 au 31/08/2026 pour l'absence de charge d'intérêts",
        source="Source : GRAND LIVRE FPPN 2024.xlsx (comptes 41195100 et 41195200) ; " + SRC_COMPTA[9:])
_ef = V9["efc_taux"]
contenu({"titre": "Sensibilité : intérêts annuels dus sur l'EFC", "sous_titre": "Au taux de 5,5 %, les intérêts dus chaque année dépassent les produits financiers du FPPN",
         "synthese": f"Encours EFC au 31/12/2025 : {F('EFC encours 2025 (t)', next(x['encours_efc'] for x in _ef if x['exercice'] == 2025))} ; intérêts d'une année à 5,5 % : {F('EFC int 5,5 2025', next(x['interets_5_5'] for x in _ef if x['exercice'] == 2025))}"},
        "Le calcul applique chaque taux écrit à l'encours EFC de fin d'année (compte 41195200), une hypothèse à la fois. La colonne à 14 % suppose un taux annuel : c'est une borne haute, car le cahier des charges ne précise pas la période. Les produits financiers sont ceux de la partie 9.",
        tableau={"titre": tab("Intérêts d'une année sur l'encours EFC selon le taux servi"), "entetes": ["Exercice", "Encours EFC (41195200)", "Intérêts à 3,5 %", "Intérêts à 5,5 %", "Intérêts à 14 % (borne haute)", "Produits financiers"],
                 "largeurs": [3.0, 6.0, 5.2, 5.2, 6.0, 5.6], "taille": 9,
                 "lignes": [[str(x["exercice"]) if x["exercice"] < 2026 else "31/08/2026", F(f"EFC enc {x['exercice']}", x["encours_efc"]), F(f"EFC i3,5 {x['exercice']}", x["interets_3_5"]),
                             F(f"EFC i5,5 {x['exercice']}", x["interets_5_5"]), F(f"EFC i14 {x['exercice']}", x["interets_14"]),
                             F(f"Prod fin (EFC) {x['exercice']}", x["produits_financiers"]) if x["produits_financiers"] is not None else "-"] for x in _ef]},
        disposition="dessous",
        analyse=[f"**Même au taux le plus bas (3,5 %), les intérêts d'une année sur l'EFC ({F('EFC i3,5 2025 (t)', next(x['interets_3_5'] for x in _ef if x['exercice'] == 2025))} en 2025) dépassent, chaque année de 2021 à 2025, les produits financiers de l'année.**",
                 "**L'épargne n'est couverte qu'en faible part par des actifs qui lui sont réservés (partie 6, adossement de l'épargne).** Les produits de ces actifs ne suffisent pas à servir un taux, quel qu'il soit, sans recourir aux cotisations des garanties."],
        implication="Le taux promis aux mutualistes doit être cohérent avec le rendement des actifs qui couvrent l'épargne. À défaut, chaque point de taux servi crée une perte récurrente que les excédents des garanties maladie et décès devront combler.",
        valable="encours EFC comptable au 31/12 de 2020 à 2025 et au 31/08/2026, une année d'intérêts, sans projection ni capitalisation",
        source=SRC_EP + " ; " + SRC_COMPTA[9:])
_e25 = next(x for x in V9["epargne_mouvements"] if x["exercice"] == 2025)
_e24 = next(x for x in V9["epargne_mouvements"] if x["exercice"] == 2024)
# ---- v10 : intérêts dus sur l'épargne depuis le 01/01/2020 (partie 6)
_EI = [e for e in V10["epargne"]]


def _ip(e, t):
    return (e["pep_debut"] + e["pep_fin"]) / 2 * t * e["mois"] / 12


def _ie(e, t):
    return (e["efc_debut"] + e["efc_fin"]) / 2 * t * e["mois"] / 12


contenu({"titre": "Épargne : intérêts dus non comptabilisés", "sous_titre": f"Du 01/01/2020 au 31/08/2026, les intérêts dus selon le manuel d'épargne atteignent {F('EI total', INT55)}",
         "synthese": f"Engagement d'épargne au 31/08/2026 : {F('EI dû', EP_DUE26)} au bilan, {F('EI engagement', EP_DUE26 + INT55)} avec les intérêts dus (PEP 3,5 %, EFC 5,5 %)"},
        "Intérêts de l'année = épargne moyenne de l'année (moyenne du début et de la fin) × taux servi × nombre de mois / 12. Taux du manuel des produits d'épargne : PEP 3,5 %, EFC 5,5 %. Les intérêts ne sont pas capitalisés d'une année sur l'autre : le montant est donc un minimum. Pour 2020, le premier solde disponible est celui du 01/09/2020 (à-nouveaux de la balance 2020) : la moyenne du 01/09 et du 31/12/2020 est appliquée à toute l'année.",
        tableau={"titre": tab("Intérêts dus sur l'épargne, par exercice"), "entetes": ["Exercice", "PEP moyen", "EFC moyenne", "Intérêts PEP (3,5 %)", "Intérêts EFC (5,5 %)", "Total de l'année", "Total si EFC à 3,5 %"],
                 "largeurs": [3.0, 4.8, 4.8, 4.4, 4.4, 4.6, 4.6], "taille": 8.5,
                 "lignes": [[str(e["exercice"]) if e["exercice"] < 2026 else "2026 (8 mois)", F(f"EI pep moy {e['exercice']}", (e["pep_debut"] + e["pep_fin"]) / 2), F(f"EI efc moy {e['exercice']}", (e["efc_debut"] + e["efc_fin"]) / 2),
                             F(f"EI ip {e['exercice']}", _ip(e, 0.035)), F(f"EI ie {e['exercice']}", _ie(e, 0.055)), F(f"EI tot {e['exercice']}", _ip(e, 0.035) + _ie(e, 0.055)),
                             F(f"EI tot35 {e['exercice']}", _ip(e, 0.035) + _ie(e, 0.035))] for e in _EI]
                 + [["Total", "", "", F("EI ip tot", sum(_ip(e, 0.035) for e in _EI)), F("EI ie tot", sum(_ie(e, 0.055) for e in _EI)), F("EI tot (t)", INT55), F("EI tot35 (t)", INT35)]]},
        disposition="dessous",
        analyse=[f"**Selon le taux retenu pour l'EFC, les intérêts dus depuis le 01/01/2020 sont compris entre {F('EI bas', INT35)} et {F('EI haut', INT55)}.** Aucun de ces montants n'est comptabilisé : il n'existe ni charge d'intérêts ni dette d'intérêts dans les balances.",
                 f"**Au 31/12/2025, les intérêts dus atteignent déjà {F('EI 2025', INT55_25)}**, soit {N('EI fois pf', INT55_25 / sum(x['produits_financiers'] for x in R['rendement_placements']), 1)} fois les produits financiers du FPPN de 2021 à 2025 ({F('EI pf', sum(x['produits_financiers'] for x in R['rendement_placements']))})."],
        implication="Les intérêts dus sont une dette envers les mutualistes. Le taux servi doit être arrêté par une décision signée, puis les intérêts calculés compte par compte et comptabilisés chaque année.",
        valable="soldes des comptes 41195100 (PEP) et 41195200 (EFC) du 01/09/2020 au 31/08/2026 (balances 2020 à 2025 et provisoire 2026), 2020 approché par les soldes du 01/09 et du 31/12 ; taux du manuel d'épargne ; intérêts antérieurs au 01/01/2020 et capitalisation non comptés",
        source=SRC_EP + " ; " + SRC_COMPTA[9:] + " ; onglet 21 du classeur des calculs")
contenu({"titre": "Rapport de gestion 2025 : l'épargne présentée est la collecte de l'année", "sous_titre": "Le rapport annonce 4,39 milliards de FCFA d'épargne, alors que la dette d'épargne atteint 13,81 milliards",
         "synthese": "Les montants du rapport de gestion 2025 sont les crédits de l'année des comptes 41195100 et 41195200, présentés comme « le montant de ces épargnes au 31/12/2025 »"},
        "Le rapport de gestion détaillé de l'exercice 2025 (page 98, tableau 2 (III)) présente l'épargne PEP et EFC. Ses montants sont comparés aux mouvements et aux soldes des mêmes comptes dans la balance 2025.",
        tableau={"titre": tab("Épargne au 31/12/2025 : rapport de gestion et balance"), "entetes": ["Compte", "Rapport de gestion 2025", "Crédits de l'année (balance)", "Solde au 31/12/2025 (balance)", "Solde au 31/12/2024 (balance)"],
                 "largeurs": [4.6, 6.4, 6.4, 6.8, 6.8], "taille": 9,
                 "lignes": [["PEP (41195100)", F("RG PEP 2025", 2038935495), F("Crédits PEP 2025", _e25["mvt_c_41195100"]), F("Solde PEP 2025 (RG)", _e25["sol_c_41195100"]), F("Solde PEP 2024 (RG)", _e24["sol_c_41195100"])],
                            ["EFC (41195200)", F("RG EFC 2025", 2348888140), F("Crédits EFC 2025", _e25["mvt_c_41195200"]), F("Solde EFC 2025 (RG)", _e25["sol_c_41195200"]), F("Solde EFC 2024 (RG)", _e24["sol_c_41195200"])],
                            ["Total", F("RG total 2025", 4387823635), F("Crédits total 2025", _e25["mvt_c_41195100"] + _e25["mvt_c_41195200"]), F("Solde total 2025 (RG)", _e25["sol_c_41195100"] + _e25["sol_c_41195200"]), F("Solde total 2024 (RG)", _e24["sol_c_41195100"] + _e24["sol_c_41195200"])]]},
        disposition="dessous",
        analyse=["**Les montants du rapport de gestion sont exactement les crédits de l'année dans la balance**, c'est-à-dire les cotisations d'épargne collectées en 2025, et non l'épargne due aux mutualistes.",
                 "**La dette d'épargne réelle au 31/12/2025 est plus de trois fois supérieure au montant présenté.** Le rapport compare aussi la collecte 2025 à la collecte 2024 en la qualifiant de « variation » de l'épargne."],
        implication="L'information donnée à l'Assemblée générale minore l'engagement d'épargne du FPPN. Le rapport de gestion doit présenter le solde dû aux mutualistes, distinct de la collecte de l'année (PBC, item 112).",
        valable="exercice 2025 ; les rapports de gestion des autres années n'ont pas été contrôlés sur ce point",
        source="Source : Rapport de Gestion  2025 Final (2).pdf, pages 97 et 98, tableau 2 (III) ; " + SRC_COMPTA[9:])
separateur(7, "Prêts du Grenier", "Encours, production, retards, précomptes")
methode_gab({"titre": "Méthode : risque de crédit", "sous_titre": "Le risque des prêts se mesure par la part de l'encours en retard",
             "synthese": "Encours en retard de plus de 30 jours / encours total, à chaque date de balance âgée"},
            [("Part en retard (montant)", "Adenka rapporte l'encours des prêts en retard de plus de 30 jours à l'encours total, sur les 37 balances âgées reçues."),
             ("Part en retard (nombre)", "Adenka rapporte le nombre de prêts en retard au nombre de prêts au 31/08/2026, avec un intervalle de confiance de Wilson à 95 %."),
             ("Précompte des échéances", "Adenka rapporte les précomptes 9X4 du mois aux échéances des prêts en cours, pour les mois de 2026."),
             ("Limite", "Le retard est celui que calcule le logiciel de prêts. Il dépend de la date à laquelle la Solde verse les précomptes.")],
            "L'encours est le capital restant dû. Le retard (en jours) est celui imprimé par le logiciel de prêts. On mesure la part en nombre de prêts et en montant d'encours.")
ge_fa = [g for g in GE if g["arrete"].startswith("31/12") or g["arrete"] == "31/08/2026"]
contenu({"titre": "Encours du Grenier", "sous_titre": "L'encours de prêts a été multiplié par 132 depuis fin 2021",
         "synthese": f"Au 31/08/2026 : {N('Prêts', GP['prets'])} prêts, {F('Encours (gr)', GP['encours'])}, dont {F('Financement FPPN', b[2026]['grenier_financement'])} financés par le FPPN"},
        f"Taux de croissance annuel moyen de l'encours de fin 2022 à fin 2025 : {pc('TCAM encours', TC['Encours des prêts du Grenier']['tcam'])}. La créance du FPPN sur le Grenier (compte 47111102) mesure le financement apporté.",
        graphique={"titre": fig("Encours des prêts et financement du FPPN"), "type": "barres", "categories": [g["arrete"] for g in ge_fa],
                   "series": [{"nom": "Encours", "valeurs": [round(g["encours"]) for g in ge_fa], "couleur": BLEU},
                              {"nom": "Financement FPPN", "valeurs": [round(g["financement_fppn_47111102"]) if g["financement_fppn_47111102"] is not None else None for g in ge_fa], "couleur": ORANGE}],
                   "format": FMT_M, "unite": "En millions de FCFA"},
        analyse=[f"**L'encours passe de {F('Encours 2021', ge_fa[0]['encours'])} fin 2021 à {F('Encours 2026 (t)', GP['encours'])} au 31/08/2026**, soit {N('Mult encours', GP['encours'] / ge_fa[0]['encours'], 2)} fois plus.",
                 f"**Le FPPN finance {pc('Part financement', b[2026]['grenier_financement'] / GP['encours'])} de l'encours.** Le Grenier repose donc presque entièrement sur les ressources du FPPN.",
                 f"**Les prêts sont courts et chers.** La durée médiane est de {N('Durée médiane', GP['duree_mediane'])} mois et le taux médian de {pcv('Taux médian', GP['taux_median'])}."],
        implication="Si le Grenier relève des activités génératrices de revenus, le règlement d'exécution n° 004/2022 les autorise à condition qu'elles ne mettent pas en péril l'activité d'assurance et que leurs pertes ne soient pas couvertes par les cotisations ni par les réserves (art. 26) : le financement du Grenier doit être encadré par une décision formelle.",
        source=SRC_GRENIER + " ; " + SRC_COMPTA[9:])
gpr = R["grenier_production"]
contenu({"titre": "Production de prêts", "sous_titre": "Les crédits scolaires et habitat portent la croissance de la production",
         "synthese": f"Montants débloqués : {F('Prod 2022', gpr['total'][1])} en 2022, {F('Prod 2025', gpr['total'][4])} en 2025 ({N('Nb prêts 2025', gpr['nombre'][4])} prêts)"},
        "Production = montants débloqués sur la période (fichiers annuels de 2021 à 2025, mensuels par sous-portefeuille en 2026). La politique de crédit de 2021 est au format Word ancien et doit être renvoyée lisible pour vérifier taux, durées et plafonds.",
        graphique={"titre": fig("Montants débloqués par type de crédit"), "type": "barres_empilees", "categories": [str(a) if a < 2026 else "2026 (8 mois)" for a in gpr["annees"]],
                   "series": [{"nom": t.capitalize(), "valeurs": [round(v) for v in gpr["valeurs"][t]]} for t in gpr["types"]],
                   "format": FMT_M, "etiquettes": False, "legende": "droite", "unite": "En millions de FCFA"},
        analyse=[f"**La production a été multipliée par {N('Mult prod', gpr['total'][4] / gpr['total'][1], 2)} entre 2022 et 2025.**",
                 f"**Les crédits scolaires et habitat font {pc('Part scol+hab', (gpr['valeurs']['CRÉDIT SCOLAIRE'][4] + gpr['valeurs']['CREDIT HABITAT'][4]) / gpr['total'][4])} des montants de 2025.**",
                 "**Le crédit habitat, aux montants unitaires plus importants, allonge la durée du risque.**"],
        implication="Une production qui croît plus vite que les ressources propres du Grenier augmente le besoin de financement du FPPN : la politique de crédit doit fixer un plafond d'encours rapporté aux ressources disponibles.",
        source=SRC_GRENIER)
ph = [p for p in R["grenier_par_historique"] if p["arrete"] != "30/09/2026"]
contenu({"titre": "Retards de remboursement", "sous_titre": "Les retards ont culminé début 2026 avant de revenir à un niveau faible",
         "synthese": f"Encours en retard de plus de 30 jours : {pc('PAR30 janv 2026', [p for p in ph if p['arrete'] == '31/01/2026'][0]['par30'])} fin janvier 2026, {pc('PAR30 août 2026', GP['par30'])} au 31/08/2026"},
        f"Au 31/08/2026, {N('Prêts retard 30', IC['retard30'])} prêts sur {N('Prêts (ic)', IC['prets'])} ont plus de 30 jours de retard, soit {pc('Part prêts retard', IC['p30'])} des prêts (IC de Wilson à 95 % : {pc('IC bas retard', IC['p30_bas'])} à {pc('IC haut retard', IC['p30_haut'])}). En montant, la part est plus faible ({pc('PAR30 (n)', GP['par30'])}) : les prêts en retard sont de petits encours.",
        graphique={"titre": fig("Part de l'encours en retard de plus de 30 et 90 jours"), "type": "lignes", "categories": [p["arrete"] for p in ph],
                   "series": [{"nom": "Plus de 30 jours", "valeurs": [round(100 * p["par30"], 2) for p in ph], "couleur": BLEU},
                              {"nom": "Plus de 90 jours", "valeurs": [round(100 * p["par90"], 2) for p in ph], "couleur": ORANGE}],
                   "format": "0.00", "etiquettes": False, "unite": "En % de l'encours"},
        analyse=["**Fin janvier 2026, la plus grande part de l'encours avait entre 31 et 90 jours de retard.**",
                 "**Ces retards se sont résorbés en quelques mois.** Ce profil fait penser à des échéances non précomptées en fin d'année 2025.",
                 f"**{N('Prêts > 180 j', IC['retard180'])} prêts ont plus de 180 jours de retard.** Ces prêts doivent être provisionnés."],
        implication="Les retards dépendent surtout du circuit de précompte, pas de la solvabilité des emprunteurs : le FPPN doit obtenir de la Solde un calendrier fiable et suivre chaque mois les échéances non précomptées.",
        source=SRC_GRENIER)
pr_ = EX["prets_retard"]
contenu({"titre": "Calcul pas à pas : prêts en retard", "sous_titre": "Les retards les plus anciens dépassent trois ans et révèlent des mécanos anormaux",
         "synthese": "Les quatre prêts les plus en retard au 31/08/2026 ont entre 1 105 et 1 308 jours de retard"},
        "Lecture : (1) le retard est imprimé en jours par le logiciel ; (2) trois de ces prêts ont un « mécano » de 2 ou 3 chiffres (142, 87, 166), alors qu'un mécano en compte 6 : ce sont des anomalies de saisie ; (3) ces prêts anciens ont encore un encours et doivent être provisionnés ou passés en perte. Noms masqués.",
        pas_a_pas={"entetes": ["Mécano", "Nom", "Montant débloqué", "Encours", "Retard (jours)"], "lignes": [[r[0], r[1], nb(float(r[2])) + " FCFA", nb(float(r[3])) + " FCFA", r[4]] for r in pr_["lignes"]],
                   "largeurs": [3.4, 4.0, 6.0, 5.6, 4.6], "rouge": [(0, 0), (0, 4), (1, 0)],
                   "explications": ["Ces prêts ont un mécano de 2 ou 3 chiffres, ce qui révèle une saisie anormale.", "Le plus ancien de ces prêts a plus de trois ans de retard.", "Ces prêts ont encore un encours inscrit, qui doit être provisionné."],
                   "fichier": "AUDIT ADENKA\\SERVICE GRENIER DU POLICIER\\FICHIER ALLER\\GRENIER AOÛT 2026\\BalanceAgeeDetaille AOUT 2026.xls"},
        analyse=["**Les prêts les plus anciens ne peuvent plus être précomptés, car leur mécano est faux.** Sans mécano valide, la Solde ne peut pas retenir l'échéance sur le salaire."],
        implication="Le Grenier doit corriger ces identifiants et décider, pour chaque prêt, d'une provision ou d'un passage en perte.",
        source=SRC_GRENIER + " ; noms masqués")
ech = [e for e in R["grenier_echeances"] if e["precompte_9x4"]]
contenu({"titre": "Échéances et précomptes", "sous_titre": "Les précomptes ne couvrent que 76 à 85 % des échéances des prêts en cours",
         "synthese": f"De janvier à août 2026, l'écart mensuel va de {F('Écart min', min(e['echeances_balance'] - e['precompte_9x4'] for e in ech))} à {F('Écart max', max(e['echeances_balance'] - e['precompte_9x4'] for e in ech))}"},
        "Échéances = somme des montants d'échéance des prêts ayant un encours dans la balance âgée du mois. Précomptes = précomptes 9X4 retour du même mois. L'écart peut être payé hors Solde (caisse, banque) ou correspondre à des rejets.",
        graphique={"titre": fig("Échéances des prêts en cours et précomptes 9X4, 2026"), "type": "barres", "categories": [e["mois"] for e in ech],
                   "series": [{"nom": "Échéances", "valeurs": [round(e["echeances_balance"]) for e in ech], "couleur": BLEU},
                              {"nom": "Précomptes 9X4", "valeurs": [round(e["precompte_9x4"]) for e in ech], "couleur": ORANGE}],
                   "format": FMT_M, "unite": "En millions de FCFA"},
        analyse=[f"**La part précomptée va de {pc('Min précompte/éch', min(e['precompte_9x4'] / e['echeances_balance'] for e in ech))} à {pc('Max précompte/éch', max(e['precompte_9x4'] / e['echeances_balance'] for e in ech))} selon le mois.**",
                 "**L'export du logiciel perd la lettre-clé du mécano** (321100C devient 321100), ce qui complique le rapprochement prêt par prêt.",
                 "**Un rapprochement avec les encaissements bancaires du Grenier est nécessaire** pour savoir si l'écart est payé autrement."],
        implication="Si l'écart n'est pas encaissé par un autre canal, il alimente chaque mois de nouveaux retards : c'est la question à poser en priorité au service du Grenier.",
        source=SRC_GRENIER + " ; " + SRC_PREC[9:])
sp = C["sensibilite_pertes_grenier"]
contenu({"titre": "Sensibilité : pertes sur prêts", "sous_titre": "Même une perte totale sur les prêts très en retard reste limitée en montant",
         "synthese": f"Encours en retard de plus de 180 jours : {F('Encours 180', sp['encours_180j'])} ; de plus de 90 jours : {F('Encours 90', sp['encours_90j'])}"},
        "Trois taux de perte théoriques (25 %, 50 %, 100 %) appliqués à l'encours en retard. Ce ne sont pas des estimations de perte, mais des ordres de grandeur pour fixer une provision. La politique de provisionnement doit être fixée par le FPPN.",
        tableau={"titre": tab("Provision selon le taux de perte"), "entetes": ["Taux de perte", "Sur l'encours > 180 jours", "Sur l'encours > 90 jours"], "largeurs": [5.0, 7.2, 7.2],
                 "lignes": [[pc(f"Perte {p['perte']}", p["perte"]), F(f"Prov180 {p['perte']}", p["provision_180j"]), F(f"Prov90 {p['perte']}", p["provision_90j"])] for p in sp["scenarios"]]},
        disposition="droite",
        analyse=[f"**Une perte totale sur les prêts en retard de plus de 90 jours coûterait {F('Perte max 90', sp['scenarios'][-1]['provision_90j'])}.** Ce montant reste faible au regard de l'encours.",
                 "**Le risque du Grenier n'est donc pas la perte sur prêts.** Il tient au volume de ressources du FPPN immobilisées dans les prêts."],
        implication="Une politique de provisionnement écrite (taux par tranche de retard) suffirait à couvrir ce risque, mais elle ne règle pas la question de la liquidité de l'épargne.",
        source=SRC_GRENIER)
retenir("Grenier", ("Le Grenier croît très vite grâce au financement du FPPN", "Encours multiplié par 132, financé à " + pc("Part financement (ret)", b[2026]["grenier_financement"] / GP["encours"]) + ", retards maîtrisés mais précomptes incomplets",
                    "Le Grenier a multiplié ses prêts en s'appuyant presque entièrement sur les ressources du FPPN."),
        f"L'encours atteint {F('Encours (ret)', GP['encours'])} au 31/08/2026. Les précomptes couvrent 76 à 85 % des échéances, et 126 prêts ont plus de 180 jours de retard.",
        "Le principal risque n'est pas la perte sur prêts, limitée, mais l'immobilisation de l'épargne des mutualistes dans les prêts.",
        "Le FPPN doit transmettre une politique de crédit lisible, les encaissements hors Solde et la décision qui encadre le financement du Grenier.",
        "Le risque principal n'est pas la perte sur prêts, limitée, mais la mobilisation de l'épargne des mutualistes pour financer cette croissance.")

# ---------------------------------------------------------------- PARTIE 8 : RETRAITÉS ET ANNEXES
separateur(8, "Retraités et annexes", "CGRAE, flotte automobile, MEDOC, projets immobiliers")
contenu({"titre": "Retraités CGRAE", "sous_titre": "Les retraités cotisants sont passés de 16 à 610 en deux ans et demi",
         "synthese": f"Précompte moyen de {F('Moyen CGRAE (r)', sum(c['moyen'] for c in CG) / len(CG))} par mois sur la pension ; données depuis février 2024"},
        "Les retraités cotisent sur leur pension par la CGRAE. Leur coût santé, supérieur à celui des actifs, devra être comparé à leur cotisation dès réception des factures. La convention IPS-CGRAE, reçue le 22/10/2026, ne fixe pas la cotisation.",
        graphique={"titre": fig("Retraités précomptés par la CGRAE, par mois"), "type": "lignes", "categories": [c["libelle"] for c in CG],
                   "series": [{"nom": "Retraités", "valeurs": [c["retraites"] for c in CG], "couleur": BLEU}], "format": "0", "etiquettes": False, "legende": None,
                   "unite": "En nombre de retraités"},
        analyse=[f"**{N('Retraités août 2026', CG[-1]['retraites'])} retraités sont précomptés en août 2026.** Leur nombre augmente chaque mois depuis février 2024.",
                 "**Plusieurs anomalies de gestion apparaissent dans les fichiers** : précomptes en double ou indus remboursés, numéros de décision erronés, précomptes non effectués.",
                 f"**La CGRAE précompte {pc('Recouvrement CGRAE (r)', Q['effectifs']['recouvrement_cgrae'])} des montants demandés** sur {N('Mois CGRAE (r)', Q['effectifs']['recouvrement_cgrae_mois'])} mois communs."],
        implication="Si la garantie santé est maintenue à vie aux retraités pour une cotisation inférieure à son coût, elle devient un engagement viager : sa valeur se calculera avec la table CIMA F (partie 11) dès que les âges et les coûts de soins seront transmis.",
        source=SRC_PREC)
# ---- v9 : convention IPS-CGRAE (partie 8)
contenu({"titre": "Convention IPS-CGRAE : santé des policiers retraités", "sous_titre": "Le FPPN assure la santé des retraités souscripteurs, même si la CGRAE reverse en retard",
         "synthese": "La convention organise le précompte des cotisations sur les pensions ; elle confirme une garantie santé des retraités, donc un risque long"},
        "La convention de partenariat entre l'IPS-CGRAE et le FPPN fixe les conditions du précompte, sur la pension de retraite, des cotisations de prise en charge médicale des policiers retraités. Le montant de la cotisation ne figure pas dans la convention : il relève du bulletin de souscription (annexe 2), non transmis.",
        tableau={"titre": tab("Clauses de la convention IPS-CGRAE utiles à la mission"), "entetes": ["Article", "Clause", "Effet sur la mission"], "largeurs": [3.0, 15.2, 13.0], "taille": 8.5,
                 "lignes": [["Art. 3", "Précompte, sur la pension, des cotisations de prise en charge médicale du policier retraité qui l'autorise", "Les retraités deviennent des cotisants à part entière (partie 8 : " + N("Retraités CGRAE (conv)", CG[-1]["retraites"]) + " en " + CG[-1]["libelle"] + ")"],
                            ["Art. 4.1 e)", "Le FPPN ne suspend aucune prestation pour cause de retard de reversement par l'IPS-CGRAE", "Le FPPN porte le risque de trésorerie entre le soin et l'encaissement"],
                            ["Art. 4.2 a)", "Précompte dans la limite de la quotité cessible ; seuls les retraités qui en disposent sont précomptés", "Une partie des retraités souscripteurs peut ne pas être précomptée"],
                            ["Art. 5", "Aucun frais de gestion facturé par l'IPS-CGRAE", "Pas de charge de gestion à prévoir"],
                            ["Art. 7", "En cas de décès, le FPPN reverse les précomptes perçus après le décès", "Dette envers la CGRAE à suivre"]]},
        disposition="dessous",
        analyse=["**La convention confirme que le FPPN assure la santé des policiers retraités qui y souscrivent.** Les statuts (art. 7 et 80) laissaient la question ouverte.",
                 "**La garantie santé des retraités relève des risques longs** : sa tarification et sa provision demandent les âges des retraités et leurs consommations. La cotisation, elle, est fixée par le barème de la CGRAE : 2,5 % de la pension de base nette de la MUGEFCI, plus 5 000 à 15 000 FCFA selon le grade."],
        implication="L'engagement santé des retraités doit être évalué séparément de celui des actifs, avec ses propres cotisations (PBC, items 108 et 114).",
        valable="texte de la convention reçu le 22/10/2026 ; son application (nombre de retraités, montants précomptés) est mesurée en partie 8 de février 2024 à septembre 2026",
        source="Source : CONVENTION DE PARTENARIAT ENTRE IPS-CGRAE ET LE FPPN.pdf (scan lu par reconnaissance de caractères, articles 3 à 7) ; " + SRC_PREC[9:])
# ---- v10 : santé des policiers retraités, ce qui est mesurable (partie 8)
_CGv = R["cgrae"]
_CGs = [c for i, c in enumerate(_CGv) if i % 4 == 0 or i == len(_CGv) - 1]
contenu({"titre": "Santé des retraités : ce que l'on peut mesurer", "sous_titre": f"Les retraités précomptés passent de {N('SR début', _CGv[0]['retraites'])} en {_CGv[0]['libelle']} à {N('SR fin', _CGv[-1]['retraites'])} en {_CGv[-1]['libelle']}",
         "synthese": "Les cotisations des retraités sont mesurées ; leurs prestations, leurs âges et donc l'engagement à vie du FPPN ne le sont pas"},
        "Cotisation moyenne du mois = montant précompté par la CGRAE / nombre de retraités précomptés. L'engagement d'une garantie santé à vie se calcule avec l'âge de chaque retraité, son coût annuel de soins et une table de survie (CIMA F, partie 11).",
        tableau={"titre": tab("Retraités précomptés par la CGRAE"), "entetes": ["Mois", "Retraités précomptés", "Montant précompté", "Cotisation moyenne"], "largeurs": [5.0, 5.0, 6.0, 5.6], "taille": 9,
                 "lignes": [[c["libelle"], N(f"SR n {c['libelle']}", c["retraites"]), F(f"SR m {c['libelle']}", c["montant"]), F(f"SR moy {c['libelle']}", c["montant"] / c["retraites"])] for c in _CGs]},
        disposition="droite",
        analyse=["**Les retraités sont une population cotisante en forte hausse.** Leurs prestations ne sont pas isolées dans les comptes ni dans les fichiers reçus.",
                 f"**Les fichiers retour de la CGRAE de février à avril 2024 donnent la date de naissance de {N('SR dn', RDN['nb'])} retraités, âgés en moyenne de {N('SR âge', RDN['age_moyen'], 1)} ans au 30/09/2026.** Aucun fichier ne donne les dates des autres retraités ni les soins de chacun : l'engagement à vie ne peut pas être calculé (annexe « Méthodes non retenues et pourquoi »)."],
        implication="Tant que ces données manquent, le FPPN ne peut pas savoir si la cotisation des retraités couvre leurs soins jusqu'à la fin de leur vie.",
        valable=f"fichiers retour de la CGRAE reçus le 15/10/2026, {_CGv[0]['libelle']} à {_CGv[-1]['libelle']} (un mois sur quatre affiché, tous les mois dans le classeur) ; ne couvre pas les retraités non précomptés",
        source="Source : " + SRC_PREC[9:] + " ; CONVENTION DE PARTENARIAT ENTRE IPS-CGRAE ET LE FPPN.pdf ; onglet 22 du classeur des calculs")
contenu({"titre": "Produits annexes", "sous_titre": "Deux produits sont précomptés sans documentation transmise",
         "synthese": "Flotte automobile (9W3) : convention avec GNA non reçue ; MEDOC (8Z7) : aucun document ; projets immobiliers suivis en compte d'attente"},
        "Ces produits ne relèvent pas directement des garanties de base, mais ils transitent par le FPPN et engagent sa responsabilité vis-à-vis des adhérents. Aucun compte n'est intitulé « SCI » dans les balances ; les projets immobiliers figurent en comptes 4711120x.",
        tableau={"titre": tab("Produits annexes et points à clarifier"), "entetes": ["Produit", "Ce que montrent les données", "À obtenir du FPPN"], "largeurs": [4.8, 13.2, 13.2],
                 "lignes": [["Flotte automobile (9W3)", "La Solde précompte environ 10 000 FCFA par mois et par adhérent de 2023 à 2026 ; le compte 47111106 « GNA ASSURANCE » suit l'assureur.", "La convention avec GNA, les primes reversées, les sinistres et les soldes par adhérent. [[GNA met mal à jour les soldes (responsable des cotisations, entretien du 19/10/2026)]]."],
                            ["MEDOC (8Z7)", "Le FPPN précompte 5 000 FCFA par mois depuis novembre 2025 ; les retours de février et avril 2026 sont vides.", "La période de fonctionnement et la décision de suspension. [[Produit décrit comme un complément pharmacie, en pause depuis environ deux mois (entretien du 19/10/2026)]]."],
                            ["Projets immobiliers", "Les comptes 47111200 et 47111201 portent des projets immobiliers ; aucun compte n'est intitulé SCI.", "Le statut juridique des projets et le lien éventuel avec la SCI « La Prévoyance »."],
                            ["Identifiants « 183FPPN »", "Les fichiers aller PEP et Grenier de 2025 et 2026 contiennent 34 identifiants sans mécano.", "L'identité et le mécano de ces personnes."]]},
        disposition="dessous",
        analyse=["**Ces activités engagent le FPPN sans que leurs règles soient documentées.** Le FPPN encaisse des précomptes pour le compte de tiers ou pour des produits dont la nature n'est pas écrite."],
        implication="Si le FPPN porte lui-même le risque de la flotte automobile, il exerce une activité d'assurance de dommages qui sort du champ des mutuelles sociales : la convention doit le préciser.",
        source=SRC_PREC + " ; " + SRC_COMPTA[9:])

# ---------------------------------------------------------------- PARTIE 9 : BILAN
separateur(9, "Bilan, provisions et placements", "Provisions par produit, adossement, fonds propres, rendement")
methode_gab({"titre": "Méthode : lecture du bilan", "sous_titre": "Le bilan est reconstitué compte par compte à partir des balances",
             "synthese": "Provisions, épargne, actifs et fonds propres sont calculés à chaque clôture, de 2020 au 31/08/2026"},
            [("Provisions techniques", "Adenka lit les comptes 19011000 et 19012000 jusqu'en 2023, puis 15801000 et 15802000 à partir de 2024 (prestations à payer)."),
             ("Fonds propres", "Les fonds propres additionnent les comptes 10, 11 et 12 et le résultat de l'exercice, à chaque clôture."),
             ("Adossement par échéance", "Adenka classe l'actif selon son délai de mobilisation et le passif selon son échéance, puis compare les deux, tranche par tranche."),
             ("Limite", "Le bilan est reconstitué compte par compte, puis confronté aux états financiers de 2020 au 31/08/2026 reçus le 22/10/2026 (partie 3). La balance au 31/08/2026 est provisoire.")],
            "Le bilan est reconstitué à partir des soldes des balances et confronté aux états financiers reçus le 22/10/2026. Fonds propres = réserves + report à nouveau + pertes reportées + résultat de l'exercice.")
pvx = EX["provisions"]
contenu({"titre": "Calcul pas à pas : provisions", "sous_titre": "Les deux comptes de provisions portent exactement le même montant sept ans de suite",
         "synthese": f"Maladie {F('Prov mal', b[2020]['provision_maladie'])} et obsèques {F('Prov obs', b[2020]['provision_obseques'])} à chaque clôture"},
        "Lecture : (1) même montant au franc près de 2020 à 2026 ; (2) seul le numéro de compte change en 2024 (reclassement du compte 19 vers le compte 158) ; (3) une estimation recalculée chaque année varierait avec les dossiers en attente. La provision n'est donc pas une estimation actuelle.",
        pas_a_pas={"entetes": ["Clôture", "Compte maladie", "Provision maladie", "Compte obsèques", "Provision obsèques"],
                   "lignes": [[a if a != "2026" else "31/08/2026", pvx[a][0][0], nb(-float(pvx[a][0][2])) + " FCFA", pvx[a][1][0], nb(-float(pvx[a][1][2])) + " FCFA"] for a in sorted(pvx)],
                   "largeurs": [3.6, 4.0, 6.8, 4.0, 6.8], "rouge": [(0, 2), (6, 2), (4, 1)],
                   "explications": ["La provision maladie vaut le même montant en 2020…", "… et au 31/08/2026, au franc près.", "Seul le numéro de compte change en 2024, du compte 19 vers le compte 158."],
                   "fichier": "AUDIT ADENKA\\DAFC\\2020 à 2026 (balances des comptes)"},
        analyse=["**La provision n'est pas recalculée.** Un montant identique au franc près pendant sept ans ne peut pas résulter d'une évaluation des dossiers en cours."],
        implication="Le règlement d'exécution n° 004/2022 impose une provision pour prestations à payer calculée exercice par exercice, dossier par dossier, augmentée des sinistres non déclarés (art. 10). Le compte 158 est par ailleurs réservé aux provisions réglementées à caractère fiscal.",
        source=SRC_COMPTA)
pm_ = C["provisions_mois"]
contenu({"titre": "Provisions techniques", "sous_titre": "La provision ne couvre plus que 21 mois de prestations, contre 115 en 2021",
         "synthese": f"Provisions / prestations mensuelles : {N('Mois 2021', pm_[1]['mois_couverts'], 2)} mois en 2021, {N('Mois 2025', pm_[5]['mois_couverts'], 2)} en 2025, {N('Mois 2026', pm_[6]['mois_couverts'], 2)} en 2026"},
        "Le nombre de mois couverts n'est pas une norme : c'est un indicateur de cohérence. Une provision figée perd son sens quand les prestations changent : elle était très supérieure aux besoins en 2021 et diminue en proportion depuis. Seul un calcul fondé sur les dossiers dira si elle est suffisante.",
        graphique={"titre": fig("Mois de prestations couverts par les provisions"), "type": "barres", "categories": cats_ex,
                   "series": [{"nom": "Mois couverts", "valeurs": [round(p["mois_couverts"], 2) for p in pm_], "couleur": BLEU}], "format": "0.00", "legende": None,
                   "unite": "En nombre de mois de prestations"},
        analyse=[f"**La provision vaut {F('Prov tot (p)', b[2026]['provisions_techniques'])} chaque année.**",
                 f"**Les prestations mensuelles passent de {F('Prest mens 2021', pm_[1]['prestations_mensuelles'])} en 2021 à {F('Prest mens 2026', pm_[6]['prestations_mensuelles'])} en 2026.** La provision représente donc de moins en moins de mois de prestations.",
                 "**La provision nécessaire est mesurée sur les factures arrivées après la clôture** (" + ref("Provision maladie : montant nécessaire") + ") : quelques millions de FCFA, très loin du montant au bilan."],
        implication="Si les délais de paiement s'allongent pendant que les prestations augmentent, la provision deviendra insuffisante : son recalcul est la première tâche dès réception des bases de prestations.",
        source=SRC_COMPTA)
# ---- v5 : provisions par produit, PSAP, méthode CIMA, Grenier, GNA, retraite du personnel (partie 9)
PR_ = {r["exercice"]: r for r in PV["psap_rapprochement"]}
contenu({"titre": "Provisions attendues par produit", "sous_titre": "Seules la maladie et le décès ont une provision, figée depuis 2020",
         "synthese": "Le règlement d'exécution n° 004/2022 prévoit quatre provisions techniques ; le FPPN n'en a constitué qu'une, sans la recalculer"},
        "Le règlement d'exécution n° 004/2022 (art. 8 à 11) prévoit la provision pour risques en cours, la provision pour prestations à payer (appelée ici PSAP, provision pour sinistres à payer), la provision mathématique, et toute autre provision fixée par la Commission de l'UEMOA. Une dépréciation d'actif (prêts) n'est pas une provision technique : elle relève du PCMS et du principe de prudence.",
        tableau={"titre": tab("Provisions exigées et provisions au bilan, par produit"), "entetes": ["Produit", "Provision exigée (texte)", "Au bilan du FPPN", "Constat"], "largeurs": [4.6, 9.4, 7.0, 10.2], "taille": 9,
                 "lignes": [["Maladie", "PSAP dossier par dossier et sinistres tardifs (art. 10)", F("PSAP maladie (p9)", PV["psap_maladie"]["2025"]), "Montant hérité d'environ 2004, non calculé ; besoin mesuré de quelques millions de FCFA (méthode courte)."],
                            ["Décès (obsèques)", "PSAP dossier par dossier et sinistres tardifs (art. 10)", F("PSAP décès (p9)", PV["psap_deces"]["2025"]), "Montant hérité d'environ 2004, inchangé ; le calcul annuel des allocations restant à payer est passé au compte 41591000."],
                            ["Toutes garanties", "Provision pour risques en cours (art. 9)", "Aucune", "Cotisation précomptée sur le mois qu'elle couvre : provision nulle (statuts, art. 88)."],
                            ["EFC, PEP", "Provision mathématique si un capital est garanti (art. 11)", "Dettes d'épargne seulement", "Intérêts dus non comptabilisés : " + F("PPP int", INT55) + " au 31/08/2026 (PEP 3,5 %, EFC 5,5 %)."],
                            ["Santé des retraités", "Provision mathématique si garantie à vie (art. 11)", "Aucune", "Âges et règles de la garantie non reçus."],
                            ["Prêts du Grenier", "Dépréciation des créances (PCMS, prudence)", "Aucune depuis 2022", "Dépréciation calculée avec la politique de crédit du Grenier (40 %, 80 %, 100 %)."],
                            ["Flotte automobile", "Aucune provision technique (risque porté par GNA)", "Compte de tiers 47111106", "Trop-perçus à traiter comme une dette envers les adhérents."],
                            ["Personnel du FPPN", "Indemnités de départ à la retraite (SYSCOHADA révisé)", F("PIDR 2025 (p9)", PV["pidr"][5]["provision_196"]), "Provision annuelle au compte 196, méthode non communiquée."]]},
        disposition="dessous",
        analyse=["**La revue produit par produit montre trois manques : provision non recalculée, provisions mathématiques absentes, dépréciation des prêts absente.** Le sens des écritures est correct (provisions au passif, au crédit), mais leur montant n'est pas une estimation."],
        implication="Le FPPN ne peut pas justifier « à tout moment » des provisions suffisantes, comme l'exige l'article 7 du règlement d'exécution n° 004/2022. Les diapositives suivantes détaillent chaque point.",
        source="Source : " + REX + ", art. 7 à 11 ; PCMS ; " + SRC_COMPTA[9:])
contenu({"titre": "PSAP et dettes envers les cliniques", "sous_titre": "Les factures des cliniques déjà enregistrées ne représentent qu'une petite part de la provision",
         "synthese": f"2025 : PSAP maladie {F('PSAP mal 2025 (c)', PR_[2025]['psap_maladie'])}, dettes envers les cliniques {F('Dettes cliniques 2025', PR_[2025]['dettes_cliniques_40110000'])}"},
        "Une facture de clinique enregistrée en dette (compte 40110000) est déjà au passif : elle n'a pas à être provisionnée une seconde fois. La PSAP doit couvrir les soins réalisés avant la clôture mais pas encore facturés ou enregistrés, plus les sinistres non déclarés (art. 10). Comparer les deux montants permet de juger la cohérence de la provision.",
        graphique={"titre": fig("PSAP maladie et dettes envers les cliniques"), "type": "barres", "categories": cats_ex,
                   "series": [{"nom": "PSAP maladie", "valeurs": [round(PR_[a]["psap_maladie"]) for a in range(2020, 2027)], "couleur": BLEU},
                              {"nom": "Dettes envers les cliniques (40110000)", "valeurs": [round(PR_[a]["dettes_cliniques_40110000"]) for a in range(2020, 2027)], "couleur": ORANGE}],
                   "format": FMT_M, "unite": "En millions de FCFA"},
        analyse=[f"**Les factures enregistrées représentent {N('Dettes cliniques mois 2025', PR_[2025]['dettes_cliniques_en_mois'], 2)} mois de prestations en 2025, la PSAP {N('PSAP mois 2025', PR_[2025]['psap_en_mois_de_prestations'], 2)} mois.** La PSAP vaut {N('Rapport PSAP 2025', PR_[2025]['psap_sur_dettes_cliniques'], 2)} fois les dettes envers les cliniques.",
                 f"**La provision décès représente {N('PSAP décès mois 2025', PR_[2025]['psap_deces_en_mois'], 2)} mois de prestations décès de 2025.** Un tel niveau n'a de sens que s'il correspond à des dossiers anciens non payés, dont la liste a été demandée.",
                 "**Sans la base des prestations, on ne peut dire si la provision est trop forte ou trop faible.** On peut seulement dire qu'elle ne suit ni les dettes, ni les prestations."],
        implication="Les trois diapositives suivantes mesurent la provision maladie nécessaire avec la méthode courte que permettent le règlement intérieur et les factures reçues.",
        source=SRC_COMPTA)
# ---- v10 : provision maladie, méthode courte (statuts, art. 89 ; règlement intérieur, art. 52, 57 et 60)
_PMA = [x for x in V10["psap_maladie_delais"]]
contenu({"titre": "Circuit d'une facture de soins", "sous_titre": "La date des soins n'est saisie nulle part : la comptabilité ne connaît que la date d'enregistrement de la facture",
         "synthese": "Le règlement intérieur fixe trois délais : paiement des cliniques en trois mois, demande de remboursement en trois mois, remboursement en six mois"},
        "Logigramme modifiable dans PowerPoint, tiré du règlement intérieur (art. 52, 57 et 60) et du journal SIN des grands livres des tiers 2023 à 2025. Cadre rouge : point de contrôle manquant.",
        logigramme="facture",
        analyse=["**La facture est enregistrée au compte 40110000 à sa réception, avec son numéro et sa date d'enregistrement.** La date des soins n'apparaît que dans le numéro de facture, et seulement pour une partie des factures.",
                 "**La provision maladie se mesure donc sur les factures d'une année enregistrées l'année suivante** (diapositives suivantes). Une base des prestations datée donnerait la date des soins de chaque facture."],
        implication="Saisir la date des soins et la date de dépôt de chaque facture permet de calculer chaque année la provision maladie et le délai réel de paiement des cliniques.",
        source="Source : " + SRC_RI + ", art. 52, 57 et 60 ; grands livres des tiers 2023 à 2025 (journal SIN)")
contenu({"titre": "Provision maladie : délais de paiement des cliniques", "sous_titre": f"Le délai moyen de paiement passe de {N('PM délai 2020', DELAI[2020], 1)} mois en 2020 à {N('PM délai 2025', DELAI[2025], 2)} mois en 2025",
         "synthese": "La règle des trois mois (règlement intérieur, art. 52) est approchée depuis 2024 mais pas respectée chaque année ; elle justifie une provision calculée sur quelques mois"},
        "Le règlement intérieur fixe le paiement d'une facture de clinique à trois mois après son dépôt (art. 52), le remboursement d'un mutualiste à six mois après sa demande, elle-même limitée à trois mois (art. 57 et 60). Délai moyen = factures non payées en fin de période / factures moyennes d'un mois.",
        tableau={"titre": tab("Factures de cliniques (compte 40110000) et délai moyen de paiement"), "entetes": ["Exercice", "Factures enregistrées", "Factures payées", "Non payées en fin de période", "Délai moyen (mois)", "Prestations maladie"],
                 "largeurs": [3.2, 5.6, 5.6, 5.8, 4.0, 6.0], "taille": 9,
                 "lignes": [[str(x["exercice"]) if x["exercice"] < 2026 else "31/08/2026", F(f"PM fact {x['exercice']}", x["factures_enregistrees"]), F(f"PM paie {x['exercice']}", x["paiements"]),
                             F(f"PM dette {x['exercice']}", x["dettes_cliniques"]), N(f"PM délai {x['exercice']} (t)", DELAI[x["exercice"]], 2), F(f"PM prest {x['exercice']}", x["prestations_maladie"])] for x in _PMA]},
        disposition="dessous",
        analyse=[f"**De 2020 à 2022, les factures non payées représentent de {N('PM délai min arr', min(DELAI[a] for a in (2020, 2021, 2022)), 0)} à {N('PM délai max arr', max(DELAI[a] for a in (2020, 2021, 2022)), 0)} mois de factures : ce sont des arriérés, non un délai normal.** Depuis 2024, le délai moyen est compris entre {N('PM délai min réc', min(DELAI[a] for a in (2024, 2025, 2026)), 2)} et {N('PM délai max réc', max(DELAI[a] for a in (2024, 2025, 2026)), 2)} mois.",
                 "**Depuis 2024, le délai moyen reste proche de la règle des trois mois (art. 52).** Cette règle limite la durée pendant laquelle une prestation reste à payer, ce qui rend inutile une méthode sur plusieurs années (annexe des méthodes non retenues)."],
        implication="La provision maladie se calcule sur quelques mois : factures reçues non payées (déjà au passif) et soins de l'année facturés plus tard.",
        valable="factures de cliniques enregistrées au compte 40110000, exercices 2020 à 2025 et du 01/01 au 31/08/2026 ; ne couvre pas les remboursements aux mutualistes ni les bons payés hors de ce compte",
        source="Source : " + SRC_RI + ", art. 52, 57 et 60 ; " + SRC_COMPTA[9:] + " ; onglet 18a du classeur des calculs")
_TM = V10["psap_maladie_tardives"]
contenu({"titre": "Provision maladie : factures arrivées en retard", "sous_titre": "Moins de 1 % des factures d'une année sont enregistrées l'année suivante",
         "synthese": "Les soins d'une année facturés l'année suivante sont le seul montant que la provision maladie doit couvrir en plus des factures reçues"},
        "Méthode de la circulaire CIMA n° 00230/2005 (cadence des déclarations tardives), réduite à sa première année : l'année de chaque facture est lue dans son numéro (par exemple « E24-000042 » pour 2024). Une facture enregistrée en 2024 avec un numéro de 2023 est un sinistre de 2023 déclaré en retard.",
        tableau={"titre": tab("Factures de cliniques d'une année enregistrées l'année suivante"), "entetes": ["Clôture", "Grand livre lu", "Factures du journal SIN", "Part dont l'année est lisible", "Factures de l'année de clôture", "Part en retard"],
                 "largeurs": [3.0, 9.4, 5.2, 4.4, 5.2, 4.0], "taille": 9,
                 "lignes": [[f"31/12/{x['cloture']}", x["grand_livre"], F(f"PMT fact {x['cloture']}", x["factures_sinistres"]), pc(f"PMT lisible {x['cloture']}", x["references_identifiees"] / x["factures_sinistres"]),
                             F(f"PMT retard {x['cloture']}", x["references_annee_precedente"]), pc(f"PMT part {x['cloture']}", x["references_annee_precedente"] / x["factures_sinistres"])] for x in _TM]},
        disposition="dessous",
        analyse=[f"**Les factures en retard représentent de {F('PMT min', min(x['references_annee_precedente'] for x in _TM))} à {F('PMT max', max(x['references_annee_precedente'] for x in _TM))} selon la clôture.** Le FPPN enregistre l'essentiel des factures de l'année avant le 31 décembre.",
                 f"**Dans le grand livre 2025, l'année n'est lisible que pour {pc('PMT lisible 2024 (t)', _TM[-1]['references_identifiees'] / _TM[-1]['factures_sinistres'])} des factures** : la clôture 2024 est donc mesurée avec moins de précision."],
        implication="La cadence de première année est faible : la provision pour sinistres tardifs en maladie est de l'ordre de quelques millions de FCFA, non de quelques milliards.",
        valable="factures de cliniques enregistrées au journal SIN en 2023, 2024 et 2025, dont l'année figure dans le numéro ; ne couvre pas les factures sans année lisible",
        source="Source : " + SRC_GLT + " (journal SIN, colonne « N° facture ») ; circulaire n° 00230/CIMA/CRCA/PDT/2005 ; onglet 18b du classeur des calculs")
contenu({"titre": "Provision maladie : montant nécessaire", "sous_titre": f"Le besoin estimé est de quelques millions de FCFA ; la provision du bilan vaut {F('PMN prov', PROV_MAL)}",
         "synthese": f"Au 31/12/2024 : besoin estimé {F('PMN besoin 2024', psap_m_elargie(2024))}, soit {N('PMN fois', PROV_MAL / psap_m_elargie(2024), 0)} fois moins que la provision"},
        "Provision nécessaire au compte 158 = soins de l'année facturés l'année suivante. Les factures déjà reçues sont au passif (compte 40110000) et ne doivent pas être provisionnées une seconde fois. Estimation élargie : factures en retard / part des factures dont l'année est lisible × (prestations maladie / factures de cliniques), pour couvrir les prestations payées hors du compte 40110000.",
        tableau={"titre": tab("Provision maladie nécessaire et provision du bilan"), "entetes": ["Clôture", "Factures reçues non payées (au passif)", "Factures en retard mesurées", "Provision nécessaire (estimation élargie)", "Provision du bilan", "Écart"],
                 "largeurs": [3.0, 5.6, 5.2, 6.0, 5.6, 5.8], "taille": 9,
                 "lignes": [[f"31/12/{c}", F(f"PMN dette {c}", _PM[c]["dettes_cliniques"]), F(f"PMN retard {c}", TARDM[c]["references_annee_precedente"]), F(f"PMN besoin {c}", psap_m_elargie(c)),
                             F(f"PMN prov {c}", PV["psap_maladie"][str(c)]), F(f"PMN écart {c}", PV["psap_maladie"][str(c)] - psap_m_elargie(c))] for c in (2022, 2023, 2024)]},
        disposition="dessous",
        analyse=[f"**La provision maladie du bilan ne correspond à aucun calcul et vaut de {N('PMN fois min', min(PV['psap_maladie'][str(c)] / psap_m_elargie(c) for c in (2022, 2023, 2024)), 0)} à {N('PMN fois max', max(PV['psap_maladie'][str(c)] / psap_m_elargie(c) for c in (2022, 2023, 2024)), 0)} fois le besoin mesuré.** Son montant est identique de 2020 au 31/08/2026, alors que les prestations maladie ont été multipliées par {N('PMN mult prest', s[2025]['prest_maladie'] / s[2020]['prest_maladie'], 1)} de 2020 à 2025.",
                 "**Le besoin n'est pas nul** : l'article 89 des statuts et l'article 10 du règlement d'exécution n° 004/2022 imposent de l'estimer chaque année."],
        implication="L'excédent est en réalité une réserve. Son reclassement relève d'une décision de l'Assemblée générale (statuts, art. 85) et du dialogue avec l'AIRMS ; il ne peut se faire qu'avec une base des prestations portant les dates des soins, du dépôt et du paiement (PBC, item 71).",
        valable="clôtures 2022, 2023 et 2024 ; prestations maladie des balances ; factures de cliniques du journal SIN",
        source="Source : " + SRC_ST + ", art. 85 et 89 ; " + REX + ", art. 10 ; " + SRC_GLT + " ; " + SRC_COMPTA[9:] + " ; onglet 18c du classeur des calculs")
contenu({"titre": "Allocations décès : méthode du FPPN", "sous_titre": "Les dossiers non payés sont valorisés au barème et passés en dette au compte 41591000",
         "synthese": "Les états de fin d'année reçus le 22/10/2026 confirment le circuit décrit en entretien et en donnent le compte"},
        "Logigramme modifiable, tel que décrit par le chef comptable le 19/10/2026 et confirmé par les états de fin d'année 2021 à 2025. Barème appliqué dans les états : sous-officier 1 400 000 FCFA, officier 1 600 000 FCFA, commissaire 2 200 000 FCFA (le barème déclaré en entretien, 1 000 000 et 2 000 000 FCFA, n'est pas celui des fichiers).",
        logigramme="deces",
        analyse=["**La méthode du FPPN se rapproche d'une évaluation dossier par dossier.** La liste des décès transmise par la direction du personnel est rapprochée des allocations payées ; les dossiers restants sont valorisés au barème par grade.",
                 "**Le calcul annuel est comptabilisé en dette au compte 41591000, non au compte de provision 15802000**, qui ne bouge pas depuis 2020 (diapositive suivante).",
                 "**Les décès de conjoints et d'enfants n'apparaissent que s'ils sont déclarés.** Ils relèvent des sinistres tardifs."],
        implication="Le règlement d'exécution n° 004/2022 (art. 10) demande cette évaluation dossier par dossier, plus les sinistres tardifs. La méthode du FPPN peut servir de base, si son résultat est comptabilisé et documenté.",
        source=SRC_ENT16)

# ---- v10 : allocations décès (états de fin d'année reçus le 22/10/2026), risques courts, Grenier
_ET = V10["deces_etats"]
contenu({"titre": "Allocations décès : états de fin d'année et compte 41591000", "sous_titre": "Le calcul de fin d'année du FPPN est comptabilisé au compte 41591000, au franc près en 2024 et 2025",
         "synthese": "Chaque 31 décembre, le FPPN liste les décès non payés, applique le barème du grade et passe le total en dette envers les bénéficiaires"},
        "Barème appliqué dans les fichiers : membre sous-officier 1 400 000 FCFA, officier 1 600 000, commissaire 2 200 000 ; conjoint 600 000 à 700 000 ; enfant 120 000 à 150 000. C'est le barème de la cartographie des risques, non celui déclaré en entretien le 19/10/2026.",
        tableau={"titre": tab("États de fin d'année des allocations décès et crédits du compte 41591000"), "entetes": ["Exercice", "Fichier", "Dossiers", "Total du fichier", "Crédit au 41591000", "Écart"],
                 "largeurs": [2.6, 12.0, 2.8, 4.6, 4.8, 4.4], "taille": 8.5,
                 "lignes": [[str(x["exercice"]), x["fichier"], N(f"DC dossiers {x['exercice']}", x["dossiers_valorises"]), F(f"DC total {x['exercice']}", x["total_imprime"]),
                             F(f"DC crédit {x['exercice']}", x["credit_41591000"]), F(f"DC écart {x['exercice']}", x["credit_41591000"] - x["total_imprime"])] for x in _ET]},
        disposition="dessous",
        analyse=["**En 2024 et 2025, le crédit de l'année au compte 41591000 est exactement le total du fichier.** Le grand livre 2024 montre l'écriture du 31/12/2024 « PROVISION ALLO DECES 2024 » de 140 800 000 FCFA.",
                 "**En 2022 et 2023, les crédits du compte diffèrent des états ; en 2021, le premier crédit reprend aussi des dossiers antérieurs à l'ouverture du compte.** Le fichier 2024 porte la mention « BILAN NON EXHAUSTIF » et la somme de ses lignes dépasse son total de 1 400 000 FCFA."],
        implication="La méthode du FPPN est bien une évaluation dossier par dossier (statuts, art. 89). Elle est comptabilisée comme une dette, non comme une provision : la provision obsèques ne doit pas la reprendre une seconde fois.",
        valable="états de fin d'année 2021 à 2025 tels que reçus le 22/10/2026 ; crédits du compte 41591000 des balances 2021 à 2025",
        source="Source : " + SRC_LOT4C + " ; « GRAND LIVRE FPPN 2024.xlsx » (compte 41591000) ; « CARTOGRAPHIE DES RISQUES COUVERTS ET DES PRESTATIONS OFFERTES.docx » ; balances 2021 à 2025 ; onglet 19a du classeur des calculs")
_DS = [x for x in V10["deces_stock"] if x["exercice"] >= 2021]


def _dette(x):
    return (x["ouv_41591000"] + x["mvt_c_41591000"] - x["mvt_d_41591000"]) + (x["ouv_41592000"] + x["mvt_c_41592000"] - x["mvt_d_41592000"])


contenu({"titre": "Allocations décès : dette réelle et provision obsèques", "sous_titre": f"La provision obsèques vaut {N('DD fois 2026', PROV_OBS / DETTE_DEC26, 1)} fois la dette décès calculée par le FPPN au 31/08/2026",
         "synthese": f"Dette décès (41591000 et 41592000) : {F('DD dette 2026', DETTE_DEC26)} ; provision obsèques inchangée depuis 2020 : {F('DD prov', PROV_OBS)}"},
        "Dette en fin d'année = dette du début + allocations ajoutées par l'état de fin d'année - allocations payées. La provision obsèques (comptes 19012000 puis 15802000) n'est rattachée à aucun calcul transmis.",
        tableau={"titre": tab("Dette décès du FPPN et provision obsèques"), "entetes": ["Exercice", "Allocations ajoutées", "Allocations payées", "Dette allocations", "Dette frais funéraires", "Dette décès totale", "Provision obsèques"],
                 "largeurs": [3.0, 4.6, 4.6, 4.6, 4.4, 4.6, 4.8], "taille": 9,
                 "lignes": [[str(x["exercice"]) if x["exercice"] < 2026 else "31/08/2026", F(f"DD ajout {x['exercice']}", x["mvt_c_41591000"]), F(f"DD payé {x['exercice']}", x["mvt_d_41591000"]),
                             F(f"DD all {x['exercice']}", x["ouv_41591000"] + x["mvt_c_41591000"] - x["mvt_d_41591000"]), F(f"DD ff {x['exercice']}", x["ouv_41592000"] + x["mvt_c_41592000"] - x["mvt_d_41592000"]),
                             F(f"DD tot {x['exercice']}", _dette(x)), F(f"DD prov {x['exercice']}", x["provision_obseques"])] for x in _DS]},
        disposition="dessous",
        analyse=[f"**Chaque année, le FPPN ajoute plus d'allocations qu'il n'en paie : la dette passe de {F('DD 2021', _dette(_DS[0]))} fin 2021 à {F('DD 2026 (a)', DETTE_DEC26)} au 31/08/2026.** Le stock contient donc des dossiers de plusieurs années.",
                 "**En 2024, des allocations ont été payées depuis les comptes bancaires de l'épargne** : 93 paiements pour 49 000 000 FCFA par le compte BOA PEP et 9 pour 4 288 000 FCFA par le compte BOA EFC (grand livre 2024)."],
        implication="La provision obsèques fait double emploi avec la dette du 41591000. La dette elle-même doit être justifiée dossier par dossier, avec la date de dépôt de chaque demande, pour écarter les demandes frappées de forclusion (règlement intérieur, art. 58 : 3 ans pour un membre, 12 mois pour un ayant droit).",
        valable="comptes 41591000, 41592000, 19012000 et 15802000, exercices 2021 à 2025 et du 01/01 au 31/08/2026 (provisoire) ; paiements par banque d'après le grand livre 2024 seulement",
        source="Source : " + SRC_COMPTA[9:] + " ; « GRAND LIVRE FPPN 2024.xlsx » (journaux BQ PEP, BQEFC) ; " + SRC_RI + ", art. 58 ; onglet 19b du classeur des calculs")
_TD = V10["deces_tardifs"]
contenu({"titre": "Allocations décès : décès connus après la clôture", "sous_titre": f"{N('DT nombre', len(_TD))} décès d'une année n'ont été enregistrés que l'année suivante, pour {F('DT montant', TARD_DEC)}",
         "synthese": "Ces décès manquent dans l'état du 31 décembre : ils donnent l'ordre de grandeur de la provision pour décès survenus non déclarés"},
        "Seuls les états 2021 et 2023 portent à la fois la date du décès et la date d'enregistrement du dossier. On retient les dossiers dont le décès est d'une année et l'enregistrement de l'année suivante.",
        tableau={"titre": tab("Décès enregistrés l'année suivant le décès"), "entetes": ["État", "Date du décès", "Date d'enregistrement", "Montant"], "largeurs": [3.6, 6.0, 6.0, 6.0], "taille": 9,
                 "lignes": [[str(x["etat"]), x["deces"][8:10] + "/" + x["deces"][5:7] + "/" + x["deces"][:4], x["enregistrement"][8:10] + "/" + x["enregistrement"][5:7] + "/" + x["enregistrement"][:4],
                             F(f"DT {i}", x["montant"])] for i, x in enumerate(_TD)]},
        disposition="dessous",
        analyse=[f"**Les décès déclarés tardivement représentent {F('DT montant (a)', TARD_DEC)} sur deux états, soit quelques dossiers par an.** Le délai de forclusion de trois ans (art. 58) borne la période pendant laquelle un décès peut encore être déclaré."],
        implication="La provision décès nécessaire est la dette du 41591000, justifiée dossier par dossier, plus une estimation des décès non encore déclarés de quelques millions de FCFA par an.",
        valable="états de fin d'année 2021 et 2023, seuls à porter la date d'enregistrement ; dossiers de membres et d'ayants droit valorisés",
        source="Source : « PROV ALLOCATION DECES 2021.xlsx » (feuille IMPAYEES GLOBAL 2021) ; « ALLOC DECES 2023 PAYER ET A PROVISIONNER .xlsx » (feuille ALL DECES  A PROVISIONNER 2023) ; onglet 19c du classeur des calculs")
contenu({"titre": "Provision pour risques courts", "sous_titre": "Aucune provision n'est nécessaire : les cotisations mensuelles couvrent leur propre mois",
         "synthese": "La provision pour risques courts (statuts, art. 88) couvre la période entre la clôture et la prochaine échéance de cotisation ; cette période est nulle au FPPN"},
        "Les cotisations maladie et décès sont précomptées chaque mois sur la solde du mois (postes 936 et 8Z6) ou sur la pension (CGRAE) : la cotisation de décembre couvre décembre. Le Code CIMA (art. 334-10), référence technique des assureurs, exclut de même les cotisations « à terme échu » du calcul de la provision pour risques en cours.",
        tableau={"titre": tab("Provision pour risques courts au 31 décembre"), "entetes": ["Exercice", "Prestations et frais de l'année", "Jours non couverts par une cotisation", "Provision"],
                 "largeurs": [3.6, 8.0, 7.0, 6.0], "taille": 9,
                 "lignes": [[str(a), F(f"RC charges {a}", s[a]["prest_maladie"] + s[a]["prest_deces"] + s[a]["frais_gestion"]), "0", F(f"RC prov {a}", 0)] for a in range(2020, 2026)]},
        disposition="dessous",
        analyse=["**La provision pour risques courts est nulle tant que les cotisations sont précomptées sur la solde ou la pension du mois qu'elles couvrent.** Elle deviendrait nécessaire si des cotisations étaient encaissées d'avance pour plusieurs mois."],
        implication="Le FPPN n'a pas de provision pour risques courts à constituer ; il doit le justifier par écrit dans sa note de calcul annuelle des provisions.",
        valable="cotisations maladie et décès précomptées par la Solde (postes 936 et 8Z6, janvier 2024 à septembre 2026) et par la CGRAE (février 2024 à septembre 2026)",
        source="Source : " + SRC_ST + ", art. 88 ; Code CIMA, art. 334-10 (référence technique) ; " + SRC_PREC[9:] + " ; onglet 20 du classeur des calculs")
_GRD10 = V10["grenier_depreciation"]
contenu({"titre": "Prêts du Grenier : dépréciation selon la politique de crédit", "sous_titre": f"La politique du Grenier conduit à déprécier {F('GD 2026', GRD['2026-08-31'])} au 31/08/2026 ; les comptes n'en portent aucune",
         "synthese": "Taux de la politique de crédit (chapitre XI.3) : 40 % de 3 à 6 mois de retard, 80 % de 6 à 12 mois, 100 % de 12 à 24 mois"},
        "Dépréciation = encours de la tranche de retard × taux de la politique de crédit du Grenier. Les tranches du logiciel de prêts (91 à 180 jours, 181 à 365 jours, plus de 365 jours) correspondent aux trois paliers.",
        tableau={"titre": tab("Dépréciation des prêts du Grenier selon sa politique de crédit"), "entetes": ["Arrêté", "Tranche de retard", "Prêts", "Encours", "Taux", "Dépréciation"], "largeurs": [3.6, 5.6, 2.8, 5.6, 3.0, 5.4], "taille": 8.5,
                 "lignes": [[x["arrete"][8:10] + "/" + x["arrete"][5:7] + "/" + x["arrete"][:4], x["tranche"], N(f"GD prêts {x['arrete']} {x['tranche']}", x["prets"]), F(f"GD enc {x['arrete']} {x['tranche']}", x["encours"]),
                             pc(f"GD taux {x['arrete']} {x['tranche']}", x["taux"]), F(f"GD dep {x['arrete']} {x['tranche']}", x["encours"] * x["taux"])] for x in _GRD10 if x["taux"] > 0]},
        disposition="dessous",
        analyse=[f"**La dépréciation passe de {F('GD 2025', GRD['2025-12-31'])} au 31/12/2025 à {F('GD 2026 (a)', GRD['2026-08-31'])} au 31/08/2026.** Le montant reste faible au regard de l'encours, mais il doit être comptabilisé."],
        implication="Le risque principal du Grenier reste l'immobilisation de l'épargne dans les prêts, non la perte sur prêts.",
        valable="prêts du Grenier aux arrêtés du 31/12/2025 et du 31/08/2026, tranches de retard du logiciel de prêts",
        source="Source : « POLITIK DE CREDIT DU GRENIER VFVF.doc », chapitre XI.3 ; " + SRC_GRENIER[9:] + " ; onglet 23 du classeur des calculs")
_omis(_snap(), {"titre": "Méthode CIMA des sinistres tardifs", "sous_titre": "Le texte UEMOA exige d'estimer les sinistres non déclarés sans dire comment",
         "synthese": "En attente des bases de prestations : la méthode CIMA sert de référence technique, appliquée garantie par garantie"},
        "L'article 10 du règlement d'exécution n° 004/2022 impose d'ajouter à la PSAP « une estimation du coût des risques survenus mais non déclarés », sans fixer de méthode. Le Code CIMA, applicable aux assureurs, en fixe une : elle sert ici de référence technique. Exemple simple : si, en moyenne, 20 sinistres de plus sont déclarés l'année suivante pour 100 déclarés dans l'année, on ajoute 20 % de sinistres au coût moyen observé.",
        tableau={"titre": tab("Étapes de la méthode CIMA des sinistres tardifs"), "entetes": ["Étape", "Calcul", "Explication simple"], "largeurs": [4.0, 12.6, 14.6], "taille": 9,
                 "lignes": [["1. Triangle", "N(n, n + k) : sinistres survenus l'année n et déclarés l'année n + k", "On compte les sinistres selon l'année des soins ou du décès et l'année où le FPPN les apprend."],
                            ["2. Taux de retard", "N(n, n + k) / N(n, n)", "Pour chaque année, on rapporte les déclarations tardives aux déclarations de l'année."],
                            ["3. Coefficient", "coef(k) = Σ des moyennes des taux de retard au-delà de k", "On fait la moyenne sur les années connues, puis on additionne les retards encore à venir."],
                            ["4. Nombre de tardifs", "coef × N(n, n)", "On estime combien de sinistres de chaque année ne sont pas encore connus."],
                            ["5. Coût moyen", "Charge cumulée / nombre cumulé de sinistres", "On mesure ce que coûte un sinistre en moyenne."],
                            ["6. Provision", "Σ nombre de tardifs × coût moyen ; PSAP = dossiers connus + tardifs", "On ajoute ce montant à la provision des dossiers connus."]]},
        disposition="dessous",
        analyse=["**Le calcul demande, pour chaque dossier, la date de survenance, la date de déclaration et le montant.** Ces bases ont été demandées pour la maladie et le décès et ne sont pas encore reçues (PBC)."],
        implication="Analyse en attente des documents. Le rapport présentera, pour chaque garantie, la PSAP du FPPN, la PSAP dossier par dossier et la PSAP selon la méthode CIMA, avec l'écart et son effet sur le résultat.",
        source="Source : " + REX + ", art. 10 ; circulaire n° 00230/CIMA/CRCA/PDT/2005")
GT = [r for r in PV["grenier_tranches"]]
gt25 = {r["tranche"]: r for r in GT if r["arrete"] == "2025-12-31"}
gt26 = {r["tranche"]: r for r in GT if r["arrete"] == "2026-08-31"}
_omis(_snap(), {"titre": "Dépréciation des prêts du Grenier", "sous_titre": "Aucune dépréciation n'est passée, alors que des prêts ont plus d'un an de retard",
         "synthese": f"Encours en retard de plus de 90 jours : {F('Grenier 90 2025', sum(gt25[t]['encours'] for t in list(gt25)[3:]))} au 31/12/2025 et {F('Grenier 90 2026', sum(gt26[t]['encours'] for t in list(gt26)[3:]))} au 31/08/2026"},
        "Un prêt en retard est un actif qui risque de ne pas être remboursé : il se déprécie (PCMS, principe de prudence). Ce n'est pas une provision technique, qui couvre une dette envers les adhérents. Les textes de la mutualité ne donnent pas de taux : la référence retenue est le barème de la BCEAO pour les systèmes financiers décentralisés (SFD), proche du Grenier (petits prêts à des membres). Ce barème est publié sur bceao.int, rubrique réglementation des SFD ; le texte est à joindre au dossier.",
        tableau={"titre": tab("Encours du Grenier par tranche de retard"), "entetes": ["Tranche de retard", "Prêts au 31/12/2025", "Encours au 31/12/2025", "Prêts au 31/08/2026", "Encours au 31/08/2026", "Taux BCEAO"], "largeurs": [4.6, 3.4, 6.0, 3.4, 6.0, 4.0], "taille": 9,
                 "lignes": [[t, N(f"Prêts {t} 2025", gt25[t]["prets"]), F(f"Encours {t} 2025", gt25[t]["encours"]), N(f"Prêts {t} 2026", gt26[t]["prets"]), F(f"Encours {t} 2026", gt26[t]["encours"]), "En attente"] for t in gt25]},
        disposition="dessous",
        analyse=["**La dépréciation sera calculée tranche par tranche dès réception du barème** : encours de la tranche multiplié par le taux de la tranche. Les tranches de plus de 90 jours portent l'essentiel du risque.",
                 "**Le remboursement par précompte réduit le risque, sauf pour les agents sortis ou non précomptés.** Les prêts à mécano anormal (partie 7) sont à examiner un par un."],
        implication="Sans dépréciation, l'actif et les fonds propres sont surévalués, et les ratios de la partie 10 un peu trop favorables. Le montant en jeu reste faible au regard de l'encours (partie 7).",
        source=SRC_GRENIER + " ; PCMS ; BCEAO (barème des SFD, à transmettre)")
GN = [r for r in PV["gna"] if r["exercice"] >= 2023]
_com = {ex: -_so(ex, r"^71800000$") for ex in range(2023, 2027)}
contenu({"titre": "Flotte automobile : risque porté par GNA", "sous_titre": "Le FPPN est apporteur d'affaires de GNA et ne doit pas de provision technique",
         "synthese": "Commission déclarée de 20 % des primes collectées, versée en fin d'année ; trop-perçus à traiter comme une dette envers les adhérents"},
        "Selon le chef comptable et le responsable des cotisations (entretiens du 19/10/2026), les primes vont à GNA, qui porte le risque. Le FPPN est apporteur d'affaires et perçoit une commission de 20 % des primes collectées. Si le FPPN portait le risque, il exercerait une activité d'assurance de dommages, hors du champ de la mutualité sociale (règlement n° 07/2009).",
        tableau={"titre": tab("Compte GNA (47111106) et produits des activités annexes (71800000)"), "entetes": ["Exercice", "Débit du compte GNA", "Solde du compte GNA", "Produits annexes (71800000)"], "largeurs": [3.4, 5.6, 5.6, 6.0],
                 "lignes": [[str(r["exercice"]) if r["exercice"] < 2026 else "31/08/2026", F(f"GNA débit {r['exercice']}", r["mouvements_debit"]), F(f"GNA solde {r['exercice']}", r["solde_debiteur"]), F(f"Produits annexes {r['exercice']}", _com[r["exercice"]])] for r in GN]},
        disposition="droite",
        analyse=["**En 2024, le débit du compte GNA est égal aux produits des activités annexes** : le compte enregistre les commissions dues par GNA. [[Ses crédits sont les versements de GNA]].",
                 f"**La commission comptabilisée en 2024 ne vaut pas 20 % des précomptes 9W3 de l'année ({F('Précomptes 9W3 2024', GNA_P[2024])}).** [[La commission porte aussi sur des primes payées hors précompte]] (convention demandée, PBC, item 89)."],
        implication="Aucune provision technique n'est due par le FPPN. La convention avec GNA doit fixer l'assiette de la commission et le traitement des trop-perçus (PBC, item 89).",
        source="Source : entretiens du 19/10/2026 ; " + SRC_COMPTA[9:] + " ; fichiers de précomptes reçus le 15/10/2026")
PI = [r for r in PV["pidr"] if r["exercice"] <= 2025]
contenu({"titre": "Départ à la retraite du personnel", "sous_titre": "La provision du personnel est reprise puis dotée en totalité chaque année",
         "synthese": f"Provision au compte 196 : {F('PIDR 2020', PI[0]['provision_196'])} en 2020, {F('PIDR 2025', PI[-1]['provision_196'])} en 2025 ; recalcul agent par agent dans les diapositives suivantes"},
        "Cette provision couvre les indemnités dues aux salariés du FPPN à leur départ à la retraite : ce n'est pas un engagement envers les mutualistes. Barème : décret n° 2017-210 du 30/03/2017 (art. 4 et 5), un pourcentage du salaire global mensuel moyen des douze derniers mois par année d'ancienneté (30 % jusqu'à 5 ans, 35 % de 6 à 10 ans, 40 % au-delà), payé à la cessation de service (art. 8) ; plafond de 25 fois le SMIG annuel et départ à 55 ans, plus un an par enfant à charge dans la limite de trois (convention collective interprofessionnelle du 19/07/1977, art. 40). Périmètre : personnel civil salarié (règlement intérieur, art. 31). Méthode attendue : unités de crédit projetées (SYSCOHADA révisé), avec probabilité de présence au départ, hausse des salaires et actualisation.",
        tableau={"titre": tab("Provision pour départ à la retraite du personnel"), "entetes": ["Exercice", "Provision (196)", "Dotation (6911)", "Reprise (7911)"], "largeurs": [3.6, 5.6, 5.6, 5.6],
                 "lignes": [[str(r["exercice"]), F(f"PIDR prov {r['exercice']}", r["provision_196"]), F(f"PIDR dot {r['exercice']}", r["dotation_6911"]), F(f"PIDR rep {r['exercice']}", r["reprise_7911"])] for r in PI]},
        disposition="droite",
        analyse=["**Chaque année, la provision de l'année précédente est reprise en totalité et la nouvelle est dotée en totalité.** Selon le chef comptable (entretien du 19/10/2026), le calcul applique la convention collective à l'ancienneté acquise.",
                 "**Les charges patronales s'ajoutent à l'indemnité** : charges fiscales de 2,80 % (IS 1,20 %, FDFP 0,40 % et 1,20 %), retraite 7,70 % plafonnée à 3 375 000 FCFA, prestations familiales 5,75 % et risques professionnels 2 % (taux du FPPN), plafonnés à 75 000 FCFA.",
                 "**Les DISA reçues le 22/10/2026 donnent l'âge, l'ancienneté et le salaire de chaque agent civil.** Le recalcul est présenté dans les diapositives suivantes."],
        implication="Le montant est faible au regard des provisions techniques, mais il entre dans les dettes à long terme et dans les ratios de la partie 10.",
        source=SRC_COMPTA + " ; décret n° 2017-210 du 30/03/2017 ; convention collective interprofessionnelle du 19/07/1977, art. 40 ; SYSCOHADA révisé")

# ---- v9 : provision pour indemnités de départ à la retraite recalculée agent par agent (partie 9)
PP, PSY, PIN, PSE = V9["pidr_parametres"], V9["pidr_synthese"], V9["pidr_individuel"], V9["pidr_sensibilites"]
SRC_DISA = "Source : DISA_256927_2021.xls, DISA_256927_2022.xls, DISA_256927_2023.xls, DISA_256927_2025.xls (déclarations CNPS) ; ETAT 301 & DISA 2024 FPPN.xls (RECAPSALAIRES 2024 CMP) ; ETAT 301 & DISA 2025 FPPN.xls (RECAPSALAIRES 2025 GRENIER)"
ENT9 = ("Direction Générale", "CMP", "Grenier", "Non documentée")
contenu({"titre": "Provision du personnel : méthode et paramètres", "sous_titre": "Chaque agent civil est évalué à partir de sa déclaration CNPS, avec des paramètres observés ou fixés par les textes",
         "synthese": "Méthode des unités de crédit projetées : indemnité projetée au départ, acquise au prorata de l'ancienneté, pondérée par la survie, la rotation et l'actualisation"},
        "L'indemnité de départ vaut un pourcentage du salaire mensuel moyen par année d'ancienneté (décret n° 2017-210, art. 4 et 5). La méthode des unités de crédit projetées, retenue par IAS 19 et utilisée comme référence de bonne pratique, ne provisionne que la part acquise à la date d'évaluation. Le calcul est fait par formules dans les onglets 14a à 14c du classeur.",
        tableau={"titre": tab("Paramètres de la provision du personnel au 31/12/2025"), "entetes": ["Paramètre", "Valeur", "Justification"], "largeurs": [6.4, 4.0, 20.8], "taille": 8.5,
                 "lignes": [["Âge de départ", N("PIDR âge", PP["age_depart"]) + " ans", "Convention collective interprofessionnelle du 19/07/1977, art. 40 : 55 ans, plus un an par enfant à charge (au plus 3) ; les enfants ne figurent pas dans les fichiers reçus."],
                            ["Barème", "30 % ; 35 % ; 40 %", "Décret n° 2017-210, art. 4 : jusqu'à 5 ans, de 6 à 10 ans, au-delà de 10 ans, fractions d'année au mois inférieur."],
                            ["Salaire", "Brut annuel / mois", "Salaire brut de la DISA divisé par la durée d'activité de l'année (salaire global mensuel moyen, art. 4)."],
                            ["Taux d'actualisation", pc("PIDR taux", PP["taux_actualisation"]), f"Moyenne de 2025 du rendement de l'obligation de l'État de Côte d'Ivoire à 10 ans, {N('PIDR cotations', PP['cotations_2025'])} cotations Investing.com (IAS 19, § 83)."],
                            ["Croissance des salaires", pc("PIDR croissance", PP["croissance_salaires"]), f"Médiane de la croissance annuelle observée sur les {N('PIDR agents croiss', PP['agents_croissance'])} agents présents douze mois en 2021 et en 2025."],
                            ["Rotation du personnel", pc("PIDR rotation", PP["rotation"]), "Moyenne des taux annuels de sortie observés entre les DISA 2021, 2022, 2023 et 2025."],
                            ["Survie", "Table CIMA F", "Table de survie, prudente pour une indemnité versée si l'agent est en vie au départ (principe de prudence, partie 11)."],
                            ["Entité", "DG, CMP, Grenier", "Listes nominatives des récapitulatifs de salaires 2024 (CMP) et 2025 (Grenier) ; agents absents de ces listes : rattachement non documenté."]]},
        disposition="dessous",
        analyse=["**Tous les paramètres viennent d'un texte ou d'une observation sur les fichiers du FPPN.** Seuls l'âge de départ, faute d'enfants déclarés, et le taux d'actualisation, faute de marché profond, relèvent d'un choix justifié."],
        implication="Les effets de l'âge de départ, du taux, de la rotation et des salaires sont mesurés plus loin, une hypothèse à la fois.",
        valable="personnel civil déclaré à la CNPS (DISA), hors travailleurs occasionnels et fonctionnaires de police détachés, évaluation au 31/12",
        source=SRC_DISA + " ; décret n° 2017-210 du 30/03/2017 ; convention collective du 19/07/1977, art. 40 ; IAS 19, § 67 et 83 ; Investing.com ; table CIMA F")
contenu({"titre": "Provision du personnel par entité de 2021 à 2025", "sous_titre": "La provision recalculée est inférieure au compte 196 ; les droits acquis en sont proches",
         "synthese": f"Au 31/12/2025 : provision recalculée {F('PIDR 2025 total', PSY[-1]['dbo'])}, droits acquis {F('PIDR 2025 acquis', PSY[-1]['acquis'])}, compte 196 {F('PIDR 2025 196', PSY[-1]['provision_196'])}"},
        "La provision recalculée suit la méthode des unités de crédit projetées. Les droits acquis sont l'indemnité que chaque agent recevrait s'il partait à la date d'évaluation, sans projection ni actualisation. Le rattachement à une entité suit les listes de 2024 et 2025, appliquées aussi aux années antérieures.",
        tableau={"titre": tab("Provision du personnel par entité, droits acquis et compte 196"), "entetes": ["Exercice", "Agents", "Direction Générale", "CMP", "Grenier", "Non documentée", "Provision recalculée", "Droits acquis", "Compte 196 (FPPN)"],
                 "largeurs": [2.4, 2.2, 3.8, 3.4, 3.2, 3.6, 4.2, 4.2, 4.2], "taille": 8.5,
                 "lignes": [[str(x["annee"]), N(f"PIDR agents {x['annee']}", x["agents"])] + [N(f"PIDR {e} {x['annee']}", x[e]["dbo"]) for e in ENT9]
                            + [N(f"PIDR tot {x['annee']}", x["dbo"]), N(f"PIDR acq {x['annee']}", x["acquis"]), N(f"PIDR 196 {x['annee']}", x["provision_196"])] for x in PSY]},
        disposition="dessous",
        analyse=[f"**Au 31/12/2025, la provision recalculée vaut {pc('PIDR ratio 196', PSY[-1]['dbo'] / PSY[-1]['provision_196'])} du compte 196 ; les droits acquis en valent {pc('PIDR ratio acquis', PSY[-1]['acquis'] / PSY[-1]['provision_196'])}.** Le montant du FPPN est donc proche d'un calcul sans projection ni actualisation, comme l'a décrit le chef comptable.",
                 f"**La Direction Générale porte {pc('PIDR part DG', PSY[-1]['Direction Générale']['dbo'] / PSY[-1]['dbo'])} de la provision 2025 avec {N('PIDR agents DG', PSY[-1]['Direction Générale']['agents'])} agents.** Le CMP compte {N('PIDR agents CMP', PSY[-1]['CMP']['agents'])} agents, dont {N('PIDR CMP depuis 2022', sum(1 for x in PIN if x['annee'] == 2025 and x['entite'] == 'CMP' and int(x['embauche'][-4:]) >= 2022))} recrutés depuis 2022 : leurs droits acquis restent faibles."],
        implication="Le compte 196 couvre la provision recalculée selon les deux approches. [[La méthode du FPPN est celle des droits acquis]] : sa note de calcul est demandée (PBC, item 113).",
        valable="exercices 2021, 2022, 2023 et 2025 (pas de DISA exploitable en 2024), personnel civil déclaré à la CNPS ; montants en FCFA",
        source=SRC_DISA + " ; " + SRC_COMPTA[9:] + " ; onglets 14b et 14c du classeur v9")


def _lignes_pidr(liste):
    return [[x["id"], str(x["naissance"]), x["embauche"], N(f"PIDR âge {x['id']}", x["age"]), N(f"PIDR anc {x['id']}", x["anciennete"], 2),
             N(f"PIDR sal {x['id']}", x["salaire_mensuel"]), N(f"PIDR acq {x['id']}", x["indemnite_acquise"]), N(f"PIDR dbo {x['id']}", x["dbo"])] for x in liste]


_P25 = [x for x in PIN if x["annee"] == 2025]
_ENTETE_I = ["Agent", "Naissance", "Embauche", "Âge", "Ancienneté (ans)", "Salaire mensuel moyen", "Droits acquis", "Provision recalculée"]
_LARG_I = [3.0, 3.0, 3.6, 2.4, 3.6, 5.0, 5.0, 5.6]
def _part5(e):
    l_ = sorted((x["dbo"] for x in _P25 if x["entite"] == e), reverse=True)
    return sum(l_[:5]) / sum(l_)


_nd = [x for x in _P25 if x["entite"] == "Non documentée"]
for _ent, _sous, _morceaux in (("Direction Générale", "Les cinq provisions les plus élevées font " + pc("PIDR part top5 DG", _part5("Direction Générale")) + " de la provision de l'entité", [None]),
                               ("CMP", "Un personnel récent : les droits acquis restent faibles", [0, 1]),
                               ("Grenier", "Trois agents, dont un recruté en 2025", [None]),
                               ("Non documentée", N("PIDR nd n", len(_nd)) + " agents recrutés en " + "2025" + " ne figurent dans aucun récapitulatif de salaires par entité", [0, 1])):
    _l = sorted([x for x in _P25 if x["entite"] == _ent], key=lambda x: -x["dbo"])
    _np = -(-len(_l) // 16)          # au plus 16 agents par diapositive
    _taille_p = -(-len(_l) // _np)
    for _k in range(_np):
        _part = _l[_k * _taille_p:(_k + 1) * _taille_p]
        _suffixe = "" if _np == 1 else f" ({_k + 1}/{_np})"
        _top = max(_part, key=lambda x: x["dbo"])
        contenu({"titre": f"Provision du personnel agent par agent : {_ent if _ent != 'Non documentée' else 'rattachement non documenté'}{_suffixe}", "sous_titre": _sous,
                 "synthese": f"{N('PIDR n ' + _ent + _suffixe, len(_part))} agents ; provision recalculée {F('PIDR somme ' + _ent + _suffixe, sum(x['dbo'] for x in _part))} ; droits acquis {F('PIDR acq somme ' + _ent + _suffixe, sum(x['indemnite_acquise'] for x in _part))}"},
                "Agents identifiés par un numéro (année d'évaluation et rang dans la DISA 2025) ; les noms figurent dans l'onglet 14b du classeur, avec toutes les formules. Montants en FCFA. Âge calculé sur l'année de naissance."
                + (" Le rattachement de ces agents à une entité est demandé au FPPN." if _ent == "Non documentée" else ""),
                tableau={"titre": tab(f"Provision du personnel au 31/12/2025 : {_ent}{_suffixe}"), "entetes": _ENTETE_I,
                         "largeurs": _LARG_I, "taille": 8,
                         "lignes": _lignes_pidr(_part)},
                disposition="dessous",
                analyse=[f"**L'agent {_top['id']} porte la provision la plus élevée de ce tableau ({F('PIDR top ' + _ent + _suffixe, _top['dbo'])})**, avec {N('PIDR top anc ' + _ent + _suffixe, _top['anciennete'], 2)} ans d'ancienneté."],
                implication="Ces montants s'additionnent dans la diapositive « Provision du personnel par entité ». L'âge de départ, la date d'embauche et l'entité de chaque agent sont demandés au FPPN (PBC, item 113).",
                valable="agents déclarés dans la DISA 2025, évaluation au 31/12/2025 ; paramètres de l'onglet 14a",
                source=SRC_DISA + " ; onglet 14b du classeur v9")
_s0 = PSE[0]["dbo"]
_chg = V9["embauche_modifiee"]
_ag55 = [x for x in _P25 if x["age"] >= PP["age_depart"]]
contenu({"titre": "Provision du personnel : sensibilités et qualité des données", "sous_titre": "La rotation du personnel pèse plus que l'âge de départ ou le taux d'actualisation",
         "synthese": f"Sans rotation, la provision passerait de {F('PIDR sens base', _s0)} à {F('PIDR sens sans rot', next(x['dbo'] for x in PSE if 'rotation' in x['hypothese']))} ; {N('PIDR embauches modifiées', len(_chg))} dates d'embauche changent dans la DISA 2025"},
        "Chaque ligne modifie une seule hypothèse et garde les autres à leur valeur centrale (IAS 19, § 145). La seconde partie relève les incohérences des fichiers du personnel qui touchent directement le calcul.",
        tableau={"titre": tab("Sensibilité de la provision du personnel au 31/12/2025"), "entetes": ["Hypothèse", "Provision", "Écart à l'hypothèse centrale"], "largeurs": [12.0, 8.0, 8.0], "taille": 9,
                 "lignes": [[x["hypothese"], F(f"PIDR sens {x['hypothese']}", x["dbo"]), pc(f"PIDR écart {x['hypothese']}", x["dbo"] / _s0 - 1)] for x in PSE]},
        disposition="dessous",
        analyse=[f"**{N('PIDR âge 55 atteint', len(_ag55))} agents ont déjà {N('PIDR âge (t)', PP['age_depart'])} ans ou plus et sont toujours en poste en 2025, dont un de {N('PIDR âge max', max(x['age'] for x in _ag55))} ans.** [[Le FPPN n'applique donc pas l'âge de 55 ans]] : la sensibilité ci-dessus chiffre un départ à 58 et à 60 ans.",
                 f"**Pour {N('PIDR embauches (t)', len(_chg))} agents, la date d'embauche de la DISA 2025 diffère de celle des DISA 2021 à 2023**, dont {N('PIDR embauches 2013', sum(1 for c in _chg if c['dates'].get('2025') == '28/10/2013'))} remplacées par le 28/10/2013. Pour l'agent A25-01, embauché le 01/01/2006 selon les DISA 2021 à 2023, l'ancienneté retenue baisse de plus de sept ans.",
                 f"**Le plafond de la convention collective (art. 40 : 25 fois le SMIG annuel de 2 080 heures) vaut {F('PIDR plafond', V11['pidr_plafond']['plafond'])}** avec le SMIG de 75 000 FCFA par mois (décret n° 2022-986 du 21/12/2022). {N('PIDR au-dessus', V11['pidr_plafond']['agents_au_dessus'])} agents le dépassent : la provision passe à {F('PIDR plafonnée', V11['pidr_plafond']['dbo_avec_plafond'])}, toujours sous le compte 196. Le décret n° 2017-210 (art. 4 et 5), plus récent, ne prévoit pas de plafond : le calcul central ne l'applique pas."],
        implication="L'âge de départ appliqué, les dates d'embauche, le rattachement des agents recrutés en 2025 et le nombre d'enfants à charge sont demandés au FPPN (PBC, item 113). Aucune de ces données ne peut porter la provision au-dessus du compte 196.",
        valable="personnel civil de la DISA 2025, évaluation au 31/12/2025 ; plafond conventionnel chiffré à part",
        source=SRC_DISA + " ; IAS 19, § 145 ; convention collective interprofessionnelle du 19/07/1977, art. 40 ; décret n° 2022-986 du 21/12/2022 (SMIG) ; onglets 14a à 14c du classeur v9")
# ---- v7 : conclusions des tâches « engagements » et « suffisance des provisions »
# ---- v10 : conclusions « engagements » et « suffisance des provisions », revue actuarielle complète
contenu({"titre": "Conclusion : engagements envers les adhérents", "sous_titre": "L'épargne et ses intérêts dus représentent l'essentiel des engagements mesurés",
         "synthese": f"Engagements mesurés au 31/08/2026 : {F('CE total', EP_DUE26 + INT55 + DETTE_DEC26 + TARD_DEC)} (épargne avec intérêts dus, allocations décès dues et décès tardifs)"},
        "Synthèse de l'évaluation des engagements du FPPN envers ses adhérents et leurs ayants droit. Chaque ligne indique la méthode retenue, le résultat et ce qui empêche d'aller plus loin.",
        tableau={"titre": tab("Engagements envers les adhérents et ayants droit : méthode et résultat"), "entetes": ["Engagement", "Méthode retenue", "Résultat", "Ce qui manque pour conclure"], "largeurs": [5.8, 9.4, 5.4, 10.6], "taille": 8.5,
                 "lignes": [["Épargne PEP et EFC", "Solde comptable + intérêts dus du 01/01/2020 au taux du manuel", F("CE épargne", EP_DUE26 + INT55), "Taux arrêté, relevés individuels, âge de départ de chaque titulaire d'EFC"],
                            ["Allocations décès dues", "Liste du FPPN des décès non payés au barème (41591000, 41592000)", F("CE décès", DETTE_DEC26), "Date de dépôt de chaque demande, pour écarter les dossiers forclos"],
                            ["Décès non encore déclarés", "Décès d'une année enregistrés l'année suivante (états 2021 et 2023)", F("CE tardifs", TARD_DEC), "Date d'enregistrement dans les états 2022, 2024 et 2025"],
                            ["Soins faits, non encore facturés", "Factures de cliniques d'une année enregistrées l'année suivante", F("CE maladie", psap_m_elargie(2024)), "Base des prestations avec dates des soins, du dépôt et du paiement"],
                            ["Santé des retraités", "Non mesurable", "Non mesuré", "Âges et prestations des retraités"],
                            ["Personnel du FPPN (hors adhérents)", "Recalcul agent par agent (décret n° 2017-210)", F("CE PIDR", PIDR25["dbo"]), "Fichier des agents policiers détachés"]]},
        disposition="dessous",
        analyse=[f"**L'épargne représente {pc('CE part épargne', (EP_DUE26 + INT55) / (EP_DUE26 + INT55 + DETTE_DEC26 + TARD_DEC))} des engagements mesurés envers les adhérents.** Les engagements maladie et décès sont des dettes de quelques mois : le délai moyen de paiement des cliniques est de " + N('CE délai 25', DELAI[2025], 2) + " mois en 2025 (" + ref('Provision maladie : délais de paiement des cliniques') + "), et la dette décès s'élève à " + F('CE dette déc', DETTE_DEC26) + " au 31/08/2026 (" + ref('Allocations décès : dette réelle et provision obsèques') + ")."],
        implication="L'évaluation des engagements dépend d'abord de l'épargne : arrêter le taux servi et tenir des comptes individuels à jour sont les deux conditions d'un engagement démontré.",
        valable="balance provisoire au 31/08/2026 ; états de fin d'année des allocations décès 2021 à 2025 ; factures de cliniques du journal SIN 2023 à 2025 ; DISA 2021 à 2025",
        source="Source : diapositives " + ", ".join(ref(t) for t in ("Épargne : intérêts dus non comptabilisés", "Allocations décès : dette réelle et provision obsèques", "Allocations décès : décès connus après la clôture", "Provision maladie : montant nécessaire")))
contenu({"titre": "Conclusion : suffisance des provisions", "sous_titre": "Les provisions du bilan suffisent en montant, mais aucune n'est calculée et deux manquent",
         "synthese": f"Provisions maladie et obsèques au bilan : {F('CP bilan', PROV_MAL + PROV_OBS)} ; besoins mesurés : {F('CP besoin', psap_m_elargie(2024) + TARD_DEC)} hors dettes déjà comptabilisées"},
        "Pour chaque provision, le tableau donne la méthode que le FPPN peut appliquer chaque année avec ses propres fichiers, conformément à l'article 89 des statuts et à l'article 10 du règlement d'exécution n° 004/2022.",
        tableau={"titre": tab("Provisions : méthode annuelle proposée et montants"), "entetes": ["Provision", "Méthode annuelle proposée", "Au bilan", "Besoin mesuré", "Fichier du FPPN à utiliser"], "largeurs": [4.8, 10.0, 4.8, 4.8, 6.8], "taille": 8.5,
                 "lignes": [["Maladie (15801000)", "Soins de l'année facturés après la clôture, mesurés sur les factures reçues de janvier à mars", F("CP mal", PROV_MAL), F("CP mal b", psap_m_elargie(2024)), "Journal SIN, numéro et date de facture"],
                            ["Décès (15802000)", "Décès survenus non déclarés ; les décès connus restent au 41591000", F("CP obs", PROV_OBS), F("CP obs b", TARD_DEC), "État de fin d'année avec date de décès et d'enregistrement"],
                            ["Risques courts (art. 88)", "Période non couverte par une cotisation à la clôture", "Aucune", F("CP rc", 0), "Calendrier des précomptes"],
                            ["Intérêts dus sur l'épargne", "Épargne moyenne × taux arrêté, compte par compte", "Aucune", F("CP int", INT55), "Relevés individuels PEP et EFC"],
                            ["Dépréciation des prêts du Grenier", "Encours par tranche de retard × taux de la politique de crédit", "Aucune", F("CP gr", GRD["2026-08-31"]), "Balance âgée du logiciel de prêts"]]},
        disposition="dessous",
        analyse=["**Le FPPN dispose déjà des fichiers nécessaires à un calcul annuel simple de chaque provision.** Ce calcul n'est pas fait : les provisions maladie et obsèques sont inchangées depuis 2020."],
        implication="Chaque année, le service comptable établit une note de calcul qui reprend ce tableau : méthode, fichiers utilisés, montant de chaque provision et comparaison au bilan. Le directeur général la signe, et elle est présentée à l'Assemblée générale avec les comptes (statuts, art. 85 et 89). Elle démontre la suffisance des provisions à l'AIRMS et au commissaire aux comptes.",
        valable="provisions au bilan du 31/08/2026 ; besoins mesurés aux dates indiquées dans les diapositives de détail",
        source="Source : " + SRC_ST + ", art. 85, 88 et 89 ; " + REX + ", art. 10 ; diapositives de détail de la partie 9")

# ---- v11 : ajouts au format du rapport de diagnostic : méthode du FPPN et recalcul, enjeux du provisionnement, dispositif actuariel
_PP = V11["pidr_plafond"]
_GR26 = GRD.get("2026-08-31", 0.0)
_RECALC = [
    ("Provision maladie (15801000)", "Montant hérité, inchangé depuis 2020", PROV_MAL, psap_m_elargie(2024), "Soins facturés l'année suivante, journal SIN"),
    ("Provision obsèques (15802000)", "Montant hérité, inchangé depuis 2020", PROV_OBS, TARD_DEC, "Décès déclarés après la clôture (dette connue au 41591000)"),
    ("Départ à la retraite du personnel (196)", "Reprise et dotation annuelles, méthode non transmise", _PP["provision_196"], _PP["dbo_sans_plafond"], "Unités de crédit projetées, agent par agent"),
    ("Intérêts dus sur l'épargne", "Aucune écriture", 0.0, INT55, "Taux du manuel d'épargne, du 01/01/2020 au 31/08/2026"),
    ("Dépréciation des prêts du Grenier", "Aucune écriture depuis 2022", 0.0, _GR26, "Taux de la politique de crédit du Grenier"),
]


def _ecart_rel(fppn, recalc):
    return pc(f"MR rel {fppn} {recalc}", (fppn - recalc) / recalc) if recalc else "Sans objet"


contenu({"titre": "Méthode du FPPN et recalcul", "sous_titre": "Le FPPN a provisionné trop pour la maladie et le décès, et rien pour les intérêts et les prêts",
         "synthese": f"Écart total sur les cinq postes : {F('MR total', sum(f - r for _, _, f, r, _ in _RECALC))}, montant du FPPN moins recalcul"},
        "Pour chaque provision, le tableau compare le montant du FPPN au 31/08/2026 au recalcul du rapport, avec la méthode utilisée. Écart absolu = FPPN - recalcul. Écart relatif = écart absolu / recalcul.",
        tableau={"titre": tab("Montants du FPPN, recalcul et écarts"), "entetes": ["Poste", "Méthode du FPPN", "Montant du FPPN", "Recalcul", "Méthode du recalcul", "Écart absolu", "Écart relatif"],
                 "largeurs": [5.6, 5.0, 4.0, 4.0, 6.0, 4.2, 2.4], "taille": 8.5,
                 "lignes": [[p, m, F(f"MR fppn {p}", f), F(f"MR rec {p}", r), mr, F(f"MR abs {p}", f - r), _ecart_rel(f, r)] for p, m, f, r, mr in _RECALC]},
        disposition="dessous",
        analyse=["**Les provisions maladie et obsèques dépassent largement les besoins mesurés, alors que les intérêts dus et la dépréciation des prêts ne sont pas provisionnés.** Le total des provisions n'est pas insuffisant : il est mal affecté."],
        implication="Le reclassement des montants hérités vers les intérêts dus sur l'épargne suppose une décision des organes et l'information de l'AIRMS (recommandation 1).",
        valable="bilan provisoire au 31/08/2026 ; recalculs des diapositives de la partie 9 ; provision du personnel au 31/12/2025",
        source=SRC_COMPTA + " ; " + ", ".join(ref(t) for t in ("Provision maladie : montant nécessaire", "Allocations décès : décès connus après la clôture", "Provision du personnel par entité de 2021 à 2025", "Épargne : intérêts dus non comptabilisés", "Prêts du Grenier : dépréciation selon la politique de crédit")))
contenu({"titre": "Enjeux d'un provisionnement mal calibré", "sous_titre": "Une provision trop forte masque les vrais engagements, une provision trop faible surévalue les fonds propres",
         "synthese": "Au FPPN, les deux erreurs coexistent : maladie et décès surprovisionnés, épargne et prêts sous-provisionnés"},
        "Le tableau décrit l'effet de chaque erreur de provisionnement sur les comptes et les ratios, puis la situation constatée au FPPN.",
        tableau={"titre": tab("Effets d'une provision mal calibrée et situation du FPPN"), "entetes": ["Erreur", "Effet sur les comptes", "Effet sur les ratios et les décisions", "Situation au FPPN"],
                 "largeurs": [4.6, 8.4, 9.0, 9.2], "taille": 8.5,
                 "lignes": [["Provision trop forte", "Résultat et fonds propres sous-évalués. Charges futures déjà couvertes deux fois", "Sinistralité et frais mal mesurés. Cotisations jugées insuffisantes à tort", f"Maladie et obsèques : {F('EN trop', PROV_MAL + PROV_OBS - psap_m_elargie(2024) - TARD_DEC)} au-dessus des besoins mesurés"],
                            ["Provision trop faible", "Dettes futures absentes du bilan. Résultat et fonds propres surévalués", "Marge de solvabilité et réserve obligatoire surévaluées. Tarifs et placements décidés sur une base fausse", f"Intérêts dus non provisionnés : {F('EN int', INT55)}. Prêts non dépréciés : {F('EN gre', _GR26)}"],
                            ["Provision non recalculée", "Montant sans lien avec les engagements de l'année", "Impossible de démontrer la suffisance à l'AIRMS ou au commissaire aux comptes", "Montants identiques de 2020 au 31/08/2026 (statuts, art. 89, non appliqué)"]]},
        disposition="dessous",
        analyse=["**Le risque principal du FPPN n'est pas un manque global de provisions, mais leur affectation.** Les engagements d'épargne, les plus importants, sont les moins provisionnés."],
        implication="Une note de calcul annuelle, signée et présentée à l'Assemblée générale, corrige les trois erreurs à la fois (recommandation 1).",
        valable="provisions et recalculs au 31/08/2026",
        source="Source : analyses Adenka, partie 9 ; " + REX + ", art. 7 à 11 ; statuts du FPPN édités le 15/06/2023, art. 89")

separateur("9.2", "Adossement actif-passif", "Liquidité, écoulement de l'épargne, tests de résistance, gouvernance")
contenu({"titre": "Structure de l'actif", "sous_titre": "L'actif repose sur des créances et des investissements peu liquides",
         "synthese": f"Au 31/08/2026 : CMP {F('CMP (a)', b[2026]['cmp_investissements'])}, terrains {F('Terrains', b[2026]['terrains'])}, trésorerie {F('Trésorerie', b[2026]['tresorerie_totale'])}"},
        "Les montants « CMP » regroupent les comptes 47111101 (« Polyclinique du FPPN »), 47111104 (« Investissements au CMP ») et 47111105 (« Charges de fonctionnement au CMP »)  ; CMP désigne la Clinique médicale du policier, autorisée par l'arrêté n° 0027/MSHP/DGS/DEPS du 10/02/2020. La créance sur l'Hôpital de la Police (47111100) est inchangée depuis 2020.",
        graphique={"titre": fig("Principaux postes de l'actif"), "type": "barres_empilees", "categories": yrs_bil,
                   "series": [{"nom": "Investissements CMP", "valeurs": [round(b[a]["cmp_investissements"]) for a in range(2020, 2027)], "couleur": BLEU},
                              {"nom": "Terrains", "valeurs": [round(b[a]["terrains"]) for a in range(2020, 2027)], "couleur": GRIS_F},
                              {"nom": "Créance Hôpital de la Police", "valeurs": [round(b[a]["creance_hopital_police"]) for a in range(2020, 2027)], "couleur": GRIS_C},
                              {"nom": "Financement du Grenier", "valeurs": [round(b[a]["grenier_financement"]) for a in range(2020, 2027)], "couleur": ORANGE},
                              {"nom": "Trésorerie", "valeurs": [round(b[a]["tresorerie_totale"]) for a in range(2020, 2027)], "couleur": NOIR}],
                   "format": FMT_M, "etiquettes": False, "legende": "droite", "unite": "En millions de FCFA"},
        analyse=[f"**Les montants « CMP » passent de {F('CMP 2020', b[2020]['cmp_investissements'])} en 2020 à {F('CMP 2026', b[2026]['cmp_investissements'])} au 31/08/2026.**",
                 f"**La créance sur l'Hôpital de la Police ({F('Hôpital', b[2026]['creance_hopital_police'])}) n'a pas bougé depuis 2020.** Aucun remboursement n'apparaît.",
                 f"**Les dettes envers les cliniques atteignent {F('Fournisseurs cliniques', b[2026]['fournisseurs_cliniques'])} au 31/08/2026.**"],
        implication="Le FPPN finance aussi des charges de fonctionnement « au CMP ». Le règlement d'exécution n° 004/2022 interdit de couvrir les pertes d'activités sanitaires par les cotisations ou les réserves obligatoires (art. 27) : une comptabilité séparée de ces activités doit le démontrer.",
        source=SRC_COMPTA)
mo = {m["cloture"]: m for m in MOB}
m26 = mo["31/08/2026"]
contenu({"titre": "Adossement de liquidité provisoire", "sous_titre": "L'actif mobilisable en moins d'un an couvre la moitié du passif à moins d'un an",
         "synthese": f"Au 31/08/2026 : actif mobilisable en moins d'un an {F('Actif < 1 an 2026', m26['immediat'] + m26['moins_1_an'])}, passif à moins d'un an hors épargne {F('Passif < 1 an 2026', m26['passif_moins_1_an'])}"},
        "Étapes 1 à 3 de la méthode d'adossement du dossier de préparation : classer l'actif par délai de mobilisation, classer le passif par échéance, comparer tranche par tranche. Actif immédiat : caisse et comptes à vue. Moins d'un an : dépôts à terme, fonds commun de placement, créances à court terme. 1 à 5 ans : prêts PEP et financement du Grenier (durée médiane des prêts de 10 mois). Plus de 5 ans ou incertain : immobilisations, terrains, avances au CMP, créance sur l'Hôpital de la Police. Passif à moins d'un an : dettes à court terme hors épargne et provisions pour prestations à payer. L'épargne est présentée à part : son échéance dépend des âges et des conditions de retrait.",
        graphique={"titre": fig("Actif par délai de mobilisation et passif par échéance"), "type": "barres_empilees",
                   "categories": ["Actif 31/12/2025", "Passif 31/12/2025", "Actif 31/08/2026", "Passif 31/08/2026"],
                   "series": [{"nom": "Actif immédiat", "valeurs": [round(mo["31/12/2025"]["immediat"]), None, round(m26["immediat"]), None], "couleur": NOIR},
                              {"nom": "Actif à moins d'un an", "valeurs": [round(mo["31/12/2025"]["moins_1_an"]), None, round(m26["moins_1_an"]), None], "couleur": GRIS_F},
                              {"nom": "Actif à 1 à 5 ans", "valeurs": [round(mo["31/12/2025"]["1_a_5_ans"]), None, round(m26["1_a_5_ans"]), None], "couleur": ORANGE},
                              {"nom": "Actif à plus de 5 ans ou incertain", "valeurs": [round(mo["31/12/2025"]["plus_5_ans_ou_incertain"]), None, round(m26["plus_5_ans_ou_incertain"]), None], "couleur": GRIS_C},
                              {"nom": "Passif à moins d'un an", "valeurs": [None, round(mo["31/12/2025"]["passif_moins_1_an"]), None, round(m26["passif_moins_1_an"])], "couleur": BLEU},
                              {"nom": "Épargne, échéance à établir", "valeurs": [None, round(mo["31/12/2025"]["epargne_echeance_a_etablir"]), None, round(m26["epargne_echeance_a_etablir"])], "couleur": ROUGE}],
                   "format": FMT_M, "etiquettes": False, "legende": "droite", "unite": "En millions de FCFA"},
        analyse=[f"**Plus des trois quarts de l'actif sont à plus de cinq ans ou de délai incertain.** Cette tranche vaut {F('Actif > 5 ans 2026', m26['plus_5_ans_ou_incertain'])} sur {F('Actif total mobilisation 2026', m26['total_actif'])} au 31/08/2026.",
                 f"**L'actif à moins d'un an ne couvre pas le passif à moins d'un an.** L'écart vaut {F('Écart < 1 an 2026', m26['passif_moins_1_an'] - m26['immediat'] - m26['moins_1_an'])} au 31/08/2026, avant même toute sortie d'épargne.",
                 f"**L'épargne ({F('Épargne mobilisation 2026', m26['epargne_echeance_a_etablir'])}) n'a pas d'échéancier connu.** Chaque départ à la retraite déclenche pourtant le versement de l'EFC."],
        implication="Cet adossement est une borne favorable (diapositive suivante) : le passif à moins d'un an comprend les provisions pour prestations à payer, dont le montant n'est pas recalculé. Le FPPN finance des engagements courts avec des actifs longs : c'est le risque de liquidité central, chiffré dans les écarts année par année.",
        source=SRC_COMPTA + " ; dossier de préparation de la mission, § 3.1")
# ---- v7 : adossement actif-passif, borne favorable, règles de valorisation, scénarios, sorties d'épargne
contenu({"titre": "Pourquoi l'adossement est une borne favorable", "sous_titre": "Toutes les corrections attendues augmentent le passif ou réduisent l'actif",
         "synthese": "L'écart mesuré entre actifs et engagements est donc le meilleur cas possible : les données manquantes ne peuvent que l'aggraver"},
        "L'adossement compare ce que le FPPN possède à ce qu'il doit, selon les échéances. Il ne vaut que si les deux côtés sont bien mesurés. Ici, le passif est incomplet (engagements non provisionnés) et une partie de l'actif est d'une valeur non démontrée. Le tableau recense chaque correction attendue et son sens.",
        tableau={"titre": tab("Corrections attendues et sens de leur effet sur l'adossement"), "entetes": ["Côté", "Élément", "Situation au 31/08/2026", "Sens de la correction", "Texte de référence"], "largeurs": [2.2, 6.2, 9.4, 5.0, 8.4], "taille": 9,
                 "lignes": [["Passif", "Provisions maladie et décès", "Montants hérités d'environ 2004, jamais recalculés ; aucune provision maladie calculée", "Baisse : besoins mesurés très inférieurs au bilan (partie 9)", REX + ", art. 10"],
                            ["Passif", "Provision mathématique (EFC, santé des retraités)", "Absente du bilan", "Hausse si un capital ou une garantie à vie est promis", REX + ", art. 11"],
                            ["Passif", "Épargne PEP et EFC", F("Épargne ALM 2026", mo["31/08/2026"]["epargne_echeance_a_etablir"]) + ", sans échéancier ; taux servi non établi (3,5 %, 5,5 % ou 14 % selon le document)", "Hausse des flux futurs (intérêts servis, non comptabilisés)", "Manuel des produits d'épargne (taux de 3,5 % et 5,5 %)"],
                            ["Actif", "Créance sur l'hôpital de la Police", F("Hôpital ALM 2026", EXA["31/08/2026"]["hop"]) + ", inchangée depuis 2020", "Baisse (recouvrement non démontré)", "PCMS (dépréciation des créances) ; " + REX + ", art. 20"],
                            ["Actif", "Clinique du FPPN (CMP)", F("CMP inv ALM 2026", EXA["31/08/2026"]["cmpi"]) + " d'investissements et " + F("CMP ch ALM 2026", EXA["31/08/2026"]["cmpc"]) + " de charges en comptes 47", "Baisse (charges consommées, valeur de marché non établie)", REX + ", art. 20 et 27"],
                            ["Actif", "Terrains", F("Terrains ALM 2026", EXA["31/08/2026"]["terr"]) + ", sans évaluation récente", "Sens inconnu sans évaluation", REX + ", art. 20"],
                            ["Actif", "Prêts du Grenier", "Aucune dépréciation des prêts en retard", "Baisse (faible montant)", "PCMS ; politique de crédit du Grenier"]]},
        disposition="dessous",
        analyse=["**Les intérêts dus ajoutent du passif, les actifs non évalués peuvent réduire l'actif.** Seule la baisse des provisions maladie et décès allège le passif : elle est trop faible pour changer la conclusion (manque cumulé de " + F("AB gap", -GAP[4][5]) + " à cinq ans)."],
        implication="Le FPPN n'est pas adossé, même dans l'hypothèse la plus favorable. L'écart est chiffré : manque cumulé de " + F("AB gap2", -GAP[4][5]) + " à cinq ans (" + ref("Adossement : écarts de liquidité année par année") + ") et couverture de " + pc("AB couv", couverture("31/08/2026", T_MOY)[2]) + " des paiements de l'année dans le scénario central (" + ref("Adossement : tests de résistance") + ").",
        source=SRC_TXT + " ; " + SRC_COMPTA[9:])
contenu({"titre": "Règles de valorisation des actifs", "sous_titre": "Les textes fixent quels actifs peuvent couvrir les engagements, et à quelle valeur",
         "synthese": "Chaque scénario applique une règle tirée d'un texte, et non un taux de décote choisi par Adenka"},
        "Le règlement d'exécution n° 004/2022 impose d'évaluer les actifs à leur valeur de marché (art. 20) et de détenir des actifs sûrs, diversifiés, liquides et rentables (art. 17 et 21). La Décision n° 05/2022 rapporte les « actifs admis » aux engagements réglementés (section II, § 2.1). L'article 27 interdit de couvrir les pertes des activités sanitaires par les cotisations ou les réserves. Le PCMS impose la dépréciation d'un actif dont la valeur actuelle est inférieure à sa valeur d'entrée, et l'évaluation à la valeur liquidative si la continuité de l'activité n'est pas assurée.",
        tableau={"titre": tab("Règles de valorisation retenues dans chaque scénario"), "entetes": ["Scénario", "Règle appliquée", "Actifs retirés", "Fondement"], "largeurs": [4.6, 9.6, 8.8, 8.2], "taille": 9,
                 "lignes": [["1. Lecture comptable", "Valeurs des balances, continuité de l'activité", "Aucun", "PCMS (coût historique, continuité de l'exploitation)"],
                            ["2. Actifs sans valeur réalisable retirés", "Une dépense consommée ou une créance non recouvrée depuis 2020 n'a pas de valeur de marché démontrée", "Créance sur l'hôpital ; charges de fonctionnement de la clinique (47111105)", REX + ", art. 20 et 27 ; PCMS (dépréciation des créances)"],
                            ["3. Actifs sans valeur de marché établie retirés", "Un actif ne couvre les engagements que si sa valeur de marché est établie", "Scénario 2, plus les investissements de la clinique (47111104) et les terrains", REX + ", art. 20 ; " + DEC + ", section II, § 2.1 (actifs admis)"],
                            ["4. Actifs liquides à un an", "Seuls les actifs mobilisables en moins d'un an paient les sorties de l'année", "Tout actif à plus d'un an", "PCMS (valeur liquidative) ; " + REX + ", art. 19 et 21 (liquidité)"]]},
        disposition="dessous",
        analyse=["**Les scénarios 2 et 3 retirent un actif tant que sa valeur n'est pas démontrée : ce ne sont pas des pertes certaines.** Un rapport d'évaluation des terrains ou de la clinique, ou un accord de recouvrement avec l'État, permettrait de les réintégrer.",
                 "**Les prêts du Grenier restent dans tous les scénarios sauf le quatrième.** Leur dépréciation sera calculée avec le barème de la BCEAO des SFD, dès réception du texte ; les montants en retard sont faibles (partie 9)."],
        implication="La décision ministérielle qui fixe, en Côte d'Ivoire, les conditions de placement et la liste des actifs admis (art. 17) n'a pas été reçue (PBC, item 54). Elle permettra de remplacer ces règles générales par la liste officielle.",
        source=SRC_TXT)
contenu({"titre": "Scénarios de couverture des engagements", "sous_titre": "Même en lecture comptable, l'actif ne couvre pas tous les engagements connus",
         "synthese": f"Au 31/08/2026, couverture des engagements connus : {pc('Scénario 1 2026', SCEN['31/08/2026'][0] / ENG['31/08/2026'])} en lecture comptable, {pc('Scénario 3 2026', SCEN['31/08/2026'][2] / ENG['31/08/2026'])} sans les actifs de valeur non établie"},
        "Couverture = actif retenu dans le scénario / engagements connus. Engagements connus = passif à moins d'un an (dettes et provisions au bilan) + épargne due aux mutualistes, sans les engagements non provisionnés. Le scénario 4 compare l'actif à moins d'un an au seul passif à moins d'un an.",
        tableau={"titre": tab("Couverture des engagements connus selon le scénario"), "entetes": ["Scénario", "Actif retenu au 31/12/2025", "Couverture 2025", "Actif retenu au 31/08/2026", "Couverture 2026"], "largeurs": [7.6, 6.2, 3.6, 6.2, 3.6], "taille": 10,
                 "lignes": [[nom, F(f"Actif sc{i} 2025", SCEN["31/12/2025"][i - 1]), pc(f"Couv sc{i} 2025", SCEN["31/12/2025"][i - 1] / (ENG["31/12/2025"] if i < 4 else mo["31/12/2025"]["passif_moins_1_an"])),
                             F(f"Actif sc{i} 2026", SCEN["31/08/2026"][i - 1]), pc(f"Couv sc{i} 2026", SCEN["31/08/2026"][i - 1] / (ENG["31/08/2026"] if i < 4 else mo["31/08/2026"]["passif_moins_1_an"]))]
                            for i, nom in ((1, "1. Lecture comptable"), (2, "2. Sans actifs de valeur non réalisable"), (3, "3. Sans actifs de valeur de marché non établie"), (4, "4. Actifs à un an / passif à un an"))]
                          + [["Engagements connus (scénarios 1 à 3)", F("Engagements 2025", ENG["31/12/2025"]), "", F("Engagements 2026", ENG["31/08/2026"]), ""]]},
        disposition="dessous",
        analyse=[f"**Même en lecture comptable, l'actif ne couvre que {pc('Scénario 1 2026 (b)', SCEN['31/08/2026'][0] / ENG['31/08/2026'])} des engagements connus au 31/08/2026.** C'est la traduction des fonds propres négatifs.",
                 f"**Sans les actifs dont la valeur n'est pas établie, la couverture tombe à {pc('Scénario 3 2026 (b)', SCEN['31/08/2026'][2] / ENG['31/08/2026'])}.** L'hôpital, la clinique et les terrains portent l'essentiel de l'actif.",
                 f"**À un an, l'actif liquide couvre {pc('Scénario 4 2026 (b)', SCEN['31/08/2026'][3] / mo['31/08/2026']['passif_moins_1_an'])} du passif à un an**, avant toute sortie d'épargne."],
        implication="Ces couvertures sont des maxima : les engagements non provisionnés (provisions recalculées, provisions mathématiques) ne sont pas encore au dénominateur. Le FPPN doit faire évaluer ses actifs immobiliers et sa clinique, et documenter le recouvrement de la créance sur l'hôpital.",
        source=SRC_TXT + " ; " + SRC_COMPTA[9:])
# ---- v10 : adossement actif-passif (méthodes justifiées, écoulement de l'épargne, écarts de liquidité, résistance, gouvernance)
SRC_ALM = {"ONT": "« ALM_ONT.pdf »", "BAN": "« Gestion des Risques Bancaires(Actif PAssif).pdf » (Sciences-Po, 2009-2010)", "IAA": "« Ch13_Gestion_Actif-Passif_French.pdf » (Association actuarielle internationale, 2016)",
           "TIC": "« Construction d'un modèle ALM.pdf »", "LAV": "« 2024081723_jerome-lavoie-pilote.pdf » (Institut des actuaires)"}
contenu({"titre": "Choix des méthodes d'adossement", "sous_titre": "Des méthodes de liquidité, adaptées à un passif court et à une épargne sans échéancier",
         "synthese": "Chaque méthode est retenue parce que ses conditions d'emploi sont remplies par les données du FPPN ; les méthodes écartées sont démontrées en annexe"},
        "La gestion actif-passif vérifie que les actifs produisent l'argent nécessaire au moment où les engagements doivent être payés. Le choix de la méthode dépend du profil du passif et des données disponibles.",
        tableau={"titre": tab("Méthodes d'adossement retenues et justification"), "entetes": ["Question", "Méthode retenue", "Pourquoi (document de référence, page)", "Donnée du FPPN qui le permet"], "largeurs": [5.0, 6.8, 11.0, 8.4], "taille": 8,
                 "lignes": [["Quel est l'enjeu pour la maladie et le décès ?", "Liquidité à court terme", "Contrats santé de durée courte : complexité actif-passif « faible », l'enjeu est la gestion du portefeuille d'actifs (" + SRC_ALM["ONT"] + ", p. 8)", f"Délai moyen de paiement des cliniques de {N('ALM délai', DELAI[2025], 2)} mois en 2025"],
                            ["Quand l'épargne sortira-t-elle ?", "Écoulement conventionnel", "Les comptes d'épargne sans échéance contractuelle ont un amortissement « à modéliser » (" + SRC_ALM["BAN"] + ", p. 33)", "Sorties annuelles des comptes 41195100 et 41195200 de 2021 à 2025"],
                            ["L'actif suit-il les sorties ?", "Écart de liquidité par échéance", "Écart entre passifs et actifs au cours du temps après écoulement des éléments du bilan (" + SRC_ALM["BAN"] + ", p. 86)", "Actifs classés par délai de mobilisation aux bilans 2025 et 2026"],
                            ["Quels actifs sont mobilisables ?", "Classes de liquidité", "Les titres sont classés selon leur degré de liquidité (" + SRC_ALM["BAN"] + ", p. 95)", "Nature comptable de chaque actif (banques, DAT, FCP, prêts, immeubles)"],
                            ["Que se passe-t-il en cas de choc ?", "Tests de résistance calibrés sur le FPPN", "Le rachat massif de 40 % de la formule standard européenne (" + SRC_ALM["TIC"] + ", p. 110) n'est pas calibré sur le FPPN", f"Taux de sortie maximal observé : {pc('ALM tmax', T_MAX)} ; hausse maximale des prestations maladie de 2024 et 2025 : {pc('ALM pic', PIC)}"],
                            ["Qui décide ?", "Gouvernance actif-passif", "Le conseil d'administration fixe les limites, un comité suit les écarts (" + SRC_ALM["IAA"] + ", p. 15 et 19 ; " + SRC_ALM["LAV"] + ", p. 8 et 11)", "Décision de placement par le seul directeur général (entretien du 19/10/2026)"]]},
        disposition="dessous",
        analyse=["**Les méthodes de taux (duration, immunisation, modèles stochastiques, VaR) ne sont pas retenues** : elles supposent des flux de passif connus et un portefeuille obligataire, absents au FPPN (annexe « Méthodes non retenues et pourquoi »)."],
        implication="Les résultats de l'adossement sont des ordres de grandeur fondés sur le comportement observé des mutualistes ; ils deviendront des flux contractuels dès que l'épargne aura un échéancier.",
        valable="choix de méthode fondé sur les documents de référence cités et sur les données du FPPN de 2020 au 31/08/2026",
        source="Source : " + " ; ".join(SRC_ALM.values()))
_EP10 = {e["exercice"]: e for e in V10["epargne"]}
contenu({"titre": "Adossement : à quel rythme l'épargne sort-elle ?", "sous_titre": f"De 2021 à 2025, il sort chaque année en moyenne {pc('ALM t moy', T_MOY)} de l'épargne du début d'année",
         "synthese": f"Taux de sortie compris entre {pc('ALM tmin', min(x['total_taux'] for x in SORT10))} et {pc('ALM tmax (s)', T_MAX)} ; le PEP sort plus vite que l'EFC"},
        "Taux de sortie de l'année = sorties de l'année (débits des comptes 41195100 et 41195200) / épargne au début de l'année. C'est l'écoulement conventionnel : faute d'échéancier, le comportement passé des mutualistes sert à prévoir les sorties futures.",
        tableau={"titre": tab("Sorties annuelles de l'épargne et taux de sortie"), "entetes": ["Exercice", "Épargne au début", "Sorties de l'année", "Taux PEP", "Taux EFC", "Taux total"], "largeurs": [3.0, 5.4, 5.4, 3.6, 3.6, 3.6], "taille": 9,
                 "lignes": [[str(x["exercice"]), F(f"ALM début {x['exercice']}", _EP10[x["exercice"]]["pep_debut"] + _EP10[x["exercice"]]["efc_debut"]),
                             F(f"ALM sorties {x['exercice']}", _EP10[x["exercice"]]["pep_sorties"] + _EP10[x["exercice"]]["efc_sorties"]),
                             pc(f"ALM tp {x['exercice']}", x["pep_taux"]), pc(f"ALM te {x['exercice']}", x["efc_taux"]), pc(f"ALM tt {x['exercice']}", x["total_taux"])] for x in SORT10]},
        disposition="droite",
        analyse=[f"**Le taux moyen ({pc('ALM t moy (a)', T_MOY)}) sert au scénario central, le taux maximal ({pc('ALM tmax (a)', T_MAX)}, en {max(SORT10, key=lambda x: x['total_taux'])['exercice']}) au test de résistance.**",
                 "**Les sorties de l'EFC suivent les départs à la retraite**, que seules les dates de naissance permettraient de prévoir."],
        implication="Les sorties d'épargne d'une année représentent plusieurs milliards de FCFA : c'est le premier besoin de liquidité du FPPN, devant les prestations.",
        valable="débits et soldes des comptes 41195100 et 41195200, balances 2020 à 2025 ; l'année 2020 sert de point de départ ; 2026 non retenue (huit mois)",
        source=SRC_COMPTA + " ; " + SRC_ALM["BAN"] + ", p. 33 à 36 ; onglet 24a du classeur des calculs")
contenu({"titre": "Adossement : écarts de liquidité année par année", "sous_titre": f"Le manque de liquidité cumulé atteint {F('GAP an5', -GAP[4][5])} à la fin de la cinquième année",
         "synthese": f"Même en comptant à leur valeur comptable les terrains, la clinique et les créances longues, il manque {F('GAP fin', -GAP[5][5])}"},
        "Écoulement au 31/08/2026, sans nouveaux versements (écart statique). Sorties d'épargne de l'année k = épargne × (1 - taux moyen) puissance (k - 1) × taux moyen. Actifs à moins d'un an en année 1 ; actifs de 1 à 5 ans répartis à parts égales sur les années 2 à 5, faute d'échéancier. Écart = actifs disponibles - paiements.",
        tableau={"titre": tab("Écarts de liquidité au 31/08/2026, scénario central"), "entetes": ["Échéance", "Actifs disponibles", "Dettes à payer", "Sorties d'épargne", "Écart de l'année", "Écart cumulé"], "largeurs": [3.8, 5.4, 5.0, 5.0, 5.4, 5.4], "taille": 9,
                 "lignes": [[g[0], F(f"GAP act {g[0]}", g[1]), F(f"GAP det {g[0]}", g[2]), F(f"GAP sor {g[0]}", g[3]), F(f"GAP éc {g[0]}", g[4]), F(f"GAP cum {g[0]}", g[5])] for g in GAP]},
        disposition="dessous",
        analyse=[f"**Dès la première année, les paiements dépassent les actifs disponibles de {F('GAP an1', -GAP[0][4])}.** L'écart se creuse chaque année : les actifs de 1 à 5 ans (prêts du Grenier, créances) sont trop faibles pour suivre les sorties.",
                 f"**Même en ajoutant les {F('GAP 5+', _m['plus_5_ans_ou_incertain'])} d'actifs longs ou incertains (terrains, clinique, créance sur l'hôpital) à leur valeur comptable, l'écart cumulé reste négatif de {F('GAP fin (a)', -GAP[5][5])}.** Leur valeur de vente n'est d'ailleurs pas établie."],
        implication="En pratique, l'écart est comblé chaque année par les nouvelles cotisations d'épargne et les excédents : le FPPN dépend de la poursuite des versements pour payer les sorties.",
        valable="bilan provisoire au 31/08/2026, actifs classés par nature comptable ; taux de sortie moyen 2021 à 2025 ; sans nouveaux versements ni intérêts dus",
        source=SRC_COMPTA + " ; " + SRC_ALM["BAN"] + ", p. 86 et 87 ; onglet 24b du classeur des calculs")
_SC = [("Central", T_MOY, 0.0, "Taux de sortie moyen 2021 à 2025"), ("Résistance", T_MAX, 0.0, "Taux de sortie maximal observé"),
       ("Résistance sévère", T_MAX, CHOC_MAL, f"Taux maximal et hausse des prestations maladie de {pc('ALM pic (sc)', PIC)}")]
_lig_sc = []
for cl in ("31/12/2025", "31/08/2026"):
    for nom, t, ch, desc in _SC:
        a_, b_, c_ = couverture(cl, t, ch)
        _lig_sc.append([cl, nom, desc, F(f"SC act {cl} {nom}", a_), F(f"SC bes {cl} {nom}", b_), F(f"SC éc {cl} {nom}", a_ - b_), pc(f"SC couv {cl} {nom}", c_)])
contenu({"titre": "Adossement : tests de résistance", "sous_titre": f"Même dans le scénario central, l'actif à moins d'un an couvre {pc('SC c26', couverture('31/08/2026', T_MOY)[2])} des paiements de l'année au 31/08/2026",
         "synthese": "Les chocs sont calibrés sur les extrêmes observés au FPPN, pour rester cohérents avec le reste du rapport"},
        "Besoin de l'année = dettes à moins d'un an + épargne × taux de sortie + choc de prestations. Couverture = actif à moins d'un an / besoin. Choc de prestations = prestations maladie 2025 × hausse annuelle la plus forte de 2024 et 2025 ; les hausses de 2021 à 2023 sont écartées car elles reflètent le rattrapage des factures impayées.",
        tableau={"titre": tab("Couverture des paiements de l'année selon trois scénarios"), "entetes": ["Clôture", "Scénario", "Hypothèse", "Actif à moins d'un an", "Besoin de l'année", "Écart", "Couverture"], "largeurs": [3.0, 3.6, 7.4, 4.4, 4.4, 4.4, 3.0], "taille": 8.5,
                 "lignes": _lig_sc},
        disposition="dessous",
        analyse=[f"**La couverture varie entre {pc('SC min', min(couverture(c, T_MAX, CHOC_MAL)[2] for c in ('31/12/2025', '31/08/2026')))} et {pc('SC max', max(couverture(c, T_MOY)[2] for c in ('31/12/2025', '31/08/2026')))} selon le scénario et la clôture.** Aucun scénario n'approche 100 %."],
        implication="Une réserve de liquidité égale aux sorties d'épargne d'une année au taux maximal observé est la cible minimale d'une politique actif-passif.",
        valable="bilans au 31/12/2025 et au 31/08/2026 (provisoire) ; taux de sortie 2021 à 2025 ; prestations maladie 2023 à 2025",
        source=SRC_COMPTA + " ; onglet 24b du classeur des calculs")
contenu({"titre": "Adossement : gouvernance à mettre en place", "sous_titre": "Aucune des pratiques de gouvernance actif-passif de référence n'est en place au FPPN",
         "synthese": "Le conseil d'administration doit fixer les limites ; un comité doit suivre les écarts et en rendre compte"},
        "Les pratiques de référence sont celles de l'Association actuarielle internationale et de l'Institut des actuaires. Elles sont rapprochées de la situation constatée au FPPN.",
        tableau={"titre": tab("Gouvernance actif-passif : pratiques de référence et situation du FPPN"), "entetes": ["Pratique de référence", "Source", "Situation du FPPN", "Action proposée"], "largeurs": [7.6, 5.4, 9.0, 9.2], "taille": 8.5,
                 "lignes": [["Le conseil d'administration s'engage et approuve les limites de risque", "AAI, ch. 13, p. 15", "Aucune limite votée ; le directeur général décide seul des placements", "Faire voter une politique actif-passif et de placement"],
                            ["Un comité actif-passif, avec un mandat écrit", "AAI, ch. 13, p. 19 ; Lavoie, p. 11", "Pas de comité", "Créer un comité : direction, comptabilité, actuaire"],
                            ["Une politique qui encadre la gestion actif-passif", "Lavoie, p. 8", "Pas de politique écrite", "Fixer réserve de liquidité, actifs admis, rendement minimal"],
                            ["Un suivi régulier des écarts", "Gestion des risques bancaires, p. 86", "Pas d'échéancier de l'épargne, pas de tableau de suivi", "Écarts de liquidité trimestriels, présentés au comité"],
                            ["Des engagements connus", "AAI, ch. 13, p. 3", "Taux servi non arrêté, comptes individuels non transmis", "Arrêter le taux, tenir les comptes individuels"]]},
        disposition="dessous",
        analyse=["**L'absence de gouvernance explique les constats chiffrés** : sans limite ni suivi, l'épargne a financé des actifs longs sans que personne ne mesure l'effet sur la liquidité."],
        implication="La gouvernance est le levier le moins coûteux : elle ne demande ni vente d'actif ni hausse de cotisation.",
        valable="pratiques des documents de référence cités ; situation du FPPN constatée en entretien le 19/10/2026 et sur les documents reçus jusqu'au 22/10/2026",
        source="Source : " + SRC_ALM["IAA"] + " ; " + SRC_ALM["LAV"] + " ; " + SRC_ALM["BAN"] + " ; " + SRC_ENT14[9:])
contenu({"titre": "Conclusion : adossement actif-passif", "sous_titre": "Le FPPN n'est pas adossé : l'épargne exigible finance des actifs longs et peu liquides",
         "synthese": f"Couverture des paiements de l'année au 31/08/2026 : {pc('CA c', couverture('31/08/2026', T_MOY)[2])} (central) à {pc('CA s', couverture('31/08/2026', T_MAX, CHOC_MAL)[2])} (sévère)"},
        "Synthèse de l'adossement actif-passif : ce qui est démontré, avec quelle méthode, et ce qui permettrait de passer à un adossement en flux contractuels.",
        tableau={"titre": tab("Adossement actif-passif : réponse et niveau de preuve"), "entetes": ["Question", "Réponse", "Méthode", "Pour aller plus loin"], "largeurs": [6.0, 10.4, 5.6, 9.2], "taille": 8.5,
                 "lignes": [["La liquidité couvre-t-elle l'année qui vient ?", "Non : " + pc("CA c (t)", couverture("31/08/2026", T_MOY)[2]) + " dans le scénario central", "Écoulement conventionnel", "Échéancier de l'épargne"],
                            ["L'écart se résorbe-t-il dans les cinq ans ?", "Non : manque cumulé de " + F("CA gap", -GAP[4][5]) + " à cinq ans", "Écarts par année", "Classement contractuel des actifs"],
                            ["Le FPPN résiste-t-il à un choc ?", "Non : " + pc("CA s (t)", couverture("31/08/2026", T_MAX, CHOC_MAL)[2]) + " dans le scénario sévère", "Tests de résistance", "Données de sorties par motif"],
                            ["Le FPPN pilote-t-il son adossement ?", "Non : ni politique, ni comité, ni limites", "Pratiques de référence", "Décision du conseil d'administration"]]},
        disposition="dessous",
        analyse=["**Le non-adossement est démontré avec les données du FPPN**, quel que soit le scénario retenu. Un échéancier individuel de l'épargne remplacerait le taux de sortie observé, sans changer la conclusion."],
        implication="Le plan de redressement doit traiter en priorité la liquidité de l'épargne : réserve de liquidité, arrêt du financement d'actifs longs par l'épargne, politique actif-passif votée.",
        valable="bilans au 31/12/2025 et au 31/08/2026 (provisoire) ; comportements de sortie observés de 2021 à 2025",
        source="Source : diapositives " + ", ".join(ref(t) for t in ("Adossement : écarts de liquidité année par année", "Adossement : tests de résistance", "Adossement : gouvernance à mettre en place")))

separateur("9.3", "Fonds propres et placements", "Fonds propres, placements, actifs immobilisés, adéquation aux engagements")

fp = EX["fonds_propres_2025"]
contenu({"titre": "Calcul pas à pas : fonds propres 2025", "sous_titre": "Les pertes reportées dépassent les réserves, d'où des fonds propres négatifs",
         "synthese": f"Fonds propres 2025 = {F('FP hors résultat 2025', b[2025]['fonds_propres_hors_resultat'])} + résultat {F('Résultat 2025 (fp)', b[2025]['resultat_exercice'])} = {F('FP 2025 (pas)', b[2025]['fonds_propres'])}"},
        "Les comptes de capitaux propres ont un solde créditeur (positif pour le FPPN), sauf le report à nouveau débiteur (pertes passées), affiché ici en négatif. Libellés exacts de la balance 2025.",
        pas_a_pas={"entetes": ["Compte", "Intitulé (balance 2025)", "Montant"], "lignes": [[c[0], c[1], nb(-float(c[2])) + " FCFA"] for c in fp["comptes"]] + [["", "Sous-total capitaux propres", F("Sous-total FP", b[2025]["fonds_propres_hors_resultat"])], ["", "Résultat de l'exercice 2025", F("Résultat (pas FP)", b[2025]["resultat_exercice"])], ["", "Fonds propres", F("FP (pas total)", b[2025]["fonds_propres"])]],
                   "largeurs": [3.4, 11.4, 6.6], "rouge": [(3, 2), (6, 2)],
                   "explications": ["Les pertes reportées (compte 129) dépassent l'ensemble des réserves.", "Les fonds propres restent négatifs malgré le résultat positif de 2025.", "Les excédents futurs devront absorber ce solde négatif."],
                   "fichier": "AUDIT ADENKA\\DAFC\\2025\\BALANCE FPPN 2025.xlsx"},
        analyse=["**Les réserves existent, mais les pertes passées les dépassent.** Le compte 10620000, intitulé « Réévaluation matériels » de 2020 à 2025, est le plus important des comptes positifs ; le chef comptable l'explique par la réévaluation du terrain du siège (partie 10)."],
        implication="La marge de solvabilité de la Décision n° 05/2022 se calcule sur ces réserves : selon que les pertes reportées sont déduites ou non, le FPPN est conforme ou non (partie 10).",
        source="Source : balance des comptes 2025")
contenu({"titre": "Fonds propres", "sous_titre": "Les fonds propres restent négatifs malgré six années d'excédents",
         "synthese": f"{F('FP 2020 (f)', b[2020]['fonds_propres'])} fin 2020, {F('FP 2025 (f)', b[2025]['fonds_propres'])} fin 2025, {F('FP 2026 (f)', b[2026]['fonds_propres'])} au 31/08/2026"},
        "Des fonds propres négatifs signifient que les dettes dépassent les actifs comptables. Les ratios de fonds propres et de marge de solvabilité de la Décision n° 05/2022 sont calculés dans la partie 10.",
        graphique={"titre": fig("Fonds propres et résultat de l'exercice"), "type": "barres", "categories": yrs_bil,
                   "series": [{"nom": "Fonds propres", "valeurs": [round(b[a]["fonds_propres"]) for a in range(2020, 2027)], "couleur": BLEU},
                              {"nom": "Résultat de l'exercice", "valeurs": [round(b[a]["resultat_exercice"]) for a in range(2020, 2027)], "couleur": ORANGE}],
                   "format": FMT_M, "unite": "En millions de FCFA"},
        analyse=["**Les excédents de 2020 à 2025 réduisent l'écart sans le combler.** Les fonds propres passent de " + F("FP 2020 (t)", b[2020]["fonds_propres"]) + " à " + F("FP 2025 (t)", b[2025]["fonds_propres"]) + ".",
                 "**Le déficit des huit premiers mois de 2026 inverse la tendance.** Les fonds propres reculent à nouveau au 31/08/2026.",
                 f"**Le ratio des fonds propres et assimilés vaut {val_ratio(RT['FEP'], 2025)} en 2025, pour une norme d'au moins 1.**"],
        implication="Au rythme des excédents de 2024 et 2025, il faudrait encore au moins un exercice excédentaire pour revenir à des fonds propres positifs. Le déficit de 2026 éloigne cette échéance : le redressement dépend de la maîtrise de la sinistralité et des frais.",
        source=SRC_COMPTA)
rp = R["rendement_placements"]
# ---- v6 : placements (entretien du 19/10/2026, 14h, avec le chef comptable)
DATB = [(ex, _mv(ex, r"^5221020", "mvt_d"), _mv(ex, r"^5221020", "mvt_c"), _so(ex, r"^5221020")) for ex in range(2021, 2027)]
FCP_ = _so(2025, r"^53300001$")
# ---- v11 : placements lus dans les grands livres (encours moyen pondéré par les jours), tous les dépôts à terme et le FCP
_PLG = V11["placements_gl"]
_DPEP = V11["dat_pep"]
_MOIS24 = _PLG["2024"]["dat_efc_fin_de_mois"]
contenu({"titre": "Placements du FPPN", "sous_titre": "Un dépôt mensuel pour payer l'EFC, un dépôt PEP de cinq ans, un FCP depuis 2025",
         "synthese": f"Encours moyen réellement placé en 2025 : {F('PL moy 2025', _PLG['2025']['placements_moyens'])} ; au 31/08/2026 : DAT BOA {F('PL dat 2026', _PLG['2026']['placements_cloture'] - FCP_)} et FCP {F('FCP 2025', FCP_)}"},
        "Les balances ne donnent que le solde de clôture. Le grand livre montre chaque mouvement des dépôts pendant l'année : le tableau donne l'encours moyen, pondéré par le nombre de jours où chaque solde est resté placé, et le plus haut niveau atteint par le DAT EFC.",
        tableau={"titre": tab("Encours des placements lus dans les grands livres (DAT et FCP)"), "entetes": ["Exercice", "DAT EFC : encours moyen", "DAT EFC : plus haut", "DAT PEP : encours moyen", "DAT Orabank et FCP : encours moyen", "Total : encours moyen", "Total à la clôture"],
                 "largeurs": [3.0, 4.6, 4.4, 4.6, 5.2, 4.6, 4.6], "taille": 9,
                 "lignes": [[a if a != "2026" else "2026 (8 mois)", F(f"PL efc {a}", _PLG[a]["dat_efc_moyen"]), F(f"PL pic {a}", _PLG[a]["dat_efc_pic"]), F(f"PL pep {a}", _PLG[a]["dat_pep_moyen"]),
                             F(f"PL autres {a}", _PLG[a]["dat_orabank_moyen"] + _PLG[a]["fcp_moyen"]), F(f"PL tot {a}", _PLG[a]["placements_moyens"]), F(f"PL clo {a}", _PLG[a]["placements_cloture"])]
                            for a in ("2021", "2022", "2023", "2024", "2025", "2026")]},
        disposition="dessous",
        analyse=[f"**Le DAT EFC (compte 52210200) est un dépôt d'un mois, renouvelé chaque mois et augmenté de l'épargne EFC précomptée.** En 2024, il passe de {F('PL fev24', _MOIS24[1])} fin février à {F('PL nov24', _MOIS24[10])} fin novembre, puis il est viré au FPPN le 12/12/2024 pour payer l'EFC. Les mouvements annuels du compte additionnent ces renouvellements et ne mesurent pas le montant placé.",
                 f"**Le DAT PEP (compte 52210300) est resté placé de 2020 à 2025 pour {F('PL dpep', _DPEP['montant'])}.** Il est soldé le {_DPEP['date_debouclage']} pour {F('PL dpep deb', _DPEP['debouclage_2025'])}, dont {F('PL dpep int', _DPEP['interets_2025'])} d'intérêts enregistrés en une fois en 2025. [[Le FCP de 500 000 000 FCFA a un rendement annoncé de 8 à 10 %, non garanti]] : il n'est pas revalorisé dans les comptes."],
        implication=f"Le DAT EFC finance les paiements de fin d'année de l'EFC : ce n'est pas une réserve. Seuls le DAT PEP, puis le FCP, ont été des placements durables, face à une épargne de {F('PL ep 2026', b[2026]['epargne_totale'])} au 31/08/2026.",
        valable="grands livres des comptes 2021 à 2026 (« GRAND LIVRE 2021.xlsx », « GRAND LIVRE DES COMPTES 2022.xlsx », grand livre des comptes 2023, « GRAND LIVRE FPPN 2024.xlsx », grand livres des comptes 2025 et du 01/01 au 31/08/2026), comptes 5221, 5330 et 77 ; 2026 du 01/01 au 31/08",
        source="Source : grands livres des comptes du FPPN de 2021 au 31/08/2026 ; " + SRC_COMPTA[9:] + " ; " + SRC_ENT14[9:])
contenu({"titre": "Décision de placement", "sous_titre": "Le directeur général décide seul, sur la base d'un point de trésorerie",
         "synthese": "Ni comité, ni avis du conseil d'administration, ni politique de placement écrite selon le chef comptable"},
        "Logigramme en couloirs, modifiable dans PowerPoint. Le circuit est celui décrit par le chef comptable le 19/10/2026. Cadre rouge : contrôle manquant.",
        logigramme="placement",
        analyse=["**Le choix des placements ne suit aucune politique validée par les organes du FPPN.** [[La banque propose, le directeur général décide après un point de trésorerie, et la comptabilité enregistre sur la base des relevés.]]"],
        implication="Le règlement d'exécution n° 004/2022 (art. 17) exige que les placements satisfassent aux conditions de sécurité, de rendement, de liquidité et de diversification : sans politique écrite, le FPPN ne peut pas démontrer le respect de ces conditions.",
        source=SRC_ENT14)

# ---- v10 : actifs immobilisés (CMP, projet de Cocody) et gouvernance des décisions (documents reçus le 22/10/2026)
SRC_CMP = "« ARRETE N° 0027-MSHP-DGS-DEPS DU 10-02-2020 PORTANT AUTORISATION D'OUVERTURE D'EXPLOITATION ET IMMATRICULATION DU CMP.pdf »"
SRC_COC = "« DEROGATION AUX REGLES D'URBANISME CONSTRUCTION 4 IMMEUBLES R+5 AVEC 01 CES DE 70% A USAGE D'HABITATION COCODY 2 PLTX.pdf »"
SRC_DG = "« ATTESTATION DE DESIGNATION DU DGFPPN AU TITRE DE LA MANDATURE 2020-2023.pdf »"
contenu({"titre": "Clinique, projet de Cocody et désignation du directeur général", "sous_titre": "Les documents établissent la propriété de la clinique et le projet de Cocody, pas leur valeur",
         "synthese": f"Clinique, terrains et créance sur l'hôpital au 31/08/2026 : {F('CC immo', b[2026]['terrains'] + b[2026]['cmp_investissements'] + b[2026]['creance_hopital_police'])}, financés en grande partie par l'épargne"},
        "Trois documents reçus le 22/10/2026 renseignent sur les actifs longs et sur l'organe qui décide des placements. Le tableau indique ce que chacun établit et ce qu'il laisse ouvert.",
        tableau={"titre": tab("Documents sur les actifs immobilisés et la gouvernance"), "entetes": ["Document", "Ce qu'il établit", "Ce qu'il n'établit pas", "Montant au bilan au 31/08/2026"], "largeurs": [6.4, 10.0, 8.6, 6.2], "taille": 8,
                 "lignes": [["Arrêté n° 0027/MSHP/DGS/DEPS du 10/02/2020", "Le « Centre médical du FPPN », établissement sanitaire privé mutualiste à Yopougon Andokoi, est autorisé et immatriculé (art. 1 et 2) ; le FPPN en est le promoteur et propriétaire (art. 3)", "Valeur de la clinique, résultat de son activité, comptabilité séparée", F("CC cmp", b[2026]["cmp_investissements"]) + " (investissements au CMP)"],
                            ["Lettre du ministre de la Construction, janvier 2025 (jour illisible)", "Accord pour quatre immeubles R+5 sur les lots 3005 à 3007, îlot 83, 5 007 m², titre foncier n° 119 230, Deux Plateaux, Cocody ; le FPPN avait demandé R+6", "Budget, financement, calendrier, rendement attendu, décision des organes", F("CC terrains", b[2026]["terrains"]) + " (terrains, tous sites)"],
                            ["Attestation du 04/09/2020 du président du conseil d'administration", "Désignation du directeur général pour la mandature 2020-2023, après l'Assemblée générale du 27/08/2020", "Désignation ou prolongation après 2023 ; délégation de pouvoir pour les placements", "Sans objet"]]},
        disposition="dessous",
        analyse=["**Le projet de Cocody engagerait l'épargne dans un actif immobilier de long terme, alors que l'épargne est exigible et que la liquidité ne couvre pas l'année qui vient.** Aucun document reçu n'en donne le coût ni le mode de financement.",
                 "**Le règlement d'exécution n° 004/2022 interdit de couvrir les pertes d'une activité sanitaire par les cotisations ou les réserves obligatoires (art. 27)** : la clinique doit avoir une comptabilité séparée pour le démontrer."],
        implication="Avant tout engagement sur le projet de Cocody, le FPPN doit présenter aux organes son budget, son financement et son effet sur la liquidité de l'épargne (PBC, nouveaux items).",
        valable="documents reçus le 22/10/2026 (scans lus par reconnaissance de caractères) ; montants des comptes 222, 47111101, 47111104, 47111105 et 47111100 au 31/08/2026 (provisoire)",
        source="Source : arrêté n° 0027/MSHP/DGS/DEPS du 10/02/2020 ; lettre de dérogation d'urbanisme de janvier 2025 ; attestation de désignation du 04/09/2020 (noms exacts en annexe « Fichiers sources ») ; " + REX + ", art. 27 ; " + SRC_COMPTA[9:])
contenu({"titre": "Rendement des placements", "sous_titre": "Le problème des placements est leur faible volume, pas leur rendement",
         "synthese": f"Rendement moyen apparent 2021 à 2025 : {pc('Rendement moyen', sum(r['produits_financiers'] for r in rp) / sum(r['placements_moyens'] for r in rp))} par an ; taux directeur BCEAO entre 2,00 % et 3,50 %"},
        "Rendement apparent = produits financiers (classe 77) / encours moyen des placements lu dans les grands livres (DAT et FCP, moyenne pondérée par les jours). Il est irrégulier : les intérêts du DAT PEP, constatés chaque année puis annulés, ne sont encaissés qu'à sa clôture le 30/10/2025. Taux directeurs : communiqués du Comité de politique monétaire de la BCEAO (validés).",
        graphique={"titre": fig("Taux directeur de la BCEAO"), "type": "lignes", "categories": [t[0] for t in R["externe"]["taux_directeur"]],
                   "series": [{"nom": "Taux directeur", "valeurs": [t[1] for t in R["externe"]["taux_directeur"]], "couleur": BLEU}], "format": "0.00", "legende": None,
                   "unite": "En %, date d'effet"},
        analyse=[f"**Les produits financiers de 2021 à 2025 atteignent {F('Produits fin', sum(r['produits_financiers'] for r in rp))}**, pour un encours moyen de placements de {F('Placements moyens', sum(r['placements_moyens'] for r in rp) / len(rp))} (grands livres).",
                 f"**Le rendement moyen dépasse le taux directeur, mais il est irrégulier.** En 2025, {F('RDP int pep', V11['dat_pep']['interets_2025'])} d'intérêts du DAT PEP, accumulés depuis 2020, sont enregistrés en une fois.",
                 "**Le taux servi sur l'épargne (3,5 % sur le PEP, 3,5 % ou 5,5 % sur l'EFC selon le document) est un engagement de rendement** que les produits des placements ne couvrent pas (partie 6).",
                 f"**Le vrai sujet est le volume placé**, face à {F('Épargne (rend)', b[2026]['epargne_totale'])} d'épargne à restituer. Le solde des comptes courants non rémunérés atteint {F('Comptes courants 2026', AG[2026]['comptes_courants_521'])} au 31/08/2026."],
        implication="La Décision n° 05/2022 demande aux mutuelles qui gèrent des risques longs un solde des comptes courants non rémunérés proche de zéro : une partie de cette trésorerie peut être placée à terme sans nuire à la liquidité.",
        source="Source : BCEAO, communiqués du Comité de politique monétaire ; " + SRC_COMPTA[9:])
# ---- v10 : conclusion « adéquation des placements aux engagements » (règlement d'exécution n° 004/2022, art. 17)
_pf_ = {x["exercice"]: x for x in R["rendement_placements"]}
contenu({"titre": "Conclusion : adéquation des placements", "sous_titre": "Les placements ne sont adéquats ni en volume, ni en liquidité, ni en rendement",
         "synthese": "Aucune des quatre conditions de l'article 17 (sécurité, rendement, liquidité, diversification) n'est démontrée"},
        "Le règlement d'exécution n° 004/2022 (art. 17) exige que les placements satisfassent aux conditions de sécurité, de rendement, de liquidité et de diversification. Chaque condition est confrontée aux chiffres du FPPN de 2020 au 31/08/2026.",
        tableau={"titre": tab("Placements et emplois du FPPN confrontés aux conditions de l'article 17"), "entetes": ["Condition", "Ce qu'il faudrait", "Ce que montrent les chiffres", "Appréciation"], "largeurs": [3.6, 8.4, 14.0, 5.2], "taille": 8.5,
                 "lignes": [["Liquidité", "Couvrir les paiements de l'année, sorties d'épargne comprises", "Couverture de " + pc("CPL liq", couverture("31/08/2026", T_MOY)[2]) + " au 31/08/2026 (scénario central)", "Non remplie"],
                            ["Rendement", "Couvrir les intérêts dus sur l'épargne", "Produits financiers 2021 à 2025 : " + F("CPL pf", sum(_pf_[a]["produits_financiers"] for a in range(2021, 2026))) + " ; intérêts dus sur la même période : " + F("CPL int", interets_dus(0.035, 0.055, 2025) - interets_dus(0.035, 0.055, 2020)), "Non remplie"],
                            ["Sécurité", "Actifs évalués et contreparties solides", "Clinique, terrains et créance sur l'hôpital non évalués ; FCP non garanti", "Non démontrée"],
                            ["Diversification", "Pas de concentration sur un emploi", "Plus des trois quarts de l'actif à plus de cinq ans ou à délai incertain", "Non remplie"]]},
        disposition="dessous",
        analyse=["**Le portefeuille est inadéquat d'abord parce que l'épargne finance des actifs qui ne produisent ni revenu financier ni liquidité.** Le rendement des placements financiers eux-mêmes n'est pas le problème principal."],
        implication="Une politique de placement votée par le conseil d'administration doit fixer une réserve de liquidité, un rendement minimal au regard du taux servi arrêté et des limites sur les actifs immobiliers, avant tout nouvel investissement.",
        valable="placements et produits financiers de 2021 à 2025 ; actifs et épargne au 31/08/2026 (provisoire) ; intérêts dus aux taux du manuel d'épargne",
        source="Source : " + REX + ", art. 17 ; " + SRC_COMPTA[9:] + " ; diapositives " + ", ".join(ref(t) for t in ("Adossement : tests de résistance", "Épargne : intérêts dus non comptabilisés", "Rendement des placements")))

retenir("bilan", ("Le bilan combine provisions figées, fonds propres négatifs et actifs peu liquides", "Trois fragilités cumulées limitent la capacité du FPPN à faire face à un choc",
                  "Le FPPN porte des engagements courts et une épargne croissante avec des actifs longs, des provisions non recalculées et des fonds propres négatifs."),
        "Les provisions gardent le même montant de 2020 à 2026, les fonds propres sont négatifs à chaque clôture, et plus des trois quarts de l'actif sont à plus de cinq ans ou de délai incertain.",
        "Le FPPN ne peut absorber ni un choc de prestations, ni une vague de retraits, sans céder des actifs peu liquides.",
        "Le FPPN doit adopter une politique de provisionnement, une politique de placement et un encadrement écrit des investissements dans les structures liées.",
        "Ces constats appellent une décision de gouvernance : politique de provisionnement, politique de placement, encadrement des investissements dans les structures liées.")

# ---------------------------------------------------------------- PARTIE 10 : CONFORMITÉ PRUDENTIELLE
separateur(10, "Conformité prudentielle", "Textes UEMOA, ratios, réserves, plan comptable, leviers de conformité",
           "Cette partie applique aux comptes du FPPN les normes de la mutualité sociale UEMOA. Les textes viennent de la documentation publique de l'AIRMS fournie le 18/10/2026 et validée comme source.")
contenu({"titre": "Textes applicables", "sous_titre": "Cinq textes fixent les règles prudentielles que le FPPN doit respecter",
         "synthese": "Les articles cités sont lus dans les textes intégraux ; les pages scannées ont été lues par reconnaissance optique de caractères (OCR) et vérifiées à l'image"},
        "Hiérarchie : le règlement de 2009 fixe les principes et renvoie au règlement d'exécution pour les règles prudentielles (art. 58) ; la Décision de 2022 chiffre les normes ; l'arrêté de 2023 fixe les sanctions en Côte d'Ivoire. Le PCMS donne les numéros de comptes que la Décision utilise dans ses formules.",
        tableau={"titre": tab("Textes de la mutualité sociale appliqués au FPPN"), "entetes": ["Texte", "Ce qu'il impose au FPPN", "Articles utilisés"], "largeurs": [8.6, 15.8, 6.8],
                 "lignes": [["Règlement n° 07/2009/CM/UEMOA du 26/06/2009", "Il impose l'agrément, une comptabilité selon le plan comptable, la répartition des excédents en réserves et le placement des fonds en banque.", "Art. 22, 28, 60 à 63, 83, 84"],
                            ["Règlement d'exécution n° 004/2022/COM/UEMOA du 01/06/2022", "Il impose des provisions suffisantes, une réserve légale, des règles propres aux risques longs et le cantonnement des patrimoines.", "Art. 3 à 27, 35"],
                            ["Décision n° 05/2022/COM/UEMOA du 01/06/2022", "Elle fixe 13 indicateurs de gouvernance, 11 indicateurs de contrôle, 9 ratios pour les risques longs, la marge de solvabilité et l'affectation des excédents.", "Annexe, tableaux 1 et 2, § 1.1 à 2.9"],
                            ["Règlement n° 03/2012/CM/UEMOA et PCMS de mai 2012", "Ils imposent le plan comptable des mutuelles sociales à partir de l'exercice ouvert le 1er janvier 2014.", "Art. 2 ; liste des comptes ; fiches 1, 2, 6, 25"],
                            ["Arrêté n° 00115/MEPS/CAB du 12/12/2023", "Il fixe les sanctions administratives et pécuniaires en cas de manquement.", "Annexe, lignes 6 et 13"]]},
        disposition="dessous",
        analyse=["**Ces textes forment un ensemble complet et chiffré depuis juin 2022.** Ils couvrent la comptabilité, les provisions, les réserves, la liquidité et les sanctions."],
        implication="Le FPPN dépasse les seuils du système normal du plan comptable (plus de 10 millions de FCFA de cotisations ou plus de 1 000 adhérents, art. 4 du règlement d'exécution) : il doit tenir une comptabilité d'engagement et produire un TAFIRE (art. 5).",
        source="Source : textes cités, documentation publique de l'AIRMS fournie le 18/10/2026 (validée)")
contenu({"titre": "Risques courts et risques longs", "sous_titre": "Le FPPN couvre des risques courts et, avec l'épargne, des risques longs",
         "synthese": "Un risque est long quand plus d'un an sépare la cotisation de la prestation (règlement d'exécution n° 004/2022, art. 1)"},
        "Définitions de l'article 1 du règlement d'exécution : risque court quand l'écart entre l'encaissement des cotisations et le versement des prestations est d'un an au plus ; risque long au-delà. L'article 3 classe la vieillesse et le décès lié à un mécanisme de capitalisation parmi les risques longs. La qualification retenue ici doit être confirmée par l'AIRMS ou par l'agrément du FPPN, non transmis.",
        tableau={"titre": tab("Qualification des garanties du FPPN"), "entetes": ["Garantie", "Délai entre cotisation et prestation", "Qualification proposée", "Conséquence"], "largeurs": [4.4, 8.4, 6.6, 11.8],
                 "lignes": [["Maladie (936, CGRAE)", "Les soins sont payés dans l'année.", "Risque court", "Ratios du tableau 1 de la Décision n° 05/2022."],
                            ["Décès (8Z6)", "Le capital est payé au décès, sans capitalisation connue.", "Risque court", "Ratios du tableau 1 ; qualification à revoir si un mécanisme de capitalisation existe."],
                            ["EFC (870)", "L'épargne est versée en fin de carrière, des années après les versements.", "Risque long (vieillesse)", "Ratios de la section II, réserve de trésorerie, étude actuarielle, certification."],
                            ["PEP (998)", "L'épargne est restituée à l'échéance du plan, en général après plus d'un an.", "Risque long (plan d'au moins trois ans, règlement intérieur, art. 66)", "Mêmes obligations que l'EFC."]]},
        disposition="dessous",
        analyse=["**Avec l'EFC et le PEP, le FPPN entre dans le champ des risques longs.** Ces deux produits représentent " + F("Épargne risques longs", b[2026]["epargne_totale"]) + " d'engagements au 31/08/2026."],
        implication="Une mutuelle qui couvre à la fois des risques courts et longs doit cantonner ses patrimoines et séparer ses opérations (art. 15), faire certifier ses comptes par un commissaire aux comptes (art. 12) et réaliser une étude actuarielle tous les cinq ans (art. 14).",
        source="Source : " + REX + ", art. 1, 3, 12 à 15 ; fichiers de précomptes ; balances du FPPN")
gab(31, {"titre": "Lire un ratio : les tiroirs du bilan", "sous_titre": "Chaque compte est rangé dans un tiroir selon le temps que met l'argent à entrer ou sortir",
         "synthese": "Un ratio compare le contenu de deux tiroirs ; le classement des comptes décide donc du résultat"},
    {"cartes": [("Ressources longues", "Argent qui reste longtemps à la disposition du FPPN : fonds propres, provisions, emprunts."),
                ("Dettes à court terme", "Argent que le FPPN doit rendre dans l'année : fournisseurs, cliniques, État."),
                ("Emplois longs", "Argent placé pour longtemps : terrains, bâtiments, prêts, avances aux structures liées."),
                ("Trésorerie", "Argent disponible tout de suite ou presque : caisse, comptes en banque, dépôts à terme.")]},
    "Image utilisée pour un lecteur non spécialiste : le bilan est une commode à tiroirs. Le plan comptable (PCMS) donne à chaque compte un numéro qui indique son tiroir : classe 1 pour les ressources longues, classe 2 pour les emplois longs, classe 4 pour les tiers, classe 5 pour la trésorerie. Un ratio de liquidité compare le tiroir « trésorerie » au tiroir « dettes à court terme ». Si un compte est rangé dans le mauvais tiroir, le ratio change sans que la réalité change. C'est pourquoi la suite présente chaque ratio deux fois : avec les comptes rangés comme le FPPN les tient, puis rangés selon leur nature.",
    source="Source : PCMS (structure des classes de comptes) ; Adenka")
contenu({"titre": "Deux lectures : comme le FPPN, puis corrigée selon le PCMS", "sous_titre": "Trois groupes de comptes sont rangés dans un tiroir qui ne reflète pas leur nature",
         "synthese": "Adenka calcule d'abord chaque ratio avec les comptes tels qu'ils sont tenus, puis avec trois reclassements motivés"},
        "Le classement « comme le FPPN » suit les numéros de comptes cités par la Décision n° 05/2022 dans ses formules. Le classement corrigé selon le PCMS ne modifie aucun montant : il déplace seulement des comptes d'un tiroir à un autre, pour une raison écrite. Le détail compte par compte (246 comptes) figure dans l'onglet « 10_classement_comptes » du classeur des calculs.",
        tableau={"titre": tab("Reclassements de la lecture corrigée selon le PCMS"), "entetes": ["Comptes", "Tiroir comme le FPPN", "Tiroir selon le PCMS", "Raison du reclassement"], "largeurs": [6.6, 5.4, 5.4, 13.8],
                 "lignes": [["41195100, 41195200 : épargne PEP et EFC", "Dette à court terme (classe 4)", "Ressource longue", "L'épargne est un risque long (art. 3 du règlement d'exécution) : elle n'est pas exigible en totalité dans l'année."],
                            ["47111100 à 47111105, 47111201, 485 : avances et créances durables", "Créance à court terme (classe 4)", "Emploi long", "Ces montants financent le CMP, le Grenier et l'Hôpital de la Police depuis plusieurs années, sans remboursement à court terme."],
                            ["10620000 : réévaluation", "Réserve libre (libellé « Réserve disponible » en 2026)", "Écart de réévaluation (terrain du siège selon le chef comptable)", "Le PCMS réserve le compte 106 aux écarts de réévaluation des immeubles, qui ne comptent pas dans la marge de solvabilité."]]},
        disposition="dessous",
        analyse=["**Les deux lectures encadrent la réalité.** La lecture comme le FPPN est prudente pour l'épargne, mais optimiste pour les avances durables ; la lecture corrigée selon le PCMS corrige les deux."],
        implication="La réponse définitive dépend de l'échéancier réel de l'épargne : les conditions de retrait et les âges des épargnants trancheront entre les deux lectures.",
        source="Source : PCMS, fiches n° 1 et 25 ; " + REX + ", art. 3 ; balances du FPPN ; classement Adenka")
methode_gab({"titre": "Règle de statut", "sous_titre": "Le statut de chaque ratio suit une règle écrite et vérifiable",
             "synthese": "Le statut compare la valeur calculée à la norme du texte ; il est jugé sur les exercices 2022 à 2025"},
            [("Conforme ou non conforme", "Un ratio est conforme quand sa valeur respecte la norme chiffrée de la Décision n° 05/2022 ; il est non conforme sinon. Pour une fourchette, une valeur au-dessus de la borne haute est signalée à part."),
             ("En attente", "Un ratio est en attente quand une donnée nécessaire n'a pas été transmise : fichiers aller des postes 936 et 8Z6 (détenus par la direction du personnel de la Police), procès-verbaux, registre des contrôles, cotisation au FNG."),
             ("À apprécier", "Un ratio est à apprécier quand la norme n'est pas chiffrée : délai contractuel, objectif du plan stratégique, solde « avoisinant zéro », taux moyen du marché."),
             ("Période jugée", "La Décision est entrée en vigueur le 1er juin 2022 : 2020 et 2021 sont indicatifs, comme les huit mois provisoires de 2026. Deux exercices non conformes de suite sont signalés, car l'article 84 du règlement n° 07/2009 prévoit alors le retrait d'agrément.")],
            "Le statut ne juge pas la gestion du FPPN : il constate l'écart entre un chiffre et une norme. Le statut global d'un ratio est non conforme s'il est non conforme au moins une année de 2022 à 2025.",
            source="Source : " + DEC + ", article 3 ; " + R07 + ", art. 84 ; Adenka")
gab(29, {"titre": "Quatre familles de ratios", "sous_titre": "Les 36 indicateurs examinés se regroupent en quatre familles, de la plus actuarielle à la gouvernance",
         "synthese": "L'ordre de priorité suit l'importance pour la protection des adhérents : solvabilité d'abord, gouvernance ensuite"},
    {"blocs": [("Sinistralité et solvabilité", "Sinistralité, marge de solvabilité, autonomie financière, excédent net : ces ratios disent si les cotisations suffisent et si le FPPN peut absorber un choc."),
               ("Réserves et fonds propres", "Réserve obligatoire, fonds propres et assimilés, fonds des risques longs : ces ratios disent si le FPPN a mis de côté assez de ressources propres."),
               ("Liquidité et équilibre", "Liquidité, trésorerie à échéance et immédiate, équilibre stable, solvabilité de l'épargne : ces ratios disent si l'argent est disponible à temps."),
               ("Gestion et gouvernance", "Frais, personnel, recouvrement, délais, contrôles, réunions des organes, FNG : ces ratios disent si l'organisation est efficace et contrôlée.")]},
    "Les familles reprennent les catégories de l'annexe de la Décision n° 05/2022 (performance, viabilité, pérennité, risques longs), réordonnées selon la priorité actuarielle demandée pour ce rapport.",
    source="Source : " + DEC + " ; classement Adenka")


def ligne_bord(x, annees=(2020, 2021, 2022, 2023, 2024, 2025, 2026)):
    vals = []
    for a in annees:
        v_ = val_ratio(x, a) if a in x["valeurs"] else "n.d."
        vals.append(v_ if v_ == "n.d." else bv(v_, boule_statut(x["statuts"].get(a, ""))))
    nom = x["nom"].replace(" (risques longs)", "")
    return [str(x["priorite"]), nom, x["norme"].replace(" (tableau 1) ; 9 mois au § 1.2.2.5.4", "").replace(" (marge disponible au moins égale à la marge exigée)", ""), *vals, x["statut_global"]]


bord_ent = ["N°", "Ratio", "Norme", "2020", "2021", "2022", "2023", "2024", "2025", "31/08/2026", "Statut"]
bord_larg = [0.8, 6.3, 4.0, 2.56, 2.56, 2.56, 2.56, 2.56, 2.56, 2.56, 2.18]
LEG_BORD = LEG_CONF
# ---- v12 : ratios par garantie, avant les ratios globaux : matrice de 2020 au 31/08/2026, puis détail des montants
_RGL = V11["ratios_garantie"]
_RGg = {(x["exercice"], x["garantie"]): x for x in _RGL}
_MAT_G = []
for _g in ("Maladie", "Décès"):
    _MAT_G.append([f"{_g} : sinistralité", "Repère 70 % à 80 %"] + [_gar(_g, a, "sinistralite") for a in _AN])
_MAT_G.append(["Toutes garanties : taux de sinistralité réglementaire", "Norme 70 % à 80 %"] + [_rt("SIN", a) for a in _AN])
for _g in ("Maladie", "Décès"):
    _MAT_G.append([f"{_g} : ratio combiné (prestations et frais)", "Au plus 100 %"] + [_gar(_g, a, "ratio_combine") for a in _AN])
for _g in ("Maladie", "Décès"):
    _MAT_G.append([f"{_g} : résultat technique", "Au moins 0 FCFA"] + [_gar(_g, a, "resultat_technique") for a in _AN])
contenu({"titre": "Ratios par garantie : matrice de 2020 au 31/08/2026", "sous_titre": "La maladie est déficitaire depuis 2025, le décès reste très excédentaire",
         "synthese": "La Décision n° 05/2022 ne fixe qu'un taux de sinistralité global ; les ratios par garantie montrent quelle garantie le fait monter"},
        "La Décision n° 05/2022 (§ 1.2.1.2.5) définit un seul taux de sinistralité : toutes les prestations de l'exercice avec la variation des provisions (comptes 60, 6901 et 7901), rapportées à toutes les cotisations acquises (comptes 70, 7906 et 6906). La norme de 70 % à 80 % s'applique à ce taux global. Les ratios par garantie sont un outil d'analyse ; la fourchette y sert de repère.",
        tableau={"titre": tab("Ratios par garantie et taux réglementaire, de 2020 au 31/08/2026"), "entetes": ["Ratio", "Norme ou repère"] + _ENT_AN,
                 "largeurs": [7.2, 3.2] + [2.97] * 7, "taille": 8, "sans_code": True, "legende": LEG_MAT[:3], "lignes": _MAT_G},
        disposition="dessous",
        encadre=["**Sinistralité d'une garantie** = prestations de la garantie / cotisations de la garantie. Elle montre la part des cotisations consommée par les soins (maladie) ou les allocations (décès).",
                 "**Ratio combiné** = (prestations + charges de fonctionnement affectées) / cotisations. Au-dessus de 100 %, la garantie coûte plus qu'elle ne rapporte, frais compris. Frais répartis au prorata des cotisations, faute de clé fournie par le FPPN.",
                 "**Résultat technique** = cotisations - prestations - charges de fonctionnement affectées, hors produits financiers et éléments exceptionnels. Négatif : la garantie est déficitaire par elle-même."],
        analyse=["**Le taux réglementaire global réunit les deux garanties : il masque une maladie déficitaire et un décès très excédentaire.** "
                 f"En 2025, la maladie consomme {pc('RGM sin m25', _RGg[(2025, 'Maladie')]['sinistralite'])} de ses cotisations, le décès {pc('RGM sin d25', _RGg[(2025, 'Décès')]['sinistralite'])}, pour un taux global de {pc('RGM sin 25', RT['SIN']['valeurs'][2025])}."],
        implication="Le respect de la norme globale ne suffit pas : la cotisation maladie doit être fixée sur le coût de la garantie maladie, sans compter sur l'excédent du décès.",
        valable="comptes de 2020 à 2025 et du 01/01 au 31/08/2026 (balance provisoire) ; en 2020, le décès comprend la prime d'assurance décès (compte 65180000), hors du taux réglementaire",
        source=SRC_COMPTA + " ; " + DEC + ", § 1.2.1.2.5 ; statuts du FPPN édités le 15/06/2023, art. 17")
contenu({"titre": "Ratios par garantie : montants", "sous_titre": "Cotisations, prestations, frais affectés et résultat technique de chaque garantie",
         "synthese": "Résultat technique = cotisations acquises - prestations (variation des provisions comprise) - charges de fonctionnement affectées, hors produits financiers et éléments exceptionnels"},
        "Charges de fonctionnement (comptes 612, 613, 618 et 62 à 68) réparties au prorata des cotisations, faute de clé fournie par le FPPN. Décès : comptes 6021 et 6024, prime d'assurance décès comprise en 2020, cotisations du compte 702 (décès et MEDOC).",
        tableau={"titre": tab("Sinistralité, frais, ratio combiné et résultat technique par garantie"), "entetes": ["Exercice", "Garantie", "Cotisations", "Prestations", "Sinistralité", "Frais affectés", "Ratio combiné", "Résultat technique"],
                 "largeurs": [2.6, 2.8, 4.6, 4.6, 3.0, 4.6, 3.0, 5.0], "taille": 8, "sans_code": True, "legende": LEG_MAT[:3],
                 "lignes": [[str(x["exercice"]) if x["exercice"] < 2026 else "2026 (8 mois)", x["garantie"], F(f"RG cot {x['exercice']} {x['garantie']}", x["cotisations"]), F(f"RG pre {x['exercice']} {x['garantie']}", x["prestations"]),
                             _gar(x["garantie"], x["exercice"], "sinistralite"), F(f"RG fr {x['exercice']} {x['garantie']}", x["frais"]), _gar(x["garantie"], x["exercice"], "ratio_combine"),
                             _gar(x["garantie"], x["exercice"], "resultat_technique")] for x in _RGL if x["exercice"] >= 2022]},
        disposition="dessous",
        analyse=[f"**Le ratio combiné de la maladie dépasse 100 % en 2025 ({pc('RG rc m25 (a)', _RGg[(2025, 'Maladie')]['ratio_combine'])}) et sur les huit mois de 2026.** Celui du décès reste sous 70 % chaque année depuis 2022."],
        implication="La cotisation maladie de 3,5 % du salaire de base (statuts, art. 17) ne couvre plus le coût complet de la garantie maladie depuis 2025. Sa révision relève du conseil d'administration et de l'Assemblée générale.",
        valable="comptes 2022 à 2025 et du 01/01 au 31/08/2026 (provisoire) ; frais répartis au prorata des cotisations",
        source=SRC_COMPTA + " ; " + DEC + ", § 1.1.1 ; statuts du FPPN édités le 15/06/2023, art. 17")

_defs_diapo(1, ("SIN", "MSE", "ROB", "FEP", "LIQ", "EQS", "AUT", "EXN", "TFP"))
contenu({"titre": "Tableau de bord (1/4) : solvabilité et réserves", "sous_titre": "La sinistralité globale, les réserves, les fonds propres et la liquidité sont non conformes",
         "synthese": "Ratios de priorité actuarielle 1 à 9, lecture comme le FPPN ; 2020, 2021 et 2026 sont indicatifs"},
        "Valeurs calculées sur les balances, avec les comptes cités par la Décision n° 05/2022. La sinistralité réglementaire est globale (maladie et décès) : elle est sous la fourchette de 70 % à 80 % parce que la garantie décès consomme peu de cotisations. La marge de solvabilité est conforme dans la lecture comme le FPPN, mais pas une fois les pertes reportées déduites (diapositives suivantes).",
        tableau={"titre": tab("Tableau de bord des ratios prudentiels, priorités 1 à 9"), "entetes": bord_ent, "largeurs": bord_larg, "statut_col": 10, "taille": 8.5, "legende": LEG_BORD,
                 "lignes": [ligne_bord(RT[c]) for c in ("SIN", "MSE", "ROB", "FEP", "LIQ", "EQS", "AUT", "EXN", "TFP")]},
        disposition="dessous",
        analyse=["**Six des neuf ratios prioritaires sont non conformes plusieurs années de suite.** Seuls l'autonomie financière, le financement propre et la marge de solvabilité calculée comme le FPPN respectent la norme."],
        implication="La non-conformité répétée de la réserve obligatoire, des fonds propres et de la liquidité expose le FPPN aux sanctions prévues pour le non-respect des règles prudentielles sur deux exercices.",
        source=SRC_REG)
_defs_diapo(2, ("RL1", "RL2", "RL3", "RL4", "RL5", "RL6", "RL7"))
contenu({"titre": "Tableau de bord (2/4) : risques longs", "sous_titre": "Aucun ratio propre aux risques longs n'est respecté",
         "synthese": "Ratios de priorité 10 à 16, applicables à l'épargne EFC et PEP qualifiée de risque long"},
        "Section II de l'annexe de la Décision n° 05/2022 et articles 19 et 21 du règlement d'exécution. Le ratio de solvabilité de l'épargne reprend la couverture calculée en partie 6. La réserve de trésorerie vaut zéro car aucun compte de réserve de trésorerie n'existe dans les balances.",
        tableau={"titre": tab("Tableau de bord des ratios des risques longs, priorités 10 à 16"), "entetes": bord_ent, "largeurs": bord_larg, "statut_col": 10, "taille": 8.5, "legende": LEG_BORD,
                 "lignes": [ligne_bord(RT[c]) for c in ("RL1", "RL2", "RL3", "RL4", "RL5", "RL6", "RL7")]},
        disposition="dessous",
        analyse=["**Les engagements d'épargne ne sont adossés ni à des actifs réservés suffisants, ni à une réserve de trésorerie, ni à des fonds propres positifs.** Les deux derniers indicateurs n'ont pas de seuil chiffré et sont à apprécier."],
        implication="Pour les risques longs, le règlement d'exécution exige aussi une étude actuarielle tous les cinq ans et la certification des comptes : le FPPN doit indiquer s'il les a réalisées.",
        source=SRC_REG)
_defs_diapo(3, ("CHF", "CHB", "CUP", "CHP", "FID", "DPP", "FOR", "CRB", "REC"))
contenu({"titre": "Tableau de bord (3/4) : gestion", "sous_titre": "Les frais dépassent nettement les normes, alors que le personnel et la fidélisation sont conformes",
         "synthese": "Ratios de gestion de priorité 17 à 25 ; le recouvrement n'est mesurable que pour les retraités précomptés par la CGRAE"},
        "Les charges de fonctionnement suivent la liste de comptes de l'annexe (612, 613, 618, 62 à 68). La fidélisation compare les matricules précomptés au poste 936 d'une année à l'autre. Le délai de paiement des prestataires est à apprécier, car les délais contractuels des conventions n'ont pas été transmis.",
        tableau={"titre": tab("Tableau de bord des ratios de gestion, priorités 17 à 25"), "entetes": bord_ent, "largeurs": bord_larg, "statut_col": 10, "taille": 8.5, "legende": LEG_BORD,
                 "lignes": [ligne_bord(RT[c]) for c in ("CHF", "CHB", "CUP", "CHP", "FID", "DPP", "FOR", "CRB")] +
                           [[str(RT["REC"]["priorite"]), RT["REC"]["nom"] + " (retraités précomptés par la CGRAE, février 2024 à août 2026)", RT["REC"]["norme"], "", "", "", "", "", "", bv(pc("REC CGRAE", Q["effectifs"]["recouvrement_cgrae"]), "🟢"), RT["REC"]["statut_global"]]]},
        disposition="dessous",
        analyse=["**Les trois ratios de frais sont non conformes chaque année.** Les frais de fonctionnement absorbent plus d'un tiers des cotisations, contre 20 % au plus selon la norme."],
        implication="La baisse du coût par unité prestée vient surtout de la hausse des prestations, pas d'une baisse des frais : le FPPN doit présenter un plan de réduction des charges de fonctionnement.",
        source=SRC_REG + " ; fichiers de précomptes")
contenu({"titre": "Tableau de bord (4/4) : données en attente", "sous_titre": "Onze indicateurs attendent des documents de gestion ou de gouvernance",
         "synthese": "Ces indicateurs ne sont pas exclus : ils seront calculés dès réception des pièces demandées"},
        "Les indicateurs de gouvernance sont placés en dernier, conformément à la priorité actuarielle du rapport. Chaque pièce manquante est ajoutée à la liste des documents demandés (PBC) et aux guides d'entretien.",
        tableau={"titre": tab("Indicateurs en attente de documents, priorités 26 à 36"), "entetes": ["N°", "Indicateur", "Norme", "Pièce à obtenir"], "largeurs": [0.9, 11.0, 6.0, 13.3], "taille": 9,
                 "lignes": [[str(RT[c]["priorite"]), RT[c]["nom"], RT[c]["norme"], RT[c]["qualitatif"]] for c in ("RBA", "PEN", "IND", "CAP", "DLQ", "RAT", "DRB", "FNG", "TTR", "OSA", "OSF")]},
        disposition="dessous",
        analyse=["**L'absence de compte du fonds national de garantie dans les balances est le point le plus sensible.** L'adhésion au fonds est obligatoire (règlement n° 07/2009, art. 63), et son défaut est sanctionné (arrêté n° 00115/MEPS/CAB, ligne 8)."],
        source=SRC_REG + " ; liste des documents demandés (PBC)")
contenu({"titre": "Liquidité et équilibre : deux lectures", "sous_titre": "Le classement de l'épargne décide de la conformité de la liquidité et de l'équilibre stable",
         "synthese": f"Liquidité 2025 : {val_ratio(RT['LIQ'], 2025)} comme le FPPN, {val_ratio(RT['LIQ'], 2025, True)} corrigée selon le PCMS ; équilibre stable 2025 : {val_ratio(RT['EQS'], 2025)} et {val_ratio(RT['EQS'], 2025, True)}"},
        "Liquidité : norme d'au moins 1. Équilibre stable : norme d'au plus 1. Lecture corrigée selon le PCMS : épargne en ressource longue, avances durables en emplois longs (diapositive « Deux lectures »). En 2020, les ressources longues comme le FPPN sont négatives : le ratio d'équilibre stable n'a pas de sens et n'est pas représenté.",
        graphique={"titre": fig("Liquidité et équilibre stable, comme le FPPN et corrigés selon le PCMS"), "type": "lignes", "categories": [str(a) for a in range(2021, 2027)],
                   "series": [{"nom": "Liquidité comme le FPPN", "valeurs": [round(RT["LIQ"]["valeurs"][a], 2) for a in range(2021, 2027)], "couleur": BLEU},
                              {"nom": "Liquidité corrigée selon le PCMS", "valeurs": [round(RT["LIQ"]["valeurs_retraitees"][a], 2) for a in range(2021, 2027)], "couleur": ORANGE},
                              {"nom": "Équilibre stable corrigé selon le PCMS", "valeurs": [round(RT["EQS"]["valeurs_retraitees"][a], 2) for a in range(2021, 2027)], "couleur": GRIS_F},
                              {"nom": "Norme (1)", "valeurs": [1.0] * 6, "couleur": ROUGE}],
                   "format": "0.00", "unite": "Ratio (sans unité)"},
        analyse=[f"**Comme le FPPN tient ses comptes, la liquidité n'atteint jamais 1.** Elle vaut au mieux {val_ratio(RT['LIQ'], 2021)} en 2021, car l'épargne figure parmi les dettes à court terme.",
                 "**Corrigée selon le PCMS, la liquidité devient conforme en 2025 et sur les huit mois de 2026.** La trésorerie a fortement augmenté pendant que les dettes à court terme hors épargne restaient stables.",
                 f"**Corrigé selon le PCMS, l'équilibre stable devient conforme en 2025** ({val_ratio(RT['EQS'], 2025, True)}), car l'épargne finance alors durablement les emplois longs."],
        implication="La conformité dépend de la part réellement exigible de l'épargne dans l'année : le FPPN doit établir l'échéancier de l'EFC (départs à la retraite) et du PEP (échéances des plans).",
        source=SRC_REG)
a25 = AG[2025]
contenu({"titre": "Calcul pas à pas : sinistralité réglementaire 2025", "sous_titre": "La sinistralité globale est sous la fourchette parce que le décès consomme peu",
         "synthese": f"Charges de prestations {F('SIN num 2025', a25['charges_prestations'])} / cotisations acquises {F('SIN den 2025', a25['cotisations_acquises'])} = {pc('SIN 2025 (pas)', a25['charges_prestations'] / a25['cotisations_acquises'])}"},
        "Formule de la Décision (§ 1.2.1.2.5) : comptes 60 + 6901 - 7901 au numérateur, comptes 70 + 7906 - 6906 au dénominateur. Le FPPN n'a pas de compte 6901, 7901, 6906 ni 7906 : les provisions techniques ne varient pas. Le compte 704 « cartes d'accès aux soins » appartient à la classe 70 et entre donc dans les cotisations acquises.",
        pas_a_pas={"entetes": ["Élément", "Comptes", "Montant 2025"],
                   "lignes": [["Cotisations maladie", "701", F("SIN cot mal", a25["cot_maladie"])],
                              ["Cotisations décès et MEDOC", "702", F("SIN cot dec", a25["cot_deces"])],
                              ["Autres produits techniques", "704", F("SIN cot autres", a25["cotisations_acquises"] - a25["cot_maladie"] - a25["cot_deces"])],
                              ["Cotisations acquises", "70", F("SIN cot acq", a25["cotisations_acquises"])],
                              ["Prestations maladie", "601, 6013, 6018, 60800007", F("SIN prest mal", a25["prest_maladie"])],
                              ["Prestations décès", "6021, 6024", F("SIN prest dec", a25["prest_deces"])],
                              ["Autres prestations", "60800008, 60850000", F("SIN prest autres", a25["charges_prestations"] - a25["prest_maladie"] - a25["prest_deces"])],
                              ["Charges de prestations", "60", F("SIN charges", a25["charges_prestations"])],
                              ["Taux de sinistralité", "60 / 70", pc("SIN taux (t)", a25["charges_prestations"] / a25["cotisations_acquises"])]],
                   "largeurs": [6.2, 6.4, 5.6], "rouge": [(3, 2), (7, 2), (8, 2)],
                   "explications": ["Les cotisations acquises additionnent toute la classe 70, y compris les autres produits techniques.", "Les charges de prestations additionnent toute la classe 60, maladie et décès confondus.", "Le taux global se situe sous la borne basse de 70 % de la norme."],
                   "fichier": "AUDIT ADENKA\\DAFC\\2025\\BALANCE FPPN 2025.xlsx"},
        analyse=["**Une sinistralité trop basse est aussi un écart à la norme.** La fourchette de 70 % à 80 % vise à ce que les cotisations reviennent aux adhérents sous forme de prestations."],
        implication=f"Appliqué à la seule maladie, le taux 2025 ({pc('S/C maladie 2025 (SIN)', s[2025]['sc_maladie'])}) se situe dans la fourchette. L'écart global vient de la garantie décès : son niveau de cotisation mérite d'être revu au vu de la mortalité observée, dès réception des données.",
        source=SRC_COMPTA + " ; " + DEC + ", § 1.2.1.2.5")
m25 = MS[2025]
contenu({"titre": "Calcul pas à pas : marge de solvabilité 2025", "sous_titre": "La marge est conforme si l'on ignore les pertes reportées, non conforme sinon",
         "synthese": f"Marge exigée {F('MSE exigée 2025', m25['marge_exigee'])} ; marge disponible {F('MSE dispo FPPN 2025', m25['marge_disponible_comme_fppn'])} comme le FPPN, {F('MSE dispo nette 2025', m25['marge_disponible_nette'])} nette des pertes"},
        "Formule (Décision, § 1.1.1) : marge exigée = le maximum entre 18 % des cotisations acquises de l'exercice précédent et de 26 % de la moyenne des prestations maladie des trois derniers exercices. Marge disponible = fonds de réserve obligatoire + fonds d'établissement + réserves libres + droits d'adhésion + dons et legs non affectés. Le texte ne mentionne pas les pertes reportées ; la lecture corrigée selon le PCMS les déduit, car le PCMS définit les réserves comme des excédents laissés à la disposition de la mutuelle (fiche n° 2).",
        pas_a_pas={"entetes": ["Élément", "Calcul", "Montant"],
                   "lignes": [["18 % des cotisations acquises 2024", "18 % x " + F("MSE cot 2024", AG[2024]["cotisations_acquises"]), F("MSE 18pc", m25["exigee_18pc_cotisations_n_1"])],
                              ["26 % des prestations maladie 2023 à 2025", "26 % x moyenne sur 3 ans", F("MSE 26pc", m25["exigee_26pc_prest_maladie_3ans"])],
                              ["Marge exigée", "le maximum des deux", F("MSE exigée (t)", m25["marge_exigee"])],
                              ["Réserve obligatoire", "compte 111", F("MSE 111", m25["reserve_obligatoire_111"])],
                              ["Réserves libres", "compte 1188", F("MSE 1188", m25["reserves_libres_1188"])],
                              ["Réserve disponible", "compte 106", F("MSE 106", m25["reserve_disponible_106"])],
                              ["Marge disponible comme le FPPN", "111 + 1188 + 106", F("MSE dispo (t)", m25["marge_disponible_comme_fppn"])],
                              ["Pertes reportées et résultat", "129, 121, résultat 2025", F("MSE pertes", m25["marge_disponible_nette"] - m25["marge_disponible_comme_fppn"])],
                              ["Marge disponible nette", "fonds propres", F("MSE nette (t)", m25["marge_disponible_nette"])]],
                   "largeurs": [6.6, 7.0, 5.0], "rouge": [(2, 2), (6, 2), (8, 2)],
                   "explications": ["La marge exigée vient de la règle des 18 % des cotisations, supérieure à celle des prestations.", "Sans déduire les pertes, la marge disponible dépasse largement la marge exigée.", "Une fois les pertes reportées déduites, la marge disponible devient négative."],
                   "fichier": "AUDIT ADENKA\\DAFC\\2024 et 2025 (balances des comptes)"},
        analyse=[f"**La conformité dépend de la lecture des pertes reportées.** Le ratio vaut {val_ratio(RT['MSE'], 2025)} comme le FPPN et {val_ratio(RT['MSE'], 2025, True)} net des pertes."],
        implication=f"Le FPPN n'a pas de fonds d'établissement, alors que la Décision exige un niveau minimal égal au tiers de la marge exigée, soit {F('Fonds établissement min 2025', m25['fonds_etablissement_minimum'])} en 2025.",
        source=SRC_COMPTA + " ; " + DEC + ", § 1.1.1 ; PCMS, fiche n° 2")
besoin6 = a25["charges_prestations"] * 6 / 12
contenu({"titre": "Calcul pas à pas : réserve obligatoire 2025", "sous_titre": "La réserve obligatoire couvre cinq mois de prestations, contre six au minimum",
         "synthese": f"{F('ROB 111 2025', a25['reserves_legales_111'])} x 12 / {F('ROB prest 2025', a25['charges_prestations'])} = {N('ROB 2025 (pas)', RT['ROB']['valeurs'][2025], 2)} mois"},
        "Formule (Décision, tableau 1, ligne 10) : réserves obligatoires x 12 / charges de prestations, norme de 6 à 9 mois ; le § 1.2.2.5.4 de l'annexe indique 9 mois. La réserve obligatoire est le compte 111 « Réserves légales ». Réserve nécessaire pour 6 mois = charges de prestations x 6 / 12.",
        pas_a_pas={"entetes": ["Élément", "Calcul", "Montant"],
                   "lignes": [["Réserve obligatoire", "compte 111", F("ROB 111 (t)", a25["reserves_legales_111"])],
                              ["Charges de prestations", "classe 60", F("ROB prest (t)", a25["charges_prestations"])],
                              ["Ratio de réserve obligatoire", "111 x 12 / 60", N("ROB ratio (t)", RT["ROB"]["valeurs"][2025], 2) + " mois"],
                              ["Réserve nécessaire pour 6 mois", "60 x 6 / 12", F("ROB besoin 6 mois", besoin6)],
                              ["Écart à combler", "besoin - réserve", F("ROB écart", besoin6 - a25["reserves_legales_111"])]],
                   "largeurs": [6.8, 5.2, 6.2], "rouge": [(2, 2), (4, 2)],
                   "explications": ["La réserve obligatoire couvre un peu plus de cinq mois de prestations.", "Il manque ce montant pour atteindre la borne basse de six mois."],
                   "fichier": "AUDIT ADENKA\\DAFC\\2025\\BALANCE FPPN 2025.xlsx"},
        analyse=["**La réserve progresse, mais moins vite que les prestations.** Elle vaut zéro jusqu'en 2022 et couvre " + N("ROB 2023 (a)", RT["ROB"]["valeurs"][2023], 2) + " mois en 2023, puis " + N("ROB 2025 (a)", RT["ROB"]["valeurs"][2025], 2) + " mois en 2025."],
        implication="Comme les prestations augmentent chaque année, le besoin de réserve augmente aussi : la dotation annuelle doit suivre la sinistralité et non seulement le résultat.",
        source=SRC_COMPTA + " ; " + DEC + ", tableau 1 et § 1.2.2.5.4")
contenu({"titre": "Calcul pas à pas : fonds propres et assimilés 2025", "sous_titre": "Les ressources propres et les provisions couvrent un cinquième des dettes",
         "synthese": f"Fonds propres et assimilés {F('FEP num 2025', RT['FEP']['num'][2025])} / dettes {F('FEP den 2025', RT['FEP']['den'][2025])} = {N('FEP 2025 (pas)', RT['FEP']['valeurs'][2025], 2)}"},
        "Formule (Décision, § 1.2.2.5.3) : fonds propres et assimilés (comptes 10 à 14, 190 et 15) / dettes, provisions techniques comprises. Le texte place les provisions techniques à la fois au numérateur et au dénominateur ; Adenka applique la formule à la lettre. Les dettes comprennent toute la classe 4 créditrice, dont l'épargne, les emprunts et les soldes bancaires créditeurs.",
        pas_a_pas={"entetes": ["Élément", "Comptes", "Montant 2025"],
                   "lignes": [["Fonds propres", "10, 11, 12, résultat", F("FEP FP", a25["fonds_propres"])],
                              ["Provisions techniques", "15801000, 15802000", F("FEP prov", a25["provisions_techniques"])],
                              ["Fonds propres et assimilés", "numérateur", F("FEP num (t)", RT["FEP"]["num"][2025])],
                              ["Dettes à court terme", "classe 4 créditrice, 5 créditeur", F("FEP dettes CT", a25["dettes_ct_pcms"])],
                              ["dont épargne PEP et EFC", "41195100, 41195200", F("FEP épargne", a25["epargne"])],
                              ["Provisions et autres dettes longues", "15, 16, 196", F("FEP autres", RT["FEP"]["den"][2025] - a25["dettes_ct_pcms"])],
                              ["Dettes, provisions comprises", "dénominateur", F("FEP den (t)", RT["FEP"]["den"][2025])],
                              ["Ratio des fonds propres et assimilés", "numérateur / dénominateur", N("FEP ratio (t)", RT["FEP"]["valeurs"][2025], 2)]],
                   "largeurs": [6.8, 6.0, 5.4], "rouge": [(0, 2), (4, 2), (7, 2)],
                   "explications": ["Les fonds propres sont négatifs : seules les provisions rendent le numérateur positif.", "L'épargne représente l'essentiel des dettes.", "Le ratio reste très loin de la norme d'au moins 1."],
                   "fichier": "AUDIT ADENKA\\DAFC\\2025\\BALANCE FPPN 2025.xlsx"},
        analyse=["**Même en comptant les provisions comme des fonds propres, le FPPN reste loin de la norme.** Le ratio passe de " + N("FEP 2022 (a)", RT["FEP"]["valeurs"][2022], 2) + " en 2022 à " + N("FEP 2025 (a)", RT["FEP"]["valeurs"][2025], 2) + " en 2025."],
        implication="Ce ratio ne peut s'améliorer que par des excédents mis en réserve pendant plusieurs années : il mesure la profondeur du retard de fonds propres, plus qu'un problème de l'année.",
        source=SRC_COMPTA + " ; " + DEC + ", § 1.2.2.5.3")
contenu({"titre": "Calcul pas à pas : liquidité 2025", "sous_titre": "La trésorerie couvre un huitième des dettes à court terme, épargne comprise",
         "synthese": f"Liquidité {val_ratio(RT['LIQ'], 2025)} comme le FPPN, {val_ratio(RT['LIQ'], 2025, True)} sans l'épargne ; trésorerie à échéance {val_ratio(RT['RL2'], 2025)} et {val_ratio(RT['RL2'], 2025, True)}"},
        "Formules : liquidité = trésorerie actif / (dettes à court terme + variation des provisions techniques) (Décision, § 1.2.2.5.2) ; trésorerie à échéance = (valeurs réalisables + disponibles) / dettes à court terme (section II, § 2.2) ; trésorerie immédiate = disponibles / dettes à court terme (§ 2.3), égale ici à la liquidité car les provisions ne varient pas.",
        pas_a_pas={"entetes": ["Élément", "Comptes", "Montant 2025"],
                   "lignes": [["Trésorerie actif", "classe 5 débitrice", F("LIQ tréso", a25["tresorerie_actif"])],
                              ["Créances à court terme", "classes 3 et 4 débitrices, hors avances durables", F("LIQ créances", a25["creances_ct"])],
                              ["Dettes à court terme", "classe 4 créditrice, 5 créditeur", F("LIQ dettes", a25["dettes_ct_pcms"])],
                              ["dont épargne PEP et EFC", "41195100, 41195200", F("LIQ épargne", a25["epargne"])],
                              ["Dettes à court terme hors épargne", "lecture corrigée selon le PCMS", F("LIQ dettes hors ép", a25["dettes_ct_hors_epargne"])],
                              ["Liquidité comme le FPPN", "trésorerie / dettes", val_ratio(RT["LIQ"], 2025)],
                              ["Liquidité corrigée selon le PCMS", "trésorerie / dettes hors épargne", val_ratio(RT["LIQ"], 2025, True)]],
                   "largeurs": [6.6, 7.4, 4.8], "rouge": [(3, 2), (5, 2), (6, 2)],
                   "explications": ["L'épargne représente presque toutes les dettes à court terme du plan comptable.", "Dans cette lecture, la trésorerie couvre environ un huitième des dettes.", "Sans l'épargne, la trésorerie couvre toutes les dettes à court terme."],
                   "fichier": "AUDIT ADENKA\\DAFC\\2025\\BALANCE FPPN 2025.xlsx"},
        analyse=["**Toute la question de la liquidité se ramène à l'épargne.** Les autres dettes à court terme sont couvertes par la trésorerie en 2025."],
        implication="Le FPPN doit connaître la part de l'épargne exigible dans l'année : sans cet échéancier, la lecture prudente du plan comptable s'impose et la liquidité est non conforme.",
        source=SRC_COMPTA + " ; " + DEC + ", § 1.2.2.5.2 et section II")
contenu({"titre": "Calcul pas à pas : équilibre stable 2025", "sous_titre": "Les emplois longs dépassent les ressources longues, sauf si l'épargne est une ressource longue",
         "synthese": f"Équilibre stable {val_ratio(RT['EQS'], 2025)} comme le FPPN, {val_ratio(RT['EQS'], 2025, True)} corrigé selon le PCMS ; norme d'au plus 1"},
        "Formule (Décision, § 1.2.3.1) : emplois moyens et longs / ressources moyennes et longues. Comme le FPPN : classe 2 nette / classes 10 à 19. Corrigé selon le PCMS : on ajoute aux emplois les avances durables suivies en classe 4 (CMP, Grenier, Hôpital de la Police, projets immobiliers, créance de cession) et aux ressources l'épargne PEP et EFC.",
        pas_a_pas={"entetes": ["Élément", "Comptes", "Montant 2025"],
                   "lignes": [["Actif immobilisé net", "classe 2", F("EQS immo", a25["actif_immobilise_net"])],
                              ["Avances et créances durables", "47111100 à 47111105, 47111201, 485", F("EQS avances", a25["emplois_longs_tiers"])],
                              ["Ressources longues", "10 à 19", F("EQS ress", RT["EQS"]["den"][2025])],
                              ["Épargne PEP et EFC", "41195100, 41195200", F("EQS épargne", a25["epargne"])],
                              ["Équilibre stable comme le FPPN", "classe 2 / classe 1", val_ratio(RT["EQS"], 2025)],
                              ["Équilibre stable corrigé selon le PCMS", "(2 + avances) / (1 + épargne)", val_ratio(RT["EQS"], 2025, True)]],
                   "largeurs": [6.6, 7.4, 4.8], "rouge": [(1, 2), (4, 2), (5, 2)],
                   "explications": ["Des avances durables supérieures à l'actif immobilisé sont suivies hors de la classe 2.", "Comme le FPPN, les emplois longs valent plus d'une fois et demie les ressources longues.", "Avec l'épargne en ressource longue, l'équilibre est respecté de peu."],
                   "fichier": "AUDIT ADENKA\\DAFC\\2025\\BALANCE FPPN 2025.xlsx"},
        analyse=["**L'épargne des adhérents finance les emplois longs du FPPN.** Sans elle, les ressources longues ne couvrent pas les terrains, les bâtiments et les avances durables."],
        implication="Financer des emplois longs avec de l'épargne n'est soutenable que si cette épargne reste longtemps : c'est précisément ce que l'échéancier de l'EFC et du PEP doit démontrer.",
        source=SRC_COMPTA + " ; " + DEC + ", § 1.2.3.1")
contenu({"titre": "Calcul pas à pas : taux d'excédent net 2025", "sous_titre": "L'excédent dépasse le plafond de 5 % des cotisations, ce qui sert à reconstituer les fonds propres",
         "synthese": f"Excédent net {F('EXN num 2025', a25['resultat_net'])} / cotisations acquises {F('EXN den 2025', a25['cotisations_acquises'])} = {pc('EXN 2025 (pas)', RT['EXN']['valeurs'][2025])}"},
        "Formule (Décision, § 1.2.2.5.5) : excédent net de l'exercice / cotisations acquises, norme d'au plus 5 %. Ce plafond vise à éviter qu'une mutuelle prélève sur ses membres plus que nécessaire. Il est en tension avec les normes de fonds propres quand ceux-ci sont négatifs.",
        pas_a_pas={"entetes": ["Élément", "Comptes", "Montant 2025"],
                   "lignes": [["Produits des activités ordinaires", "classe 7", F("EXN produits", a25["produits_7"])],
                              ["Charges des activités ordinaires", "classe 6", F("EXN charges", a25["charges_6"])],
                              ["Résultat hors activités ordinaires", "classe 8", F("EXN hao", a25["resultat_hao"])],
                              ["Excédent net", "7 - 6 + 8", F("EXN excédent", a25["resultat_net"])],
                              ["Cotisations acquises", "classe 70", F("EXN cot", a25["cotisations_acquises"])],
                              ["Taux d'excédent net", "excédent / cotisations", pc("EXN taux (t)", RT["EXN"]["valeurs"][2025])]],
                   "largeurs": [6.8, 5.6, 5.6], "rouge": [(3, 2), (5, 2)],
                   "explications": ["L'excédent net 2025 est positif, mais il baisse depuis 2021.", "Le taux dépasse le plafond de 5 % fixé par la norme."],
                   "fichier": "AUDIT ADENKA\\DAFC\\2025\\BALANCE FPPN 2025.xlsx"},
        analyse=["**Ce dépassement n'est pas un défaut en soi pour le FPPN.** Ses fonds propres sont négatifs, et les excédents sont le seul moyen de les reconstituer et de doter la réserve obligatoire."],
        implication="Un plan de mise en conformité présenté à l'AIRMS doit justifier ce dépassement temporaire par l'objectif de fonds propres, avec un calendrier de retour sous le plafond.",
        source=SRC_COMPTA + " ; " + DEC + ", § 1.2.2.5.5")
contenu({"titre": "Calcul pas à pas : frais de fonctionnement 2025", "sous_titre": "Les charges de fonctionnement absorbent plus d'un tiers des cotisations",
         "synthese": f"Taux des charges de fonctionnement {pc('CHF 2025 (synth)', RT['CHF']['valeurs'][2025])} (norme 20 %), brut {pc('CHB 2025 (synth)', RT['CHB']['valeurs'][2025])} (norme 15 %), par unité prestée {pc('CUP 2025 (synth)', RT['CUP']['valeurs'][2025])} (norme 15 %)"},
        "Formules (Décision, § 1.2.1.2.1 à 1.2.1.2.4). Charges de fonctionnement : comptes 612, 613, 618, 62 à 68. Charges de fonctionnement brut : comptes 61 à 68, 691 et 697. Le compte 65180000 (prime d'assurance décès de 2020) entre dans la classe 65 et donc dans les charges de fonctionnement de 2020.",
        pas_a_pas={"entetes": ["Élément", "Comptes", "Montant 2025"],
                   "lignes": [["Charges de fonctionnement", "612, 613, 618, 62 à 68", F("CHF num", a25["charges_fonctionnement"])],
                              ["Cotisations acquises", "70", F("CHF den", a25["cotisations_acquises"])],
                              ["Taux des charges de fonctionnement", "norme : au plus 20 %", pc("CHF taux (t)", RT["CHF"]["valeurs"][2025])],
                              ["Charges de fonctionnement brut", "61 à 68, 691, 697", F("CHB num", a25["charges_fonctionnement_brut"])],
                              ["Total des produits", "liste de l'annexe", F("CHB den", a25["total_produits"])],
                              ["Taux brut", "norme : au plus 15 %", pc("CHB taux (t)", RT["CHB"]["valeurs"][2025])],
                              ["Charges de prestations", "60", F("CUP den", a25["charges_prestations"])],
                              ["Coût par unité prestée", "norme : au plus 15 %", pc("CUP taux (t)", RT["CUP"]["valeurs"][2025])],
                              ["Charges de personnel / fonctionnement brut", "66, norme : moins de 85 %", pc("CHP taux (t)", RT["CHP"]["valeurs"][2025])]],
                   "largeurs": [7.2, 6.2, 4.8], "rouge": [(2, 2), (5, 2), (7, 2)],
                   "explications": ["Les charges de fonctionnement représentent plus d'un tiers des cotisations.", "Le taux brut dépasse aussi sa norme de 15 %.", "Le FPPN dépense 63 FCFA de fonctionnement pour 100 FCFA de prestations."],
                   "fichier": "AUDIT ADENKA\\DAFC\\2025\\BALANCE FPPN 2025.xlsx"},
        analyse=["**Le personnel n'explique pas l'écart.** Les charges de personnel ne représentent qu'un quart des charges de fonctionnement brut ; les indemnités des organes, les services extérieurs et les amortissements pèsent ensemble davantage."],
        implication="Pour atteindre 20 % des cotisations de 2025, les charges de fonctionnement doivent baisser de " + F("CHF baisse nécessaire", a25["charges_fonctionnement"] - 0.20 * a25["cotisations_acquises"]) + " : l'effort doit porter sur les postes autres que le personnel.",
        source=SRC_COMPTA + " ; " + DEC + ", § 1.2.1.2.1 à 1.2.1.2.4")
trim = RT["RL4"]["den"][2025]
contenu({"titre": "Calcul pas à pas : réserves des risques longs 2025", "sous_titre": "Aucune réserve de trésorerie n'est constituée pour l'épargne, et les fonds propres sont négatifs",
         "synthese": f"Besoin d'un trimestre d'engagements et de frais : {F('RL4 besoin 2025', trim)} ; réserve de trésorerie constituée : 0 FCFA"},
        "Règlement d'exécution, art. 19 : une réserve de trésorerie ou fonds de roulement au moins égale à un trimestre d'engagements et de charges de fonctionnement, pour chaque risque long. Engagements : débits des comptes 41195100 et 41195200 (sorties d'épargne). Charges de fonctionnement : quote-part au prorata de l'épargne dans le total épargne + cotisations. Art. 21 : fonds propres au moins égaux aux prestations des risques longs des deux derniers exercices.",
        pas_a_pas={"entetes": ["Élément", "Calcul", "Montant 2025"],
                   "lignes": [["Sorties d'épargne", "débits 41195100 et 41195200", F("RL4 sorties", a25["sorties_epargne"])],
                              ["Frais attribués à l'épargne", "prorata épargne / (épargne + cotisations)", F("RL4 frais", 4 * trim - a25["sorties_epargne"])],
                              ["Besoin d'un trimestre", "(sorties + frais) / 4", F("RL4 besoin (t)", trim)],
                              ["Réserve de trésorerie constituée", "aucun compte identifié", "0 FCFA"],
                              ["Actifs réservés à l'épargne (pour comparaison)", "banques, DAT et FCP réservés", F("RL4 actifs réservés", a25["actifs_dedies_epargne"])],
                              ["Fonds propres", "10 à 14 + résultat", F("RL5 FP", a25["fonds_propres"])],
                              ["Sorties d'épargne 2024 et 2025", "art. 21", F("RL5 sorties 2 ans", RT["RL5"]["den"][2025])],
                              ["Fonds propres / sorties de deux ans", "norme : au moins 1", N("RL5 ratio (t)", RT["RL5"]["valeurs"][2025], 2)]],
                   "largeurs": [7.0, 7.0, 4.6], "rouge": [(2, 2), (3, 2), (7, 2)],
                   "explications": ["Le besoin d'un trimestre se calcule à partir des sorties réelles d'épargne.", "Aucun compte de réserve de trésorerie n'existe dans les balances.", "Les fonds propres négatifs rendent le ratio de l'article 21 négatif."],
                   "fichier": "AUDIT ADENKA\\DAFC\\2024 et 2025 (balances des comptes)"},
        analyse=["**Les actifs réservés à l'épargne dépassent le besoin d'un trimestre, mais ils ne forment pas une réserve.** Une réserve est une part des fonds propres affectée, et les fonds propres du FPPN sont négatifs."],
        implication="La constitution de cette réserve passe par l'affectation d'excédents futurs, après reconstitution des fonds propres : elle doit figurer dans le plan de mise en conformité.",
        source=SRC_COMPTA + " ; " + REX + ", art. 19 et 21")
methode_gab({"titre": "Excédent par risque : pourquoi le calculer", "sous_titre": "La norme d'affectation des excédents oblige à connaître le résultat de chaque risque",
             "synthese": "Le FPPN ne présente qu'un résultat global ; Adenka reconstitue le résultat de la maladie et du décès"},
            [("Ce que dit la norme", "La Décision n° 05/2022 impose d'affecter à la réserve obligatoire au moins 30 % de l'excédent net du risque maladie et au moins 50 % de celui des autres risques (§ 1.1.1). Le règlement d'exécution fixe une dotation de 20 à 50 % du résultat excédentaire de l'exercice précédent (art. 18)."),
             ("Ce qui nous y a conduits", "Le résultat global du FPPN est positif, mais la sinistralité maladie monte vite pendant que le décès reste très excédentaire (partie 5). Un excédent global peut donc cacher un déficit de la maladie, ce que la norme veut précisément rendre visible."),
             ("Ce que fait le FPPN", "La comptabilité sépare les cotisations et les prestations de chaque risque, mais pas les frais de fonctionnement. Adenka reprend d'abord la dotation réellement passée en réserve, telle que le FPPN l'a décidée."),
             ("Pourquoi répartir les frais", "L'excédent par risque est la différence entre les produits et toutes les charges affectées au risque (Décision, § 1.1.1). Sans clé fournie par le FPPN, Adenka répartit les frais au prorata des cotisations de chaque risque, la clé la plus neutre.")],
            "La répartition au prorata des cotisations est une hypothèse : une autre clé (nombre de dossiers, temps passé) donnerait d'autres excédents. Le FPPN est invité à fournir sa propre clé, qui remplacera celle-ci.",
            source="Source : " + DEC + ", § 1.1.1 ; " + REX + ", art. 18 ; Adenka")
contenu({"titre": "Excédent par risque", "sous_titre": "Depuis 2025, la maladie est déficitaire et le décès finance l'ensemble",
         "synthese": f"Excédent maladie 2025 : {F('Exc mal 2025', EXR[2025]['excedent_maladie'])} ; excédent décès 2025 : {F('Exc dec 2025', EXR[2025]['excedent_deces'])}"},
        "Excédent d'un risque = cotisations du risque - prestations du risque - charges de fonctionnement (612, 613, 618, 62 à 68) attribuées au prorata des cotisations. La prime d'assurance décès de 2020 (65180000) est rattachée au décès. Les autres éléments (produits financiers, transferts de charges, hors activités ordinaires) ne sont pas répartis et figurent dans la colonne « autres » du classeur des calculs.",
        graphique={"titre": fig("Excédent technique par risque, frais répartis au prorata des cotisations"), "type": "barres", "categories": cats_ex,
                   "series": [{"nom": "Maladie", "valeurs": [round(EXR[a]["excedent_maladie"]) for a in range(2020, 2027)], "couleur": BLEU},
                              {"nom": "Décès", "valeurs": [round(EXR[a]["excedent_deces"]) for a in range(2020, 2027)], "couleur": ORANGE}],
                   "format": FMT_M, "unite": "En millions de FCFA"},
        analyse=[f"**La maladie devient déficitaire en 2025, une fois ses frais pris en compte.** Son excédent passe de {F('Exc mal 2022', EXR[2022]['excedent_maladie'])} en 2022 à {F('Exc mal 2025 (t)', EXR[2025]['excedent_maladie'])} en 2025.",
                 f"**Le décès est excédentaire chaque année depuis 2022.** Il dégage {F('Exc dec 2025 (t)', EXR[2025]['excedent_deces'])} en 2025.",
                 f"**Sur les huit mois de 2026, le déficit maladie atteint {F('Exc mal 2026', EXR[2026]['excedent_maladie'])}.**"],
        implication="Le décès subventionne la maladie : avec la clé des statuts (art. 17 : 3,5 % sur 5 %, soit 70 % pour la maladie), la cotisation maladie ne couvre plus son coût complet depuis 2025 et doit être revue.",
        source=SRC_COMPTA + " ; répartition des frais par Adenka")
contenu({"titre": "Affectation des excédents", "sous_titre": "Depuis 2024, la dotation à la réserve obligatoire atteint le minimum exigé par risque",
         "synthese": "Dotation de la réserve légale (compte 111) rapportée au résultat de l'exercice précédent, puis comparée aux deux normes"},
        "Lecture : la colonne « dotation réelle » est l'augmentation du compte 111 entre deux clôtures ; elle traduit l'affectation décidée par l'assemblée générale. Le minimum du règlement d'exécution est de 20 % du résultat excédentaire de l'exercice précédent (art. 18), exigible dans un délai de trois ans après le 1er juin 2022. Le minimum par risque applique 30 % à l'excédent maladie et 50 % à l'excédent décès (Décision, § 1.1.1).",
        tableau={"titre": tab("Dotation de la réserve obligatoire et normes"), "entetes": ["Affectation en", "Résultat de l'exercice précédent", "Dotation réelle", "Taux", "Minimum de 20 %", "Minimum par risque", "Écart au minimum par risque"],
                 "largeurs": [3.2, 5.0, 4.4, 2.6, 4.4, 4.4, 4.8], "taille": 10,
                 "lignes": [[str(a), F(f"AFF res {a}", AFF[a]["resultat_n_1"]), F(f"AFF dot {a}", AFF[a]["dotation_reserve_legale"]), pc(f"AFF taux {a}", AFF[a]["taux_dotation_legale"]),
                             F(f"AFF min20 {a}", 0.20 * AFF[a]["resultat_n_1"]), F(f"AFF minrisque {a}", AFF[a]["dotation_minimale_par_risque"]), F(f"AFF écart {a}", AFF[a]["ecart_dotation"])] for a in (2023, 2024, 2025)]},
        disposition="dessous",
        analyse=[f"**La dotation passe de {pc('AFF taux 2023 (a)', AFF[2023]['taux_dotation_legale'])} du résultat en 2023 à {pc('AFF taux 2025 (a)', AFF[2025]['taux_dotation_legale'])} en 2025.** En 2023, elle restait sous les deux minimums, pendant le délai de mise en conformité de trois ans.",
                 "**En 2024, une partie des excédents a aussi servi à apurer les pertes reportées et à doter les réserves facultatives.**"],
        implication="Le FPPN respecte désormais la règle d'affectation, mais la réserve reste sous six mois de prestations : la règle d'affectation seule ne suffit pas à rattraper le retard.",
        source=SRC_COMPTA + " ; " + DEC + ", § 1.1.1 ; " + REX + ", art. 18")
contenu({"titre": "Marge de solvabilité", "sous_titre": "La marge disponible ne couvre la marge exigée que si l'on ignore les pertes reportées",
         "synthese": f"2025 : marge exigée {F('MS exigée 2025 (g)', MS[2025]['marge_exigee'])}, disponible {F('MS dispo 2025 (g)', MS[2025]['marge_disponible_comme_fppn'])} comme le FPPN, {F('MS nette 2025 (g)', MS[2025]['marge_disponible_nette'])} nette des pertes"},
        "Marge exigée = le maximum entre 18 % des cotisations acquises de l'exercice précédent et de 26 % de la moyenne des prestations maladie des trois derniers exercices (Décision, § 1.1.1). L'exercice 2020 n'est pas calculé, faute de balances 2018 et 2019.",
        graphique={"titre": fig("Marge de solvabilité exigée et disponible"), "type": "barres", "categories": [str(a) if a < 2026 else "31/08/2026" for a in range(2021, 2027)],
                   "series": [{"nom": "Marge exigée", "valeurs": [round(MS[a]["marge_exigee"]) for a in range(2021, 2027)], "couleur": GRIS_C},
                              {"nom": "Disponible comme le FPPN", "valeurs": [round(MS[a]["marge_disponible_comme_fppn"]) for a in range(2021, 2027)], "couleur": BLEU},
                              {"nom": "Disponible nette des pertes", "valeurs": [round(MS[a]["marge_disponible_nette"]) for a in range(2021, 2027)], "couleur": ORANGE}],
                   "format": FMT_M, "unite": "En millions de FCFA"},
        analyse=[f"**La marge disponible comme le FPPN repose surtout sur le compte 10620000** ({F('MS 106 (g)', MS[2025]['reserve_disponible_106'])}). Le chef comptable l'explique par la réévaluation du terrain du siège, ce que son intitulé de 2020 à 2025 (« Réévaluation matériels ») ne contredit pas.",
                 "**Nette des pertes reportées, la marge disponible est négative chaque année**, même si elle se rapproche de zéro jusqu'en 2025.",
                 f"**La marge exigée augmente avec les cotisations** : elle passe de {F('MS exigée 2021 (g)', MS[2021]['marge_exigee'])} en 2021 à {F('MS exigée 2026 (g)', MS[2026]['marge_exigee'])} en 2026."],
        implication="Un écart de réévaluation ne compte pas dans la marge de solvabilité : la lecture corrigée selon le PCMS, qui l'exclut, devient la lecture de référence, et le FPPN est non conforme. Le rapport de réévaluation est demandé au FPPN (PBC, item 86).",
        source=SRC_COMPTA + " ; " + DEC + ", § 1.1.1 ; PCMS, fiche n° 1")
# ---- v11 : dispositif actuariel (qui calcule quoi, grille de conformité)
contenu({"titre": "Dispositif actuariel : qui calcule quoi", "sous_titre": "Aucun calcul actuariel n'a de responsable désigné ni de procédure écrite",
         "synthese": "Le service comptable passe les écritures, mais aucun calcul de provision, d'intérêts ou de ratio n'est documenté"},
        "Pour chaque calcul exigé par les textes, le tableau indique qui le fait aujourd'hui au FPPN, sur quelle base, et ce que la procédure devrait prévoir.",
        tableau={"titre": tab("Calculs actuariels : situation actuelle et organisation cible"), "entetes": ["Calcul", "Texte", "Situation actuelle", "Organisation cible"],
                 "largeurs": [5.6, 6.0, 9.6, 10.0], "taille": 8.5,
                 "lignes": [["Provisions maladie et décès", "Statuts, art. 89. " + REX + ", art. 10", "Montants hérités, jamais recalculés", "Note de calcul annuelle du service comptable, signée par le directeur général, présentée à l'Assemblée générale"],
                            ["Allocations décès restant à payer", "Statuts, art. 89", "État de fin d'année au barème, comptabilisé au 41591000 (2024 et 2025)", "Même état, avec date de dépôt de chaque demande pour écarter les dossiers forclos"],
                            ["Intérêts servis sur l'épargne", "Manuel des produits d'épargne", "Aucun calcul, aucune écriture", "Taux arrêté par le conseil d'administration, relevé annuel par épargnant"],
                            ["Provision du personnel (196)", "Décret n° 2017-210, art. 4 et 5", "Reprise et dotation annuelles, méthode non transmise", "Calcul agent par agent avec âge de départ, ancienneté et plafond conventionnel documentés"],
                            ["Ratios de la Décision n° 05/2022", DEC, "[[Étude des ratios du FPPN, sans provisions]]", "Tableau de bord semestriel selon les formules du texte"],
                            ["Étude des risques longs", REX + ", art. 14", "Aucune étude transmise", "Étude actuarielle tous les cinq ans"]]},
        disposition="dessous",
        analyse=["**Les calculs existent en partie, mais ils dépendent de personnes, pas de procédures.** Aucun ne fait l'objet d'un contrôle par une seconde personne ni d'une validation par les organes."],
        implication="La recommandation 9 organise ces calculs : un responsable, une procédure écrite et un contrôle pour chacun.",
        valable="situation constatée sur les documents reçus jusqu'au 22/10/2026 et les entretiens des 16 et 19/10/2026",
        source="Source : analyses Adenka ; statuts du FPPN, art. 89 ; " + REX + " ; " + DEC + " ; ETUDE DES RATIO MALADIE DECES 2.xlsx")
_GRILLE = [["Provisions techniques calculées dossier par dossier", REX + ", art. 7 à 10", "Non conforme", ref("Conclusion : suffisance des provisions")],
           ["Provision mathématique des risques longs", REX + ", art. 11", "Non conforme", ref("Provisions attendues par produit")],
           ["Comptes certifiés (risques longs)", REX + ", art. 12", "Non conforme", ref("Exigences non chiffrées")],
           ["États financiers par risque long", REX + ", art. 13", "Non conforme", ref("Exigences non chiffrées")],
           ["Étude actuarielle quinquennale", REX + ", art. 14", "Non conforme", ref("Exigences non chiffrées")],
           ["Placements sûrs, rentables, liquides, diversifiés", REX + ", art. 17", "Non conforme", ref("Conclusion : adéquation des placements")],
           ["Réserve de trésorerie des risques longs", REX + ", art. 19", "Non conforme", ref("Calcul pas à pas : réserves des risques longs 2025")],
           ["Comptabilité séparée de l'activité sanitaire", REX + ", art. 27", "Non conforme", ref("Clinique, projet de Cocody et désignation du directeur général")],
           ["Affectation des excédents à la réserve", DEC + ", § 1.1.1", "Conforme", ref("Affectation des excédents")],
           ["Adhésion au fonds national de garantie", "Règlement n° 07/2009, art. 63", "Non démontrée", ref("Tableau de bord (4/4) : données en attente")]]
contenu({"titre": "Dispositif actuariel : grille de conformité", "sous_titre": "Une seule exigence actuarielle des textes est respectée",
         "synthese": "Grille établie sur les pièces reçues ; une exigence sans pièce est classée « Non démontrée »"},
        "La grille reprend les exigences actuarielles du règlement d'exécution n° 004/2022, de la Décision n° 05/2022 et du règlement n° 07/2009, avec la diapositive qui établit le statut.",
        tableau={"titre": tab("Grille de conformité des exigences actuarielles"), "entetes": ["Exigence", "Texte", "Statut", "Détail"], "largeurs": [10.4, 8.6, 4.2, 8.0], "taille": 8.5, "lignes": _GRILLE},
        disposition="dessous",
        analyse=["**Les exigences propres aux risques longs (EFC et PEP) ne sont pas appliquées.** Le FPPN gère pourtant " + F("GR épargne", b[2026]["epargne_totale"]) + " d'épargne au 31/08/2026."],
        implication="Le plan de mise en conformité (recommandation 4) doit couvrir ces exigences, en plus des ratios chiffrés.",
        valable="exigences des textes applicables au 31/08/2026 ; pièces reçues jusqu'au 22/10/2026",
        source="Source : " + REX + " ; " + DEC + " ; règlement n° 07/2009/CM/UEMOA ; analyses Adenka")

contenu({"titre": "Contrôle du plan comptable", "sous_titre": "Deux écarts au plan comptable touchent directement les calculs actuariels",
         "synthese": "Contrôle limité aux comptes qui alimentent les provisions, les réserves, l'épargne, les cotisations, les prestations et les placements"},
        "Le contrôle ne constitue pas une revue comptable complète : il vérifie seulement que les comptes utilisés dans ce rapport sont rangés comme le prévoit le PCMS, et mesure l'effet d'un écart sur les ratios.",
        tableau={"titre": tab("Contrôle du classement des comptes selon le PCMS"), "entetes": ["Point", "Attendu selon le PCMS", "Constat", "Incidence"], "largeurs": [4.6, 8.2, 9.2, 9.2], "taille": 9,
                 "lignes": [[p["point"], p["attendu_pcms"], p["constat"], p["incidence"]] for p in Q["pcms"]]},
        disposition="dessous",
        analyse=["**Le reclassement des provisions techniques en provisions réglementées depuis 2024 et l'usage du compte 106 sont les deux écarts qui changent la lecture du bilan.** Les autres comptes suivent le plan comptable."],
        source="Source : PCMS (mai 2012), liste des comptes et fiches n° 1, 2, 6 et 25 ; balances du FPPN")
contenu({"titre": "Exigences non chiffrées", "sous_titre": "Plusieurs obligations du règlement d'exécution ne se mesurent pas par un ratio",
         "synthese": "Chaque obligation est comparée aux pièces reçues ; une obligation sans pièce est non démontrée"},
        "Ces exigences complètent les ratios. Elles relèvent du cadre actuariel quand elles portent sur les provisions, les risques longs, les placements et les réserves. Elles sont reprises dans les guides d'entretien.",
        tableau={"titre": tab("Exigences non chiffrées et situation du FPPN"), "entetes": ["Exigence", "Texte", "Situation constatée"], "largeurs": [9.6, 6.2, 15.4], "taille": 9,
                 "lignes": [["Provisions techniques suffisantes à tout moment, calculées dossier par dossier", "Règl. d'exécution, art. 7 à 11", "Montant identique de 2020 à 2026 : la provision n'est pas recalculée (partie 9)."],
                            ["Comptes certifiés par un commissaire aux comptes (risques longs)", "Règl. d'exécution, art. 12", "Aucun rapport de commissaire aux comptes n'a été transmis."],
                            ["États financiers propres à chaque risque long", "Règl. d'exécution, art. 13", "Les balances ne séparent pas l'EFC et le PEP en états distincts."],
                            ["Étude actuarielle tous les cinq ans (risques longs)", "Règl. d'exécution, art. 14", "Aucune étude antérieure n'a été transmise ; la présente mission peut y contribuer."],
                            ["Cantonnement des patrimoines entre risques courts et longs", "Règl. d'exécution, art. 15", "Des comptes bancaires réservés existent, mais ils couvrent seulement 16,28 % de l'épargne."],
                            ["Placements sûrs, rentables, liquides et diversifiés", "Règl. d'exécution, art. 17", "Aucune politique de placement. Aucune des quatre conditions n'est démontrée (partie 9.3)."],
                            ["Pertes des activités sanitaires non couvertes par les cotisations", "Règl. d'exécution, art. 27", "Le FPPN finance des charges de fonctionnement « au CMP » sans comptabilité séparée transmise."],
                            ["Adhésion et cotisation au fonds national de garantie", "Règl. n° 07/2009, art. 63", "Aucun compte du fonds de garantie n'apparaît dans les balances."]]},
        disposition="dessous",
        analyse=["**Les exigences propres aux risques longs ne sont pas documentées.** Elles découlent pourtant directement de la gestion de l'EFC et du PEP."],
        source="Source : " + REX + " ; " + R07 + " ; documents reçus le 15/10/2026")
contenu({"titre": "Exposition aux sanctions", "sous_titre": "Le non-respect des règles prudentielles expose le FPPN à des sanctions graduées",
         "synthese": "Barème de l'arrêté n° 00115/MEPS/CAB du 12/12/2023 pour une mutuelle de plus de 10 000 000 FCFA de cotisations"},
        "Les sanctions sont prononcées par le ministre chargé de la mutualité sociale, sur rapport ou avis motivé de l'AIRMS, après que la mutuelle a été entendue (règlement d'exécution, art. 38 et 39). Ce tableau décrit une exposition théorique : la décision appartient à l'autorité de contrôle.",
        tableau={"titre": tab("Manquements liés aux constats du rapport et sanctions prévues"), "entetes": ["Manquement", "Ligne de l'arrêté", "Sanction", "Peine pécuniaire"], "largeurs": [10.0, 3.4, 10.2, 7.6], "taille": 10,
                 "lignes": [["Non-respect des règles comptables et prudentielles durant un exercice", "6", "Avertissement, injonction de régularisation, suspension des dirigeants ou suspension provisoire de l'agrément", "1 000 000 FCFA"],
                            ["Défaut d'adhésion ou de versement au fonds national de garantie", "8", "Mêmes sanctions que la ligne 6", "1 000 000 FCFA"],
                            ["Non-respect des règles prudentielles sur deux exercices", "13", "Retrait de l'agrément (règlement n° 07/2009, art. 84)", "3 000 000 FCFA"]]},
        disposition="dessous",
        analyse=["**Les peines pécuniaires sont faibles ; la sanction lourde est le retrait d'agrément.** Celui-ci emporte la dissolution de la mutuelle dès sa publication au Journal officiel (règlement n° 07/2009, art. 84)."],
        implication="Un plan de mise en conformité daté et chiffré, présenté à l'AIRMS, est le meilleur moyen de démontrer la volonté de régularisation et d'éviter les sanctions les plus lourdes.",
        source="Source : arrêté n° 00115/MEPS/CAB du 12/12/2023, annexe ; " + R07 + ", art. 83 et 84 ; " + REX + ", art. 38 à 41")
retenir("conformité prudentielle", ("Les normes de solidité financière de l'UEMOA ne sont pas respectées", f"{NB_NC} ratios non conformes sur 36, souvent plusieurs années de suite",
                                    "Le FPPN respecte les ratios d'activité courante, mais pas ceux de réserves, de fonds propres, de liquidité et de frais."),
        f"La réserve obligatoire couvre {N('ROB 2025 (ret)', RT['ROB']['valeurs'][2025], 2)} mois de prestations, les fonds propres et assimilés {N('FEP 2025 (ret)', RT['FEP']['valeurs'][2025], 2)} fois les dettes, et les frais {pc('CHF 2025 (ret)', RT['CHF']['valeurs'][2025])} des cotisations.",
        "Des non-conformités répétées exposent le FPPN au retrait d'agrément prévu par l'article 84 du règlement n° 07/2009.",
        "Le FPPN doit présenter à l'AIRMS un plan de mise en conformité daté, et documenter le compte 106, le fonds de garantie et l'échéancier de l'épargne.",
        "Les ratios sont calculés sur les balances : leur conclusion ne dépend pas des données manquantes, sauf pour la liquidité, qui dépend de l'échéancier de l'épargne.")

# ---- v8 : analyses de sensibilité et leviers de conformité (classeur à formules recalculé, onglets 13a à 13i)
_SW = _ox.load_workbook(SENS, data_only=True)
K8 = json.load(open(CARTE8, encoding="utf-8"))
A8 = list(range(2020, 2027))
CODES8 = K8["CODES"]
NOM8 = {"SIN": "Sinistralité", "ROB": "Réserve obligatoire", "FEP": "Fonds propres et assimilés", "LIQ": "Liquidité", "EQS": "Équilibre stable",
        "EXN": "Excédent net", "RL1": "Solvabilité de l'épargne", "RL2": "Trésorerie à échéance", "RL3": "Trésorerie immédiate", "RL4": "Réserve de trésorerie",
        "RL5": "Fonds propres / prestations longues", "CHF": "Charges de fonctionnement", "CHB": "Charges de fonctionnement brut", "CUP": "Coût par unité prestée"}
NORME8 = {"SIN": "70,00 % à 80,00 %", "ROB": "6 à 9 mois", "FEP": "Au moins 1", "LIQ": "Au moins 1", "EQS": "Au plus 1", "EXN": "0,00 % à 5,00 %",
          "RL1": "Au moins 1", "RL2": "Plus de 1", "RL3": "Au moins 1", "RL4": "Au moins 1 trimestre", "RL5": "Au moins 1", "CHF": "Au plus 20,00 %",
          "CHB": "Au plus 15,00 %", "CUP": "Au plus 15,00 %"}
PCT8 = {"SIN", "EXN", "CHF", "CHB", "CUP"}
SRC_SENS = "Source : classeur des calculs v8, onglets 13a à 13i (formules) ; " + DEC + " ; balances du FPPN ; calculs Adenka"


def s13(onglet, r, c):
    return _SW[onglet].cell(r, c).value


def d13(code, cle, ex):
    return s13("13d_ecarts", K8["D"][code][cle], 3 + A8.index(ex))


def q13(code, c):
    return s13("13d_ecarts", K8["D"][code]["val"], c)


def f13(cle, ex):
    return s13("13f_leviers", K8["EQR"][cle], 4 + A8.index(ex))


def h13(r, ex):
    return s13("13h_synthese", r, 2 + A8.index(ex))


def v8(code, x, lab):
    if x is None or x == "":
        return "-"
    return pc(lab, x) if code in PCT8 else N(lab, x, 2)


N_STRUCT = sum(1 for c in CODES8 if q13(c, 16) == "Structurel")
FJ = K8["FJ"]
ANREF = 2025

contenu({"titre": "Analyses de sensibilité : méthode", "sous_titre": "Huit étapes, chacune appuyée sur un texte ou une norme",
         "synthese": "Les montants à combler viennent de la formule du texte inversée ; les leviers viennent des statuts et des textes UEMOA ; tout est recalculé par formules dans le classeur"},
        "Une analyse de sensibilité mesure comment un résultat change quand une hypothèse change. Ici, le résultat est le statut de chaque ratio prudentiel, et les hypothèses sont les leviers que les textes donnent au FPPN. Chaque calcul est une formule du classeur v8 (onglets 13a à 13i) : modifier un paramètre de l'onglet 13b met à jour l'ensemble.",
        tableau={"titre": tab("Étapes des analyses de sensibilité et références"), "entetes": ["Étape", "Ce qui est fait", "Référence"], "largeurs": [4.2, 18.8, 8.2], "taille": 9,
                 "lignes": [["1. Données", "Numérateurs et dénominateurs des ratios, recalculés depuis les balances de 2020 au 31/08/2026", DEC + ", tableaux 1 et 2"],
                            ["2. Statut", "Conforme si la valeur respecte la borne du texte ; un excédent ou des ressources longues négatifs sont non conformes", DEC],
                            ["3. Qualification", "Structurel : non conforme toutes les années mesurées. Récurrent : au moins deux exercices de 2022 à 2025", R07[0].upper() + R07[1:] + ", art. 83 et 84 ; arrêté n° 00115/MEPS/CAB, ligne 13"],
                            ["4. Montant à combler", "Montant qui ramène le ratio sur la borne la plus proche, les autres termes restant inchangés", DEC + " (formule inversée)"],
                            ["5. Facteurs récurrents", "Regroupement des ratios par comptes qui expliquent l'écart chaque année", "Balances du FPPN ; PCMS"],
                            ["6. Leviers", "Montant d'équilibre de chaque levier, et effet d'un choc de 100 000 000 FCFA enregistré en partie double, une hypothèse à la fois", "Statuts, art. 17, 79, 81, 85, 90 et 93 ; " + REX + ", art. 16, 17, 19 et 21 ; IAS 19, § 145"],
                            ["7. Faisabilité conjointe", "Démonstration algébrique : les normes de flux peuvent-elles être respectées ensemble ?", DEC + ", tableau 1, lignes 2 et 5 ; tableau 2, lignes 5 et 10"],
                            ["8. Inflation et rendement", "Sous-indices de prix de la santé pondérés par les prestations du FPPN ; rendement comparé aux taux du marché", "ANStat (IHPC) ; BCEAO, Rapport sur la Politique Monétaire, mars 2026, tableaux 12, 17 et 22 ; Investing.com"]]},
        disposition="dessous",
        analyse=["**La démarche suit la norme internationale de pratique actuarielle ISAP 1 (§ 2.7.7) : l'actuaire examine la sensibilité des résultats aux variations des hypothèses clés.** Les normes IAS 19 (§ 145) et IFRS 17 (§ 128) fixent la présentation : une hypothèse à la fois, changements raisonnablement possibles, méthode et limites décrites."],
        implication="Les normes ISAP 1, IAS 19 et IFRS 17 sont des références de bonne pratique, retenues pour la méthode ; elles ne sont pas opposables au FPPN. Les textes UEMOA et les statuts, eux, le sont.",
        source="Source : Association Actuarielle Internationale, ISAP 1 (adoptée le 01/12/2018), § 2.7.7 ; IAS 19 révisée, § 145 ; IFRS 17, § 128 et 129 ; IFRS 7, § 40 ; textes cités ; classeur v8, onglet 13a")

def _t96(c, a):
    b_ = {"Conforme": "🟢", "Non conforme": "🔴"}.get(d13(c, "st", a), "")
    if c in RT and RT[c]["valeurs"].get(a) is not None:
        return bv(val_ratio(RT[c], a), b_)
    return b_ or "-"


contenu({"titre": "Ratios non conformes de 2020 au 31/08/2026", "sous_titre": f"{N_STRUCT} ratios sur 14 sont non conformes toutes les années mesurées",
         "synthese": "Les écarts ne sont pas accidentels : ils se répètent chaque année, ce qui les rend structurels au sens des textes"},
        "Statut recalculé par formule pour chaque année (onglet 13d) ; il est identique à celui de la partie 10. Chaque case donne la valeur du ratio et une boule : 🔴 non conforme, 🟢 conforme ; « - » : non mesurable (pas d'exercice précédent). 2020 et 2021 sont antérieurs à la Décision du 1er juin 2022 ; 2026 couvre huit mois : ces années sont indicatives. Les ratios conformes chaque année (marge de solvabilité dans la lecture du FPPN, charges de personnel, formation) ne sont pas repris.",
        tableau={"titre": tab("Conformité de chaque ratio, année par année"), "entetes": ["Ratio", "Norme"] + [str(a) if a < 2026 else "31/08/2026" for a in A8] + ["Qualification"], "largeurs": [7.0, 4.4] + [2.3] * 7 + [3.7], "taille": 8.5,
                 "legende": LEG_CONF,
                 "lignes": [[NOM8[c], NORME8[c]] + [_t96(c, a) for a in A8] + [q13(c, 16)] for c in CODES8]},
        disposition="dessous",
        analyse=[f"**{N_STRUCT} ratios sont non conformes toutes les années mesurées, et la sinistralité l'est de 2020 à 2025.** Seule la sinistralité de 2026 respecte sa norme, sur huit mois.",
                 "**Le règlement n° 07/2009 (art. 84) prévoit le retrait d'agrément pour un non-respect des règles prudentielles sur deux exercices.** Tous ces ratios dépassent ce seuil sur la période 2022 à 2025."],
        implication="Un écart structurel ne se corrige pas par un ajustement d'une année : il faut traiter les causes qui se répètent. Les diapositives suivantes les identifient et chiffrent ce qu'il faudrait changer.",
        source=SRC_SENS + " ; " + R07 + ", art. 84")

_pr0 = 5 + len(K8["FACT"]) + 4
_PR = {lib: _pr0 + i for i, (_f, lib) in enumerate(K8["PREUVES"])}


def p13(lib, ex):
    return s13("13e_facteurs", _PR[lib], 3 + A8.index(ex))


_FN = {f[0]: f for f in K8["FACT"]}
_ST = lambda f: s13("13e_facteurs", K8["FR"][f], 4 + len(CODES8))  # noqa: E731
contenu({"titre": "Cinq facteurs récurrents expliquent les écarts", "sous_titre": "Chaque facteur se retrouve dans les comptes de chacune des sept années",
         "synthese": "L'épargne non adossée, les pertes reportées, la réserve tardive, des frais au-delà des plafonds et des prestations faibles"},
        "Un facteur regroupe les ratios dont l'écart s'explique par les mêmes comptes. La preuve chiffrée compare 2020 et 2025 (onglet 13e). La dernière colonne compte les ratios du facteur non conformes toutes les années mesurées.",
        tableau={"titre": tab("Facteurs récurrents, comptes en cause et preuves"), "entetes": ["Facteur", "Comptes en cause", "Ratios touchés", "Preuve 2020", "Preuve 2025", "Structurels"], "largeurs": [6.4, 6.0, 5.6, 5.0, 5.0, 3.2], "taille": 8.5,
                 "lignes": [["F1. " + _FN["F1"][1], _FN["F1"][2], ", ".join(NOM8[c] for c in _FN["F1"][3]),
                             "Épargne couverte à " + pc("S8 couv ép 2020", p13("Couverture de l'épargne par les actifs dédiés", 2020)), "Épargne couverte à " + pc("S8 couv ép 2025", p13("Couverture de l'épargne par les actifs dédiés", 2025)), _ST("F1")],
                            ["F2. " + _FN["F2"][1], _FN["F2"][2], ", ".join(NOM8[c] for c in _FN["F2"][3]),
                             F("S8 129 2020", p13("Pertes reportées (129)", 2020)), F("S8 129 2025", p13("Pertes reportées (129)", 2025)), _ST("F2")],
                            ["F3. " + _FN["F3"][1], _FN["F3"][2], NOM8["ROB"],
                             "Réserve nulle", F("S8 res 2025", p13("Réserve obligatoire (111)", 2025)) + " pour " + F("S8 6m 2025", p13("Six mois de prestations (borne basse de ROB)", 2025)) + " requis", _ST("F3")],
                            ["F4. " + _FN["F4"][1], _FN["F4"][2], ", ".join(NOM8[c] for c in _FN["F4"][3]),
                             pc("S8 CHF 2020", d13("CHF", "val", 2020)) + " des cotisations", pc("S8 CHF 2025", d13("CHF", "val", 2025)) + " des cotisations", _ST("F4")],
                            ["F5. " + _FN["F5"][1], _FN["F5"][2], ", ".join(NOM8[c] for c in _FN["F5"][3]),
                             "Prestations : " + pc("S8 PC 2020", p13("Prestations / cotisations (toutes garanties)", 2020)) + " des cotisations", "Prestations : " + pc("S8 PC 2025", p13("Prestations / cotisations (toutes garanties)", 2025)) + " des cotisations", _ST("F5")]]},
        disposition="dessous",
        analyse=["**Les facteurs F1 et F2 expliquent à eux seuls huit ratios : ce sont des problèmes de bilan, hérités et cumulés, que l'excédent d'une année ne peut pas corriger.**",
                 "**Les facteurs F4 et F5 sont des problèmes de compte de résultat** : les frais dépassent les plafonds fixés en proportion des cotisations et des prestations, et les prestations restent sous 70 % des cotisations jusqu'en 2025."],
        implication="Les leviers doivent viser les facteurs, pas les ratios un par un : un même levier agit sur plusieurs ratios, parfois en sens contraire.",
        source=SRC_SENS)

contenu({"titre": "Montant à combler par ratio", "sous_titre": "L'épargne et les fonds propres concentrent l'essentiel des écarts",
         "synthese": f"En 2025 : {F('S8 gap LIQ 2025', d13('LIQ', 'gap', 2025))} de trésorerie et {F('S8 gap FEP 2025', d13('FEP', 'gap', 2025))} de fonds propres manquent pour atteindre la norme"},
        "Montant à combler = montant qui ramène le ratio sur la borne la plus proche, les autres termes restant inchangés (formule du texte inversée, onglet 13d). Pour la sinistralité, un montant positif est un supplément de prestations ; pour l'excédent net de 2026, le montant est le déficit à combler.",
        tableau={"titre": tab("Montant à combler par ratio (FCFA)"), "entetes": ["Ratio", "Variable d'ajustement", "2020", "2025", "31/08/2026", "Moyenne 2020-2026"], "largeurs": [5.6, 8.6, 4.2, 4.2, 4.2, 4.4], "taille": 8.5,
                 "lignes": [[NOM8[c], s13("13d_ecarts", K8["D"][c]["gap"], 2).split(" : ", 1)[1].capitalize(),
                             "-" if d13(c, "gap", 2020) in (None, "") else F(f"S8 gap {c} 2020", d13(c, "gap", 2020)),
                             F(f"S8 gap {c} 2025 (t)", d13(c, "gap", 2025)), F(f"S8 gap {c} 2026", d13(c, "gap", 2026)), F(f"S8 gap {c} moy", q13(c, 15))] for c in CODES8]},
        disposition="dessous",
        analyse=[f"**Les écarts de bilan se comptent en milliards : {F('S8 gap RL1 2025', d13('RL1', 'gap', 2025))} d'actifs à affecter à l'épargne en 2025, et {F('S8 gap FEP 2025 (b)', d13('FEP', 'gap', 2025))} de fonds propres.**",
                 f"**Les écarts de compte de résultat se comptent en centaines de millions : {F('S8 gap CHF 2025', d13('CHF', 'gap', 2025))} de charges de fonctionnement à réduire pour le taux des charges en 2025.**"],
        implication="Les écarts de bilan ne peuvent pas être comblés par les seuls excédents annuels : ils demandent des décisions de structure (adossement de l'épargne, apurement des pertes, politique de réserves).",
        source=SRC_SENS)

_SENS = K8["SENS"]
_CHOCS = [n for n in _SENS]
_SY = {"Améliore": "+", "Dégrade": "-", "Sans effet": "0", "Sans effet sur l'écart": "0"}
contenu({"titre": "Effet de chaque levier sur chaque ratio", "sous_titre": "Aucun levier n'améliore tous les ratios ; plusieurs en dégradent certains",
         "synthese": "Choc de 100 000 000 FCFA sur l'exercice 2025, enregistré en partie double, une hypothèse à la fois"},
        "« + » : le levier rapproche le ratio de sa norme ; « - » : il l'en éloigne ; « 0 » : sans effet. Écritures : cotisations encaissées (+ produits, + résultat, + fonds propres, + trésorerie) ; prestations payées (inverse) ; charges évitées (+ résultat, + trésorerie) ; dotation à la réserve légale (affectation interne aux fonds propres) ; réallocation d'un placement long en disponibilités dédiées à l'épargne ; apport en fonds propres (dons et legs). Le montant de 100 000 000 FCFA est une convention de présentation, modifiable dans l'onglet 13b.",
        tableau={"titre": tab("Sens de l'effet d'un choc de 100 000 000 FCFA (exercice 2025)"), "entetes": ["Levier"] + [NOM8[c] for c in CODES8], "largeurs": [5.4] + [1.84] * 14, "taille": 7,
                 "lignes": [[{"L1 cotisations +": "Hausse des cotisations", "L2 prestations +": "Hausse des prestations", "L3 charges -": "Baisse des charges", "L4 réserve légale +": "Dotation à la réserve", "L5 réallocation": "Réallocation des placements", "L7 fonds propres +": "Apport en fonds propres"}[n]]
                            + [_SY.get(s13("13h_synthese", _SENS[n], 2 + j), "0") for j in range(len(CODES8))] for n in _CHOCS]},
        disposition="dessous",
        analyse=["**La hausse des cotisations améliore les frais et la liquidité, mais éloigne la sinistralité et l'excédent net de leur norme.** La hausse des prestations fait l'inverse.",
                 "**La dotation à la réserve, la réallocation des placements et l'apport en fonds propres n'ont aucun effet défavorable ; la baisse des charges n'éloigne que l'excédent net de sa norme.**"],
        implication="Le plan de redressement prévu par les statuts (art. 93) doit combiner les leviers : agir sur les cotisations seules ne peut pas rendre le FPPN conforme.",
        source=SRC_SENS + " ; statuts, art. 79, 81, 85, 90 et 93")

_fe = lambda k: f13(k, ANREF)  # noqa: E731
_cf25, _pre25, _emp25, _res25 = AG[2025]["charges_fonctionnement"], RT["SIN"]["num"][2025], RT["EQS"]["num"][2025], RT["ROB"]["num"][2025]
_moyRN = s13("13f_leviers", K8["EQR"]["L7|MOY"], 4)
contenu({"titre": "Montants d'équilibre de chaque levier en 2025", "sous_titre": "Sur quoi agir, et de combien, pour rendre chaque ratio conforme",
         "synthese": f"Taux de cotisation : {pc('S8 T SIN 2025', _fe('L1|T_SIN'))} pour la sinistralité, mais {pc('S8 T CHF 2025', _fe('L1|T_CHF'))} pour les frais ; réduction des frais de {pc('S8 L3 CHF pct', _fe('L3|CHF') / _cf25)} pour le taux des charges"},
        "Montant d'équilibre = montant du levier qui ramène un ratio sur sa borne en 2025, toutes choses égales par ailleurs (onglet 13f). Signe négatif : baisse. La dernière colonne compte les années de 2020 au 31/08/2026 où un effort est nécessaire. Taux de cotisation équivalent = 5 % × (cotisations + variation) / cotisations, à assiette et recouvrement inchangés.",
        tableau={"titre": tab("Montants d'équilibre 2025 par levier et par ratio"), "entetes": ["Levier (texte)", "Ratio visé", "Montant 2025", "Lecture", "Années d'effort"], "largeurs": [7.2, 5.4, 5.0, 9.6, 3.0], "taille": 8.5,
                 "lignes": [["Cotisations (statuts, art. 17 et 93)", NOM8["SIN"], F("S8 L1 SIN", _fe("L1|SIN")), "Taux de cotisation ramené à " + pc("S8 T SIN (t)", _fe("L1|T_SIN")), N("S8 n L1 SIN", s13("13f_leviers", K8["EQR"]["L1|SIN"], 11))],
                            ["Cotisations (statuts, art. 17 et 93)", NOM8["CHF"], F("S8 L1 CHF", _fe("L1|CHF")), "Taux de cotisation porté à " + pc("S8 T CHF (t)", _fe("L1|T_CHF")), N("S8 n L1 CHF", s13("13f_leviers", K8["EQR"]["L1|CHF"], 11))],
                            ["Prestations (statuts, art. 93)", NOM8["SIN"], F("S8 L2 SIN", _fe("L2|SIN")), pc("S8 L2 SIN pct", _fe("L2|SIN") / _pre25) + " de prestations en plus", N("S8 n L2 SIN", s13("13f_leviers", K8["EQR"]["L2|SIN"], 11))],
                            ["Prestations (statuts, art. 93)", NOM8["EXN"], F("S8 L2 EXN", _fe("L2|EXN")), pc("S8 L2 EXN pct", _fe("L2|EXN") / _pre25) + " de prestations en plus", N("S8 n L2 EXN", s13("13f_leviers", K8["EQR"]["L2|EXN"], 11))],
                            ["Charges (Décision n° 05/2022)", NOM8["CHF"], F("S8 L3 CHF", _fe("L3|CHF")), pc("S8 L3 CHF pct (t)", _fe("L3|CHF") / _cf25) + " des charges en moins", N("S8 n L3 CHF", s13("13f_leviers", K8["EQR"]["L3|CHF"], 11))],
                            ["Charges (Décision n° 05/2022)", NOM8["CUP"], F("S8 L3 CUP", _fe("L3|CUP")), pc("S8 L3 CUP pct", _fe("L3|PCT_CUP")) + " des charges en moins", N("S8 n L3 CUP", s13("13f_leviers", K8["EQR"]["L3|CUP"], 11))],
                            ["Réserve légale (statuts, art. 90)", NOM8["ROB"], F("S8 L4 ROB", _fe("L4|ROB")), "Dotation supplémentaire, interne aux fonds propres", N("S8 n L4 ROB", s13("13f_leviers", K8["EQR"]["L4|ROB"], 11))],
                            ["Placements (statuts, art. 85)", NOM8["LIQ"], F("S8 L5 LIQ", _fe("L5|LIQ")), "Supérieur à toute la classe 2 (" + F("S8 classe 2 2025", _emp25) + ")", N("S8 n L5 LIQ", s13("13f_leviers", K8["EQR"]["L5|LIQ"], 11))],
                            ["Placements (statuts, art. 85)", NOM8["RL1"], F("S8 L5 RL1", _fe("L5|RL1")), "Actifs à affecter à l'épargne", N("S8 n L5 RL1", s13("13f_leviers", K8["EQR"]["L5|RL1"], 11))],
                            ["Fonds propres (statuts, art. 79 et 81)", NOM8["FEP"], F("S8 L7 FEP", _fe("L7|FEP")), N("S8 L7 ans", _fe("L7|ANS"), 2) + " années d'excédent moyen 2023 à 2025", N("S8 n L7 FEP", s13("13f_leviers", K8["EQR"]["L7|FEP"], 11))]]},
        disposition="dessous",
        analyse=[f"**Pour la sinistralité, il faudrait baisser le taux de cotisation ou servir {F('S8 L2 SIN (b)', _fe('L2|SIN'))} de prestations en plus ; pour les frais, il faudrait au contraire presque doubler le taux.** Le levier des cotisations est donc contradictoire.",
                 f"**Le levier le plus direct est la baisse des charges de fonctionnement** : {F('S8 L3 CHF (b)', _fe('L3|CHF'))} en 2025 pour le taux des charges, davantage pour le coût par unité prestée.",
                 f"**Les écarts de bilan dépassent les moyens du FPPN** : la trésorerie manquante excède toute la classe 2, et les fonds propres demanderaient {N('S8 L7 ans (b)', _fe('L7|ANS'), 2)} années d'excédent moyen."],
        implication="Les ratios de compte de résultat peuvent être corrigés en quelques exercices par une maîtrise des frais et une revue du régime des prestations. Les ratios de bilan demandent un plan pluriannuel, présenté à l'AIRMS.",
        source=SRC_SENS + " ; statuts, art. 17, 79, 81, 85, 90 et 93")

contenu({"titre": "Cotisations : deux ratios demandent des mouvements opposés", "sous_titre": "La sinistralité appelle une baisse du taux, les frais une hausse",
         "synthese": f"En 2025, taux de cotisation équivalent : {pc('S8 T SIN 2025 (g)', _fe('L1|T_SIN'))} pour la sinistralité, {pc('S8 T CHF 2025 (g)', _fe('L1|T_CHF'))} pour le taux des charges, contre 5,00 % dans les statuts"},
        "Taux de cotisation qui ramènerait chaque ratio sur sa borne, toutes choses égales par ailleurs (onglet 13f). Taux actuel : 5 % de la solde de base (statuts, art. 17). Les courbes montrent l'écart entre les deux exigences, année par année.",
        graphique={"titre": fig("Taux de cotisation équivalent selon le ratio visé"), "type": "lignes", "categories": cats_ex,
                   "series": [{"nom": "Pour la sinistralité (70 %)", "valeurs": [round(100 * f13("L1|T_SIN", a), 2) for a in A8], "couleur": ORANGE},
                              {"nom": "Pour le taux des charges (20 %)", "valeurs": [round(100 * f13("L1|T_CHF", a), 2) for a in A8], "couleur": BLEU},
                              {"nom": "Taux des statuts", "valeurs": [5.0] * 7, "couleur": GRIS_F}],
                   "format": "0.00", "unite": "En % de la solde de base"},
        analyse=[f"**L'écart entre les deux exigences se réduit, mais reste large : de {pc('S8 T SIN 2020', f13('L1|T_SIN', 2020))} contre {pc('S8 T CHF 2020', f13('L1|T_CHF', 2020))} en 2020, à {pc('S8 T SIN 2025 (b)', _fe('L1|T_SIN'))} contre {pc('S8 T CHF 2025 (b)', _fe('L1|T_CHF'))} en 2025.**",
                 "**La hausse des prestations maladie rapproche la sinistralité de sa norme** : en 2026, sur huit mois, le taux de 5 % suffit à la respecter."],
        implication="Le taux de cotisation n'est pas le bon levier pour les frais : c'est le niveau des charges qui dépasse les plafonds du texte. Le taux devra être revu une fois les provisions recalculées, car elles entrent dans la sinistralité.",
        source=SRC_SENS + " ; statuts, art. 17")

_ANS_L = {a: f13("L7|ANS", a) for a in A8}
contenu({"titre": "Liquidité : la vente des placements longs ne suffirait pas", "sous_titre": "Même en vendant toute la classe 2, la liquidité resterait sous la norme",
         "synthese": f"Liquidité maximale théorique : {N('S8 LIQ max 2025', f13('L5|MAX', 2025), 2)} en 2025 ; liquidité hors épargne : {N('S8 LIQ r 2025', f13('L6|LIQ', 2025), 2)}, conforme"},
        "Ligne 1 : trésorerie à constituer pour une liquidité de 1. Ligne 2 : toute la classe 2 (siège, terrains, créances longues), plafond théorique d'une réallocation. Ligne 3 : liquidité si toute la classe 2 devenait disponible. Ligne 4 : liquidité calculée en classant l'épargne en ressource longue (présentation des risques longs, partie 10).",
        tableau={"titre": tab("Liquidité : écart, plafond de réallocation et lecture hors épargne"), "entetes": ["Indicateur"] + [str(a) if a < 2026 else "31/08/2026" for a in A8], "largeurs": [8.2] + [3.28] * 7, "taille": 8.5,
                 "lignes": [["Trésorerie manquante"] + [F(f"S8 L5 LIQ {a}", f13("L5|LIQ", a)) for a in A8],
                            ["Classe 2 nette"] + [F(f"S8 EMP {a}", RT["EQS"]["num"][a]) for a in A8],
                            ["Liquidité maximale théorique"] + [N(f"S8 LIQ max {a}", f13("L5|MAX", a), 2) for a in A8],
                            ["Liquidité hors épargne"] + [N(f"S8 LIQ r {a}", f13("L6|LIQ", a), 2) for a in A8],
                            ["Conforme hors épargne ?"] + [{"Conforme": "Oui", "Non conforme": "Non"}[f13("L6|LIQ_st", a)] for a in A8]]},
        disposition="dessous",
        analyse=["**La trésorerie manquante dépasse, chaque année, toute la classe 2 : la réallocation des placements, seule, ne peut pas rendre la liquidité conforme.**",
                 "**Hors épargne, la liquidité devient conforme en 2025 et au 31/08/2026.** Le problème de liquidité est donc celui de l'épargne PEP et EFC, classée en dettes à court terme."],
        implication="Le levier de liquidité est l'adossement de l'épargne : des actifs dédiés, constitués au fil des versements, et un échéancier qui justifie le classement de l'épargne en ressource longue. Sans échéancier, ce classement n'est pas démontré.",
        source=SRC_SENS + " ; " + REX + ", art. 3, 17 et 19")

contenu({"titre": "Réserve et fonds propres : la règle des statuts aurait suffi pour la réserve", "sous_titre": "Avec 20 % des excédents depuis 2020, la réserve aurait été conforme en 2025",
         "synthese": f"Réserve obligatoire 2025 : {N('S8 ROB 2025', d13('ROB', 'val', 2025), 2)} mois réels ; {N('S8 ROB20 2025', f13('L4|D20', 2025), 2)} mois avec 20 % des excédents passés ; norme de 6 à 9 mois"},
        "Réserve hypothétique = maximum entre la réserve comptabilisée et 20 % (ou 50 %) des excédents positifs cumulés depuis 2020 (statuts, art. 90). Les excédents antérieurs à 2020 n'ont pas été reçus. En réalité, une partie des excédents a apuré les pertes reportées (affectations 2023 et 2024). Fonds propres : nombre d'années d'excédent moyen 2023 à 2025 (" + F("S8 RN moy", _moyRN) + ") pour atteindre la norme.",
        tableau={"titre": tab("Réserve obligatoire selon la dotation statutaire, et horizon des fonds propres"), "entetes": ["Indicateur"] + [str(a) if a < 2026 else "31/08/2026" for a in A8], "largeurs": [8.2] + [3.28] * 7, "taille": 8.5,
                 "lignes": [["Réserve réelle (mois)"] + [N(f"S8 ROB {a}", d13("ROB", "val", a), 2) for a in A8],
                            ["Avec 20 % des excédents (mois)"] + [N(f"S8 ROB20 {a}", f13("L4|D20", a), 2) for a in A8],
                            ["Conforme avec 20 % ?"] + [{"Conforme": "Oui", "Non conforme": "Non"}[f13("L4|D20_st", a)] for a in A8],
                            ["Avec 50 % des excédents (mois)"] + [N(f"S8 ROB50 {a}", f13("L4|D50", a), 2) for a in A8],
                            ["Années d'excédent pour les fonds propres"] + [N(f"S8 ANS FEP {a}", _ANS_L[a], 2) for a in A8]]},
        disposition="dessous",
        analyse=["**Une dotation de 20 % des excédents dès 2020 aurait donné une réserve conforme en 2025 ; au-delà de 9 mois, la réserve dépasse la borne haute.** Le problème de la réserve tient à l'affectation des excédents, pas à leur montant.",
                 f"**Les fonds propres demanderaient {N('S8 ANS FEP 2025 (b)', _ANS_L[2025], 2)} années d'excédent moyen pour atteindre la norme** : l'excédent seul ne peut pas les reconstituer dans un délai raisonnable."],
        implication="La réserve obligatoire peut être mise en conformité par la seule application de l'article 90 des statuts. Les fonds propres demandent d'autres ressources (dons et legs, art. 81) ou une réduction des dettes, à inscrire dans le plan de redressement.",
        source=SRC_SENS + " ; statuts, art. 81 et 90")

contenu({"titre": "Les normes de flux ne peuvent pas être respectées ensemble", "sous_titre": "Avec les formules de la Décision, sinistralité, coût par unité prestée et excédent net se contredisent",
         "synthese": f"Il faudrait que les autres éléments du résultat soient une charge d'au moins {pc('S8 seuil A', -h13(FJ[2], 2025))} des cotisations ; ils sont un produit de {pc('S8 A 2025', h13(FJ[1], 2025))} en 2025"},
        "Démonstration (onglet 13h). L'excédent vaut : cotisations - prestations - charges de fonctionnement + autres éléments (A : produits financiers et autres, moins les autres charges). L'excédent net plafonné à 5 % impose A ≤ 5 % des cotisations - cotisations + prestations + charges. Or les prestations sont plafonnées à 80 % des cotisations (sinistralité) et les charges à 15 % des prestations, soit 12 % des cotisations (coût par unité prestée). Donc A ≤ (5 % - 100 % + 80 % × 115 %) des cotisations = -3,00 % des cotisations.",
        tableau={"titre": tab("Condition de faisabilité conjointe des normes de flux"), "entetes": ["Indicateur"] + [str(a) if a < 2026 else "31/08/2026" for a in A8], "largeurs": [8.2] + [3.28] * 7, "taille": 8.5,
                 "lignes": [["Autres éléments du résultat (A)"] + [F(f"S8 A {a}", h13(FJ[0], a)) for a in A8],
                            ["A en % des cotisations"] + [pc(f"S8 A pct {a}", h13(FJ[1], a)) for a in A8],
                            ["Seuil maximal de A"] + [pc(f"S8 seuil {a}", h13(FJ[2], a)) for a in A8],
                            ["Normes conciliables ?"] + [h13(FJ[3], a) for a in A8],
                            ["Exercices pour 6 mois de réserve à 5 % d'excédent"] + [N(f"S8 ex ROB {a}", h13(FJ[4], a), 2) for a in A8]]},
        disposition="dessous",
        analyse=["**Tant que les produits financiers et autres produits dépassent les autres charges, aucune combinaison des cotisations, des prestations et des frais ne respecte en même temps la sinistralité, le coût par unité prestée et l'excédent net.**",
                 f"**Les normes se contraignent aussi dans le temps** : avec un excédent plafonné à 5 % et la sinistralité de 2025, constituer six mois de réserve prend {N('S8 ex ROB 2025 (b)', h13(FJ[4], 2025), 2)} exercices."],
        implication="Cette tension tient aux formules du texte, pas à la gestion du FPPN. Elle doit être exposée à l'AIRMS : le plan de mise en conformité doit dire quels ratios de flux sont visés en priorité et sur quel horizon.",
        source=SRC_SENS + " ; " + DEC + ", tableau 1, lignes 2 et 5 ; tableau 2, lignes 5 et 10")

_IN = K8["INF"]
_MOY = K8["MOY"]
_LABO = {(int(r[0]), int(r[1])): r[3] for r in _SW["13g_inflation"].iter_rows(min_row=5, min_col=14, max_col=21, values_only=True) if isinstance(r[0], (int, float))}
_POIDS_LABO = sum(s13("13g_inflation", K8["MAP_r1"] + k, 4) for k, (nat, po) in enumerate(K8["MAP"]) if po == "Services de laboratoires et de radiologie")
_RET = ["Services hospitaliers", "Services de laboratoires et de radiologie", "Médicaments modernes", "Services médicaux et dentaires"]
contenu({"titre": "Inflation médicale : sous-indices de prix de la santé", "sous_titre": "Les prix des analyses et de l'imagerie ont le plus augmenté depuis 2020",
         "synthese": f"Inflation médicale pondérée par les prestations du FPPN : {pc('S8 inf moy', s13('13g_inflation', _IN['MOY'], 2))} par an en moyenne de 2021 à 2024, {pc('S8 inf max', s13('13g_inflation', _IN['MAX'], 2))} au plus (2022)"},
        "Indices harmonisés des prix à la consommation (IHPC) de l'ANStat, par poste de la santé, mensuels de janvier 2017 à octobre 2024 (moyennes annuelles ; 2024 : janvier à octobre). Poids : prestations maladie du FPPN de 2025 par nature (hospitalisations, analyses et imagerie, pharmacie, consultations) ; les natures « autres » ne sont pas détaillées et sont exclues. Les fichiers codent les mois 2, 8 et 12 par « 0 » : le mois est reconstitué par la position de la ligne (onglet 13g).",
        graphique={"titre": fig("Variation annuelle des prix de la santé retenus et inflation pondérée du FPPN"), "type": "lignes", "categories": [str(a) for a in range(2020, 2025)],
                   "series": [{"nom": p.replace("Services de laboratoires et de radiologie", "Laboratoires et radiologie"), "valeurs": [round(100 * s13("13g_inflation", K8["VAR"][p], 5 + j), 2) for j in range(5)], "couleur": col}
                              for p, col in zip(_RET, (BLEU, ORANGE, GRIS_F, GRIS_C))]
                             + [{"nom": "Inflation pondérée du FPPN", "valeurs": [round(100 * s13("13g_inflation", _IN[str(a)], 2), 2) for a in range(2020, 2025)], "couleur": NOIR}],
                   "format": "0.00", "unite": "En % ; 2024 : glissement d'octobre pour l'inflation pondérée, moyenne de janvier à octobre pour les postes"},
        analyse=[f"**Les analyses et l'imagerie, qui pèsent {pc('S8 poids labo', _POIDS_LABO)} des prestations détaillées du FPPN, ont connu un saut de prix en 2022.** L'indice passe de {N('S8 labo avr22', _LABO[(2022, 4)], 2)} en avril à {N('S8 labo juil22', _LABO[(2022, 7)], 2)} en juillet 2022.",
                 f"**En 2025, faute de série ANStat, l'inflation de la santé de l'UEMOA est négative ({pc('S8 inf 2025', s13('13g_inflation', _IN['2025'], 2))} en moyenne de trois trimestres)** selon la BCEAO."],
        implication=f"Les prix de la santé ont augmenté d'au plus {pc('S8 inf max (c)', s13('13g_inflation', _IN['MAX'], 2))} par an, alors que les prestations maladie ont été multipliées par {N('S8 mult prest', AG[2025]['prest_maladie'] / AG[2020]['prest_maladie'], 2)} de 2020 à 2025 : l'essentiel de la hausse vient du volume des soins, pas des prix. Après octobre 2024, seule la série de l'UEMOA publiée par la BCEAO est disponible : elle ne montre pas de hausse des prix de la santé en 2025.",
        source="Source : ANStat, IHPC par poste de consommation (national), extraction du 19/10/2026 ; BCEAO, Rapport sur la Politique Monétaire dans l'UMOA, mars 2026, tableau 12 ; prestations du FPPN (balances 2025) ; classeur v8, onglet 13g")

_g0 = _IN["GR0"]
_GR = [s13("13g_inflation", _g0 + k, 1) for k in range(9)]
contenu({"titre": "Sinistralité face à une hausse des prestations", "sous_titre": f"Une hausse de {pc('S8 S80 2026 (st)', s13('13g_inflation', _IN['S80'], 4))} des prestations de 2026 ferait dépasser 80 %",
         "synthese": f"Seuil de dépassement : {pc('S8 S80 2026', s13('13g_inflation', _IN['S80'], 4))} de hausse sur 2026 ; l'inflation pondérée maximale observée ({pc('S8 inf max (b)', s13('13g_inflation', _IN['MAX'], 2))}) le franchirait"},
        "Sinistralité = prestations × (1 + hausse) / cotisations, à cotisations inchangées : elles sont assises sur la solde de base, qui n'est pas indexée sur les prix de la santé (onglet 13g). 2026 : huit mois, indicatif.",
        tableau={"titre": tab("Sinistralité selon la hausse des prestations"), "entetes": ["Hausse des prestations", "Sinistralité 2025", "Conforme 2025 ?", "Sinistralité 2026", "Conforme 2026 ?"], "largeurs": [8.4, 5.4, 4.4, 5.4, 4.4], "taille": 9,
                 "lignes": [[("Inflation moyenne 2021-2024 : " if k == 7 else "Inflation maximale 2018-2024 : " if k == 8 else "") + pc(f"S8 hausse {k}", _GR[k]),
                             pc(f"S8 SIN25 {k}", s13("13g_inflation", _g0 + k, 2)), {"Conforme": "Oui", "Non conforme": "Non"}[s13("13g_inflation", _g0 + k, 3)],
                             pc(f"S8 SIN26 {k}", s13("13g_inflation", _g0 + k, 4)), {"Conforme": "Oui", "Non conforme": "Non"}[s13("13g_inflation", _g0 + k, 5)]] for k in range(9)]},
        disposition="droite",
        analyse=[f"**En 2026, la sinistralité est déjà proche de sa borne haute** : une hausse de {pc('S8 S80 2026 (b)', s13('13g_inflation', _IN['S80'], 4))} des prestations la porterait au-delà de 80 %.",
                 f"**En 2025, il faudrait au contraire {pc('S8 S70 2025', s13('13g_inflation', _IN['S70'], 2))} de prestations en plus pour atteindre 70 %.**"],
        implication="La sinistralité est passée en deux ans d'un excès de marge à une marge faible : la revue du régime des prestations doit intégrer une hypothèse d'inflation médicale, et le recalcul des provisions, qui entre dans le ratio.",
        source="Source : classeur v8, onglet 13g ; ANStat ; " + DEC + ", tableau 1, ligne 5")

_RD = K8["RDT"]
_tx = lambda nom, col: s13("13i_rendement", _RD["TX"][nom], col)  # noqa: E731
_ri = lambda cle, col: s13("13i_rendement", _RD[cle], col)  # noqa: E731
_DAT = "Dépôts à terme, Côte d'Ivoire"
_OK = {"Conforme": "Oui", "Non conforme": "Non", None: "-", "": "-"}
_A5 = list(range(2021, 2026))
_RGL = {x["exercice"]: x["rendement"] for x in V11["rendement_gl"]}
SRC_INV = "Investing.com, « Rendement de l'Obligation Cote d'Ivoire 5-Year / 10-Year - Données Historiques », export du 19/10/2026"
contenu({"titre": "Rendement des placements et taux du marché", "sous_titre": "Le rendement du FPPN est inférieur aux taux du marché de 2022 à 2024",
         "synthese": f"Obligations de l'État de Côte d'Ivoire : {pc('S8 inv5 der', _ri('INV5', 9))} à 5 ans ({_ri('INV5', 10)}), {pc('S8 inv10 der', _ri('INV10', 9))} à 10 ans ({_ri('INV10', 10)}) ; taux servi sur l'épargne : [[3,50 % selon le chef comptable]], 5,50 % selon le manuel d'épargne"},
        "Le taux brut de rentabilité des placements doit atteindre au moins le taux moyen du marché (Décision n° 05/2022, section II), sans que le texte précise le marché. Deux lectures : dépôts à terme en Côte d'Ivoire (BCEAO, tableau 22, lecture principale, 2024 et 2025) et obligations de l'État de Côte d'Ivoire à 5 ans (Investing.com, moyenne annuelle des cotations quotidiennes, lecture exigeante). Investing.com est un agrégateur de cotations du marché secondaire : source secondaire, en complément de la BCEAO. Produit possible = solde moyen non rémunéré × taux moyen des dépôts à terme (onglet 13i).",
        tableau={"titre": tab("Rendement du FPPN et taux du marché, 2021 à 2025"), "entetes": ["Indicateur"] + [str(a) for a in _A5] + ["Source"], "largeurs": [8.8] + [2.9] * 5 + [7.9], "taille": 8,
                 "lignes": [["Rendement brut des placements du FPPN"] + [pc(f"S8 RL6 {a}", _RGL[a]) for a in _A5] + ["Grands livres"],
                            ["Dépôts à terme, Côte d'Ivoire", "-", "-", "-", pc("S8 DAT 2024", _tx(_DAT, 11)), pc("S8 DAT 2025", _tx(_DAT, 12)), "BCEAO, tableau 22"],
                            ["État de Côte d'Ivoire, 5 ans"] + [pc(f"S8 inv5 {a}", _ri("INV5", 3 + j)) for j, a in enumerate(_A5)] + ["Investing.com"],
                            ["État de Côte d'Ivoire, 10 ans"] + [pc(f"S8 inv10 {a}", _ri("INV10", 3 + j)) for j, a in enumerate(_A5)] + ["Investing.com"],
                            ["Obligations de l'UMOA, 5 ans", "-", "-", "-", pc("S8 O5 2024", _tx("Obligations à 5 ans (UMOA)", 11)), pc("S8 O5 2025", _tx("Obligations à 5 ans (UMOA)", 12)), "BCEAO, tableau 17"],
                            ["Conforme face aux dépôts à terme ?", "-", "-", "-", "Oui" if _RGL[2024] >= _tx(_DAT, 11) else "Non", "Oui" if _RGL[2025] >= _tx(_DAT, 12) else "Non", "Grands livres, BCEAO"],
                            ["Conforme face à l'État à 5 ans ?"] + ["Oui" if _RGL[a] >= _ri("INV5", 3 + j) else "Non" for j, a in enumerate(_A5)] + ["Grands livres, Investing.com"],
                            ["Solde moyen non rémunéré"] + [F(f"S8 RL7 {a}", _ri("RL7", 3 + j)) for j, a in enumerate(_A5)] + ["Partie 10"],
                            ["Produit possible au taux des dépôts à terme", "-", "-", "-", F("S8 gain 2024", _ri("GAIN", 6)), F("S8 gain 2025", _ri("GAIN", 7)), "Onglet 13i"]]},
        disposition="dessous",
        analyse=["**De 2022 à 2024, le rendement des placements est inférieur à toutes les références de marché disponibles.** Celui de 2025 les dépasse seulement parce que les intérêts du DAT PEP accumulés depuis 2020 sont enregistrés à sa clôture, le 30/10/2025.",
                 f"**Le taux servi de 3,5 % sur l'épargne est inférieur de {pc('S8 écart inv5', _ri('ECARTINV', 3))} au rendement actuel des obligations de l'État à 5 ans** : un adossement de l'épargne en titres publics couvrirait le taux servi avec une marge."],
        implication=f"Placer le solde non rémunéré rapporterait {F('S8 gain 2025 (i)', _ri('GAIN', 7))} par an au taux des dépôts à terme de 2025, sans réduire la liquidité au sens du texte, car un dépôt à terme reste en classe 5. Les échéances doivent suivre l'échéancier de l'épargne, que le FPPN ne tient pas encore (recommandation 2).",
        source="Source : BCEAO, Rapport sur la Politique Monétaire dans l'UMOA, mars 2026, tableaux 17 et 22 ; " + SRC_INV + " ; classeur v8, onglet 13i ; entretien du 19/10/2026 (taux servi)")

retenir("leviers de conformité", ("Les écarts sont structurels ; seuls des leviers combinés peuvent rendre le FPPN conforme", f"{N_STRUCT} ratios non conformes toutes les années, cinq facteurs récurrents, des leviers parfois contradictoires",
                                  "Les écarts de bilan (épargne, fonds propres) dépassent les moyens annuels du FPPN ; les écarts de compte de résultat (frais, réserve) peuvent être corrigés en quelques exercices."),
        f"Trésorerie manquante de {F('S8 gap LIQ 2025 (r)', d13('LIQ', 'gap', 2025))} en 2025, supérieure à toute la classe 2 ; fonds propres manquants de {F('S8 gap FEP 2025 (r)', d13('FEP', 'gap', 2025))} ; frais à réduire de {F('S8 L3 CHF (r)', _fe('L3|CHF'))} ; une dotation de 20 % des excédents aurait suffi pour la réserve en 2025.",
        "Le taux de cotisation ne peut pas, seul, rendre le FPPN conforme : il améliore les frais mais éloigne la sinistralité et l'excédent net de leur norme. Les normes de flux ne sont pas conciliables tant que les autres produits dépassent les autres charges.",
        "Le FPPN doit proposer à l'AIRMS un plan pluriannuel : maîtrise des frais, dotation statutaire de la réserve, adossement de l'épargne sur un échéancier, ressources propres, revue du régime des prestations.",
        "Les montants sont recalculés par formules dans le classeur v8 (onglets 13a à 13i) : chaque paramètre peut être modifié pour tester d'autres hypothèses.")

# ---------------------------------------------------------------- PARTIE 11 : MORTALITÉ
separateur(11, "Tables de mortalité", "Rôle, choix des tables, indicateurs de référence")
methode_gab({"titre": "Pourquoi des tables de mortalité", "sous_titre": "Les tables servent à mesurer et à évaluer tout engagement lié à la vie humaine",
             "synthese": "Elles préparent deux analyses de la mission : la mortalité des policiers et les engagements viagers"},
            [("qx", "La table donne, pour chaque âge x, la probabilité de décéder avant l'âge x + 1. C'est la brique de base de tous les calculs."),
             ("Lx", "La table donne aussi le nombre de survivants à chaque âge, sur 1 000 000 de naissances. Le rapport de deux Lx donne une probabilité de survie."),
             ("Premier usage", "Comparer la mortalité réelle des policiers à la table de référence dit si la cotisation décès est bien calibrée."),
             ("Second usage", "Les tables permettent de calculer la valeur des engagements qui durent toute la vie, comme une prise en charge santé maintenue aux retraités.")],
            "Ces deux analyses demandent les dates de naissance (PBC, items 9 et 28). Le règlement d'exécution n° 004/2022 définit la table de mortalité comme un tableau qui donne les probabilités de décès ou de survie et l'espérance de vie selon l'âge et le sexe (art. 1).",
            source=SRC_CIMA)
contenu({"titre": "Choix des tables", "sous_titre": "Le principe de prudence fixe la table selon le risque porté",
         "synthese": "CIMA H pour le risque décès, CIMA F pour le risque de survie, TD et TV en comparaison historique"},
        "Principe de prudence : on choisit la table qui défavorise l'organisme. Pour un risque décès (le FPPN paie au décès), une table qui prévoit plus de décès est prudente : CIMA H. Pour un risque de survie (le FPPN paie tant que la personne vit), une table qui prévoit des vies plus longues est prudente : CIMA F. Les fichiers fournis portent eux-mêmes cet usage dans le nom de leur feuille. Les tables TD et TV sont les anciennes tables réglementaires remplacées en 2012 (Code CIMA, article 338 modifié).",
        tableau={"titre": tab("Tables de mortalité CIMA et usage retenu"), "entetes": ["Table", "Feuille du fichier", "Risque", "Usage dans la mission"], "largeurs": [3.8, 5.6, 7.0, 14.8],
                 "lignes": [["CIMA H", MORT["CIMA H"]["feuille"], "Décès (paiement au décès)", "Calibrage de la cotisation décès et comparaison à la mortalité observée"],
                            ["CIMA F", MORT["CIMA F"]["feuille"], "Survie (paiement tant que la personne vit)", "Engagements viagers : prise en charge des retraités, rentes"],
                            ["TD CIMA", MORT["TD CIMA"]["feuille"], "Décès (ancienne table)", "Comparaison historique"],
                            ["TV CIMA", MORT["TV CIMA"]["feuille"], "Survie (ancienne table)", "Comparaison historique"]]},
        disposition="dessous",
        analyse=["**La prudence change de sens selon le risque.** Une même table peut être prudente pour un capital décès et imprudente pour une rente : c'est pourquoi deux tables sont retenues."],
        implication="Le FPPN porte les deux risques (capital décès et, potentiellement, prise en charge viagère des retraités) : chaque engagement sera évalué avec la table qui le défavorise.",
        source=SRC_CIMA)
cs = C["courbes_survie"]
contenu({"titre": "Courbes de survie", "sous_titre": "Les tables de survie prévoient des vies plus longues que les tables de décès",
         "synthese": f"Survivants à 60 ans sur 100 naissances : {N('Survie60 H', 100 * MORT['CIMA H']['survie_60'], 2)} (CIMA H), {N('Survie60 F', 100 * MORT['CIMA F']['survie_60'], 2)} (CIMA F)"},
        "Une courbe de survie donne la part des personnes encore en vie à chaque âge. L'écart entre les courbes mesure l'écart de prudence entre une table de décès et une table de survie.",
        graphique={"titre": fig("Survivants pour 100 naissances, par âge"), "type": "lignes", "categories": [str(a) for a in cs["ages"]],
                   "series": [{"nom": n, "valeurs": [round(v, 2) for v in cs[n]], "couleur": c} for n, c in (("CIMA H", BLEU), ("CIMA F", ORANGE), ("TD CIMA", GRIS_F), ("TV CIMA", GRIS_C))],
                   "format": "0.00", "etiquettes": False, "unite": "En % des naissances"},
        analyse=[f"**L'espérance de vie à la naissance vaut {N('e0 H', MORT['CIMA H']['e0'], 2)} ans avec CIMA H et {N('e0 F', MORT['CIMA F']['e0'], 2)} ans avec CIMA F.**",
                 f"**Les anciennes tables prévoient davantage de décès** : {N('e0 TD', MORT['TD CIMA']['e0'], 2)} ans avec TD et {N('e0 TV', MORT['TV CIMA']['e0'], 2)} ans avec TV.",
                 "**Le choix de la table change donc fortement les montants calculés.**"],
        implication="Un engagement viager évalué avec une ancienne table est sous-estimé : les calculs de la mission utiliseront les tables en vigueur depuis 2012.",
        source=SRC_CIMA)
cq = C["courbes_qx"]
contenu({"titre": "Probabilités de décès", "sous_titre": "La probabilité de décès double environ tous les dix ans aux âges actifs",
         "synthese": f"Probabilité de décès à 40 ans : {N('q40 H', 1000 * MORT['CIMA H']['q40'], 2)} pour mille (CIMA H), {N('q40 F', 1000 * MORT['CIMA F']['q40'], 2)} pour mille (CIMA F)"},
        "qx en pour mille aux âges de 20 à 80 ans. Ce sont les âges de la population du FPPN (actifs puis retraités). La comparaison avec la mortalité observée se fera par le rapport décès observés / décès attendus (voir partie 12).",
        graphique={"titre": fig("Probabilité de décès dans l'année (qx), par âge"), "type": "lignes", "categories": [str(a) for a in cq["ages"]],
                   "series": [{"nom": n, "valeurs": [round(v, 2) for v in cq[n]], "couleur": c} for n, c in (("CIMA H", BLEU), ("CIMA F", ORANGE), ("TD CIMA", GRIS_F), ("TV CIMA", GRIS_C))],
                   "format": "0.00", "etiquettes": False, "unite": "En pour mille"},
        analyse=[f"**À 55 ans, la probabilité de décès vaut {N('q55 H', 1000 * MORT['CIMA H']['q55'], 2)} pour mille avec CIMA H et {N('q55 F', 1000 * MORT['CIMA F']['q55'], 2)} pour mille avec CIMA F.**",
                 f"**Entre 20 et 60 ans, CIMA H prévoit {pc('H vs TD', -C['mortalite_comparaison']['deces_20_60_H_vs_TD'])} de décès de moins que TD**, et CIMA F {pc('F vs TV', -C['mortalite_comparaison']['deces_20_60_F_vs_TV'])} de moins que TV.",
                 "**La table ronde CIMA FANAF de 2011 indique des écarts de 17 % pour CIMA H et d'environ 4 % pour CIMA F, sur l'ensemble des âges.**"],
        implication="Pour une population de policiers, exposée à des risques professionnels, la mortalité réelle peut s'écarter de la table : seul le rapport décès observés / décès attendus le dira.",
        source=SRC_CIMA)
contenu({"titre": "Indicateurs de référence", "sous_titre": "Une rente viagère vaut entre 12 et 18 années de paiement selon l'âge et la table",
         "synthese": f"Valeur d'une rente de 1 FCFA par an à 60 ans, au taux de 3,5 % : {N('a60 F', MORT['CIMA F']['a60'], 2)} (CIMA F), {N('a60 H', MORT['CIMA H']['a60'], 2)} (CIMA H)"},
        "Rente viagère : valeur aujourd'hui de 1 FCFA payé chaque année, au début de l'année, tant que la personne vit, actualisé au taux de 3,5 % (plafond du Code CIMA). Formule : ä_x = Σ (L_(x+k) / L_x) × v^k, avec v = 1 / 1,035. Exemple : si le FPPN prenait en charge 100 000 FCFA par an de soins pour un retraité de 60 ans, l'engagement vaudrait environ 100 000 × ä_60. Ce sont des valeurs de référence, pas des engagements du FPPN.",
        tableau={"titre": tab("Espérance de vie et valeur de la rente viagère (taux de 3,5 %)"), "entetes": ["Table", "e55 (ans)", "e60 (ans)", "ä55", "ä60", "ä65"], "largeurs": [3.8, 3.0, 3.0, 3.0, 3.0, 3.0],
                 "lignes": [[t, N(f"e55 {t}", MORT[t]["e55"], 2), N(f"e60 {t}", MORT[t]["e60"], 2), N(f"a55 {t}", MORT[t]["a55"], 2), N(f"a60 {t}", MORT[t]["a60"], 2), N(f"a65 {t}", MORT[t]["a65"], 2)] for t in ("CIMA H", "CIMA F", "TD CIMA", "TV CIMA")]},
        disposition="droite",
        analyse=[f"**Une rente servie à partir de 60 ans vaut {N('a60 F (t)', MORT['CIMA F']['a60'], 2)} années de paiement avec CIMA F.** Avec CIMA H, elle n'en vaut que {N('a60 H (t)', MORT['CIMA H']['a60'], 2)}.",
                 "**L'écart entre les deux tables dépasse deux années de paiement à 60 ans.** C'est l'ordre de grandeur de l'enjeu du choix de la table."],
        implication="Ces facteurs serviront à évaluer tout engagement viager du FPPN dès que les âges et les coûts annuels seront connus.",
        source=SRC_CIMA)
retenir("tables de mortalité", ("Les tables CIMA sont prêtes ; il manque les dates de naissance pour les appliquer", "Choix prudent : CIMA H pour le décès, CIMA F pour la survie",
                                "Les outils de calcul sont en place, mais ils ne peuvent pas être appliqués aux mutualistes du FPPN sans leurs âges."),
        "Une rente à 60 ans vaut 16,84 années de paiement avec CIMA F, contre 14,26 avec CIMA H. Les anciennes tables sous-estiment la durée de vie.",
        "La mortalité des policiers et la valeur des engagements viagers restent inconnues tant que les dates de naissance manquent.",
        "Le FPPN doit transmettre le fichier des mutualistes avec date de naissance et sexe (PBC, item 9) et la base des décès (item 28).",
        "Les indicateurs présentés sont des références techniques. Ils deviendront des engagements chiffrés du FPPN dès réception des données.")

# ---------------------------------------------------------------- PARTIE 12 : PROJECTION DE L'ÉQUILIBRE TECHNIQUE
separateur(12, "Projection de l'équilibre technique", "Avertissement, point de blocage, deux scénarios de 2026 à 2030")
_PA, _PB = V10["projection"]["A"], V10["projection"]["B"]
_TC = V10["tcam"]
_CRP = {x["exercice"]: x["prest_maladie"] for x in V10["croissance_prestations"]}
contenu({"titre": "Projection : avertissement", "sous_titre": "Sans les âges des mutualistes, la projection ne peut être qu'agrégée",
         "synthese": "La projection actuarielle bloque dès sa première étape : aucun fichier reçu ne donne la date de naissance des mutualistes en activité"},
        "Une projection actuarielle de l'équilibre technique suit six étapes. Le tableau indique, pour chacune, la donnée nécessaire, si elle a été reçue, et ce qu'il faut faire à la place.",
        tableau={"titre": tab("Étapes d'une projection actuarielle et point de blocage"), "entetes": ["Étape", "Donnée nécessaire", "Reçue ?", "Conséquence"], "largeurs": [5.0, 9.2, 3.2, 13.8], "taille": 8.5,
                 "lignes": [["1. Population par âge", "Date de naissance et sexe de chaque mutualiste", "Non", f"Blocage : les fichiers de la Solde et le fichier des effectifs ne portent que la date de recrutement ; la CGRAE ne donne la date de naissance que de {N('PJ ret dn', RDN['nb'])} retraités"],
                            ["2. Vieillissement et décès", "Âge de chaque mutualiste et table CIMA H", "Table seulement", "Impossible sans l'étape 1 : la table ne s'applique qu'à un âge"],
                            ["3. Départs à la retraite", "Âge et grade (âge légal de départ)", "Grade seulement", "Impossible : on ne peut dater ni les départs, ni les versements de l'EFC qu'ils déclenchent"],
                            ["4. Entrées", "Promotions de recrues prévues", "Passées seulement", "Possible en tendance (partie 4)"],
                            ["5. Cotisations et prestations par âge", "Salaires et coût des soins par âge", "Non", "Impossible : seuls les totaux annuels sont connus"],
                            ["6. Équilibre de chaque année", "Résultats des étapes 1 à 5", "Partiel", "Remplacé par le prolongement des tendances agrégées de 2023 à 2025"]]},
        disposition="dessous",
        analyse=["**La projection présentée ensuite prolonge les tendances passées, sans les expliquer.** Elle ne dit pas pourquoi les prestations augmentent (vieillissement, prix des soins, nombre de bénéficiaires) et ne peut donc pas tester l'effet d'une mesure ciblée."],
        implication="Le fichier des mutualistes avec date de naissance et sexe (PBC, item 9) est la donnée qui débloque à la fois la projection, la mortalité observée et l'engagement santé des retraités.",
        valable="fichiers reçus jusqu'au 22/10/2026 : précomptes de la Solde (postes 936, 8Z6, 998, 870), fichier des effectifs du FPPN, fichiers retour de la CGRAE (dates de naissance de février à avril 2024 seulement)",
        source="Source : liste des documents demandés (PBC) pointée au 22/10/2026 ; " + SRC_PREC[9:])
contenu({"titre": "Projection agrégée : deux scénarios", "sous_titre": f"Si les tendances de 2023 à 2025 se prolongent, le résultat technique devient négatif dès 2026 ({F('PJ A 2026', _PA[0]['resultat'])})",
         "synthese": f"Scénario A (tendance) : {F('PJ A 2030', _PA[-1]['resultat'])} en 2030 ; scénario B (prestations au rythme des cotisations) : {F('PJ B 2030', _PB[-1]['resultat'])}"},
        f"Chaque poste de 2025 croît chaque année à son taux annuel moyen. Cotisations : {pc('PJ tc', _TC['cotisations'])} ; frais : {pc('PJ tf', _TC['frais'])} (2023 à 2025). Prestations : {pc('PJ tp', _TC['prestations'])} dans le scénario A, taux des cotisations dans le scénario B. Résultat technique = cotisations - prestations - frais. Les années 2021 et 2022 sont écartées : leurs prestations maladie sont anormalement basses.",
        tableau={"titre": tab("Projection agrégée du résultat technique, maladie et décès"), "entetes": ["Exercice", "Cotisations", "Prestations A", "Résultat A", "Prestations B", "Résultat B"], "largeurs": [3.0, 5.4, 5.4, 5.4, 5.4, 5.4], "taille": 9,
                 "lignes": [[str(a["exercice"]), F(f"PJ cot {a['exercice']}", a["cotisations"]), F(f"PJ pA {a['exercice']}", a["prestations"]), F(f"PJ rA {a['exercice']}", a["resultat"]),
                             F(f"PJ pB {a['exercice']}", bb["prestations"]), F(f"PJ rB {a['exercice']}", bb["resultat"])] for a, bb in zip(_PA, _PB)]},
        disposition="dessous",
        analyse=[f"**Au rythme de 2023 à 2025, les prestations croissent de {pc('PJ tp (a)', _TC['prestations'])} par an et les cotisations de {pc('PJ tc (a)', _TC['cotisations'])} : l'équilibre technique ne tient pas.**",
                 f"**La hausse des prestations maladie ralentit déjà : {pc('PJ cr 2024', _CRP[2024])} en 2024, {pc('PJ cr 2025', _CRP[2025])} en 2025.** Prolonger le rythme moyen est donc une hypothèse pessimiste.",
                 "**Le scénario B montre l'équilibre atteint si les prestations suivent les cotisations.** La réalité se situe entre les deux scénarios ; seule une projection par âge permettrait de la situer."],
        implication="Le FPPN doit suivre chaque trimestre le rapport prestations / cotisations de la maladie : c'est l'indicateur avancé du déséquilibre que montre le scénario A.",
        valable="cotisations, prestations et frais de gestion des comptes 2023 à 2025, garanties maladie et décès prises ensemble ; hors épargne, Grenier et produits financiers",
        source=SRC_COMPTA + " ; onglet 25 du classeur des calculs")
contenu({"titre": "Projection agrégée : évolution du résultat", "sous_titre": "L'écart entre les deux scénarios mesure l'incertitude laissée par l'absence des âges",
         "synthese": "Le scénario A sert d'alerte, le scénario B de cible : le pilotage des prestations maladie fait passer de l'un à l'autre"},
        "Résultat technique projeté de 2026 à 2030 dans les deux scénarios, en millions de FCFA.",
        graphique={"titre": fig("Résultat technique projeté selon deux scénarios"), "type": "barres", "categories": [str(a["exercice"]) for a in _PA],
                   "series": [{"nom": "Scénario A : tendance 2023 à 2025", "valeurs": [round(a["resultat"]) for a in _PA], "couleur": ORANGE},
                              {"nom": "Scénario B : prestations au rythme des cotisations", "valeurs": [round(a["resultat"]) for a in _PB], "couleur": BLEU}],
                   "format": FMT_M, "unite": "En millions de FCFA"},
        analyse=[f"**En 2030, l'écart entre les deux scénarios atteint {F('PJ écart', _PB[-1]['resultat'] - _PA[-1]['resultat'])}.**",
                 "**Aucun des deux scénarios n'est une prévision** : ce sont deux bornes construites sur les seules données agrégées disponibles."],
        implication="La projection par âge doit être refaite dès réception du fichier des mutualistes ; elle remplacera ces deux bornes par un scénario central expliqué.",
        valable="mêmes données et hypothèses que la diapositive précédente",
        source=SRC_COMPTA + " ; onglet 25 du classeur des calculs")

# ---------------------------------------------------------------- PARTIE 13 : OUTILS
separateur(13, "Outils et traçabilité", "Chaîne de traitement, extraits de code, contrôle des chiffres")
contenu({"titre": "Environnement de travail", "sous_titre": "Les traitements sont écrits en Python et exécutés dans VS Code sous Windows",
         "synthese": "Dix scripts conservés, rejouables à chaque nouvel envoi du FPPN"},
        "Les images des diapositives suivantes reproduisent la fenêtre de VS Code sous Windows et y placent des extraits réels des scripts de la mission. Ce sont des reconstitutions, pas des captures d'écran ; le terminal n'est affiché que lorsque sa sortie a réellement été obtenue.",
        tableau={"titre": tab("Scripts de la mission"), "entetes": ["Script", "Rôle", "Contrôle intégré"], "largeurs": [7.4, 12.4, 11.4],
                 "lignes": [["extraction_fppn.py", "Extraction sans perte, archive découpée pour l'envoi", "Empreintes SHA-256, lignes et cellules"],
                            ["envoi_fichiers_fppn.py", "Envoi de fichiers précis en parties", "Empreinte de chaque partie"],
                            ["base_fppn.py", "Base d'analyse : lecture par en-tête ou par contenu", "Lignes de total, montants doublés"],
                            ["grenier_fppn.py, compta_fppn.py", "Lecture des exports du logiciel de prêts et des balances Sage", "Totaux imprimés retrouvés, débit = crédit"],
                            ["analyses_fppn.py, analyses_complements_fppn.py", "Indicateurs, tests, sensibilités, tables CIMA", "Classeur des calculs"],
                            ["ratios_airms_fppn.py", "Ratios prudentiels, classement des comptes, excédent par risque", "Rapprochement avec les agrégats du premier traitement"],
                            ["provisions_fppn.py, formules_fppn.py", "Provisions, variantes de sinistralité, registre des formules", "Rapprochement avec les balances"],
                            ["rapport_v9_fppn.py, donnees_v9_fppn.py, sensibilites_fppn.py, logigramme_fppn.py, classeur_v2_fppn.py", "Rapport, logigrammes et graphiques natifs, classeur", "Contrôle automatique des chiffres"]]},
        disposition="dessous",
        analyse=["**Chaque chiffre du rapport peut être recalculé par un tiers.** Les scripts sont conservés dans le dossier de la mission et rejoués à chaque nouvel envoi."],
        source="Source : dépôt de la mission (dossier scripts)")
ETAPES = [
        ("Réception", ("4 287 fichiers envoyés par le FPPN, dont des exports de logiciels et des PDF.", "Le script dresse l'inventaire, repère doublons, fichiers vides et formats mal nommés.", "Chaque fichier reçoit une empreinte numérique qui prouve qu'il n'a pas changé.", "Inventaire des fichiers et liste des fichiers à renvoyer.")),
        ("Extraction", ("Tableurs Excel, pages web nommées « .xls » et fichiers texte de la Solde.", "Le script lit les premiers octets pour reconnaître le vrai format, puis convertit chaque feuille en CSV sans perte.", "Les lignes et les cellules sont comptées avant et après conversion : 491 feuilles identiques.", "Fichiers CSV et journal des feuilles converties.")),
        ("Base", ("Fichiers CSV des précomptes, des balances et du Grenier.", "Le script reconnaît les colonnes par leur en-tête ou par leur contenu, puis réunit 4,45 millions de lignes.", "Les lignes de total et les montants doublés sont écartés et enregistrés à part.", "Base d'analyse au format Parquet et journal des rejets.")),
        ("Contrôles", ("Base d'analyse et totaux imprimés par les logiciels.", "Le script recalcule les totaux et rapproche les fichiers aller et retour, poste par poste et mois par mois.", "Les 37 balances âgées retrouvent leurs totaux et les 7 exercices sont équilibrés.", "Tableaux de contrôle et anomalies signalées au FPPN.")),
        ("Analyses", ("Base contrôlée, données externes validées et textes de l'UEMOA.", "Les scripts calculent les indicateurs, les tests statistiques, les tables CIMA et les 36 ratios prudentiels.", "Les agrégats recalculés sont comparés à ceux de la version précédente au franc près.", "Fichiers de résultats et classeur des calculs.")),
        ("Rapport", ("Fichiers de résultats et gabarits Timesaver Adenka.", "Le script construit les graphiques et tableaux natifs, les schémas modifiables et les notes.", "Chaque chiffre affiché est relu dans le PowerPoint produit et comparé à sa valeur exacte.", "Rapport PowerPoint et onglet de traçabilité."))]
contenu({"titre": "Étapes de la chaîne de traitement", "sous_titre": "Chaque étape transforme une entrée en sortie, avec un contrôle automatique",
         "synthese": "Six étapes du logigramme de la partie 2, chacune avec son entrée, son traitement, son contrôle et sa sortie"},
        "Ce tableau détaille les étapes du logigramme de la chaîne de traitement (partie 2). Chaque étape correspond à un script conservé dans le dossier de la mission.",
        tableau={"titre": tab("Entrée, traitement, contrôle et sortie de chaque étape"), "entetes": ["Étape", "Entrée", "Traitement", "Contrôle", "Sortie"], "largeurs": [2.6, 6.6, 8.2, 7.8, 6.0], "taille": 9,
                 "lignes": [[e] + list(v) for e, v in ETAPES]},
        disposition="dessous",
        analyse=["**Chaque étape produit une trace vérifiable.** L'inventaire, le journal des conversions, le journal des rejets, les tableaux de contrôle et l'onglet de traçabilité permettent à un tiers de refaire le chemin."],
        source="Source : scripts Python Adenka (dossier scripts de la mission)")
for nomimg, titre, sous, synth, expl, impl in [
        ("code_extraction", "Code : reconnaître le vrai format", "Le script lit les premiers octets pour reconnaître le vrai format d'un fichier", "C'est ainsi que les 70 fichiers « .xls » du Grenier ont été reconnus comme des pages web lisibles",
         ["**Un fichier Excel moderne commence par les octets « PK », un ancien par d'autres octets.**", "**Une page web contient « <html » ou « <table ».** Le script la lit alors comme un tableau.", "**Le terminal montre la sortie réelle de l'extraction sur le poste de la mission** : 91,5 Mo en 5 parties."],
         "Ce test évite de déclarer illisible un fichier simplement mal nommé : aucun fichier du Grenier n'a été perdu."),
        ("code_total", "Code : écarter les lignes de total", "Une ligne égale à la somme des autres est écartée automatiquement", "Règle appliquée feuille par feuille, avec une tolérance de 1 %",
         ["**Le script calcule la somme de la feuille.**", "**Une ligne dont le montant égale la somme des autres est un total**, pas un mutualiste.", "**Chaque ligne écartée est enregistrée** avec son fichier et son montant."],
         "La règle est automatique et tracée : le FPPN peut vérifier chaque ligne écartée dans le journal des rejets."),
        ("code_double", "Code : corriger les montants doublés", "Les montants écrits deux fois à la suite sont corrigés et signalés", "Correction limitée au poste décès, au-delà d'un million de FCFA",
         ["**Le script découpe le montant en deux moitiés.**", "**Si les deux moitiés sont identiques, il retient la première.**", "**Le montant d'origine est conservé** dans une colonne à part."],
         "La correction est réversible : le montant d'origine reste disponible pour le FPPN."),
        ("code_grenier", "Code : lire les montants du Grenier", "Les montants imprimés avec espaces et « XOF » sont convertis sans perte", "Les totaux recalculés sont égaux aux totaux imprimés dans les 37 balances âgées",
         ["**Le script retire les espaces, y compris insécables.**", "**Il retire aussi la mention « XOF ».**", "**Le nombre obtenu est comparé au total imprimé** par le logiciel."],
         "L'égalité des totaux sur les 37 balances prouve que la lecture n'a rien perdu."),
        ("code_sc", "Code : calculer la sinistralité", "Le ratio S/C est calculé compte par compte, pour chaque exercice", "Chaque compte utilisé est cité dans l'annexe des méthodes",
         ["**Les cotisations viennent des comptes 7010 et 7021.**", "**Les prestations viennent des comptes 601, 6013, 6018, 60800007, 6021 et 6024.**", "**Le résultat est écrit dans le classeur des calculs.**"],
         "Le choix des comptes est explicite dans le code : un lecteur peut le contester et le modifier."),
        ("code_stats", "Code : intervalle de confiance", "L'intervalle de Wilson encadre la proportion de prêts en retard", "Méthode adaptée aux proportions, même faibles",
         ["**La proportion vaut le nombre de prêts en retard divisé par le nombre de prêts.**", "**L'intervalle est centré et élargi** pour rester entre 0 et 1.", "**z vaut 1,96 pour un niveau de confiance de 95 %.**"],
         "L'intervalle dit dans quelle plage se situe la vraie proportion de prêts en retard, compte tenu du hasard."),
        ("code_statut", "Code : statut d'un ratio prudentiel", "Le statut de chaque ratio suit une règle écrite dans le code", "Conforme, non conforme, en attente ou à apprécier, selon la norme du texte",
         ["**Une valeur absente donne « En attente » ; une norme non chiffrée donne « À apprécier ».**", "**Le test compare la valeur à la norme** : fourchette, minimum ou maximum.", "**Les années 2020, 2021 et 2026 portent la mention « indicatif ».** Le terminal montre la sortie réelle du script."],
         "La règle de statut est la même pour les 36 indicateurs : elle ne laisse aucune place à l'appréciation au cas par cas.")]:
    contenu({"titre": titre, "sous_titre": sous, "synthese": synth},
            "Extrait réel du script, placé dans une fenêtre de VS Code sous Windows reconstituée (thème sombre, fil d'Ariane du dossier de la mission, numéros de ligne). Le code complet est dans le dossier scripts de la mission ; il peut être relu et rejoué par un tiers.",
            image={"fichier": os.path.join(IMG, nomimg + ".png"), "legende": fig(f"Extrait du script dans VS Code sous Windows ({titre.split(': ')[1]})")}, analyse=expl, implication=impl,
            source="Source : scripts de la mission (fenêtre VS Code reconstituée, pas une capture d'écran)")
methode_gab({"titre": "Contrôle automatique des chiffres", "sous_titre": "Chaque chiffre du rapport est recalculé et comparé au texte affiché",
             "synthese": "Le contrôle relit le PowerPoint produit et vérifie que chaque chiffre enregistré figure sur sa diapositive"},
            [("Enregistrement", "À la construction, chaque chiffre est produit par une fonction qui garde sa valeur exacte, sa diapositive et le texte affiché."),
             ("Relecture", "Après la construction, un script relit le fichier PowerPoint, textes, tableaux et notes compris."),
             ("Comparaison", "Le script vérifie la présence de chaque texte enregistré sur sa diapositive et compte les chiffres retrouvés."),
             ("Résultat", "Le résultat figure dans l'onglet « Traçabilité » du classeur des calculs, ligne par ligne, avec la valeur exacte de chaque chiffre.")],
            "Principe : aucun chiffre n'est saisi à la main. Le contrôle garantit que le rapport affiche exactement les valeurs calculées, avec deux décimales pour les pourcentages et des montants exacts en FCFA.")

# ---------------------------------------------------------------- PARTIE 14 : ANALYSES EN ATTENTE ET RISQUES
separateur(14, "Analyses en attente", "Documents à obtenir, risques, points à clarifier")
# ---- v11 : partie 14 : analyses en attente (méthodes retenues), questions au FPPN, documents prioritaires (PBC du 23/10/2026)
def _n(*anciens):
    return ", ".join(str(x) for x in sorted({PBCM[a] for a in anciens}))


# ---- v13 : matrice données × analyses (statut de chaque donnée dans la PBC du 23/10/2026)
_DON = [("Base des prestations datée", 47, "🟠"), ("Dossiers non payés", 7, "🟠"), ("Fichiers des décès", 64, "🟠"), ("Dates de naissance des mutualistes", 4, "🟠"),
        ("Dates de naissance des retraités", 84, "🔴"), ("Fichier des ayants droit", 43, "🔴"), ("Comptes individuels d'épargne", 9, "🟠"),
        ("Taux servi arrêté", 69, "🟠"), ("Inventaire des placements", 55, "🟠")]
_BES = {  # analyse : données nécessaires (indices de _DON)
    "Provision maladie dossier par dossier": (0, 1),
    "Provision décès dossier par dossier": (1, 2),
    "Mortalité observée par âge": (2, 3),
    "Engagement santé des retraités": (0, 4),
    "Engagement d'épargne par épargnant": (6, 7),
    "Adossement par échéance": (6, 8),
    "Projection par garantie": (0, 3, 5),
}
_MAT_D = [[an_] + [_DON[k][2] if k in ks_ else "" for k in range(len(_DON))] for an_, ks_ in _BES.items()]
contenu({"titre": "Matrice des données et des analyses", "sous_titre": "Aucune analyse actuarielle centrale ne dispose de toutes ses données",
         "synthese": "Analyses en lignes, données nécessaires en colonnes, avec leur statut dans la PBC du 23/10/2026"},
        "Une case vide : la donnée n'est pas nécessaire à l'analyse. Une boule : la donnée est nécessaire, et la couleur donne son statut dans la PBC.",
        tableau={"titre": tab("Données nécessaires à chaque analyse et statut dans la PBC"), "entetes": ["Analyse"] + [f"{n_}\nPBC n° {p_}" for n_, p_, _ in _DON],
                 "largeurs": [5.6] + [2.84] * 9, "taille": 8, "sans_code": True, "centre": True, "hauteur_ligne": 0.85,
                 "legende": [("🔴", "Non reçu", "donnée non reçue"), ("🟠", "Partiel", "donnée reçue en partie"), ("🟢", "Reçu", "donnée reçue")], "lignes": _MAT_D},
        disposition="dessous",
        analyse=["**La base des prestations datée et les dates de naissance interviennent dans quatre des sept analyses.** Les dates de naissance des retraités et le fichier des ayants droit n'ont pas été reçus du tout."],
        valable="PBC pointée au 23/10/2026",
        source="Source : PBC du 23/10/2026 (« FPPN_PBC_complete_pointee_20261023.xlsx ») ; méthodes retenues du rapport")

contenu({"titre": "Analyses en attente", "sous_titre": "Quatre analyses attendent des données du FPPN, toutes avec une méthode déjà choisie",
         "synthese": "Les conclusions du rapport sont établies sur les documents reçus ; ces analyses les préciseront sans en changer le sens"},
        "Chaque analyse en attente utilise la méthode retenue dans le rapport, pas une méthode écartée. Les numéros renvoient à la PBC du 23/10/2026.",
        tableau={"titre": tab("Analyses en attente, méthode retenue et données nécessaires"), "entetes": ["Analyse", "Méthode retenue", "Ce que l'analyse précisera", "Données nécessaires (PBC n°)"],
                 "largeurs": [5.6, 9.4, 8.4, 7.8], "taille": 9,
                 "lignes": [["Provision maladie", "Soins de l'année facturés après la clôture, sur la base des factures datées (date des soins, du dépôt et du paiement)", f"Montant exact, aujourd'hui borné à {F('AT psap', psap_m_elargie(2024))} au 31/12/2024", "Base des prestations datée (" + _n(71) + ")"],
                            ["Provision décès", "Dossiers connus au compte 41591000 et décès survenus déclarés après la clôture", "Part des dossiers forclos dans la dette de " + F("AT dette", DETTE_DEC26), "Liste des dossiers du 41591000, fichiers des décès (" + _n(115, 92) + ")"],
                            ["Mortalité observée et santé des retraités", "Décès observés comparés aux décès attendus des tables CIMA H et F, par âge", "Écart de mortalité des policiers, engagement de santé à vie des retraités", "Dates de naissance des mutualistes et des retraités (" + _n(9, 120) + ")"],
                            ["Engagement d'épargne par épargnant", "Solde de chaque compte individuel augmenté des intérêts au taux arrêté", "Échéancier des sorties, à la place du taux de sortie observé", "Comptes individuels, taux servi (" + _n(17, 99) + ")"]]},
        disposition="dessous",
        analyse=["**Aucune conclusion du rapport ne dépend de ces analyses pour son sens.** Les provisions du bilan dépassent les besoins mesurés, le non-adossement est démontré et les placements sont insuffisants, quelles que soient les valeurs exactes."],
        implication="Les données listées sont demandées au FPPN avec la priorité « Urgent ». Leur réception permettra une mise à jour chiffrée des diapositives concernées.",
        valable="analyses en attente au 23/10/2026",
        source="Source : PBC actuarielle du 23/10/2026 ; méthodes de la partie 9 et de l'annexe des méthodes non retenues")
QFPPN = [
    ["Cotisations", "Pourquoi la Solde précompte-t-elle 3,77 % et 1,23 % du salaire de base, alors que les statuts (art. 17) fixent 3,5 % et 1,5 % ?", _n(56), "Urgent"],
    ["Cotisations", "Le délai de trois mois de cotisation avant les prestations s'applique-t-il à la maladie et au décès ? L'entretien du 16/10/2026 indiquait l'absence de délai de carence.", _n(59), "Urgent"],
    ["Cotisations", "Quelle décision a porté l'EFC de 6 000 FCFA à 10 000 FCFA en 2021 ?", _n(58), "Moyen"],
    ["Cotisations", "Que représentent les références PTI, RPEP, EP, TI, R-PEP et PTQ précomptées au poste 998 ?", "85", "Moyen"],
    ["Cotisations", "Le FPPN rapproche-t-il les fichiers aller et retour de chaque produit ? Si oui, qui le fait et selon quelle procédure ?", "88", "Faible"],
    ["Épargne", "Quel taux est servi sur l'EFC (3,5 %, 5,5 % ou 14 %), et quelle décision signée le fixe ?", _n(99), "Urgent"],
    ["Épargne", "Le DAT du compte 52210200 est-il un placement du FPPN ou un placement affecté à l'EFC ?", _n(80), "Urgent"],
    ["Provisions", "Quelle décision permettra de reprendre ou de justifier les provisions passées vers 2004 ?", _n(72), "Urgent"],
    ["Provisions", "Pourquoi les sinistres décès 2021 valent-ils 156 020 000 FCFA dans l'étude des ratios et 401 856 666 FCFA dans la balance 2021 ?", "86", "Moyen"],
    ["Provisions", "À quel grade correspond l'allocation de 2 200 000 FCFA des états décès, et quel texte fonde l'exclusion du suicide ?", _n(116), "Moyen"],
    ["Retraités", "La garantie santé des retraités est-elle accordée à vie ou pour une durée limitée ? (Le barème de cotisation de la CGRAE a été reçu le 23/10/2026.)", _n(79), "Urgent"],
    ["Personnel", "Quel âge de départ le FPPN applique-t-il à son personnel civil (cinq agents de 55 ans ou plus en poste) ?", _n(113), "Urgent"],
    ["Personnel", "Pourquoi douze dates d'embauche diffèrent-elles entre les DISA 2021 à 2023 et la DISA 2025 ?", _n(113), "Urgent"],
    ["Personnel", "À quelle entité (Direction Générale, clinique, Grenier) sont rattachés les dix-neuf agents recrutés en 2025 ?", _n(113), "Urgent"],
    ["Placements", "Quelle est la nature de la créance de 3 407 519 850 FCFA sur l'hôpital de la Police, et un remboursement est-il prévu ?", _n(87), "Moyen"],
    ["Placements", "Le compte 10620000 enregistre-t-il la réévaluation du terrain du siège ou de matériels, comme l'indique son intitulé ?", _n(86), "Urgent"],
    ["Placements", "Quel est le budget du projet de Cocody, son financement et la part financée par l'épargne ?", _n(118), "Urgent"],
    ["Grenier", "Pourquoi l'essentiel de l'encours était-il en retard fin janvier 2026, et comment sont réglées les échéances non précomptées ?", "90", "Moyen"],
    ["Flotte automobile", "Sur quelle assiette la commission de 20 % versée par GNA est-elle calculée ?", _n(89), "Moyen"],
    ["MEDOC", "Sur quelle période le produit MEDOC a-t-il fonctionné, et quand a-t-il été suspendu ?", _n(69), "Faible"],
    ["Gouvernance", "Quelle version des comptes 2023 a été approuvée, et pourquoi le résultat au 31/08/2026 diffère-t-il entre états financiers et balance ?", _n(110), "Urgent"],
    ["Gouvernance", "Le directeur général a-t-il été désigné ou prolongé après 2023, et qui autorise les placements ?", _n(117), "Urgent"],
    ["Conformité", "Quels sont le numéro, la date et les risques autorisés de l'agrément AIRMS ?", _n(44), "Urgent"],
    ["Conformité", "Le FPPN adhère-t-il au FOMUS-CI, et quel montant a-t-il payé depuis 2022 ?", _n(49), "Moyen"],
    ["Conformité", "Une étude actuarielle des risques longs ou une certification des comptes a-t-elle été réalisée ?", _n(48, 47), "Moyen"],
    ["Conformité", "Quel est l'effectif total de la Police nationale au 31/12 de 2020 à 2025 et au 31/08/2026 ?", "89", "Moyen"],
]
_moit = (len(QFPPN) + 1) // 2
for _k, _lig in enumerate((QFPPN[:_moit], QFPPN[_moit:]), 1):
    contenu({"titre": f"Questions au FPPN ({_k}/2)", "sous_titre": f"{len(QFPPN)} questions simples, chacune reliée à un constat du rapport et à la PBC",
             "synthese": "Réponse écrite attendue avec la pièce qui l'établit ; liste complète des documents dans la PBC du 23/10/2026"},
            "Questions transmises au FPPN dans le fichier « FPPN_Demandes_au_FPPN_20261023.xlsx », avec une colonne de réponse.",
            tableau={"titre": tab(f"Questions au FPPN ({_k}/2)"), "entetes": ["N°", "Thème", "Question", "PBC n°", "Priorité"], "largeurs": [1.0, 3.6, 20.0, 3.6, 3.0], "taille": 8.5,
                     "lignes": [[str(i), *row] for i, row in enumerate(_lig, 1 + (_k - 1) * _moit)]},
            disposition="dessous",
            valable="questions ouvertes au 23/10/2026",
            source="Source : analyses Adenka ; PBC actuarielle du 23/10/2026 ; FPPN_Demandes_au_FPPN_20261023.xlsx")
separateur("A", "Annexes", "Méthodes non retenues, validité, listes, abréviations, références")
# ---- v10 : annexe « Méthodes non retenues et pourquoi » (conditions d'emploi citées, démonstration chiffrée, conséquence)
_MB = [x for x in V9["mortalite_brute"] if 2021 <= x["annee"] <= 2025]
_DMB = sum(x["deces_adherents"] for x in _MB)
_LAT = {c: TARDM[c]["references_annee_precedente"] / TARDM[c]["factures_sinistres"] for c in TARDM}
_pm25a = next(x["placements_moyens"] for x in R["rendement_placements"] if x["exercice"] == 2025)
SRC_CIRC = "circulaire n° 00230/CIMA/CRCA/PDT/2005"


def methode_non_retenue(titre, sous_titre, synthese, principe, lignes, consequence, a_la_place, valable_, source_):
    contenu({"titre": titre, "sous_titre": sous_titre, "synthese": synthese}, principe,
            tableau={"titre": tab("Conditions d'emploi de la méthode et données du FPPN"), "entetes": ["Condition d'emploi", "Source de la condition", "Ce que montrent les données du FPPN", "Remplie ?"],
                     "largeurs": [6.0, 7.4, 14.2, 3.6], "taille": 8, "lignes": lignes},
            disposition="dessous", analyse=[consequence], implication=a_la_place, valable=valable_, source=source_)


contenu({"titre": "Méthodes non retenues : vue d'ensemble", "sous_titre": "Sept méthodes courantes ne sont pas appliquées : leurs conditions d'emploi ne sont pas remplies",
         "synthese": "Appliquer une méthode hors de ses conditions d'emploi produirait des chiffres précis en apparence et faux en réalité"},
        "Pour chaque méthode, les diapositives suivantes citent les conditions d'emploi dans un texte ou un ouvrage de référence, montrent avec les chiffres du FPPN pourquoi elles ne sont pas remplies, et indiquent la méthode appliquée à la place.",
        tableau={"titre": tab("Méthodes non retenues et méthodes appliquées à la place"), "entetes": ["Méthode non retenue", "Condition principale non remplie", "Méthode appliquée à la place", "Détail"], "largeurs": [6.4, 11.0, 9.0, 4.8], "taille": 8.5,
                 "lignes": [["Chain Ladder et méthode CIMA des tardifs", "Des paiements étalés sur plusieurs années de développement", "Méthode courte sur les factures en retard", ref("Méthode non retenue : Chain Ladder et méthode CIMA")],
                            ["Mortalité observée par âge (A/E)", "La date de naissance de chaque mutualiste", "Taux brut de mortalité", ref("Méthode non retenue : mortalité observée par âge")],
                            ["Provision mathématique santé des retraités", "L'âge et le coût des soins de chaque retraité", "Description des cotisations", ref("Méthode non retenue : provision santé des retraités")],
                            ["Duration et immunisation", "Des flux de passif connus et des actifs à revenu fixe", "Écarts de liquidité", ref("Méthode non retenue : duration et immunisation")],
                            ["Modèle actif-passif stochastique", "Des contrats décrits par âge et des scénarios économiques", "Tests de résistance calibrés sur le FPPN", ref("Méthode non retenue : modèle stochastique")],
                            ["Valeur en risque (VaR)", "Des actifs évalués au prix du marché", "Classes de liquidité", ref("Méthode non retenue : valeur en risque")],
                            ["Capital de solvabilité de type Solvabilité II", "Un cadre réglementaire qui l'impose et des engagements en meilleure estimation", "Marge de solvabilité de la Décision n° 05/2022", ref("Méthode non retenue : capital de solvabilité")]]},
        disposition="dessous",
        analyse=["**Une même cause revient : l'absence de données individuelles (âges, dates de soins, comptes d'épargne).** Les méthodes retenues exploitent au mieux les données agrégées reçues."],
        implication="Chaque méthode non retenue pourra être appliquée dès réception des données citées dans sa diapositive.",
        valable="données reçues jusqu'au 22/10/2026", source="Source : Adenka, d'après les références citées dans chaque diapositive")
methode_non_retenue("Méthode non retenue : Chain Ladder et méthode CIMA", "Les factures se règlent en quelques mois : un triangle annuel n'aurait qu'une seule colonne utile",
                    f"Moins de {pc('AN lat max', max(_LAT.values()))} des factures d'une année sont enregistrées l'année suivante",
                    "Chain Ladder et la méthode CIMA des tardifs projettent le rythme passé de règlement, année de survenance par année de développement. Elles ne sont utiles que si une part notable des sinistres se règle après la fin de l'année de survenance.",
                    [["Paiements étalés sur plusieurs années de développement", SRC_CIRC + " (cadences sur quatre années de survenance)", f"Part des factures enregistrées l'année suivante : {', '.join(pc(f'AN lat {c}', _LAT[c]) for c in sorted(_LAT))} (clôtures {', '.join(str(c) for c in sorted(_LAT))})", "Non"],
                     ["Rythme de règlement stable d'une année à l'autre", "Hypothèse de base de Chain Ladder", f"Délai moyen de paiement des cliniques : {N('AN d20', DELAI[2020], 0)} mois en 2020, {N('AN d21', DELAI[2021], 0)} en 2021, {N('AN d25', DELAI[2025], 2)} en 2025 : le rythme a changé de nature", "Non"],
                     ["Base des prestations datée (soins, dépôt, paiement)", REX + ", art. 10", "Non reçue (PBC, item 71) ; seule l'année de la facture est lisible dans son numéro", "Non"],
                     ["Délai de règlement borné par les textes", SRC_RI + ", art. 52, 57 et 60", "Paiement à trois mois du dépôt, demande de remboursement limitée à trois mois", "Oui"]],
                    "**Appliquée aux données du FPPN, la méthode projetterait le rattrapage des arriérés de 2020 à 2022 comme un rythme normal et surestimerait fortement la provision.**",
                    "Méthode appliquée : soins de l'année facturés l'année suivante, mesurés facture par facture (diapositive « Provision maladie : factures arrivées en retard »).",
                    "factures de cliniques du journal SIN 2023 à 2025 ; délais calculés sur le compte 40110000 de 2020 au 31/08/2026",
                    "Source : " + SRC_CIRC + " ; " + REX + ", art. 10 ; " + SRC_RI + " ; " + SRC_GLT)
methode_non_retenue("Méthode non retenue : mortalité observée par âge", "Sans date de naissance, les décès attendus par la table ne peuvent pas être calculés",
                    f"{N('AN déc', _DMB)} décès de membres de 2021 à 2025 : même avec les âges, l'écart à la table ne serait mesuré qu'à ±{pc('AN ic', 1.96 / _DMB ** 0.5)} près",
                    "Le rapport A/E divise les décès observés par les décès attendus : somme, âge par âge, de l'exposition au risque multipliée par la probabilité de décès de la table CIMA H.",
                    [["Exposition au risque par âge", "Définition de la table de mortalité, " + REX + ", art. 1", "Aucun fichier reçu ne porte la date de naissance des mutualistes", "Non"],
                     ["Décès datés et rattachés à un âge", "Même source", f"Décès de membres comptés par année ({N('AN déc (t)', _DMB)} de 2021 à 2025), sans âge", "Non"],
                     ["Nombre de décès suffisant", "Loi de Poisson : intervalle à 95 % d'environ ±1,96 / racine du nombre de décès", f"Intervalle de ±{pc('AN ic (t)', 1.96 / _DMB ** 0.5)} sur cinq ans", "Partiellement"]],
                    "**La mortalité des policiers ne peut pas être comparée à la table CIMA H.** Seul le taux brut, toutes tranches d'âge confondues, est calculable.",
                    "Méthode appliquée : taux brut de mortalité des membres participants (diapositive « Mortalité brute des membres participants »).",
                    "décès de membres recensés par le FPPN de 2021 à 2025", SRC_RATIO + " ; " + SRC_CIMA[9:])
methode_non_retenue("Méthode non retenue : provision santé des retraités", "L'engagement à vie ne se calcule qu'avec l'âge et le coût des soins de chaque retraité",
                    f"{N('AN ret', CG[-1]['retraites'])} retraités précomptés en {CG[-1]['libelle']}, dont {N('AN ret dn', RDN['nb'])} avec une date de naissance, sans prestations isolées",
                    "Provision mathématique = valeur actuelle des soins futurs - valeur actuelle des cotisations futures, retraité par retraité, avec une table de survie (CIMA F) et un taux d'actualisation.",
                    [["Âge de chaque retraité", "[[Code CIMA, art. 334-8]] (référence technique)", f"Date de naissance pour {N('AN ret dn (t)', RDN['nb'])} retraités seulement (fichiers retour de la CGRAE de février à avril 2024)", "Partiellement"],
                     ["Coût annuel des soins par âge", "Même source", "Prestations des retraités non isolées dans les comptes ni dans les fichiers", "Non"],
                     ["Règle de la garantie (durée, cotisation)", "Bulletin de souscription, annexe 2 de la convention", "Non transmis", "Non"]],
                    "**Calculer une provision avec un âge moyen supposé donnerait un chiffre sans fondement** : la valeur d'une rente à 60 ans varie de plus de deux années de paiement selon la seule table retenue (partie 11).",
                    "Méthode appliquée : description des cotisations des retraités (diapositive « Santé des retraités : ce que l'on peut mesurer »).",
                    f"retraités précomptés par la CGRAE de {CG[0]['libelle']} à {CG[-1]['libelle']} ; dates de naissance des fichiers retour de février à avril 2024", "Source : CONVENTION DE PARTENARIAT ENTRE IPS-CGRAE ET LE FPPN.pdf ; Code CIMA ; " + SRC_PREC[9:])
methode_non_retenue("Méthode non retenue : duration et immunisation", "L'immunisation suppose des flux de passif connus et des actifs à revenu fixe",
                    f"Placements financiers moyens 2025 : {pc('AN pm', _pm25a / EP_DUE25)} de l'épargne due ; taux servi non arrêté",
                    "L'immunisation rééquilibre les actifs pour que leur valeur varie comme celle des passifs quand les taux changent. Elle compare les durations (durée moyenne pondérée des flux) de l'actif et du passif.",
                    [["Flux de passif prévisibles", SRC_ALM["IAA"] + ", p. 3 et 23", "Épargne sans échéancier ; taux servi de 3,5 %, 5,5 % ou 14 % selon le document", "Non"],
                     ["Actifs à revenu fixe en volume suffisant", "Même source", f"DAT soldé chaque année et FCP ; placements moyens 2025 de {F('AN pm (t)', _pm25a)}", "Non"],
                     ["Risque principal : le taux d'intérêt", SRC_ALM["ONT"] + ", p. 8", "Risque principal : la liquidité (couverture de l'année inférieure à 40 %)", "Non"]],
                    "**Le risque du FPPN n'est pas une variation de taux, mais un manque de liquidité** : mesurer une duration ne dirait rien de la capacité à payer les sorties de l'année.",
                    "Méthode appliquée : écarts de liquidité par échéance (diapositive « Adossement : écarts de liquidité année par année »).",
                    "bilans au 31/12/2025 et au 31/08/2026", "Source : " + SRC_ALM["IAA"] + " ; " + SRC_ALM["ONT"] + " ; " + SRC_COMPTA[9:])
methode_non_retenue("Méthode non retenue : modèle stochastique", "Un modèle stochastique demande des contrats décrits par âge et des scénarios économiques",
                    "Aucune des deux entrées du modèle n'existe pour le FPPN",
                    "Un modèle actif-passif stochastique projette des milliers de scénarios économiques (méthode de Monte-Carlo) sur des groupes de contrats décrits par âge, sexe, ancienneté et taux garanti.",
                    [["Groupes de contrats par âge, sexe, ancienneté et taux garanti", SRC_ALM["TIC"] + ", p. 54", "Âges non reçus ; taux garanti non arrêté", "Non"],
                     ["Actifs en valeur de marché", "Même source", "Terrains, clinique et créances en valeur comptable non réévaluée", "Non"],
                     ["Générateur de scénarios économiques", "Même source, section 3.2", "Aucune série de prix des actifs du FPPN", "Non"]],
                    "**Le modèle produirait une distribution de résultats à partir d'entrées inventées.**",
                    "Méthode appliquée : trois scénarios déterministes calibrés sur les extrêmes observés au FPPN (diapositive « Adossement : tests de résistance »).",
                    "données reçues jusqu'au 22/10/2026", "Source : " + SRC_ALM["TIC"])
methode_non_retenue("Méthode non retenue : valeur en risque", "La VaR mesure une perte de valeur de marché, alors que les actifs du FPPN ne sont pas évalués au marché",
                    "Les actifs du FPPN sont tenus en valeur comptable ; aucun n'est coté à l'exception du FCP",
                    "La valeur en risque (VaR) est la perte de valeur de marché qui n'est dépassée qu'avec une faible probabilité sur une période donnée.",
                    [["Indicateur adapté à la gestion actif-passif", SRC_ALM["BAN"] + ", p. 9", "La source indique que la VaR n'est pas adaptée à l'ALM, « monde des intérêts courus »", "Non"],
                     ["Séries de prix des actifs", "Définition de la VaR", "Aucune série de prix pour les terrains, la clinique, les prêts et les créances", "Non"]],
                    "**Une VaR ne mesurerait que le FCP, une petite part de l'actif.**",
                    "Méthode appliquée : classes de liquidité des actifs (diapositive « Choix des méthodes d'adossement »).",
                    "bilan au 31/08/2026", "Source : " + SRC_ALM["BAN"])
methode_non_retenue("Méthode non retenue : capital de solvabilité", "Le capital de type Solvabilité II n'est pas exigé du FPPN et demande des engagements en meilleure estimation",
                    "Le cadre applicable est la marge de solvabilité de la Décision n° 05/2022, calculée en partie 10",
                    "Le capital de solvabilité requis (SCR) de la formule standard européenne applique des chocs (par exemple 40 % de rachats massifs) à la meilleure estimation des engagements.",
                    [["Cadre réglementaire qui l'impose", "Décision n° 05/2022 ; " + REX, "Le texte applicable fixe une marge de solvabilité, non un SCR", "Non"],
                     ["Engagements en meilleure estimation", SRC_ALM["TIC"] + ", p. 110", "Épargne sans échéancier ni taux arrêté ; provisions non calculées", "Non"],
                     ["Chocs calibrés sur le portefeuille", "Même source", f"Choc de 40 % calibré sur des contrats européens ; sortie maximale observée au FPPN : {pc('AN tmax', T_MAX)}", "Non"]],
                    "**Le SCR mélangerait un cadre étranger et des engagements non estimés.**",
                    "Méthode appliquée : marge de solvabilité de la Décision n° 05/2022 et tests de résistance calibrés sur le FPPN (partie 10 et partie 9).",
                    "ratios de 2020 au 31/08/2026", "Source : " + SRC_ALM["TIC"] + " ; Décision n° 05/2022 ; " + REX)
# ---- v10 : annexe « Fichiers sources » (noms exacts des fichiers cités par les lignes « Valable pour » et les sources)
_sm = pd.read_parquet(os.path.join(BASE, "sources_mois.parquet"))
_sm = _sm[_sm.retenu]
_gb = pd.read_parquet(os.path.join(BASE, "grenier_balances.parquet"), columns=["fichier"]).fichier.drop_duplicates()
_gp = pd.read_parquet(os.path.join(BASE, "grenier_production.parquet"), columns=["periode_du"])
_cb = pd.read_parquet(os.path.join(BASE, "compta_balances.parquet"), columns=["nature", "fichier"]).drop_duplicates()
_POSTE = {"936": "Cotisation maladie des actifs", "8Z6": "Cotisation décès des actifs", "998": "Épargne PEP", "870": "Épargne EFC", "9X4": "Remboursements du Grenier", "9W3": "Primes GNA (flotte automobile)", "8Z7": "MEDOC", "CGRAE": "Cotisations des retraités"}
_lsrc = []
for po, lib in _POSTE.items():
    g = _sm[_sm.poste == po]
    if len(g):
        dossiers = sorted({d.split("\\")[0] + "\\" + d.split("\\")[1] + ("\\" + d.split("\\")[2] if len(d.split("\\")) > 2 and not d.split("\\")[2].isdigit() else "") for d in g.dossier})
        _lsrc.append([f"Précomptes, poste {po}", lib, (" ; ".join("« " + d + " »" for d in dossiers) + " : ", ("FS n " + po, g.fichier.nunique()), " fichiers"),
                      f"{int(g[g.annee == g.annee.min()].mois.min()):02d}/{int(g.annee.min())} à {int(g[g.annee == g.annee.max()].mois.max()):02d}/{int(g.annee.max())}"])
_lsrc = [["Balances des comptes", "Soldes et mouvements par compte", " ; ".join("« " + f + " »" for f in _cb[_cb.nature == "comptes"].fichier), "01/09/2020 (à-nouveaux) au 31/08/2026"],
         ["Balances des tiers", "Soldes par tiers", " ; ".join("« " + f + " »" for f in _cb[_cb.nature == "tiers"].fichier), "2020, 2022 à 2024, 2026"],
         ["Grands livres", "Écritures détaillées", "« GRAND LIVRE FPPN 2024.xlsx » ; « GRAND LIVRE DES TIERS FPPN 2023.xlsx » ; « GRAND LIVRE DES TIERS 2024.xlsx » ; « GRAND LIVRES DES TIERS 2025.xlsx »", "2023 à 2025"],
         ["Allocations décès (fin d'année)", "Décès non payés valorisés au barème", SRC_LOT4, "2021 à 2025"],
         ["Grenier, balances âgées", "Encours et retards par prêt", ("", ("FS gb", len(_gb)), " fichiers « BalanceAgeeDetaille … .xls » (annuels 2021 à 2025, mensuels 2026 par portefeuille CRF, CRO, CRS)"), "31/12/2021 au 30/09/2026"],
         ["Grenier, production", "Prêts débloqués", "fichiers « production_de_la_periode … .xls » (annuels 2021 à 2025, mensuels 2026)", "2021 à 09/2026"]] + _lsrc + [
         ["Personnel", "Déclarations des salaires", SRC_DISA[9:], "2021 à 2025"],
         ["Actifs et gouvernance", "Clinique, projet de Cocody, directeur général", SRC_CMP + " ; " + SRC_COC + " ; " + SRC_DG, "2020 à 2025"],
         ["Tables de mortalité", "Lx, qx, dx par âge", "« CIMA_H.xlsx » ; « CIMA_F.xlsx » ; « TD_CIMA.xlsx » ; « TV_CIMA.xlsx »", "Âges 0 à 110 ans"]]
for _k in range(0, len(_lsrc), 7):
    contenu({"titre": f"Fichiers sources ({_k // 7 + 1}/{(len(_lsrc) + 6) // 7})", "sous_titre": "Noms exacts des fichiers du FPPN cités dans les lignes « Valable pour » et dans les sources",
             "synthese": "Les lignes « Valable pour » citent les familles de fichiers ; ce tableau donne les noms exacts"},
            "Tous les fichiers ont été reçus du FPPN le 15/10/2026 (précomptes, balances, Grenier), le 18/10/2026 (tables de mortalité) ou le 22/10/2026 (autres documents).",
            tableau={"titre": tab("Fichiers sources et période couverte"), "entetes": ["Famille", "Contenu", "Fichiers exacts", "Période"], "largeurs": [4.6, 4.6, 17.4, 4.6], "taille": 7.5, "lignes": [[(c[0] + N(*c[1]) + c[2]) if isinstance(c, tuple) else c for c in lg_] for lg_ in _lsrc[_k:_k + 7]]},
            disposition="dessous", valable="fichiers reçus jusqu'au 22/10/2026 et retenus après les contrôles de la partie 3", source="Source : inventaire des fichiers reçus du FPPN ; traitements Adenka")

SL.append({"type": "annexe", "titre": "Tableau des formules : lecture", "notes": "Le tableau complet (quinze colonnes) figure dans l'onglet 11_formules du classeur des calculs v9.", "tableau": {"titre": tab("Colonnes et statuts du tableau des formules"),
           "entetes": ["Élément", "Signification"], "largeurs": [7.0, 24.2], "taille": 9,
           "lignes": [["Identifiant", "F-partie-numéro : renvoie à la ligne du classeur, où figurent la formule détaillée, les données, l'écart chiffré, les risques et le calcul."],
                      ["Formule", "Ce que calcule l'indicateur, en mots ; la formule en comptes ou en colonnes est dans le classeur."],
                      ["Source de la méthode", "Texte réglementaire (article ou paragraphe) ou, à défaut, référence technique."],
                      ["Méthode du FPPN", "Méthode du FPPN quand elle est connue ; « Non communiquée » en rouge quand aucun calcul n'a été transmis."],
                      ["Conforme au texte", "La formule du texte est appliquée telle quelle."],
                      ["Divergent, justifié", "La formule diffère du texte ; la raison et l'écart sont expliqués dans le rapport et dans le classeur."],
                      ["Sans texte applicable", "Aucun texte ne fixe la méthode ; une référence technique est citée."],
                      ["En attente", "Le calcul attend un document du FPPN (PBC)."],
                      ["À apprécier", "Le texte ne fixe pas de norme chiffrée."]]},
           "source": "Source : Adenka"})
_REG = FORM.registre(Q, PV, R)
for _k in range(0, len(_REG), 6):
    _part = _REG[_k:_k + 6]
    SL.append({"type": "annexe", "titre": f"Tableau des formules ({_part[0]['id']} à {_part[-1]['id']})", "notes": "Extrait du tableau des formules ; colonnes complètes dans l'onglet 11_formules du classeur des calculs v9.",
               "tableau": {"titre": None, "entetes": ["Identifiant", "Indicateur", "Formule", "Source de la méthode", "Méthode du FPPN", "Statut"], "largeurs": [1.9, 4.6, 9.4, 5.4, 5.6, 4.3], "taille": 8,
                           "lignes": [[x["id"], x["indicateur"], x["litterale"], x["source"], x["fppn"], x["statut"]] for x in _part]},
               "source": "Source : Adenka ; textes cités"})
SL.append({"type": "annexe", "titre": "Outils et bibliothèques", "notes": "", "tableau": {"titre": tab("Outils et bibliothèques utilisés"),
           "entetes": ["Outil", "Usage", "Version"], "largeurs": [7.0, 18.2, 6.0],
           "lignes": [["Python", "Langage des traitements, scripts conservés et rejouables", "3.14 (poste de la mission)"],
                      ["Visual Studio Code", "Éditeur et exécution des scripts sous Windows", "dernière"],
                      ["pandas, numpy", "Tables de données, calculs, agrégations", "3.0 et 2.5"],
                      ["statsmodels", "Régressions et tests statistiques", "0.15"],
                      ["python-calamine, openpyxl, lxml", "Lecture sans perte des fichiers Excel et des pages web", "dernière"],
                      ["pyarrow", "Stockage de la base d'analyse (Parquet)", "26.0"],
                      ["pymupdf, pdftotext, tesseract", "Lecture des PDF comptables et des textes scannés (OCR)", "dernière"],
                      ["pygments, Pillow", "Mise en forme des extraits de code", "dernière"],
                      ["python-pptx", "Rapport, graphiques et tableaux natifs", "1.0"]]},
           "source": "Source : Adenka"})
SL.append({"type": "liste_figures", "notes": ""})
SL.append({"type": "liste_tableaux", "notes": ""})
# ---- v10 : validité exacte par diapositive, renvois numérotés, retrait des mentions internes
from val_v10_fppn import VAL  # noqa: E402  (dictionnaire titre -> « Valable pour » exact)
for x in SL:
    if x["type"] == "contenu" and any(c in x for c in ("tableau", "graphique", "graphiques")):
        t_ = x["entete"]["titre"]
        if t_ in VAL:
            x["valable"] = VAL[t_]
for x in SL:
    if x["type"] == "contenu" and x["entete"].get("titre") == "Conformité prudentielle en bref":
        x["valable"] = "balances de 2020 au 31/08/2026 (provisoire) ; statut global jugé de 2022 à 2025"
_POS = {}
for i, x in enumerate(SL, 1):
    t_ = x["entete"].get("titre") if x["type"] == "contenu" else None
    if t_:
        assert t_ not in _POS, ("titre en double", t_)
        _POS[t_] = i
_NET = [(r"§§(.+?)§§", None), (r"(classeur des calculs|classeur) v\d+\b", "classeur des calculs"), (r"[Ee]mpreintes SHA-256", "empreintes numériques"), (r"empreinte SHA-256", "empreinte numérique"),
        (r"de la chaîne de traitement", "du traitement des données"), (r"la chaîne de traitement", "le traitement des données"), (r"Chaîne de traitement", "Traitement des données"),
        (r"chaîne de traitement", "traitement des données")]


def _nettoie(v):
    if isinstance(v, str):
        for motif, rempl in _NET:
            if rempl is None:
                def _r(m):
                    assert m.group(1) in _POS, ("renvoi introuvable", m.group(1))
                    return f"diapositive {_POS[m.group(1)]}"
                v = re.sub(motif, _r, v)
            else:
                v = re.sub(motif, rempl, v)
        return v
    if isinstance(v, list):
        return [_nettoie(e) for e in v]
    if isinstance(v, tuple):
        return tuple(_nettoie(e) for e in v)
    if isinstance(v, dict):
        return {k: _nettoie(e) for k, e in v.items()}
    return v


for i, x in enumerate(SL):
    SL[i] = _nettoie(x)
for i, x in enumerate(SL, 1):
    if len(x.get("source") or "") > 430:
        print("SOURCE LONGUE", i, len(x["source"]), (x.get("entete") or {}).get("titre", x.get("titre")))
TITRES_P = {x["num"]: x["titre"] for x in SL if x["type"] == "separateur"}
VAL_LIGNES = [[("Partie " + str(x["_partie"])) if x["_partie"] not in ("R", "A") else ("Résumé" if x["_partie"] == "R" else "Annexe"), x["entete"]["titre"], str(i), x["valable"]]
              for i, x in enumerate(SL, 1) if x["type"] == "contenu" and x.get("valable")]
for _k in range(0, len(VAL_LIGNES), 14):
    SL.append({"type": "annexe", "titre": f"Périmètre et validité des analyses ({_k // 14 + 1}/{(len(VAL_LIGNES) + 13) // 14})", "notes": "Niveau 3 de la validité : la ligne « Valable pour » de chaque diapositive. Colonne « Valable pour » dans l'onglet 9_figures_tableaux du classeur des calculs.",
               "tableau": {"titre": None, "entetes": ["Partie", "Analyse ou diapositive", "Diapo.", "Valable pour"], "largeurs": [2.6, 9.6, 3.0, 16.0], "taille": 8,
                           "lignes": VAL_LIGNES[_k:_k + 14]},
               "source": "Source : Adenka, d'après les fichiers cités dans chaque diapositive"})
_sig = sorted({(nom, defi) for nom, _, defi in SIGLES}, key=lambda t: t[0].lower())
_moit = (len(_sig) + 1) // 2
SL.append({"type": "annexe", "titre": "Signes et abréviations", "notes": "Les sigles de chaque diapositive sont aussi définis en bas de la diapositive.", "tableau": {"titre": None,
           "entetes": ["Sigle", "Signification", "Sigle", "Signification"], "largeurs": [2.6, 13.0, 2.6, 13.0], "taille": 8, "hauteur_ligne": 0.5,
           "lignes": [[_sig[i][0], _sig[i][1], _sig[i + _moit][0] if i + _moit < len(_sig) else "", _sig[i + _moit][1] if i + _moit < len(_sig) else ""] for i in range(_moit)]},
           "source": None})
SL.append({"type": "annexe", "titre": "Références et sources", "notes": "", "tableau": {"titre": None,
           "entetes": ["Référence", "Contenu utilisé", "Lien ou emplacement"], "largeurs": [10.2, 9.8, 11.2], "taille": 8,
           "lignes": [["Envoi du FPPN du 15/10/2026", "Précomptes, balances, exports du Grenier", "Dossier 04-Documents reçus du client"],
                      ["Grands livres 2021, 2022, 2024, 2025 (PDF)", "Détail des écritures", "Envoi complémentaire du 18/10/2026"],
                      ["Entretien du 19/10/2026 avec le responsable des cotisations", "Circuit des précomptes, règles des produits, nouvelles recrues", "FPPN_CR_entretien_cotisations_20261019.pptx"],
                      ["Entretien du 19/10/2026 (14h) avec le chef comptable", "Placements, autres actifs, ressources", "FPPN_CR_entretien_placements_20261019.pptx"],
                      ["Entretien du 19/10/2026 (16h) avec le chef comptable", "Provisions, allocations décès, retraite du personnel", "FPPN_CR_entretien_provisions_20261019.pptx"],
                      ["Règlement n° 07/2009/CM/UEMOA du 26/06/2009", "Cadre de la mutualité sociale", "airms-ci.com, documents publics"],
                      ["Règlement d'exécution n° 004/2022/COM/UEMOA du 01/06/2022", "Règles prudentielles, garantie, contrôle", "airms-ci.com, documents publics"],
                      ["Décision n° 05/2022/COM/UEMOA du 01/06/2022", "Indicateurs et ratios prudentiels", "airms-ci.com, documents publics"],
                      ["Règlement n° 03/2012/CM/UEMOA et PCMS (mai 2012)", "Plan comptable des mutuelles sociales", "airms-ci.com, documents publics"],
                      ["Arrêté n° 00115/MEPS/CAB du 12/12/2023", "Sanctions administratives", "airms-ci.com, documents publics"],
                      ["Arrêtés n° 092 et 093/MEPS/CAB du 06/09/2023", "Agrément, immatriculation, frais", "airms-ci.com, documents publics"],
                      ["Répertoire AIRMS des mutuelles agréées (15/10/2026)", "Agrément et immatriculation du FPPN", "airms-ci.com, documents publics"],
                      ["Ordonnance n° 2024-1043", "FOMUS-CI", "civlii.laws.africa"],
                      ["Code CIMA, article 338", "Tables et taux technique de 3,5 %", "cima-afrique.org"],
                      ["Tables CIMA H, CIMA F, TD, TV ; table ronde CIMA FANAF (2011)", "Mortalité de référence", "Fichiers fournis le 18/10/2026 ; fanaf.org"],
                      ["BCEAO, communiqués du CPM 2020 à 2025", "Taux directeurs", "bceao.int"],
                      ["Banque mondiale, FP.CPI.TOTL.ZG", "Inflation 2020 à 2025", "donnees.banquemondiale.org"],
                      ["Code CIMA, art. 334-12 ; circulaire n° 00230/CIMA/CRCA/PDT/2005 du 24/10/2005", "Provision pour sinistres déclarés tardifs : méthode de la cadence des déclarations tardives", "Extrait du Code CIMA 2019 reçu le 23/10/2026 (pages 523 et 524)"],
                      ["BCEAO, instruction sur les créances en souffrance des SFD", "Barème de dépréciation des prêts", "bceao.int, réglementation des SFD (texte à joindre)"],
                      ["Décret n° 2017-210 du 30/03/2017 ; convention collective interprofessionnelle du 19/07/1977, art. 40", "Indemnité de départ à la retraite", "Textes fournis le 19/10/2026"],
                      ["Statuts et règlement intérieur du FPPN (15/06/2023)", "Barème, provisions, réserves, produits, personnel", "Textes fournis le 19/10/2026"],
                      ["BCEAO, Rapport sur la Politique Monétaire dans l'UMOA, mars 2026", "Inflation de la santé (tableau 12), taux des titres publics (17) et des dépôts à terme (22)", "bceao.int"],
                      ["ANStat, IHPC par poste de consommation (national)", "Inflation médicale de la Côte d'Ivoire", "Extraction du 19/10/2026"],
                      ["Investing.com, rendements des obligations de l'État de Côte d'Ivoire à 5 et 10 ans", "Référence de marché du rendement des placements (source secondaire)", "Exports du 19/10/2026, cotations de mars 2021 à octobre 2026"],
                      ["Documents reçus le 22/10/2026 (124 fichiers)", "États financiers, rapports de gestion, cahiers des charges, manuel d'épargne, étude des ratios, DISA, grands livres, rapprochements bancaires, conventions, états des allocations décès, arrêté du CMP, dérogation de Cocody, désignation du directeur général", "Noms exacts en annexe « Fichiers sources »"],
                      ["Classeur des calculs", "Tous les calculs du rapport, expliqués onglet par onglet : tableaux, ratios, provisions, formules, sensibilités, provision du personnel, mortalité, EFC, contrôles croisés, revue actuarielle (onglets 18 à 25), traçabilité et validité", "FPPN_Calculs_rapport_20261022.xlsx"]]},
           "source": None})
SL.append({"type": "merci", "notes": ""})

# ===========================================================================
# CONSTRUCTION
# ===========================================================================
GAB = {"couverture": 1, "sommaire": 5, "separateur": 2, "contenu": 5, "annexe": 5, "liste_figures": 5, "liste_tableaux": 5, "merci": 110}
tmp = SORTIE.replace(".pptx", "_tmp.pptx")
selection(TIMESAVER, tmp, [x["gab"] if x["type"] == "gab" else GAB[x["type"]] for x in SL])


def paragraphes(x):
    """Paragraphes d'analyse d'une diapositive : constat et preuves, puis l'implication dans un paragraphe distinct."""
    paras = list(x.get("analyse") or x.get("texte") or [])
    if x.get("implication"):
        paras.append(x["implication"])
    return paras


def _brut(t):
    return str(t).replace("**", "").replace("[[", "").replace("]]", "")


def longueur(paras):
    return sum(len(_brut(p)) for p in paras)


JAUNE_SURL = "FFFF00"


def surligner(r):
    """Surlignage jaune d'un morceau de texte (affirmation à confirmer par une pièce ou une réponse du FPPN)."""
    rPr = r._r.get_or_add_rPr()
    for h in rPr.findall(qn("a:highlight")):
        rPr.remove(h)
    hl = rPr.makeelement(qn("a:highlight"), {})
    clr = hl.makeelement(qn("a:srgbClr"), {"val": JAUNE_SURL})
    hl.append(clr)
    latin = rPr.find(qn("a:latin"))
    if latin is not None:
        latin.addprevious(hl)
    else:
        rPr.append(hl)


def morceaux_balises(texte):
    """Découpe un texte en (morceau, gras, surligné) selon les balises **…** et [[…]]."""
    out = []
    for m1 in re.split(r"(\[\[.+?\]\])", str(texte)):
        if not m1:
            continue
        surl = m1.startswith("[[") and m1.endswith("]]")
        m1 = m1[2:-2] if surl else m1
        for m2 in re.split(r"(\*\*.+?\*\*)", m1):
            if not m2:
                continue
            gras = m2.startswith("**") and m2.endswith("**")
            out.append((m2[2:-2] if gras else m2, gras, surl))
    return out


def disposition_graphique(paras):
    L = longueur(paras)
    if L <= 650:
        return 18.2, 20.2, 12.4
    if L <= 1000:
        return 16.6, 18.6, 14.0
    return 15.0, 17.0, 15.6


CJ = {}
for i, x in enumerate(SL, 1):
    k = str(i)
    if x["type"] == "couverture":
        CJ[k] = {"Google Shape;138;p25": "Audit Comptable et financier FPPN", "Text Placeholder 4": "23 octobre 2026"}
    elif x["type"] == "separateur":
        CJ[k] = {"Rectangle 24": str(x["num"]), "Rectangle 9": x["titre"], "Rectangle 5": x["sous_titre"]}
    elif x["type"] in ("sommaire", "liste_figures", "liste_tableaux", "annexe"):
        CJ[k] = {"@vider": True}
    elif x["type"] == "gab":
        # Les gabarits mis à l'échelle par compresser() ne doivent pas être réduits une première fois par l'en-tête
        CJ[k] = {"@entete": dict(x["entete"], ajuster=x["gab"] not in (9, 13, 29, 31, 39, 41))}
        if x.get("source"):
            CJ[k]["@source"] = x["source"]
    elif x["type"] == "contenu":
        d = {"@vider": True, "@entete": x["entete"]}
        if "graphique" in x:
            lg, _, _ = disposition_graphique(paragraphes(x))
            g = dict(x["graphique"])
            d["@texte1"] = {"texte": g.pop("titre"), "style": "titre_graphique", "x": 1.4, "y": 3.85, "largeur": lg, "hauteur": 0.7}
            if g.get("unite"):
                d["@texte3"] = {"texte": g.pop("unite"), "style": "unite", "x": 1.4, "y": 4.45, "largeur": lg, "hauteur": 0.45}
            g.update({"x": 1.4, "y": 4.95, "largeur": lg, "hauteur": BAS_CONTENU - 4.95 - _hv(x, lg)})
            d["@graphique"] = g
        if "graphiques" in x:
            hg = ((13.4 - 4.95) if x.get("bandeau") else (BAS_CONTENU - 4.95)) - _hv(x)
            for j, gg in enumerate(x["graphiques"]):
                g = dict(gg)
                xx = 1.4 + j * 16.0
                d[f"@texte1{j}"] = {"texte": g.pop("titre"), "style": "titre_graphique", "x": xx, "y": 3.85, "largeur": 15.2, "hauteur": 0.7}
                if g.get("unite"):
                    d[f"@texte3{j}"] = {"texte": g.pop("unite"), "style": "unite", "x": xx, "y": 4.45, "largeur": 15.2, "hauteur": 0.45}
                g.update({"x": xx, "y": 4.95, "largeur": 15.2, "hauteur": hg})
                d[f"@graphique{j}"] = g
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


def hauteur_texte(paras, largeur, taille, espace=6):
    cpl = max(1.0, (largeur - 0.3) / (0.0181 * taille))
    lignes = sum(math.ceil(len(_brut(p)) / cpl) for p in paras)
    return lignes * taille * 1.22 * 0.03528 + len(paras) * espace * 0.03528


def corps(slide, x, y, l, h, paras, tmax=14, tmin=9):
    """Texte d'analyse justifié, **gras** accepté ; la taille est la plus grande qui tient dans la zone."""
    if not paras or h < 0.6:
        return None
    taille = tmin
    for t in (14, 13, 12, 11.5, 11, 10.5, 10, 9.5, 9):
        if t <= tmax and hauteur_texte(paras, l, t) <= h:
            taille = t
            break
    tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(l), Cm(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = tf.margin_right = Cm(0.15)
    for i, bloc in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.JUSTIFY
        p.space_after = Pt(6)
        for morceau, gras, surl in morceaux_balises(bloc):
            r = p.add_run()
            r.text = morceau
            r.font.name = "Helvetica"
            r.font.size = Pt(taille)
            r.font.bold = gras
            r.font.color.rgb = RGBColor.from_string(NOIR)
            if surl:
                surligner(r)
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


COUL_STATUT = [("Conforme", BLEU, "FFFFFF"), ("Non conforme", ROUGE, "FFFFFF"), ("Au-dessus", ORANGE, NOIR),
               ("En attente", GRIS_C, NOIR), ("À apprécier", GRIS_C, NOIR)]


VERT_C, JAUNE_C = "70AD47", "FFD966"
# ---- v12 : codes couleur par boules (pas de remplissage). Police des boules : Segoe UI Emoji (Windows), repli automatique ailleurs
BOULE = {ROUGE: "🔴", ORANGE: "🟠", JAUNE_C: "🟡", VERT_C: "🟢", GRIS_C: "⚪"}
LECT83 = {}  # texte du bandeau « Comment lire » de la carte des risques


def runs_texte(p, texte, taille, gras, coul, surl=False):
    """Écrit un texte en runs : les boules dans une police d'émojis, le reste en Helvetica."""
    for seg in re.split("([" + BOULES + "])", str(texte)):
        if not seg:
            continue
        r = p.add_run()
        r.text = seg
        emoji = seg in BOULES
        r.font.name = "Segoe UI Emoji" if emoji else "Helvetica"
        r.font.size = Pt(taille + 1 if emoji else taille)
        r.font.bold = False if emoji else gras
        r.font.color.rgb = RGBColor.from_string(coul)
        if surl and not emoji:
            surligner(r)


def encadre(slide, x, y, l, h, lignes, titre="Définition et lecture"):
    """Encadré gris à filet orange : titre, puis une ligne par définition (**gras** accepté)."""
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(x), Cm(y), Cm(l), Cm(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = RGBColor.from_string("F4F4F4")
    sh.line.color.rgb = RGBColor.from_string(ORANGE)
    sh.line.width = Pt(1.25)
    sh.shadow.inherit = False
    tf = sh.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = tf.margin_right = Cm(0.3)
    tf.margin_top = tf.margin_bottom = Cm(0.15)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    runs_texte(p, titre, 10.5, True, ORANGE)
    for li in lignes:
        p = tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(3)
        for morceau, gras, surl in morceaux_balises(li):
            runs_texte(p, morceau, 10, gras, NOIR, surl)
    return y + h
MOTS_COL = ("Statut", "Priorité", "Criticité", "Appréciation", "Verdict", "Horizon", "Conforme", "Reçu", "Disponible", "Utilisable", "Résolu", "État")
CODES = [  # (début de valeur en minuscules, fond, police, libellé de légende, sens)
    ("non conforme", ROUGE, "FFFFFF", "Non conforme", "norme non respectée"),
    ("non remplie", ROUGE, "FFFFFF", "Non remplie", "condition non remplie"),
    ("non démontr", ORANGE, NOIR, "Non démontré", "non démontré par les documents reçus"),
    ("non reçu", ROUGE, "FFFFFF", "Non reçu", "document non reçu"),
    ("non communiqu", ROUGE, "FFFFFF", "Non communiquée", "méthode du FPPN non transmise"),
    ("non mesur", ORANGE, NOIR, "Non mesuré", "donnée absente"),
    ("insuffisant", ROUGE, "FFFFFF", "Insuffisant", "exigence non satisfaite"),
    ("critique", ROUGE, "FFFFFF", "Critique", "non-conformité réglementaire, impact stratégique majeur, risque élevé"),
    ("urgent", ROUGE, "FFFFFF", "Urgent", "à traiter en premier"),
    ("haute", ROUGE, "FFFFFF", "Haute", "à traiter en premier"),
    ("court terme", ROUGE, "FFFFFF", "Court terme", "moins de 6 mois"),
    ("non", ROUGE, "FFFFFF", "Non", "réponse négative à la question de la colonne"),
    ("partiel", ORANGE, NOIR, "Partiel", "en partie seulement"),
    ("élevé", ORANGE, NOIR, "Élevé", "impact opérationnel significatif, risque important"),
    ("moyen terme", ORANGE, NOIR, "Moyen terme", "6 à 18 mois"),
    ("divergent", ORANGE, NOIR, "Divergent", "formule différente du texte"),
    ("au-dessus", ORANGE, NOIR, "Au-dessus", "au-dessus de la norme"),
    ("grade seulement", ORANGE, NOIR, "Partiel", "en partie seulement"),
    ("passées seulement", ORANGE, NOIR, "Partiel", "en partie seulement"),
    ("table seulement", ORANGE, NOIR, "Partiel", "en partie seulement"),
    ("en attente", JAUNE_C, NOIR, "En attente", "document du FPPN attendu"),
    ("à apprécier", GRIS_C, NOIR, "À apprécier", "norme non chiffrée"),
    ("long terme", VERT_C, NOIR, "Long terme", "plus de 18 mois"),
    ("faible", VERT_C, NOIR, "Faible", "opportunité d'optimisation ou de bonne pratique"),
    ("sans texte", GRIS_C, NOIR, "Sans texte applicable", "aucun texte ne fixe la règle"),
    ("conforme", VERT_C, NOIR, "Conforme", "norme respectée"),
    ("oui", VERT_C, NOIR, "Oui", "réponse positive à la question de la colonne"),
    ("reçu", VERT_C, NOIR, "Reçu", "document reçu"),
    ("remplie", VERT_C, NOIR, "Remplie", "condition remplie"),
    ("démontré", VERT_C, NOIR, "Démontré", "démontré par les documents"),
    ("résolu", VERT_C, NOIR, "Résolu", "réponse établie par les documents"),
]


def code_couleur(entete, v):
    """Couleur d'une cellule de statut : renvoie (fond, police, libellé, sens) ou None."""
    e, t = str(entete), str(v).strip().lower()
    if not t or t in ("-", "sans objet"):
        return None
    if "criticité" in e.lower() and t.startswith("moyen"):
        return JAUNE_C, NOIR, "Moyen", "amélioration souhaitable, impact ponctuel"
    if "priorité" in e.lower() and t.startswith("faible"):
        return VERT_C, NOIR, "Faible", "à traiter en dernier"
    if t.startswith(("moyenne", "moyen")) and not t.startswith("moyen terme"):
        return ORANGE, NOIR, "Moyenne", "à traiter ensuite"
    for debut, f_, c_, lib, sens in CODES:
        if t.startswith(debut):
            return f_, c_, lib, sens
    return None


def colonnes_codees(spec):
    """Colonnes de statut d'un tableau : en-tête reconnu et au moins une valeur codée."""
    cols = []
    for j, e in enumerate(spec["entetes"]):
        if spec.get("sans_code"):
            break
        if any(m.lower() in str(e).lower() for m in MOTS_COL) or str(e).strip().endswith("?") or spec.get("statut_col") == j:
            vals = [l[j] for l in spec["lignes"] if j < len(l)]
            if vals and sum(1 for v in vals if code_couleur(e, v)) >= max(1, len(vals) // 2):
                cols.append(j)
    return cols


FAMILLES = [  # échelles complètes : la légende montre toute l'échelle, même si le tableau n'en utilise qu'une partie
    [("🔴", "Critique", "non-conformité réglementaire, impact stratégique majeur, risque élevé"), ("🟠", "Élevé", "impact opérationnel significatif, risque important"),
     ("🟡", "Moyen", "amélioration souhaitable, impact ponctuel"), ("🟢", "Faible", "opportunité d'optimisation ou de bonne pratique")],
    [("🔴", "Urgent", "à traiter en premier"), ("🟠", "Moyenne", "à traiter ensuite"), ("🟢", "Faible", "à traiter en dernier")],
    [("🔴", "Court terme", "moins de 6 mois"), ("🟠", "Moyen terme", "6 à 18 mois"), ("🟢", "Long terme", "plus de 18 mois")],
    [("🔴", "Non conforme", "norme non respectée"), ("🟠", "Non démontré", "non démontré par les documents reçus"), ("🟡", "En attente", "document du FPPN attendu"),
     ("⚪", "À apprécier", "norme non chiffrée"), ("🟢", "Conforme", "norme respectée")],
    [("🔴", "Non", "réponse négative à la question de la colonne"), ("🟠", "Partiel", "en partie seulement"), ("🟢", "Oui", "réponse positive à la question de la colonne")],
    [("🔴", "Non reçu", "document non reçu"), ("🟠", "Partiel", "en partie seulement"), ("🟢", "Reçu", "document reçu")],
    [("🔴", "Non remplie", "condition non remplie"), ("🟠", "Partiel", "en partie seulement"), ("🟢", "Remplie", "condition remplie")],
]


def legende_codes(slide, spec, cols, x0, y, largeur):
    """Tableau de signification des boules, sous le tableau : l'échelle complète de la famille de codes utilisée."""
    vus = [tuple(v) for v in (spec.get("legende") or [])]
    if not vus:
        presents = []
        for j in cols:
            for l in spec["lignes"]:
                cc = code_couleur(spec["entetes"][j], l[j]) if j < len(l) else None
                if cc and cc[2] not in [v[1] for v in presents]:
                    presents.append((BOULE.get(cc[0], ""), cc[2], cc[3]))
        libs = {v[1] for v in presents}
        fam = max(FAMILLES, key=lambda f_: sum(1 for v in f_ if v[1] in libs))
        if not any(v[1] in libs for v in fam):
            fam = FAMILLES[3]  # échelle de conformité par défaut (tableaux de formules)
        vus = list(fam) + [v for v in presents if v[1] not in {w[1] for w in fam}]
    if not vus:
        return y
    par_ligne = max(1, min(len(vus), int(largeur // 6.0)))
    lignes_ = [vus[k:k + par_ligne] for k in range(0, len(vus), par_ligne)]
    h = 0.45
    shp = slide.shapes.add_table(len(lignes_), 1 + 2 * par_ligne, Cm(x0), Cm(y), Cm(largeur), Cm(h * len(lignes_)))
    t = shp.table
    t._tbl.tblPr.set("firstRow", "0")
    t._tbl.tblPr.set("bandRow", "0")
    lc = (largeur - 2.2) / par_ligne
    t.columns[0].width = Cm(2.2)
    for k in range(par_ligne):
        t.columns[1 + 2 * k].width = Cm(0.7)
        t.columns[2 + 2 * k].width = Cm(lc - 0.7)
    for i, lg in enumerate(lignes_):
        t.rows[i].height = Cm(h)
        cells = [("Code couleur" if i == 0 else "", "", True)]
        for boule, lib, sens in lg:
            cells += [(boule, "", False), (lib, " : " + sens, True)]
        cells += [("", "", False)] * (1 + 2 * par_ligne - len(cells))
        for j, (texte, suite, g_) in enumerate(cells):
            c = t.cell(i, j)
            c.margin_left = c.margin_right = Cm(0.06)
            c.margin_top = c.margin_bottom = Cm(0.01)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = c.text_frame.paragraphs[0]
            p.text = ""
            if texte in BOULES and texte:
                p.alignment = PP_ALIGN.CENTER
            runs_texte(p, texte, 7, g_, NOIR)
            if suite:
                runs_texte(p, suite, 7, False, NOIR)
            c.fill.solid()
            c.fill.fore_color.rgb = RGBColor.from_string("FFFFFF")
    return y + h * len(lignes_)


def hauteurs_lignes(spec, taille):
    hs = []
    for i, vals in enumerate([spec["entetes"]] + spec["lignes"]):
        h = 0
        for v, w in zip(vals, spec["largeurs"]):
            cpl = max(1.0, (w - 0.3) / (0.0186 * taille))
            n = sum(max(1, math.ceil(len(_brut(seg)) / cpl)) for seg in str(v).split("\n"))
            h = max(h, n * taille * 1.2 * 0.03528 + 0.16 + (0.22 if any(b_ in str(v) for b_ in BOULES) else 0))
        hs.append(max(h, spec.get("hauteur_ligne") or 0.5))
    return hs


def tableau(slide, spec, y0=3.85, x0=1.4, taille=None, hauteur_max=None):
    """Tableau natif ; renvoie le bas du tableau (cm). La police diminue si le tableau dépasse la hauteur disponible."""
    y = y0
    if spec.get("titre"):
        txt(slide, x0, y, sum(spec["largeurs"]), 0.6, spec["titre"], 12, False, ORANGE, PP_ALIGN.CENTER)
        y += 0.75
    lignes, larg = spec["lignes"], spec["largeurs"]
    taille = taille or spec.get("taille") or (11 if len(lignes) <= 8 else 10)
    hmax = (hauteur_max or (BAS_CONTENU - y))
    cols_c = colonnes_codees(spec)
    if cols_c or spec.get("legende"):
        hmax -= 0.62
    hs = hauteurs_lignes(spec, taille)
    while sum(hs) > hmax and taille > 7.5:
        taille -= 0.5
        hs = hauteurs_lignes(spec, taille)
    n = len(lignes) + 1
    shape = slide.shapes.add_table(n, len(larg), Cm(x0), Cm(y), Cm(sum(larg)), Cm(sum(hs)))
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
        t.rows[i].height = Cm(hs[i])
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
            fond, coul = (NOIR, "FFFFFF") if i == 0 else ((GRIS if i % 2 == 0 else "FFFFFF"), NOIR)
            gras_c = (i == 0) or (j == 0)
            if i > 0 and j in cols_c:
                cc = code_couleur(spec["entetes"][j], v)
                if cc:
                    b_, t_ = BOULE.get(cc[0], ""), str(v).strip()
                    v = b_ if (len(t_) <= 32 and ":" not in t_ and "\n" not in t_) else b_ + " " + t_
            if str(v) in BOULES or (spec.get("centre") and j > 0 and i > 0):
                p.alignment = PP_ALIGN.CENTER
            if i > 0 and spec.get("fonds") and (i - 1, j) in spec["fonds"]:
                fond = spec["fonds"][(i - 1, j)]
                coul, gras_c = ("FFFFFF" if fond == "C00000" else NOIR), True
            for li_, ligne_ in enumerate(str(v).split("\n")):
                if li_:
                    p = c.text_frame.add_paragraph()
                    p.alignment = c.text_frame.paragraphs[0].alignment
                for morceau, gras, surl in morceaux_balises(ligne_):
                    runs_texte(p, morceau, taille, gras_c or gras, coul, surl)
            c.fill.solid()
            c.fill.fore_color.rgb = RGBColor.from_string(fond)
    if cols_c or spec.get("legende"):
        return legende_codes(slide, spec, cols_c, x0, y + sum(hs) * 1.07 + 0.12, sum(larg))
    return y + sum(hs)


def valable(slide, x, y, largeur):
    """Niveau 2 de la validité : ligne « Valable pour » sous le tableau ou le graphique."""
    if x.get("valable"):
        txt(slide, 1.4, y, largeur, _hv(x, largeur) - 0.1, "Valable pour : " + x["valable"] + ".", 8.5, False, GRIS_F)
        return y + _hv(x, largeur)
    return y


def tableau_avec_analyse(slide, x):
    spec, paras = dict(x["tableau"]), paragraphes(x)
    enc = x.get("encadre")
    h_enc = (hauteur_texte(enc, 30.2, 10, 3) + 0.95) if enc else 0.0
    if not paras:
        bas = tableau(slide, spec, hauteur_max=BAS_CONTENU - 3.85 - 0.75 - 0.25 - _hv(x, sum(spec["largeurs"])) - (h_enc + 0.25 if enc else 0))
        bas = valable(slide, x, bas + 0.3, sum(spec["largeurs"]))
        if enc:
            encadre(slide, 1.4, bas + 0.25, 31.2, h_enc, enc, x.get("encadre_titre", "Définition et lecture"))
        return
    if x.get("disposition") == "droite":
        larg = spec["largeurs"]
        if sum(larg) > 19.6:
            f_ = 19.6 / sum(larg)
            spec["largeurs"] = [round(w * f_, 2) for w in larg]
        bas = tableau(slide, spec, hauteur_max=BAS_CONTENU - 3.85 - 0.75 - 0.25 - _hv(x, sum(spec["largeurs"])))
        valable(slide, x, bas + 0.3, sum(spec["largeurs"]))
        xt = 1.4 + sum(spec["largeurs"]) + 0.6
        y_d = 3.85
        if enc:
            h_d = hauteur_texte(enc, 32.6 - xt - 0.8, 10, 3) + 0.95
            encadre(slide, xt, y_d, 32.6 - xt, h_d, enc, x.get("encadre_titre", "Définition et lecture"))
            y_d += h_d + 0.3
        corps(slide, xt, y_d, 32.6 - xt, BAS_CONTENU - y_d, paras)
    else:
        besoin = max(1.6, min(4.2, hauteur_texte(paras, 31.2, 11) + 0.2))
        bas = tableau(slide, spec, hauteur_max=BAS_CONTENU - 3.85 - 0.75 - besoin - 0.55 - _hv(x) - (h_enc + 0.25 if enc else 0))
        bas = valable(slide, x, bas + 0.3, 31.2)
        if enc:
            encadre(slide, 1.4, bas + 0.25, 31.2, h_enc, enc, x.get("encadre_titre", "Définition et lecture"))
            bas += h_enc + 0.25
        corps(slide, 1.4, bas + 0.3, 31.2, BAS_CONTENU - bas - 0.3, paras, 13)


def pastille(slide, x, y, num):
    return boite(slide, x, y, 0.75, 0.75, str(num), ROUGE, "FFFFFF", 11, True, MSO_SHAPE.OVAL)


def pas_a_pas(slide, spec, paras):
    lignes = spec["lignes"]
    taille = 10 if not spec.get("petit") else 9
    txt(slide, 1.4, 3.75, sum(spec["largeurs"]), 0.5, "Extrait : " + spec["fichier"], 9, False, GRIS_F)
    t_spec = {"titre": None, "entetes": spec["entetes"], "lignes": lignes, "largeurs": spec["largeurs"], "hauteur_ligne": 0.55}
    y_tab = 4.3
    hs = hauteurs_lignes(t_spec, taille)
    bas = tableau(slide, t_spec, y_tab, 1.4, taille)
    larg = spec["largeurs"]
    for k, (ri, cj) in enumerate(spec["rouge"], 1):
        xx = 1.4 + sum(larg[:cj])
        yy = y_tab + sum(hs[:ri + 1])
        cadre = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(xx - 0.05), Cm(yy - 0.03), Cm(larg[cj] + 0.1), Cm(hs[ri + 1] + 0.06))
        cadre.fill.background()
        cadre.line.color.rgb = RGBColor.from_string(ROUGE)
        cadre.line.width = Pt(2.25)
        pastille(slide, xx + larg[cj] - 0.5, yy - 0.45, k)
    xr = 1.4 + sum(larg) + 0.6
    lr = 33.0 - xr - 0.4
    if lr < 7:
        xr, lr = 1.4, sum(larg)
        y = bas + 0.35
    else:
        y = 4.3
    for k, e in enumerate(spec["explications"], 1):
        h = max(0.8, hauteur_texte([e], lr - 1.0, 12) + 0.1)
        pastille(slide, xr, y, k)
        txt(slide, xr + 1.0, y - 0.05, lr - 1.0, h, e, 12)
        y += h + 0.25
    if lr == sum(larg):           # explications sous le tableau : l'analyse suit, pleine largeur
        corps(slide, 1.4, y + 0.1, 31.2, BAS_CONTENU - y - 0.1, paras, 12)
    else:
        y_an = max(y + 0.2, 4.3)
        if BAS_CONTENU - y_an >= 2.0:
            corps(slide, xr, y_an, lr, BAS_CONTENU - y_an, paras, 12)
        else:
            corps(slide, 1.4, bas + 0.3, sum(larg), BAS_CONTENU - bas - 0.3, paras, 12)


def bandeau_attente(slide, y=15.15):
    boite(slide, 1.4, y, 31.2, 0.95, "Travail en attente des données : voir la liste des données manquantes (PBC pointée et partie 14)", ORANGE, NOIR, 13)


def triangle(slide):
    ent = ["Survenance", "Délai 1", "Délai 2", "Délai 3", "Délai 4"]
    lignes = [["Année n-3", "C(1,1)", "C(1,2)", "C(1,3)", "C(1,4)"], ["Année n-2", "C(2,1)", "C(2,2)", "C(2,3)", "?"],
              ["Année n-1", "C(3,1)", "C(3,2)", "?", "?"], ["Année n", "C(4,1)", "?", "?", "?"]]
    tableau(slide, {"titre": None, "entetes": ent, "lignes": lignes, "largeurs": [4.6, 3.6, 3.6, 3.6, 3.6], "hauteur_ligne": 1.1}, 4.3, 1.4, 14)
    txt(slide, 1.4, 10.4, 19.0, 4.4, "C(i, j) : paiements cumulés de l'année de survenance i, au délai j.\n« ? » : montants futurs à estimer.\nFacteur de développement : f(j) = Σ C(i, j+1) / Σ C(i, j).\nCharge ultime = dernier montant connu × produit des facteurs suivants.\nProvision = charge ultime - paiements déjà faits.", 13)
    boite(slide, 21.6, 4.3, 11.0, 4.9, "Hypothèse\nLe rythme de règlement passé se répète dans le futur.", GRIS, NOIR, 15, True, MSO_SHAPE.ROUNDED_RECTANGLE)
    boite(slide, 21.6, 9.6, 11.0, 4.9, "Données\nPaiements par année de survenance et par délai, issus des bases de prestations (PBC 13, 28, 40).", GRIS, NOIR, 15, True, MSO_SHAPE.ROUNDED_RECTANGLE)
    bandeau_attente(slide)


def schema(slide, nom, paras):
    if nom == "fonctionnement":
        boite(slide, 1.4, 6.9, 6.0, 3.2, "Mutualistes\nactifs et retraités", BLEU, "FFFFFF", 16)
        boite(slide, 9.0, 4.4, 6.2, 2.4, "Solde (actifs)", GRIS, NOIR, 14, contour=BLEU)
        boite(slide, 9.0, 10.2, 6.2, 2.4, "CGRAE (retraités)", GRIS, NOIR, 14, contour=BLEU)
        boite(slide, 16.8, 6.5, 6.6, 4.0, "FPPN\nCotisations 2025\n" + F("Cot (schéma)", s[2025]["cot_maladie"] + s[2025]["cot_deces"]), ORANGE, NOIR, 16)
        sorties = [("Prestations maladie\n" + F("Prest mal (schéma)", s[2025]["prest_maladie"]), 5.0),
                   ("Prestations décès\n" + F("Prest déc (schéma)", s[2025]["prest_deces"]), 7.6),
                   ("Épargne PEP et EFC\n" + F("Épargne (schéma)", b[2026]["epargne_totale"]), 10.2),
                   ("Prêts du Grenier\n" + F("Grenier (schéma)", b[2026]["grenier_financement"]), 12.8),
                   ("CMP et terrains\n" + F("CMP+terrains (schéma)", b[2026]["cmp_investissements"] + b[2026]["terrains"]), 15.1)]
        for t_, yy in sorties:
            boite(slide, 25.2, yy - 1.05, 7.4, 2.1, t_, BLEU, "FFFFFF", 12)
            fleche(slide, 23.4, 8.5, 25.2, yy, GRIS_F)
        fleche(slide, 7.4, 7.9, 9.0, 5.6, GRIS_F)
        fleche(slide, 7.4, 9.1, 9.0, 11.4, GRIS_F)
        fleche(slide, 15.2, 5.6, 16.8, 7.7, GRIS_F)
        fleche(slide, 15.2, 11.4, 16.8, 9.3, GRIS_F)
        corps(slide, 1.4, 13.2, 22.6, BAS_CONTENU - 13.2, paras, 12)
    elif nom == "precomptes":
        etapes_ = [("1. FPPN", "Le FPPN prépare le fichier aller : qui précompter et combien."), ("2. Solde / CGRAE", "La Solde ou la CGRAE retient les montants sur le salaire ou la pension."),
                   ("3. Fichier retour", "Le fichier retour liste les montants réellement retenus."), ("4. FPPN", "Le FPPN rapproche aller et retour, comptabilise et relance les rejets.")]
        for k, (t_, c_) in enumerate(etapes_):
            x = 1.4 + k * 8.0
            boite(slide, x, 4.6, 6.6, 2.2, t_, BLEU if k % 2 == 0 else ORANGE, "FFFFFF" if k % 2 == 0 else NOIR, 16)
            txt(slide, x, 7.1, 6.6, 2.4, c_, 13, False, NOIR, PP_ALIGN.CENTER)
            if k < 3:
                fleche(slide, x + 6.6, 5.7, x + 8.0, 5.7, GRIS_F, 3)
        boite(slide, 1.4, 9.8, 31.2, 1.9, "Contrôle Adenka : taux de précompte = montant retour / montant aller, poste par poste et mois par mois (partie 3)", GRIS, NOIR, 14, contour=BLEU)
        corps(slide, 1.4, 12.2, 31.2, BAS_CONTENU - 12.2, paras, 13)


def compresser(slide, y0, y1, n0, n1, uniforme=False):
    """Ramène dans [n0, n1] les formes du gabarit situées dans [y0, y1] (cm) : place de l'en-tête et des sigles.
    En mode uniforme, chaque forme est réduite dans les deux sens autour du centre de la diapositive."""
    f_ = (n1 - n0) / (y1 - y0)
    for sh in slide.shapes:
        if sh.top is None:
            continue
        top, h = sh.top / 360000, sh.height / 360000
        if h > 18 or top < y0 - 0.05 or top > y1 or sh.name.startswith(("ZoneTexte", "Title", "Adenka_")):
            continue
        centre = top + h / 2
        if uniforme:
            l_, w_ = sh.left / 360000, sh.width / 360000
            cx = 16.93 + (l_ + w_ / 2 - 16.93) * f_
            cy = n0 + (centre - y0) * f_
            sh.width, sh.height = Cm(w_ * f_), Cm(max(h * f_, 0.01))
            sh.left, sh.top = Cm(cx - w_ * f_ / 2), Cm(cy - h * f_ / 2)
            continue
        if sh.shape_type in (1, 9, 17) and not sh.name.startswith("Google Shape"):
            nt, nh = n0 + (top - y0) * f_, h * f_
            sh.top, sh.height = Cm(nt), Cm(max(nh, 0.01))
        else:
            sh.top = Cm(n0 + (centre - y0) * f_ - h / 2)


def ecrire(sh, lignes_, tailles=None, gras=None):
    tf = sh.text_frame
    paras = [p for p in tf.paragraphs]
    for k, ligne in enumerate(lignes_):
        if k < len(paras):
            p = paras[k]
            runs = p.runs
            if runs:
                runs[0].text = ligne
                for r in runs[1:]:
                    r._r.getparent().remove(r._r)
            else:
                p.text = ligne
        else:
            p = tf.add_paragraph()
            p.text = ligne
        for r in p.runs:
            if tailles:
                r.font.size = Pt(tailles[min(k, len(tailles) - 1)])
            if gras is not None:
                r.font.bold = gras[min(k, len(gras) - 1)]
    for p in paras[len(lignes_):]:
        p._p.getparent().remove(p._p)


def remplir_gab(slide, num, d, analyse=None):
    formes = list(slide.shapes)

    def par_pos(pred):
        return sorted([sh for sh in formes if pred(sh)], key=lambda sh: (round(sh.top / 360000), sh.left))

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
        if d.get("lecture"):
            LECT83[id(slide)] = d["lecture"]
    elif num in (39, 41):
        compresser(slide, 3.5, 16.2, 4.0, 14.9 if d.get("attente") else 16.1)
        formes = list(slide.shapes)
        nums = {sh.text_frame.text.strip(): sh for sh in formes if sh.has_text_frame and sh.text_frame.text.strip() in ("01", "02", "03", "04", "05", "06")}
        noms_r = ("Rectangle 18", "Rectangle 19", "Rectangle 21", "Rectangle 22", "Rectangle 28", "Rectangle 29")
        rects = [sh for sh in formes if sh.name in noms_r]
        for k, (t_, c_) in enumerate(d["etapes"], 1):
            nsh = nums[f"0{k}"]
            r = min(rects, key=lambda sh: abs(sh.top - nsh.top) + abs(sh.left - nsh.left))
            ecrire(r, [t_, c_], [14, 12] if num == 39 else [13, 11], [True, False])
        if d.get("attente"):
            bandeau_attente(slide)
    elif num == 9:
        for sh in formes:
            if sh.name == "Rectangle 27":
                ecrire(sh, [""])
        compresser(slide, 2.0, 17.2, 3.75, 14.85 if d.get("attente") else 16.25)
        formes = list(slide.shapes)
        blocs = par_pos(lambda sh: sh.name in ("Rectangle 22", "Rectangle 23", "Rectangle 32", "Rectangle 33"))
        for sh, (t_, c_) in zip(blocs, d["blocs"]):
            ecrire(sh, [t_, c_], [13, 11.5], [True, False])
            sh.text_frame.word_wrap = True
        if d.get("attente"):
            bandeau_attente(slide)
    elif num == 13:
        compresser(slide, 2.6, 16.9, 3.9, 16.25, uniforme=True)
        formes = list(slide.shapes)
        titres = par_pos(lambda sh: sh.name in ("Rectangle 3", "Rectangle 23", "Rectangle 39", "Rectangle 44"))
        corps_ = par_pos(lambda sh: sh.name in ("Rectangle 30", "Rectangle 31", "Rectangle 32", "Rectangle 33"))
        for sh, (t_, _) in zip(titres, d["cartes"]):
            ecrire(sh, [t_], [14], [True])
        for sh, (_, c_) in zip(corps_, d["cartes"]):
            ecrire(sh, [c_], [12.5], [False])
            sh.text_frame.word_wrap = True
            for p in sh.text_frame.paragraphs:
                pPr = p._p.get_or_add_pPr()
                for bu in list(pPr):
                    if bu.tag.endswith(("buChar", "buAutoNum")):
                        pPr.remove(bu)
                pPr.insert(0, pPr.makeelement(qn("a:buNone"), {}))
                pPr.set("marL", "0")
                pPr.set("indent", "0")
    elif num == 29:
        compresser(slide, 3.5, 15.6, 4.2, 16.1)
        formes = list(slide.shapes)
        blocs = par_pos(lambda sh: sh.name in ("Rectangle 30", "Rectangle 31", "Rectangle 33", "Rectangle 41"))
        for sh, (t_, c_) in zip(blocs, d["blocs"]):
            ecrire(sh, [t_, c_], [14, 11.5], [True, False])
    elif num == 31:
        for sh in formes:
            if sh.name == "Title 2":
                ecrire(sh, [""])
        compresser(slide, 3.2, 16.1, 4.0, 16.2, uniforme=True)
        formes = list(slide.shapes)
        blocs = par_pos(lambda sh: sh.name in ("Rectangle 5", "Rectangle 7", "Rectangle 9", "Rectangle 12"))
        for sh, (t_, c_) in zip(blocs, d["cartes"]):
            ecrire(sh, [t_, c_], [14, 12], [True, False])
            sh.text_frame.word_wrap = True
    elif num == 107:
        charts = sorted([sh for sh in formes if getattr(sh, "has_chart", False) and sh.has_chart], key=lambda sh: sh.left)
        for sh, (v, lib, cle) in zip(charts, d["anneaux"]):
            cd = CategoryChartData()
            cd.categories = ["Valeur", "Reste"]
            cd.add_series("Série", (round(100 * v, 2), round(100 - 100 * v, 2)))
            sh.chart.replace_data(cd)
            sh.chart.has_title = False
            nh = Cm(7.4)
            delta = sh.top - Cm(4.4) + (sh.height - nh) // 2
            for o in formes:
                if o is not sh and sh.left <= o.left + o.width / 2 <= sh.left + sh.width and sh.top <= o.top + o.height / 2 <= sh.top + sh.height:
                    o.top = o.top - delta
            sh.top, sh.height = Cm(4.4), nh
            pts = sh.chart.plots[0].series[0].points
            for k_, coul in enumerate((ORANGE, BLEU)):
                pts[k_].format.fill.solid()
                pts[k_].format.fill.fore_color.rgb = RGBColor.from_string(coul)
            txt(slide, sh.left / 360000 - 0.3, sh.top / 360000 + sh.height / 360000 + 0.05, sh.width / 360000 + 0.6, 0.9, pc(cle, v), 20, True, BLEU, PP_ALIGN.CENTER)
            txt(slide, sh.left / 360000 - 0.3, sh.top / 360000 + sh.height / 360000 + 0.9, sh.width / 360000 + 0.6, 1.6, lib, 11, False, NOIR, PP_ALIGN.CENTER)
        if analyse:
            corps(slide, 1.4, 14.35, 31.2, BAS_CONTENU - 14.35, [analyse], 13)


def sommaire(slide, entrees):
    txt(slide, 2.7, 0.85, 20, 1.1, "SOMMAIRE", 24, True, BLEU)
    pas = 1.37 if len(entrees) <= 9 else 13.4 / len(entrees)
    taille = 16 if len(entrees) <= 9 else (13 if len(entrees) <= 14 else 12)
    centres = [3.4 + i * pas + pas / 2 for i in range(len(entrees))]
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
    if x["type"] == "separateur" and x["num"] != "R" and "." not in str(x["num"]):
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
        # Couverture retouchée par l'utilisateur sur la version 5 : titre, sous-titre orange, logo du FPPN, date
        for sh in sl.shapes:
            if sh.name == "Google Shape;138;p25":
                sh.left, sh.top, sh.width, sh.height = Emu(908861), Emu(2534894), Emu(5918659), Emu(1231106)
                for r in sh.text_frame.paragraphs[0].runs:
                    r.font.size, r.font.bold, r.font.name = Pt(40), True, "Helvetica"
                    r.font.color.rgb = RGBColor.from_string("FFFFFF")
            if sh.name == "Text Placeholder 4":
                sh.left, sh.width = Emu(10506325), Emu(1187835)
        st_ = sl.shapes.add_textbox(Emu(817421), Emu(5135034), Emu(6497320), Emu(461665))
        st_.name = "ZoneTexte 2"
        r = st_.text_frame.paragraphs[0].add_run()
        r.text = "Rapport actuariel et statistique"
        r.font.size, r.font.bold, r.font.name = Pt(24), True, "Helvetica"
        r.font.color.rgb = RGBColor.from_string(ORANGE)
        lg_ = sl.shapes.add_picture(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo_fppn.jpg"), Emu(938577), Emu(4017907), Emu(1844040), Emu(865220))
        lg_.crop_top, lg_.crop_bottom = 0.23449, 0.29631
        lg_.name = "Logo FPPN"
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
        remplir_gab(sl, x["gab"], x["remplir"], x.get("analyse"))
    elif x["type"] == "contenu":
        paras = paragraphes(x)
        if "tableau" in x:
            tableau_avec_analyse(sl, x)
        if "pas_a_pas" in x:
            pas_a_pas(sl, x["pas_a_pas"], paras)
        if x.get("triangle"):
            triangle(sl)
        if "schema" in x:
            schema(sl, x["schema"], paras)
        if "logigramme" in x:
            # Logigrammes en formes modifiables, puis l'analyse en dessous
            if x["logigramme"] == "circuit":
                LG.circuit_precomptes(sl, 3.9, 2.35)
                y_ = 13.55
            elif x["logigramme"] == "chaine":
                LG.chaine_traitement(sl, 3.9)
                y_ = 13.55
            elif x["logigramme"] == "placement":
                LG.decision_placement(sl, 3.9, 2.35)
                y_ = 13.55
            elif x["logigramme"] == "deces":
                LG.provision_deces(sl, 4.0)
                y_ = 10.6
            elif x["logigramme"] == "facture":
                LG.circuit_facture(sl, 4.0)
                y_ = 10.6
            elif x["logigramme"] in ("epargne", "gna", "maladie_deces"):
                getattr(LG, "circuit_" + x["logigramme"])(sl, 3.9)
                y_ = 14.4
            elif x["logigramme"] == "cgrae":
                LG.circuit_cgrae(sl, 3.9)
                y_ = 14.3
            else:
                LG.nouvelles_recrues(sl, 4.0)
                y_ = 10.6
            corps(sl, 1.4, y_, 31.2, BAS_CONTENU - y_, paras, 12)
        if "graphique" in x:
            lg_v, xt, lt = disposition_graphique(paras)
            corps(sl, xt, 4.0, lt, BAS_CONTENU - 4.0, paras)
            valable(sl, x, BAS_CONTENU - _hv(x, lg_v), lg_v)
        if "graphiques" in x:
            valable(sl, x, (13.4 if x.get("bandeau") else BAS_CONTENU) - _hv(x), 31.2)
        if x.get("bandeau"):
            corps(sl, 1.4, 13.75, 31.2, BAS_CONTENU - 13.75, [x["bandeau"]], 13)
        if "image" in x:
            from PIL import Image
            im = Image.open(x["image"]["fichier"])
            larg = 21.8
            haut = larg * im.height / im.width
            if haut > BAS_CONTENU - 4.75:
                haut = BAS_CONTENU - 4.75
                larg = haut * im.width / im.height
            sl.shapes.add_picture(x["image"]["fichier"], Cm(1.4), Cm(4.65), Cm(larg), Cm(haut))
            txt(sl, 1.4, 3.85, larg, 0.7, x["image"]["legende"], 12, False, ORANGE, PP_ALIGN.CENTER)
            corps(sl, 1.4 + larg + 0.6, 4.0, 32.6 - (1.4 + larg + 0.6), BAS_CONTENU - 4.0, paras, 13)
    elif x["type"] == "annexe":
        txt(sl, 1.52, 0.4, 31.24, 1.1, x["titre"], 20, True, NOIR)
        tableau(sl, x["tableau"], 2.0, hauteur_max=14.3)
        if x.get("source"):
            txt(sl, 1.02, 17.15, 31.75, 0.51, x["source"], 7.5, False, "808080")
    elif x["type"] in ("liste_figures", "liste_tableaux"):
        liste = FIG if x["type"] == "liste_figures" else TAB
        nom = "Figure" if x["type"] == "liste_figures" else "Tableau"
        txt(sl, 1.52, 0.4, 31.24, 1.1, "Liste des figures" if nom == "Figure" else "Liste des tableaux", 20, True, NOIR)
        moitie = (len(liste) + 1) // 2
        for col, part in enumerate((liste[:moitie], liste[moitie:])):
            if part:
                tableau(sl, {"titre": None, "entetes": [nom, "Titre", "Diapo."], "largeurs": [1.6, 12.4, 1.4],
                             "lignes": [[str(n_), t_, str(d_)] for n_, t_, d_ in part], "hauteur_ligne": 0.3}, 1.8, 1.4 + col * 16.0, 7, hauteur_max=15.6)
    if x.get("notes"):
        sl.notes_slide.notes_text_frame.text = x["notes"]
COURANTE = None


def textes_visibles(shapes):
    for sh in shapes:
        if sh.shape_type == 6:
            yield from textes_visibles(sh.shapes)
            continue
        if sh.has_text_frame:
            yield sh.text_frame.text
        if getattr(sh, "has_table", False) and sh.has_table:
            for row in sh.table.rows:
                for c in row.cells:
                    yield c.text
        if getattr(sh, "has_chart", False) and sh.has_chart:
            for t_ in sh.chart._chartSpace.iter(qn("c:v")):
                yield t_.text or ""


def sans_point_virgule(tf):
    """Charte Adenka : pas de point-virgule dans les phrases ; on coupe en deux phrases."""
    for p_ in tf.paragraphs:
        for r in p_.runs:
            while " ; " in r.text:
                avant, apres = r.text.split(" ; ", 1)
                r.text = avant + ". " + apres[:1].upper() + apres[1:]


for i, (sl, x) in enumerate(zip(prs.slides, SL), 1):
    for sh in sl.shapes:
        if sh.has_text_frame:
            sans_point_virgule(sh.text_frame)
        if getattr(sh, "has_table", False) and sh.has_table:
            for row in sh.table.rows:
                for c in row.cells:
                    sans_point_virgule(c.text_frame)
    # v11 : numéros d'items de la PBC du 23/10/2026
    def _items(m):
        nums = []
        for a_, b_ in re.findall(r"(\d+)(?: à (\d+))?", m.group(2)):
            for k_ in range(int(a_), int(b_ or a_) + 1):
                if PBCM.get(k_) and PBCM[k_] not in nums:
                    nums.append(PBCM[k_])
        nums.sort()
        if not nums:
            return m.group(0)
        txt_ = ", ".join(str(n_) for n_ in nums[:-1]) + (" et " if len(nums) > 1 else "") + str(nums[-1])
        pre_ = m.group(1)
        if "item" in pre_:
            pre_ = "PBC, items " if len(nums) > 1 else "PBC, item "
        return pre_ + txt_
    surl_ = False
    for sh in sl.shapes:
        tfs = [sh.text_frame] if sh.has_text_frame else []
        if getattr(sh, "has_table", False) and sh.has_table:
            tfs += [c.text_frame for row in sh.table.rows for c in row.cells]
        for tf_ in tfs:
            for p_ in tf_.paragraphs:
                for r in p_.runs:
                    r.text = re.sub(r"(PBC,? items? |PBC )((?:\d+(?: à \d+)?(?:, | et )?)+)", _items, r.text)
                    if r._r.find(qn("a:rPr")) is not None and r._r.find(qn("a:rPr")).find(qn("a:highlight")) is not None:
                        surl_ = True
    # Sigles de la diapositive, définis en bas de page (au-dessus de la source)
    if x["type"] in ("contenu", "gab") or (x["type"] == "annexe" and x["titre"] != "Signes et abréviations"):
        texte = "\n".join(t_ for t_ in textes_visibles(sl.shapes) if t_)
        sig_ = sigles_du_texte(texte)
        if sig_ or surl_:
            ligne_s = ("Surligné en jaune : affirmation à confirmer par une pièce ou une réponse écrite du FPPN. " if surl_ else "") + ("Sigles : " + " | ".join(sig_) if sig_ else "")
            taille_s = 7 if len(ligne_s) < 330 else 6.5
            txt(sl, 1.02, 16.48, 31.75, 0.66, ligne_s, taille_s, False, "595959", PP_ALIGN.LEFT, MSO_ANCHOR.BOTTOM)
_corr = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "corrections_mise_en_page_v5.json"), encoding="utf-8"))
_appl = 0
for sl in prs.slides:
    t_ = [sh.text_frame.text for sh in sl.shapes if sh.has_text_frame and sh.name == "Adenka_Titre"]
    for c_ in _corr:
        if t_ and t_[0] == c_["titre"]:
            for f_ in c_["formes"]:
                for sh in sl.shapes:
                    if sh.name == f_["nom"] and all(abs(a_ - b_) < 2000 for a_, b_ in zip((sh.left, sh.top, sh.width, sh.height), f_["avant"])):
                        sh.left, sh.top, sh.width, sh.height = (Emu(v_) for v_ in f_["apres"])
                        _appl += 1
                        break
print("retouches de mise en page réappliquées :", _appl)
# v12 : bandeau « Comment lire » de la carte des risques, posé après les retouches de l'utilisateur
for sl in prs.slides:
    if id(sl) in LECT83 or any(sh.has_text_frame and sh.name == "Adenka_Titre" and sh.text_frame.text == "Carte des risques" for sh in sl.shapes):
        lect_ = LECT83.get(id(sl)) or next(iter(LECT83.values()), None)
        if lect_:
            compresser(sl, 3.5, 16.3, 6.1, 16.3)
            encadre(sl, 1.4, 3.85, 31.2, 2.0, [lect_], "Comment lire")
prs.save(SORTIE)
finaliser(SORTIE)
os.remove(tmp)

# ---------------------------------------------------------------- contrôle automatique des chiffres
prs2 = Presentation(SORTIE)
textes = []
for sl in prs2.slides:
    morceaux = list(textes_visibles(sl.shapes))
    morceaux.append(sl.notes_slide.notes_text_frame.text if sl.has_notes_slide else "")
    textes.append("\n".join(morceaux))
for t_ in TRACE:
    t_["retrouve"] = t_["texte"] in textes[t_["diapositive"] - 1] if t_["diapositive"] <= len(textes) else False
for _l in FIG + TAB:
    _l.append(SL[_l[2] - 1].get("valable") or "")
json.dump({"figures": FIG, "tableaux": TAB, "diapositives": len(SL), "trace": TRACE, "validite": VAL_LIGNES},
          open(SORTIE.replace(".pptx", "_index.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
ok = sum(t_["retrouve"] for t_ in TRACE)
for x in _REG:
    dia = sorted({t_["diapositive"] for t_ in TRACE if x["traces"] and re.search(x["traces"], t_["indicateur"])})
    x["diapositives"] = ", ".join(str(d_) for d_ in dia)
json.dump(_REG, open(SORTIE.replace(".pptx", "_formules.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(f"{len(SL)} diapositives ; {len(FIG)} figures ; {len(TAB)} tableaux ; chiffres retrouvés {ok}/{len(TRACE)}")
for t_ in TRACE:
    if not t_["retrouve"]:
        print("  NON RETROUVÉ", t_["diapositive"], t_["indicateur"], t_["texte"])
