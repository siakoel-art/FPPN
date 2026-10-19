# -*- coding: utf-8 -*-
"""
FPPN - Comptes rendus des entretiens du 19/10/2026 avec M. Koffi, chef comptable (PowerPoint, charte Adenka)
==========================================================================================================

Deux comptes rendus : placements (14h) et provisions (16h). Les déclarations sont présentées comme telles
(« à confirmer par document ») et confrontées aux balances des comptes reçues le 15/10/2026. Les propos sensibles
figurent dans les notes internes confidentielles. Les points de contrôle interne restent dans les comptes rendus.
Utilisation : python cr_placements_provisions_fppn.py <dossier_base> <skill_adenka_powerpoint> <dossier_sortie>
"""
import os
import sys

import pandas as pd

BASE, SKILL, SORTIE = sys.argv[1:4]
sys.path.insert(0, os.path.join(SKILL, "scripts"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx_tools import selection, remplir, finaliser  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.dml.color import RGBColor  # noqa: E402
from pptx.enum.shapes import MSO_SHAPE  # noqa: E402
from pptx.util import Cm  # noqa: E402
import logigramme_fppn as LG  # noqa: E402
from outils_slides_fppn import BAS_CONTENU, GRIS, GRIS_F, ORANGE, corps, ligne_sigles, tableau, txt  # noqa: E402

TIMESAVER = os.path.join(SKILL, "assets", "Timesaver_Adenka.pptx")
SRC_P = "Source : entretien du 19/10/2026 (14h) avec M. Koffi, chef comptable ; déclarations à confirmer par document"
SRC_V = "Source : entretien du 19/10/2026 (16h) avec M. Koffi, chef comptable ; déclarations à confirmer par document"
BAL = "balances des comptes du FPPN (Sage 100), 2020 au 31/08/2026, reçues le 15/10/2026 ; traitements Adenka"


def nb(x, dec=0):
    return f"{x:,.{dec}f}".replace(",", " ").replace(".", ",")


def F(x):
    return nb(round(x)) + " FCFA"


def pc(x):
    return nb(100 * x, 2) + " %"


# ---------------------------------------------------------------------------
# Confrontation aux balances
# ---------------------------------------------------------------------------
bal = pd.read_parquet(os.path.join(BASE, "compta_balances.parquet"))
bal = bal[bal.nature == "comptes"].sort_values("periode_au").groupby(["exercice", "compte"]).tail(1)


def sol(ex, c):
    g = bal[(bal.exercice == ex) & bal.compte.str.match(c)]
    return float((g.sol_d - g.sol_c).sum())


def mvd(ex, c):
    g = bal[(bal.exercice == ex) & bal.compte.str.match(c)]
    return float(g.mvt_d.sum())


def mvc(ex, c):
    g = bal[(bal.exercice == ex) & bal.compte.str.match(c)]
    return float(g.mvt_c.sum())


DAT_EFC = {ex: mvd(ex, r"^5221020") for ex in range(2021, 2027)}
DAT_EFC_SOLDE = sol(2026, r"^52210200$")
INT_2026 = -sol(2026, r"^77110000$")
DAT_PEP = sol(2024, r"^52210300$")
ORA_OUV, ORA_REMB = mvd(2022, r"^52210400$"), mvc(2023, r"^52210400$")
FCP = sol(2025, r"^53300001$")
TERR = {ex: sol(ex, r"^22200000$") for ex in (2020, 2024, 2025)}
REEV = -sol(2025, r"^10620000$")
HOP = sol(2026, r"^47111100$")
CMP_I, CMP_C = sol(2026, r"^47111104$"), sol(2026, r"^47111105$")
POLY21 = sol(2021, r"^47111101$")
GRE = sol(2026, r"^47111102$")
EMP21 = -sol(2021, r"^16200000$")
LOY = {ex: -sol(ex, r"^71810000$") for ex in range(2022, 2027)}
COM = {ex: -sol(ex, r"^71800000$") for ex in range(2022, 2027)}
CARTES = -sol(2025, r"^71800001$")
PSAP_M, PSAP_D = -sol(2025, r"^15801"), -sol(2025, r"^15802")
PRIME_2020 = sol(2020, r"^65180000$")
PIDR = {ex: -sol(ex, r"^196") for ex in (2020, 2025)}
pre = pd.read_parquet(os.path.join(BASE, "precomptes.parquet"))
GNA_PREC = pre[pre.retenu & (pre.sens == "retour") & (pre.poste == "9W3")].groupby("annee").montant.sum().to_dict()


# ---------------------------------------------------------------------------
# Rendus communs
# ---------------------------------------------------------------------------
def tuiles(s, tiles):
    tw = (31.2 - 3 * 0.4) / 4
    for i, (a, b) in enumerate(tiles):
        x = 1.4 + i * (tw + 0.4)
        r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(x), Cm(3.85), Cm(tw), Cm(1.8))
        r.fill.solid()
        r.fill.fore_color.rgb = RGBColor.from_string(GRIS)
        r.line.fill.background()
        r.shadow.inherit = False
        txt(s, x + 0.3, 3.95, tw - 0.6, 0.8, a, 15, True, ORANGE)
        txt(s, x + 0.3, 4.85, tw - 0.6, 0.6, b, 9, False, GRIS_F)


def infos(tiles, participants, paras):
    def r(s):
        tuiles(s, tiles)
        bas = tableau(s, {"titre": None, "entetes": ["Structure", "Nom", "Fonction", "Rôle dans l'entretien"], "largeurs": [4.6, 7.0, 10.6, 9.0],
                          "lignes": participants}, 6.2, taille=11)
        corps(s, 1.4, bas + 0.5, 31.2, BAS_CONTENU - bas - 0.5, paras, 12)
    return r


def table_texte(spec, paras, taille_txt=12):
    def r(s):
        bas = tableau(s, spec, 3.85, hauteur_max=BAS_CONTENU - 3.85 - 3.2)
        corps(s, 1.4, bas + 0.35, 31.2, BAS_CONTENU - bas - 0.35, paras, taille_txt)
    return r


def logi(fonction, y_texte, paras):
    def r(s):
        fonction(s)
        corps(s, 1.4, y_texte, 31.2, BAS_CONTENU - y_texte, paras, 11)
    return r


def construire(sl, couverture, date, sortie):
    tmp = sortie.replace(".pptx", "_tmp.pptx")
    selection(TIMESAVER, tmp, [1] + [5] * len(sl) + [110])
    cj = {"1": {"Google Shape;138;p25": couverture, "Text Placeholder 4": date}}
    for i, x in enumerate(sl, 2):
        cj[str(i)] = {"@vider": True, "@entete": x["entete"], "@source": x["source"]}
    cj[str(len(sl) + 2)] = {}
    remplir(tmp, cj, tmp)
    prs = Presentation(tmp)
    slides = list(prs.slides)
    for sh in slides[0].shapes:
        if sh.name == "Google Shape;138;p25":
            sh.width = Cm(12.8)
    for i, x in enumerate(sl, 1):
        s = slides[i]
        x["rendu"](s)
        # Charte Adenka : pas de point-virgule dans les phrases (sources exceptées)
        cadres = [sh.text_frame for sh in s.shapes if sh.has_text_frame and not sh.text_frame.text.startswith(("Source", "Sigles"))]
        cadres += [c.text_frame for sh in s.shapes if getattr(sh, "has_table", False) and sh.has_table for r_ in sh.table.rows for c in r_.cells]
        for tf in cadres:
            for p_ in tf.paragraphs:
                for r_ in p_.runs:
                    while " ; " in r_.text:
                        av, ap = r_.text.split(" ; ", 1)
                        r_.text = av + ". " + ap[:1].upper() + ap[1:]
        if x["notes"]:
            s.notes_slide.notes_text_frame.text = x["notes"]
        textes = []
        for sh in s.shapes:
            if sh.has_text_frame:
                textes.append(sh.text_frame.text)
            if getattr(sh, "has_table", False) and sh.has_table:
                textes += [c.text for r_ in sh.table.rows for c in r_.cells]
        ligne_sigles(s, textes)
    prs.save(sortie)
    finaliser(sortie)
    os.remove(tmp)
    print(os.path.basename(sortie), len(sl) + 2, "diapositives")


def d(sl, titre, sous_titre, synthese, rendu, source, notes=""):
    sl.append({"entete": {"titre": titre, "sous_titre": sous_titre, "synthese": synthese}, "rendu": rendu, "notes": notes, "source": source})


# ---------------------------------------------------------------------------
# Compte rendu « placements »
# ---------------------------------------------------------------------------
P = []
d(P, "Informations de l'entretien", "Entretien avec le chef comptable sur les placements", "Un entretien consacré aux placements, aux autres actifs, aux ressources et aux contrôles comptables",
  infos([("19 octobre 2026", "Entretien de 14h"), ("Placements", "Décision, suivi, actifs"), ("M. Koffi", "Chef comptable"), ("Documents", "Demande complémentaire, items 80 et suivants")],
        [["FPPN", "M. Koffi", "Chef comptable", "Interlocuteur"],
         ["Adenka", "Emmanuel Dacoury", "Équipe Adenka", "Participant"],
         ["Adenka", "Marlène Mody", "Équipe Adenka", "Participante"],
         ["Adenka", "Cédric Kouamé", "Équipe Adenka", "Participant"],
         ["Adenka", "Jean-Noël Sialou", "Consultant actuaire", "Placements, engagements et provisions"]],
        ["**Objet de l'entretien.** Comprendre comment les placements sont décidés, suivis et comptabilisés, et quels autres actifs et ressources porte le FPPN.",
         "**Méthode.** Chaque déclaration est présentée comme telle et confrontée aux balances des comptes reçues le 15/10/2026. Les documents manquants sont repris dans la demande complémentaire."]),
  "Source : entretien du 19/10/2026 (14h)")
d(P, "Décision de placement", "Le directeur général décide seul, sur la base d'un point de trésorerie",
  "Aucun comité, aucun avis du conseil d'administration, aucune procédure écrite selon le chef comptable",
  logi(lambda s: LG.decision_placement(s, 3.9, 2.35), 13.55,
       ["**Le circuit de décision tient en une personne.** La banque propose un placement au directeur général, qui l'analyse et demande au DAF un point de trésorerie, préparé par la comptabilité. Le directeur général décide ; une lettre d'ouverture est adressée à la banque, qui prélève les versements. La comptabilité enregistre les opérations sur la base des relevés bancaires.",
        "Sans politique de placement écrite, ni comité, le choix des placements, de leur durée et de leur liquidité ne peut pas être rapproché des échéances de l'épargne et des prestations."]),
  SRC_P, "Logigramme modifiable : chaque forme et chaque flèche peuvent être déplacées dans PowerPoint. Les formes à contour rouge signalent un contrôle manquant.")
d(P, "Placements identifiés", "Les balances confirment le DAT annuel, le FCP et l'ancien DAT Orabank",
  f"DAT BOA : {F(DAT_EFC_SOLDE)} au 31/08/2026 ; FCP United Capital : {F(FCP)} depuis 2025",
  table_texte({"titre": "Placements déclarés et confrontés aux balances", "entetes": ["Placement", "Déclaration du chef comptable", "Ce que montrent les balances", "À confirmer"],
               "largeurs": [4.4, 9.4, 10.6, 6.8], "taille": 9,
               "lignes": [["DAT BOA (compte 52210200)", "Placement du FPPN, alimenté chaque mois (60 millions par mois en 2026) jusqu'à fin décembre, qui sert aussi à payer l'EFC en fin d'année.",
                           f"Versements de {F(DAT_EFC[2022])} en 2022, {F(DAT_EFC[2025])} en 2025 et {F(DAT_EFC[2026])} au 31/08/2026. Solde de {F(DAT_EFC_SOLDE)}, dont {F(INT_2026)} d'intérêts en 2026. Le compte s'intitule « DAT BOA Entraide fin de carrière ».", "Convention, taux, relevés."],
                          ["FCP (compte 53300001)", "Fonds commun de placement de 500 millions, géré par United Capital (SGI), rendement annoncé de 8 à 10 %, non garanti.",
                           f"{F(FCP)} en 2025 et au 31/08/2026, sans revalorisation enregistrée.", "Souscription, valeur liquidative, relevés."],
                          ["DAT BOA PEP (52210300)", "Non cité.", f"{F(DAT_PEP)} de 2020 à 2024, soldé en 2025.", "Convention et motif du débouclage."],
                          ["DAT Orabank (52210400)", "Versements mensuels arrêtés par la banque après deux prélèvements ; capital et intérêts reversés au terme, DAT non reconduit.",
                           f"Ouvert en 2022 pour {F(ORA_OUV)}, soldé en 2023 pour {F(ORA_REMB)}. Le chef comptable situait l'incident en 2021.", "Année et correspondance avec la banque."],
                          ["Ancienneté", "Placements depuis 2005 environ.", "Les balances reçues commencent en 2020 ; des DAT existent déjà à l'ouverture de 2020.", "Historique des placements 2016 à 2026."]]},
              [f"**Le DAT annuel est le principal placement du FPPN : il est soldé à chaque clôture de 2021 à 2025.** Les versements vont de {F(min(DAT_EFC[e] for e in range(2021, 2026)))} à {F(max(DAT_EFC[e] for e in range(2021, 2026)))} par an.",
               "Le FCP n'a pas été revalorisé dans les comptes : sa valeur au 31/08/2026 et son rendement réel restent à documenter."], 11),
  SRC_P + " ; " + BAL)
d(P, "Immobilier et autres actifs", "Terrains, hôpital et clinique pèsent plus lourd que les placements financiers",
  "Plusieurs actifs ne produisent pas de revenu ou ne sont plus sous le contrôle du FPPN",
  table_texte({"titre": "Autres actifs déclarés et confrontés aux balances", "entetes": ["Actif", "Déclaration du chef comptable", "Ce que montrent les balances"], "largeurs": [5.0, 13.0, 13.2], "taille": 9,
               "lignes": [["Terrains", "Deux grands terrains, dont un avec un projet ; quelques terrains achetés, sans programme immobilier engagé.", f"Terrains nus (22200000) : {F(TERR[2024])} en 2024, {F(TERR[2025])} en 2025."],
                          ["Immeubles d'Aboisso", "Deux villas et un autre bâtiment, loués à usage d'habitation.", f"Revenus des immeubles de rapport (71810000) : {F(LOY[2025])} en 2025."],
                          ["Siège", "Le terrain du siège a été réévalué par un expert immobilier.", f"Compte 10620000 : {F(REEV)}, intitulé « Réévaluation matériels » de 2020 à 2025, puis « Réserve disponible » en 2026 ; montant inchangé depuis 2020."],
                          ["Hôpital de la Police", "Équipé et financé par le FPPN, puis repris par la direction des services de santé de la Police. Démarches en cours auprès de l'État.", f"Créance 47111100 : {F(HOP)}, inchangée depuis 2020."],
                          ["Clinique du FPPN (CMP)", "Dépenses suivies en comptes 47 en attendant une comptabilité propre ; un cabinet prépare le bilan d'ouverture.", f"Compte « Polyclinique du FPPN » ({F(POLY21)} en 2021) reclassé en 2022. Au 31/08/2026 : investissements {F(CMP_I)}, charges de fonctionnement {F(CMP_C)}."],
                          ["Grenier du Policier", "Sans personnalité morale, créé en novembre 2021. Le FPPN a emprunté à la banque pour le financer et a remboursé l'emprunt ; le Grenier n'a encore rien remboursé.", f"Emprunt (16200000) : {F(EMP21)} fin 2021, soldé en 2023. Financement du Grenier (47111102) : {F(GRE)} au 31/08/2026."]]},
              ["**Les déclarations sont cohérentes avec les balances, à une réserve près : l'intitulé du compte 10620000 parle de matériels de 2020 à 2025, et non du terrain du siège.** Le rapport de réévaluation est demandé.",
               "Les créances sur l'hôpital et les avances à la clinique et au Grenier sont des actifs peu liquides, dont le recouvrement n'est pas assuré : elles comptent pourtant dans les fonds propres."], 11),
  SRC_P + " ; " + BAL)
d(P, "Ressources et produits annexes", "Hors cotisations, le FPPN tire des revenus faibles de ses activités annexes",
  "Commissions de GNA, loyers, cartes d'accès aux soins et intérêts des placements",
  table_texte({"titre": "Ressources déclarées et confrontées aux balances", "entetes": ["Ressource", "Déclaration du chef comptable", "Ce que montrent les balances"], "largeurs": [5.0, 12.4, 13.8], "taille": 9,
               "lignes": [["Cotisations obligatoires", "5 % du salaire de base, 3,5 % pour la maladie et 1,5 % pour le décès, précomptés chaque mois.", "La Solde retient 3,77 % et 1,23 % (rapport, partie 4)."],
                          ["Épargne PEP et EFC", "Produits facultatifs, enregistrés en dettes envers les mutualistes, restitués en fin de contrat ou à la retraite.", "Comptes 41195100 et 41195200 (rapport, partie 6)."],
                          ["Flotte automobile (GNA)", "Le FPPN est apporteur d'affaires : il perçoit 20 % des primes collectées, en fin d'année.",
                           f"Produits des activités annexes (71800000) : {F(COM[2024])} en 2024, égal au débit du compte GNA. Précomptes 9W3 de 2024 : {F(GNA_PREC.get(2024, 0))}. Le rapport des deux n'est pas de 20 %."],
                          ["Cartes d'accès aux soins", "Les cartes sont payantes.", f"{F(CARTES)} en 2025 (71800001)."],
                          ["Subventions", "Aucune subvention.", "Aucun compte de subvention mouvementé."]]},
              ["**Les commissions de GNA ne correspondent pas à 20 % des précomptes du poste 9W3.** L'assiette réelle (primes collectées par précompte et hors précompte) et la convention avec GNA doivent l'expliquer.",
               "Ces ressources annexes restent marginales : l'équilibre du FPPN dépend des cotisations et du rendement de ses placements."], 11),
  SRC_P + " ; " + BAL + " ; fichiers de précomptes reçus le 15/10/2026")
d(P, "Cotisations : réception et contrôle", "Les cotisations sont saisies en montant global, sans rapprochement formalisé",
  "Le FPPN ne rapproche pas la comptabilité de la plateforme des adhérents (Easy Money)",
  table_texte({"titre": "Circuit comptable des cotisations", "entetes": ["Étape", "Déclaration du chef comptable"], "largeurs": [6.0, 25.2], "taille": 10,
               "lignes": [["Réception", "La Solde transmet l'état de retour sur papier, vers le 28 du mois."],
                          ["Encaissement", "Le Trésor vire les fonds entre le 5 et le 10 du mois suivant, parfois en partie quand il a d'autres priorités. Le reste est versé plus tard, sans avertissement."],
                          ["Saisie", "Le comptable saisit le montant global à la main, à partir de l'état de retour."],
                          ["Contrôle", "Le montant viré est comparé au relevé bancaire, sans formalisation. Un écart est signalé au DAF, qui relance le Trésor."],
                          ["Rapprochement", "Aucun rapprochement entre la comptabilité et Easy Money, ni avec le fichier envoyé à la Solde."]]},
              ["**Le contrôle des cotisations se limite à vérifier que l'argent annoncé est arrivé.** Personne ne vérifie que chaque mutualiste a été précompté du bon montant.",
               "Les retards de virement du Trésor créent une créance (compte 41111000) dont le suivi relève de la comptabilité."], 11),
  SRC_P)
d(P, "Contrôle interne comptable et bancaire", "Le logiciel comptable et la banque présentent des faiblesses de contrôle",
  "Points relevés pendant l'entretien, sans effet direct sur les calculs actuariels",
  table_texte({"titre": "Points de contrôle interne relevés", "entetes": ["Domaine", "Déclaration du chef comptable", "Risque"], "largeurs": [5.0, 15.0, 11.2], "taille": 10,
               "lignes": [["Saisie", "Toutes les écritures sont saisies à la main ; aucune validation par une seconde personne.", "Erreurs de saisie non détectées."],
                          ["Modification", "Une écriture validée peut être modifiée ou supprimée.", "Piste d'audit non garantie."],
                          ["Accès", "Tous les utilisateurs ont les mêmes droits dans le logiciel.", "Absence de séparation des tâches."],
                          ["Archives", "Une nouvelle base de données est créée tous les cinq à six ans.", "Historique dispersé entre plusieurs bases."],
                          ["Rapprochements bancaires", "Réalisés, mais plus supervisés.", "Écarts non revus."],
                          ["Chèques en suspens", "Annulés après trois ans, lors du changement de signataires.", "Dettes éteintes sans analyse."],
                          ["Comptes bancaires", "Des comptes inactifs ne sont pas fermés.", "Fonds ou mouvements non suivis."]]},
              ["**Ces faiblesses réduisent la fiabilité des balances utilisées par Adenka.** Elles justifient les contrôles de cohérence appliqués aux données (rapport, partie 3) et une revue par sondage des écritures."], 11),
  SRC_P)
d(P, "Documents et points d'attention", "Les placements et les autres actifs doivent être documentés",
  "Les documents demandés sont repris dans la demande complémentaire (items 80 et suivants)",
  table_texte({"titre": "Documents demandés", "entetes": ["Sujet", "Document", "Suite"], "largeurs": [6.0, 18.2, 7.0], "taille": 10,
               "lignes": [["Placements", "Conventions et relevés du DAT BOA et du FCP United Capital ; inventaire et tableau de variation des placements de 2016 à 2026", "PBC, items 80 à 82"],
                          ["Orabank", "Correspondance sur l'arrêt des versements et le remboursement", "PBC, item 83"],
                          ["Immobilier", "Titres et actes des terrains, baux des immeubles d'Aboisso, rapport de réévaluation du siège", "PBC, items 84 à 86"],
                          ["Hôpital et clinique", "Dossier de la créance sur l'hôpital, bilan d'ouverture de la clinique", "PBC, items 87 et 88"],
                          ["GNA", "Convention d'apporteur d'affaires et décomptes de commissions", "PBC, item 89"],
                          ["Banque", "Liste des comptes bancaires et des signataires, état des chèques en suspens", "PBC, items 90 et 91"]]},
              ["**L'entretien éclaire le fonctionnement, mais ne remplace pas les documents.** Les déclarations seront intégrées au rapport avec la mention « à confirmer par document »."], 12),
  SRC_P)

# ---------------------------------------------------------------------------
# Compte rendu « provisions »
# ---------------------------------------------------------------------------
V = []
d(V, "Informations de l'entretien", "Entretien avec le chef comptable sur les provisions", "Un entretien consacré aux provisions techniques, à la provision pour départ à la retraite et à l'épargne",
  infos([("19 octobre 2026", "Entretien de 16h"), ("Provisions", "Maladie, décès, retraite"), ("M. Koffi", "Chef comptable"), ("Documents", "Demande complémentaire, items 92 et suivants")],
        [["FPPN", "M. Koffi", "Chef comptable", "Interlocuteur"],
         ["Adenka", "Issa Bamba", "Équipe Adenka", "Conduite de l'entretien"],
         ["Adenka", "Jean-Noël Sialou", "Consultant actuaire", "Provisions et engagements"]],
        ["**Objet de l'entretien.** Comprendre comment les provisions maladie, décès et retraite sont calculées, validées et comptabilisées.",
         "**Méthode.** Chaque déclaration est présentée comme telle et confrontée aux balances des comptes reçues le 15/10/2026. Les documents manquants sont repris dans la demande complémentaire."]),
  "Source : entretien du 19/10/2026 (16h)")
d(V, "Provisions maladie et montants hérités", "Le FPPN ne constitue aucune provision maladie ; les montants au bilan datent de 2004",
  f"Maladie {F(PSAP_M)} et obsèques {F(PSAP_D)} : montants hérités, jamais recalculés selon le chef comptable",
  table_texte({"titre": "Provisions : déclarations et balances", "entetes": ["Sujet", "Déclaration du chef comptable", "Ce que montrent les balances"], "largeurs": [5.2, 14.0, 12.0], "taille": 10,
               "lignes": [["Maladie", "Aucune provision : les factures de l'année reçues jusqu'en mars de l'année suivante sont rattachées à l'exercice.", "Dettes envers les cliniques (40110000) à chaque clôture (rapport, partie 9)."],
                          ["Montants hérités", "Provisions passées vers 2004, avant le PCMS, selon une méthode de l'époque ; plusieurs tentatives d'annulation ; archives perdues lors d'un sinistre en 2010.", f"Mêmes montants de 2020 au 31/08/2026 : {F(PSAP_M)} et {F(PSAP_D)}."],
                          ["Reclassement", "Reclassement du compte 19 vers le compte 158 à la demande de l'AIRMS, date non précisée.", "Reclassement visible dans la balance 2024."],
                          ["Gestion du décès", "Allocation décès confiée à LMAI, puis à Sanlam, avant le retour en gestion interne, au même barème.", f"Prime d'assurance (65180000) : {F(PRIME_2020)} en 2020, aucune ensuite."]]},
              ["**Les provisions au bilan ne sont pas des estimations : ce sont des montants hérités que personne ne sait plus justifier.** Les factures de maladie reçues jusqu'en mars couvrent une partie des soins survenus ; les soins facturés plus tard ne sont couverts par rien.",
               "Le recalcul des provisions selon le règlement d'exécution n° 004/2022 (art. 10) suppose la base des prestations avec les dates de soins et de déclaration."], 11),
  SRC_V + " ; " + BAL)
d(V, "Allocations décès en fin d'année", "Le calcul part de la liste des décès transmise par la direction du personnel",
  "Les dossiers non payés sont valorisés au barème par grade ; le compte où ce montant est passé reste à préciser",
  logi(lambda s: LG.provision_deces(s, 4.0), 10.6,
       ["**Chaque mois, l'informaticien reçoit la liste des policiers décédés, pour arrêter leurs précomptes.** En fin d'année, le chef comptable la rapproche des allocations payées ; les dossiers restants sont valorisés au barème : sous-officier 1 000 000 FCFA, commissaire 2 000 000 FCFA, officier à confirmer.",
        "**Les décès de conjoints et d'enfants échappent à cette liste tant qu'ils ne sont pas déclarés.** Ils forment une part des sinistres tardifs, à estimer avec la méthode retenue pour la provision.",
        "Le seul compte de provision décès des balances (15802000) ne bouge pas : il faut savoir où ce calcul annuel est comptabilisé."]),
  SRC_V, "Logigramme modifiable. Le montant de l'allocation des officiers n'est pas lisible dans la transcription : il est à confirmer.")
d(V, "Départ à la retraite du personnel", "La provision suit la convention collective, à l'ancienneté",
  f"Provision au compte 196 : {F(PIDR[2020])} en 2020, {F(PIDR[2025])} en 2025",
  table_texte({"titre": "Éléments de calcul de la provision pour départ à la retraite", "entetes": ["Élément", "Valeur ou déclaration", "Source"], "largeurs": [7.0, 15.2, 9.0], "taille": 10,
               "lignes": [["Barème de l'indemnité", "Pourcentage du salaire mensuel par année d'ancienneté, selon la convention collective", "Chef comptable"],
                          ["Charges fiscales patronales", "2,80 % : IS 1,20 %, FDFP 0,40 % et 1,20 %", "Note de taux Adenka (source à préciser)"],
                          ["Retraite", "7,70 %, plafonné à 3 375 000 FCFA", "Note de taux Adenka (source à préciser)"],
                          ["Prestations familiales", "5,75 %, plafonné à 75 000 FCFA", "Note de taux Adenka (source à préciser)"],
                          ["Risques professionnels", "2 % au FPPN (3 % dans la note de taux), plafonné à 75 000 FCFA", "Chef comptable"],
                          ["Documents", "Notification du taux de risques professionnels, déclarations CNPS, fichier de calcul 2020 à 2026", "Demandés"]]},
              ["**Selon la description, la méthode du FPPN applique la convention collective à l'ancienneté acquise.** Elle ne mentionne ni la probabilité d'être présent au départ, ni l'actualisation, que prévoit la méthode actuarielle des unités de crédit projetées.",
               "Le recalcul attend le fichier du personnel et le fichier de calcul du FPPN."], 11),
  SRC_V + " ; " + BAL)
d(V, "Autres points abordés", "Épargne, réserve de réévaluation, Grenier et cotisations",
  "Points utiles au rapport, à confirmer par document",
  table_texte({"titre": "Autres déclarations du chef comptable", "entetes": ["Sujet", "Déclaration", "Suite"], "largeurs": [5.4, 18.6, 7.2], "taille": 10,
               "lignes": [["Taux servi sur l'épargne", "3,5 % par an sur le PEP et l'EFC.", "Décision fixant le taux"],
                          ["Compte 10620000", "Réévaluation du terrain du siège par un expert.", "Rapport de réévaluation"],
                          ["Clé de répartition", "3,5 % maladie et 1,5 % décès, sans document qui la fixe.", "Barème officiel"],
                          ["Grenier", "Comptabilité propre dans un logiciel de microfinance ; lien avec le FPPN par un compte 47 où sont retracées les charges payées pour lui depuis novembre 2021.", "Comptes du Grenier"],
                          ["Arriérés du Trésor", "Virements parfois partiels ; le reste peut tarder.", "État des arriérés"]]},
              ["**Le taux de 3,5 % servi sur l'épargne est un engagement de rendement.** Il doit être couvert par les produits des placements, ce que le rapport mesurera."], 12),
  SRC_V)
d(V, "Documents et points d'attention", "Les provisions héritées ne peuvent être justifiées que par des documents",
  "Les documents demandés sont repris dans la demande complémentaire (items 92 et suivants)",
  table_texte({"titre": "Documents demandés", "entetes": ["Sujet", "Document", "Suite"], "largeurs": [6.0, 18.2, 7.0], "taille": 10,
               "lignes": [["Décès", "Fichier des décès de l'année et calcul des allocations restant à payer ; barème des allocations par grade", "PBC, items 92 et 93"],
                          ["Gestion déléguée", "Contrats avec LMAI et Sanlam, et décomptes", "PBC, item 94"],
                          ["Reclassement", "Courrier de l'AIRMS demandant le reclassement en compte 158", "PBC, item 95"],
                          ["Retraite du personnel", "Notification du taux de risques professionnels, déclarations CNPS, fichier de calcul 2020 à 2026", "PBC, items 96 à 98"],
                          ["Épargne", "Décision fixant le taux servi de 3,5 %", "PBC, item 99"]]},
              ["**Le compte où est passé le calcul annuel des allocations décès est la question prioritaire.** Sans lui, le rapport ne peut pas dire si le FPPN constate ses allocations restant à payer."], 12),
  SRC_V)

construire(P, "Entretien\nPlacements", "19 octobre 2026", os.path.join(SORTIE, "FPPN_CR_entretien_placements_20261019.pptx"))
construire(V, "Entretien\nProvisions", "19 octobre 2026", os.path.join(SORTIE, "FPPN_CR_entretien_provisions_20261019.pptx"))
