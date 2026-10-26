"""Résultats, hypothèses et conclusions des quatre axes actuariels (engagements, provisions, adossement, placements),
avec les méthodes expliquées en annexe. Tous les chiffres sont repris du rapport actuariel et statistique du 23/10/2026.

Le moteur de mise en page (charte Adenka, Timesaver, tableaux, sigles en bas de diapositive) est repris du script du
rapport : rapport_v12_fppn.py (en-tête, aides et construction), exécutés avec le contenu ci-dessous.

Usage : python3 deck_resultats_axes_fppn.py <skill adenka-powerpoint> <sortie.pptx>
"""
import os
import sys

ICI = os.path.dirname(os.path.abspath(__file__))
SKILL, SORTIE = sys.argv[1:3]
SRC = open(os.path.join(ICI, "rapport_v12_fppn.py"), encoding="utf-8").read().split("\n")


def lignes(a, b):
    """Lignes a à b (numérotation de l'éditeur, bornes comprises) du script du rapport."""
    return "\n".join(SRC[a - 1:b])


ENTETE = "\n".join([lignes(82, 87), lignes(90, 91), lignes(94, 94), lignes(96, 104), lignes(136, 187), lignes(238, 375), lignes(396, 411), lignes(602, 635)])
debut = next(i for i, l in enumerate(SRC, 1) if l.startswith("tmp = SORTIE.replace"))
fin = next(i for i, l in enumerate(SRC, 1) if l.startswith("finaliser(SORTIE)"))
MOTEUR = lignes(debut, fin + 1)
for avant, apres in [('"Audit Comptable et financier FPPN", "Text Placeholder 4": "23 octobre 2026"', '"Audit Comptable et financier FPPN", "Text Placeholder 4": "26 octobre 2026"'),
                     ('r.text = "Rapport actuariel et statistique"', 'r.text = "Résultats, hypothèses et conclusions"'),
                     ('entrees = [("Résumé exécutif", 3)]', 'entrees = [("Synthèse des quatre axes", 3)]')]:
    assert avant in MOTEUR, avant
    MOTEUR = MOTEUR.replace(avant, apres)

CONTENU = r'''
VALIDITE = {}
PBCM = {}
GAB = {"couverture": 1, "sommaire": 5, "separateur": 2, "contenu": 5, "annexe": 5, "merci": 110}
R_ = "Source : rapport actuariel et statistique du 23/10/2026, diapositives "


def t(titre, sous_titre, synthese, entetes, largeurs, lignes_, source, notes="", taille=9.5, analyse=None, valable=None):
    """Diapositive de contenu : en-tête standard et un tableau, avec un paragraphe d'analyse facultatif."""
    k = {"tableau": {"titre": None, "entetes": entetes, "largeurs": largeurs, "taille": taille, "sans_code": True, "lignes": lignes_}, "source": source}
    if analyse:
        k["analyse"] = analyse
    if valable:
        k["valable"] = valable
    contenu({"titre": titre, "sous_titre": sous_titre, "synthese": synthese}, notes, **k)


def annexe(titre, entetes, largeurs, lignes_, source, taille=8.5, notes=""):
    SL.append({"type": "annexe", "titre": titre, "notes": notes,
               "tableau": {"titre": None, "entetes": entetes, "largeurs": largeurs, "taille": taille, "sans_code": True, "lignes": lignes_}, "source": source})


SL.append({"type": "couverture", "notes": "Résultats exhaustifs, hypothèses et conclusions des quatre axes de la mission actuarielle, chiffres du rapport du 23/10/2026. Méthodes expliquées en annexe."})
SL.append({"type": "sommaire", "notes": ""})

t("Synthèse des quatre axes", "Engagements et provisions partiellement démontrés, adossement et placements insuffisants",
  "Le FPPN connaît ses dettes de court terme, mais pas la valeur de ses engagements d'épargne : c'est la cause commune des verdicts 1, 3 et 4",
  ["Axe", "Verdict", "Résultat clé", "Conclusion"], [5.0, 4.0, 11.0, 11.2],
  [["1. Engagements envers les adhérents", "Partiellement démontré", "Épargne due 15 743 516 391 FCFA au 31/08/2026, 19 022 429 629 FCFA avec les intérêts dus", "L'épargne représente 97,46 % des engagements mesurés, et ses intérêts ne sont pas comptabilisés"],
   ["2. Suffisance des provisions", "Partiellement démontré", "Provisions maladie et obsèques 4 829 298 290 FCFA, besoins mesurés 19 581 392 FCFA hors dettes déjà comptabilisées", "Le total suffit en montant, mais aucune provision n'est calculée et deux manquent"],
   ["3. Adossement actif-passif", "Insuffisant", "Actif à moins d'un an : 34,42 % des paiements de l'année. Manque cumulé à cinq ans : 12 596 052 822 FCFA", "L'épargne exigible finance des actifs longs et peu liquides"],
   ["4. Adéquation des placements", "Insuffisant", "Placements moyens 2025 : 763 752 135 FCFA, soit 5,53 % de l'épargne due. Rendement moyen 2021 à 2025 : 4,90 %", "Aucune des quatre conditions de l'article 17 n'est démontrée"]],
  R_ + "4 à 8", taille=10)

# ================================================================ AXE 1
separateur(1, "Engagements", "Résultats, hypothèses et conclusions")
t("Engagements : résultats", "Les engagements mesurés atteignent 19 518 949 298 FCFA au 31/08/2026",
  "L'épargne et ses intérêts dus en représentent 97,46 %. Les autres engagements sont des dettes de quelques mois",
  ["Engagement", "Résultat", "Ce que le chiffre montre"], [9.6, 6.4, 15.2],
  [["Épargne PEP et EFC due au 31/08/2026 (41195100 et 41195200)", "15 743 516 391 FCFA", "Cotisations d'épargne nettes des sorties, sans intérêts"],
   ["Intérêts dus du 01/01/2020 au 31/08/2026 (PEP 3,50 %, EFC 5,50 %)", "3 278 913 238 FCFA", "Intérêts promis par le manuel d'épargne, absents des comptes"],
   ["Épargne avec intérêts dus", "19 022 429 629 FCFA", "Engagement d'épargne réel au taux du manuel"],
   ["Intérêts dus si l'EFC est servie à 3,50 % (borne basse)", "2 541 799 542 FCFA", "Taux déclaré par le chef comptable"],
   ["Allocations décès et frais funéraires dus (41591000 et 41592000)", "489 419 669 FCFA", "Liste du FPPN des décès non payés, valorisés au barème"],
   ["Décès non encore déclarés (états 2021 et 2023)", "7 100 000 FCFA", "Six décès d'une année enregistrés l'année suivante"],
   ["Soins faits, non encore facturés (clôture 2024)", "12 481 392 FCFA", "Dette maladie non encore facturée à la clôture"],
   ["Provision du personnel du FPPN au 31/12/2025 (hors adhérents)", "24 446 566 FCFA", "Compte 196 du FPPN : 43 666 678 FCFA"],
   ["Santé des policiers retraités", "Non mesurable", "610 retraités cotisants en août 2026, date de naissance connue pour 105 seulement"]],
  R_ + "5, 98, 112, 123, 126, 127, 133 et 143", taille=9)
t("Engagements : intérêts dus par exercice", "Aucun intérêt n'est comptabilisé : il n'existe ni charge ni dette d'intérêts dans les balances",
  "Au 31/12/2025, les intérêts dus atteignent déjà 2 804 987 325 FCFA, soit 16,1 fois les produits financiers de 2021 à 2025",
  ["Exercice", "PEP moyen", "EFC moyenne", "Intérêts PEP (3,50 %)", "Intérêts EFC (5,50 %)", "Total de l'année", "Total si EFC à 3,50 %"], [3.0, 4.6, 4.6, 4.6, 4.6, 4.9, 4.9],
  [["2020", "7 539 594 956 FCFA", "3 490 677 433 FCFA", "263 885 823 FCFA", "191 987 259 FCFA", "455 873 082 FCFA", "386 059 534 FCFA"],
   ["2021", "6 864 293 020 FCFA", "3 623 520 438 FCFA", "240 250 256 FCFA", "199 293 624 FCFA", "439 543 880 FCFA", "367 073 471 FCFA"],
   ["2022", "5 636 626 650 FCFA", "4 119 506 634 FCFA", "197 281 933 FCFA", "226 572 865 FCFA", "423 854 798 FCFA", "341 464 665 FCFA"],
   ["2023", "4 235 501 312 FCFA", "5 007 709 623 FCFA", "148 242 546 FCFA", "275 424 029 FCFA", "423 666 575 FCFA", "323 512 383 FCFA"],
   ["2024", "3 695 214 698 FCFA", "6 258 105 806 FCFA", "129 332 514 FCFA", "344 195 819 FCFA", "473 528 334 FCFA", "348 366 218 FCFA"],
   ["2025", "4 397 772 850 FCFA", "7 901 792 848 FCFA", "153 922 050 FCFA", "434 598 607 FCFA", "588 520 656 FCFA", "430 484 799 FCFA"],
   ["2026 (8 mois)", "5 097 233 672 FCFA", "9 681 558 005 FCFA", "118 935 452 FCFA", "354 990 460 FCFA", "473 925 913 FCFA", "344 838 472 FCFA"],
   ["Total", "", "", "1 251 850 574 FCFA", "2 027 062 663 FCFA", "3 278 913 238 FCFA", "2 541 799 542 FCFA"]],
  R_ + "98", taille=8.5)
t("Engagements : hypothèses", "Chaque hypothèse est tirée d'un document du FPPN ou retenue pour ne pas surévaluer la dette",
  "Les hypothèses donnent un minimum : intérêts non capitalisés et intérêts antérieurs à 2020 non comptés",
  ["Hypothèse", "Valeur retenue", "Source ou justification"], [8.4, 8.0, 14.8],
  [["Période", "01/01/2020 au 31/08/2026", "Demande de la mission. 2020 approché par les soldes du 01/09 et du 31/12, seule balance reçue pour 2020"],
   ["Montant de l'épargne", "Soldes des comptes 41195100 (PEP) et 41195200 (EFC)", "Aucun relevé individuel transmis : la balance est la seule source complète"],
   ["Taux servi sur le PEP", "3,50 % par an", "« MANUEL DE PRODUIT D'EPARGNE FPPN VF VF.docx »"],
   ["Taux servi sur l'EFC", "5,50 % par an, borne basse 3,50 %", "Manuel d'épargne et cartographie des risques. 3,50 % déclaré par le chef comptable. 14 % du cahier des charges 2021 non retenu (période non précisée)"],
   ["Calcul des intérêts", "Intérêts simples sur l'encours moyen de l'année", "Encours moyen = moyenne des soldes d'ouverture et de clôture. Pas de capitalisation, faute de règle écrite"],
   ["Allocations décès", "Barème de 1 400 000, 1 600 000 ou 2 200 000 FCFA selon le grade", "États de fin d'année du FPPN 2021 à 2025, comptes 41591000 et 41592000"],
   ["Décès tardifs", "7 100 000 FCFA", "Seuls les états 2021 et 2023 portent la date d'enregistrement"],
   ["Provision du personnel", "Unités de crédit projetées, agent par agent", "Décret n° 2017-210, DISA 2021, 2022, 2023 et 2025"]],
  R_ + "95, 98, 125, 127 et 133", taille=9)
t("Engagements : conclusions", "Le FPPN ne connaît pas la valeur de son principal engagement, l'épargne de ses adhérents",
  "Le verdict deviendra « démontré » quand le taux servi sera arrêté et les comptes individuels transmis",
  ["Conclusion", "Implication et risque pour le FPPN", "Recommandation"], [10.4, 11.0, 9.8],
  [["L'épargne due vaut 15 743 516 391 FCFA, et 19 022 429 629 FCFA avec les intérêts dus", "La dette est sous-estimée de 3 278 913 238 FCFA. Le résultat et les fonds propres sont surévalués d'autant", "Comptabiliser les intérêts dus chaque année (recommandation 2)"],
   ["Trois taux coexistent pour l'EFC (3,50 %, 5,50 %, 14 %), aucun n'est adopté par une décision signée", "L'engagement ne peut pas être évalué, et un épargnant peut contester le montant versé", "Faire arrêter le taux servi par le conseil d'administration (recommandation 2)"],
   ["Même à 3,50 %, les intérêts d'une année sur l'EFC (309 709 577 FCFA en 2025) dépassent les produits financiers de chaque année", "Chaque point de taux servi crée une perte que les garanties maladie et décès doivent combler", "Rendre le taux servi cohérent avec le rendement des actifs (recommandation 5)"],
   ["Le rapport de gestion 2025 présente 4 387 823 635 FCFA d'épargne, contre 13 814 066 964 FCFA dus", "L'Assemblée générale reçoit une information qui minore l'engagement", "Présenter le solde dû, distinct de la collecte de l'année"],
   ["L'engagement santé des retraités n'est pas mesurable", "Le FPPN ne sait pas si la cotisation des retraités couvre leurs soins à vie", "Étude actuarielle des risques longs tous les cinq ans (recommandation 8)"]],
  R_ + "95 à 99, 112 et 143", taille=9)

# ================================================================ AXE 2
separateur(2, "Provisions", "Résultats, hypothèses et conclusions")
t("Provisions : résultats", "Le FPPN a trop provisionné pour la maladie et le décès, et rien pour les intérêts et les prêts",
  "Écart total sur les cinq postes : 1 536 007 381 FCFA (montant du FPPN moins recalcul)",
  ["Poste", "Montant du FPPN", "Recalcul", "Écart absolu", "Écart relatif"], [9.4, 5.6, 5.4, 6.0, 4.8],
  [["Provision maladie (15801000)", "2 867 602 878 FCFA", "12 481 392 FCFA", "2 855 121 486 FCFA", "22 875,03 %"],
   ["Provision obsèques (15802000)", "1 961 695 412 FCFA", "7 100 000 FCFA", "1 954 595 412 FCFA", "27 529,51 %"],
   ["Départ à la retraite du personnel (196)", "43 666 678 FCFA", "24 446 566 FCFA", "19 220 112 FCFA", "78,62 %"],
   ["Intérêts dus sur l'épargne", "0 FCFA", "3 278 913 238 FCFA", "-3 278 913 238 FCFA", "-100,00 %"],
   ["Dépréciation des prêts du Grenier", "0 FCFA", "14 016 392 FCFA", "-14 016 392 FCFA", "-100,00 %"],
   ["Risques courts (statuts, art. 88)", "Aucune", "0 FCFA", "0 FCFA", "-"]],
  R_ + "128 et 145", taille=10,
  analyse=["**Deux lectures du besoin décès.** Hors dettes déjà comptabilisées, le besoin maladie et décès vaut 19 581 392 FCFA, et l'excédent 4 809 716 898 FCFA. Si la provision obsèques devait aussi couvrir les décès connus (489 419 669 FCFA, déjà en dette au 41591000), l'excédent serait de 4 327 397 229 FCFA (diapositive 6)."])
t("Provisions : maladie", "Le délai moyen de paiement des cliniques passe de 53,38 mois en 2020 à 2,62 mois en 2025",
  "La provision maladie nécessaire est de quelques millions de FCFA. Celle du bilan vaut 2 867 602 878 FCFA depuis 2020",
  ["Exercice", "Factures enregistrées", "Non payées en fin de période", "Délai moyen (mois)", "Prestations maladie"], [3.6, 7.0, 7.4, 4.6, 8.6],
  [["2020", "313 578 298 FCFA", "1 394 774 547 FCFA", "53,38", "205 812 056 FCFA"],
   ["2021", "103 385 027 FCFA", "875 220 557 FCFA", "101,59", "101 453 171 FCFA"],
   ["2022", "104 085 535 FCFA", "533 144 348 FCFA", "61,47", "115 237 773 FCFA"],
   ["2023", "312 091 366 FCFA", "271 270 756 FCFA", "10,43", "521 004 362 FCFA"],
   ["2024", "613 863 474 FCFA", "236 353 873 FCFA", "4,62", "1 060 419 627 FCFA"],
   ["2025", "1 107 859 142 FCFA", "242 140 869 FCFA", "2,62", "1 771 635 361 FCFA"],
   ["31/08/2026", "1 104 671 563 FCFA", "690 234 451 FCFA", "5,00", "1 646 039 117 FCFA"]],
  R_ + "121 à 123", taille=9.5,
  analyse=["**Provision nécessaire (estimation élargie).** 5 644 139 FCFA au 31/12/2022, 2 599 652 FCFA au 31/12/2023 et 12 481 392 FCFA au 31/12/2024. Les factures d'une année enregistrées l'année suivante ne représentent que 0,89 %, 0,21 % et 0,14 % des factures."])
t("Provisions : décès", "La provision obsèques vaut 4,0 fois la dette décès calculée par le FPPN au 31/08/2026",
  "La dette décès est déjà inscrite aux comptes 41591000 et 41592000 : la provision obsèques fait double emploi",
  ["Exercice", "Allocations ajoutées", "Allocations payées", "Dette allocations", "Dette frais funéraires", "Dette décès totale", "Provision obsèques"], [3.2, 4.4, 4.4, 4.6, 4.2, 4.6, 5.8],
  [["2021", "370 503 332 FCFA", "0 FCFA", "370 503 332 FCFA", "0 FCFA", "370 503 332 FCFA", "1 961 695 412 FCFA"],
   ["2022", "97 653 333 FCFA", "87 638 748 FCFA", "380 517 917 FCFA", "0 FCFA", "380 517 917 FCFA", "1 961 695 412 FCFA"],
   ["2023", "120 690 000 FCFA", "64 112 728 FCFA", "437 095 189 FCFA", "0 FCFA", "437 095 189 FCFA", "1 961 695 412 FCFA"],
   ["2024", "140 800 000 FCFA", "185 943 302 FCFA", "391 951 887 FCFA", "0 FCFA", "391 951 887 FCFA", "1 961 695 412 FCFA"],
   ["2025", "130 600 000 FCFA", "107 207 220 FCFA", "415 344 667 FCFA", "17 100 000 FCFA", "432 444 667 FCFA", "1 961 695 412 FCFA"],
   ["31/08/2026", "111 000 000 FCFA", "69 324 998 FCFA", "457 019 669 FCFA", "32 400 000 FCFA", "489 419 669 FCFA", "1 961 695 412 FCFA"]],
  R_ + "125 à 127", taille=8.5,
  analyse=["**Chaque année, le FPPN ajoute plus d'allocations qu'il n'en paie.** Le stock contient donc des dossiers de plusieurs années, à justifier un par un avec la date de dépôt pour écarter ceux qui sont forclos. Les décès déclarés après la clôture représentent 7 100 000 FCFA sur deux états."])
t("Provisions : hypothèses", "Les méthodes suivent les délais réels de paiement et les obligations des textes",
  "Le besoin est mesuré sur ce qui arrive après la clôture, sans double compte avec les dettes déjà inscrites",
  ["Hypothèse", "Valeur retenue", "Source ou justification"], [8.4, 8.6, 14.2],
  [["Maladie : périmètre de la provision", "Soins de l'année facturés après la clôture", "Les factures reçues non payées sont déjà au passif (compte 40110000)"],
   ["Maladie : mesure", "Factures de l'année enregistrées de janvier à mars suivant (journal SIN)", "Règlement intérieur, art. 52 : paiement des cliniques en trois mois. Délai observé de 2,62 à 5,00 mois depuis 2024"],
   ["Maladie : estimation élargie", "Extrapolée aux factures sans année lisible et aux prestations hors factures de cliniques", "L'année n'est lisible que pour 15,57 % des factures du grand livre 2025"],
   ["Décès : périmètre", "Décès survenus non déclarés. Les décès connus restent en dette au 41591000", "Règlement d'exécution n° 004/2022, art. 10 : dossier par dossier, plus les sinistres tardifs"],
   ["Décès : délai de déclaration", "Forclusion de 3 ans (membre) et 12 mois (ayant droit)", "Règlement intérieur, art. 58 et 60"],
   ["Risques courts", "Provision nulle", "Cotisation précomptée sur le mois qu'elle couvre (statuts, art. 88)"],
   ["Prêts du Grenier", "40 % de 3 à 6 mois de retard, 80 % de 6 à 12 mois, 100 % de 12 à 24 mois", "Politique de crédit du Grenier, chapitre XI.3"],
   ["Intérêts dus", "PEP 3,50 %, EFC 5,50 %, du 01/01/2020 au 31/08/2026", "Manuel d'épargne (voir l'axe 1)"]],
  R_ + "120 à 129 et 144", taille=9)
t("Provisions : conclusions", "Les provisions suffisent en montant, mais elles sont mal affectées et non calculées",
  "Le risque n'est pas un manque global de provisions : les engagements d'épargne, les plus importants, sont les moins provisionnés",
  ["Conclusion", "Implication et risque pour le FPPN", "Recommandation"], [10.4, 11.0, 9.8],
  [["Provisions maladie et obsèques identiques de 2020 au 31/08/2026, sans calcul", "Non-respect des statuts (art. 89) et du règlement d'exécution n° 004/2022 (art. 7 et 10). Suffisance non démontrable à l'AIRMS", "Note de calcul annuelle signée, présentée à l'Assemblée générale (recommandation 1)"],
   ["Maladie et obsèques : 4 809 716 898 FCFA au-dessus des besoins mesurés", "Résultat et fonds propres sous-évalués, sinistralité mal mesurée. L'excédent est en réalité une réserve", "Reclassement après décision des organes et information de l'AIRMS (recommandation 1)"],
   ["Intérêts dus (3 278 913 238 FCFA) et dépréciation des prêts (14 016 392 FCFA) non provisionnés", "Dettes absentes du bilan. Marge de solvabilité et réserve obligatoire surévaluées", "Comptabiliser les intérêts et déprécier les prêts chaque année (recommandations 2 et 7)"],
   ["Provision du personnel : 43 666 678 FCFA au bilan pour 24 446 566 FCFA recalculés", "Écart faible, méthode non documentée", "Transmettre la note de calcul (PBC, item 78)"],
   ["La base des prestations datée manque", "Pas de calcul dossier par dossier ni de boni et mali de liquidation", "Transmettre la base (recommandation 3)"]],
  R_ + "144 à 146", taille=9)

# ================================================================ AXE 3
separateur(3, "Adossement actif-passif", "Résultats, hypothèses et conclusions")
t("Adossement : structure du bilan", "Plus des trois quarts de l'actif sont à plus de cinq ans ou à délai incertain",
  "Au 31/08/2026, l'actif à moins d'un an ne couvre pas le passif à moins d'un an : l'écart vaut 3 559 869 546 FCFA avant toute sortie d'épargne",
  ["Poste", "31/12/2025", "31/08/2026"], [13.2, 9.0, 9.0],
  [["Actif immédiat (trésorerie, DAT)", "1 464 171 543 FCFA", "2 373 558 814 FCFA"],
   ["Actif à moins d'un an", "1 250 837 284 FCFA", "1 084 308 137 FCFA"],
   ["Actif à 1 à 5 ans (prêts du Grenier, créances)", "1 198 959 085 FCFA", "1 297 315 010 FCFA"],
   ["Actif à plus de 5 ans ou incertain (clinique, terrains, créance sur l'hôpital)", "16 040 901 123 FCFA", "17 170 238 905 FCFA"],
   ["Passif à moins d'un an (hors épargne)", "6 535 377 593 FCFA", "7 017 736 497 FCFA"],
   ["Épargne, échéance à établir", "13 814 066 964 FCFA", "15 743 516 391 FCFA"],
   ["Épargne couverte par des actifs réservés", "11,64 %", "16,28 %"]],
  R_ + "90 et 149", taille=10)
t("Adossement : couverture des engagements", "Même en lecture comptable, l'actif ne couvre pas tous les engagements connus",
  "Au 31/08/2026 : 96,33 % en lecture comptable, 27,19 % sans les actifs dont la valeur de marché n'est pas établie",
  ["Scénario", "Actif retenu 31/12/2025", "Couverture 2025", "Actif retenu 31/08/2026", "Couverture 2026"], [9.6, 6.0, 3.8, 6.0, 3.8],
  [["1. Lecture comptable", "19 954 869 035 FCFA", "98,06 %", "21 925 420 866 FCFA", "96,33 %"],
   ["2. Sans actifs de valeur non réalisable", "15 732 740 110 FCFA", "77,31 %", "17 237 600 156 FCFA", "75,73 %"],
   ["3. Sans actifs de valeur de marché non établie", "5 298 475 488 FCFA", "26,04 %", "6 188 677 951 FCFA", "27,19 %"],
   ["4. Actifs à un an / passif à un an", "2 715 008 827 FCFA", "41,54 %", "3 457 866 951 FCFA", "49,27 %"],
   ["Engagements connus (scénarios 1 à 3)", "20 349 444 557 FCFA", "", "22 761 252 888 FCFA", ""]],
  R_ + "151 et 152", taille=10,
  analyse=["**Ces couvertures sont des maxima.** Les engagements non comptabilisés, dont les intérêts dus sur l'épargne, ne sont pas au dénominateur."])
t("Adossement : sorties de l'épargne", "De 2021 à 2025, il sort chaque année en moyenne 19,24 % de l'épargne du début d'année",
  "Taux de sortie compris entre 12,60 % et 27,55 %. Le PEP sort plus vite que l'EFC",
  ["Exercice", "Épargne au début", "Sorties de l'année", "Taux PEP", "Taux EFC", "Taux total"], [3.6, 7.6, 7.6, 4.0, 4.0, 4.4],
  [["2021", "10 828 205 640 FCFA", "1 652 904 458 FCFA", "17,73 %", "10,02 %", "15,26 %"],
   ["2022", "10 147 421 276 FCFA", "2 317 974 441 FCFA", "30,55 %", "9,88 %", "22,84 %"],
   ["2023", "9 364 845 293 FCFA", "2 579 886 392 FCFA", "46,17 %", "7,02 %", "27,55 %"],
   ["2024", "9 121 576 576 FCFA", "1 634 861 234 FCFA", "33,45 %", "7,99 %", "17,92 %"],
   ["2025", "10 785 064 432 FCFA", "1 358 821 103 FCFA", "23,60 %", "6,54 %", "12,60 %"],
   ["Moyenne 2021 à 2025", "", "", "", "", "19,24 %"]],
  R_ + "154", taille=10)
t("Adossement : écarts de liquidité", "Le manque de liquidité cumulé atteint 12 596 052 822 FCFA à la fin de la cinquième année",
  "Même en comptant à leur valeur comptable les terrains, la clinique et les créances longues, il manque 835 832 022 FCFA",
  ["Échéance", "Actifs disponibles", "Dettes à payer", "Sorties d'épargne", "Écart de l'année", "Écart cumulé"], [4.4, 5.4, 5.0, 5.2, 5.6, 5.6],
  [["Année 1", "3 457 866 951 FCFA", "7 017 736 497 FCFA", "3 028 376 687 FCFA", "-6 588 246 233 FCFA", "-6 588 246 233 FCFA"],
   ["Année 2", "324 328 752 FCFA", "0 FCFA", "2 445 847 020 FCFA", "-2 121 518 268 FCFA", "-8 709 764 501 FCFA"],
   ["Année 3", "324 328 752 FCFA", "0 FCFA", "1 975 371 053 FCFA", "-1 651 042 301 FCFA", "-10 360 806 802 FCFA"],
   ["Année 4", "324 328 752 FCFA", "0 FCFA", "1 595 394 465 FCFA", "-1 271 065 713 FCFA", "-11 631 872 515 FCFA"],
   ["Année 5", "324 328 752 FCFA", "0 FCFA", "1 288 509 060 FCFA", "-964 180 307 FCFA", "-12 596 052 822 FCFA"],
   ["Au-delà de 5 ans", "17 170 238 905 FCFA", "0 FCFA", "5 410 018 105 FCFA", "11 760 220 800 FCFA", "-835 832 022 FCFA"]],
  R_ + "155", taille=9,
  analyse=["**En pratique, l'écart est comblé chaque année par les nouvelles cotisations d'épargne et les excédents.** Le FPPN dépend de la poursuite des versements pour payer les sorties."])
t("Adossement : tests de résistance", "La couverture des paiements de l'année varie entre 22,30 % et 34,42 % selon le scénario",
  "Aucun scénario n'approche 100 %, même dans le scénario central",
  ["Clôture", "Scénario", "Actif à moins d'un an", "Besoin de l'année", "Écart", "Couverture"], [3.6, 9.4, 5.0, 5.0, 5.0, 3.2],
  [["31/12/2025", "Central : taux de sortie moyen (19,24 %)", "2 715 008 827 FCFA", "9 192 611 042 FCFA", "-6 477 602 215 FCFA", "29,53 %"],
   ["31/12/2025", "Résistance : taux de sortie maximal (27,55 %)", "2 715 008 827 FCFA", "10 340 963 512 FCFA", "-7 625 954 685 FCFA", "26,25 %"],
   ["31/12/2025", "Résistance sévère : taux maximal et prestations maladie +103,53 %", "2 715 008 827 FCFA", "12 175 203 736 FCFA", "-9 460 194 909 FCFA", "22,30 %"],
   ["31/08/2026", "Central : taux de sortie moyen (19,24 %)", "3 457 866 951 FCFA", "10 046 113 184 FCFA", "-6 588 246 233 FCFA", "34,42 %"],
   ["31/08/2026", "Résistance : taux de sortie maximal (27,55 %)", "3 457 866 951 FCFA", "11 354 859 261 FCFA", "-7 896 992 310 FCFA", "30,45 %"],
   ["31/08/2026", "Résistance sévère : taux maximal et prestations maladie +103,53 %", "3 457 866 951 FCFA", "13 189 099 485 FCFA", "-9 731 232 534 FCFA", "26,22 %"]],
  R_ + "156", taille=9,
  analyse=["**Ratios prudentiels liés au 31/08/2026.** Liquidité 0,19 (norme au moins 1), équilibre stable 1,72 (norme au plus 1), réserve de trésorerie des risques longs 0,00 (norme un trimestre)."])
t("Adossement : hypothèses", "Les hypothèses sont calibrées sur le comportement observé des adhérents",
  "Toutes les corrections attendues augmentent le passif ou réduisent l'actif : le résultat est une borne favorable au FPPN",
  ["Hypothèse", "Valeur retenue", "Source ou justification"], [8.0, 9.0, 14.2],
  [["Valeur des actifs", "Valeur comptable des balances", "Aucune expertise de la clinique ni des terrains. Choix favorable au FPPN"],
   ["Délai de mobilisation des actifs", "Classement par nature comptable", "Pas d'échéancier contractuel des actifs"],
   ["Actifs à 1 à 5 ans", "Répartis à parts égales sur les années 2 à 5 (324 328 752 FCFA par an)", "Prêts du Grenier de durée médiane 10 mois, créances sans échéancier"],
   ["Échéance de l'épargne", "Sortie annuelle de 19,24 % du solde restant (moyenne 2021 à 2025)", "Sorties des comptes 41195100 et 41195200. Aucun échéancier individuel"],
   ["Test de résistance", "Taux de sortie maximal observé : 27,55 % (2023)", "Calibré sur le FPPN, et non sur le rachat de 40 % de la formule standard européenne"],
   ["Test de résistance sévère", "27,55 % et hausse des prestations maladie de 103,53 %", "Plus forte hausse observée en 2024 et 2025"],
   ["Dettes à moins d'un an", "7 017 736 497 FCFA, toutes en année 1", "Comprennent les provisions non recalculées. Seule leur baisse allégerait le passif"],
   ["Nouveaux versements", "Non comptés", "Écoulement statique : on mesure ce que le bilan actuel permet de payer"]],
  R_ + "149, 150 et 153 à 156", taille=9)
t("Adossement : conclusions", "Le FPPN n'est pas adossé : l'épargne exigible finance des actifs longs et peu liquides",
  "Le non-adossement est démontré quel que soit le scénario. Un échéancier de l'épargne préciserait les montants sans changer la conclusion",
  ["Conclusion", "Implication et risque pour le FPPN", "Recommandation"], [10.4, 11.0, 9.8],
  [["La liquidité ne couvre pas l'année qui vient : 34,42 % dans le scénario central au 31/08/2026", "Risque de liquidité : une hausse des départs ou des retraits empêche de payer à l'échéance", "Réserve de liquidité égale aux sorties d'épargne d'une année au taux maximal (recommandation 5)"],
   ["L'écart ne se résorbe pas : manque cumulé de 12 596 052 822 FCFA à cinq ans", "Le FPPN dépend des nouveaux versements pour payer les sorties", "Arrêter le financement d'actifs longs par l'épargne (recommandation 7)"],
   ["Le FPPN ne résiste pas à un choc : 26,22 % dans le scénario sévère", "Il devrait céder des actifs peu liquides, avec des retards de paiement aux adhérents", "Échéancier de l'épargne et suivi trimestriel des écarts"],
   ["Ni politique, ni comité, ni limites actif-passif", "Personne ne mesure l'effet des investissements sur la liquidité", "Politique actif-passif votée et comité actif-passif (recommandations 5 et 9)"],
   ["Clinique, terrains et créance sur l'hôpital : 15 736 951 165 FCFA, non évalués", "La couverture tombe à 27,19 % sans ces actifs", "Faire évaluer les actifs immobiliers et documenter le recouvrement de la créance"]],
  R_ + "150 à 158", taille=9)

# ================================================================ AXE 4
separateur(4, "Placements", "Résultats, hypothèses et conclusions")
t("Placements : encours", "Un dépôt mensuel pour payer l'EFC, un dépôt PEP de cinq ans, un FCP depuis 2025",
  "Encours moyen réellement placé en 2025 : 763 752 135 FCFA. Au 31/08/2026 : DAT BOA 485 541 804 FCFA et FCP 500 000 000 FCFA",
  ["Exercice", "DAT EFC : encours moyen", "DAT PEP : encours moyen", "DAT Orabank et FCP", "Total : encours moyen", "Total à la clôture"], [3.6, 5.6, 5.6, 5.0, 5.6, 5.8],
  [["2021", "196 540 324 FCFA", "505 383 858 FCFA", "0 FCFA", "701 924 182 FCFA", "505 383 858 FCFA"],
   ["2022", "173 255 937 FCFA", "505 383 858 FCFA", "24 640 000 FCFA", "703 279 795 FCFA", "563 924 385 FCFA"],
   ["2023", "136 861 631 FCFA", "505 383 858 FCFA", "17 440 000 FCFA", "659 685 489 FCFA", "505 383 858 FCFA"],
   ["2024", "211 637 502 FCFA", "505 383 858 FCFA", "0 FCFA", "717 021 360 FCFA", "505 383 858 FCFA"],
   ["2025", "266 146 860 FCFA", "418 153 219 FCFA", "79 452 055 FCFA", "763 752 135 FCFA", "500 000 000 FCFA"],
   ["2026 (8 mois)", "247 695 588 FCFA", "0 FCFA", "500 000 000 FCFA", "747 695 588 FCFA", "985 541 804 FCFA"]],
  R_ + "162", taille=9.5,
  analyse=["**Le DAT EFC n'est pas une réserve.** C'est un dépôt d'un mois, renouvelé chaque mois, puis viré en décembre pour payer l'EFC. Seuls le DAT PEP, puis le FCP, ont été des placements durables."])
t("Placements : rendement et volume", "Le problème des placements est leur faible volume, pas leur rendement",
  "Produits financiers 2021 à 2025 : 173 881 316 FCFA. Intérêts dus sur la même période : 2 349 114 243 FCFA",
  ["Mesure", "Valeur", "Ce que le chiffre montre"], [10.4, 6.6, 14.2],
  [["Placements moyens 2025 rapportés à l'épargne due au 31/12/2025", "5,53 %", "763 752 135 FCFA pour 13 814 066 964 FCFA d'épargne"],
   ["Rendement moyen apparent 2021 à 2025", "4,90 % par an", "173 881 316 FCFA de produits pour 709 132 592 FCFA d'encours moyen. Taux directeur BCEAO entre 2,00 % et 3,50 %"],
   ["Produits financiers 2025", "117 912 488 FCFA", "Dont 132 663 262 FCFA d'intérêts du DAT PEP accumulés depuis 2020, enregistrés en une fois"],
   ["Intérêts d'une année sur l'EFC à 3,50 % (2025)", "309 709 577 FCFA", "Supérieurs aux produits financiers de chaque année de 2021 à 2025"],
   ["Clinique, terrains et créance sur l'hôpital au 31/08/2026", "15 736 951 165 FCFA", "Actifs peu liquides, sans évaluation récente, financés en grande partie par l'épargne"],
   ["Financement des prêts du Grenier au 31/08/2026", "915 477 260 FCFA", "87,66 % de l'encours des prêts, durée médiane de 10 mois"],
   ["Comptes courants non rémunérés au 31/08/2026", "2 356 399 123 FCFA", "Trésorerie qui peut être placée à terme"],
   ["Politique de placement votée", "Aucune", "Le directeur général décide seul (entretien du 19/10/2026), à confirmer par pièce"]],
  R_ + "8, 97, 102, 162, 163 et 165", taille=9)
t("Placements : hypothèses", "Le rendement est mesuré sur l'argent réellement placé chaque jour",
  "Les règles appliquées sont celles de la mutualité sociale de l'UEMOA, régime du FPPN",
  ["Hypothèse", "Valeur retenue", "Source ou justification"], [8.0, 9.2, 14.0],
  [["Périmètre des placements", "Dépôts à terme (comptes 52) et FCP (comptes 53)", "Grands livres 2021 à 2025 et 2026"],
   ["Encours", "Moyenne pondérée par les jours", "Les DAT sont placés et dénoués en cours d'année : les soldes de clôture fausseraient le rendement"],
   ["Produits", "Produits financiers de la classe 77", "Balances 2021 à 2025"],
   ["Référence de marché", "Taux directeur BCEAO et titres publics de l'État de Côte d'Ivoire", "Référence de placement sûr dans la zone"],
   ["FCP", "500 000 000 FCFA à la valeur d'achat", "Rendement annoncé de 8 à 10 %, non garanti, non revalorisé dans les comptes"],
   ["Conditions à respecter", "Sécurité, rendement, liquidité, diversification", "Règlement d'exécution n° 004/2022, art. 17. Décision ministérielle sur les actifs admis non reçue (PBC, item 32)"]],
  R_ + "151, 162 et 165", taille=9.5)
t("Placements : conclusions", "Les placements ne sont adéquats ni en volume, ni en liquidité, ni en rendement",
  "Le portefeuille est inadéquat d'abord parce que l'épargne finance des actifs qui ne produisent ni revenu financier ni liquidité",
  ["Condition (art. 17)", "Ce que montrent les chiffres", "Appréciation", "Recommandation"], [4.2, 13.0, 3.4, 10.6],
  [["Liquidité", "Couverture de 34,42 % des paiements de l'année au 31/08/2026", "Non remplie", "Part liquide minimale fixée par une politique votée (recommandation 5)"],
   ["Rendement", "Produits 2021 à 2025 : 173 881 316 FCFA, intérêts dus : 2 349 114 243 FCFA", "Non remplie", "Rendement minimal au regard du taux servi arrêté (recommandation 5)"],
   ["Sécurité", "Clinique, terrains et créance sur l'hôpital non évalués. FCP non garanti", "Non démontrée", "Évaluation des actifs, comptabilité séparée de la clinique (recommandation 7)"],
   ["Diversification", "Plus des trois quarts de l'actif à plus de cinq ans ou à délai incertain", "Non remplie", "Limites sur l'immobilier. Projet de Cocody présenté aux organes avant tout engagement"],
   ["Gouvernance", "Ni politique écrite, ni comité, ni délibération de placement reçue", "Non démontrée", "Placements décidés par le conseil d'administration (statuts, art. 85)"]],
  R_ + "163, 164 et 166", taille=9.5)

# ================================================================ CONCLUSIONS GÉNÉRALES
separateur(5, "Conclusions générales", "Ce que les quatre axes disent ensemble du FPPN")
t("Conclusions générales", "Le FPPN dégage des excédents, mais il ne peut pas démontrer qu'il remboursera l'épargne de ses adhérents",
  "Le risque principal est la liquidité de l'épargne, et l'enjeu réglementaire est l'agrément",
  ["Constat transversal", "Chiffre", "Conséquence"], [9.6, 8.6, 13.0],
  [["Fonds propres négatifs à chaque clôture", "-438 242 200 FCFA (2025), -879 498 700 FCFA (31/08/2026)", "Aucun coussin pour absorber un choc. Le déficit des huit premiers mois de 2026 inverse le redressement"],
   ["Ratios de la Décision n° 05/2022", "14 non conformes, 7 conformes, 15 en attente", "Une non-conformité sur deux exercices de suite permet à l'AIRMS de retirer l'agrément (règlement n° 07/2009, art. 84)"],
   ["Épargne non couverte par des actifs réservés", "16,28 % de couverture au 31/08/2026", "L'épargne finance la clinique, les terrains, la créance sur l'hôpital et les prêts du Grenier"],
   ["Sinistralité maladie", "6,00 % (2021), 73,53 % (2025), 97,40 % (31/08/2026)", "La garantie maladie est déficitaire depuis 2025, le décès finance l'ensemble"],
   ["Données manquantes", "Base des prestations datée, dates de naissance, comptes individuels, taux servi", "Les montants seront précisés, sans changer le sens des verdicts"]],
  R_ + "9, 11, 161, 178 et 195", taille=9.5)
t("Recommandations", "Neuf recommandations, dont quatre à engager dans les six mois",
  "Chaque recommandation répond à un constat d'un ou plusieurs axes",
  ["Recommandation", "Axes concernés", "Horizon"], [17.2, 7.4, 6.6],
  [["1. Calculer les provisions maladie et décès à chaque clôture", "Provisions", "Moins de 6 mois"],
   ["2. Arrêter le taux servi sur l'épargne et tenir les comptes individuels", "Engagements, provisions", "Moins de 6 mois"],
   ["3. Transmettre les données qui manquent aux calculs actuariels", "Tous", "Moins de 6 mois"],
   ["4. Présenter à l'AIRMS un plan de mise en conformité prudentielle", "Tous", "Moins de 6 mois"],
   ["5. Constituer une réserve de liquidité et adopter une politique de placement", "Adossement, placements", "6 à 18 mois"],
   ["6. Rééquilibrer la garantie maladie", "Provisions", "6 à 18 mois"],
   ["7. Encadrer les investissements dans la clinique, le Grenier et le projet de Cocody", "Adossement, placements", "6 à 18 mois"],
   ["8. Faire réaliser l'étude actuarielle des risques longs et certifier les comptes", "Engagements", "Plus de 18 mois"],
   ["9. Mettre en place un dispositif actuariel permanent", "Tous", "Plus de 18 mois"]],
  R_ + "25 et 26", taille=10)

# ================================================================ ANNEXES : MÉTHODES
separateur("A", "Annexes : méthodes", "Chaque méthode expliquée, pourquoi elle a été retenue et ses limites")
annexe("Méthodes : engagements envers les adhérents",
       ["Méthode", "Ce qu'elle fait", "Pourquoi elle est retenue", "Limite"], [6.0, 9.2, 9.0, 7.0],
       [["Inventaire des soldes comptables", "Additionne les soldes des comptes d'épargne PEP et EFC à chaque clôture", "Seule source complète de 2020 au 31/08/2026, en l'absence de relevés individuels", "Montant global, non vérifiable épargnant par épargnant"],
        ["Intérêts simples sur encours moyen", "Multiplie chaque année la moyenne des soldes d'ouverture et de clôture par le taux servi", "Tient compte des versements et retraits de l'année sans connaître leurs dates. Pas de capitalisation faute de règle écrite : résultat minimal", "Dépend du taux retenu (3,50 % ou 5,50 % pour l'EFC)"],
        ["Évaluation dossier par dossier au barème", "Valorise chaque décès non payé au barème du grade", "Méthode du FPPN, conforme aux statuts (art. 89), rapprochée du compte 41591000", "Date de dépôt non saisie : forclusion non vérifiable"],
        ["Décès enregistrés l'année suivante", "Mesure les décès d'une année déclarés après le 31 décembre", "Donne l'ordre de grandeur des décès survenus non déclarés", "Deux états seulement portent la date d'enregistrement"],
        ["Unités de crédit projetées", "Calcule agent par agent l'indemnité de départ acquise et projetée", "Méthode de référence pour les indemnités de départ (décret n° 2017-210)", "Hors policiers détachés. DISA 2024 scannée"],
        ["Provision mathématique (art. 334-2 du Code CIMA)", "Valeur actuelle des engagements moins valeur actuelle des versements futurs, compte par compte", "Méthode cible, présentée pour la suite", "Non applicable : comptes individuels, âges et taux arrêté manquants"]],
       R_ + "98, 124, 127, 133 et 143")
annexe("Méthodes : suffisance des provisions",
       ["Méthode", "Ce qu'elle fait", "Pourquoi elle est retenue", "Limite"], [6.0, 9.2, 9.0, 7.0],
       [["Règlements postérieurs à la clôture (maladie)", "Mesure les soins d'une année facturés de janvier à mars suivant, puis l'étend aux factures sans année lisible", "Les cliniques sont payées en quelques mois (règlement intérieur, art. 52 ; délai observé de 2,62 à 5,00 mois depuis 2024)", "Date des soins non saisie. Année lisible pour 15,57 % des factures 2025"],
        ["Délai moyen de paiement", "Divise les factures non payées par les factures enregistrées par mois", "Vérifie que la règle des trois mois s'applique, et donc qu'une méthode courte suffit", "Ne couvre pas les remboursements aux mutualistes"],
        ["Cadence des déclarations tardives (décès)", "Estime les décès survenus non encore déclarés à partir des délais de déclaration observés", "Méthode reconnue dans la zone (Code CIMA, art. 334-12, circulaire n° 00230). Tient compte de la forclusion de 3 ans et 12 mois", "Peu d'observations (deux états)"],
        ["Dépréciation par tranche de retard (Grenier)", "Applique 40 %, 80 % ou 100 % à l'encours selon l'ancienneté du retard", "Taux fixés par la politique de crédit du Grenier", "Barème de la BCEAO des SFD non reçu"],
        ["Provision pour risques courts (statuts, art. 88)", "Mesure la période non couverte par une cotisation à la clôture", "Obligation des statuts", "Nulle tant que la cotisation est précomptée chaque mois"],
        ["Méthode écartée : Chain Ladder", "Projette les paiements futurs à partir de triangles de paiements sur plusieurs années", "Écartée : inutile pour des soins réglés en quelques mois, et base datée non transmise", "-"]],
       R_ + "120 à 129, 144 et annexe « Méthodes non retenues »")
annexe("Méthodes : adossement actif-passif",
       ["Méthode", "Ce qu'elle fait", "Pourquoi elle est retenue", "Limite"], [6.0, 9.2, 9.0, 7.0],
       [["Classement des actifs par liquidité", "Range chaque actif selon le temps nécessaire pour le transformer en argent", "Méthode standard (Gestion des risques bancaires, p. 95). Possible avec la nature comptable de chaque actif", "Pas d'échéancier contractuel des actifs"],
        ["Écoulement conventionnel de l'épargne", "Fait sortir chaque année 19,24 % du solde d'épargne restant", "Méthode prévue pour les dépôts sans échéance (Gestion des risques bancaires, p. 33). Calibrée sur les sorties réelles de 2021 à 2025", "Remplace un échéancier individuel, à obtenir"],
        ["Écart de liquidité par échéance (gap)", "Compare, année par année, les actifs disponibles aux dettes et sorties d'épargne", "Mesure directement le manque de liquidité (Gestion des risques bancaires, p. 86)", "Statique : sans nouveaux versements"],
        ["Scénarios de couverture", "Rapporte l'actif retenu aux engagements selon quatre règles de valorisation tirées des textes", "Montre l'effet des actifs non évalués (règlement d'exécution n° 004/2022, art. 17 à 27)", "Valeur de marché des actifs non établie"],
        ["Tests de résistance", "Simule un retrait au taux maximal (27,55 %) et une hausse des prestations maladie de 103,53 %", "Chocs calibrés sur les extrêmes observés au FPPN", "Pas de données de sorties par motif"],
        ["Méthodes écartées : duration, immunisation, VaR, modèle stochastique", "Mesurent la sensibilité aux taux et le risque de perte", "Écartées : supposent des flux de passif connus, une courbe de taux et un portefeuille obligataire, absents au FPPN", "-"]],
       R_ + "149 à 156 et annexe « Méthodes non retenues »")
annexe("Méthodes : adéquation des placements",
       ["Méthode", "Ce qu'elle fait", "Pourquoi elle est retenue", "Limite"], [6.0, 9.2, 9.0, 7.0],
       [["Encours moyen pondéré par les jours", "Calcule l'argent réellement placé chaque jour, lu dans les grands livres", "Les DAT sont placés et dénoués en cours d'année : les soldes de clôture fausseraient le rendement", "Grands livres 2021 à 2026 seulement"],
        ["Taux de rendement", "Divise les produits financiers par l'encours moyen", "Formule de la Décision n° 05/2022 (taux brut de rentabilité des placements)", "Produits irréguliers (intérêts enregistrés en une fois en 2025)"],
        ["Taux de couverture de l'épargne par les placements", "Divise l'encours placé par l'épargne due", "Mesure le volume placé face à l'engagement", "Taux servi non arrêté"],
        ["Comparaison au marché", "Compare le rendement au taux directeur de la BCEAO et aux titres de l'État", "Référence de placement sûr dans la zone", "Sources secondaires pour les titres de l'État"],
        ["Comparaison aux intérêts dus", "Compare les produits financiers aux intérêts promis aux épargnants", "Vérifie si les placements peuvent servir le taux promis", "Dépend du taux retenu"],
        ["Contrôle des quatre conditions de l'article 17", "Teste la sécurité, le rendement, la liquidité et la diversification", "Exigence du règlement d'exécution n° 004/2022", "Décision ministérielle sur les actifs admis non reçue"]],
       R_ + "162, 165, 166 et 179")
SL.append({"type": "merci", "notes": ""})
'''

code = ENTETE + "\n" + CONTENU + "\n" + MOTEUR + "\nprint(SORTIE, len(SL), 'diapositives')\n"
g = {"__name__": "__main__", "__file__": os.path.join(ICI, "rapport_v12_fppn.py"), "SKILL": SKILL, "SORTIE": SORTIE}
exec(compile(code, "deck_resultats_axes_fppn (assemblé)", "exec"), g)
