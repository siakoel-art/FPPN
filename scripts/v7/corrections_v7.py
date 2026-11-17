# -*- coding: utf-8 -*-
"""Corrections du rapport FPPN v6 (16/11/2026) -> v7, validées le 17/11/2026 sur le diagnostic de clarté.
Utilisation : python corrections_v7.py <v6.pptx> <v7.pptx> <journal.xlsx>
Ne modifie jamais le fichier v6."""
import os
import re
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from pptx import Presentation

sys.path.insert(0, os.path.dirname(__file__))
from moteur_v7 import JOURNAL, Deck, Erreur, modifie, texte_para  # noqa: E402

V6, V7, JR = sys.argv[1:4]
assert os.path.abspath(V6) != os.path.abspath(V7)
d = Deck(Presentation(V6))
SY, ST, TI = "Adenka_Synthese", "Adenka_SousTitre", "Adenka_Titre"

# ---------------------------------------------------------------- 1. Constats à clarifier
# D9 : dépenses de la clinique inscrites en compte 47111105
d.reecrit(9, "TextBox 32", 1, "S'y ajoutent",
          "S'y ajoutent 1 280 300 860 FCFA de dépenses de la clinique payées par le Fonds et inscrites comme une somme que la "
          "clinique lui doit (compte 47111105) (4), au lieu d'une dépense. Elles ne deviennent des dépenses de soins que "
          "lorsqu'elles sont déduites des factures de la clinique (3). Chaque année, le Fonds paie plus qu'il ne déduit : le reste "
          "est passé de 469 980 456 FCFA fin 2022 à 1 280 300 860 FCFA au 31/08/2026 (4). Tant qu'il n'est pas déduit, les biens "
          "et le résultat du Fonds paraissent plus importants qu'ils ne sont. [[La règle de déduction est à confirmer par le service comptable.]]")
d.rep(9, "comptes 47 et 27 (11/11/2026)", "comptes 47 et 27 (11/11/2026). (4) « BALANCE FPPN 2022.xlsx » et « BALANCE FPPN DU "
      "01-01-2026 AU 31-08-2026.xlsx », compte 47111105")

# D80 : factures de 2025 comptées deux fois
d.reecrit(80, TI, 0, "Somme de 2025", "Factures de 2025 comptées deux fois")
d.reecrit(80, ST, 0, "Une somme", "Des factures de 2025 sont comptées une seconde fois dans les dépenses de 2026")
d.reecrit(80, SY, 0, "Après correction",
          "Après correction, la maladie dépense 82,31 FCFA pour 100 FCFA de cotisations, soit 82,31 %, et non 97,40")
d.rep(80, "Effet de l'extourne oubliée sur les chiffres de 2026", "Effet de la correction des factures comptées deux fois")
d.rep(80, "Après extourne", "Après correction")
d.rep(80, "Prestations maladie pour 100 FCFA de cotisations maladie", "Prestations maladie pour 100 FCFA de cotisations maladie, soit en %")
d.rep(80, "Prestations pour 100 FCFA de cotisations, selon le texte (norme 70 à 80)",
      "Prestations pour 100 FCFA de cotisations, soit en %, selon le texte (norme 70 % à 80 %)")
d.reecrit(80, "TextBox 32", 0, "Provision du 31/12/2025",
          "Le 31/12/2025, le Fonds attend les factures de sa clinique pour octobre, novembre et décembre 2025. Il inscrit donc "
          "255 000 000 FCFA en dépense de 2025 et en dette envers la clinique (écritures OD 805, 806 et 807, « CMP PROV FACT "
          "10/2025 » à « 12/2025 ») (1). Le 31/01/2026, les vraies factures de ces trois mois arrivent : 240 471 119 FCFA (pièces "
          "SIN 462, 552 et 619, « OCT 25 », « NOVEMBRE 2025 », « DECEMBRE ») (2). Le Fonds les inscrit en dépense de 2026 sans "
          "annuler les 255 000 000 FCFA de fin 2025. Les mêmes soins sont donc comptés en 2025 et en 2026, et la dette envers la "
          "clinique deux fois.")
d.reecrit(80, "TextBox 32", 1, "Le rapport retient",
          "Le rapport retient les chiffres corrigés : le Fonds doit annuler en 2026 les 255 000 000 FCFA de fin 2025 (écriture "
          "d'extourne). [[À confirmer par le service comptable et par le lettrage du compte 40110000, qui rapproche chaque facture "
          "de son paiement.]]")
d.rep(80, "extourne ajoutée par Adenka", "correction ajoutée par Adenka")

# D105 : intérêts du plan d'épargne pris sur l'épargne
d.reecrit(105, SY, 0, "198 millions",
          "198 millions de FCFA d'intérêts pris sur l'épargne de 2020 à 2026, jamais comptés en dépenses")
d.reecrit(105, "TextBox 32", 1, "En 2024, 355",
          "Quand un policier quitte le plan d'épargne, le Fonds lui rend ce qu'il a versé plus 3,5 %. Ces 3,5 % sont pris sur le "
          "compte d'épargne 41195100, comme s'ils faisaient partie des sommes versées, au lieu d'être comptés en dépense. De 09/2020 "
          "à 08/2026, ils font 198 040 675 FCFA. Le compte d'épargne affiche donc 198 040 675 FCFA de moins que ce que les autres "
          "policiers ont réellement versé, et les dépenses du Fonds sont trop basses du même montant.")
d.reecrit(105, "TextBox 32", 2, "Si un taux",
          "Pour l'entraide de fin de carrière, les 355 règlements de 2024 (442 318 346 FCFA) sont des montants ronds : [[on ne "
          "sait pas encore s'ils contiennent des intérêts]]. Si c'est le cas, la dette d'épargne inscrite au bilan est trop basse "
          "du montant de ces intérêts.")
d.rep(105, "Traitements Adenka", "Traitements Adenka. Grands livres 2020 à 2026, compte 41195100 : remboursements égaux au "
      "versé × 1,035, calcul Adenka")

# D108 : intérêts dus selon la règle
d.reecrit(108, "TextBox 32", 0, "Selon le taux",
          "Le montant dû dépend de la règle. Selon la règle appliquée par le Fonds (3,5 % du montant versé, une seule fois), il "
          "doit 183 023 590 FCFA sur le plan d'épargne au 31/08/2026, soit 3,5 % des 5 229 245 425 FCFA inscrits au compte "
          "41195100, payés au départ de chaque policier. Selon le manuel (3,5 % par an sur le plan d'épargne, 5,5 % par an sur "
          "l'entraide), il devrait 3 278 913 238 FCFA depuis le 01/01/2020. Aucun de ces montants n'est inscrit dans les comptes.")
d.reecrit(108, "TextBox 32", 1, "Au 31/12/2025",
          "Selon le manuel, les intérêts dus atteignent déjà 2 804 987 325 FCFA au 31/12/2025. [[La règle de l'entraide reste à "
          "recevoir.]]")

# D169 : effet des montants bloqués
d.reecrit(169, "TextBox 32", 0, "Les intérêts dus",
          "Les 4 829 298 290 FCFA mis de côté pour la maladie et les obsèques sont comptés comme des dettes à payer dans "
          "l'année, alors que le besoin estimé par Adenka n'est que d'au moins 19 581 392 FCFA (le FPPN ne calcule pas ses "
          "provisions). Les ramener au besoin change beaucoup la mesure : pour 100 FCFA à payer dans l'année, le Fonds dispose "
          "de 69,42 FCFA au lieu de 35,32, soit 69,42 % des paiements au lieu de 35,32 %.")
d.reecrit(169, "TextBox 32", 1, "Le FPPN n'est pas adossé",
          "Le manque reste réel : l'épargne à rendre dans l'année (3 028 376 687 FCFA) et les autres dettes de l'année "
          "(1 933 438 207 FCFA) dépassent l'argent disponible (3 457 866 951 FCFA). Sur cinq ans, il manque 7 531 335 924 FCFA (diapositive 179).")

# D162 : lecture des comptes tels quels
d.rep(162, "Figure 27 : Actif par délai de mobilisation et passif par échéance",
      "Figure 27 : Biens selon leur délai de vente, dettes selon leur date de paiement")
d.rep(162, "L'actif à moins d'un an ne couvre pas", "Dans les comptes tels quels, l'actif à moins d'un an ne couvre pas")
d.soit(162, "le Fonds a entre 35 et 69 FCFA", ", soit 35 % à 69 % des paiements")

# D167 : lectures A, B et C
d.rep(167, "La lecture B est la plus juste",
      "La lecture B, qui ramène l'argent mis de côté pour les soins et les décès au besoin estimé par Adenka, est la plus juste")
d.rep(167, "Le Fonds n'a que 69,42 FCFA (lecture B, la plus juste)", "Le Fonds n'a que 69,42 FCFA, soit 69,42 % (lecture B, la plus juste)", mini=0)
d.soit(167, "n'a que 69,42 FCFA", ", soit 69,42 %", [SY])
d.soit(167, "66,96 FCFA pour 100 FCFA à payer", ", soit 66,96 % des paiements de l'année")

# D33 : surplus 2025 corrigé
d.ajoute(33, "TextBox 32", 3, 3,
         "En 2025, le Fonds a payé 26 440 750 FCFA d'intérêts du plan d'épargne sans les compter en dépense, et la part des "
         "dépenses de la clinique inscrite comme somme due (compte 47111105) a augmenté de 196 285 018 FCFA. Comptés en dépenses, "
         "ces montants ramènent le surplus de 437 899 911 FCFA à 215 174 143 FCFA : pour 100 FCFA de surplus affiché, 49,14 FCFA "
         "seulement sont réels, soit 49,14 % du surplus.")
d.rep(33, "Traitements Adenka", "Traitements Adenka. Intérêts 2025 : grand livre 2025, compte 41195100 (remboursements égaux au "
      "versé × 1,035, calcul Adenka). Compte 47111105 : « BALANCE FPPN 2024.xlsx » et « BALANCE FPPN 2025.xlsx »")

# D92 et D173 : « inscrit à l'actif »
d.rep(92, "une partie de la santé des policiers était financée par le FPPN mais inscrite à l'actif (clinique, hôpital, État), ce "
          "qui gonflait les excédents et l'actif.",
      "une partie des soins des policiers était payée par le Fonds mais inscrite comme une somme que la clinique, l'hôpital ou "
      "l'État lui devaient, au lieu d'une dépense. Ces soins n'apparaissaient pas dans les dépenses de soins, ce qui gonflait "
      "les excédents et les biens du Fonds.")
d.rep(173, "une partie des soins de la clinique est payée par le FPPN et inscrite à l'actif, nette des prestations facturées, au "
           "lieu d'être en charges.",
      "une partie des dépenses de la clinique est payée par le FPPN et inscrite comme une somme que la clinique lui doit, puis "
      "déduite de ses factures de soins, au lieu d'être comptée en dépense. Le reste non déduit (1 280 300 860 FCFA) fait "
      "paraître les biens plus importants qu'ils ne sont.")

# D172 : réserve de réévaluation
d.reecrit(172, "TextBox 32", 0, "Une réserve de réévaluation",
          "Quand un bien prend de la valeur, la comptabilité peut relever sa valeur et inscrire le gain dans une réserve de "
          "réévaluation. Ce gain doit être la différence entre la nouvelle valeur et l'ancienne. Ici, la réserve égale la valeur "
          "entière du siège fixée par l'expert en 2019 (2 899 400 000 FCFA). [[Si le prix d'achat du terrain figure encore au "
          "compte des terrains (22200000), le terrain est compté deux fois, dans les biens comme dans les fonds propres.]]")
d.rep(172, "Le FPPN doit produire l'écriture d'origine et le coût d'acquisition du terrain.",
      "L'écriture est antérieure à septembre 2020, date du premier grand livre reçu : le compte 10620000 porte déjà "
      "2 899 400 000 FCFA à l'ouverture. [[Le FPPN doit produire l'écriture d'origine et le coût d'achat du terrain.]]")

# D139 : compte de liaison
d.rep(139, "Le FPPN et le Grenier doivent rapprocher leur compte de liaison chaque mois",
      "Le Fonds et le Grenier tiennent chacun un compte de ce qu'ils se doivent l'un à l'autre (compte de liaison). Les deux "
      "devraient donner le même montant. Ils ne concordent qu'en 2025 : l'écart est de 67 280 924 FCFA en 2022, 87 274 688 FCFA "
      "en 2023, 53 178 248 FCFA en 2024 et 70 191 500 FCFA au 31/08/2026. Le FPPN et le Grenier doivent rapprocher ce compte "
      "chaque mois")

# D128 : PSAP
d.reecrit(128, "TextBox 33", 0, "Les factures enregistrées",
          "Fin 2025, les factures de cliniques reçues et non payées font 1,64 mois de soins, l'argent mis de côté pour la maladie "
          "(PSAP) 19,42 mois. Pour 100 FCFA de factures à payer, le Fonds a mis de côté 1 184 FCFA, soit 1 184 % de ces factures "
          "(11,84 fois leur montant).")
d.rep(128, "Figure 25 : PSAP maladie et dettes envers les cliniques",
      "Figure 25 : Argent mis de côté pour la maladie (PSAP) et factures à payer")

# D125 : provision
d.rep(125, "La provision n'est pas recalculée.",
      "La provision (argent mis de côté pour payer les soins déjà faits dont la facture n'est pas encore arrivée ou pas encore "
      "payée) n'est pas recalculée.")

# D131 : factures tardives
d.rep(131, "Les factures en retard représentent de 1 215 600 FCFA à 2 748 440 FCFA selon la clôture.",
      "Peu de factures d'une année arrivent l'année suivante : de 1 215 600 FCFA à 2 748 440 FCFA selon la clôture.")
d.reecrit(131, "TextBox 32", 2, "La cadence",
          "Ces factures en retard (sinistres tardifs) sont si faibles que l'argent à mettre de côté pour elles se compte en "
          "millions de FCFA, pas en milliards.")
d.rep(131, "Au moins 0,89 FCFA sur 100 FCFA de factures arrive l'année suivante.",
      "Au moins 0,89 FCFA sur 100 FCFA de factures, soit 0,89 %, arrive l'année suivante.")

# D86 : provision mathématique
d.reecrit(86, "TextBox 33", 0, "Le taux de sinistralité",
          "Le taux de sinistralité ne mesure bien que les risques courts. Pour une promesse longue, comme rendre l'épargne ou "
          "payer des soins à vie, il faut calculer aujourd'hui la somme à avoir pour tenir la promesse plus tard (provision "
          "mathématique). Ce calcul tient compte de la durée de vie probable (table de mortalité) et de ce que l'argent placé "
          "rapporte d'ici là (taux d'intérêt technique) (règlement d'exécution n° 004/2022, art. 11).")

# D83 : jargon et mention interne
d.rep(83, "Les quotients écrits donnent", "Les divisions écrites dans la note donnent")
d.rep(83, "le registre des formules (onglet 11_formules) la reprend", "le tableau des formules en annexe la reprend")

# D24 : statuts
d.reecrit(24, "TextBox 32", 0, "Les statuts imposent",
          "Les statuts imposent eux-mêmes de calculer l'argent à mettre de côté (provisions) et de verser une part du surplus à "
          "la réserve légale. Les écarts relevés en parties 4.7 et 4.8 sont donc aussi des écarts aux propres règles du FPPN.")
d.reecrit(24, "TextBox 32", 1, "Les délais de forclusion",
          "Le règlement intérieur fixe les délais au-delà desquels un policier ne peut plus réclamer un paiement (forclusion). "
          "Ces délais disent combien de temps des factures peuvent encore arriver. Ils serviront à mesurer l'argent à mettre de "
          "côté pour elles dès réception de la liste des soins demandée au FPPN.")

# D113, D170 : jargon
d.rep(113, "Ces prêts doivent être provisionnés.",
      "Leur valeur doit être baissée dans les comptes du Grenier (dépréciation), parce qu'une partie ne sera pas remboursée.")
d.rep(170, "avec le barème de la BCEAO des SFD",
      "avec le barème de la banque centrale (BCEAO) pour les établissements de microfinance (SFD)")

# D143 et D155 : retraite du personnel
d.ajoute(143, "TextBox 32", 1, 1,
         "Le Fonds calcule ce qu'il paierait si chaque salarié partait le 31 décembre. Adenka calcule ce qu'il paiera au "
         "départ en retraite : les salaires montent d'ici là, certains salariés partent avant et ne toucheront pas l'indemnité "
         "(rotation), et une somme payée dans vingt ans vaut moins aujourd'hui (actualisation).")
d.reecrit(155, SY, 0, "72,1 millions",
          "Sans départs avant la retraite, il faudrait 72,1 millions de FCFA, plus que les 43,7 millions des comptes")

# D91 : montant exact de la perte maladie
d.reecrit(91, SY, 0, "En 2025, 103,90",
          "En 2025, 103,90 FCFA de soins et de frais pour 100 FCFA de cotisations, soit 103,90 % : la maladie perd 93 891 850 FCFA",
          tout_bleu=False)

# ---------------------------------------------------------------- 2. Taux sans « soit X % »
S = d.soit
S(4, "16,28 FCFA seulement sont à la banque", ", soit 16,28 % de l'épargne")
S(4, "103,90 FCFA sont sortis en soins et en frais", ", soit 103,90 % des cotisations maladie")
d.rep(4, "après correction d'une somme de 2025 oubliée", "après correction de factures de 2025 comptées deux fois")
S(5, "le Fonds doit 71,80 FCFA d'épargne aux policiers", ", soit 71,80 % de ses biens")
S(6, "le Fonds a 69,42 FCFA sous la main", ", soit 69,42 % des paiements")
S(6, "le Fonds a gagné 4,90 FCFA par an de 2021 à 2025", ", soit 4,90 % des sommes placées")
d.rep(6, "3,5 FCFA pour 100 FCFA versés, une seule fois, soit 183 023 590 FCFA",
      "3,5 FCFA pour 100 FCFA versés, une seule fois, soit 3,5 % du montant versé, ce qui fait 183 023 590 FCFA", mini=0)
S(7, "16,28 FCFA seulement sont restés à la banque", ", soit 16,28 % de l'épargne")
S(7, "177,05 FCFA disponibles pour 100 FCFA à payer", ", soit 177,05 % des paiements de l'année")
S(8, "103,90 FCFA pour 100 FCFA reçus", ", soit 103,90 % des cotisations", [ST])
S(8, "19,16 FCFA pour 100 FCFA reçus", ", soit 19,16 % des cotisations")
S(8, "103,90 FCFA sont sortis en soins et en frais", ", soit 103,90 % des cotisations maladie")
d.rep(8, "64,54 FCFA pour 100 FCFA, sous la norme.",
      "64,54 FCFA pour 100 FCFA de cotisations, soit 64,54 % des cotisations, sous la norme de 70 % à 80 %.")
for n in (9, 90):
    S(n, "27,52 FCFA sur 100 FCFA de dépenses maladie", ", soit 27,52 % des dépenses maladie")
S(13, "la limite de 20 FCFA pour 100 FCFA de cotisations", ", soit 20 % des cotisations")
S(13, "103,90 FCFA de soins et de frais pour 100 FCFA reçus", ", soit 103,90 % des cotisations")
S(15, "76 à 85 FCFA pour 100 FCFA dus", ", soit 76 % à 85 % des sommes dues")
d.rep(17, "le Fonds a 69,42 FCFA. Sans l'épargne, il a 177,05 FCFA.",
      "le Fonds a 69,42 FCFA, soit 69,42 % des paiements. Sans l'épargne, il a 177,05 FCFA, soit 177,05 %.")
d.rep(17, "il a 69,42 FCFA. Sans l'épargne, il en a 177,05", "il a 69,42 FCFA, soit 69,42 %. Sans l'épargne, il en a 177,05, soit 177,05 %")
S(19, "couvrent 22 FCFA pour 100 FCFA de dettes", ", soit 22 % des dettes")
S(19, "Les frais prennent 36,34 FCFA pour 100 FCFA de cotisations", ", soit 36,34 % des cotisations")
d.rep(19, "36,34 FCFA pour 100 FCFA de cotisations pour une limite de 20",
      "36,34 FCFA pour 100 FCFA de cotisations, soit 36,34 %, pour une limite de 20 %")
S(55, "97 à 100 FCFA pour 100 FCFA demandés", ", soit 97 % à 100 % des sommes demandées")
S(56, "76,96 FCFA sont retenus pour 100 FCFA demandés", ", soit 76,96 %")
S(56, "En juillet 2024, 150,62 FCFA", ", soit 150,62 %")
S(68, "5 FCFA pour 100 FCFA de salaire de base", ", soit 5 % du salaire de base")
d.rep(70, "73,53 FCFA pour 100 FCFA reçus selon les comptes, 68,75 FCFA selon la Solde",
      "73,53 FCFA pour 100 FCFA reçus selon les comptes, soit 73,53 %, et 68,75 FCFA selon la Solde, soit 68,75 %")
S(78, "73,53 FCFA de soins pour 100 FCFA reçus", ", soit 73,53 % des cotisations", [ST])
d.rep(79, "6,00 FCFA pour 100 FCFA en 2021, 73,53 en 2025 et 82,31 à fin août 2026, après correction",
      "6,00 FCFA pour 100 FCFA en 2021, soit 6,00 %, puis 73,53 en 2025 et 82,31 à fin août 2026 après correction, soit 73,53 % "
      "et 82,31 %")
for n in (82, 218):
    S(n, "57,48 FCFA de soins et d'aides pour 100 FCFA de cotisations", ", soit 57,48 % des cotisations", [ST])
d.rep(89, "le Fonds rendait moins de 10 FCFA de soins pour 100 FCFA reçus",
      "moins de 10 FCFA de soins pour 100 FCFA reçus, soit moins de 10 % des cotisations", [ST])
S(89, "soins et frais sous 100 FCFA pour 100 FCFA", ", soit moins de 100 % des cotisations")
d.rep(89, "pour 100 FCFA de cotisations (20,77 FCFA en 2021)",
      "pour 100 FCFA de cotisations, soit moins de 10 % des cotisations (20,77 FCFA en 2021, soit 20,77 %)")
S(96, "36,34 FCFA de frais pour 100 FCFA de cotisations", ", soit 36,34 % des cotisations", [ST])
S(99, "le Fonds en garde 16,28 en banque", ", soit 16,28 % de l'épargne")
d.rep(103, "16,28 FCFA en banque pour 100 FCFA d'épargne : un retrait de 15,46 % vide la caisse disponible",
      "16,28 FCFA en banque pour 100 FCFA d'épargne, soit 16,28 % : un retrait de 15,46 % de l'épargne vide la caisse")
S(106, "Le Fonds rend 3,5 FCFA pour 100 FCFA versés, une seule fois", ", soit 3,5 % du montant versé", [ST])
S(115, "76 à 85 FCFA pour 100 FCFA dus", ", soit 76 % à 85 % des sommes dues")
S(145, "33 salariés sur 100 en août 2026", ", soit 33 %")
S(171, "122,63 FCFA de biens pour 100 FCFA de dettes", ", soit 122,63 % des dettes")
d.rep(178, "30,30 FCFA sur 100 FCFA du plan d'épargne sortent, contre 8,29 FCFA pour l'entraide",
      "30,30 FCFA sur 100 FCFA du plan d'épargne sortent, soit 30,30 %, contre 8,29 FCFA pour l'entraide, soit 8,29 %")
S(180, "69,42 FCFA pour 100 FCFA à payer", ", soit 69,42 % des paiements")
for n in (187, 278):
    S(n, "5,53 FCFA pour 100 FCFA d'épargne", ", soit 5,53 % de l'épargne")
d.rep(190, "4,90 FCFA pour 100 FCFA par an", "4,90 FCFA pour 100 FCFA placés par an, soit 4,90 % des sommes placées")
for n in (212, 213):
    d.rep(n, "36,34 FCFA de frais pour 100 FCFA de cotisations, pour une limite de 20 :",
          "36,34 FCFA de frais pour 100 FCFA de cotisations, soit 36,34 %, pour une limite de 20 % :")
S(221, "22 FCFA pour 100 FCFA de dettes", ", soit 22 % des dettes", [ST])
S(222, "13 FCFA pour 100 FCFA de dettes à court terme", ", soit 13 % des dettes à court terme")
S(224, "12,81 FCFA pour 100 FCFA de cotisations", ", soit 12,81 % des cotisations", [ST])
S(225, "36,34 FCFA pour 100 FCFA de cotisations", ", soit 36,34 % des cotisations", [ST])
S(225, "63 FCFA de fonctionnement pour 100 FCFA de prestations", ", soit 63 % des prestations")
d.rep(258, "81,51 personnes vivent jusqu'à 60 ans avec CIMA H et 89,58 avec CIMA F",
      "81,51 personnes vivent jusqu'à 60 ans avec CIMA H, soit 81,51 %, et 89,58 avec CIMA F, soit 89,58 %")

# ---------------------------------------------------------------- 3. Extourne : même mot partout
for n in range(1, len(d.prs.slides) + 1):
    if n == 80:
        continue
    for anc, neu in (("après extourne de la provision de 2025", "après correction des factures de 2025 comptées deux fois"),
                     ("extourne de la provision de 2025 ajoutée par Adenka",
                      "correction des factures de 2025 comptées deux fois ajoutée par Adenka"),
                     ("après extourne", "après correction"), ("extourne ajoutée par Adenka", "correction ajoutée par Adenka"),
                     ("extourne comprise", "correction comprise"),
                     ("Effet de l'extourne oubliée sur les chiffres de 2026", "Effet de la correction des factures comptées deux fois")):
        d.rep(n, anc, neu, mini=0)

# ---------------------------------------------------------------- 4. Richesse propre -> fonds propres
RP = [(7, "la richesse propre du Fonds devient négative", "les fonds propres du Fonds deviennent négatifs"),
      (10, "Richesse propre face", "Fonds propres face"),
      (16, "La richesse propre affichée n'est pas fiable.", "Les fonds propres affichés ne sont pas fiables."),
      (18, "Sa richesse propre est négative.", "Ses fonds propres sont négatifs."),
      (18, "La richesse propre est négative à chaque fin d'année.", "Les fonds propres sont négatifs à chaque fin d'année."),
      (18, "richesse propre négative", "fonds propres négatifs"), (192, "richesse propre négative", "fonds propres négatifs"),
      (19, "sa richesse propre,", "ses fonds propres,"),
      (159, "Effet sur la richesse propre du Fonds", "Effet sur les fonds propres du Fonds"),
      (159, "fausse la richesse propre du Fonds", "fausse les fonds propres du Fonds"),
      (174, "la richesse propre du Fonds tombe", "les fonds propres du Fonds tombent"),
      (183, "Richesse propre et placements", "Fonds propres et placements"),
      (184, "la richesse propre du Fonds est de", "les fonds propres du Fonds sont de"),
      (185, "La richesse propre du Fonds est négative, mais son signe", "Les fonds propres du Fonds sont négatifs, mais leur signe"),
      (185, "Richesse propre du Fonds", "Fonds propres du Fonds"),
      (186, "Richesse propre une fois", "Fonds propres une fois"),
      (186, "la richesse propre est négative", "les fonds propres sont négatifs"),
      (199, "sa richesse propre", "ses fonds propres"),
      (205, "richesse propre et argent", "fonds propres et argent"),
      (208, "ni richesse propre positive", "ni fonds propres positifs"),
      (215, "et richesse propre", "et fonds propres"),
      (216, "la richesse propre et les frais", "les fonds propres et les frais"),
      (224, "sert à refaire une richesse propre négative", "sert à combler des fonds propres négatifs"),
      (242, "norme de richesse propre", "norme de fonds propres"), (243, "norme de richesse propre", "norme de fonds propres"),
      (314, "Effet sur la richesse propre", "Effet sur les fonds propres"),
      (315, "Richesse propre corrigée", "Fonds propres corrigés")]
for n, a, b in RP:
    d.rep(n, a, b)
for lieu, _, p in list(d.paras(315)):
    if texte_para(p) == "Richesse propre":
        modifie(p, "Fonds propres", False, lieu)

# ---------------------------------------------------------------- 5. Termes expliqués à la première apparition
CORPS_EXCLUS = (TI, ST, SY, "Adenka_Source")
DEBUTS_EXCLUS = ("Valable pour", "Sigles", "Surligné", "Source", "Extrait", "Tableau ", "Figure ", "En millions", "En %")
ADJ = ("mathématique", "maladie", "décès", "obsèques", "technique", "techniques", "figée", "figées")


def premiere(n, motif, deja):
    """Première occurrence utilisable : texte courant d'abord, tableaux ensuite. Jamais dans une parenthèse ni dans un
    intitulé court (moins de huit mots sans point final) ni sur une diapositive sans titre (séparateur, annexe)."""
    noms = {sh.name for sh in d.slide(n).shapes}
    if TI not in noms:
        return None
    tout = " ".join(texte_para(p) for _, _, p in d.paras(n))
    if deja in tout:
        return None
    for tables in (False, True):
        for lieu, nom, p in d.paras(n):
            t = texte_para(p)
            if ("[" in lieu) != tables or nom in CORPS_EXCLUS or t.startswith(DEBUTS_EXCLUS):
                continue
            if not tables and len(t.split()) < 8 and not t.rstrip().endswith("."):
                continue
            if tables and (" / " in t or "=" in t):
                continue
            for m in motif.finditer(t):
                if t[:m.start()].count("(") > t[:m.start()].count(")"):
                    continue
                return lieu, p, t, m
    return None


PROV = re.compile(r"\b([Pp]rovisions?)\b(?! \()")
FP = re.compile(r"\b[Ff]onds propres\b(?! \()")
SANS_PROV = {34, 43, 146, 269}
SANS_FP = {34, 43, 159, 184, 221, 239, 246}
for n in range(4, 282):
    r = None if n in SANS_PROV else premiere(n, PROV, "(argent mis de côté")
    if r and "mis de côté" not in r[2] and "mettre de côté" not in r[2]:
        lieu, p, t, m = r
        fin = m.end()
        reste = t[fin:]
        mm = re.match(r" (" + "|".join(ADJ) + r")\b", reste)
        if mm:
            fin += mm.end()
        modifie(p, t[:fin] + " (argent mis de côté)" + t[fin:], False, lieu)
        d.modifiees.add(n)
    r = None if n in SANS_FP else premiere(n, FP, "(ce que le Fonds possède")
    if r:
        lieu, p, t, m = r
        fin, suite = m.end(), t[m.end():]
        if suite.startswith(" du Fonds"):
            suite = suite[len(" du Fonds"):]
        modifie(p, t[:fin] + " (ce que le Fonds possède moins ce qu'il doit)" + suite, False, lieu)
        d.modifiees.add(n)

# ---------------------------------------------------------------- 6. Renvois dans les tableaux : « Diapo » et « D42 »
RENVOI = re.compile(r"\b[Dd]iapositives? (\d+)(?:( et | à |, )(\d+))?")
for n in range(1, len(d.prs.slides) + 1):
    for sh in d.slide(n).shapes:
        if not sh.has_table:
            continue
        for ri, row in enumerate(sh.table.rows):
            for ci, c in enumerate(row.cells):
                for i, pa in enumerate(c.text_frame.paragraphs):
                    t = texte_para(pa._p) if pa._p.findall("{http://schemas.openxmlformats.org/drawingml/2006/main}r") else ""
                    if not t:
                        continue
                    neu = t
                    if ri == 0 and t.strip() in ("Diapo.", "Diapositive", "Diapositives"):
                        neu = "Diapo"
                    else:
                        neu = RENVOI.sub(lambda m: "D" + m.group(1) + ((m.group(2) + "D" + m.group(3)) if m.group(3) else ""), t)
                    if neu != t:
                        modifie(pa._p, neu, False, f"D{n} {sh.name}[{ri},{ci}]#{i}")
                        d.modifiees.add(n)

# ---------------------------------------------------------------- 7. Mention du surlignage jaune en bas de page
NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
for n in sorted(d.modifiees):
    s = d.slide(n)
    xml = "".join(sh._element.xml for sh in s.shapes)
    if 'highlight' not in xml or "Surligné en jaune" in xml:
        continue
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.startswith("Sigles"):
            p = sh.text_frame.paragraphs[0]._p
            modifie(p, "Surligné en jaune : affirmation à confirmer par une pièce ou une réponse écrite du FPPN. " + texte_para(p),
                    False, f"D{n} {sh.name}#0")
            break
    else:
        print(f"D{n} : surlignage sans mention en bas de page")

# ---------------------------------------------------------------- 8. Taille du texte là où les ajouts débordent
TAILLE = {(33, "TextBox 32"): 12.5, (139, "TextBox 32"): 11.5, (143, "TextBox 32"): 11, (169, "TextBox 32"): 10.5,
          (172, "TextBox 32"): 11}
for (n, nom), pt in TAILLE.items():
    for el in d.forme(n, nom)._element.iter(NS + "rPr", NS + "endParaRPr"):
        el.set("sz", str(int(pt * 100)))

d.prs.save(V7)

# ---------------------------------------------------------------- journal
wb = Workbook()
ws = wb.active
ws.title = "Corrections_v7"
ws.append(["Diapo", "Endroit", "Texte v6", "Texte v7"])
for c in ws[1]:
    c.font, c.fill = Font(name="Helvetica", bold=True, color="FFFFFF"), PatternFill("solid", fgColor="060606")
for lieu, a, b in JOURNAL:
    ws.append([lieu.split()[0], lieu.split(" ", 1)[1], a, b])
for col, w in zip("ABCD", (7, 26, 80, 80)):
    ws.column_dimensions[col].width = w
for row in ws.iter_rows(min_row=2):
    for c in row:
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.font = Font(name="Helvetica", size=9)
wb.save(JR)
print(len(JOURNAL), "modifications sur", len(d.modifiees), "diapositives")
os.makedirs("donnees_intermediaires/v7", exist_ok=True)
open("donnees_intermediaires/v7/diapos_modifiees.txt", "w").write(" ".join(map(str, sorted(d.modifiees))))
