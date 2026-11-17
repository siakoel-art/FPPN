# -*- coding: utf-8 -*-
"""
FPPN - Diagnostic de clarté du rapport v6 (16/11/2026), diapositive par diapositive (17/11/2026).
Onglets : 1 constats à clarifier (rédaction proposée), 2 taux sans « soit X % », 3 renvois à écrire « D42 »,
4 questions de forme. Rien n'est modifié dans le rapport : liste soumise à validation.
Utilisation : python diagnostic_v6_fppn.py <v6_texte.json> <sortie.xlsx>
"""
import json
import re
import sys

sys.path.insert(0, __import__('os').path.dirname(__file__))
from diagnostic_v6_taux import proposition
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

SRC, SORTIE = sys.argv[1:3]
L = json.load(open(SRC, encoding="utf-8"))
I, E, J = "Incohérence", "Explication à ajouter", "Mot technique à expliquer"

# (diapo, nature, ce qui est écrit, ce qui ne va pas, rédaction proposée)
C = [
 (9, E, "« S'y ajoutent 1 280 300 860 FCFA de dépenses de la clinique comptées comme des biens. Elles font paraître le Fonds plus riche qu'il n'est. »",
  "Le lecteur ne sait pas pourquoi une dépense devient un bien, ni pourquoi cela enrichit le Fonds en apparence.",
  "S'y ajoutent 1 280 300 860 FCFA de dépenses de la clinique payées par le Fonds et inscrites comme une somme que la clinique lui doit (compte 47111105), au lieu d'une dépense. Elles ne deviennent des dépenses de soins que lorsqu'elles sont déduites des factures de la clinique. Chaque année, le Fonds paie plus qu'il ne déduit : le reste est passé de 469 980 456 FCFA fin 2022 à 1 280 300 860 FCFA au 31/08/2026. Tant qu'il n'est pas déduit, les biens et le résultat du Fonds paraissent plus élevés qu'ils ne sont. [[La règle de déduction est à confirmer par le service comptable.]]"),
 (80, E, "Rubrique « Somme de 2025 oubliée dans les comptes ». Sous-titre « Une somme mise de côté fin 2025 et oubliée gonfle les dépenses de 2026 ». Texte : « Provision du 31/12/2025… Aucune écriture n'annule la provision… extourne ».",
  "« Oubliée » est inexact : c'est l'annulation qui manque. « Provision » et « extourne » ne sont pas expliqués. Le lecteur ne voit pas pourquoi ce sont les mêmes soins.",
  "Rubrique : « Factures de 2025 comptées deux fois ». Sous-titre : « Des factures de 2025 sont comptées une seconde fois dans les dépenses de 2026 ». Texte : « Le 31/12/2025, le Fonds attend les factures de sa clinique pour octobre, novembre et décembre 2025. Il inscrit donc 255 000 000 FCFA en dépense de 2025 et en dette envers la clinique (écritures OD 805, 806 et 807, « CMP PROV FACT 10/2025 » à « 12/2025 »). Le 31/01/2026, les vraies factures de ces trois mois arrivent : 240 471 119 FCFA (pièces SIN 462, 552 et 619, « OCT 25 », « NOVEMBRE 2025 », « DECEMBRE »). Le Fonds les inscrit en dépense de 2026 sans annuler les 255 000 000 FCFA de fin 2025. Les mêmes soins sont donc comptés deux fois, et la dette envers la clinique aussi. Le rapport retire ces 255 000 000 FCFA de 2026 : la maladie dépense alors 82,31 FCFA pour 100 FCFA de cotisations maladie, soit 82,31 % des cotisations, au lieu de 97,40. » [[À confirmer par le lettrage du compte 40110000.]] Tableau : « Après extourne » devient « Après correction »."),
 (105, I, "Sous-titre : « l'épargne due paraît plus petite ». Texte : « Le grand livre ne dit pas si ces règlements comprennent des intérêts calculés hors comptabilité. »",
  "Le texte contredit le constat du 13/11 (les intérêts du plan d'épargne sont prouvés) et le mécanisme « paraît plus petite » n'est pas expliqué.",
  "Quand un policier quitte le plan d'épargne, le Fonds lui rend ce qu'il a versé plus 3,5 %. Ces 3,5 % sont pris sur le compte d'épargne, comme s'ils faisaient partie des sommes versées, au lieu d'être comptés en dépense. De 09/2020 à 08/2026, ils font 198 040 675 FCFA. Le compte d'épargne affiche donc 198 040 675 FCFA de moins que ce que les autres policiers ont réellement versé, et les dépenses du Fonds sont trop basses du même montant. Pour l'entraide de fin de carrière, les règlements sont des montants ronds : [[on ne sait pas encore s'ils contiennent des intérêts]]."),
 (108, I, "« Selon le taux retenu pour l'EFC, les intérêts dus depuis le 01/01/2020 sont compris entre 2 541 799 542 FCFA et 3 278 913 238 FCFA… soit 16,1 fois les produits financiers du FPPN de 2021 à 2025 ».",
  "Le texte reprend l'ancienne fourchette et ignore la règle des entretiens donnée en sous-titre. « 16,1 fois » compare un total de six ans aux produits de cinq ans (un stock à des flux).",
  "Le montant dû dépend de la règle. Selon la règle appliquée par le Fonds (3,5 % du montant versé, une seule fois), il doit 183 023 590 FCFA sur le plan d'épargne au 31/08/2026, soit 3,5 % des 5 229 245 425 FCFA versés. Selon le manuel (3,5 % par an sur le plan d'épargne, 5,5 % par an sur l'entraide), il devrait 3 278 913 238 FCFA depuis le 01/01/2020. Aucun de ces montants n'est inscrit dans les comptes. [[La règle de l'entraide reste à recevoir.]] Supprimer « 16,1 fois »."),
 (169, I, "« Seule la baisse des provisions maladie et décès allège le passif : elle est trop faible pour changer la conclusion… Le FPPN n'est pas adossé, même dans l'hypothèse la plus favorable. »",
  "C'est l'ancien raisonnement « borne favorable », corrigé depuis : retirer l'argent mis de côté en trop fait passer de 35,32 à 69,42 FCFA pour 100 FCFA à payer. La phrase contredit le sous-titre et D162.",
  "Les 4 829 298 290 FCFA mis de côté pour la maladie et les obsèques sont comptés comme des dettes à payer dans l'année, alors que le besoin estimé par Adenka n'est que d'au moins 19 581 392 FCFA. Les ramener au besoin change beaucoup la mesure : pour 100 FCFA à payer dans l'année, le Fonds dispose de 69,42 FCFA au lieu de 35,32, soit 69,42 % des paiements au lieu de 35,32 %. Le manque reste réel, parce que l'épargne à rendre dans l'année dépasse l'argent disponible."),
 (33, E, "Sous-titre : « Ce surplus oublie les intérêts dus sur l'épargne et des dépenses de la clinique comptées comme des biens ».",
  "L'effet n'est pas chiffré : le lecteur ne sait pas de combien le surplus est surestimé.",
  "Ajouter : « En 2025, le Fonds a payé 26 440 750 FCFA d'intérêts du plan d'épargne sans les compter en dépense, et la part des dépenses de la clinique inscrite comme somme due a augmenté de 196 285 018 FCFA. Comptés en dépenses, ces montants ramènent le surplus de 437 899 911 FCFA à 215 174 143 FCFA : pour 100 FCFA de surplus affiché, 49,14 FCFA seulement sont réels, soit 49,14 % du surplus. »"),
 (92, J, "« une partie de la santé des policiers était financée par le FPPN mais inscrite à l'actif (clinique, hôpital, État) »",
  "« Inscrite à l'actif » n'est pas compris d'un dirigeant non comptable.",
  "« une partie des soins des policiers était payée par le Fonds mais inscrite comme une somme que la clinique, l'hôpital ou l'État lui devaient, au lieu d'une dépense. Ces soins n'apparaissaient donc pas dans les dépenses de soins, ce qui faisait paraître le taux maladie très bas (6,47 FCFA pour 100 FCFA de cotisations en 2022, soit 6,47 %, au lieu de 32,86). »"),
 (173, J, "« inscrite à l'actif, nette des prestations facturées, au lieu d'être en charges »",
  "Trois termes comptables dans la même phrase.",
  "Reprendre l'explication de D9 : dépenses payées par le Fonds, inscrites comme une somme due par la clinique, puis déduites de ses factures de soins. Le reste non déduit fait paraître les biens plus élevés."),
 (172, E, "« Une réserve de réévaluation doit égaler l'écart entre la valeur réévaluée et la valeur comptable. Ici, elle égale la valeur entière… »",
  "La notion de réévaluation n'est pas expliquée.",
  "Quand un bien prend de la valeur, la comptabilité peut relever sa valeur et inscrire le gain dans une réserve de réévaluation. Ce gain doit être la différence entre la nouvelle valeur et l'ancienne. Ici, la réserve égale la valeur entière du siège fixée par l'expert en 2019 (2 899 400 000 FCFA). Si le prix d'achat du terrain figure encore au compte des terrains (22200000), le terrain est compté deux fois, dans les biens comme dans la richesse propre. [[Écriture d'origine à recevoir.]]"),
 (139, J, "« Le FPPN et le Grenier doivent rapprocher leur compte de liaison chaque mois »",
  "« Compte de liaison » n'est pas expliqué.",
  "Le Fonds et le Grenier tiennent chacun un compte de ce qu'ils se doivent l'un à l'autre (compte de liaison). Les deux comptes devraient donner le même montant. Ils ne concordent qu'en 2025 : l'écart est de 67 280 924 FCFA en 2022, 87 274 688 FCFA en 2023, 53 178 248 FCFA en 2024 et 70 191 500 FCFA au 31/08/2026. Tant que l'écart n'est pas expliqué, on ne sait pas lequel des deux comptes est juste."),
 (128, J, "« Les factures enregistrées représentent 1,64 mois de prestations en 2025, la PSAP 19,42 mois. La PSAP vaut 11,84 fois les dettes envers les cliniques. »",
  "Sigle PSAP et rapport « 11,84 fois » sans forme « pour 100 FCFA ».",
  "« Fin 2025, les factures de cliniques reçues et non payées font 1,64 mois de soins, l'argent mis de côté pour la maladie 19,42 mois. Pour 100 FCFA de factures à payer, le Fonds a mis de côté 1 184 FCFA, soit 11,84 fois ces factures. »"),
 (125, J, "« La provision n'est pas recalculée… provision pour prestations à payer »",
  "La rubrique dit « argent mis de côté », le texte « provision » : le lien n'est pas fait.",
  "Ajouter une fois : « Cet argent mis de côté (la provision) sert à payer les soins déjà faits mais dont la facture n'est pas encore arrivée ou pas encore payée. »"),
 (131, J, "« La cadence de première année est faible : la provision pour sinistres tardifs en maladie est de l'ordre de quelques millions »",
  "« Cadence » et « sinistres tardifs » non expliqués, « quelques millions » non chiffré.",
  "« Peu de factures arrivent l'année suivante : de 1 215 600 FCFA à 2 748 440 FCFA selon l'année. L'argent à mettre de côté pour ces factures en retard se compte donc en millions de FCFA, pas en milliards. »"),
 (86, J, "« l'engagement se mesure par une provision mathématique, calculée avec une table de mortalité et un taux d'intérêt technique »",
  "Trois notions actuarielles sans explication.",
  "« Pour une promesse longue, comme rendre l'épargne ou payer des soins à vie, on calcule aujourd'hui la somme qu'il faut avoir pour tenir la promesse plus tard. Ce calcul tient compte de la durée de vie probable (table de mortalité) et de ce que l'argent placé rapporte d'ici là (taux d'intérêt). »"),
 (143, J, "« Hausse des salaires de 10,36 % par an et taux d'actualisation de 7,19 % »",
  "« Taux d'actualisation », « rotation », « droits acquis » non expliqués dans la série D142 à D155.",
  "Ajouter en D142 : « Le Fonds doit une indemnité à chaque salarié civil qui part en retraite. Lui calcule ce qu'il paierait si chacun partait le 31 décembre (droits déjà gagnés). Adenka calcule ce qu'il paiera vraiment au départ, en tenant compte des hausses de salaire, des salariés qui partent avant la retraite (rotation) et du fait qu'une somme payée dans vingt ans vaut moins aujourd'hui (actualisation). »"),
 (155, E, "« 72,1 millions de FCFA sans départs contre 43,7 millions dans les comptes : tout dépend du taux de départ »",
  "Le lecteur ne sait pas ce qu'est le « taux de départ » ni pourquoi il pèse autant.",
  "« Le calcul suppose que 9,79 salariés sur 100 quittent le Fonds chaque année avant la retraite, soit 9,79 % : ceux-là ne toucheront pas l'indemnité. Si personne ne part, le Fonds devra 72 102 465 FCFA, plus que les 43 666 678 FCFA inscrits. »"),
 (167, J, "« La lecture B est la plus juste… La lecture A… La lecture C… »",
  "Les trois lectures ne sont définies nulle part sur la diapositive.",
  "Ajouter en tête : « Lecture A : les comptes tels quels. Lecture B : l'argent mis de côté ramené au besoin estimé par Adenka. Lecture C : la lecture B plus les intérêts du plan d'épargne, s'ils étaient tous payés dans l'année. »"),
 (162, J, "Figure « Actif par délai de mobilisation et passif par échéance ». « L'écart vaut 3 304 869 546 FCFA… avant même toute sortie d'épargne. »",
  "« Délai de mobilisation » est un terme technique et l'écart cité correspond aux comptes tels quels (lecture A) sans le dire.",
  "Figure : « Biens classés selon le temps qu'il faut pour en faire de l'argent, dettes selon leur date de paiement ». Phrase : « Dans les comptes tels quels, … »"),
 (113, J, "« 126 prêts ont plus de 180 jours de retard. Ces prêts doivent être provisionnés. »", "« Provisionnés » non expliqué.",
  "« 126 prêts ont plus de 180 jours de retard : leur valeur doit être baissée dans les comptes du Grenier, parce qu'une partie ne sera pas remboursée. »"),
 (24, J, "« dotation de la réserve légale… délais de forclusion… cadences de la méthode CIMA… (PBC, item 47) »",
  "Quatre termes techniques et une mention interne (PBC).",
  "« Les statuts imposent eux-mêmes de calculer l'argent à mettre de côté et de verser une part du surplus à la réserve. Les écarts relevés plus loin sont donc aussi des écarts aux propres règles du Fonds. Les délais au-delà desquels un policier ne peut plus réclamer un remboursement serviront à mesurer les factures en retard, dès réception de la liste des soins demandée au Fonds. »"),
 (83, J, "« Les quotients écrits donnent 12,95 % et 99,23 %… le registre des formules (onglet 11_formules) la reprend »",
  "« Quotients » et mention d'un onglet du classeur interne.",
  "« Les calculs écrits dans la note donnent 12,95 % pour le décès et 99,23 % pour la maladie. » Retirer la mention de l'onglet."),
 (170, J, "« barème de la BCEAO des SFD »", "Sigles et jargon.", "« barème de la banque centrale pour les établissements de microfinance »"),
 (174, I, "Sous-titre « richesse propre », texte « fonds propres ».",
  "Deux mots pour la même chose dans la même diapositive (même écart en D184 à D186).",
  "Garder « richesse propre » et ajouter une fois : « (ce que le Fonds possède moins ce qu'il doit, appelé fonds propres en comptabilité) »."),
]

wb = Workbook()
noir, gris = "060606", "BBBCBC"
fh = Font(name="Helvetica", size=10, bold=True, color="FFFFFF")
f = Font(name="Helvetica", size=10)
bd = Border(*(Side(style="thin", color=gris),) * 4)


def feuille(ws, titre, entetes, lignes, larg):
    ws.append([titre])
    ws["A1"].font = Font(name="Helvetica", size=13, bold=True)
    ws.append([])
    ws.append(entetes)
    for c in ws[3]:
        c.font, c.fill, c.border = fh, PatternFill("solid", fgColor=noir), bd
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for l in lignes:
        ws.append(list(l))
        for c in ws[ws.max_row]:
            c.font, c.border, c.alignment = f, bd, Alignment(wrap_text=True, vertical="top")
    for i, w in enumerate(larg):
        ws.column_dimensions[chr(65 + i)].width = w
    ws.freeze_panes = "A4"


ws = wb.active
ws.title = "1_Constats"
feuille(ws, "Constats et passages à clarifier, rapport v6 du 16/11/2026 (diagnostic du 17/11/2026)",
        ["Diapo", "Nature", "Ce qui est écrit", "Ce qui ne va pas", "Rédaction proposée"], sorted(C, key=lambda x: x[0]), [7, 18, 45, 40, 90])

# 2. taux sans « soit X % »
TX = []
for d in L:
    T = "\n".join([d["titre"], d["sous_titre"]] + d["texte"] + [" | ".join(r) for t in d["tableaux"] for r in t])
    for phr in re.split(r"(?<=[.!])\s+|\n", T):
        if "pour 100 FCFA" in phr and not re.search(r"soit [\d ,]+ ?%|représente environ", phr):
            prop = proposition(d["n"], phr)
            TX.append((d["n"], phr.strip()[:300], prop[:400]))
ws2 = wb.create_sheet("2_Taux_sans_soit")
feuille(ws2, f"Taux écrits « X FCFA pour 100 FCFA » sans « soit X % » : {len(TX)} phrases (rédaction proposée)",
        ["Diapo", "Phrase actuelle", "Phrase proposée"], TX, [7, 80, 90])

# 3. renvois
RV = []
for d in L:
    for t in d["tableaux"]:
        if t and any(re.search(r"^(Diapo\.?|Diapositive|Détail)$", c.strip()) for c in t[0]):
            RV.append((d["n"], "en-tête : " + " | ".join(t[0])[:120], "en-tête de la colonne : « Diapo »"))
        for r in t:
            for c in r:
                for m in re.finditer(r"[Dd]iapositives? (\d+)", c):
                    RV.append((d["n"], c.strip()[:120], re.sub(r"[Dd]iapositives? (\d+)", r"D\1", c.strip())[:120]))
ws3 = wb.create_sheet("3_Renvois_D42")
feuille(ws3, f"Renvois dans les tableaux à écrire « D42 » : {len(RV)} cellules ou en-têtes", ["Diapo", "Actuel", "Proposé"], RV, [7, 70, 70])

# 4. questions
Q = [("Montants arrondis", "Plusieurs sous-titres de la v6 arrondissent (« 15,7 milliards », « 487,5 millions », « 2,87 milliards », « 1,52 milliard »). La règle demande des montants exacts en FCFA.", "Garder l'arrondi dans les sous-titres pour la lisibilité, avec le montant exact dans le texte, ou revenir partout au montant exact ?"),
     ("Argent mis de côté / provision", "La v6 dit « argent mis de côté » dans les rubriques et « provision » dans les textes et tableaux.", "Garder les deux, avec « (provision) » expliqué une fois par diapositive, ou remplacer « provision » partout ?"),
     ("Richesse propre / fonds propres", "Même dédoublement entre sous-titres et textes (D174, D184 à D186).", "Garder « richesse propre », avec « fonds propres » expliqué une fois ?"),
     ("Renvois dans le texte courant", "La règle « D42 » est appliquée aux tableaux.", "L'appliquer aussi dans les paragraphes (« voir D42 ») ou garder « diapositive 42 » ?")]
ws4 = wb.create_sheet("4_Questions")
feuille(ws4, "Questions de forme à trancher", ["Sujet", "Constat", "Question"], Q, [28, 80, 70])
wb.save(SORTIE)
print(len(C), "constats ;", len(TX), "taux ;", len(RV), "renvois")
