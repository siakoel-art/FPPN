"""Constats par activité du FPPN, présentation schématique pour le client (langage simple).

Structure : vue d'ensemble (carte des activités, matrice activités × types de constat, synthèse chiffrée), puis une
activité par diapositive, une ligne par constat lue de gauche à droite : constat et commentaire, recommandation,
complexité, priorité. L'arborescence des 14 activités sert de repère en haut de chaque diapositive. Les circuits des logigrammes du FPPN sont en annexe.
Numérotation et classement repris du résumé exécutif du 26/10/2026 (constats 1 à 28, sous-numéros pour les constats
propres à une activité). Chiffres du rapport actuariel et statistique du 23/10/2026.

Usage : python3 deck_constats_activites_fppn.py <skill adenka-powerpoint> <sortie.pptx>
"""
import os
import sys

ICI = os.path.dirname(os.path.abspath(__file__))
SKILL, SORTIE = sys.argv[1:3]
SRC = open(os.path.join(ICI, "rapport_v12_fppn.py"), encoding="utf-8").read().split("\n")


def lignes(a, b):
    return "\n".join(SRC[a - 1:b])


ENTETE = "\n".join([lignes(82, 87), lignes(90, 91), lignes(94, 94), lignes(96, 104), lignes(136, 187), lignes(238, 375), lignes(396, 411), lignes(602, 635)])
debut = next(i for i, l in enumerate(SRC, 1) if l.startswith("tmp = SORTIE.replace"))
fin = next(i for i, l in enumerate(SRC, 1) if l.startswith("finaliser(SORTIE)"))
MOTEUR = lignes(debut, fin + 1)
for avant, apres in [('"Audit Comptable et financier FPPN", "Text Placeholder 4": "23 octobre 2026"', '"Audit Comptable et financier FPPN", "Text Placeholder 4": "26 octobre 2026"'),
                     ('r.text = "Rapport actuariel et statistique"', 'r.text = "Les constats par activité"'),
                     ('entrees = [("Résumé exécutif", 3)]', 'entrees = []'),
                     ('        if x.get("triangle"):', '        if "activites" in x:\n            dessin_activites(sl, x["activites"])\n        if "carte" in x:\n            dessin_carte(sl)\n        if "matrice_types" in x:\n            dessin_matrice(sl)\n        if x.get("triangle"):')]:
    assert avant in MOTEUR, avant
    MOTEUR = MOTEUR.replace(avant, apres)

CONTENU = r'''
VALIDITE = {}
PBCM = {}
GAB = {"couverture": 1, "sommaire": 5, "separateur": 2, "contenu": 5, "annexe": 5, "merci": 110}
RAP = "Source : rapport actuariel et statistique du 23/10/2026"
CRIT = {"C": ("C00000", "Critique"), "E": ("ED7D31", "Élevé"), "M": ("FFC000", "Moyen"), "F": ("70AD47", "Faible")}
PRIO = {"P1": "C00000", "P2": "ED7D31", "P3": "FFC000", "P4": "70AD47"}
TYPES = [("D", "Données"), ("C", "Comptabilité"), ("A", "Actuariel"), ("R", "Conformité"), ("G", "Gouvernance")]


def I(num, crit, types, titre, texte, reco, cplx, prio):
    return {"num": num, "crit": crit, "types": types, "titre": titre, "texte": texte, "reco": reco, "cplx": cplx, "prio": prio}


ACT = {
 "Maladie": [
  I("19", "C", "A", "La garantie maladie perd de l'argent depuis 2025.", "Les soins consomment 73,53 % des cotisations maladie en 2025 et 97,40 % sur huit mois de 2026 (109,87 % frais compris en 2025). Le niveau bas de 2020 à 2022 venait d'une couverture limitée à l'hospitalisation et de factures impayées, pas d'une bonne santé.", "Recalculer la cotisation maladie sur le coût complet et revoir le barème et les conventions avec les cliniques.", "Élevée", "P2"),
  I("1.a", "E", "CA", "La provision maladie est figée depuis 2020.", "2 867 602 878 FCFA au bilan pour un besoin mesuré de 12 481 392 FCFA.", "Calculer la provision chaque année dans une note signée.", "Moyenne", "P1"),
  I("21", "E", "C", "Le paiement des cliniques ralentit.", "2,62 mois en 2025, 5,00 mois au 31/08/2026, pour une règle de 3 mois. 690 234 451 FCFA de factures en attente.", "Suivre chaque mois le délai de paiement et les factures en attente.", "Moyenne", "P2"),
  I("26.a", "M", "D", "La date des soins n'est saisie nulle part.", "Le délai réel de paiement et les soins non encore facturés ne peuvent pas être mesurés précisément.", "Saisir la date des soins et la date de dépôt de chaque facture.", "Moyenne", "P1"),
  I("19.a", "M", "R", "Le contrôle des factures après paiement n'est pas démontré.", "Aucun registre des contrôles n'a été transmis, alors que la réglementation demande d'en contrôler au moins 5 %.", "Contrôler après paiement au moins 5 % des factures.", "Moyenne", "P2"),
  I("26.b", "M", "D", "Le fichier des cartes de santé actives manque.", "Le nombre de personnes couvertes, ayants droit compris, n'est pas connu.", "Transmettre le fichier des cartes actives, ayants droit compris.", "Faible", "P2")],
 "Décès": [
  I("1.b", "E", "C", "La provision obsèques fait double emploi.", "1 961 695 412 FCFA figés depuis 2020, alors que la dette décès est déjà comptée au compte 41591000. Besoin mesuré : 7 100 000 FCFA de décès déclarés en retard.", "Faire décider le reclassement par les organes, après information de l'AIRMS.", "Moyenne", "P1"),
  I("2", "E", "C", "La dette envers les familles grossit.", "370 503 332 FCFA fin 2021, 489 419 669 FCFA au 31/08/2026. La date de dépôt des demandes n'est pas saisie.", "Dresser la liste nominative des dossiers, payer les dus, clôturer les prescrits.", "Moyenne", "P2"),
  I("3", "E", "CR", "Des allocations sont payées avec l'argent de l'épargne.", "102 allocations payées en 2024 depuis les comptes de l'épargne, pour 53 288 000 FCFA.", "Payer depuis les comptes des garanties et rembourser l'épargne.", "Faible", "P2"),
  I("20.a", "M", "C", "Le décès finance la maladie.", "La Solde prélève 1,23 % au lieu de 1,50 %. Le décès reste excédentaire : 449 151 155 FCFA en 2025.", "Appliquer la clé des statuts et suivre le résultat de chaque garantie.", "Faible", "P2"),
  I("27.a", "M", "D", "Des fichiers décès sont vides ou à valider.", "Fichiers de décembre 2024 vides (0 octet). Montants supérieurs à 1 000 000 FCFA à valider dans les fichiers retour.", "Transmettre les fichiers et valider ces montants.", "Faible", "P2"),
  I("28.a", "M", "C", "Les décès 2021 diffèrent selon la source.", "156 020 000 FCFA dans l'étude du FPPN, 401 856 666 FCFA en comptabilité.", "Expliquer l'écart.", "Faible", "P3"),
  I("26.c", "M", "D", "Le barème officiel et la règle du suicide ne sont pas documentés.", "Quatre décès valorisés à 0 FCFA en 2025 sans texte transmis.", "Transmettre le texte et le barème signé.", "Faible", "P3")],
 "Retraités (CGRAE)": [
  I("22", "E", "A", "Le coût des soins des retraités n'est pas suivi.", "610 retraités cotisent en août 2026 (16 en février 2024), 11 954 FCFA par mois en moyenne. On ne sait pas si cela couvre leurs soins.", "Suivre séparément les soins des retraités.", "Moyenne", "P3"),
  I("22.a", "E", "AR", "La durée de la garantie santé n'est pas fixée.", "Aucun texte transmis ne dit si elle est à vie ou limitée.", "Fixer la durée de la garantie par décision écrite.", "Faible", "P2"),
  I("22.b", "M", "D", "Les âges des retraités manquent.", "Date de naissance connue pour 105 retraités seulement.", "Obtenir de la CGRAE les dates de naissance.", "Moyenne", "P2"),
  I("27.b", "M", "C", "Une partie des retraités n'est pas prélevée.", "Sans quotité cessible suffisante, pas de prélèvement, et la liste n'a pas été transmise. Le virement de la CGRAE n'est pas rapproché du fichier retour.", "Suivre les non-prélevés et rapprocher chaque virement.", "Moyenne", "P2"),
  I("26.d", "F", "D", "Deux fichiers retour manquent.", "Janvier et juin 2024.", "Transmettre les deux fichiers.", "Faible", "P3")],
 "EFC": [
  I("6", "C", "CA", "Les intérêts promis ne sont pas comptabilisés.", "Pour l'EFC : 2 027 062 663 FCFA à 5,50 % de 2020 au 31/08/2026, sur 3 278 913 238 FCFA au total.", "Comptabiliser chaque année les intérêts dus.", "Moyenne", "P1"),
  I("7", "C", "RG", "Aucun taux n'est officiellement fixé.", "3,50 %, 5,50 % ou 14 % selon le document. Même à 3,50 %, une année d'intérêts (309 709 577 FCFA en 2025) dépasse les produits de tous les placements.", "Faire fixer le taux par le conseil d'administration.", "Faible", "P1"),
  I("11.a", "C", "A", "Les sorties ne sont pas prévues.", "L'EFC sort aux départs à la retraite, imprévisibles sans dates de naissance. Aucun échéancier n'existe.", "Établir l'échéancier des départs à la retraite.", "Élevée", "P1"),
  I("10.b", "E", "A", "L'EFC a triplé sans réserve dédiée.", "3 461 975 893 FCFA en 2020, 10 514 270 966 FCFA au 31/08/2026. Le dépôt à terme EFC est vidé chaque décembre pour payer l'EFC.", "Constituer une réserve réservée à l'épargne.", "Moyenne", "P2"),
  I("8", "E", "D", "Pas de compte individuel par épargnant.", "Le solde de chaque policier n'est pas connu.", "Tenir un compte par épargnant et envoyer un relevé annuel.", "Élevée", "P2"),
  I("26.e", "M", "D", "La hausse de l'EFC à 10 000 FCFA n'est pas documentée.", "Décision de 2021 non transmise. [[Le premier prélèvement des nouvelles recrues peut être décalé d'un mois.]]", "Transmettre la décision et dater chaque souscription.", "Faible", "P3")],
 "PEP": [
  I("6.a", "C", "C", "Les intérêts promis ne sont pas comptabilisés.", "Pour le PEP : 1 251 850 574 FCFA à 3,50 % de 2020 au 31/08/2026.", "Comptabiliser chaque année les intérêts dus.", "Moyenne", "P1"),
  I("10.a", "E", "A", "Le PEP sort vite.", "46,17 % de l'épargne PEP est sortie en 2023, 23,60 % en 2025.", "Prévoir la liquidité selon les échéances des plans.", "Moyenne", "P2"),
  I("27.c", "M", "D", "Des prélèvements PEP sont de nature inconnue.", "174 461 860 FCFA de références PTI, RPEP, EP, TI et PTQ de 2020 à 2026.", "Préciser la nature de ces prélèvements.", "Faible", "P2"),
  I("23.a", "M", "C", "Les intérêts du dépôt PEP sont enregistrés en une fois.", "Dépôt de 505 383 858 FCFA placé de 2020 à 2025 : 132 663 262 FCFA d'intérêts enregistrés en 2025 seulement.", "Enregistrer les intérêts chaque année.", "Faible", "P3")],
 "MEDOC": [
  I("26.f", "E", "D", "MEDOC n'est pas documenté.", "Ni conventions avec les pharmacies, ni période de fonctionnement, ni décision de suspension.", "Transmettre la documentation de MEDOC.", "Faible", "P3"),
  I("27.d", "M", "C", "Les prélèvements MEDOC sont incomplets.", "Fichiers reçus de novembre 2025 à août 2026 seulement. Produit ou garantie : nature à préciser.", "Transmettre les fichiers manquants et préciser la nature.", "Faible", "P3")],
 "Grenier": [
  I("25", "E", "AG", "Le Grenier grandit vite avec l'épargne.", "Encours de 7 900 000 FCFA fin 2021 à 1 044 302 107 FCFA au 31/08/2026, financé à 87,66 % par le FPPN.", "Fixer par décision un plafond de financement du Grenier.", "Moyenne", "P3"),
  I("4", "F", "C", "Les pertes probables ne sont pas comptées.", "14 016 392 FCFA selon la politique de crédit (40 %, 80 % ou 100 % selon le retard).", "Déprécier chaque année les prêts en retard.", "Faible", "P3"),
  I("4.a", "M", "C", "126 prêts ont plus de 180 jours de retard.", "Les retards de fin janvier 2026 se sont résorbés, signe d'échéances non prélevées fin 2025.", "Suivre chaque mois les retards et les échéances non prélevées.", "Faible", "P2"),
  I("25.a", "F", "A", "Les prêts sont courts mais l'habitat allonge le risque.", "Durée médiane de 10 mois, taux médian de 18,00 %. Scolaire et habitat : 66,12 % des montants de 2025.", "Encadrer la part des crédits longs.", "Faible", "P3")],
 "Flotte automobile": [
  I("18.a", "M", "C", "Les trop-perçus ne sont pas suivis comme une dette.", "Le risque est porté par GNA : pas de provision au FPPN. Mais le compte 47111106 n'est pas justifié de 2023 à 2026.", "Justifier les mouvements et traiter les trop-perçus comme une dette envers les adhérents.", "Faible", "P3"),
  I("26.g", "M", "D", "La convention avec GNA manque.", "Convention d'apporteur d'affaires et décomptes non transmis : l'assiette de la commission de 20 % n'est pas vérifiable.", "Transmettre la convention et rapprocher chaque année la commission des primes.", "Faible", "P3"),
  I("27.e", "F", "D", "Un fichier retour manque.", "Janvier 2026.", "Transmettre le fichier.", "Faible", "P3")],
 "Clinique (CMP)": [
  I("17", "E", "CR", "La clinique n'a pas de comptabilité séparée.", "7 083 534 565 FCFA au bilan, dont 1 280 300 860 FCFA de charges de fonctionnement. On ne sait pas si elle gagne ou perd de l'argent.", "Tenir une comptabilité séparée et présenter son résultat chaque année.", "Moyenne", "P2"),
  I("12.a", "E", "A", "Sa valeur n'est pas établie.", "Aucune expertise. Elle est financée en grande partie par l'épargne.", "Faire évaluer la clinique par un expert.", "Élevée", "P2")],
 "Cotisations et précomptes": [
  I("20", "M", "C", "La répartition maladie et décès diffère selon la source.", "Total de 5 % respecté, mais la Solde prélève 3,77 % et 1,23 % au lieu des 3,50 % et 1,50 % des statuts. La comptabilité retient 3,52 % et 1,48 %.", "Notifier à la Solde la clé des statuts et l'appliquer en comptabilité.", "Faible", "P2"),
  I("27", "E", "C", "Les fichiers ne sont pas rapprochés.", "Le fichier envoyé n'est jamais comparé au fichier retour. Aucun fichier envoyé transmis sur 80 mois.", "Rapprocher chaque mois l'aller et le retour, avec un état des rejets signé.", "Moyenne", "P2"),
  I("26.h", "M", "D", "2020 à 2022 ne sont pas vérifiables.", "Prélèvements mensuels disponibles seulement depuis 2023 ou 2024 selon le poste.", "Transmettre les fichiers de 2020 à 2022.", "Moyenne", "P3"),
  I("27.f", "M", "D", "Rejets et paiements manuels non suivis.", "États des rejets de la Solde et liste des paiements manuels non transmis.", "Transmettre ces états chaque mois.", "Faible", "P3"),
  I("27.g", "F", "C", "La dette de l'État est soldée.", "704 189 157 FCFA dus fin 2025, compte soldé au 31/08/2026.", "Suivre chaque mois les reversements de l'État.", "Faible", "P4")],
 "Comptabilité et provisions": [
  I("1", "C", "CR", "Les provisions ne sont pas calculées.", "4 829 298 290 FCFA identiques depuis 2020, soit 4 809 716 898 FCFA au-dessus du besoin mesuré. Rien pour les intérêts de l'épargne.", "Note de calcul annuelle signée et reclassement décidé par les organes.", "Moyenne", "P1"),
  I("15", "E", "C", "Les fonds propres sont négatifs.", "-879 498 700 FCFA au 31/08/2026. Comptes remis en ordre (constats 1, 4 et 6), ils seraient positifs d'environ 637 millions de FCFA.", "Affecter les excédents aux réserves et faire les reclassements.", "Moyenne", "P2"),
  I("9", "E", "G", "L'épargne due est minorée en Assemblée générale.", "Rapport de gestion 2025 : 4 387 823 635 FCFA présentés, pour 13 814 066 964 FCFA dus.", "Présenter le solde dû, distinct de la collecte.", "Faible", "P2"),
  I("18", "M", "C", "Des sommes durables sont en comptes d'attente.", "11 603 989 675 FCFA d'avances et de créances suivis comme du court terme.", "Reclasser dans les comptes prévus par le plan comptable.", "Faible", "P3"),
  I("28", "F", "C", "Deux versions des comptes 2023 existent.", "Excédent de 928 699 102 FCFA ou de 930 629 102 FCFA.", "Transmettre la version approuvée.", "Faible", "P3"),
  I("5", "F", "C", "La provision du personnel est suffisante.", "43 666 678 FCFA pour 24 446 566 FCFA recalculés. Méthode non écrite.", "Rédiger la note de calcul annuelle.", "Faible", "P4")],
 "Trésorerie et placements": [
  I("11", "C", "A", "L'argent disponible ne couvre pas l'année.", "34,42 % des paiements de l'année couverts au 31/08/2026. Manque cumulé de 12 596 052 822 FCFA à cinq ans.", "Réserve de trésorerie et échéancier des sorties d'épargne.", "Élevée", "P1"),
  I("10", "E", "AR", "L'épargne n'est pas séparée.", "16,28 % de l'épargne couverte par des actifs qui lui sont réservés.", "Comptes réservés à l'épargne et objectif de couverture.", "Moyenne", "P2"),
  I("12", "E", "A", "Les biens ne se vendent pas vite.", "Plus des trois quarts de l'actif à plus de cinq ans. Couverture des dettes : 96,33 % sur le papier, 27,19 % sans les biens non évalués. Créance sur l'hôpital de 3 407 519 850 FCFA inchangée depuis 2020.", "Évaluer les biens et obtenir un accord de l'État sur la créance.", "Élevée", "P2"),
  I("23", "E", "AR", "Les placements sont trop faibles.", "5,53 % de l'épargne réellement placée en 2025. Rendement de 4,90 % par an, mais les intérêts dus dépassent les produits de 470 608 168 FCFA en 2025.", "Faire voter une politique de placement.", "Moyenne", "P2"),
  I("13", "M", "A", "Une partie de la trésorerie dort.", "2 356 399 123 FCFA sur des comptes non rémunérés.", "Placer à court terme la trésorerie non utilisée.", "Faible", "P3"),
  I("23.b", "M", "R", "Le fonds commun de placement n'est pas suivi.", "500 000 000 FCFA placés en 2025, non garantis et non revalorisés dans les comptes.", "Suivre la valeur liquidative.", "Faible", "P3")],
 "Gouvernance": [
  I("14", "C", "R", "Plusieurs règles de solidité ne sont pas respectées.", "14 indicateurs sur 36 non conformes de 2022 à 2025, dont la réserve obligatoire, la liquidité et les fonds propres. Retrait d'agrément possible.", "Plan de mise en conformité chiffré présenté à l'AIRMS.", "Élevée", "P1"),
  I("24", "E", "G", "Les placements ne passent pas par le conseil.", "[[Selon le chef comptable, le directeur général décide seul.]] Projet de Cocody sans budget. Mandat du directeur général documenté pour 2020 à 2023.", "Délibération du conseil d'administration pour chaque placement.", "Faible", "P2"),
  I("24.a", "E", "G", "Aucun pilotage actif-passif.", "Ni politique votée, ni comité, ni limites.", "Politique actif-passif votée et comité actif-passif.", "Moyenne", "P2"),
  I("16", "M", "R", "Les frais dépassent la norme.", "36,34 % des cotisations en 2025, 44,12 % sur huit mois de 2026, pour un maximum de 20 %.", "Budget de fonctionnement avec un objectif de 20 %.", "Moyenne", "P3")],
 "Données et informatique": [
  I("26", "E", "D", "Des données essentielles manquent.", "Pas de base des prestations datée, ni de dates de naissance, ni de fichier des ayants droit.", "Transmettre les données demandées.", "Moyenne", "P1"),
  I("26.i", "M", "D", "Le fichier des mutualistes est incomplet.", "Ni date de naissance ni sexe des actifs, ni date d'adhésion, ni statut, ni motif de sortie.", "Extraire du logiciel le fichier complet des mutualistes.", "Moyenne", "P1"),
  I("26.k", "M", "D", "Le lien entre matricule et mécano manque.", "La table de correspondance n'a pas été transmise.", "Transmettre la table de correspondance.", "Faible", "P3"),
  I("26.j", "F", "D", "Certains exports sont difficiles à lire.", "Fichiers PEP de 2020 à 2022 au format texte. Exports du Grenier sans la lettre-clé du mécano.", "Fournir des exports Excel complets.", "Faible", "P3")],
}
PRODUITS = ["Maladie", "Décès", "Retraités (CGRAE)", "EFC", "PEP", "MEDOC", "Grenier", "Flotte automobile", "Clinique (CMP)"]
TRANSV = ["Cotisations et précomptes", "Comptabilité et provisions", "Trésorerie et placements", "Gouvernance", "Données et informatique"]
ORDRE_C = {"C": 0, "E": 1, "M": 2, "F": 3}


def pire(act):
    return min((it["crit"] for it in ACT[act]), key=ORDRE_C.get)


def run(p, texte, taille, gras=False, couleur=NOIR, surl=False):
    r_ = p.add_run()
    r_.text = texte
    r_.font.name = "Helvetica"
    r_.font.size = Pt(taille)
    r_.font.bold = gras
    r_.font.color.rgb = RGBColor.from_string(couleur)
    if surl:
        surligner(r_)
    return r_


def texte_balise(p, texte, taille, gras=False):
    for morceau, g_, s_ in morceaux_balises(texte):
        run(p, morceau, taille, gras or g_, NOIR, s_)


def boite_txt(slide, x, y, l, h, fond=None, bord="BBBCBC"):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(x), Cm(y), Cm(l), Cm(h))
    if fond:
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(fond)
    else:
        sh.fill.background()
    sh.line.color.rgb = RGBColor.from_string(bord)
    sh.line.width = Pt(0.75)
    sh.shadow.inherit = False
    tf = sh.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = tf.margin_right = Cm(0.15)
    tf.margin_top = tf.margin_bottom = Cm(0.08)
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.paragraphs[0].alignment = PP_ALIGN.LEFT
    return sh


def gauche(tf):
    for p_ in tf.paragraphs:
        if p_.alignment is None or p_.alignment == PP_ALIGN.CENTER:
            p_.alignment = PP_ALIGN.LEFT


def trait(slide, x1, y1, x2, y2, couleur="112A40", ep=1.25):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x1), Cm(y1), Cm(x2), Cm(y2))
    c.line.color.rgb = RGBColor.from_string(couleur)
    c.line.width = Pt(ep)
    return c


def lignes_est(texte, largeur, taille):
    cpl = max(10, int(largeur / (taille * 0.0176)))
    n = 0
    for par in str(texte).split("\n"):
        n += max(1, -(-len(par.replace("**", "").replace("[[", "").replace("]]", "")) // cpl))
    return n


def h_ligne(taille):
    return taille * 0.0353 * 1.2


COURT = {"Retraités (CGRAE)": "Retraités", "Flotte automobile": "Flotte auto", "Clinique (CMP)": "Clinique", "Cotisations et précomptes": "Cotisations",
         "Comptabilité et provisions": "Comptabilité", "Trésorerie et placements": "Trésorerie", "Données et informatique": "Données"}
LC, LR, LX, LP = 16.6, 9.4, 2.6, 2.6   # colonnes : constat et commentaire, recommandation, complexité, priorité
Y_TAB, Y_FIN = 6.4, 15.8


def hauteur_item(it, taille):
    h1 = lignes_est(f"● {it['num']} {it['titre']} {it['texte']}", LC - 0.35, taille) * h_ligne(taille)
    h2 = lignes_est(f"→ {it['reco']}", LR - 0.35, taille) * h_ligne(taille)
    return max(h1, h2, 0.55) + 0.22


def paquets_activite(a):
    """Découpe une activité en diapositives : texte entre 10,5 et 9 points, au plus la hauteur du tableau."""
    dispo = Y_FIN - Y_TAB
    for t_ in (10.5, 10.0, 9.5, 9.0):
        if sum(hauteur_item(it, t_) for it in ACT[a]) <= dispo:
            return [(ACT[a], t_)]
    t_, res, cour, h = 9.0, [], [], 0.0
    for it in ACT[a]:
        hi = hauteur_item(it, t_)
        if cour and h + hi > dispo:
            res.append((cour, t_))
            cour, h = [], 0.0
        cour.append(it)
        h += hi
    res.append((cour, t_))
    return res


def pilule(slide, x, y, l, h, texte, fond, couleur="FFFFFF", taille=9.5):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(x), Cm(y), Cm(l), Cm(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = RGBColor.from_string(fond)
    sh.line.fill.background()
    sh.shadow.inherit = False
    tf = sh.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Cm(0.02)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run(p, texte, taille, True, couleur)
    return sh


def dessin_activites(slide, spec):
    blocs, taille = spec["blocs"], spec["taille"]
    actives = {b_["act"] for b_ in blocs}
    x0 = 1.4
    # arborescence des activités : repère de navigation, l'activité de la diapositive en noir
    toutes = PRODUITS + TRANSV
    lch = (31.2 - 0.15 * (len(toutes) - 1) - 0.5) / len(toutes)
    xs, x = [], x0
    for k, nom in enumerate(toutes):
        if k == len(PRODUITS):
            x += 0.5
        xs.append(x)
        x += lch + 0.15
    for lib, debut_, fin_ in (("PRODUITS", 0, len(PRODUITS) - 1), ("FONCTIONS TRANSVERSES", len(PRODUITS), len(toutes) - 1)):
        b = boite_txt(slide, xs[debut_], 3.92, xs[fin_] + lch - xs[debut_], 0.42, BLEU, BLEU)
        b.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = b.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run(p, lib, 8, True, "FFFFFF")
    for k, nom in enumerate(toutes):
        actif = nom in actives
        b = boite_txt(slide, xs[k], 4.42, lch, 0.5, NOIR if actif else "F2F2F2", NOIR if actif else "D9D9D9")
        b.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        b.text_frame.margin_left = b.text_frame.margin_right = Cm(0.03)
        p = b.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run(p, "● ", 7.5, True, CRIT[pire(nom)][0])
        run(p, COURT.get(nom, nom), 7.5, actif, "FFFFFF" if actif else "53565A")
    y = 5.1
    reste_total = max(0.0, (Y_FIN - 5.1) - sum(1.25 + sum(hauteur_item(it, taille) for it in b_["items"]) for b_ in blocs))
    n_lignes = sum(len(b_["items"]) for b_ in blocs)
    for bloc in blocs:
        a, items = bloc["act"], bloc["items"]
        n_p1 = sum(1 for it in ACT[a] if it["prio"] == "P1")
        b = boite_txt(slide, x0, y, 31.2, 0.62, NOIR, NOIR)
        b.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = b.text_frame.paragraphs[0]
        run(p, "● ", 12, True, CRIT[pire(a)][0])
        run(p, ("PRODUITS" if a in PRODUITS else "FONCTIONS TRANSVERSES") + "  ›  ", 10, False, "BBBCBC")
        run(p, a.upper(), 12, True, "FFFFFF")
        run(p, f"     {len(ACT[a])} constats, dont {n_p1} à traiter dans les trois mois" if n_p1 else f"     {len(ACT[a])} constats", 9.5, False, "FFFFFF")
        if bloc.get("partie"):
            run(p, f"     ({bloc['partie']})", 9.5, False, "FFFFFF")
        y += 0.7
        x = x0
        for lib, l in (("Constat et commentaire", LC), ("Recommandation", LR), ("Complexité", LX), ("Priorité", LP)):
            b = boite_txt(slide, x, y, l, 0.5, "D9D9D9", "D9D9D9")
            b.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = b.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if l > 5 else PP_ALIGN.CENTER
            run(p, lib, 9.5, True)
            x += l
        y += 0.55
        for k_, it in enumerate(items):
            h = hauteur_item(it, taille) + min(reste_total / max(1, n_lignes), 0.5)
            fond = "FFFFFF" if k_ % 2 == 0 else "F7F7F7"
            c = boite_txt(slide, x0, y, LC, h, fond, "D9D9D9")
            c.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = c.text_frame.paragraphs[0]
            run(p, "● ", taille + 2, True, CRIT[it["crit"]][0])
            run(p, it["num"] + "   ", taille, True, BLEU)
            run(p, it["titre"] + " ", taille, True)
            texte_balise(p, it["texte"], taille)
            c = boite_txt(slide, x0 + LC, y, LR, h, fond, "D9D9D9")
            c.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = c.text_frame.paragraphs[0]
            run(p, "→ ", taille + 1, True, ORANGE)
            run(p, it["reco"], taille)
            for xx, l in ((x0 + LC + LR, LX), (x0 + LC + LR + LX, LP)):
                boite_txt(slide, xx, y, l, h, fond, "D9D9D9")
            ph = min(0.6, h - 0.12)
            pilule(slide, x0 + LC + LR + 0.2, y + (h - ph) / 2, LX - 0.4, ph, it["cplx"], {"Faible": "D9EAD3", "Moyenne": "FFF2CC", "Élevée": "F4CCCC"}[it["cplx"]], NOIR, 9)
            pilule(slide, x0 + LC + LR + LX + 0.45, y + (h - ph) / 2, LP - 0.9, ph, it["prio"], PRIO[it["prio"]], "FFFFFF", 10)
            y += h
        y += 0.08
    # légende
    tb = slide.shapes.add_textbox(Cm(x0), Cm(Y_FIN + 0.08), Cm(31.2), Cm(0.42))
    tb.text_frame.word_wrap = True
    p = tb.text_frame.paragraphs[0]
    run(p, "Criticité : ", 7.5, True)
    for code in ("C", "E", "M", "F"):
        run(p, "● ", 8.5, True, CRIT[code][0])
        run(p, CRIT[code][1] + "   ", 7.5)
    run(p, "Priorité : P1 dans les 3 mois, P2 dans les 6 mois, P3 de 6 à 18 mois, P4 au-delà.   Complexité : Faible (une décision suffit), Moyenne (travail interne), Élevée (tiers ou financement).   Numéros : ceux du résumé exécutif.", 7.5)


def dessin_carte(slide):
    def act_box(x, y, l, a):
        b = boite_txt(slide, x, y, l, 0.72, "FFFFFF", "112A40")
        b.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = b.text_frame.paragraphs[0]
        run(p, "● ", 11, True, CRIT[pire(a)][0])
        run(p, a, 9.5, True)
        run(p, f"  ({len(ACT[a])})", 8.5, False, "53565A")

    def titre_box(x, y, l, h, t, sous, fond=BLEU):
        b = boite_txt(slide, x, y, l, h, fond, fond)
        b.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = b.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run(p, t, 10.5, True, "FFFFFF")
        if sous:
            p2 = b.text_frame.add_paragraph()
            p2.alignment = PP_ALIGN.CENTER
            run(p2, sous, 8, False, "FFFFFF")

    def fleche_lib(x1, y1, x2, y2, lib, dx=0.0, dy=-0.45):
        c = trait(slide, x1, y1, x2, y2, "53565A", 1.75)
        ln = c.line._get_or_add_ln()
        tail = ln.makeelement(qn("a:tailEnd"), {"type": "triangle"})
        ln.append(tail)
        tb = slide.shapes.add_textbox(Cm(min(x1, x2) + dx), Cm((y1 + y2) / 2 + dy), Cm(abs(x2 - x1) + 0.2), Cm(0.4))
        p = tb.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run(p, lib, 7.5, False, "53565A")

    titre_box(1.4, 4.4, 5.6, 1.5, "Policiers en activité", "cotisations et épargne prélevées sur la solde (Solde)")
    titre_box(1.4, 7.6, 5.6, 1.5, "Policiers retraités", "cotisation santé prélevée sur la pension (CGRAE)")
    titre_box(10.3, 5.3, 4.6, 2.9, "FPPN", "cotisations : 3 418 141 722 FCFA en 2025\népargne due : 15 743 516 391 FCFA", NOIR)
    fleche_lib(7.0, 5.15, 10.3, 6.2, "5 % du salaire, EFC, PEP", dy=0.15)
    fleche_lib(7.0, 8.35, 10.3, 7.4, "barème CGRAE", dy=0.05)
    gx, gl = 18.6, 14.0
    titre_box(gx, 3.95, gl, 0.5, "Garanties", None)
    l3 = (gl - 0.4) / 3
    for k, a in enumerate(["Maladie", "Décès", "Retraités (CGRAE)"]):
        act_box(gx + k * (l3 + 0.2), 4.5, l3, a)
    titre_box(gx, 5.5, gl, 0.5, "Épargne", None)
    l2 = (gl - 0.2) / 2
    for k, a in enumerate(["EFC", "PEP"]):
        act_box(gx + k * (l2 + 0.2), 6.05, l2, a)
    titre_box(gx, 7.05, gl, 0.5, "Services aux policiers", None)
    for k, a in enumerate(["Grenier", "Flotte automobile", "MEDOC", "Clinique (CMP)"]):
        act_box(gx + (k % 2) * (l2 + 0.2), 7.6 + (k // 2) * 0.8, l2, a)
    fleche_lib(14.9, 5.6, gx, 4.85, "soins, allocations", dy=-0.55)
    fleche_lib(14.9, 6.75, gx, 6.4, "remboursement de l'épargne", dy=-0.5)
    fleche_lib(14.9, 7.9, gx, 8.35, "prêts, assurance, soins", dy=0.1)
    titre_box(1.4, 10.35, 31.2, 0.55, "Fonctions transverses", None, "53565A")
    lx = (31.2 - 0.4 * 4) / 5
    for k, a in enumerate(TRANSV):
        act_box(1.4 + k * (lx + 0.4), 11.05, lx, a)
    tb = slide.shapes.add_textbox(Cm(1.4), Cm(12.15), Cm(31.2), Cm(3.4))
    tb.text_frame.word_wrap = True
    tf = tb.text_frame
    for j, (b_, t_) in enumerate((("Comment lire. ", "Chaque activité porte une boule de la couleur de son constat le plus grave et, entre parenthèses, son nombre de constats. Les flèches suivent l'argent : il entre par les prélèvements sur la solde et sur la pension, il sort en prestations, en remboursements d'épargne et en services."),
                                 ("Ce que montre la carte. ", "Les activités les plus exposées touchent à l'épargne (EFC, PEP, trésorerie et placements) et aux règles de solidité (gouvernance, comptabilité et provisions). Les services aux policiers posent surtout la question de l'épargne qu'ils immobilisent (Grenier, clinique) et de leurs documents."))):
        p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
        p.space_after = Pt(4)
        run(p, b_, 11, True)
        run(p, t_, 11)


def dessin_matrice(slide):
    toutes = PRODUITS + TRANSV
    x0, l0, lc = 1.4, 6.4, 3.9
    y0, h0, hl = 3.95, 0.62, 0.74
    entetes = ["Activité"] + [t[1] for t in TYPES] + ["Total"]
    larg = [l0] + [lc] * 5 + [5.3]
    x = x0
    for e, l in zip(entetes, larg):
        b = boite_txt(slide, x, y0, l, h0, NOIR, NOIR)
        b.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = b.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run(p, e, 9.5, True, "FFFFFF")
        x += l
    fonds = {"C": "F4CCCC", "E": "FCE5CD", "M": "FFF2CC", "F": "D9EAD3"}
    for i, a in enumerate(toutes):
        y = y0 + h0 + i * hl
        x = x0
        b = boite_txt(slide, x, y, l0, hl, "F2F2F2")
        b.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        run(b.text_frame.paragraphs[0], a, 9, True)
        x += l0
        for code, _ in TYPES:
            its = [it for it in ACT[a] if code in it["types"]]
            if its:
                cr = min((it["crit"] for it in its), key=ORDRE_C.get)
                b = boite_txt(slide, x, y, lc, hl, fonds[cr])
                b.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
                p = b.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                run(p, "● ", 10, True, CRIT[cr][0])
                run(p, str(len(its)), 9.5, True)
            else:
                boite_txt(slide, x, y, lc, hl, "FFFFFF")
            x += lc
        b = boite_txt(slide, x, y, 5.3, hl, "FFFFFF")
        b.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = b.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run(p, f"{len(ACT[a])} constats", 9)
    tb = slide.shapes.add_textbox(Cm(x0), Cm(y0 + h0 + len(toutes) * hl + 0.05), Cm(31.2), Cm(0.45))
    p = tb.text_frame.paragraphs[0]
    run(p, "Dans chaque case : nombre de constats du type, et boule de la couleur du plus grave. ", 7.5, True)
    for code in ("C", "E", "M", "F"):
        run(p, "● ", 8.5, True, CRIT[code][0])
        run(p, CRIT[code][1] + "   ", 7.5)
    run(p, "Un constat peut relever de plusieurs types.", 7.5)


# ---------------------------------------------------------------- diapositives
SL.append({"type": "couverture", "notes": "Les constats du volet actuariel et statistique, présentés activité par activité. Numéros, complexité et priorité repris du résumé exécutif du 26/10/2026."})
SL.append({"type": "sommaire", "notes": ""})
separateur(1, "Vue d'ensemble", "Les activités du FPPN, les types de constats et leur gravité")
n_tot = sum(len(v) for v in ACT.values())
contenu({"titre": "Carte des activités du FPPN", "sous_titre": "L'argent des policiers entre par la solde et la pension, et sort en prestations, en épargne et en services",
         "synthese": f"{n_tot} constats sur 14 activités. Les plus graves touchent l'épargne, la liquidité et les règles de solidité"},
        "Présenter la carte de gauche à droite : d'où vient l'argent, ce que fait le FPPN, où va l'argent. Puis les fonctions transverses qui font tourner l'ensemble.",
        carte=True, source=RAP + " ; résumé exécutif du 26/10/2026")
_nd = sum(1 for a in PRODUITS + TRANSV if any("D" in it["types"] for it in ACT[a]))
contenu({"titre": "Activités et types de constats", "sous_titre": f"Les manques de données touchent {_nd} activités sur 14, les constats critiques portent sur l'épargne et la solidité",
         "synthese": "Matrice des 14 activités et des cinq types de constats : données, comptabilité, actuariel, conformité, gouvernance"},
        "Lecture en ligne : une activité, ses types de constats. Lecture en colonne : un type de constat, les activités touchées.",
        matrice_types=True, source=RAP + " ; analyses Adenka")
_syn = []
for a in PRODUITS + TRANSV:
    its = ACT[a]
    pg = min(its, key=lambda it: (ORDRE_C[it["crit"]], it["prio"]))
    _syn.append([a, str(len(its)), CRIT[pire(a)][1], f"{pg['num']} : {pg['titre']}", min(it["prio"] for it in its), str(sum(1 for it in its if it["prio"] == "P1"))])
contenu({"titre": "Synthèse chiffrée par activité", "sous_titre": "Cinq activités ont au moins un constat critique à traiter dans les trois mois",
         "synthese": "Nombre de constats, criticité la plus haute, constat le plus grave et priorité la plus haute de chaque activité"},
        "Le tableau sert de sommaire détaillé : chaque activité est développée dans les diapositives suivantes.",
        tableau={"titre": None, "entetes": ["Activité", "Constats", "Criticité la plus haute", "Constat le plus grave", "Priorité la plus haute", "Constats en P1"],
                 "largeurs": [6.4, 2.4, 4.2, 12.0, 3.6, 2.6], "taille": 9, "sans_code": True, "lignes": _syn},
        source=RAP + " ; résumé exécutif du 26/10/2026")

SOUS = {"Maladie": "La garantie maladie perd de l'argent depuis 2025, et ses provisions ne sont pas calculées",
        "Décès": "La garantie décès est excédentaire, mais sa dette envers les familles grossit",
        "Retraités (CGRAE)": "Les retraités cotisent de plus en plus, mais leurs soins et leurs âges ne sont pas suivis",
        "EFC": "L'épargne de fin de carrière a triplé, sans taux fixé ni intérêts comptabilisés",
        "PEP": "L'épargne volontaire sort vite et ses intérêts ne sont pas comptabilisés",
        "MEDOC": "MEDOC fonctionne sans documentation ni fichiers complets",
        "Grenier": "Le Grenier grandit vite, financé presque entièrement par l'épargne",
        "Flotte automobile": "Le risque est porté par GNA, mais les documents et le suivi manquent",
        "Clinique (CMP)": "La clinique n'a ni comptes séparés ni valeur établie",
        "Cotisations et précomptes": "Le total prélevé est juste, mais les fichiers ne sont pas rapprochés",
        "Comptabilité et provisions": "Les comptes ne mesurent pas les vraies dettes envers les policiers",
        "Trésorerie et placements": "L'argent disponible ne couvre pas l'année, et l'épargne n'est pas mise à part",
        "Gouvernance": "Les règles de solidité ne sont pas respectées, et les décisions ne sont pas encadrées",
        "Données et informatique": "Des données essentielles manquent pour finir les calculs"}
def h_bloc(items, t_):
    return 1.25 + sum(hauteur_item(it, t_) for it in items) + 0.08


def grave(items):
    return min(items, key=lambda it: (ORDRE_C[it["crit"]], it["prio"]))


def libelle(it):
    t = it["titre"].rstrip(".")
    return t if t[:2].isupper() else t[0].lower() + t[1:]


for groupe, liste in (("Produits", PRODUITS), ("Fonctions transverses", TRANSV)):
    separateur(2 if groupe == "Produits" else 3, groupe, "Chaque constat se lit sur une ligne : son commentaire, sa recommandation, sa complexité et sa priorité")
    diapos, k = [], 0
    while k < len(liste):
        a = liste[k]
        paq = paquets_activite(a)
        if len(paq) == 1 and k + 1 < len(liste):
            b_ = liste[k + 1]
            for t_ in (10.0, 9.5):
                if len(ACT[a]) + len(ACT[b_]) <= 6 and h_bloc(ACT[a], t_) + h_bloc(ACT[b_], t_) <= Y_FIN - 5.1:
                    diapos.append(([{"act": a, "items": ACT[a]}, {"act": b_, "items": ACT[b_]}], t_))
                    k += 2
                    break
            else:
                diapos.append(([{"act": a, "items": ACT[a]}], paq[0][1]))
                k += 1
            continue
        for n_, (items, t_) in enumerate(paq):
            diapos.append(([{"act": a, "items": items, "partie": f"{n_ + 1}/{len(paq)}" if len(paq) > 1 else None}], t_))
        k += 1
    for blocs, t_ in diapos:
        tous = [it for b_ in blocs for it in b_["items"]]
        g_ = grave(tous)
        act_g = next(b_["act"] for b_ in blocs if g_ in b_["items"])
        titre = " et ".join(b_["act"] for b_ in blocs) + (f" ({blocs[0]['partie']})" if blocs[0].get("partie") else "")
        contenu({"titre": titre, "sous_titre": SOUS[act_g], "synthese": f"Constat le plus grave : {g_['num']}, {libelle(g_)}"},
                "Lire chaque ligne de gauche à droite : le constat et son commentaire, la recommandation, la complexité de l'action et sa priorité. Le bandeau du haut situe l'activité parmi les 14 activités du FPPN.",
                activites={"blocs": blocs, "taille": t_},
                source=RAP + " ; logigrammes du FPPN (« FLOW_CHART_COTISATION_ET_PRODUITS_DEPARGNE_1.pptx ») ; résumé exécutif du 26/10/2026")

separateur("A", "Annexes", "Circuits des cotisations par produit, avec les contrôles absents")
_SRC_FLOW = "Source : « FLOW_CHART_COTISATION_ET_PRODUITS_DEPARGNE_1.pptx » (FPPN, reçu le 23/10/2026). Contrôles absents : analyses Adenka"
for titre_, sous_, log_, ana_ in [
        ("Circuit de la maladie et du décès", "Le fichier envoyé à la Solde est produit hors du FPPN : le recouvrement ne peut pas être mesuré", "maladie_deces",
         "**Le service du personnel calcule et transmet le fichier sans le communiquer au FPPN.** Le FPPN ne peut ni vérifier les taux retenus (3,77 % et 1,23 % au lieu de 3,50 % et 1,50 %), ni mesurer le taux de recouvrement (constats 20 et 27)."),
        ("Circuit de l'EFC et du PEP", "Le FPPN produit le fichier de l'épargne, mais ne compare pas l'envoi et le retour", "epargne",
         "**Le circuit s'arrête à l'enregistrement des prélèvements.** Aucune étape ne compare les fichiers, ne calcule les intérêts dus ni ne produit de relevé par épargnant (constats 6, 8 et 27)."),
        ("Circuit de la flotte automobile (GNA)", "Le FPPN encaisse une commission de 20 % dont la base n'est pas documentée", "gna",
         "**La convention avec GNA n'a pas été transmise.** La commission ne peut pas être vérifiée, et le fichier retour de janvier 2026 manque (constats 18.a, 26.g et 27.e)."),
        ("Circuit des retraités (CGRAE)", "La cotisation suit le barème de la CGRAE, dans la limite de la quotité cessible", "cgrae",
         "**Seuls les retraités qui ont une quotité cessible suffisante sont prélevés, et la liste des autres n'a pas été transmise** (constats 22.a, 27.b et 26.d).")]:
    contenu({"titre": titre_, "sous_titre": sous_, "synthese": "Logigramme du FPPN redessiné : les cadres rouges signalent les contrôles absents"},
            "Couloirs repris du logigramme du FPPN. Cadre rouge : contrôle absent ou fichier manquant.", logigramme=log_, analyse=[ana_], source=_SRC_FLOW)
SL.append({"type": "merci", "notes": ""})
print("constats :", n_tot, {a: len(v) for a, v in ACT.items()})
'''

code = ENTETE + "\n" + CONTENU + "\n" + MOTEUR + "\nprint(SORTIE, len(SL), 'diapositives')\n"
g = {"__name__": "__main__", "__file__": os.path.join(ICI, "rapport_v12_fppn.py"), "SKILL": SKILL, "SORTIE": SORTIE}
exec(compile(code, "deck_constats_activites_fppn (assemblé)", "exec"), g)
