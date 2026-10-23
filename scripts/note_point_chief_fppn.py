"""Note interne : fil conducteur du point d'avancement de la mission FPPN avec le Chief (réunion de 30 minutes).

Ordre du fil conducteur : contexte, périmètre, tâches principales, tâches ajoutées, travaux réalisés, conclusions,
limites, reste à faire, recommandations, taux d'avancement, prochaines étapes.

Usage : python3 note_point_chief_fppn.py <sortie.docx>
"""
import sys

sys.path.insert(0, "/root/.claude/skills/synced/891d7168-f164-499a-9885-98883788a845_be017d1e-653f-49e6-8fd9-683fb0156b08/adenka-rapport-word/scripts")
from adenka_docx import RapportAdenka  # noqa: E402
from docx import Document  # noqa: E402
from docx.enum.text import WD_COLOR_INDEX  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

SORTIE = sys.argv[1]
# assertions à confirmer : surlignées en jaune après l'enregistrement
JAUNE = []


def jaune(t):
    JAUNE.append(t)
    return t


r = RapportAdenka(titre="Point d'avancement de la mission", client="FPPN", type_livrable="Note interne de présentation", date="Octobre 2026")
r.sigles({
    "AIRMS": "Agence ivoirienne de régulation de la mutualité sociale",
    "CGRAE": "Caisse générale de retraite des agents de l'État",
    "CIMA": "Conférence interafricaine des marchés d'assurances",
    "CMP": "Clinique médicale du policier",
    "EFC": "Entraide fin de carrière (épargne versée au départ à la retraite)",
    "FCP": "Fonds commun de placement",
    "FPPN": "Fonds de prévoyance de la Police nationale",
    "GNA": "Assureur de la flotte automobile",
    "PBC": "Liste des documents demandés au client",
    "PCMS": "Plan comptable des mutuelles sociales",
    "PEP": "Plan Épargne Plus (épargne volontaire)",
    "TDR": "Termes de référence",
    "UEMOA": "Union économique et monétaire ouest-africaine",
})
r.abreviations({"DAT": "Dépôt à terme", "h-j": "Homme-jour", "j": "Jour de travail", "S/C": "Prestations / cotisations (sinistralité)", "n°": "Numéro"})

# ---------------------------------------------------------------- synthèse
r.synthese("L'ESSENTIEL EN UNE PAGE")
r.para("**Le volet actuariel et statistique de l'audit du FPPN doit répondre à quatre questions : les engagements envers les adhérents, la suffisance des provisions, l'adossement des actifs aux dettes et l'adéquation des placements.** "
       "Le FPPN est la mutuelle des policiers de Côte d'Ivoire. Il couvre la maladie et le décès, gère deux produits d'épargne (EFC et PEP) et prête aux policiers par le Grenier du Policier. "
       "La période couverte va du 01/01/2020 au 31/08/2026.")
r.para("**Sur les documents reçus, deux réponses sont partiellement démontrées et deux sont insuffisantes.** "
       "L'épargne due aux adhérents atteint 15 743 516 391 FCFA au 31/08/2026, sans les intérêts promis (3 278 913 238 FCFA non comptabilisés). "
       "Les provisions maladie et obsèques (4 829 298 290 FCFA) n'ont pas été recalculées depuis 2020. "
       "Les actifs disponibles à moins d'un an couvrent 34,42 % des paiements de l'année, et 14 des 36 ratios de l'UEMOA ne sont pas respectés.")
r.para("**L'avancement est de 63,85 %, ce qui est bon pour la partie qui dépend d'Adenka, mais le reste dépend de données que le FPPN n'a pas encore transmises.** "
       "Le rapport compte 291 diapositives, appuyées sur un classeur Excel de calcul de 99 onglets. Les 2 885 chiffres du rapport sont contrôlés automatiquement contre les calculs. "
       "La restitution provisoire au FPPN est prévue les 23 et 24/11/2026, le rapport définitif le 30/11/2026.")
r.tableau(["Étape", "Durée", "Ce qu'il faut dire en une phrase", "À montrer"],
          [["1. Contexte", "2 min", "Qui est le FPPN et pourquoi cet audit", "Diapositive 42"],
           ["2. Mon périmètre", "1 min", "Le volet actuariel et statistique, sur 2020 au 31/08/2026", "Diapositive 44"],
           ["3. Tâches principales", "2 min", "Les quatre questions de l'offre technique", "Diapositive 4"],
           ["4. Tâches ajoutées", "2 min", "Ce que l'offre et la nature du FPPN imposent en plus", "Partie 1.4 de cette note"],
           ["5. Ce que j'ai fait", "4 min", "Collecte, contrôles, analyses, outils, livrables", "Diapositives 46 à 58, 229 à 239 (survol)"],
           ["6. Conclusions", "5 min", "Les quatre verdicts, pourquoi et comment", "Diapositives 4 à 8, 10 et 11"],
           ["7. Limites", "2 min", "Ce que les conclusions ne disent pas encore", "Diapositive 241"],
           ["8. Reste à faire", "1 min", "Les calculs qui attendent les données", "Diapositive 242"],
           ["9. Recommandations", "2 min", "Neuf actions, dont quatre dans les six mois", "Diapositives 25 à 27"],
           ["10. Avancement", "1 min", "63,85 %, bon pour notre part, à risque pour celle du FPPN", "Plan de travail (Excel)"],
           ["11. Prochaines étapes", "1 min", "Relance du FPPN, calculs, restitution, rapport définitif", "Partie 1.11 de cette note"],
           ["Questions", "7 min", "Retour au rapport selon ses questions", "Carte du rapport (partie 2)"]],
          titre="Déroulé de la réunion (30 minutes)", source="Note interne Adenka.", largeurs_cm=[3.6, 1.5, 7.6, 5.7], lignes_total=())

# ---------------------------------------------------------------- partie 1
r.partie(1, "Fil conducteur", introduction="Le texte en italique est une proposition de ce qu'il faut dire. Il peut être lu tel quel ou reformulé. Chaque étape se termine par une phrase de transition.")

r.h2("1.1. LE CONTEXTE DE LA MISSION (2 MINUTES)")
r.para("**Le message : le FPPN gère l'argent de prévoyance et d'épargne des policiers, et l'audit doit dire s'il peut tenir ses promesses.**")
r.para("*« Le FPPN est la mutuelle des policiers de Côte d'Ivoire. Les cotisations sont prélevées directement sur la solde, ou sur la pension pour les retraités. "
       "En échange, le FPPN rembourse les soins, verse une allocation au décès, gère une épargne de fin de carrière (l'EFC) et une épargne volontaire (le PEP), et prête aux policiers par le Grenier du Policier. "
       "Il relève de la mutualité sociale de l'UEMOA et il est suivi par l'AIRMS, le régulateur ivoirien. "
       "Adenka réalise l'audit comptable et financier du FPPN sur la période du 01/01/2020 au 31/08/2026, sur la base d'une offre technique validée par le FPPN. »*", gras_auto=False)
r.para("**Transition.** *« Dans cet audit, mon intervention porte sur un volet précis. »*", gras_auto=False)

r.h2("1.2. MON PÉRIMÈTRE (1 MINUTE)")
r.para("**Le message : j'interviens sur le volet actuariel et statistique, c'est-à-dire sur ce que le FPPN devra payer plus tard et sur sa capacité à le payer.**")
r.para("*« L'audit comptable vérifie que les comptes sont justes. Mon volet regarde vers l'avenir : combien le FPPN doit à ses adhérents, s'il a mis assez de côté, et si son argent sera disponible au bon moment. "
       "Mon profil dans l'offre technique (page 74) et le pôle 4 (page 70) fixent ce périmètre. La valorisation et le rendement des placements, ainsi que les ratios des TDR, relèvent de l'expert financier (pages 70 et 74). »*", gras_auto=False)
r.para("**Transition.** *« Concrètement, ce périmètre se traduit par quatre questions. »*", gras_auto=False)

r.h2("1.3. LES TÂCHES PRINCIPALES DE MON PÉRIMÈTRE (2 MINUTES, DIAPOSITIVE 4)")
r.para("**Le message : quatre questions structurent tout le rapport, et chacune reçoit une réponse explicite dès la diapositive 4.**")
r.tableau(["Question de l'offre technique", "Ce que cela veut dire en clair"],
          [["1. Évaluer les engagements du FPPN envers ses adhérents et leurs ayants droit", "Combien le FPPN doit-il aujourd'hui et devra-t-il plus tard : épargne, intérêts, prestations maladie et décès ?"],
           ["2. Apprécier la suffisance des provisions au regard des prestations servies", "Les sommes mises de côté au bilan suffisent-elles à payer les prestations déjà nées ?"],
           ["3. Apprécier l'adossement actif-passif", "L'argent du FPPN sera-t-il disponible au moment où il devra payer ?"],
           ["4. Analyser l'adéquation des placements aux engagements", "Les placements sont-ils sûrs, disponibles et assez rentables pour servir l'épargne promise ?"]],
          titre="Les quatre tâches principales", source="Offre technique Adenka validée par le FPPN, profil de l'actuaire (p. 74) ; rapport du 23/10/2026, diapositive 4.", largeurs_cm=[7.6, 10.8], lignes_total=())
r.para("**Transition.** *« Pour y répondre correctement, l'offre technique et la nature du FPPN imposent d'autres travaux. »*", gras_auto=False)

r.h2("1.4. LES TÂCHES AJOUTÉES AU VU DE L'OFFRE TECHNIQUE ET DE LA NATURE DE LA MISSION (2 MINUTES)")
r.para("**Le message : ces travaux ne sont pas un dépassement du périmètre, ils sont prévus par l'offre ou indispensables pour répondre aux quatre questions.**")
r.tableau(["Tâche ajoutée", "Pourquoi elle est nécessaire", "Fondement"],
          [["Contrôler les données de base (effectifs, cotisations, fichiers de précomptes)", "Un calcul actuariel ne vaut que par ses données. Les fichiers du FPPN contenaient des lignes de total, des montants doublés et des fichiers vides.", "Offre technique, pôle 4 (p. 70)"],
           ["Calculer les 36 ratios de la Décision n° 05/2022 de l'UEMOA", "Apprécier l'adéquation des réserves aux exigences prudentielles et comparer les réserves aux indicateurs de l'AIRMS.", "Offre technique, pôle 4 (p. 70) et risque R3 (p. 36)"],
           ["Projeter l'équilibre technique", "Vérifier que les cotisations futures couvriront les prestations et les frais futurs.", "Offre technique, pôle 4 (p. 70)"],
           ["Volet statistique : qualité des données, décès, santé", "Mesurer la sinistralité, les délais, les dossiers non payés, la mortalité comparée aux tables CIMA.", "Plan d'analyse statistique du 16/10/2026"],
           ["Revoir les autres provisions (prêts du Grenier, compte GNA, départ à la retraite du personnel)", "Elles pèsent sur les fonds propres et donc sur la capacité à payer les adhérents.", "Plan de travail, tâche 4.4"],
           ["Établir le cadre réglementaire applicable", "Le FPPN relève de la mutualité sociale de l'UEMOA. Le Code CIMA ne sert que de référence technique, et ce choix doit être justifié.", "Règlement n° 07/2009, règlement d'exécution n° 004/2022, Décision n° 05/2022, PCMS"],
           ["Décrire les circuits des cotisations par produit", "Repérer où l'argent peut se perdre entre la Solde, la CGRAE, le GNA et la comptabilité.", "Logigrammes du FPPN, entretien du 19/10/2026"]],
          titre="Les tâches ajoutées et leur fondement", source="Offre technique Adenka ; plan d'analyse du 16/10/2026 ; plan de travail du 23/10/2026.", largeurs_cm=[5.4, 8.0, 5.0], lignes_total=())
r.para("**Transition.** *« Voici maintenant ce qui a été réalisé. »*", gras_auto=False)

r.h2("1.5. CE QUE J'AI PU FAIRE (4 MINUTES)")
r.para("**Le message : les travaux couvrent tout le périmètre, des fichiers bruts du FPPN jusqu'aux recommandations, avec des outils qui rendent chaque chiffre vérifiable.**")
r.tableau(["Étape", "Ce qui a été fait", "Chiffre à citer"],
          [["Lancement", "Réunions de lancement des 13 et 14/10/2026, liste des documents demandés (PBC)", "90 documents demandés"],
           ["Collecte", "Sept envois du FPPN du 15/10 au 23/10/2026, pointés un par un", "6 reçus, 41 partiels, 43 non reçus"],
           ["Entretiens", "Responsable des cotisations, chef comptable, service des prestations (16 et 19/10/2026), avec guides et comptes rendus", "4 entretiens documentés"],
           ["Données", "Précomptes sur salaire et sur pension, balances comptables, grands livres, états financiers, fichiers du personnel, prêts du Grenier", "4 454 416 lignes de précomptes contrôlées"],
           ["Contrôles", "Vrai format des fichiers, lignes de total, montants doublés, fichiers vides, rapprochement aller et retour", "7 balances et 506 feuilles de précomptes relues"],
           ["Analyses", "Effectifs, cotisations, sinistralité, épargne, prêts, retraités, provisions, adossement, placements, 36 ratios de l'UEMOA, sensibilités, mortalité, projection", "14 parties de rapport"],
           ["Synthèse", "Constats, matrices des risques et de conformité, recommandations avec horizon", "9 recommandations"]],
          titre="Les travaux réalisés en sept étapes", source="Rapport du 23/10/2026 ; PBC du 23/10/2026 ; plan de travail du 23/10/2026.", largeurs_cm=[2.6, 10.0, 5.8], lignes_total=())
r.para("**Les outils utilisés.** C'est un point fort à mettre en avant devant un profil finance ou audit (diapositives 229 à 239).", gras_auto=False)
r.puces(["**Python** pour lire et contrôler les fichiers bruts du FPPN : plus de quatre millions de lignes, plusieurs formats différents, fichiers scannés lus par reconnaissance de caractères.",
         "**Une empreinte numérique de chaque fichier reçu** : elle prouve que les fichiers analysés sont identiques à ceux transmis.",
         "**Un classeur Excel de 99 onglets** : chaque calcul y est refait par une formule visible, avec une explication en tête d'onglet. L'onglet « 0_controle » compare chaque formule au chiffre du rapport : 0 écart.",
         "**Un contrôle automatique du rapport** : les 2 885 chiffres écrits dans le rapport sont retrouvés un par un dans les calculs (diapositive 239).",
         "**Une vérification indépendante** : les chiffres ont été recalculés une seconde fois depuis les fichiers bruts. 18 anomalies ont été relevées et corrigées.",
         "**Des logigrammes modifiables** dans PowerPoint, et un code couleur par boules expliqué sous chaque tableau."])
r.para("**Les livrables produits à ce jour.** Les citer sans les ouvrir, pour montrer l'exhaustivité.", gras_auto=False)
r.puces(["Rapport actuariel et statistique (291 diapositives, PowerPoint et PDF).",
         "Classeur de calcul de 99 onglets.",
         "Liste des documents demandés (PBC), pointée à chaque envoi.",
         "Liste des 26 questions et 84 documents à transmettre par le FPPN.",
         "Fichier de vérification de cohérence des chiffres et des sources.",
         "Plan de travail pointé, avec le taux de réalisation par phase.",
         "Logigrammes du FPPN mis à jour, avec les contrôles manquants signalés.",
         "Guides d'entretien, comptes rendus et note de point d'étape du 21/10/2026."])
r.para("*« Le rapport fait 291 diapositives parce que chaque chiffre est sourcé et démontré. Le résumé exécutif, aux diapositives 3 à 27, se suffit à lui-même. Le reste sert de preuve. »*", gras_auto=False)
r.para("**Transition.** *« Ces travaux conduisent à quatre conclusions. »*", gras_auto=False)

r.h2("1.6. LES CONCLUSIONS, POURQUOI ET COMMENT J'Y ARRIVE (5 MINUTES, DIAPOSITIVES 4 À 8)")
r.para("**Le message : le FPPN dégage des excédents chaque année, mais il n'est pas en mesure de démontrer qu'il pourra rembourser l'épargne de ses adhérents.** "
       "Montrer la diapositive 4, puis donner pour chaque question le verdict, le chiffre et la façon dont il a été obtenu.")
r.tableau(["Question et verdict", "Conclusion", "Pourquoi et comment"],
          [["1. Engagements\nPartiellement démontré",
            "L'épargne due vaut 15 743 516 391 FCFA au 31/08/2026. Avec les intérêts promis, l'engagement atteint 19 022 429 629 FCFA.",
            "Soldes d'épargne lus dans les balances 2020 au 31/08/2026. Intérêts recalculés sur l'encours moyen de chaque année, aux taux de 3,50 % (PEP) et 5,50 % (EFC) : 3 278 913 238 FCFA jamais comptabilisés. Sans comptes individuels, le montant ne peut pas être vérifié épargnant par épargnant."],
           ["2. Provisions\nPartiellement démontré",
            "Les provisions maladie et obsèques (4 829 298 290 FCFA) sont figées depuis 2020 et dépassent de 4 809 716 898 FCFA le besoin mesuré (19 581 392 FCFA).",
            "Besoin mesuré sur les factures de soins arrivées après la clôture (journal des sinistres) et sur les décès déclarés après la clôture. Le total n'est pas trop faible, il est mal affecté : rien n'est provisionné pour les intérêts de l'épargne."],
           ["3. Adossement\nInsuffisant",
            "Les actifs disponibles à moins d'un an couvrent 34,42 % des paiements attendus dans l'année. Le manque cumulé atteint 12 596 052 822 FCFA à cinq ans.",
            "Actifs classés selon leur délai de mobilisation, dettes classées par échéance, sorties d'épargne au taux observé. Les actifs sont retenus à leur valeur comptable : c'est une estimation favorable au FPPN."],
           ["4. Placements\nInsuffisant",
            "En 2025, l'argent réellement placé (763 752 135 FCFA) représente 5,53 % de l'épargne due. Le rendement moyen est de 4,90 % par an de 2021 à 2025.",
            "Encours des dépôts à terme et des FCP lu jour par jour dans les grands livres. L'épargne a surtout financé des actifs peu liquides : clinique, terrains, prêts. Aucune politique de placement votée n'a été transmise."]],
          titre="Les quatre conclusions et leur démonstration", source="Rapport du 23/10/2026, diapositives 4 à 8.", largeurs_cm=[3.2, 6.8, 8.4], lignes_total=())
r.para("**Les constats transversaux à citer.** Les fonds propres sont négatifs à chaque clôture depuis 2020 (-879 498 700 FCFA au 31/08/2026), et 14 ratios de l'UEMOA sur 36 ne sont pas respectés, dont la réserve obligatoire, les fonds propres et la liquidité (diapositive 20). "
       "La synthèse des constats (diapositives 10 et 11) présente chaque constat avec le fait chiffré, la norme qui s'applique, l'écart et son effet.", gras_auto=False)
r.para("*« La non-conformité sur deux exercices de suite permet à l'AIRMS de retirer l'agrément (règlement n° 07/2009, art. 84). Ce n'est pas automatique, mais c'est l'enjeu principal pour le FPPN. »*", gras_auto=False)
r.para("**Transition.** *« Ces conclusions ont des limites qu'il faut connaître. »*", gras_auto=False)

r.h2("1.7. LES LIMITES DE CES CONCLUSIONS (2 MINUTES, DIAPOSITIVE 241)")
r.para("**Le message : les limites changent la précision des montants, pas le sens des verdicts.**")
r.puces(["**Trois données manquantes bloquent les calculs actuariels centraux** : la base des prestations datée (date des soins, du dépôt et du paiement de chaque facture), les dates de naissance des mutualistes et des retraités, et le fichier des ayants droit. Sans elles, ni provision dossier par dossier, ni mortalité par âge, ni engagement santé des retraités.",
         "**L'engagement d'épargne est global** : aucun compte individuel ni taux servi arrêté n'a été transmis.",
         "**L'adossement est une estimation favorable** : les actifs sont retenus à leur valeur comptable, sans expertise de la clinique ni des terrains.",
         "**Les méthodes avancées ne sont pas applicables** : duration, valeur à risque ou modèle stochastique exigent une courbe de taux, un échéancier et un taux servi fiable.",
         "**Certains taux sont mesurés sur une période courte** : le taux de cotisation maladie de 3,77 % repose sur 8 mois de fichiers (04 à 10/2024 et 01/2025).",
         "**Quelques constats reposent sur des déclarations du FPPN** : ils sont surlignés en jaune dans le rapport et seront confirmés ou retirés à réception des pièces.",
         "**Les PDF du Code CIMA et du référentiel comptable des SFD sont illisibles** (police mal encodée) : seules les pages de la circulaire n° 00230 ont pu être lues."])
r.para("**Transition.** *« D'où le travail qui reste à faire. »*", gras_auto=False)

r.h2("1.8. LE RESTE À FAIRE (1 MINUTE, DIAPOSITIVE 242)")
r.para("**Le message : il reste surtout des calculs qui attendent les données du FPPN, puis la finalisation du rapport.**")
r.puces(["**Engagements** : calculs par âge et engagement par épargnant, dès réception des dates de naissance et des comptes individuels.",
         "**Provisions** : provision maladie et décès dossier par dossier, boni et mali par exercice, à partir de la base des prestations datée.",
         "**Placements** : adossement par échéance, à partir de l'inventaire et de l'échéancier signés des placements.",
         "**Projection de l'équilibre technique** par garantie, après les calculs d'engagements et de provisions.",
         "**Ratios** : onze indicateurs de gouvernance en attente de pièces.",
         "**Rapport** : restitution, intégration des commentaires du FPPN, version définitive."])
r.para("**Transition.** *« Sans attendre ces données, les premières conclusions permettent déjà de recommander des actions. »*", gras_auto=False)

r.h2("1.9. LES RECOMMANDATIONS (2 MINUTES, DIAPOSITIVES 25 À 27)")
r.para("**Le message : neuf recommandations, chacune liée à un constat, avec un horizon. Les quatre premières sont à engager dans les six mois.**")
r.tableau(["Recommandation", "Horizon"],
          [["1. Calculer les provisions maladie et décès à chaque clôture", "Court terme (moins de 6 mois)"],
           ["2. Arrêter le taux servi sur l'épargne et tenir les comptes individuels", "Court terme (moins de 6 mois)"],
           ["3. Transmettre les données qui manquent aux calculs actuariels", "Court terme (moins de 6 mois)"],
           ["4. Présenter à l'AIRMS un plan de mise en conformité prudentielle", "Court terme (moins de 6 mois)"],
           ["5. Constituer une réserve de liquidité et adopter une politique de placement", "Moyen terme (6 à 18 mois)"],
           ["6. Rééquilibrer la garantie maladie", "Moyen terme (6 à 18 mois)"],
           ["7. Encadrer les investissements dans la clinique, le Grenier et le projet de Cocody", "Moyen terme (6 à 18 mois)"],
           ["8. Faire réaliser l'étude actuarielle des risques longs et certifier les comptes", "Long terme (plus de 18 mois)"],
           ["9. Mettre en place un dispositif actuariel permanent", "Long terme (plus de 18 mois)"]],
          titre="Les neuf recommandations", source="Rapport du 23/10/2026, diapositives 25 et 26.", largeurs_cm=[12.6, 5.8], lignes_total=())
r.para("*« La matrice de criticité (diapositive 12) croise la gravité de chaque constat et l'horizon de l'action. La feuille de route (diapositive 27) enchaîne quatre étapes : collecter, calculer, vérifier, projeter. »*", gras_auto=False)
r.para("**Transition.** *« Où en est-on aujourd'hui ? »*", gras_auto=False)

r.h2("1.10. LE TAUX D'AVANCEMENT (1 MINUTE)")
r.para("**Le message : avec 63,85 % de réalisation, l'avancement est bon pour la partie qui dépend d'Adenka, mais il est à risque pour la partie qui dépend des données du FPPN.** "
       "Montrer l'onglet « Synthèse-PT_Actuariat » du plan de travail.")
r.puces(["**Pourquoi c'est un bon avancement.** La mission a démarré le 13/10/2026. Au 23/10/2026, 9 jours ouvrés sur 35 se sont écoulés jusqu'au rapport définitif du 30/11/2026, soit 25,71 % du calendrier, pour 63,85 % des travaux réalisés.",
         "**Pourquoi il reste un risque.** Le reste (36,15 %) se concentre dans l'évaluation des engagements (54,37 %) et la projection (40,00 %), qui attendent les données du FPPN, et dans la finalisation du rapport (42,00 %), prévue en fin de mission.",
         "**Ce que cela représente.** Environ 14,10 jours de travail restent à réaliser sur 39 jours estimés, et 20 jours ouvrés séparent le 26/10/2026 de la restitution du 23/11/2026."])
r.tableau(["Phase", "Taux de réalisation", "Ce qui reste"],
          [["1. Cadrage : textes, garanties, règles", "87,50 %", "Versions des statuts antérieures à 2023, âge de départ du personnel"],
           ["2. Contrôle et rapprochement des données", "73,44 %", "Base des prestations datée, comptes individuels d'épargne"],
           ["3. Évaluation des engagements", "54,37 %", "Calculs par âge, bloqués sans dates de naissance"],
           ["4. Revue des provisions", "71,50 %", "Boni et mali par exercice, faute de dates de survenance"],
           ["5. Adossement et placements", "75,62 %", "Inventaire et échéancier signés des placements"],
           ["6. Ratios de l'AIRMS", "87,50 %", "Onze indicateurs de gouvernance en attente de pièces"],
           ["7. Projection de l'équilibre technique", "40,00 %", "Projection par garantie, après les phases 3 et 4"],
           ["8. Rapport et revues", "42,00 %", "Restitution, intégration des commentaires, version définitive"],
           ["Ensemble de la mission", "63,85 %", "39 jours estimés dans le scénario « données dégradées »"]],
          titre="Avancement par phase au 23/10/2026", source="FPPN_Plan_de_travail_Actuariat_20261023.xlsx, onglet « Synthèse-PT_Actuariat ». Taux pondérés par le temps estimé de chaque tâche.", largeurs_cm=[6.6, 3.2, 8.6], lignes_total=(8,))
r.para("**Transition.** *« Pour tenir le calendrier, voici les prochaines étapes. »*", gras_auto=False)

r.h2("1.11. LES PROCHAINES ÉTAPES (1 MINUTE)")
r.tableau(["Étape", "Qui", "Quand"],
          [["Envoyer au FPPN la liste des 26 questions et 84 documents, avec une date limite de réponse", "Adenka", "Dès validation interne"],
           ["Relancer le FPPN en priorité sur la base des prestations datée, les dates de naissance, les comptes individuels d'épargne et la décision fixant le taux servi", "Adenka et équipe d'audit", "Chaque semaine jusqu'à réception"],
           ["Calculer les engagements et les provisions dossier par dossier, puis projeter l'équilibre technique", "Adenka", "À réception des données"],
           ["Rapprocher les ratios et les placements avec les travaux de l'expert financier", "Adenka et expert financier", "Avant la restitution"],
           ["Faire relire le rapport par une seconde personne", "Adenka", "Avant la restitution"],
           ["Restitution provisoire au FPPN", "Adenka", "23 et 24/11/2026"],
           ["Intégration des commentaires du FPPN", "Adenka", "Du 25 au 27/11/2026"],
           ["Transmission du rapport définitif", "Adenka", "30/11/2026"]],
          titre="Prochaines étapes", source="Plan de travail du 23/10/2026 ; rapport du 23/10/2026, diapositive 27.", largeurs_cm=[10.6, 3.8, 4.0], lignes_total=())
r.para("**Clôture de la présentation.** *« En résumé, le travail est avancé et chaque chiffre est vérifiable. Les conclusions sont déjà solides sur l'adossement et les placements. Les montants des engagements et des provisions seront précisés dès que le FPPN transmettra ses données. »*", gras_auto=False)

r.h2("1.12. QUESTIONS ET RETOUR AU RAPPORT (7 MINUTES)")
r.para("**Laisser le Chief choisir les sujets à approfondir.** La carte du rapport (partie 2) donne la diapositive à ouvrir pour chaque sujet. Les questions les plus probables et leurs réponses figurent en partie 3.")

# ---------------------------------------------------------------- partie 2
r.partie(2, "Préparation et carte du rapport", introduction="Cette partie liste ce qu'il faut ouvrir avant la réunion, la carte du rapport et les points à faire arbitrer par le Chief.")
r.h2("2.1. DOCUMENTS À OUVRIR")
r.puces(["**Le rapport en PDF** (FPPN_Rapport_actuariel_statistique_20261023.pdf) : plus fluide que le PowerPoint pour sauter d'une page à l'autre.",
         "**Le classeur de calcul** (FPPN_Calculs_rapport_20261023.xlsx), ouvert sur l'onglet « 0_controle » : il montre que chaque chiffre est recalculé par une formule.",
         "**Le plan de travail** (FPPN_Plan_de_travail_Actuariat_20261023.xlsx), onglet « Synthèse-PT_Actuariat » : avancement par phase.",
         "**La liste des demandes au FPPN** (FPPN_Demandes_au_FPPN_20261023.xlsx) : 26 questions et 84 documents attendus."])
r.h2("2.2. CONSEILS DE PRÉSENTATION")
r.puces(["**Dérouler les onze étapes sans entrer dans les diapositives de détail.** Le Chief ne connaît pas la mission : il doit d'abord voir l'étendue des travaux.",
         "**Ne pas lire les tableaux.** Montrer la diapositive, donner le chiffre clé en une phrase, passer à la suite.",
         "**Garder la carte du rapport à portée de main** : elle permet de retrouver en quelques secondes la diapositive qui répond à une question.",
         "**Dire clairement ce qui est établi et ce qui reste à confirmer.**"])
r.h2("2.3. CARTE DU RAPPORT")
r.tableau(["Partie", "Diapositives", "Ce qu'elle démontre", "À ouvrir si besoin"],
          [["Résumé exécutif", "3 à 27", "Les réponses, les constats, les matrices et les recommandations", "4, 10, 23, 25"],
           ["1. Comprendre le FPPN", "28 à 40", "Garanties, textes, circuits des cotisations par produit d'après les logigrammes du FPPN", "32, 35 à 38"],
           ["2. Contexte et démarche", "41 à 45", "Mission, périmètre, documents reçus", "42, 44"],
           ["3. Données et qualité", "46 à 58", "Couverture des fichiers, contrôles, fichiers vides, qualité variable par variable", "48, 53, 54"],
           ["4. Effectifs et cotisations", "59 à 71", "Effectifs, taux réellement prélevés (3,77 % et 1,23 % au lieu de 3,50 % et 1,50 %)", "61, 65"],
           ["5. Sinistralité et résultat", "72 à 86", "La sinistralité maladie passe de 6,00 % en 2021 à 73,53 % en 2025", "75"],
           ["6. Épargne des mutualistes", "87 à 99", "Intérêts dus non comptabilisés : 3 278 913 238 FCFA", "96, 98"],
           ["7. Prêts du Grenier", "100 à 108", "Encours de 1 044 302 107 FCFA, sans dépréciation", "102"],
           ["8. Retraités et annexes", "109 à 113", "Retraités précomptés par la CGRAE, flotte automobile, projets", "110"],
           ["9. Bilan, provisions et placements", "114 à 167", "Provisions recalculées, adossement, tests de résistance, placements", "123, 145, 155, 162"],
           ["10. Conformité prudentielle", "168 à 217", "Ratios par garantie, 36 ratios de l'UEMOA avec leur définition, sensibilités", "175, 178, 205"],
           ["11. Tables de mortalité", "218 à 224", "Choix des tables pour les calculs par âge", "220"],
           ["12. Projection de l'équilibre", "225 à 228", "Deux scénarios de 2026 à 2030", "227"],
           ["13. Outils et traçabilité", "229 à 239", "Traitement des données, extraits de code, contrôle des chiffres", "231, 239"],
           ["14. Analyses en attente", "240 à 244", "Données manquantes, questions au FPPN", "241, 243"],
           ["Annexes", "245 à 291", "Méthodes écartées, fichiers sources, formules, validité, abréviations", "246, 257"]],
          titre="Carte du rapport du 23/10/2026", source="Rapport du 23/10/2026 (291 diapositives).", largeurs_cm=[4.4, 2.4, 8.4, 3.2], lignes_total=())
r.h2("2.4. POINTS À FAIRE ARBITRER PAR LE CHIEF")
r.puces(["**La formulation des verdicts « insuffisant »** et de l'enjeu de retrait d'agrément avant la restitution au FPPN.",
         "**L'envoi au FPPN de la liste des 26 questions et 84 documents**, avec une date limite de réponse compatible avec la restitution du 23/11/2026.",
         "**La date de restitution provisoire** si les données manquantes n'arrivent pas : maintenir la date avec des verdicts partiels, ou la décaler.",
         "**Le partage avec l'expert financier** : le rapport calcule les 36 ratios et le rendement des placements parce qu'ils servent l'appréciation des réserves. Les chiffres communs sont à rapprocher avant la restitution.",
         jaune("**Le temps passé** : l'offre technique prévoit 14 hommes-jours pour le pôle 4 (p. 70), alors que le plan de travail estime 26 jours avec des données correctes et 39 jours avec des données dégradées. Le périmètre exact des 14 hommes-jours est à confirmer dans l'offre."),
         "**La relecture qualité** du rapport par une seconde personne avant la restitution."])

# ---------------------------------------------------------------- partie 3
r.partie(3, "Questions probables", introduction="Réponses courtes, à adapter. Les chiffres renvoient au rapport du 23/10/2026.")
qr = [
    ("Pouvez-vous conclure sans les données manquantes ?",
     "Oui pour l'essentiel. Les conclusions reposent sur les balances, les grands livres, les fichiers de précomptes et les états financiers, qui sont complets. Les données manquantes permettront de préciser les montants, pas de changer le sens des verdicts. C'est pourquoi deux réponses sont « partiellement démontrées » et non « démontrées »."),
    ("Comment être sûr que les chiffres sont justes ?",
     "Par trois contrôles. Chaque chiffre du rapport est retrouvé automatiquement dans les calculs (2 885 sur 2 885). Chaque calcul est refait par une formule Excel, avec 0 écart. Une vérification indépendante a recalculé les chiffres depuis les fichiers bruts et relevé 18 anomalies, toutes corrigées."),
    ("Pourquoi avoir fait plus que les quatre questions ?",
     "Parce que l'offre technique le prévoit : le pôle 4 demande de contrôler les données de base, d'apprécier les réserves au regard des exigences prudentielles et de projeter l'équilibre technique. Sans ces travaux, les quatre réponses ne seraient pas démontrées."),
    ("Le FPPN est-il en faillite ?",
     "Non. Il dégage des excédents chaque année depuis 2020, sauf sur les huit premiers mois de 2026. Son problème est la liquidité et la mesure de ses engagements d'épargne. Ses fonds propres sont négatifs (-879 498 700 FCFA au 31/08/2026) et il ne pourrait pas rembourser une vague de retraits sans vendre des actifs peu liquides."),
    ("Les provisions sont supérieures au besoin : est-ce un problème ?",
     "Oui, parce qu'elles ne couvrent aucun risque identifié. Les provisions maladie et obsèques dépassent le besoin mesuré de 4 809 716 898 FCFA, alors que les intérêts dus sur l'épargne (3 278 913 238 FCFA) ne sont pas provisionnés. Le résultat et les fonds propres sont donc mal mesurés."),
    ("Qu'est-ce que l'adossement, et pourquoi est-il insuffisant ?",
     "C'est la capacité à payer chaque dette à son échéance avec des actifs disponibles à la même date. Au FPPN, l'épargne, exigible à tout moment, finance des actifs longs : clinique, terrains, prêts. Les actifs disponibles à moins d'un an couvrent 34,42 % des paiements de l'année."),
    ("En quoi votre travail diffère-t-il de l'audit comptable ?",
     "L'audit comptable vérifie que les comptes sont exacts. Le volet actuariel vérifie que les engagements futurs sont bien mesurés et couverts : provisions, épargne, adossement, placements, ratios de l'UEMOA. Les deux se complètent, et les chiffres partagés ont été rapprochés."),
    ("Pourquoi un rapport aussi long ?",
     "Pour que chaque conclusion soit démontrée et chaque chiffre retraçable jusqu'au fichier du FPPN. Le résumé exécutif (diapositives 3 à 27) se lit seul. Le reste sert de preuve en cas de contestation."),
    ("Tiendrez-vous le délai ?",
     "Oui pour la partie qui dépend d'Adenka : 63,85 % des travaux sont réalisés après 25,71 % du calendrier. Le délai dépend de la date à laquelle le FPPN transmettra ses données. Sans elles, la restitution se fera avec des verdicts partiels, ce qui est clairement indiqué dans le rapport."),
    ("Quels risques pour Adenka ?",
     "Le principal est de conclure sur des déclarations non documentées : elles sont surlignées en jaune et seront confirmées ou retirées. Le second est le délai, pour la raison ci-dessus."),
    ("Qu'attendez-vous du FPPN maintenant ?",
     "Les réponses aux 26 questions et les 84 documents de la liste, en priorité la base des prestations datée, les dates de naissance, les comptes individuels d'épargne et la décision fixant le taux servi."),
]
for q_, a_ in qr:
    r.h3(q_)
    r.para(a_, gras_auto=False)

# ---------------------------------------------------------------- annexe
r.annexe(1, "Glossaire simple")
r.tableau(["Terme", "Explication en une phrase"],
          [["Actuariat", "Discipline qui mesure les engagements futurs (prestations, épargne, retraites) à partir des probabilités et du temps."],
           ["Adossement actif-passif", "Capacité à payer chaque dette à son échéance avec des actifs disponibles à la même date."],
           ["Ayant droit", "Conjoint ou enfant d'un adhérent, couvert par les garanties du FPPN."],
           ["Engagement", "Somme que le FPPN devra payer plus tard à ses adhérents (épargne, intérêts, prestations)."],
           ["Épargne EFC", "Épargne prélevée chaque mois sur la solde et versée au policier à son départ à la retraite."],
           ["Épargne PEP", "Épargne volontaire des policiers, rémunérée."],
           ["Fonds propres", "Ressources propres du FPPN : réserves et résultats cumulés. Négatifs, ils signifient que les pertes passées dépassent les réserves."],
           ["Lecture corrigée selon le PCMS", "Les mêmes comptes, avec l'épargne classée en ressource longue et les avances durables en emplois longs, comme le prévoit le plan comptable."],
           ["Précompte", "Retenue de la cotisation directement sur le salaire (Solde) ou sur la pension (CGRAE)."],
           ["Provision", "Montant mis de côté à la clôture pour payer des prestations déjà nées mais pas encore payées."],
           ["Ratio combiné", "Prestations et frais d'une garantie divisés par ses cotisations. Au-dessus de 100,00 %, la garantie perd de l'argent."],
           ["Ratios prudentiels", "Les 36 indicateurs de la Décision n° 05/2022 de l'UEMOA que doivent respecter les mutuelles sociales."],
           ["Sinistralité (S/C)", "Part des cotisations consommée par les prestations. La norme de l'UEMOA est comprise entre 70,00 % et 80,00 %."],
           ["Table de mortalité", "Tableau qui donne la probabilité de décéder à chaque âge, nécessaire aux calculs par âge."],
           ["Test de résistance", "Simulation d'un choc défavorable (retraits massifs, hausse des prestations) pour voir si le FPPN tient."]],
          titre="Glossaire des termes du rapport", source="Adenka.", largeurs_cm=[4.6, 13.8], lignes_total=())
r.enregistrer(SORTIE)

# ---- retouches : cadre des pages de séparation élargi (titres coupés au milieu des mots), surlignage jaune
doc = Document(SORTIE)
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
for dr in doc.element.body.iter(qn("w:drawing")):
    txt = "".join(t.text or "" for t in dr.iter(qn("w:t")))
    if txt.startswith("PARTIE") or txt.startswith("ANNEXE"):
        for e in list(dr.iter(WP + "extent")) + list(dr.iter(A + "ext")):
            if e.get("cx") == "1691640":
                e.set("cx", "4320000")
        for w_ in list(dr.iter(WP + "wrapSquare")):
            w_.addprevious(w_.makeelement(WP + "wrapNone", {}))
            w_.getparent().remove(w_)
for p in doc.element.body.iter(qn("w:p")):
    txt = "".join(t.text or "" for t in p.iter(qn("w:t")))
    if any(j.replace("**", "")[:60] in txt for j in JAUNE):
        for rr in p.iter(qn("w:r")):
            rpr = rr.find(qn("w:rPr"))
            if rpr is None:
                rpr = rr.makeelement(qn("w:rPr"), {})
                rr.insert(0, rpr)
            h = rpr.makeelement(qn("w:highlight"), {qn("w:val"): "yellow"})
            apres = [rpr.find(qn("w:" + n_)) for n_ in ("u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath")]
            apres = [e for e in apres if e is not None]
            if apres:
                min(apres, key=list(rpr).index).addprevious(h)
            else:
                rpr.append(h)
doc.save(SORTIE)
print(SORTIE)
