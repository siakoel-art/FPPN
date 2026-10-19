# -*- coding: utf-8 -*-
"""
FPPN - Registre des formules du rapport (version 5)
===================================================

Une ligne par formule d'indicateur utilisée dans le rapport : formule littérale et détaillée, données, méthode du
FPPN, méthode réglementaire, analyse de l'écart et faisabilité, impact chiffré, risques, niveau de risque, source de
la méthode, calcul (script et onglet), statut. Les totaux de tableaux, moyennes, parts et conversions d'unités ne
sont pas répertoriés un par un (décision du 19/10/2026).

Statuts : « Conforme au texte » (formule du texte appliquée telle quelle) ; « Divergent, justifié » (formule
différente du texte, avec la raison) ; « Sans texte applicable » (référence technique) ; « En attente » (donnée
manquante).

Utilisé par rapport_v5_fppn.py (annexe) et classeur_v2_fppn.py (onglet 11_formules).
"""

NC = "Non communiquée : aucun calcul du FPPN transmis"
DEC = "Décision n° 05/2022/COM/UEMOA, annexe"
REX = "Règlement d'exécution n° 004/2022/COM/UEMOA"
BAL = "Balances des comptes (Sage 100), 2020 au 31/08/2026"
PREC = "Fichiers de précomptes reçus le 15/10/2026"
GRE = "Balances âgées et production du Grenier reçues le 15/10/2026"
ATT = "En attente"


def _nb(x, dec=0):
    return f"{x:,.{dec}f}".replace(",", " ").replace(".", ",")


def _pc(x):
    return _nb(100 * x, 2) + " %"


def _f(x):
    return _nb(round(x)) + " FCFA"


def registre(Q, P, R):
    """Q : ratios.json ; P : provisions.json ; R : resultats.json. Retourne la liste des formules."""
    sin = {r["exercice"]: r for r in R["sinistralite"]}
    ps = {r["exercice"]: r for r in P["sinistralite"]}
    rt = {x["code"]: x for x in Q["ratios"]}
    pr = {r["exercice"]: r for r in P["psap_rapprochement"]}
    vchf = rt["CHF"]["valeurs"]
    chf25 = float(vchf.get("2025", vchf.get(2025)))
    L = []

    def f(id_, partie, indicateur, litterale, detaillee, donnees, fppn, reglementaire, ecart, impact, risques, niveau,
          source, calcul, statut, traces=""):
        L.append({"id": id_, "partie": partie, "indicateur": indicateur, "litterale": litterale, "detaillee": detaillee,
                  "donnees": donnees, "fppn": fppn, "reglementaire": reglementaire, "ecart": ecart, "impact": impact,
                  "risques": risques, "niveau": niveau, "source": source, "calcul": calcul, "statut": statut,
                  "traces": traces})

    # ------------------------------------------------------------------ Partie 3 : données
    f("F-03-01", 3, "Taux de précompte", "Montant précompté (fichier retour) divisé par le montant demandé (fichier aller), sur les mois où les deux fichiers existent.",
      "Somme des montants retour / somme des montants aller, par poste et par mois", PREC,
      "Non communiquée : le FPPN ne rapproche pas les fichiers aller et retour (entretien du 19/10/2026)",
      DEC + ", tableau 2 : taux de recouvrement = cotisations recouvrées / cotisations émises, au moins 90 %",
      "Même formule que le texte. Elle n'est calculable que si le FPPN détient le fichier aller, ce qui n'est pas le cas pour la maladie et le décès.",
      "Non mesurable pour la maladie et le décès (fichiers aller détenus par la direction du personnel)",
      "Fichiers aller incomplets (taux supérieur à 100 %), mois manquants", "Élevé", DEC, "analyses_fppn.py, onglet 1_aller_retour",
      "Conforme au texte", r"^(Taux min moyen|Taux max moyen|Min Grenier|Min PEP|Taux PEP|Taux janv25|Taux juil24)")
    f("F-03-02", 3, "Lignes de total écartées", "Une ligne dont le montant égale la somme des autres lignes de la feuille, à 1 % près, est un total et non un mutualiste.",
      "|montant de la ligne - (somme de la feuille - montant de la ligne)| ≤ 1 % de la somme", PREC, "Sans objet (contrôle des données)",
      "Aucun texte applicable", "Contrôle de qualité des données, nécessaire pour ne pas compter deux fois un montant.", "Lignes et montants écartés journalisés",
      "Un vrai mutualiste au montant égal à la somme serait écarté (cas non rencontré)", "Faible", "Pratique de contrôle des données (ISA 500, caractère probant des données)",
      "base_fppn.py, journal des rejets", "Sans texte applicable", r"^(Lignes total écartées|Montant total écarté|Total écartés)")
    f("F-03-03", 3, "Montants doublés corrigés", "Un montant décès formé de deux moitiés identiques écrites à la suite est ramené à sa première moitié.",
      "Si montant > 1 000 000 FCFA et moitié gauche = moitié droite, montant retenu = moitié gauche", PREC, "Sans objet (contrôle des données)",
      "Aucun texte applicable", "Correction d'une double frappe, limitée au poste décès et réversible.", "Lignes corrigées journalisées, montant d'origine conservé",
      "Correction à tort d'un montant réel symétrique", "Faible", "Pratique de contrôle des données", "base_fppn.py, colonne montant_origine",
      "Sans texte applicable", r"^(Montants corrigés|Corrigés)")

    # ------------------------------------------------------------------ Partie 4 : effectifs et cotisations
    f("F-04-01", 4, "Effectif cotisant du mois", "Nombre de matricules distincts du fichier maladie 936 du mois.", "nunique(matricule) du poste 936, mois m",
      PREC, NC, "Aucun texte ne fixe le mode de comptage", "Comptage direct sur les fichiers de la Solde.", "Trente et un mois disponibles",
      "Mois manquants avant 2024", "Moyen", "Pratique statistique (comptage d'unités distinctes)", "analyses_fppn.py, onglet 2_effectifs",
      "Sans texte applicable", r"^(Cotisants|Hausse cotisants)")
    f("F-04-02", 4, "Entrées, sorties et fidélisation", "Une entrée est un matricule présent un mois et absent le mois précédent, une sortie l'inverse. La fidélisation rapporte les matricules présents deux années de suite à ceux de l'année précédente.",
      "Entrées(m) = |M(m) \\ M(m-1)| ; sorties(m) = |M(m-1) \\ M(m)| ; fidélisation = |M(n) ∩ M(n-1)| / |M(n-1)|", PREC, NC,
      DEC + ", tableau 2 : taux de fidélisation = renouvellements n / adhérents n-1, au moins 80 %",
      "La fidélisation suit le texte. Entrées et sorties n'ont pas de norme.", "Fidélisation conforme à la norme", "Motif de sortie inconnu (retraite, décès, radiation)",
      "Moyen", DEC, "analyses_fppn.py, onglet 2_flux", "Conforme au texte", r"^(Entrées|Sorties|Fidélisation)")
    f("F-04-03", 4, "Taux de cotisation individuel", "Précompte maladie (ou décès) du mois divisé par le salaire de base du même mois, pour chaque agent.",
      "montant 936 / salaire de base ; montant 8Z6 / salaire de base", "Fichiers 936 détaillés et 8Z6",
      "Solde : 3,77 % et 1,23 % (fichiers). Barème déclaré : 3,5 % et 1,5 % (entretien du 16/10/2026, à confirmer). Répartition renvoyée à la comptabilité (entretien du 19/10/2026)",
      "Aucun texte UEMOA ne fixe le taux : le barème relève des statuts ou d'une décision du FPPN (non transmis)",
      "Trois répartitions coexistent pour le même total de 5,00 %.", "Change la sinistralité de chaque garantie (F-04-05)",
      "Excédent par risque et réserve obligatoire incalculables sans clé unique", "Élevé", "Statuts ou décision du FPPN (à transmettre)",
      "analyses_fppn.py, onglet 2_taux_936", ATT, r"^(Taux total|Taux maladie observé|Lignes taux)")
    f("F-04-04", 4, "Clé comptable maladie", "Part des cotisations maladie dans les cotisations maladie et décès, multipliée par 5 %.",
      "7010 / (7010 + 7021) × 5 %", BAL, "Clé propre à la comptabilité, non documentée",
      "Aucun texte : la clé doit découler du barème", "La comptabilité ne suit ni la Solde ni le barème déclaré.", "Écart avec la clé de la Solde mesuré en partie 4",
      "Résultat par garantie dépendant d'une écriture non documentée", "Élevé", "Statuts ou décision du FPPN (à transmettre)",
      "analyses_complements_fppn.py, onglet 8_sensibilite_cle", ATT, r"^(Clé comptable maladie)")
    f("F-04-05", 4, "Sinistralité selon la clé de répartition", "Prestations maladie divisées par les cotisations totales multipliées par la part maladie de la clé retenue.",
      "(601 + 6013 + 6018 + 60800007) / [(7010 + 7021) × part maladie]", BAL, NC,
      DEC + ", § 1.2.1.2.5 (taux global) ; § 1.1.1 (excédent par risque)", "Sensibilité : seule l'affectation des cotisations change, pas les prestations.",
      "Écart de plusieurs points entre les clés (partie 4)", "Conclusion sur l'équilibre maladie incertaine", "Élevé", DEC,
      "analyses_complements_fppn.py, onglet 8_sensibilite_cle", "Divergent, justifié", r"^(S/C clé|S/C solde|S/C décès solde|S/C barème)")
    f("F-04-06", 4, "Valeur réelle", "Valeur nominale multipliée par 100 et divisée par l'indice des prix de l'année (base 100 en 2020).",
      "V_réelle(n) = V_nominale(n) × 100 / I(n) ; I(n) = I(n-1) × (1 + inflation(n))", "Banque mondiale, FP.CPI.TOTL.ZG", "Sans objet",
      "Aucun texte applicable", "Retire l'effet des prix pour comparer des montants d'années différentes.", "Hausse réelle des cotisations mesurée en partie 4",
      "Inflation générale différente de l'inflation médicale", "Faible", "Pratique de comptabilité nationale (déflation par un indice des prix)",
      "analyses_complements_fppn.py, onglet 8_reel", "Sans texte applicable", r"^(Croissance nominale|Croissance prix|Croissance réelle|Prix|Réelle|Inflation)")
    f("F-04-07", 4, "Taux de croissance annuel moyen (TCAM)", "Taux constant qui mène de la valeur de départ à la valeur d'arrivée en n années.",
      "TCAM = (V_fin / V_début)^(1/n) - 1", "Selon l'indicateur", "Sans objet", "Aucun texte applicable",
      "Mesure de croissance standard, sensible aux deux bornes choisies.", "Voir le tableau des TCAM (partie 5)", "Bornes atypiques (2021 bas, 2026 sur 8 mois)",
      "Faible", "Pratique statistique (moyenne géométrique)", "analyses_complements_fppn.py, onglet 8_tcam", "Sans texte applicable", r"^(TCAM|Mult |V0 |V1 )")
    f("F-04-08", 4, "Tendance et intervalle de confiance", "Droite des moindres carrés ajustée sur la série, avec un intervalle de confiance à 95 % sur la pente et un test de Student.",
      "y = a + b·t ; IC95 % = b ± t(0,975 ; n-2) × écart-type(b)", "Séries mensuelles ou annuelles", "Sans objet", "Aucun texte applicable",
      "Distingue une hausse réelle d'une variation due au hasard.", "Pentes significatives au seuil de 5 %", "Peu de points (six années)",
      "Faible", "Statistique classique (régression linéaire, loi de Student)", "analyses_complements_fppn.py, onglet 8_tendances",
      "Sans texte applicable", r"^(Pente|IC bas|IC haut|Croissance log)")

    # ------------------------------------------------------------------ Partie 5 : sinistralité
    f("F-05-01", 5, "Taux de sinistralité réglementaire", "Charges de prestations, provisions comprises, divisées par les cotisations acquises, provisions sur cotisations comprises.",
      "(60 + 6901 - 7901) / (70 + 7906 - 6906)", BAL, NC,
      DEC + ", § 1.2.1.2.5 et tableau 1, ligne 5 : même formule, norme de 70 % à 80 %",
      "Formule du texte appliquée telle quelle. Les comptes 6901, 7901, 7906 et 6906 n'existent pas au FPPN : la variation de provisions vaut zéro.",
      f"Taux 2025 : {_pc(ps[2025]['taux_texte'])}", "Provision figée depuis 2020 : le taux n'intègre aucune variation de provision", "Élevé",
      DEC, "provisions_fppn.py, onglet 12_sinistralite_variantes ; ratios_airms_fppn.py", "Conforme au texte", r"^(SIN |SIN num|SIN den|SIN cot|SIN prest|SIN charges|SIN taux)")
    f("F-05-02", 5, "Sinistralité maladie par garantie", "Prestations maladie comptabilisées divisées par les cotisations maladie comptabilisées de la même année.",
      "(601 + 6013 + 6018 + 60800007) / 7010", BAL, NC,
      DEC + " : taux global seulement ; le suivi par risque découle de l'excédent par risque (§ 1.1.1)",
      "Le texte ne fixe pas de taux par garantie, mais impose de calculer l'excédent de chaque risque : le taux par garantie en découle. La norme de 70 % à 80 % ne s'applique qu'au taux global.",
      f"2025 : {_pc(ps[2025]['taux_maladie'])} par garantie contre {_pc(ps[2025]['taux_texte'])} en global", "Clé de répartition des cotisations non documentée (F-04-04)",
      "Élevé", DEC + ", § 1.1.1", "analyses_fppn.py, onglet 3_sinistralite_resultat", "Divergent, justifié", r"^(S/C maladie|S/C 20|S/C anneau|Mult S/C|Prest mal|Cot mal|Total prest)")
    f("F-05-03", 5, "Sinistralité décès par garantie", "Prestations décès (capitaux, allocations, frais funéraires) divisées par les cotisations décès.",
      "(6021 + 6024) / 7021 ; en 2020, + 65180000 (prime d'assurance décès)", BAL, NC,
      DEC + " : taux global seulement ; suivi par risque exigé par § 1.1.1",
      "Même logique que la maladie. La prime d'assurance décès de 2020 est ajoutée, car un assureur portait alors la garantie.",
      f"2025 : {_pc(ps[2025]['taux_deces'])}", "Garantie largement excédentaire, provision obsèques figée", "Moyen", DEC + ", § 1.1.1",
      "analyses_fppn.py, onglet 3_sinistralite_resultat", "Divergent, justifié", r"^(Min S/C décès|Max S/C décès|Prest dec|Cot dec)")
    f("F-05-04", 5, "Sinistralité sans MEDOC et sans bons de prise en charge", "Variantes du taux réglementaire : cotisations MEDOC retirées du dénominateur, ou bons de prise en charge et bons de caisse retirés du numérateur.",
      "(60 + 6901 - 7901) / (70 - 7022 + 7906 - 6906) ; (60 - 60800007 - 60800008 + 6901 - 7901) / (70 + 7906 - 6906)", BAL, NC,
      DEC + ", § 1.2.1.2.5 (formule de base)", "Variantes de sensibilité : elles montrent l'effet de comptes dont la nature est à préciser.",
      f"2025 : {_pc(ps[2025]['taux_sans_medoc'])} sans MEDOC ; {_pc(ps[2025]['taux_sans_bons'])} sans bons", "Écarts faibles : la conclusion ne dépend pas de ces comptes",
      "Faible", DEC, "provisions_fppn.py, onglet 12_sinistralite_variantes", "Divergent, justifié", r"^(Variante)")
    f("F-05-05", 5, "Taux de frais de gestion", "Frais de gestion (comptes 61 à 66, hors prime d'assurance décès) divisés par les cotisations maladie et décès.",
      "(61 + 62 + 63 + 64 + 65 + 66 - 6518) / (7010 + 7021)", BAL, NC,
      DEC + ", tableau 1 : charges de fonctionnement = (612, 613, 618, 62 à 68) / cotisations acquises, au plus 20 %",
      "Le texte exclut 611 et inclut 67 et 68, et prend toutes les cotisations acquises (70). Le rapport présente les deux ; seule la formule du texte est comparée à la norme.",
      f"2025 : {_pc(sin[2025]['taux_frais'])} (rapport) contre {_pc(chf25)} (texte)", "Classement des charges entre fonctionnement et prestations",
      "Moyen", DEC, "analyses_fppn.py ; ratios_airms_fppn.py (CHF)", "Divergent, justifié", r"^(Taux frais|Min frais|Max frais)")
    f("F-05-06", 5, "Résultat net", "Produits moins charges, plus le résultat hors activités ordinaires.", "classe 7 - classe 6 + classe 8", BAL,
      "Résultat des balances du FPPN", "PCMS (règlement n° 03/2012/CM/UEMOA) : résultat = produits - charges",
      "Même définition que le plan comptable.", "Résultat 2026 provisoire (8 mois)", "Écritures de fin d'année absentes en 2026", "Moyen",
      "PCMS", "analyses_fppn.py, onglet 3_sinistralite_resultat", "Conforme au texte", r"^(Résultat|Produits 2025|Charges 2025)")

    # ------------------------------------------------------------------ Partie 6 : épargne
    f("F-06-01", 6, "Dette d'épargne", "Somme des comptes PEP et EFC dus aux mutualistes.", "41195100 + 41195200 (+ 47120000 depuis 2025)", BAL,
      "Comptes de dettes du FPPN", REX + ", art. 11 (provision mathématique) et art. 12 à 15 (risques longs)",
      "Le FPPN suit l'épargne en dettes. Si l'EFC promet plus que l'épargne versée, une provision mathématique est due (règlement de l'EFC non reçu).",
      "Non mesurable sans règlement de l'EFC", "Engagement sous-estimé si l'EFC garantit un capital", "Élevé", REX, "analyses_fppn.py, onglet 6_bilan",
      ATT, r"^(Épargne|EFC 20|PEP 20|Hausse épargne|Var PEP)")
    f("F-06-02", 6, "Couverture de l'épargne", "Actifs réservés à l'épargne divisés par la dette d'épargne.", "(banques et DAT dédiés + FCP) / (41195100 + 41195200)", BAL, NC,
      DEC + ", section II : ratio de solvabilité de la branche épargne = actifs admis / engagements réglementés, au moins 1",
      "Même logique que le texte (RL1).", "Couverture très inférieure à 1", "Actifs non dédiés par écrit", "Élevé", DEC, "analyses_fppn.py ; ratios_airms_fppn.py (RL1)",
      "Conforme au texte", r"^(Couverture|Actifs dédiés|Actifs réservés|Min couverture|Max couverture|Total actifs)")
    f("F-06-03", 6, "Sensibilité aux retraits", "Part de l'épargne qui peut être retirée avant d'épuiser la trésorerie ou les actifs réservés.", "seuil = trésorerie / épargne ; couverture d'un retrait de 10 %",
      BAL, "Sans objet", REX + ", art. 15 (réserve de trésorerie d'un trimestre de prestations des risques longs)",
      "Test de résistance complémentaire au texte.", "Seuil de retrait mesuré en partie 6", "Retraits massifs non modélisés par ancienneté", "Moyen",
      "Pratique de gestion actif-passif (test de résistance)", "analyses_complements_fppn.py", "Sans texte applicable", r"^(Retrait|Seuil tréso)")

    # ------------------------------------------------------------------ Partie 7 : Grenier
    f("F-07-01", 7, "Portefeuille en retard (PAR30, PAR90)", "Encours des prêts en retard de plus de 30 (ou 90) jours divisé par l'encours total.",
      "Σ encours(retard > 30 j) / Σ encours", GRE, NC,
      "Aucun texte de la mutualité. Référence technique : indicateurs de portefeuille à risque des systèmes financiers décentralisés (BCEAO)",
      "Indicateur standard du crédit, faute de texte de la mutualité.", f"PAR30 de {_pc(R['grenier_par']['par30'])} au 31/08/2026", "Retard mesuré sur l'export du logiciel, pas sur les précomptes",
      "Moyen", "BCEAO, instruction applicable aux SFD (texte à transmettre)", "analyses_fppn.py, onglet 5_grenier_retards", "Sans texte applicable", r"^(PAR30|Prêts retard|Part prêts retard)")
    f("F-07-02", 7, "Intervalle de Wilson", "Encadrement à 95 % de la proportion de prêts en retard.",
      "(p + z²/2n ± z √(p(1-p)/n + z²/4n²)) / (1 + z²/n), z = 1,96", GRE, "Sans objet", "Aucun texte applicable",
      "Méthode adaptée aux proportions faibles.", "Intervalle publié en partie 7", "Indépendance des prêts supposée", "Faible",
      "Statistique classique (Wilson, 1927)", "analyses_complements_fppn.py, onglet 8_ic_grenier", "Sans texte applicable", r"^(IC bas retard|IC haut retard)")
    f("F-07-03", 7, "Dépréciation des prêts en retard", "Encours de chaque tranche de retard multiplié par le taux de dépréciation de la tranche, puis additionné.",
      "Σ encours(tranche k) × taux(k)", GRE, "Aucune dépréciation passée depuis 2022 (compte 49 non mouvementé)",
      "PCMS et SYSCOHADA révisé : principe de prudence. Taux par tranche : barème BCEAO des SFD (texte à transmettre)",
      "La dépréciation réduit l'actif ; elle n'est pas une provision technique (" + REX + ", art. 8). Calcul en attente du barème.",
      "Encours de plus de 90 jours connu par tranche (partie 9)", "Actif et fonds propres surévalués sans dépréciation", "Moyen",
      "PCMS ; BCEAO (instruction SFD, à transmettre)", "provisions_fppn.py, onglet 12_grenier_tranches", ATT, r"^(Encours 180|Encours 90|Perte|Prov180|Prov90|Grenier 90)")
    f("F-07-04", 7, "Précomptes rapportés aux échéances", "Montant précompté au poste 9X4 divisé par les échéances dues des prêts en cours.", "précompte 9X4 du mois / Σ échéances du mois",
      PREC + " ; " + GRE, NC, "Aucun texte applicable", "Contrôle de cohérence entre deux sources.", "76 à 85 % des échéances précomptées",
      "Paiements hors précompte non connus", "Moyen", "Rapprochement de sources (pratique d'audit)", "analyses_fppn.py, onglet 5_grenier_echeances",
      "Sans texte applicable", r"^(Écart min|Écart max|Min précompte|Max précompte)")

    # ------------------------------------------------------------------ Partie 9 : bilan et provisions
    f("F-09-01", 9, "PSAP en mois de prestations", "Provision pour sinistres à payer divisée par les prestations mensuelles moyennes de l'année.",
      "PSAP / (prestations de l'année / nombre de mois)", BAL, "Aucune provision maladie calculée ; montants hérités d'environ 2004, jamais recalculés (chef comptable, 19/10/2026, à confirmer)",
      REX + ", art. 10 : PSAP calculée exercice par exercice, dossier par dossier, plus les sinistres non déclarés",
      "Indicateur de cohérence, pas une norme : une provision figée ne suit pas les prestations.", f"Maladie : {_nb(pr[2025]['psap_en_mois_de_prestations'], 2)} mois de prestations en 2025 ; décès : {_nb(pr[2025]['psap_deces_en_mois'], 2)} mois",
      "Provision ni suffisante ni excessive démontrable sans base de prestations", "Élevé", REX, "provisions_fppn.py, onglet 12_psap_rapprochement",
      ATT, r"^(Mois 20|Prov tot|Prest mens|PSAP)")
    f("F-09-02", 9, "PSAP et dettes envers les cliniques", "Provision maladie comparée aux factures des cliniques déjà enregistrées et non payées.",
      "PSAP maladie / 40110000", BAL, "Aucune provision maladie calculée ; montants hérités d'environ 2004, jamais recalculés (chef comptable, 19/10/2026, à confirmer)",
      REX + ", art. 10 (dossiers connus + tardifs) ; PCMS (dettes fournisseurs en classe 4)",
      "Une facture enregistrée en dette n'a plus à être provisionnée : la PSAP doit couvrir les soins survenus et non encore facturés.",
      "Rapport mesuré chaque année (partie 9)", "Double compte possible entre dettes et provision", "Élevé", REX + " ; PCMS",
      "provisions_fppn.py, onglet 12_psap_rapprochement", ATT, r"^(Dettes cliniques|Rapport PSAP)")
    f("F-09-03", 9, "Sinistres tardifs (méthode CIMA)", "Coefficients de déclaration tardive observés, multipliés par les sinistres déclarés l'année de survenance et par le coût moyen.",
      "IBNR = Σ_n coef(n) × N_déclarés(n, n) × coût moyen(n) ; coef = Σ_k moyenne[N(n, n+k) / N(n, n)]", "Bases de prestations avec dates de survenance et de déclaration (non reçues)",
      NC, REX + ", art. 10 (estimation des sinistres non déclarés, sans méthode fixée) ; Code CIMA (méthode de référence, article à citer)",
      "Le texte UEMOA impose l'estimation sans dire comment : la méthode CIMA sert de référence technique.", "Non calculable sans les bases", "Peu d'années d'historique",
      "Élevé", "Code CIMA (référence technique, texte à transmettre)", "À produire à réception des bases", ATT, r"^(IBNR)")
    f("F-09-04", 9, "Provision mathématique", "Valeur actuelle des prestations promises moins valeur actuelle des cotisations futures.",
      "PM = Σ_t prestation(t) × p(t) × v^t - Σ_t cotisation(t) × p(t) × v^t ; v = 1 / (1 + i)", "Règlements de l'EFC et de la santé des retraités, âges (non reçus)",
      "Aucune provision mathématique au bilan", REX + ", art. 11 : table de mortalité, taux d'intérêt technique, chargements de gestion",
      "Due si l'EFC garantit un capital ou si la santé est garantie à vie aux retraités.", "Non calculable sans règlements et âges", "Engagement long absent du bilan",
      "Élevé", REX + ", art. 11 ; Code CIMA, art. 338 (tables, taux de 3,5 %)", "À produire à réception des documents", ATT, r"^(Provision mathématique)")
    f("F-09-05", 9, "Provision pour risques en cours", "Coût des risques entre la date d'inventaire et la prochaine échéance de cotisation.", "PREC = coût estimé des risques de la période non couverte par les cotisations acquises",
      BAL, "Aucune provision pour risques en cours", REX + ", art. 9", "Cotisations mensuelles précomptées : la période non couverte est au plus d'un mois. Provision probablement faible.",
      "Non chiffrée", "Faible tant que la cotisation reste mensuelle", "Faible", REX + ", art. 9", "À apprécier", ATT)
    f("F-09-06", 9, "Indemnités de départ à la retraite du personnel", "Somme, pour chaque salarié, de l'indemnité prévue par la convention collective à son départ, pondérée par la probabilité d'être présent et actualisée, au prorata de l'ancienneté acquise.",
      "PIDR = Σ_i IDR_i × p_i × v^(t_i) × ancienneté acquise_i / ancienneté au départ_i ; IDR = salaire × (30 % par année jusqu'à 5 ans, 35 % de 6 à 10, 40 % au-delà) ; charges patronales : fiscales 2,80 %, retraite 7,70 % (plafond 3 375 000 FCFA), prestations familiales 5,75 % et risques professionnels 2 % (plafond 75 000 FCFA)",
      "Fichier du personnel, convention collective, note de calcul (non reçus)", "Convention collective appliquée à l'ancienneté acquise ; provision reprise et dotée en totalité chaque année (chef comptable, 19/10/2026, à confirmer)",
      "SYSCOHADA révisé (engagements de retraite, méthode actuarielle) ; convention collective interprofessionnelle de Côte d'Ivoire (barème, à confirmer par le texte)",
      "Montant du FPPN à recalculer selon la méthode des unités de crédit projetées.", "Provision 2025 : " + _f(P["pidr"][5]["provision_196"]), "Méthode et hypothèses inconnues",
      "Moyen", "SYSCOHADA révisé ; convention collective interprofessionnelle de Côte d'Ivoire", "provisions_fppn.py, onglet 12_pidr_196", ATT, r"^(PIDR)")
    f("F-09-10", 9, "Allocations décès restant à payer (méthode du FPPN)", "Décès de l'année dont l'allocation n'est pas payée, valorisés au barème par grade.",
      "Σ dossiers non payés × allocation du grade (sous-officier 1 000 000 FCFA, officier à confirmer, commissaire 2 000 000 FCFA)", "Liste des décès de la direction du personnel ; allocations payées",
      "Liste des décès moins allocations payées, valorisée au barème (chef comptable, 19/10/2026, à confirmer) ; compte de comptabilisation non précisé",
      REX + ", art. 10 : dossier par dossier, plus les sinistres non déclarés",
      "Méthode proche du dossier par dossier, sans les sinistres tardifs (décès d'ayants droit non déclarés). Le compte utilisé est à préciser : le compte 15802000 ne bouge pas.",
      "Non mesurable sans les fichiers des décès (PBC, item 92)", "Allocations restant à payer peut-être non comptabilisées", "Élevé", REX + ", art. 10",
      "À produire à réception des fichiers", ATT, r"^(Allocations décès)")
    f("F-09-07", 9, "Adossement par échéance", "Actif mobilisable dans chaque tranche de délai comparé au passif exigible dans la même tranche.",
      "Écart(tranche) = actif(tranche) - passif(tranche)", BAL, NC, DEC + " (liquidité, équilibre stable) ; " + REX + ", art. 15",
      "Lecture par échéance plus fine que les ratios globaux.", "Écart à moins d'un an mesuré en partie 9", "Classement des comptes par jugement", "Moyen",
      "Pratique de gestion actif-passif", "ratios_airms_fppn.py, onglet 10_mobilisation", "Sans texte applicable", r"^(Actif < 1 an|Passif < 1 an|Actif > 5 ans|Écart < 1 an|Actif total mobilisation|Épargne mobilisation)")
    f("F-09-08", 9, "Fonds propres", "Réserves, report à nouveau, pertes reportées et résultat de l'exercice.", "10 + 11 + 12 + 13 + 14 + résultat", BAL,
      "Fonds propres des balances", "PCMS ; " + DEC + ", § 1.1.1 (marge de solvabilité)", "Même définition que le plan comptable.", "Fonds propres négatifs de 2020 à 2026",
      "Bilan reconstitué (non transmis)", "Élevé", "PCMS", "analyses_fppn.py, onglet 6_bilan", "Conforme au texte", r"^(FP |FP 20|FP hors|Sous-total FP|Fonds propres)")
    f("F-09-09", 9, "Rendement des placements", "Produits financiers divisés par les placements moyens de l'année.", "77 / moyenne(placements début, placements fin)", BAL, NC,
      DEC + ", section II : taux brut de rentabilité = intérêts perçus / capital investi", "Même logique que le texte (RL6).", "Rendement moyen publié en partie 9",
      "Placements moyens estimés sur deux dates", "Faible", DEC, "analyses_fppn.py, onglet 7_placements", "Conforme au texte", r"^(Rendement moyen|Produits fin|Placements moyens)")

    # ------------------------------------------------------------------ Partie 10 : ratios (repris de ratios.json)
    for k, x in enumerate(Q["ratios"], 1):
        sg = x["statut_global"]
        attente = sg == "En attente"
        L.append({"id": f"F-10-{k:02d}", "partie": 10, "indicateur": f"{x['nom']} ({x['code']})", "litterale": x["formule"] + ".",
                  "detaillee": x.get("comptes") or "", "donnees": BAL if not attente else "Documents demandés (PBC)",
                  "fppn": NC, "reglementaire": f"{x['source']} : {x['formule']}, norme : {x['norme']}",
                  "ecart": ("Formule du texte appliquée telle quelle. " if not attente else "Calcul impossible sans les documents demandés. ")
                  + "Les états éventuellement transmis par le FPPN à l'AIRMS permettront la comparaison.",
                  "impact": f"Statut 2022 à 2025 : {sg}", "risques": x.get("qualitatif") or "Balances 2026 provisoires ; classement de certains comptes par jugement",
                  "niveau": "Élevé" if sg == "Non conforme" and x["famille"].startswith("Actuariel") else ("Moyen" if sg in ("Non conforme", "En attente") else "Faible"),
                  "source": x["source"], "calcul": "ratios_airms_fppn.py, onglet 10_ratios", "statut": ATT if attente else ("À apprécier" if sg == "À apprécier" else "Conforme au texte"),
                  "traces": rf"^{x['code']} "})
    f("F-10-90", 10, "Excédent par risque", "Cotisations du risque moins ses prestations et sa part des frais, répartie au prorata des cotisations.",
      "Excédent(r) = cotisations(r) - prestations(r) - frais × cotisations(r) / cotisations totales", BAL, NC,
      DEC + ", § 1.1.1 : réserve obligatoire dotée d'au moins 30 % de l'excédent maladie et 50 % des autres risques",
      "Le texte ne fixe pas la clé des frais : Adenka retient le prorata des cotisations, à défaut de clé du FPPN.", "Excédents 2025 publiés en partie 10",
      "Dépend de la clé maladie et décès (F-04-04)", "Élevé", DEC + ", § 1.1.1", "ratios_airms_fppn.py, onglet 10_excedent_par_risque", "Divergent, justifié", r"^(Exc )")
    f("F-10-91", 10, "Affectation de l'excédent", "Dotation à la réserve obligatoire comparée au minimum exigé.", "écart = dotation - max(20 % excédent ; 30 % excédent maladie + 50 % excédent autres)",
      "Balances (compte 111)", NC, DEC + ", § 1.1.1 ; règlement n° 07/2009/CM/UEMOA", "Formule du texte.", "Écarts 2023 à 2025 publiés en partie 10",
      "PV d'AG non reçus", "Élevé", DEC, "ratios_airms_fppn.py, onglet 10_affectation", "Conforme au texte", r"^(AFF )")

    # ------------------------------------------------------------------ Partie 11 : mortalité
    f("F-11-01", 11, "Espérance de vie", "Nombre moyen d'années restant à vivre à l'âge x selon la table.", "e_x = Σ_(k≥1) L_(x+k) / L_x + 0,5", "Tables CIMA H, F, TD, TV",
      "Sans objet", "Code CIMA, art. 338 (tables de référence)", "Indicateur de lecture des tables.", "Valeurs publiées en partie 11", "Table non validée sur la population du FPPN",
      "Faible", "Code CIMA, art. 338", "analyses_complements_fppn.py, onglet 8_mortalite", "Conforme au texte", r"^(e0|e55|e60|Survie60)")
    f("F-11-02", 11, "Valeur d'une rente viagère", "Somme actualisée des paiements de 1 par an tant que la personne est en vie.", "ä_x = Σ_k (L_(x+k) / L_x) × (1,035)^(-k)", "Tables CIMA",
      "Sans objet", "Code CIMA, art. 338 : taux d'intérêt technique au plus 3,5 %", "Base du calcul de la provision mathématique.", "Valeurs publiées en partie 11",
      "Taux de 3,5 % à justifier pour une mutuelle sociale", "Moyen", "Code CIMA, art. 338", "analyses_complements_fppn.py, onglet 8_mortalite", "Conforme au texte", r"^(a5|a6|ä)")
    f("F-11-03", 11, "Probabilité de décès", "Probabilité de mourir entre x et x + 1 ans.", "q_x = 1 - L_(x+1) / L_x", "Tables CIMA", "Sans objet", "Code CIMA, art. 338",
      "Base de la mortalité observée (A/E).", "Valeurs publiées en partie 11", "A/E en attente des âges", "Faible", "Code CIMA, art. 338",
      "analyses_complements_fppn.py, onglet 8_mortalite", "Conforme au texte", r"^(q40|q55|H vs TD|F vs TV)")
    return L


COLONNES = [("id", "Identifiant", 9), ("diapositives", "Diapositives", 12), ("indicateur", "Indicateur", 26), ("litterale", "Formule littérale", 46),
            ("detaillee", "Formule détaillée (comptes, colonnes)", 40), ("donnees", "Données sources", 28), ("fppn", "Méthode du FPPN", 34),
            ("reglementaire", "Méthode réglementaire (texte, article, formule, norme)", 46), ("ecart", "Analyse de l'écart et faisabilité", 46),
            ("impact", "Impact chiffré", 28), ("risques", "Risques et limites", 32), ("niveau", "Niveau de risque", 11),
            ("source", "Source de la méthode", 32), ("calcul", "Calcul (script, onglet)", 30), ("statut", "Statut", 16)]
