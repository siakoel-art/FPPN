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
