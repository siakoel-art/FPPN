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
