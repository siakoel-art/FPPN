# ---- v13 (13/11/2026) : résumé exécutif en 8 diapositives (REGLES_REDACTION_FPPN.md, structure validée le 13/11/2026)
# 1 messages, 2 le FPPN en bref, 3 réponses aux quatre questions, 4 pourquoi, 5 maladie et décès, 6 dépenses hors soins,
# 7 règles du régulateur, 8 que faire et quand. Chiffres : chiffres_v13.json (K13) et données du générateur (base v13).
B26 = "« BALANCE FPPN DU 01-01-2026 AU 31-08-2026.xlsx »"
B25 = "« BALANCE FPPN 2025.xlsx »"
GL25 = "« GRAND LIVRE DES COMPTES FPPN 2025.xlsx »"
GL26 = "« GRAND LIVRE DES COMPTES FPPN DU 01-01-2026 AU 31-08-2026.xlsx »"
_pp = lambda x: n13(x, 2)  # noqa: E731  « pour 100 » à deux décimales
_ff = lambda x: n13(x, 0) + " FCFA"  # noqa: E731
_ep, _ban = K13["epargne"], 2562885661
_mal25 = s[2025]["prest_maladie"] / s[2025]["cot_maladie"]
_dec25 = s[2025]["prest_deces"] / s[2025]["cot_deces"]
_comb25 = (1284143358 + 875420236 + 0.705 * 487492003) / 2409281030
_mob = [("Disponible tout de suite (banques, caisse)", 2373558814), ("Disponible dans l'année", 1084308137),
        ("Disponible en un à cinq ans", 1297315010), ("Plus de cinq ans ou incertain (clinique, terrains, créances)", 17170238905)]
_tot = sum(v for _, v in _mob)

# 1. Les cinq messages
contenu(blt13({"titre": "L'épargne n'est pas gardée à part et la maladie dépense plus qu'elle reçoit",
               "sous_titre": fk13("Au même rythme qu'à fin août, il manquera {res_tech_26_dec_abs:f} FCFA à fin décembre 2026 pour payer les soins, les allocations et les frais"),
               "synthese": "Cinq messages, chacun avec son chiffre et la diapositive qui le démontre"}),
        "Diapositive à lire seule : elle résume le rapport en cinq faits chiffrés.",
        tableau={"titre": tab("Les cinq messages à retenir"), "entetes": ["Message", "Chiffre qui le prouve", "Détail"], "largeurs": [10.4, 15.6, 5.2], "taille": 12,
                 "lignes": [["Le Fonds garde peu d'épargne en banque", f"{_pp(100 * _ban / _ep)} FCFA en banque pour 100 FCFA d'épargne reçus, soit {_pp(100 * _ban / _ep)} %", ref("Calcul pas à pas : couverture de l'épargne")],
                            ["Le Fonds paie 3,5 % une seule fois, en le prenant sur l'épargne", f"{_ff(K13['interets_payes_pep'])} d'intérêts payés de 09/2020 à 08/2026, sans passer en charges", ref("Intérêts servis : aucune écriture dans les comptes")],
                            ["Le Fonds ne calcule pas ses provisions", f"{_ff(K13['provisions'])} figés depuis 2020, pour un besoin estimé par Adenka à au moins {_ff(K13['besoin'])}", ref("Calcul pas à pas : provisions")],
                            ["La maladie dépense plus qu'elle ne reçoit", f"En 2025, {_pp(100 * _comb25)} FCFA de soins et de frais pour 100 FCFA de cotisations maladie, soit {_pp(100 * _comb25)} %", ref("Sinistralité des soins seuls")],
                            ["2026 est déjà déficitaire", fk13("{res_tech_26:f} FCFA de résultat technique à fin août, après extourne de la provision de 2025 oubliée"), ref("Projection agrégée : deux scénarios")]]},
        disposition="dessous",
        analyse=["**Le Fonds sait ce qu'il doit, mais il ne garde pas à part l'argent de l'épargne et ses garanties de base ne s'équilibrent plus.**"],
        implication="Les trois actions de la diapositive « Que faire et quand » répondent à ces cinq messages.",
        valable="comptes 2020 au 31/08/2026 (2026 provisoire, extourne de la provision de 2025 ajoutée par Adenka) ; documents reçus jusqu'au 11/11/2026",
        source=f"Sources : (1) {B26}, comptes 41195100, 41195200, 52 et 53. (2) Grands livres 2020 à 2026, compte 41195100, calcul Adenka. (3) {B26}, comptes 15801000 et 15802000. (4) {GL25}, comptes 6011 à 6080. (5) {B26} et {GL25}, OD 805 à 807")

# 2. Le FPPN en bref
contenu(blt13({"titre": fk13("Le Fonds garde {epargne:f} FCFA d'épargne pour les policiers"),
               "sous_titre": fk13("L'épargne est la plus grosse dette du Fonds : {epargne_sur_cotisations:p} fois les cotisations de 2025 ({cotisations_2025:f} FCFA)"),
               "synthese": "Les chiffres qui situent le Fonds au 31/08/2026"}),
        "Ordres de grandeur du Fonds, pour situer chaque constat.",
        tableau={"titre": tab("Le Fonds en chiffres"), "entetes": ["Indicateur", "Valeur", "Source"], "largeurs": [12.0, 8.0, 11.2], "taille": 12,
                 "lignes": [["Policiers cotisants (août 2026)", "31 198", "(1)"],
                            ["Retraités cotisants (août 2026)", N("R13 retraités", CG[-1]["retraites"]), "(2)"],
                            ["Cotisations maladie et décès 2025", _ff(K13["cotisations_2025"]), "(3)"],
                            ["Prestations maladie et décès 2025", _ff(s[2025]["prest_maladie"] + s[2025]["prest_deces"]), "(3)"],
                            ["Épargne des policiers au 31/08/2026", _ff(_ep), "(4)"],
                            ["Prêts du Grenier au 31/08/2026", _ff(GP["encours"]), "(5)"],
                            ["Total des biens du Fonds au 31/08/2026", _ff(_tot), "(4)"]]},
        disposition="dessous",
        analyse=[f"**Pour 100 FCFA de biens, le Fonds doit {_pp(100 * _ep / _tot)} FCFA d'épargne aux policiers.** C'est ce qui fait de l'épargne le sujet central du rapport."],
        valable="effectifs d'août 2026 ; comptes au 31/12/2025 et au 31/08/2026 (provisoire)",
        source=f"Sources : (1) fichiers retour de la Solde, poste 936, août 2026. (2) fichiers retour de la caisse de retraite (CGRAE), août 2026. (3) {B25}, comptes 7010, 7021, 601, 60800007, 6021 et 6024. (4) {B26}, comptes 41195100 et 41195200, classes 2 à 5. (5) balance âgée du Grenier au 31/08/2026")

# 3. Réponses aux quatre questions
contenu(blt13({"titre": "Sur les quatre questions de la mission, deux réponses sont négatives",
               "sous_titre": "Le Fonds sait ce qu'il doit, mais il n'a ni calculé ses provisions ni gardé l'argent pour payer à temps",
               "synthese": "Une phrase simple et un chiffre par question"}),
        "Réponse aux quatre questions de la mission, du 01/09/2020 au 31/08/2026.",
        tableau={"titre": tab("Réponse aux quatre questions de la mission"), "entetes": ["Question", "Réponse", "En une phrase, avec son chiffre"], "largeurs": [7.0, 4.4, 19.8], "taille": 12, "statut_col": 1,
                 "lignes": [["1. Ce que le Fonds doit à ses membres", "Partiellement démontré", f"Le Fonds doit {_ff(_ep)} d'épargne. S'y ajoutent les intérêts du plan d'épargne : 3,5 FCFA pour 100 FCFA versés, une seule fois, soit 3,5 % du montant versé, {_ff(K13['interets_pep_dus'])}. Règle de l'entraide : [[à compléter]]."],
                            ["2. Les provisions suffisent-elles", "Partiellement démontré", f"Le Fonds ne calcule pas ses provisions : {_ff(K13['provisions'])} figés depuis 2020. Adenka estime le besoin à au moins {_ff(K13['besoin'])}."],
                            ["3. L'argent arrive-t-il à temps pour payer", "Insuffisant", fk13("Pour 100 FCFA à payer dans l'année, le Fonds dispose de {couv_besoin:p} FCFA, soit {couv_besoin:p} % des paiements.")],
                            ["4. Les placements conviennent-ils", "Insuffisant", f"Pour 100 FCFA placés, ils ont rapporté 4,90 FCFA par an de 2021 à 2025, soit 4,90 %, sous les obligations de l'État (5,19 % à 7,73 %)."]]},
        disposition="dessous",
        analyse=["**La même cause explique trois réponses : le Fonds n'a ni taux d'intérêt écrit, ni calendrier de remboursement, ni argent mis de côté pour l'épargne.**"],
        valable="période du 01/09/2020 au 31/08/2026, documents reçus jusqu'au 11/11/2026",
        source=f"Sources : (1) {B26}, comptes 41195100 et 41195200. (2) entretiens avec le Fonds, [[service et date à compléter]]. (3) {B26}, comptes 15801000 et 15802000. (4) classeur de calcul, onglets 24b et 24c. (5) grands livres 2021 à 2025, comptes 5221, 5330 et 77")

# 4. Pourquoi : épargne, argent disponible, biens comptés trop cher
contenu(blt13({"titre": f"Pour 100 FCFA d'épargne reçus, le Fonds n'en a gardé que {_pp(100 * _ban / _ep)} en banque",
               "sous_titre": f"Les {_pp(100 - 100 * _ban / _ep)} FCFA restants, soit {_ff(_ep - _ban)}, sont dans des terrains, des immeubles et des prêts : si beaucoup de policiers retirent leur épargne en même temps, le Fonds ne peut pas les payer tout de suite",
               "synthese": "Où se trouve l'argent du Fonds, et quels biens sont comptés trop cher"}),
        "Les biens du Fonds classés selon le temps qu'il faut pour en faire de l'argent.",
        tableau={"titre": tab("Où se trouve l'argent du Fonds au 31/08/2026"), "entetes": ["Biens", "FCFA", "Pour 100 FCFA de biens"], "largeurs": [15.0, 8.2, 8.0], "taille": 12,
                 "lignes": [[a, n13(v), _pp(100 * v / _tot)] for a, v in _mob] + [["Total", n13(_tot), "100,00"]]},
        disposition="dessous",
        analyse=[f"**Des biens sont comptés trop cher :** deux terrains sans titre et la créance sur l'hôpital pèsent {_ff(K13['actifs_contestes'])}. Sans eux, les fonds propres tombent à {_ff(K13['fp_sans_contestes'])}. [[La réévaluation du siège (2 899 400 000 FCFA) compte peut-être deux fois le terrain.]]",
                 fk13("Hors épargne, les dettes de l'année sont couvertes à {couv_hors_epargne:p} FCFA pour 100 FCFA : c'est l'épargne que le Fonds ne peut pas rendre à temps.")],
        valable="bilan au 31/08/2026 (provisoire), biens à leur valeur comptable",
        source=f"Sources : (1) {B26}, classes 2 à 5, classées par délai (classeur, onglet 10_actif_mobilisation). (2) mémo du service comptable sur les comptes 47 et 27 (11/11/2026). (3) rapport d'expertise du siège du 19/03/2019, compte 10620000")

# 5. Maladie et décès
_ans = (2023, 2024, 2025, 2026)
contenu(blt13({"titre": f"En 2025, la maladie a dépensé {_pp(100 * _comb25)} FCFA pour 100 FCFA de cotisations",
               "sous_titre": f"Le décès, qui ne dépense que {_pp(100 * _dec25)} FCFA pour 100 FCFA reçus, paie la différence",
               "synthese": "Prestations pour 100 FCFA de cotisations, par garantie (2026 : huit mois, extourne de la provision de 2025 ajoutée)"}),
        "Sinistralité par garantie et sinistralité réglementaire, de 2023 au 31/08/2026.",
        tableau={"titre": tab("Prestations pour 100 FCFA de cotisations"), "entetes": ["Garantie"] + [str(a) if a < 2026 else "31/08/2026" for a in _ans], "largeurs": [11.2, 5.0, 5.0, 5.0, 5.0], "taille": 12,
                 "lignes": [["Maladie (prestations seules)"] + [_pp(100 * s[a]["prest_maladie"] / s[a]["cot_maladie"]) for a in _ans],
                            ["Décès"] + [_pp(100 * s[a]["prest_deces"] / s[a]["cot_deces"]) for a in _ans],
                            ["Ensemble, selon le texte (norme 70 à 80)"] + [_pp(100 * _rv13("SIN", a)) for a in _ans]]},
        disposition="dessous",
        analyse=[f"**En 2025, la maladie a dépensé {_pp(100 * _comb25)} FCFA de soins et de frais pour 100 FCFA de cotisations, soit {_pp(100 * _comb25)} %, une fois les dépenses hors soins rangées en frais.**",
                 f"Le 31/12/2025, le Fonds a provisionné 255 000 000 FCFA de factures de la clinique (OD 805 à 807) sans les extourner en 2026 : les vraies factures (240 471 119 FCFA) ont été comptées une seconde fois. Après extourne, le taux réglementaire de 2026 est de {_pp(K13['regl_26'])} FCFA pour 100 FCFA, sous la norme."],
        valable="comptes 2023 au 31/08/2026 ; 2026 provisoire (huit mois), extourne ajoutée par Adenka",
        source=f"Sources : (1) {B25} et {B26}, comptes 7010, 7021, 601, 60800007, 6021, 6024. (2) {GL25}, OD 805, 806 et 807, compte 60120000. (3) {GL26}, pièces SIN 462, 552 et 619. (4) classeur, onglet 35_sinistralite")

# 6. Dépenses hors soins
contenu(blt13({"titre": "En 2025, 487 492 003 FCFA payés comme des soins ont servi à autre chose",
               "sous_titre": "Ces dons, gadgets et honoraires représentent 27,52 FCFA sur 100 FCFA de prestations maladie, soit 27,52 % : sans eux, la maladie dépenserait moins qu'elle ne reçoit",
               "synthese": "[[« Paiement non justifié » seulement si aucune pièce n'est reçue pour ces sorties]]"}),
        "Dépenses enregistrées comme prestations maladie qui ne sont pas des soins, classées d'après le libellé de chaque écriture.",
        tableau={"titre": tab("Dépenses hors soins dans les prestations maladie de 2025"), "entetes": ["Nature d'après le libellé", "FCFA"], "largeurs": [20.0, 11.2], "taille": 12,
                 "lignes": [["Dons (autorités, élus, délégués)", "203 689 140"], ["Gadgets, imprimés et cartes", "167 972 000"],
                            ["Honoraires d'un cabinet (factures trimestrielles)", "115 830 863"], ["Total classé hors soins", "487 492 003"],
                            ["Libellés oubliés par le classement (imprimeurs, cérémonies, hadj, lobbying)", "au moins 55 067 300"]]},
        disposition="dessous",
        analyse=["**Les comptes de soins servent aussi à payer des dons, des gadgets et des honoraires.** Jusqu'en 2024, les dons passaient par le compte 6583 (52 270 446 FCFA en 2021). Il est vide en 2025.",
                 "Les 1 280 300 860 FCFA de charges de la clinique encore inscrits à l'actif s'ajoutent à ces sorties : ils grossissent les biens du Fonds au lieu de ses charges."],
        valable="grand livre 2025, comptes 6011 à 6080, classement par libellé",
        source=f"Sources : (1) {GL25}, comptes 6011 à 6080, classement Adenka (classeur, onglet 35_sinistralite). (2) {B25}, compte 65830000. (3) mémo du service comptable sur les comptes 47 et 27 (11/11/2026)")

# 7. Règles du régulateur
_lig7 = [[NOMS_CTRL[c], RT[c]["norme"], val_ratio(RT[c], 2024), val_ratio(RT[c], 2025)] for c in NC_T1 + NC_RL]
contenu(blt13({"titre": f"Le Fonds ne respecte pas {len(NC_T1) + len(NC_RL)} règles de contrôle du régulateur deux ans de suite",
               "sous_titre": f"{len(NC_T1)} ratios du tableau 1 et {len(NC_RL)} ratios des risques longs : pour ces règles, le régulateur peut sanctionner le Fonds",
               "synthese": "Seuls les ratios de contrôle entraînent des sanctions (Décision n° 05/2022, tableau 1 et section II)"}),
        "Ratios de contrôle hors norme en 2024 et en 2025, lecture des comptes tels quels.",
        tableau={"titre": tab("Ratios de contrôle hors norme en 2024 et 2025"), "entetes": ["Ratio", "Norme", "2024", "2025"], "largeurs": [12.6, 8.6, 5.0, 5.0], "taille": 12,
                 "lignes": _lig7},
        disposition="dessous",
        analyse=["**Le non-respect des ratios de contrôle « entraîne des sanctions administratives » et l'article 84 du règlement n° 07/2009 permet un retrait d'agrément après deux exercices.**"],
        valable="ratios recalculés sur les balances 2024 et 2025",
        source="Sources : (1) Décision n° 05/2022/COM/UEMOA, tableau 1 et section II. (2) règlement n° 07/2009/CM/UEMOA, art. 84. (3) balances 2024 et 2025, ratios recalculés par Adenka (classeur, onglet 10_ratios)")

# 8. Que faire et quand
contenu(blt13({"titre": "Trois actions dans les trois mois protègent l'épargne des policiers",
               "sous_titre": "Calculer les provisions, fixer par écrit la règle des intérêts de l'épargne et sortir des comptes de soins ce qui n'en est pas, avant fin février 2027",
               "synthese": "Actions à trois mois, puis feuille de route"}),
        "Actions proposées par Adenka, rattachées aux constats du rapport.",
        tableau={"titre": tab("Que faire et quand"), "entetes": ["Action", "Qui", "Quand", "Effet attendu"], "largeurs": [12.4, 6.0, 4.2, 8.6], "taille": 12,
                 "lignes": [["1. Calculer chaque année les provisions maladie, décès et personnel, et extourner la provision de 2025", "Direction financière", "28/02/2027", "Provisions au besoin réel, résultat 2026 juste"],
                            ["2. Fixer par écrit la règle des intérêts du plan d'épargne et de l'entraide, et les passer en charges", "Conseil d'administration", "28/02/2027", "Fin de l'écart avec le manuel, dette d'épargne exacte"],
                            ["3. Sortir des comptes de soins les dons, gadgets et honoraires", "Direction financière", "31/01/2027", "Sinistralité fiable"],
                            ["4. Mettre à part une réserve pour rendre l'épargne", "Direction générale", "Après trois mois", "Retraits payés à temps"],
                            ["5. Écrire une politique de placement et créer un comité", "Conseil d'administration", "Après trois mois", "Placements suivis"],
                            ["6. Transmettre les dates de naissance et la base des prestations", "Direction générale", "Après trois mois", "Calculs par âge possibles"]]},
        disposition="dessous",
        analyse=["**Les trois premières actions ne coûtent rien et rendent les comptes justes : elles conditionnent toutes les autres.**"],
        valable="actions proposées au 13/11/2026",
        source="Source : constats du rapport, actions proposées par Adenka")
