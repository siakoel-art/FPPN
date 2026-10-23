# ---- v11 : placements lus dans les grands livres (encours moyen pondéré par les jours), tous les dépôts à terme et le FCP
_PLG = V11["placements_gl"]
_DPEP = V11["dat_pep"]
_MOIS24 = _PLG["2024"]["dat_efc_fin_de_mois"]
contenu({"titre": "Placements du FPPN", "sous_titre": "Un dépôt mensuel pour payer l'EFC, un dépôt PEP de cinq ans, un FCP depuis 2025",
         "synthese": f"Encours moyen réellement placé en 2025 : {F('PL moy 2025', _PLG['2025']['placements_moyens'])} ; au 31/08/2026 : DAT BOA {F('PL dat 2026', _PLG['2026']['placements_cloture'] - FCP_)} et FCP {F('FCP 2025', FCP_)}"},
        "Les balances ne donnent que le solde de clôture. Le grand livre montre chaque mouvement des dépôts pendant l'année : le tableau donne l'encours moyen, pondéré par le nombre de jours où chaque solde est resté placé, et le plus haut niveau atteint par le DAT EFC.",
        tableau={"titre": tab("Encours des placements lus dans les grands livres (DAT et FCP)"), "entetes": ["Exercice", "DAT EFC : encours moyen", "DAT EFC : plus haut", "DAT PEP : encours moyen", "DAT Orabank et FCP : encours moyen", "Total : encours moyen", "Total à la clôture"],
                 "largeurs": [3.0, 4.6, 4.4, 4.6, 5.2, 4.6, 4.6], "taille": 9,
                 "lignes": [[a if a != "2026" else "2026 (8 mois)", F(f"PL efc {a}", _PLG[a]["dat_efc_moyen"]), F(f"PL pic {a}", _PLG[a]["dat_efc_pic"]), F(f"PL pep {a}", _PLG[a]["dat_pep_moyen"]),
                             F(f"PL autres {a}", _PLG[a]["dat_orabank_moyen"] + _PLG[a]["fcp_moyen"]), F(f"PL tot {a}", _PLG[a]["placements_moyens"]), F(f"PL clo {a}", _PLG[a]["placements_cloture"])]
                            for a in ("2021", "2022", "2023", "2024", "2025", "2026")]},
        disposition="dessous",
        analyse=[f"**Le DAT EFC (compte 52210200) est un dépôt d'un mois, renouvelé chaque mois et augmenté de l'épargne EFC précomptée.** En 2024, il passe de {F('PL fev24', _MOIS24[1])} fin février à {F('PL nov24', _MOIS24[10])} fin novembre, puis il est viré au FPPN le 12/12/2024 pour payer l'EFC. Les mouvements annuels du compte additionnent ces renouvellements et ne mesurent pas le montant placé.",
                 f"**Le DAT PEP (compte 52210300) est resté placé de 2020 à 2025 pour {F('PL dpep', _DPEP['montant'])}.** Il est soldé le {_DPEP['date_debouclage']} pour {F('PL dpep deb', _DPEP['debouclage_2025'])}, dont {F('PL dpep int', _DPEP['interets_2025'])} d'intérêts enregistrés en une fois en 2025. [[Le FCP de 500 000 000 FCFA a un rendement annoncé de 8 à 10 %, non garanti]] : il n'est pas revalorisé dans les comptes."],
        implication=f"Le DAT EFC finance les paiements de fin d'année de l'EFC : ce n'est pas une réserve. Seuls le DAT PEP, puis le FCP, ont été des placements durables, face à une épargne de {F('PL ep 2026', b[2026]['epargne_totale'])} au 31/08/2026.",
        valable="grands livres des comptes 2021 à 2026 (« GRAND LIVRE 2021.xlsx », « GRAND LIVRE DES COMPTES 2022.xlsx », grand livre des comptes 2023, « GRAND LIVRE FPPN 2024.xlsx », grand livres des comptes 2025 et du 01/01 au 31/08/2026), comptes 5221, 5330 et 77 ; 2026 du 01/01 au 31/08",
        source="Source : grands livres des comptes du FPPN de 2021 au 31/08/2026 ; " + SRC_COMPTA[9:] + " ; " + SRC_ENT14[9:])
