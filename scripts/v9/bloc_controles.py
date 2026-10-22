# ---- v9 : contrôles croisés (partie 3)
EFB = V9["ef_contre_balances"]
_e23 = next(x for x in EFB if x["exercice"] == 2023)
_e26 = next(x for x in EFB if x["exercice"] == 2026)
_nok = [x["exercice"] for x in EFB if x["ef_brut"] != x["balance_brut"] or x["ef_resultat"] != x["balance_resultat"]]
contenu({"titre": "Contrôle croisé : états financiers et balances", "sous_titre": f"Les états financiers concordent avec les balances {len(EFB) - len(_nok)} années sur {len(EFB)}",
         "synthese": "Total de l'actif, résultat et trésorerie identiques au franc près, sauf en 2023 (deux versions des comptes) et au 31/08/2026 (résultat)"},
        "Les montants des états financiers ont été relevés sur les pages « Bilan Actif » et « Bilan Passif » des PDF scannés, puis comparés aux balances utilisées dans tout le rapport. La trésorerie-actif des états financiers est comparée à la somme des soldes débiteurs de la classe 5, car les découverts figurent au passif.",
        tableau={"titre": tab("États financiers et balances : actif brut, résultat, trésorerie"), "entetes": ["Exercice", "Actif brut (états financiers)", "Écart avec la balance", "Résultat (états financiers)", "Écart avec la balance", "Trésorerie : écart"],
                 "largeurs": [3.6, 6.0, 4.8, 6.0, 4.8, 4.2], "taille": 9,
                 "lignes": [[str(x["exercice"]) if x["exercice"] < 2026 else "31/08/2026", F(f"EF brut {x['exercice']}", x["ef_brut"]), F(f"EF écart brut {x['exercice']}", x["ef_brut"] - x["balance_brut"]),
                             F(f"EF résultat {x['exercice']}", x["ef_resultat"]), F(f"EF écart rés {x['exercice']}", x["ef_resultat"] - x["balance_resultat"]),
                             F(f"EF écart trés {x['exercice']}", x["ef_tresorerie"] - x["balance_classe5"])] for x in EFB]},
        disposition="dessous",
        analyse=[f"**En 2023, les états financiers scannés portent un actif brut inférieur de {F('EF écart 2023 (t)', _e23['balance_brut'] - _e23['ef_brut'])} et un excédent de {F('EF rés 2023 (t)', _e23['ef_resultat'])}, contre {F('Bal rés 2023 (t)', _e23['balance_resultat'])} dans la balance.** Le rapport de gestion 2023 « corrigé et finalisé » et la colonne 2023 des états financiers 2024 reprennent les montants de la balance : le scan est une version antérieure à la correction.",
                 f"**Au 31/08/2026, le résultat des états financiers ({F('EF rés 2026 (t)', _e26['ef_resultat'])}) diffère de celui de la balance provisoire ({F('Bal rés 2026 (t)', _e26['balance_resultat'])}).** L'actif brut et la trésorerie sont pourtant identiques."],
        implication="Les balances utilisées dans le rapport sont confirmées par les comptes publiés. Le FPPN doit indiquer quelle version des comptes 2023 a été approuvée et expliquer l'écart de résultat au 31/08/2026 (PBC, item 110).",
        valable="exercices 2020 à 2025 et situation au 31/08/2026, totaux du bilan et résultat ; les rubriques détaillées n'ont pas été comparées",
        source="Source : ETATS FINANCIERS 2020.pdf à ETATS FINANCIERS 2025.pdf, ETATS FINANCIERS AU 31-08-2026.pdf, BILAN_DE_L'EXERCICE_CLOS_31.12.2023_(corrigé_et_finalisé)[1].pdf ; " + SRC_COMPTA[9:])
_ad = V9["controle_adherents_8z6"]
_rm = {(x["annee"], x["mois"]): x for x in V9["rapport_8z6_936_mois"]}
_ra = V9["rapport_8z6_936_annee"]
contenu({"titre": "Contrôle croisé : effectifs du FPPN et fichiers de la Solde", "sous_titre": "Les effectifs déclarés par le FPPN sont ceux des fichiers 8Z6, à dix agents près en 2026",
         "synthese": "Le nombre d'adhérents du FPPN égale le nombre de lignes 8Z6 en novembre 2024 et décembre 2025 ; la clé 3,77 / 1,23 s'applique depuis mars 2024"},
        "Deux sources indépendantes sont comparées : l'étude des ratios du FPPN (nombre d'adhérents par période) et les fichiers retour 8Z6 de la Solde (une ligne par policier précompté). La seconde partie du tableau rapporte, agent par agent, le précompte décès (8Z6) au précompte maladie (936) du même mois : la clé 3,77 / 1,23 donne un rapport de 0,3263.",
        tableau={"titre": tab("Adhérents du FPPN, lignes 8Z6 et rapport 8Z6 / 936"), "entetes": ["Période", "Adhérents (FPPN)", "Lignes 8Z6 (Solde)", "Écart", "Rapport médian 8Z6 / 936", "Lignes à la clé 3,77 / 1,23"],
                 "largeurs": [4.0, 5.0, 5.0, 3.4, 6.0, 6.0], "taille": 9,
                 "lignes": [[f"{x['mois']:02d}/{x['annee']}", N(f"Adh FPPN {x['annee']}-{x['mois']}", x["adherents_fppn"]), N(f"Lignes 8Z6 {x['annee']}-{x['mois']}", x["lignes_8z6"]),
                             N(f"Écart adh {x['annee']}-{x['mois']}", x["adherents_fppn"] - x["lignes_8z6"]),
                             nb(_rm[(x["annee"], x["mois"])]["mediane"], 4) if (x["annee"], x["mois"]) in _rm else "-",
                             pc(f"Part clé {x['annee']}-{x['mois']}", _rm[(x["annee"], x["mois"])]["part_cle"]) if (x["annee"], x["mois"]) in _rm else "-"] for x in _ad]
                          + [[f"{m:02d}/2024", "-", N(f"Lignes 8Z6 2024-{m}", _rm[(2024, m)]["lignes"]), "-", nb(_rm[(2024, m)]["mediane"], 4), pc(f"Part clé 2024-{m}", _rm[(2024, m)]["part_cle"])] for m in (1, 2, 3)]},
        disposition="dessous",
        analyse=[f"**La répartition 3,77 / 1,23 s'applique à {pc('Part clé 2025 (t)', next(x['part_cle'] for x in _ra if x['annee'] == 2025))} des lignes de 2025 et à {pc('Part clé 2026 (t)', next(x['part_cle'] for x in _ra if x['annee'] == 2026))} des lignes de 2026.** En janvier et février 2024, le rapport médian vaut {nb(_rm[(2024, 1)]['mediane'], 4)} et {nb(_rm[(2024, 2)]['mediane'], 4)} : une autre clé s'appliquait alors.",
                 "**De janvier à juin 2026, le FPPN compte chaque mois dix adhérents de moins que les fichiers 8Z6.** Les cotisations décès déclarées par le FPPN dépassent aussi la somme des lignes 8Z6 du mois."],
        implication="Les effectifs utilisés par le FPPN pour ses ratios sont fiables à l'unité près. La clé appliquée avant mars 2024 et l'origine des dix agents d'écart sont à demander (PBC, item 111).",
        valable="mois cités, policiers précomptés par la Solde (8Z6 et 936) ; les mois de 2024 sans fichier 8Z6 ne sont pas couverts",
        source="Source : ETUDE DES RATIO MALADIE DECES 2.xlsx, feuille Tableau, lignes « NB D'ADHERENTS » et « COTISATIONS PAYEES » ; " + SRC_PREC[9:])
_rb = V9["rapprochements"]
_r25 = next(x for x in _rb if x["exercice"] == 2025)
contenu({"titre": "Contrôle croisé : rapprochements bancaires", "sous_titre": "Les soldes bancaires de la comptabilité sont ceux de la balance ; des chèques restent en suspens depuis 2021",
         "synthese": f"Au 31/12/2025, {N('Chèques suspens 2025', _r25['suspens_plus_un_an'])} chèques émis depuis plus d'un an ne sont pas encaissés, pour {F('Montant suspens 2025', _r25['montant_suspens_plus_un_an'])}"},
        "Chaque état de rapprochement compare, à une date, le solde du compte dans la comptabilité du FPPN et le solde du relevé de la banque, puis liste les opérations qui expliquent l'écart. Les chèques émis par le FPPN mais pas encore présentés à la banque figurent dans cette liste.",
        tableau={"titre": tab("Rapprochements bancaires au 31 décembre"), "entetes": ["Exercice", "Comptes rapprochés", "Soldes égaux à la balance", "Solde comptable", "Solde bancaire", "Chèques en suspens depuis plus d'un an", "Montant"],
                 "largeurs": [2.8, 3.4, 3.8, 5.2, 5.2, 5.4, 4.6], "taille": 9,
                 "lignes": [[str(x["exercice"]), N(f"ERB comptes {x['exercice']}", x["comptes"]), N(f"ERB concord {x['exercice']}", x["concordants"]), F(f"ERB compta {x['exercice']}", x["solde_compta"]),
                             F(f"ERB banque {x['exercice']}", x["solde_banque"]), N(f"ERB suspens {x['exercice']}", x["suspens_plus_un_an"]), F(f"ERB montant {x['exercice']}", x["montant_suspens_plus_un_an"])] for x in _rb]},
        disposition="dessous",
        analyse=["**Chaque solde comptable des rapprochements est égal au solde du compte de banque dans la balance, pour toutes les années.** Les balances sont donc cohérentes avec le suivi bancaire du FPPN.",
                 f"**Le nombre de chèques en suspens depuis plus d'un an passe de {N('Suspens 2022 (t)', next(x['suspens_plus_un_an'] for x in _rb if x['exercice'] == 2022))} fin 2022 à {N('Suspens 2025 (t)', _r25['suspens_plus_un_an'])} fin 2025**, le plus ancien datant du {_r25['plus_ancien'][8:10]}/{_r25['plus_ancien'][5:7]}/{_r25['plus_ancien'][:4]}. Une partie concerne le compte BOA de l'EFC, donc des règlements à des mutualistes."],
        implication="Les chèques non encaissés restent une dette envers les bénéficiaires : leur montant est faible, mais leur ancienneté doit être expliquée et les chèques périmés régularisés (PBC, item 91).",
        valable="comptes bancaires présents dans les états de rapprochement (BOA, Banque Atlantique, Orabank), au 31/12 de 2020 à 2025 ; les DAT, le FCP et la caisse ne sont pas couverts",
        source="Source : ETAT DE RAPPROCHEMENT BANCAIRE FPPN 2020.xlsx à ETAT DE RAPPROCHEMENT BANCAIRE FPPN 2026.xlsx ; " + SRC_COMPTA[9:])
