# ---- v13 (13/11/2026) : diapositives nouvelles, placées par passe_v13.py après la diapositive indiquée par « _apres » (titre v12)
_TI13 = json.load(open(os.path.join(os.path.dirname(BASE), "test_interets", "regle_interets_pep.json"), encoding="utf-8"))
_f = lambda x: n13(x, 0)  # noqa: E731
_p = lambda x: n13(x, 2)  # noqa: E731


def _apres(titre):
    SL[-1]["_apres"] = titre


# 1. extourne oubliée de la provision de la clinique
contenu(blt13({"titre": "Une provision de 2025 oubliée gonfle les dépenses de 2026",
               "sous_titre": fk13("Après extourne des 255 000 000 FCFA provisionnés fin 2025, la maladie dépense {sc_mal_26:p} FCFA pour 100 FCFA de cotisations à fin août 2026, au lieu de 97,40"),
               "synthese": "Les factures de la clinique d'octobre à décembre 2025 sont comptées en 2025 (provision) et une seconde fois en 2026 (factures)"}),
        "Le 31/12/2025, le Fonds a provisionné trois mois de factures de la clinique. En 2026, il a passé les vraies factures en charges sans annuler la provision.",
        tableau={"titre": tab("Effet de l'extourne oubliée sur les chiffres de 2026"), "entetes": ["Chiffre au 31/08/2026", "Comptes tels quels", "Après extourne"], "largeurs": [13.2, 9.0, 9.0], "taille": 12,
                 "lignes": [["Prestations maladie pour 100 FCFA de cotisations maladie", "97,40", _p(K13["sc_mal_26"])],
                            ["Prestations pour 100 FCFA de cotisations, selon le texte (norme 70 à 80)", "75,20", _p(K13["regl_26"])],
                            ["Résultat technique maladie et décès (FCFA)", "-462 119 274", _f(K13["res_tech_26"])],
                            ["Résultat net (FCFA)", "-441 256 500", _f(K13["res_net_26"])],
                            ["Factures des cliniques à payer, compte 40110000 (FCFA)", "690 234 451", _f(K13["dette_cliniques_26"])]]},
        disposition="dessous",
        analyse=["**Provision du 31/12/2025 : 255 000 000 FCFA (OD 805 à 807). Vraies factures passées le 31/01/2026 : 240 471 119 FCFA (pièces SIN 462, 552 et 619).** Aucune écriture n'annule la provision en 2026.",
                 "Le rapport retient les chiffres après extourne : c'est l'écriture que le Fonds doit passer. [[Montant à confirmer par le service comptable.]]"],
        valable="comptes 2025 et du 01/01 au 31/08/2026 (provisoire)",
        source="Sources : (1) « GRAND LIVRE DES COMPTES FPPN 2025.xlsx », OD 805, 806 et 807 du 31/12/2025, comptes 60120000 et 40110000. (2) « GRAND LIVRE DES COMPTES FPPN DU 01-01-2026 AU 31-08-2026.xlsx », pièces SIN 462, 552 et 619 du 31/01/2026, comptes 6011 et 60120000. (3) « BALANCE FPPN DU 01-01-2026 AU 31-08-2026.xlsx », extourne ajoutée par Adenka")
_apres("Sinistralité comptable")

# 2. cotisation par adhérent
_c20, _a20, _c25, _a25 = 2344618501, 18654, 3417551722, 31678
contenu(blt13({"titre": "Chaque adhérent cotise moins qu'en 2020",
               "sous_titre": f"{_f(_c25 / _a25)} FCFA par adhérent en 2025 contre {_f(_c20 / _a20)} FCFA en 2020, soit {_p(100 * (_c25 / _a25 / (_c20 / _a20) - 1))} % : les recrues ont des salaires plus bas, alors que les soins coûtent plus cher",
               "synthese": "Cotisations comptabilisées rapportées au nombre d'adhérents de fin d'année"}),
        "La hausse des cotisations vient du nombre d'adhérents. La cotisation de chacun baisse.",
        tableau={"titre": tab("Cotisation moyenne par adhérent"), "entetes": ["Année", "Cotisations (FCFA)", "Adhérents", "Par adhérent (FCFA)"], "largeurs": [5.0, 9.4, 7.0, 9.8], "taille": 12,
                 "lignes": [["2020", _f(_c20), _f(_a20), _f(_c20 / _a20)], ["2025", _f(_c25), _f(_a25), _f(_c25 / _a25)],
                            ["2025 aux prix de 2020", "", "", _f(_c25 / _a25 / 1.185)]]},
        disposition="dessous",
        analyse=[f"**Hors hausse des prix (18,50 %), la cotisation par adhérent baisse de {_p(100 * (_c25 / _a25 / 1.185 / (_c20 / _a20) - 1))} % de 2020 à 2025.** C'est l'effet de ciseau le plus direct : chaque adhérent apporte moins pendant que ses soins coûtent plus."],
        valable="cotisations comptabilisées 2020 et 2025, adhérents au 31 décembre",
        source="Sources : (1) « BALANCE FORMAT EXCEL DE 09-2020 AU 31-12-2020.xlsx » et « BALANCE FPPN 2025.xlsx », comptes 7010 et 7021. (2) états des effectifs du Fonds (diapositive « Mortalité brute des membres participants »). (3) Banque mondiale, FP.CPI.TOTL.ZG, Côte d'Ivoire")
_apres("Cotisations comptabilisées")

# 3. règle des intérêts du plan d'épargne
_ex = _TI13["A_remboursements_par_exercice"]
contenu(blt13({"titre": "Le Fonds rend 3,5 FCFA pour 100 FCFA versés, une seule fois",
               "sous_titre": f"soit 3,5 % du montant versé (cotisé), quelle que soit la durée d'épargne : {_f(_TI13['A_total']['avec_3_5'])} remboursements sur {_f(_TI13['A_total']['remboursements'])} suivent cette règle de 09/2020 à 08/2026",
               "synthese": f"Pour 100 000 FCFA versés, le policier reçoit 103 500 FCFA à sa sortie. Intérêts ainsi payés : {_f(_TI13['A_total']['interet_paye'])} FCFA"}),
        "Les entretiens conduits avec le Fonds font ressortir que la règle du plan d'épargne est de 3,5 % une seule fois. Les comptes le confirment.",
        tableau={"titre": tab("Remboursements du plan d'épargne qui suivent la règle des 3,5 %"), "entetes": ["Exercice", "Remboursements", "Avec 3,5 %", "Pour 100", "Intérêts payés (FCFA)"], "largeurs": [4.6, 6.4, 6.0, 4.8, 9.4], "taille": 12,
                 "lignes": [[str(int(e["exercice"])), _f(e["remboursements"]), _f(e["avec_3_5"]), _p(100 * e["avec_3_5"] / e["remboursements"]), _f(e["interet_paye"])] for e in _ex]},
        disposition="dessous",
        analyse=["**Les entretiens conduits avec le Fonds font ressortir que le plan d'épargne rend 3,5 % du montant versé, une seule fois.** [[Service et date de l'entretien à compléter.]] Un remboursement suit la règle quand son montant divisé par 1,035 est un multiple de 5 000 FCFA, au franc près.",
                 "Les intérêts sont pris sur le compte d'épargne lui-même, sans charge : l'épargne due aux autres policiers est sous-estimée d'autant. [[Le manuel d'épargne et l'entretien du 19/10/2026 parlent de 3,5 % « tous les ans » : le texte qui fixe la règle reste à recevoir.]]"],
        valable="remboursements du compte 41195100, 09/2020 au 31/08/2026 (2020 : quatre mois)",
        source="Sources : (1) grands livres des comptes 2020 à 2026, compte 41195100, débits libellés REMB, REGLT, SOLDE ou RESIL, journal RAN exclu, calcul Adenka. (2) « MANUEL DE PRODUIT D'EPARGNE FPPN VF VF.docx », partie PEP. (3) « FPPN_CR_entretien_provisions_20261019.pptx », diapositive 6")
_apres("Intérêts servis : aucune écriture dans les comptes")

# 4. santé des retraités : ordre de grandeur
_cot_an, _nr, _a60 = 11954 * 12, 610, 16.84
contenu(blt13({"titre": f"La santé des retraités peut coûter jusqu'à {_f(_nr * _cot_an * _a60)} FCFA",
               "sous_titre": "Ordre de grandeur, pas une provision : tant que les soins d'un retraité coûtent moins que sa cotisation, le Fonds ne perd rien",
               "synthese": "Engagement à vie estimé pour 610 retraités, avec une rente de 16,84 FCFA par FCFA annuel (table CIMA F, 3,5 %, 60 ans)"}),
        "Sans l'âge ni le coût des soins de chaque retraité, seul un ordre de grandeur est possible.",
        tableau={"titre": tab("Engagement santé des retraités selon le coût des soins"), "entetes": ["Coût des soins d'un retraité", "Coût net par an (FCFA)", "Engagement à vie (FCFA)"], "largeurs": [12.4, 9.0, 9.8], "taille": 12,
                 "lignes": [[f"{k_} fois sa cotisation ({_f(_cot_an)} FCFA par an)", _f(_cot_an * (k_ - 1)), _f(_nr * _cot_an * (k_ - 1) * _a60)] for k_ in (1, 1.5, 2)]},
        disposition="dessous",
        analyse=["**En 2025, un policier actif a coûté 55 926 FCFA de soins par an, moins que la cotisation d'un retraité (143 448 FCFA).** Si les retraités coûtent comme les actifs, l'engagement est nul. S'ils coûtent deux fois leur cotisation, il dépasse un milliard de FCFA.",
                 "[[La rente à 60 ans majore la valeur pour un âge moyen de 63,3 ans : le chiffre exact demande l'âge et les soins de chaque retraité.]]"],
        valable="610 retraités cotisants en août 2026, table CIMA F, taux de 3,5 %",
        source="Sources : (1) fichiers retour de la caisse de retraite (CGRAE), août 2026. (2) table CIMA F, rente payable d'avance à 60 ans (classeur, onglet 14t_table_cima_f). (3) « BALANCE FPPN 2025.xlsx », comptes 601 et 60800007, 31 678 adhérents")
_apres("Santé des retraités : ce que l'on peut mesurer")

# 5. passage des fonds propres comptables aux fonds propres corrigés
_fp = K13["fonds_propres_26"]
_pont = [("Fonds propres dans les comptes, extourne comprise", _fp),
         ("Plus : provisions au-delà du besoin estimé par Adenka", K13["pont_provisions"]),
         ("Moins : intérêts du plan d'épargne dus (règle des entretiens)", -K13["interets_pep_dus"]),
         ("Moins : dépréciation des prêts du Grenier selon sa politique", K13["pont_grenier"])]
_s1 = sum(v for _, v in _pont)
_pont2 = [("Moins : terrains sans titre et créance sur l'hôpital", -K13["actifs_contestes"]),
          ("Moins : prêts aux adhérents de 2004 à 2015, jugés perdus", -379204261),
          ("Moins : charges de la clinique inscrites à l'actif", -1280300860)]
_s2 = _s1 + sum(v for _, v in _pont2)
contenu(blt13({"titre": "Une fois les biens incertains retirés, les fonds propres sont négatifs",
               "sous_titre": f"Corrigés des provisions et des intérêts, les fonds propres sont de {_f(_s1)} FCFA : sans les biens de valeur non établie, ils tombent à {_f(_s2)} FCFA",
               "synthese": "Passage des fonds propres comptables aux fonds propres corrigés au 31/08/2026, poste par poste"}),
        "Chaque correction est chiffrée par le rapport. Le signe des fonds propres dépend de la valeur des biens.",
        tableau={"titre": tab("Des fonds propres comptables aux fonds propres corrigés au 31/08/2026"), "entetes": ["Poste", "FCFA", "Statut"], "largeurs": [17.0, 7.4, 6.8], "taille": 12,
                 "lignes": [[a, _f(v), "Mesuré" if i_ != 2 else "Règle des entretiens"] for i_, (a, v) in enumerate(_pont)]
                 + [["Fonds propres corrigés avant valeur des biens", _f(_s1), ""]]
                 + [[a, _f(v), "Sensibilité"] for a, v in _pont2]
                 + [["Fonds propres corrigés", _f(_s2), ""], ["[[Moins 2 899 400 000 FCFA si le terrain du siège est compté deux fois]]", "[[-2 899 400 000]]", "À confirmer"]]},
        disposition="dessous",
        analyse=["**Les provisions en trop rendent les fonds propres positifs, la valeur incertaine des biens les rend négatifs.** La conclusion dépend donc de pièces que le Fonds doit produire : titres des terrains, accord de remboursement de l'hôpital, écriture d'origine du siège."],
        valable="bilan au 31/08/2026 (provisoire), extourne ajoutée par Adenka ; intérêts de l'entraide non compris",
        source="Sources : (1) « BALANCE FPPN DU 01-01-2026 AU 31-08-2026.xlsx », classes 1 à 7. (2) estimations Adenka des provisions (classeur, onglets 18 à 20). (3) mémo du service comptable sur les comptes 47 et 27 (11/11/2026). (4) politique de crédit du Grenier, chapitre XI.3")
_apres("Fonds propres")

# 6. emplacements des six diapositives sur les dettes (remplies après le rendu par dettes_v13.py)
for _k13, _t13 in enumerate(("Dettes 1 : l'épargne est la plus grosse dette", "Dettes 2 : dettes de court terme", "Dettes 3 : détail (1/2)",
                             "Dettes 4 : détail (2/2)", "Dettes 5 : trois lectures de l'argent disponible", "Dettes 6 : intérêts promis"), 1):
    contenu({"titre": _t13, "sous_titre": "emplacement", "synthese": ""}, "", analyse=["emplacement"])
    SL[-1]["_dettes13"] = _k13
    SL[-1]["_apres"] = "Adossement de liquidité provisoire"
