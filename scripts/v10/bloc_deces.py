# ---- v10 : allocations décès (états de fin d'année reçus le 22/10/2026), risques courts, Grenier
_ET = V10["deces_etats"]
contenu({"titre": "Allocations décès : états de fin d'année et compte 41591000", "sous_titre": "Le calcul de fin d'année du FPPN est comptabilisé au compte 41591000, au franc près en 2024 et 2025",
         "synthese": "Chaque 31 décembre, le FPPN liste les décès non payés, applique le barème du grade et passe le total en dette envers les bénéficiaires"},
        "Barème appliqué dans les fichiers : membre sous-officier 1 400 000 FCFA, officier 1 600 000, commissaire 2 200 000 ; conjoint 600 000 à 700 000 ; enfant 120 000 à 150 000. C'est le barème de la cartographie des risques, non celui déclaré en entretien le 19/10/2026.",
        tableau={"titre": tab("États de fin d'année des allocations décès et crédits du compte 41591000"), "entetes": ["Exercice", "Fichier", "Dossiers", "Total du fichier", "Crédit au 41591000", "Écart"],
                 "largeurs": [2.6, 12.0, 2.8, 4.6, 4.8, 4.4], "taille": 8.5,
                 "lignes": [[str(x["exercice"]), x["fichier"], N(f"DC dossiers {x['exercice']}", x["dossiers_valorises"]), F(f"DC total {x['exercice']}", x["total_imprime"]),
                             F(f"DC crédit {x['exercice']}", x["credit_41591000"]), F(f"DC écart {x['exercice']}", x["credit_41591000"] - x["total_imprime"])] for x in _ET]},
        disposition="dessous",
        analyse=["**En 2024 et 2025, le crédit de l'année au compte 41591000 est exactement le total du fichier.** Le grand livre 2024 montre l'écriture du 31/12/2024 « PROVISION ALLO DECES 2024 » de 140 800 000 FCFA.",
                 "**En 2022 et 2023, les écarts restent à expliquer ; en 2021, le premier crédit reprend aussi des dossiers antérieurs à l'ouverture du compte.** Le fichier 2024 porte la mention « BILAN NON EXHAUSTIF » et la somme de ses lignes dépasse son total de 1 400 000 FCFA."],
        implication="La méthode du FPPN est bien une évaluation dossier par dossier (statuts, art. 89). Elle est comptabilisée comme une dette, non comme une provision : la provision obsèques ne doit pas la reprendre une seconde fois.",
        valable="états de fin d'année 2021 à 2025 tels que reçus le 22/10/2026 ; crédits du compte 41591000 des balances 2021 à 2025",
        source="Source : " + SRC_LOT4C + " ; « GRAND LIVRE FPPN 2024.xlsx » (compte 41591000) ; « CARTOGRAPHIE DES RISQUES COUVERTS ET DES PRESTATIONS OFFERTES.docx » ; balances 2021 à 2025 ; onglet 19a du classeur des calculs")
_DS = [x for x in V10["deces_stock"] if x["exercice"] >= 2021]


def _dette(x):
    return (x["ouv_41591000"] + x["mvt_c_41591000"] - x["mvt_d_41591000"]) + (x["ouv_41592000"] + x["mvt_c_41592000"] - x["mvt_d_41592000"])


contenu({"titre": "Allocations décès : dette réelle et provision obsèques", "sous_titre": f"La provision obsèques vaut {N('DD fois 2026', PROV_OBS / DETTE_DEC26, 1)} fois la dette décès calculée par le FPPN au 31/08/2026",
         "synthese": f"Dette décès (41591000 et 41592000) : {F('DD dette 2026', DETTE_DEC26)} ; provision obsèques inchangée depuis 2020 : {F('DD prov', PROV_OBS)}"},
        "Dette en fin d'année = dette du début + allocations ajoutées par l'état de fin d'année - allocations payées. La provision obsèques (comptes 19012000 puis 15802000) n'est rattachée à aucun calcul transmis.",
        tableau={"titre": tab("Dette décès du FPPN et provision obsèques"), "entetes": ["Exercice", "Allocations ajoutées", "Allocations payées", "Dette allocations", "Dette frais funéraires", "Dette décès totale", "Provision obsèques"],
                 "largeurs": [3.0, 4.6, 4.6, 4.6, 4.4, 4.6, 4.8], "taille": 9,
                 "lignes": [[str(x["exercice"]) if x["exercice"] < 2026 else "31/08/2026", F(f"DD ajout {x['exercice']}", x["mvt_c_41591000"]), F(f"DD payé {x['exercice']}", x["mvt_d_41591000"]),
                             F(f"DD all {x['exercice']}", x["ouv_41591000"] + x["mvt_c_41591000"] - x["mvt_d_41591000"]), F(f"DD ff {x['exercice']}", x["ouv_41592000"] + x["mvt_c_41592000"] - x["mvt_d_41592000"]),
                             F(f"DD tot {x['exercice']}", _dette(x)), F(f"DD prov {x['exercice']}", x["provision_obseques"])] for x in _DS]},
        disposition="dessous",
        analyse=[f"**Chaque année, le FPPN ajoute plus d'allocations qu'il n'en paie : la dette passe de {F('DD 2021', _dette(_DS[0]))} fin 2021 à {F('DD 2026 (a)', DETTE_DEC26)} au 31/08/2026.** Le stock contient donc des dossiers de plusieurs années.",
                 "**En 2024, des allocations ont été payées depuis les comptes bancaires de l'épargne** : 93 paiements pour 49 000 000 FCFA par le compte BOA PEP et 9 pour 4 288 000 FCFA par le compte BOA EFC (grand livre 2024)."],
        implication="La provision obsèques fait double emploi avec la dette du 41591000. La dette elle-même doit être justifiée dossier par dossier, avec la date de dépôt de chaque demande, pour écarter les demandes frappées de forclusion (règlement intérieur, art. 58 : 3 ans pour un membre, 12 mois pour un ayant droit).",
        valable="comptes 41591000, 41592000, 19012000 et 15802000, exercices 2021 à 2025 et du 01/01 au 31/08/2026 (provisoire) ; paiements par banque d'après le grand livre 2024 seulement",
        source="Source : " + SRC_COMPTA[9:] + " ; « GRAND LIVRE FPPN 2024.xlsx » (journaux BQ PEP, BQEFC) ; " + SRC_RI + ", art. 58 ; onglet 19b du classeur des calculs")
_TD = V10["deces_tardifs"]
contenu({"titre": "Allocations décès : décès connus après la clôture", "sous_titre": f"{N('DT nombre', len(_TD))} décès d'une année n'ont été enregistrés que l'année suivante, pour {F('DT montant', TARD_DEC)}",
         "synthese": "Ces décès manquent dans l'état du 31 décembre : ils donnent l'ordre de grandeur de la provision pour décès survenus non déclarés"},
        "Seuls les états 2021 et 2023 portent à la fois la date du décès et la date d'enregistrement du dossier. On retient les dossiers dont le décès est d'une année et l'enregistrement de l'année suivante.",
        tableau={"titre": tab("Décès enregistrés l'année suivant le décès"), "entetes": ["État", "Date du décès", "Date d'enregistrement", "Montant"], "largeurs": [3.6, 6.0, 6.0, 6.0], "taille": 9,
                 "lignes": [[str(x["etat"]), x["deces"][8:10] + "/" + x["deces"][5:7] + "/" + x["deces"][:4], x["enregistrement"][8:10] + "/" + x["enregistrement"][5:7] + "/" + x["enregistrement"][:4],
                             F(f"DT {i}", x["montant"])] for i, x in enumerate(_TD)]},
        disposition="dessous",
        analyse=[f"**Les décès déclarés tardivement représentent {F('DT montant (a)', TARD_DEC)} sur deux états, soit quelques dossiers par an.** Le délai de forclusion de trois ans (art. 58) borne la période pendant laquelle un décès peut encore être déclaré."],
        implication="La provision décès nécessaire est la dette du 41591000, justifiée dossier par dossier, plus une estimation des décès non encore déclarés de quelques millions de FCFA par an.",
        valable="états de fin d'année 2021 et 2023, seuls à porter la date d'enregistrement ; dossiers de membres et d'ayants droit valorisés",
        source="Source : « PROV ALLOCATION DECES 2021.xlsx » (feuille IMPAYEES GLOBAL 2021) ; « ALLOC DECES 2023 PAYER ET A PROVISIONNER .xlsx » (feuille ALL DECES  A PROVISIONNER 2023) ; onglet 19c du classeur des calculs")
contenu({"titre": "Provision pour risques courts", "sous_titre": "Aucune provision n'est nécessaire : les cotisations mensuelles couvrent leur propre mois",
         "synthese": "La provision pour risques courts (statuts, art. 88) couvre la période entre la clôture et la prochaine échéance de cotisation ; cette période est nulle au FPPN"},
        "Les cotisations maladie et décès sont précomptées chaque mois sur la solde du mois (postes 936 et 8Z6) ou sur la pension (CGRAE) : la cotisation de décembre couvre décembre. Le Code CIMA (art. 334-10), référence technique des assureurs, exclut de même les cotisations « à terme échu » du calcul de la provision pour risques en cours.",
        tableau={"titre": tab("Provision pour risques courts au 31 décembre"), "entetes": ["Exercice", "Prestations et frais de l'année", "Jours non couverts par une cotisation", "Provision"],
                 "largeurs": [3.6, 8.0, 7.0, 6.0], "taille": 9,
                 "lignes": [[str(a), F(f"RC charges {a}", s[a]["prest_maladie"] + s[a]["prest_deces"] + s[a]["frais_gestion"]), "0", F(f"RC prov {a}", 0)] for a in range(2020, 2026)]},
        disposition="dessous",
        analyse=["**La provision pour risques courts est nulle tant que les cotisations sont précomptées sur la solde ou la pension du mois qu'elles couvrent.** Elle deviendrait nécessaire si des cotisations étaient encaissées d'avance pour plusieurs mois."],
        implication="Le FPPN n'a pas de provision pour risques courts à constituer ; il doit le justifier par écrit dans sa note de calcul annuelle des provisions.",
        valable="cotisations maladie et décès précomptées par la Solde (postes 936 et 8Z6, janvier 2024 à septembre 2026) et par la CGRAE (février 2024 à septembre 2026)",
        source="Source : " + SRC_ST + ", art. 88 ; Code CIMA, art. 334-10 (référence technique) ; " + SRC_PREC[9:] + " ; onglet 20 du classeur des calculs")
_GRD10 = V10["grenier_depreciation"]
contenu({"titre": "Prêts du Grenier : dépréciation selon la politique de crédit", "sous_titre": f"La politique du Grenier conduit à déprécier {F('GD 2026', GRD['2026-08-31'])} au 31/08/2026 ; les comptes n'en portent aucune",
         "synthese": "Taux de la politique de crédit (chapitre XI.3) : 40 % de 3 à 6 mois de retard, 80 % de 6 à 12 mois, 100 % de 12 à 24 mois"},
        "Dépréciation = encours de la tranche de retard × taux de la politique de crédit du Grenier. Les tranches du logiciel de prêts (91 à 180 jours, 181 à 365 jours, plus de 365 jours) correspondent aux trois paliers.",
        tableau={"titre": tab("Dépréciation des prêts du Grenier selon sa politique de crédit"), "entetes": ["Arrêté", "Tranche de retard", "Prêts", "Encours", "Taux", "Dépréciation"], "largeurs": [3.6, 5.6, 2.8, 5.6, 3.0, 5.4], "taille": 8.5,
                 "lignes": [[x["arrete"][8:10] + "/" + x["arrete"][5:7] + "/" + x["arrete"][:4], x["tranche"], N(f"GD prêts {x['arrete']} {x['tranche']}", x["prets"]), F(f"GD enc {x['arrete']} {x['tranche']}", x["encours"]),
                             pc(f"GD taux {x['arrete']} {x['tranche']}", x["taux"]), F(f"GD dep {x['arrete']} {x['tranche']}", x["encours"] * x["taux"])] for x in _GRD10 if x["taux"] > 0]},
        disposition="dessous",
        analyse=[f"**La dépréciation passe de {F('GD 2025', GRD['2025-12-31'])} au 31/12/2025 à {F('GD 2026 (a)', GRD['2026-08-31'])} au 31/08/2026.** Le montant reste faible au regard de l'encours, mais il doit être comptabilisé."],
        implication="Le risque principal du Grenier reste l'immobilisation de l'épargne dans les prêts, non la perte sur prêts.",
        valable="prêts du Grenier aux arrêtés du 31/12/2025 et du 31/08/2026, tranches de retard du logiciel de prêts",
        source="Source : « POLITIK DE CREDIT DU GRENIER VFVF.doc », chapitre XI.3 ; " + SRC_GRENIER[9:] + " ; onglet 23 du classeur des calculs")
