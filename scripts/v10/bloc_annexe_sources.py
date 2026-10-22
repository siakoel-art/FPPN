# ---- v10 : annexe « Fichiers sources » (noms exacts des fichiers cités par les lignes « Valable pour » et les sources)
_sm = pd.read_parquet(os.path.join(BASE, "sources_mois.parquet"))
_sm = _sm[_sm.retenu]
_gb = pd.read_parquet(os.path.join(BASE, "grenier_balances.parquet"), columns=["fichier"]).fichier.drop_duplicates()
_gp = pd.read_parquet(os.path.join(BASE, "grenier_production.parquet"), columns=["periode_du"])
_cb = pd.read_parquet(os.path.join(BASE, "compta_balances.parquet"), columns=["nature", "fichier"]).drop_duplicates()
_POSTE = {"936": "Cotisation maladie des actifs", "8Z6": "Cotisation décès des actifs", "998": "Épargne PEP", "870": "Épargne EFC", "9X4": "Remboursements du Grenier", "9W3": "Primes GNA (flotte automobile)", "8Z7": "MEDOC", "CGRAE": "Cotisations des retraités"}
_lsrc = []
for po, lib in _POSTE.items():
    g = _sm[_sm.poste == po]
    if len(g):
        dossiers = sorted({d.split("\\")[0] + "\\" + d.split("\\")[1] + ("\\" + d.split("\\")[2] if len(d.split("\\")) > 2 and not d.split("\\")[2].isdigit() else "") for d in g.dossier})
        _lsrc.append([f"Précomptes, poste {po}", lib, (" ; ".join("« " + d + " »" for d in dossiers) + " : ", ("FS n " + po, g.fichier.nunique()), " fichiers"),
                      f"{int(g[g.annee == g.annee.min()].mois.min()):02d}/{int(g.annee.min())} à {int(g[g.annee == g.annee.max()].mois.max()):02d}/{int(g.annee.max())}"])
_lsrc = [["Balances des comptes", "Soldes et mouvements par compte", " ; ".join("« " + f + " »" for f in _cb[_cb.nature == "comptes"].fichier), "01/09/2020 (à-nouveaux) au 31/08/2026"],
         ["Balances des tiers", "Soldes par tiers", " ; ".join("« " + f + " »" for f in _cb[_cb.nature == "tiers"].fichier), "2020, 2022 à 2024, 2026"],
         ["Grands livres", "Écritures détaillées", "« GRAND LIVRE FPPN 2024.xlsx » ; « GRAND LIVRE DES TIERS FPPN 2023.xlsx » ; « GRAND LIVRE DES TIERS 2024.xlsx » ; « GRAND LIVRES DES TIERS 2025.xlsx »", "2023 à 2025"],
         ["Allocations décès (fin d'année)", "Décès non payés valorisés au barème", SRC_LOT4, "2021 à 2025"],
         ["Grenier, balances âgées", "Encours et retards par prêt", ("", ("FS gb", len(_gb)), " fichiers « BalanceAgeeDetaille … .xls » (annuels 2021 à 2025, mensuels 2026 par portefeuille CRF, CRO, CRS)"), "31/12/2021 au 30/09/2026"],
         ["Grenier, production", "Prêts débloqués", "fichiers « production_de_la_periode … .xls » (annuels 2021 à 2025, mensuels 2026)", "2021 à 09/2026"]] + _lsrc + [
         ["Personnel", "Déclarations des salaires", SRC_DISA[9:], "2021 à 2025"],
         ["Actifs et gouvernance", "Clinique, projet de Cocody, directeur général", SRC_CMP + " ; " + SRC_COC + " ; " + SRC_DG, "2020 à 2025"],
         ["Tables de mortalité", "Lx, qx, dx par âge", "« CIMA_H.xlsx » ; « CIMA_F.xlsx » ; « TD_CIMA.xlsx » ; « TV_CIMA.xlsx »", "Âges 0 à 110 ans"]]
for _k in range(0, len(_lsrc), 7):
    contenu({"titre": f"Fichiers sources ({_k // 7 + 1}/{(len(_lsrc) + 6) // 7})", "sous_titre": "Noms exacts des fichiers du FPPN cités dans les lignes « Valable pour » et dans les sources",
             "synthese": "Les lignes « Valable pour » citent les familles de fichiers ; ce tableau donne les noms exacts"},
            "Tous les fichiers ont été reçus du FPPN le 15/10/2026 (précomptes, balances, Grenier), le 18/10/2026 (tables de mortalité) ou le 22/10/2026 (autres documents).",
            tableau={"titre": tab("Fichiers sources et période couverte"), "entetes": ["Famille", "Contenu", "Fichiers exacts", "Période"], "largeurs": [4.6, 4.6, 17.4, 4.6], "taille": 7.5, "lignes": [[(c[0] + N(*c[1]) + c[2]) if isinstance(c, tuple) else c for c in lg_] for lg_ in _lsrc[_k:_k + 7]]},
            disposition="dessous", valable="fichiers reçus jusqu'au 22/10/2026 et retenus après les contrôles de la partie 3", source="Source : inventaire des fichiers reçus du FPPN ; traitements Adenka")
