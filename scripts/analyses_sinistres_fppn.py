# -*- coding: utf-8 -*-
"""
FPPN - Analyses statistiques des prestations décès et santé (volet actuariel)
=============================================================================

Données attendues (format des onglets « Modèle données » de la demande complémentaire) :
  - FICHIER_DECES        : une ligne par dossier décès (demande n°28)
  - FICHIER_SANTE        : une ligne par facture santé (demande n°40)
  - FICHIER_COTISATIONS  : une ligne par année : Annee, Cotisations_deces, Cotisations_maladie
                           (issues de la comptabilité, encaissées ou appelées : à préciser dans HYPOTHESES)
  - FICHIER_MUTUALISTES  : (facultatif) une ligne par mutualiste : Matricule, Date_naissance, Grade,
                           Statut, Date_entree, Date_sortie

Étapes du script :
  1. contrôle de la qualité des données (champs manquants, dates incohérentes, doublons, montants anormaux) ;
  2. analyses décès : sinistralité, fréquence, coût moyen, délais, stock non payé, triangle et Chain Ladder ;
  3. analyses santé : coût par bénéficiaire, concentration des prestataires, doublons, éligibilité ;
  4. comparaison avec les chiffres cités en entretien ;
  5. écriture d'un fichier Excel de résultats, avec un onglet par analyse et un onglet « Méthode ».

Lancer : python analyses_sinistres_fppn.py   (après avoir renseigné les chemins ci-dessous)
Un fichier manquant n'arrête pas le script : les analyses qui en dépendent sont signalées « non réalisées ».
"""

# =============================================================================
# PARAMÈTRES
# =============================================================================
DOSSIER = r"C:\Dev\Missions\05-FPPN\05-Analyses\donnees"
FICHIER_DECES = DOSSIER + r"\sinistres_deces.xlsx"
FICHIER_SANTE = DOSSIER + r"\prestations_sante.xlsx"
FICHIER_COTISATIONS = DOSSIER + r"\cotisations_par_annee.xlsx"
FICHIER_MUTUALISTES = DOSSIER + r"\mutualistes.xlsx"
DOSSIER_SORTIE = r"C:\Dev\Missions\05-FPPN\05-Analyses\resultats"

DATE_ARRETE = "2026-08-31"   # date d'arrêté de la mission
ANNEE_DEBUT = 2019
CHIFFRES_CITES = {           # entretien du 16/10/2026 (sinistralité décès citée)
    2019: 0.755,
    2021: 0.21,
    2026: 1.13,
}
SEUIL_MONTANT_ANORMAL = 5    # un montant > 5 fois la médiane de son groupe est signalé
HYPOTHESES = [
    "Sinistralité = prestations de l'année ÷ cotisations de l'année (définition à confirmer avec le FPPN).",
    "Deux versions sont calculées : par année de paiement (vision comptable) et par année de survenance (vision technique).",
    "Chain Ladder : appliqué aux montants payés cumulés par année de survenance, si au moins 4 années sont disponibles.",
    "Les dossiers rejetés sont exclus des montants dus.",
]

# =============================================================================
import importlib, subprocess, sys
for _m, _p in {"pandas": "pandas", "openpyxl": "openpyxl", "numpy": "numpy"}.items():
    try:
        importlib.import_module(_m)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", _p])

import os
from datetime import datetime
import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ARRETE = pd.Timestamp(DATE_ARRETE)
RESULTATS = {}      # nom d'onglet -> (DataFrame, note)
JOURNAL = []        # messages


def note(msg):
    print(msg)
    JOURNAL.append(msg)


def lire(chemin, colonnes_dates=()):
    if not os.path.exists(chemin):
        note(f"Fichier absent : {chemin}")
        return None
    df = pd.read_excel(chemin) if chemin.lower().endswith((".xlsx", ".xls")) else pd.read_csv(chemin, sep=None, engine="python")
    df.columns = [str(c).strip() for c in df.columns]
    for c in colonnes_dates:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
    note(f"Lu : {os.path.basename(chemin)} ({len(df)} lignes)")
    return df


def num(s):
    return pd.to_numeric(s.astype(str).str.replace(r"[\s\u202f]", "", regex=True).str.replace(",", "."), errors="coerce")


# =============================================================================
# 1. QUALITÉ DES DONNÉES
# =============================================================================
def qualite(df, nom, obligatoires, dates_ordre, cle, montants):
    lignes = []
    n = len(df)
    for c in obligatoires:
        if c not in df.columns:
            lignes.append([nom, "Colonne absente", c, n, "Bloquant pour les analyses qui l'utilisent"])
        else:
            vides = int(df[c].isna().sum())
            if vides:
                lignes.append([nom, "Valeurs manquantes", c, vides, f"{vides / n:.1%} des lignes"])
    for avant, apres in dates_ordre:
        if avant in df.columns and apres in df.columns:
            k = int((df[apres] < df[avant]).sum())
            if k:
                lignes.append([nom, "Dates incohérentes", f"{apres} < {avant}", k, "À corriger avec le FPPN"])
    for c in [x for x in df.columns if x.startswith("Date")]:
        k = int((df[c] > ARRETE + pd.Timedelta(days=60)).sum())
        if k:
            lignes.append([nom, "Date postérieure à l'arrêté", c, k, "Saisie probablement erronée"])
    if cle in df.columns:
        k = int(df[cle].duplicated(keep=False).sum())
        if k:
            lignes.append([nom, "Doublons d'identifiant", cle, k, "Lignes partageant le même identifiant"])
    for c in montants:
        if c in df.columns:
            v = df[c]
            k = int((v < 0).sum())
            if k:
                lignes.append([nom, "Montants négatifs", c, k, ""])
            med = v[v > 0].median()
            if pd.notna(med) and med > 0:
                k = int((v > SEUIL_MONTANT_ANORMAL * med).sum())
                if k:
                    lignes.append([nom, "Montants anormalement élevés", c, k,
                                   f"> {SEUIL_MONTANT_ANORMAL} fois la médiane ({med:,.0f})".replace(",", " ")])
    if not lignes:
        lignes.append([nom, "Aucune anomalie détectée", "", 0, ""])
    return lignes


# =============================================================================
# 2. ANALYSES DÉCÈS
# =============================================================================
def analyses_deces(dc, cot):
    dc = dc.copy()
    for c in ["Montant_capital", "Montant_frais_funeraires", "Montant_paye", "Salaire_base"]:
        if c in dc.columns:
            dc[c] = num(dc[c])
    dc["Montant_du"] = dc.get("Montant_capital", 0).fillna(0) + dc.get("Montant_frais_funeraires", 0).fillna(0)
    statut = dc.get("Statut_dossier", pd.Series("", index=dc.index)).astype(str).str.lower()
    dc = dc[~statut.str.contains("rejet")]
    dc["Annee_survenance"] = dc["Date_deces"].dt.year
    dc["Annee_paiement"] = dc["Date_paiement"].dt.year
    dc["Paye"] = dc["Date_paiement"].notna() & (dc["Date_paiement"] <= ARRETE)

    # A1 sinistralité
    if cot is not None:
        cot = cot.copy()
        cot["Annee"] = num(cot["Annee"]).astype("Int64")
        cot["Cotisations_deces"] = num(cot["Cotisations_deces"])
        pay = dc[dc["Paye"]].groupby("Annee_paiement")["Montant_paye"].sum()
        surv = dc.groupby("Annee_survenance")["Montant_du"].sum()
        t = cot.set_index("Annee")[["Cotisations_deces"]].copy()
        t["Prestations payées dans l'année"] = pay
        t["Prestations dues des décès de l'année"] = surv
        t = t.fillna(0)
        t["Sinistralité (année de paiement)"] = t["Prestations payées dans l'année"] / t["Cotisations_deces"]
        t["Sinistralité (année de survenance)"] = t["Prestations dues des décès de l'année"] / t["Cotisations_deces"]
        t = t.replace([np.inf, -np.inf], np.nan).reset_index()
        RESULTATS["A1 Sinistralité"] = (t, "Sinistralité = prestations ÷ cotisations décès. Deux visions : paiement (comptable) et survenance (technique). Les décès récents sont incomplets (dossiers non encore déclarés ou payés).")
        # A7 comparaison
        comp = []
        for a, v in CHIFFRES_CITES.items():
            ligne = t[t["Annee"] == a]
            calc_p = ligne["Sinistralité (année de paiement)"].iloc[0] if len(ligne) else np.nan
            calc_s = ligne["Sinistralité (année de survenance)"].iloc[0] if len(ligne) else np.nan
            comp.append([a, v, calc_p, calc_s, calc_p - v if pd.notna(calc_p) else np.nan])
        RESULTATS["A7 Chiffres cités"] = (pd.DataFrame(comp, columns=["Année", "Sinistralité citée", "Recalculée (paiement)", "Recalculée (survenance)", "Écart (paiement)"]),
                                          "Comparaison avec les chiffres cités en entretien. Un écart important appelle une explication de la méthode utilisée par le FPPN.")
    else:
        note("A1 et A7 non réalisées : fichier des cotisations absent.")

    # A2 fréquence par bénéficiaire, grade et âge
    dc["Age_deces"] = ((dc["Date_deces"] - dc.get("Date_naissance_decede", dc.get("Date_naissance_policier"))).dt.days / 365.25).round(0)
    dc["Tranche_age"] = pd.cut(dc["Age_deces"], [0, 18, 30, 40, 50, 55, 60, 70, 120], right=False,
                               labels=["< 18", "18-29", "30-39", "40-49", "50-54", "55-59", "60-69", "70 et +"])
    f1 = dc.pivot_table(index="Annee_survenance", columns="Beneficiaire_decede", values="N_dossier", aggfunc="count", fill_value=0) if "Beneficiaire_decede" in dc else None
    f2 = dc.pivot_table(index="Tranche_age", columns="Annee_survenance", values="N_dossier", aggfunc="count", fill_value=0, observed=False)
    RESULTATS["A2 Nombre de décès"] = (f1.reset_index() if f1 is not None else pd.DataFrame(), "Nombre de décès par année de survenance et par personne décédée.")
    RESULTATS["A2b Décès par âge"] = (f2.reset_index(), "Nombre de décès par tranche d'âge. Pour un taux de décès, rapporter à l'effectif exposé (fichier des mutualistes).")
    if "Grade" in dc.columns:
        f3 = dc.pivot_table(index="Grade", columns="Annee_survenance", values="N_dossier", aggfunc="count", fill_value=0)
        RESULTATS["A2c Décès par grade"] = (f3.reset_index(), "Nombre de décès par grade et par année de survenance.")

    # A3 coût moyen
    g = dc.groupby(["Annee_survenance", "Beneficiaire_decede"] if "Beneficiaire_decede" in dc else ["Annee_survenance"])
    a3 = g["Montant_du"].agg(["count", "mean", "median", "min", "max", "sum"]).reset_index()
    a3.columns = list(a3.columns[:-6]) + ["Nombre", "Coût moyen", "Coût médian", "Minimum", "Maximum", "Total"]
    RESULTATS["A3 Coût moyen"] = (a3, "Montant dû (capital + frais funéraires) par dossier. Comparer aux barèmes en vigueur chaque année.")

    # A4 délais
    rows = []
    for nom, d1, d2 in [("Décès → déclaration", "Date_deces", "Date_declaration"),
                        ("Déclaration → dossier complet", "Date_declaration", "Date_dossier_complet"),
                        ("Dossier complet → paiement", "Date_dossier_complet", "Date_paiement"),
                        ("Décès → paiement", "Date_deces", "Date_paiement")]:
        if d1 in dc and d2 in dc:
            j = (dc[d2] - dc[d1]).dt.days.dropna()
            j = j[j >= 0]
            if len(j):
                rows.append([nom, len(j), j.mean(), j.quantile(.25), j.median(), j.quantile(.75), j.quantile(.9), j.max()])
    RESULTATS["A4 Délais"] = (pd.DataFrame(rows, columns=["Étape", "Dossiers", "Moyenne (jours)", "1er quartile", "Médiane", "3e quartile", "90e centile", "Maximum"]),
                              "Délais en jours calendaires. Le 90e centile sert à estimer la part des décès d'une année encore inconnus à la clôture (tardifs).")

    # A5 stock non payé
    np_ = dc[~dc["Paye"]]
    a5 = np_.groupby("Annee_survenance").agg(Dossiers=("N_dossier", "count"), Montant_du=("Montant_du", "sum")).reset_index()
    if "Statut_dossier" in np_:
        a5b = np_.pivot_table(index="Annee_survenance", columns="Statut_dossier", values="N_dossier", aggfunc="count", fill_value=0).reset_index()
        a5 = a5.merge(a5b, on="Annee_survenance", how="left")
    RESULTATS["A5 Stock non payé"] = (a5, f"Dossiers non payés au {ARRETE:%d/%m/%Y}, par année du décès. À comparer à la provision pour prestations à payer au bilan.")

    # A6 triangle et Chain Ladder
    p = dc[dc["Paye"]].copy()
    p["Developpement"] = p["Annee_paiement"] - p["Annee_survenance"]
    tri = p.pivot_table(index="Annee_survenance", columns="Developpement", values="Montant_paye", aggfunc="sum", fill_value=0).sort_index()
    tri = tri.loc[tri.index >= ANNEE_DEBUT] if len(tri) else tri
    if len(tri) >= 4:
        cum = tri.cumsum(axis=1)
        derniere = ARRETE.year
        for a in cum.index:  # masque les cellules futures
            for d in cum.columns:
                if a + d > derniere:
                    cum.loc[a, d] = np.nan
        facteurs = []
        for j in range(len(cum.columns) - 1):
            c0, c1 = cum.columns[j], cum.columns[j + 1]
            m = cum[c1].notna() & cum[c0].notna()
            facteurs.append(cum.loc[m, c1].sum() / cum.loc[m, c0].sum() if cum.loc[m, c0].sum() else 1.0)
        ult = []
        for a in cum.index:
            ligne = cum.loc[a].dropna()
            dern_dev = ligne.index.max()
            val = ligne.iloc[-1]
            f = np.prod(facteurs[list(cum.columns).index(dern_dev):]) if dern_dev in cum.columns else 1
            ult.append([a, val, f, val * f, val * f - val])
        cl = pd.DataFrame(ult, columns=["Année de survenance", "Payé cumulé", "Facteur jusqu'à l'ultime", "Charge ultime estimée", "Reste à payer estimé (dont tardifs)"])
        cum_out = cum.reset_index().rename(columns={"Annee_survenance": "Survenance \\ Développement"})
        fac = pd.DataFrame([facteurs], columns=[f"{cum.columns[j]}→{cum.columns[j + 1]}" for j in range(len(facteurs))])
        RESULTATS["A6 Triangle"] = (cum_out, "Montants payés cumulés, par année du décès (lignes) et nombre d'années après le décès (colonnes).")
        RESULTATS["A6b Facteurs"] = (fac, "Facteurs de développement Chain Ladder : f(j) = Σ C(i,j+1) ÷ Σ C(i,j), sur les années connues.")
        RESULTATS["A6c Chain Ladder"] = (cl, "Charge ultime = payé cumulé × produit des facteurs restants. Reste à payer = ultime − payé. Résultat fragile si l'historique est court ou si les règles ont changé (2021, 2023).")
    else:
        note("A6 non réalisée : moins de 4 années de survenance avec des paiements.")
    return dc


# =============================================================================
# 3. ANALYSES SANTÉ
# =============================================================================
def analyses_sante(sa, mut):
    sa = sa.copy()
    for c in ["Montant_facture", "Montant_pris_en_charge", "Montant_paye"]:
        if c in sa.columns:
            sa[c] = num(sa[c])
    sa["Annee"] = sa["Date_soins"].dt.year
    # S1 coût par bénéficiaire
    s1 = sa.groupby(["Annee", "Lien"] if "Lien" in sa else ["Annee"]).agg(
        Factures=("N_facture", "count"), Beneficiaires=("Beneficiaire", "nunique"), Montant_paye=("Montant_paye", "sum")).reset_index()
    s1["Coût par bénéficiaire"] = s1["Montant_paye"] / s1["Beneficiaires"]
    RESULTATS["S1 Coût santé"] = (s1, "Coût santé par année et par lien avec le mutualiste. Comparer à la cotisation maladie (3,5 % du salaire de base).")
    # S2 concentration prestataires
    s2 = sa.groupby("Prestataire")["Montant_paye"].sum().sort_values(ascending=False).reset_index()
    s2["Part"] = s2["Montant_paye"] / s2["Montant_paye"].sum()
    s2["Part cumulée"] = s2["Part"].cumsum()
    RESULTATS["S2 Prestataires"] = (s2, "Montants payés par prestataire, du plus élevé au plus faible. Une forte concentration oriente les tests de factures.")
    # S3 doublons
    cles = [c for c in ["Beneficiaire", "Prestataire", "Date_soins", "Montant_facture"] if c in sa]
    d = sa[sa.duplicated(subset=cles, keep=False)].sort_values(cles)
    RESULTATS["S3 Doublons possibles"] = (d, "Factures ayant le même bénéficiaire, prestataire, date de soins et montant. À vérifier une par une.")
    # S4 éligibilité
    if mut is not None and "Matricule" in sa:
        mut = mut.copy()
        connus = set(mut["Matricule"].astype(str))
        hors = sa[~sa["Matricule"].astype(str).isin(connus)]
        RESULTATS["S4 Hors fichier"] = (hors, "Factures dont le matricule est absent du fichier des mutualistes : bénéficiaires potentiellement non éligibles.")
        if {"Date_sortie", "Matricule"} <= set(mut.columns):
            mut["Date_sortie"] = pd.to_datetime(mut["Date_sortie"], errors="coerce", dayfirst=True)
            m = sa.merge(mut[["Matricule", "Date_sortie", "Statut"]].astype({"Matricule": str}), left_on=sa["Matricule"].astype(str), right_on="Matricule", how="left", suffixes=("", "_mut"))
            apres = m[m["Date_sortie"].notna() & (m["Date_soins"] > m["Date_sortie"])]
            RESULTATS["S5 Soins après sortie"] = (apres, "Soins postérieurs à la date de sortie du mutualiste (retraite non réinscrite, décès, radiation).")
    else:
        note("S4 et S5 non réalisées : fichier des mutualistes absent.")
    # S7 factures sans prise en charge
    if "N_prise_en_charge" in sa:
        sans = sa[sa["N_prise_en_charge"].isna() | (sa["N_prise_en_charge"].astype(str).str.strip() == "")]
        RESULTATS["S7 Sans prise en charge"] = (sans, "Factures payées sans numéro de prise en charge signée par le médecin.")
    # S6 délais de facturation
    j = (sa["Date_facture"] - sa["Date_soins"]).dt.days.dropna()
    j = j[j >= 0]
    if len(j):
        RESULTATS["S6 Délais santé"] = (pd.DataFrame([[len(j), j.mean(), j.median(), j.quantile(.9), j.max()]],
                                                   columns=["Factures", "Moyenne (jours)", "Médiane", "90e centile", "Maximum"]),
                                      "Délai entre les soins et la réception de la facture. Sert à estimer les soins réalisés mais non encore facturés à la clôture.")


# =============================================================================
# 4. ÉCRITURE DES RÉSULTATS
# =============================================================================
def ecrire(chemin, qualite_lignes):
    BN, GC, GB = "112A40", "F2F2F2", "BBBCBC"
    bord = Border(*(Side(style="thin", color=GB),) * 4)
    with pd.ExcelWriter(chemin, engine="openpyxl") as xw:
        meth = pd.DataFrame({"Élément": ["Date d'arrêté", "Généré le"] + [f"Hypothèse {i + 1}" for i in range(len(HYPOTHESES))] + [f"Journal {i + 1}" for i in range(len(JOURNAL))],
                             "Valeur": [f"{ARRETE:%d/%m/%Y}", f"{datetime.now():%d/%m/%Y %H:%M}"] + HYPOTHESES + JOURNAL})
        onglets = {"Méthode": (meth, "Hypothèses, paramètres et journal d'exécution."),
                   "Qualité données": (pd.DataFrame(qualite_lignes, columns=["Fichier", "Contrôle", "Champ", "Nombre", "Commentaire"]),
                                       "À traiter avant d'interpréter les résultats. Toute anomalie doit être expliquée ou corrigée avec le FPPN.")}
        onglets.update(RESULTATS)
        for nom, (df, txt) in onglets.items():
            df = df if df is not None else pd.DataFrame()
            df.to_excel(xw, sheet_name=nom[:31], index=False, startrow=3)
            ws = xw.sheets[nom[:31]]
            ws["A1"] = nom
            ws["A1"].font = Font(name="Helvetica", size=14, bold=True)
            ws["A2"] = txt
            ws["A2"].font = Font(name="Helvetica", size=9, italic=True, color="53565A")
            for j, c in enumerate(df.columns, 1):
                cell = ws.cell(4, j)
                cell.font = Font(name="Helvetica", size=10, bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor=BN)
                cell.alignment = Alignment(wrap_text=True)
                ws.column_dimensions[get_column_letter(j)].width = max(12, min(40, len(str(c)) + 4))
                serie = df.iloc[:, j - 1] if len(df) else None
                pct = any(k in str(c) for k in ["Sinistralité", "Part", "Écart"])
                for i in range(5, 5 + len(df)):
                    cl = ws.cell(i, j)
                    cl.font = Font(name="Helvetica", size=9)
                    cl.border = bord
                    if i % 2 == 0:
                        cl.fill = PatternFill("solid", fgColor=GC)
                    if isinstance(cl.value, (int, float)):
                        cl.number_format = "0.0%" if pct else ("# ##0" if abs(cl.value) >= 100 else "0.00")
                    elif isinstance(cl.value, datetime):
                        cl.number_format = "dd/mm/yyyy"
            ws.freeze_panes = "A5"


def lancer():
    note(f"Analyses FPPN, arrêté au {ARRETE:%d/%m/%Y}")
    dc = lire(FICHIER_DECES, ["Date_naissance_policier", "Date_naissance_decede", "Date_deces", "Date_declaration", "Date_dossier_complet", "Date_paiement"])
    sa = lire(FICHIER_SANTE, ["Date_naissance_beneficiaire", "Date_soins", "Date_facture", "Date_paiement"])
    cot = lire(FICHIER_COTISATIONS)
    mut = lire(FICHIER_MUTUALISTES, ["Date_naissance", "Date_entree", "Date_sortie"])
    q = []
    if dc is not None:
        for c in ["Montant_capital", "Montant_frais_funeraires", "Montant_paye"]:
            if c in dc:
                dc[c] = num(dc[c])
        q += qualite(dc, "Décès", ["N_dossier", "Matricule", "Date_deces", "Date_declaration", "Montant_capital", "Statut_dossier", "Beneficiaire_decede"],
                     [("Date_deces", "Date_declaration"), ("Date_declaration", "Date_dossier_complet"), ("Date_dossier_complet", "Date_paiement"), ("Date_deces", "Date_paiement")],
                     "N_dossier", ["Montant_capital", "Montant_frais_funeraires", "Montant_paye"])
        analyses_deces(dc, cot)
    else:
        note("Analyses décès non réalisées : fichier des sinistres absent.")
    if sa is not None:
        for c in ["Montant_facture", "Montant_paye"]:
            if c in sa:
                sa[c] = num(sa[c])
        q += qualite(sa, "Santé", ["N_facture", "Matricule", "Beneficiaire", "Prestataire", "Date_soins", "Montant_paye"],
                     [("Date_soins", "Date_facture"), ("Date_facture", "Date_paiement")], "N_facture", ["Montant_facture", "Montant_paye"])
        analyses_sante(sa, mut)
    else:
        note("Analyses santé non réalisées : fichier santé absent.")
    os.makedirs(DOSSIER_SORTIE, exist_ok=True)
    sortie = os.path.join(DOSSIER_SORTIE, f"Analyses_sinistres_FPPN_{datetime.now():%Y%m%d_%H%M}.xlsx")
    ecrire(sortie, q)
    print(f"\nRésultats : {sortie}")


if __name__ == "__main__":
    if len(sys.argv) >= 2:   # option : python analyses_sinistres_fppn.py <dossier_donnees> <dossier_sortie>
        DOSSIER = sys.argv[1]
        FICHIER_DECES = os.path.join(DOSSIER, "sinistres_deces.xlsx")
        FICHIER_SANTE = os.path.join(DOSSIER, "prestations_sante.xlsx")
        FICHIER_COTISATIONS = os.path.join(DOSSIER, "cotisations_par_annee.xlsx")
        FICHIER_MUTUALISTES = os.path.join(DOSSIER, "mutualistes.xlsx")
    if len(sys.argv) >= 3:
        DOSSIER_SORTIE = sys.argv[2]
    lancer()
