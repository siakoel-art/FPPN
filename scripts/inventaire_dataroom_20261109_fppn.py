"""Inventaire de la dataroom FPPN du 09/11/2026 comparé aux fichiers déjà reçus.

Entrées :
  dataroom/inventaire_20261109/FPPN_inventaire_dataroom_20261109.csv  (sortie du script PowerShell chez l'utilisateur)
  dataroom/FPPN_fichiers_deja_recus.csv                               (fichiers déjà reçus, avec empreinte SHA-256)
Sorties (dataroom/inventaire_20261109/) :
  FPPN_Inventaire_dataroom_20261109.xlsx   inventaire détaillé, synthèse, liste à transmettre
  FPPN_a_transmettre_20261109.csv          liste lue par FPPN_preparer_envoi_20261109.ps1

Règles de classement :
  1. contenu identique : l'empreinte SHA-256 figure dans les fichiers déjà reçus ;
  2. même chemin : la fin du chemin (2 à 4 derniers éléments, nom compris) correspond à un fichier déjà reçu ;
     on distingue alors « taille différente » (vraie mise à jour) et « même taille » (même document, empreinte différente) ;
  3. sinon : nouveau.
La comparaison sur le nom seul est écartée : des noms génériques (CNI.pdf, Scan_0001.pdf) reviennent des centaines de fois.
"""
import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

RAC = Path(__file__).resolve().parent.parent / "dataroom"
DOS = RAC / "inventaire_20261109"
INV = DOS / "FPPN_inventaire_dataroom_20261109.csv"
REF = RAC / "FPPN_fichiers_deja_recus.csv"
XLSX = DOS / "FPPN_Inventaire_dataroom_20261109.xlsx"
LISTE = DOS / "FPPN_a_transmettre_20261109.csv"
RACINES = {"03-Adenka-Audit": r"D:\03-Adenka-Audit",
           "05-Adenka Audit Juridique et Fiscal": r"D:\05-Adenka Audit Juridique et Fiscal"}
MO = 2 ** 20

# Thème et priorité par motif de sous-dossier (premier motif trouvé)
THEMES = [
    (r"Balance générale et Grand livre|GRAND LIVRE 2025", "Comptabilité : balances et grands livres", "P1"),
    (r"Credit aux mutualistes|ETAT FINANCIER GRENIER", "Grenier du Policier : crédits et états financiers", "P1"),
    (r"\\EDI\\|COMPTABLES & FINANCIERS\\AIRMS|COMPTABLES & FINANCIERS$|PROVISION POUR CONGES", "Comptabilité : états financiers, AIRMS et provisions", "P1"),
    (r"CYCLE PRESTATION ET SINISTRE|\\Prestations", "Prestations et sinistres", "P1"),
    (r"CYCLE PLACEMENT|ETAT DE RAPPROCHEMENT|Trésorerie et placements\\2026$", "Placements et trésorerie", "P1"),
    (r"Adhésion et cotisation|04-CENTRALISATION", "Adhésions, cotisations et épargne", "P1"),
    (r"DOCUMENTS JURIDIQUES|^PV ET AG|^$", "Juridique et gouvernance", "P1"),
    (r"ETAT 301", "Personnel : effectifs (état 301)", "P1"),
    (r"06-Inventaire", "Textes réglementaires", "P1"),
    (r"01-PBC|03-Memos", "Suivi de mission (PBC, mémos, logigrammes)", "P1"),
    (r"Immobilisations|Contrôle interne", "Immobilisations et contrôle interne", "P1"),
    (r"RA-GL-RE|RELEVE|RAPPROCHEMENT BANCAIRE", "Relevés et rapprochements bancaires", "P2"),
    (r"Revue budgétaire", "Budgets", "P2"),
    (r"Fiscalité|CNPS", "Fiscalité et obligations sociales", "P2"),
    (r"PAIE", "Paie du personnel", "P2"),
    (r"Systèmes d’information", "Systèmes d'information", "P3"),
    (r"CYCLE ACHAT", "Achats", "P3"),
    (r"07-Modèle PT", "Modèles de papiers de travail", "P3"),
]
LIB_PRIO = {"P1": "P1 : indispensable au volet actuariel", "P2": "P2 : utile (contrôles, rapprochements)",
            "P3": "P3 : non nécessaire au volet actuariel", "Doublon": "Doublon d'un autre fichier de la dataroom", "": ""}


def theme(dossier, sd):
    s = sd if dossier.startswith("03") else (sd if sd else "")
    for motif, nom, prio in THEMES:
        if re.search(motif, s):
            return nom, prio
    return "Autres", "P2"


def charger():
    d = pd.read_csv(INV, sep=";", dtype=str, keep_default_na=False, encoding="utf-8-sig")
    r = pd.read_csv(REF, sep=";", dtype=str, keep_default_na=False, encoding="utf-8-sig")
    d["taille"] = d.taille_octets.astype("int64")
    r["taille"] = r.taille_octets.astype("int64")
    d["rel"] = [f"{s}\\{f}" if s else f for s, f in zip(d.sous_dossier, d.fichier)]
    return d, r


def classer(d, r):
    par_sha = {}
    for _, x in r.iterrows():
        par_sha.setdefault(x.sha256.lower(), x)
    idx = {}
    for i, c in enumerate(r.chemin_origine.str.lower().str.replace("/", "\\")):
        seg = c.split("\\")
        for k in (2, 3, 4):
            if len(seg) >= k:
                idx.setdefault("\\".join(seg[-k:]), []).append(i)
    out = []
    for _, x in d.iterrows():
        res = {"statut_final": "Nouveau", "ref_envoi": "", "ref_chemin": "", "ref_taille": None}
        sha = x.sha256.lower()
        if sha and sha in par_sha:
            y = par_sha[sha]
            res.update(statut_final="Déjà reçu (contenu identique)", ref_envoi=y.envoi, ref_chemin=y.chemin_origine,
                       ref_taille=int(y.taille))
        else:
            p = x.rel.lower().split("\\")
            for k in (4, 3, 2):
                cle = "\\".join(p[-k:])
                if len(p) >= k and cle in idx:
                    sub = r.iloc[idx[cle]]
                    meme = sub[sub.taille == x.taille]
                    y = (meme if len(meme) else sub).iloc[0]
                    st = ("Même chemin, même taille, empreinte différente" if len(meme)
                          else "Version modifiée (même chemin, taille différente)")
                    res.update(statut_final=st, ref_envoi=y.envoi, ref_chemin=y.chemin_origine, ref_taille=int(y.taille))
                    break
        out.append(res)
    d = pd.concat([d, pd.DataFrame(out, index=d.index)], axis=1)
    th = [theme(a, b) for a, b in zip(d.dossier_dataroom, d.sous_dossier)]
    d["theme"] = [t[0] for t in th]
    pr = []
    for (nom, p), st, dbl in zip(th, d.statut_final, d.doublon_dataroom):
        if st.startswith("Déjà"):
            pr.append("")
        elif st.startswith("Même chemin"):
            pr.append("P3")
        elif dbl:
            pr.append("Doublon")
        else:
            pr.append(p)
    d["priorite"] = pr
    # Lots d'envoi : P1 tableurs et textes, P1 PDF et images, P2
    def lot(p, ext):
        if p == "P1":
            return "Lot A" if ext not in (".pdf", ".png", ".jpg", ".jpeg", ".tif", ".tiff") else "Lot B"
        return "Lot C" if p == "P2" else ""
    d["lot"] = [lot(p, e) for p, e in zip(d.priorite, d.extension)]
    return d


# ---------- Excel ----------
NOIR, GRIS, GRIS_CL, ORANGE = "060606", "D9D9D9", "F2F2F2", "EFAC62"
FIN = Side(style="thin", color="BFBFBF")
BORD = Border(left=FIN, right=FIN, top=FIN, bottom=FIN)
F = "Helvetica"
COUL_STATUT = {"Nouveau": "C6EFCE", "Version modifiée (même chemin, taille différente)": "FFEB9C",
               "Même chemin, même taille, empreinte différente": "EDEDED", "Déjà reçu (contenu identique)": "FFFFFF"}


def titre(ws, texte, sous=None):
    ws["A1"] = texte
    ws["A1"].font = Font(name=F, size=14, bold=True)
    if sous:
        ws["A2"] = sous
        ws["A2"].font = Font(name=F, size=9, italic=True, color="53565A")


def tableau(ws, ligne, entetes, lignes, largeurs=None, formats=None, couleur_col=None):
    for j, h in enumerate(entetes, 1):
        c = ws.cell(ligne, j, h)
        c.font = Font(name=F, bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor=NOIR)
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = BORD
    for i, row in enumerate(lignes, ligne + 1):
        for j, v in enumerate(row, 1):
            c = ws.cell(i, j, v)
            c.font = Font(name=F, size=10, bold=bool(row and str(row[0]).startswith("Total")))
            c.border = BORD
            if formats and formats.get(j):
                c.number_format = formats[j]
            if (i - ligne) % 2 == 0:
                c.fill = PatternFill("solid", fgColor=GRIS_CL)
        if couleur_col:
            v = row[couleur_col - 1]
            if v in COUL_STATUT:
                ws.cell(i, couleur_col).fill = PatternFill("solid", fgColor=COUL_STATUT[v])
    if largeurs:
        for j, w in enumerate(largeurs, 1):
            ws.column_dimensions[get_column_letter(j)].width = w
    return ligne + len(lignes) + 1


def synthese(wb, d):
    ws = wb.active
    ws.title = "Synthèse"
    titre(ws, "Inventaire de la dataroom FPPN du 09/11/2026",
          "Source : inventaire des dossiers « 03-Adenka-Audit » et « 05-Adenka Audit Juridique et Fiscal » (empreintes SHA-256), "
          "comparé aux fichiers déjà reçus par Adenka (envois du 15/10/2026, lots du 22/10/2026 et fichiers transmis)")
    l = 4
    ws.cell(l, 1, "1. Répartition par statut").font = Font(name=F, bold=True, size=11)
    ordre = ["Nouveau", "Version modifiée (même chemin, taille différente)",
             "Même chemin, même taille, empreinte différente", "Déjà reçu (contenu identique)"]
    lig = []
    for dos in sorted(d.dossier_dataroom.unique()):
        s = d[d.dossier_dataroom == dos]
        for st in ordre:
            x = s[s.statut_final == st]
            if len(x):
                lig.append([dos, st, len(x), round(x.taille.sum() / MO, 1)])
    lig.append(["Total", "", len(d), round(d.taille.sum() / MO, 1)])
    l = tableau(ws, l + 1, ["Dossier", "Statut", "Fichiers", "Taille (Mo)"], lig, [38, 52, 12, 14],
                {3: "# ##0", 4: "# ##0.0"}, couleur_col=2) + 1
    ws.cell(l, 1, "2. Fichiers à transmettre à Adenka, par priorité et par lot").font = Font(name=F, bold=True, size=11)
    lig = []
    for (p, lt), x in d[d.priorite != ""].groupby(["priorite", "lot"]):
        lib = {"Lot A": "Lot A : tableurs, textes et présentations", "Lot B": "Lot B : PDF et images",
               "Lot C": "Lot C"}.get(lt, "Non transmis")
        lig.append([LIB_PRIO[p], lib, len(x), round(x.taille.sum() / MO, 1)])
    l = tableau(ws, l + 1, ["Priorité", "Lot d'envoi", "Fichiers", "Taille (Mo)"], lig, None, {3: "# ##0", 4: "# ##0.0"}) + 1
    ws.cell(l, 1, "3. Fichiers nouveaux ou modifiés par thème").font = Font(name=F, bold=True, size=11)
    lig = []
    a = d[d.statut_final.isin(ordre[:2])]
    for th, x in a.groupby("theme"):
        n = x[x.statut_final == "Nouveau"]
        m = x[x.statut_final != "Nouveau"]
        lig.append([th, x.priorite.iloc[0] if x.priorite.iloc[0] != "Doublon" else x.priorite.mode().iloc[0],
                    len(n), len(m), round(x.taille.sum() / MO, 1)])
    lig.sort(key=lambda r: (r[1], -r[4]))
    tableau(ws, l + 1, ["Thème", "Priorité", "Nouveaux", "Versions modifiées", "Taille (Mo)"], lig, None,
            {3: "# ##0", 4: "# ##0", 5: "# ##0.0"})
    l = ws.max_row + 2
    notes = [
        "Lecture des statuts :",
        "Nouveau : aucun fichier déjà reçu n'a ce contenu ni ce chemin.",
        "Version modifiée : un fichier déjà reçu a le même chemin, mais une taille différente. C'est une mise à jour à analyser.",
        "Même chemin, même taille, empreinte différente : très probablement le même document retéléchargé. Il n'est pas à transmettre.",
        "Déjà reçu : contenu strictement identique (même empreinte SHA-256) à un fichier déjà reçu.",
        "Doublon : fichier présent plusieurs fois dans la dataroom. Seul le premier exemplaire est transmis.",
    ]
    for i, t in enumerate(notes):
        c = ws.cell(l + i, 1, t)
        c.font = Font(name=F, size=9, bold=(i == 0))
    ws.column_dimensions["E"].width = 14


def feuille_liste(wb, nom, d, sous):
    ws = wb.create_sheet(nom)
    titre(ws, nom, sous)
    cols = [("Dossier dataroom", "dossier_dataroom", 26), ("Sous-dossier", "sous_dossier", 60), ("Fichier", "fichier", 48),
            ("Extension", "extension", 9), ("Taille (octets)", "taille", 14), ("Date de modification", "date_modification", 16),
            ("Statut", "statut_final", 40), ("Thème", "theme", 34), ("Priorité", "priorite", 9), ("Lot d'envoi", "lot", 10),
            ("Déjà reçu dans", "ref_envoi", 24), ("Chemin du fichier déjà reçu", "ref_chemin", 60),
            ("Taille déjà reçue (octets)", "ref_taille", 14), ("Doublon dans la dataroom", "doublon_dataroom", 50),
            ("Empreinte SHA-256", "sha256", 30), ("Remarque", "remarque", 20)]
    lig = [[(None if (v := x[c]) is None or (isinstance(v, float) and pd.isna(v)) else v) for _, c, _ in cols]
           for _, x in d.iterrows()]
    tableau(ws, 4, [c[0] for c in cols], lig, [c[2] for c in cols], {5: "# ##0", 13: "# ##0"}, couleur_col=7)
    ws.freeze_panes = "D5"
    ws.auto_filter.ref = f"A4:{get_column_letter(len(cols))}{4 + len(lig)}"


def main():
    d, r = charger()
    d = classer(d, r)
    d = d.sort_values(["dossier_dataroom", "sous_dossier", "fichier"]).reset_index(drop=True)
    wb = Workbook()
    synthese(wb, d)
    t = d[d.lot != ""].sort_values(["lot", "dossier_dataroom", "sous_dossier", "fichier"])
    feuille_liste(wb, "À transmettre", t, "Fichiers à transmettre à Adenka (priorités P1 et P2), dans l'ordre des lots d'envoi")
    feuille_liste(wb, "Inventaire complet", d, "Tous les fichiers des deux dossiers, avec leur statut et le fichier déjà reçu correspondant")
    wb.save(XLSX)
    t.assign(chemin_relatif=t.rel)[["lot", "priorite", "dossier_dataroom", "chemin_relatif", "taille_octets", "sha256"]].to_csv(
        LISTE, sep=";", index=False, encoding="utf-8-sig")
    print(d.groupby(["statut_final"]).taille.agg(["size", "sum"]).assign(sum=lambda x: (x["sum"] / MO).round(1)))
    print(t.groupby("lot").taille.agg(["size", "sum"]).assign(sum=lambda x: (x["sum"] / MO).round(1)))
    print(d.groupby(["priorite", "statut_final"]).size())


if __name__ == "__main__":
    sys.exit(main())
