# -*- coding: utf-8 -*-
"""
FPPN - Construction de la base d'analyse à partir de l'extraction (extraction_fppn.py)
======================================================================================

Entrée  : le dossier « donnees » de l'extraction (tableaux/, bruts/, journal_fichiers.csv).
Sortie  : des tables propres au format Parquet, une par famille de données :
  - precomptes.parquet : une ligne par mutualiste, poste, sens (aller / retour) et mois
                         (postes 936, 8Z6, 870, 998, 9X4, 9W3, 8Z7, retraités CGRAE) ;
  - effectifs_detail.parquet : fichiers maladie 936 avec grade, date de recrutement, salaire ;
  - sources_mois.parquet : pour chaque poste, sens et mois, le fichier retenu et les fichiers
                           concurrents (contrôle des doublons non identiques) ;
  - rejets.parquet : lignes non reprises et motif (en-têtes répétés, totaux, lignes vides).

Chaque ligne garde l'identifiant du fichier source (colonne src) pour la traçabilité.
"""
import csv
import io
import os
import re
import sys
import unicodedata
from collections import defaultdict
from datetime import datetime
from functools import lru_cache

import pandas as pd

csv.field_size_limit(2**31 - 1)

MOIS = {"JANVIER": 1, "JANV": 1, "FEVRIER": 2, "FEVIER": 2, "FEV": 2, "MARS": 3, "AVRIL": 4, "MAI": 5,
        "JUIN": 6, "JUILLET": 7, "JUIILLET": 7, "AOUT": 8, "SEPTEMBRE": 9, "SEPT": 9, "OCTOBRE": 10,
        "OCT": 10, "NOVEMBRE": 11, "DECEMBRE": 12}


def norm(s):
    return unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().upper().strip()


# ---------------------------------------------------------------------------
# Mois couvert par un fichier, d'après son nom et son dossier
# ---------------------------------------------------------------------------
def mois_fichier(dossier, fichier):
    D, F = norm(dossier), norm(os.path.splitext(fichier)[0])
    annees_f = re.findall(r"(20\d\d)", F)
    annees_d = re.findall(r"(?:^|\\)(20\d\d)(?:\\|$| )", D)
    trouves = set()
    for k, v in sorted(MOIS.items(), key=lambda x: -len(x[0])):
        if re.search(r"(?<![A-Z])" + k + r"(?![A-Z])", F):
            trouves.add(v)
    if "FEV-MARS" in F:
        trouves = {2, 3}
    if not trouves:
        m = re.search(r"(20\d\d)-(\d\d)(?!\d)", F)
        if m:
            return [(int(m.group(1)), int(m.group(2)))]
        m = re.search(r"(?:^|_)(\d\d)[-_](20\d\d)", F) or re.search(r"(?<!\d)(\d\d)(20\d\d)(?!\d)", F)
        if m and 1 <= int(m.group(1)) <= 12:
            return [(int(m.group(2)), int(m.group(1)))]
        m = re.search(r"\\(\d\d) (20\d\d)$", D)
        if m:
            return [(int(m.group(2)), int(m.group(1)))]
        dernier = D.split("\\")[-1]
        for k, v in MOIS.items():
            if re.search(r"(?<![A-Z])" + k + r"(?![A-Z])", dernier):
                trouves.add(v)
    an = int(annees_f[-1]) if annees_f else (int(annees_d[-1]) if annees_d else None)
    return [(an, m) for m in sorted(trouves)] if an else []


# ---------------------------------------------------------------------------
# Poste et sens d'une table
# ---------------------------------------------------------------------------
def poste_sens(dossier, fichier, feuille):
    D, F, S = norm(dossier), norm(fichier), norm(feuille)
    if "EFFECTIF POLICE" in D:
        return "936", "retour"
    if "ALLOCATION DECES" in D:
        return "8Z6", "retour"
    if "ESPACE SOLDE" in D:
        return None, None                     # octobre 2026 : hors période
    if "MEDOC" in F or "MEDOC" in D:
        return "8Z7", ("aller" if "ALLER" in D or "ALLER" in F else "retour")
    if "\\EFC\\FICHIER ALLER" in D or "\\PEP\\FICHIER ALLER" in D:
        if "PEP" in S or S.startswith("998"):
            return "998", "aller"
        if "EFC" in S or S.startswith("870"):
            return "870", "aller"
        return None, None
    if "\\EFC\\FICHIER RETOUR" in D:
        return "870", "retour"
    if "\\PEP\\FICHIER RETOUR" in D:
        return "998", "retour"
    if "CGRAE\\PRECOMTES" in D:
        return "CGRAE", ("aller" if "ALLER" in F or "FICHIERS ALLER" in D else "retour")
    if D.startswith("FPPN-GNA"):
        return "9W3", ("aller" if "ALLER" in F or "FICHIERS ALLER" in D else "retour")
    if "FPPN-SOLDE\\GRENIER" in D:
        retour = ("RETOUR" in D.split("\\")[-1]) or F.startswith(("RETOUR", "FICHIER RETOUR", "PRECOMPTE", "MTG"))
        return "9X4", ("retour" if retour else "aller")
    return None, None


# ---------------------------------------------------------------------------
# Lecture d'une table brute et repérage de l'en-tête
# ---------------------------------------------------------------------------
ALIAS = {
    "poste": ["POSTE", "LEBI", "CODE PRODUIT"],
    "reference": ["REFERENCE", "REF", "REFRENCE"],
    "org": ["ORG"],
    "matricule": ["MATRICULE", "MATRICU", "MATRIC", "MECANO", "MATRICULE_MILITAIRE"],
    "nom": ["NOM", "NOM ET PRENOMS", "NOM ET PRENOM", "NOM ET PRENOM(S)", "NOM & PRENOMS", "NOM ET PRENOMS"],
    "prenoms": ["PRENOMS"],
    "principal": ["PRINCIPAL"],
    "montant": ["MENSUEL", "MONTANT", "MONTANT PRECOMPTE", "MONTANT_PRELEVE", "FPPN", "ECHEANCE"],
    "debut": ["DEBUT", "DATE DE 1ER PRECOMPTE"],
    "fin": ["FIN"],
    "grade": ["GRADE", "GR"],
    "indice": ["IND"],
    "emploi": ["CODE, ET LIBELLE EMPLOI", "CODE ET LIBELLE EMPLOI", "LIBELLE EMPLOI"],
    "date_recrutement": ["DATE_RECRU"],
    "brut_mois": ["BRUT_MOIS"],
    "salaire_base": ["BASE"],
    "date_naissance": ["DATE_NAISSANCE"],
    "matricule_cgrae": ["MATRICULE CGRAE", "MATRICULE_CGRAE"],
    "decision_cgrae": ["N DE DECISION", "NUM_DECISION_CGRAE"],
    "categorie": ["CATEGORIE"],
    "observations": ["OBSERVATIONS", "OBSERVATION"],
    "statut": ["STATUT"],
    "annee": ["ANNEE"],
    "mois_col": ["MOIS"],
}
ALIAS_INV = {}
for k, vs in ALIAS.items():
    for v in vs:
        ALIAS_INV[re.sub(r"[^A-Z0-9]", "", norm(v))] = k


def cle_col(c):
    return ALIAS_INV.get(re.sub(r"[^A-Z0-9]", "", norm(c)))


def lire_lignes(chemin):
    """Lit un CSV (séparateur ; ou ,) en liste de listes de texte."""
    brut = open(chemin, "rb").read()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            texte = brut.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    premieres = texte[:5000]
    sep = ";" if premieres.count(";") >= premieres.count(",") else ","
    return list(csv.reader(io.StringIO(texte), delimiter=sep))


def trouver_entete(lignes, max_l=30):
    for i, l in enumerate(lignes[:max_l]):
        cles = [cle_col(c) for c in l]
        if "matricule" in cles and ("montant" in cles or "nom" in cles):
            return i, cles
    return None, None


RE_MAT = re.compile(r"^\s*(\d{4,7})(?:\.0+)?\s*([A-Z])?\s*$")
RE_ID_FPPN = re.compile(r"^\d{3}FPPN$")


def nettoyer_matricule(v):
    s = re.sub(r"[^A-Z0-9.]", "", norm(v))
    if RE_ID_FPPN.match(s):
        return s, s                      # identifiant interne au FPPN (sans mécano) : gardé, signalé
    m = RE_MAT.match(s)
    if not m:
        return s or None, None
    chiffres, lettre = m.group(1), m.group(2) or ""
    return chiffres + lettre, chiffres


def nombre(v):
    s = str(v).strip().replace("\u00a0", "").replace(" ", "")
    if s in ("", "-", "nan", "None"):
        return None
    s = s.replace("XOF", "")
    if re.match(r"^-?\d+,\d+$", s):
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


@lru_cache(maxsize=None)
def _date_txt(s):
    if not s or s.startswith("-"):
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%y", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def date_txt(v):
    return _date_txt(str(v).strip())



# ---------------------------------------------------------------------------
# Repérage des colonnes d'après leur contenu (en-têtes absents, faux ou décalés)
# ---------------------------------------------------------------------------
RE_MAT_STRICT = re.compile(r"^\d{5,7}(?:\.0+)?[A-Z]$")
RE_MECANO = re.compile(r"^\d{6}$")
RE_REF = re.compile(r"^[A-Z]{2,6}\d{3,4}[A-Z0-9]*$")
RE_NOM = re.compile(r"^[A-Z][A-Z' .\-&]{2,}$")
POSTES = {"936", "8Z6", "8.0Z6", "870", "998", "9X4", "9W3", "8Z7"}


def profil(lignes, debut, n=400):
    echant = [l for l in lignes[debut:debut + 3 * n] if any(c.strip() for c in l)][:n]
    ncol = max((len(l) for l in echant), default=0)
    prof = []
    for j in range(ncol):
        vals = [norm(l[j]).replace(" ", "") if j < len(l) else "" for l in echant]
        brut = [norm(l[j]) if j < len(l) else "" for l in echant]
        nv = [v for v in vals if v]
        if not nv:
            prof.append({}); continue
        k = len(nv)
        prof.append({
            "mat": sum(bool(RE_MAT_STRICT.match(v)) for v in nv) / k,
            "mec": sum(bool(RE_MECANO.match(v)) for v in nv) / k,
            "poste": sum(v in POSTES for v in nv) / k,
            "num": sum(nombre(v) is not None for v in nv) / k,
            "date": sum(date_txt(v) is not None and not RE_MECANO.match(v) for v in nv) / k,
            "ref": sum(bool(RE_REF.match(v)) for v in nv) / k,
            "nom": sum(bool(RE_NOM.match(b)) and " " in b for b in brut if b) / k,
            "rempli": k / max(len(vals), 1),
            "moy": (sum(nombre(v) or 0 for v in nv) / k),
        })
    return prof


def colonnes_par_contenu(lignes, debut):
    prof = profil(lignes, debut)
    if not prof:
        return None
    cles = [None] * len(prof)
    def meilleur(critere, seuil, exclure=()):
        cand = [(p.get(critere, 0), j) for j, p in enumerate(prof) if j not in exclure and cles[j] is None]
        cand = [c for c in cand if c[0] >= seuil]
        return max(cand)[1] if cand else None
    j = meilleur("mat", 0.6)
    if j is None:
        j = meilleur("mec", 0.6)
    if j is not None:
        cles[j] = "matricule"
    for crit, cle, seuil in (("poste", "poste", 0.8), ("nom", "nom", 0.5), ("ref", "reference", 0.6)):
        j = meilleur(crit, seuil)
        if j is not None:
            cles[j] = cle
    # montant : colonne numérique (hors matricule), la plus à droite du nom parmi les entières
    nums = [j for j, p in enumerate(prof) if cles[j] is None and p.get("num", 0) >= 0.8
            and p.get("date", 0) < 0.5 and p.get("moy", 0) >= 100]
    j_nom = cles.index("nom") if "nom" in cles else -1
    apres = [j for j in nums if j > j_nom] or nums
    if apres:
        cles[apres[0]] = "montant"
    dates = [j for j, p in enumerate(prof) if cles[j] is None and p.get("date", 0) >= 0.6]
    for j, cle in zip(dates, ("debut", "fin")):
        cles[j] = cle
    # org : petite colonne de 2 caractères avant le matricule
    if "matricule" in cles:
        jm = cles.index("matricule")
        for j in range(jm - 1, -1, -1):
            if cles[j] is None and prof[j]:
                cles[j] = "org"
                break
    if "matricule" not in cles and not ("nom" in cles and "montant" in cles):
        return None
    return cles


def entete_valide(lignes, i, cles):
    """L'en-tête est fiable si sa colonne matricule contient bien des matricules et sa colonne montant des nombres."""
    if "matricule" not in cles:
        return False
    prof = profil(lignes, i + 1, 200)
    jm = cles.index("matricule")
    if jm >= len(prof) or max(prof[jm].get("mat", 0), prof[jm].get("mec", 0)) < 0.6:
        return False
    if "montant" in cles:
        ja = cles.index("montant")
        if ja >= len(prof) or prof[ja].get("num", 0) < 0.8:
            return False
    return True

# ---------------------------------------------------------------------------
# PEP 998 : fichiers texte à largeur fixe (une ligne par mutualiste et par mois)
# ---------------------------------------------------------------------------
RE_PEP = re.compile(r"^\s*(\S+)\s+(\d{4,7}[A-Z]?)\s+(.*?)\s+(?:([A-Za-z][A-Za-z\-]{1,6} ?\d{3,6}\S*)\s+)?(998)\s+(\d+)\s+(\d\d/\d\d/\d{4})")


def lire_pep_txt(chemin, src):
    lignes, rejets = [], 0
    for l in open(chemin, encoding="latin-1"):
        m = RE_PEP.match(l)
        if not m:
            if l.strip() not in ("", "/"):
                rejets += 1
            continue
        d = _date_txt(m.group(7))
        mat, m6 = nettoyer_matricule(m.group(2))
        lignes.append({"poste": "998", "sens": "retour", "annee": d.year, "mois": d.month,
                       "org": m.group(1), "matricule": mat, "mat6": m6, "nom": m.group(3).strip(),
                       "reference": (m.group(4) or "").strip(), "montant": float(m.group(6)), "src": src})
    return lignes, rejets


# ---------------------------------------------------------------------------
def construire(dossier_donnees, dossier_sortie):
    os.makedirs(dossier_sortie, exist_ok=True)
    J = list(csv.DictReader(open(os.path.join(dossier_donnees, "journal_fichiers.csv"), encoding="utf-8-sig"),
                            delimiter=";"))
    feuilles = {}
    for f in csv.DictReader(open(os.path.join(dossier_donnees, "journal_feuilles.csv"), encoding="utf-8-sig"),
                            delimiter=";"):
        feuilles.setdefault(f["id"], []).append(f)

    tables = []         # (src, dossier, fichier, feuille, chemin, nb_feuilles)
    for j in J:
        if j["action"].startswith("Converti"):
            fs = feuilles.get(j["id"], [])
            for f in fs:
                tables.append((j["id"], j["dossier"], j["fichier"], f["feuille"],
                               os.path.join(dossier_donnees, "tableaux", f["csv"]), len(fs)))
        elif j["action"].startswith("Copié") and j["extension"] in (".csv", ".txt"):
            tables.append((j["id"], j["dossier"], j["fichier"], "",
                           os.path.join(dossier_donnees, j["sortie"].replace("\\", os.sep)), 1))

    # Feuilles avec un en-tête reconnu, par fichier (pour écarter les feuilles annexes sans en-tête)
    avec_entete = defaultdict(set)
    lectures = {}
    for src, dossier, fichier, feuille, chemin, nbf in tables:
        if fichier.lower().endswith(".txt") or poste_sens(dossier, fichier, feuille)[0] is None:
            continue
        lignes = lire_lignes(chemin)
        i, cles = trouver_entete(lignes)
        lectures[(src, feuille)] = (lignes, i, cles)
        if i is not None:
            avec_entete[src].add(feuille)

    lignes_pre, lignes_eff, rejets, inventaire = [], [], [], []
    for src, dossier, fichier, feuille, chemin, nbf in tables:
        poste, sens = poste_sens(dossier, fichier, feuille)
        if poste is None:
            continue
        if fichier.lower().endswith(".txt"):
            ls, nrej = lire_pep_txt(chemin, src)
            for x in ls:
                x["feuille"] = feuille
            lignes_pre += ls
            df = pd.DataFrame(ls)
            for (a, m), g in df.groupby(["annee", "mois"]):
                inventaire.append({"src": src, "dossier": dossier, "fichier": fichier, "feuille": feuille,
                                   "poste": poste, "sens": sens, "annee": a, "mois": m, "lignes": len(g),
                                   "montant": g["montant"].sum(), "lecture": "texte largeur fixe"})
            if nrej:
                rejets.append({"src": src, "fichier": fichier, "motif": "ligne texte non reconnue", "n": nrej})
            continue
        mois = mois_fichier(dossier, fichier)
        if len(mois) != 1 or not (2019 <= mois[0][0] <= 2026):
            rejets.append({"src": src, "fichier": fichier, "motif": f"mois non déterminé {mois}", "n": 0})
            continue
        an, mo = mois[0]
        lignes, i, cles = lectures[(src, feuille)]
        if i is not None and entete_valide(lignes, i, cles):
            debut, lecture = i + 1, "en-tête"
        else:
            generique = re.match(r"^(FEUIL|SHEET|FEUILLE)\s*\d*$", norm(feuille)) is not None
            if i is None and generique and avec_entete[src] - {feuille}:
                rejets.append({"src": src, "fichier": f"{fichier} | {feuille}",
                               "motif": "feuille annexe sans en-tête (autre feuille du fichier retenue)",
                               "n": len(lignes)})
                continue
            debut = (i + 1) if i is not None else 0
            cles_c = colonnes_par_contenu(lignes, debut)
            if cles_c is None:
                rejets.append({"src": src, "fichier": f"{fichier} | {feuille}", "motif": "colonnes non reconnues",
                               "n": len(lignes)})
                continue
            # on garde les colonnes de détail de l'en-tête quand elles existent
            if cles:
                cles_c = [c or (cles[k] if k < len(cles) and cles[k] not in cles_c else None)
                          for k, c in enumerate(cles_c)]
            cles, lecture = cles_c, "contenu"
        n_ok, total = 0, 0.0
        detail = "salaire_base" in cles or "date_recrutement" in cles
        feuille_pre, feuille_eff = [], []
        for l in lignes[debut:]:
            if not any(c.strip() for c in l):
                continue
            rec = {}
            for c, k in zip(l, cles):
                if k and k not in rec:
                    rec[k] = c
            mat_brut = rec.get("matricule", "")
            nm = norm(mat_brut)
            if nm.startswith("---") or nm in ("MATRICULE", "MATRICU", "MECANO", "TOTAL", "MATRICULE_MILITAIRE"):
                continue
            mat, m6 = nettoyer_matricule(mat_brut)
            mt = nombre(rec.get("montant", ""))
            nom = re.sub(r"\s+", " ", " ".join(x for x in (rec.get("nom", ""), rec.get("prenoms", "")) if x)).strip()
            if not m6:
                if mt is not None and nom and not nm:
                    mat, m6 = None, None          # ligne sans matricule mais avec nom et montant : gardée, signalée
                else:
                    if mt is not None or nm:
                        rejets.append({"src": src, "fichier": fichier, "motif": f"matricule invalide : {mat_brut[:20]}",
                                       "n": 1})
                    continue
            if mt is None and not nom:
                continue
            if not nom and (mat is None or not re.search(r"[A-Z]", mat)):
                rejets.append({"src": src, "fichier": fichier, "motif": "ligne de total ou sans nom", "n": 1,
                               "montant": mt})
                continue
            if re.match(r"^(TOTAL|SOUS[- ]?TOTAL)", norm(nom)):
                rejets.append({"src": src, "fichier": fichier, "motif": "ligne de total ou sans nom", "n": 1,
                               "montant": mt})
                continue
            r = {"poste": poste, "sens": sens, "annee": an, "mois": mo, "org": rec.get("org", "").strip(),
                 "matricule": mat, "mat6": m6, "nom": nom,
                 "reference": rec.get("reference", "").strip(), "montant": mt,
                 "principal": nombre(rec.get("principal", "")),
                 "debut": date_txt(rec.get("debut", "")), "fin": date_txt(rec.get("fin", "")),
                 "src": src, "feuille": feuille}
            if poste == "CGRAE":
                r.update({"grade": rec.get("grade", "").strip(), "categorie": rec.get("categorie", "").strip(),
                          "matricule_cgrae": rec.get("matricule_cgrae", "").strip(),
                          "date_naissance": date_txt(rec.get("date_naissance", "")),
                          "observations": rec.get("observations", "").strip()})
            feuille_pre.append(r)
            if detail:
                feuille_eff.append({"annee": an, "mois": mo, "matricule": mat, "mat6": m6, "nom": nom,
                                   "grade": rec.get("grade", "").strip(), "indice": nombre(rec.get("indice", "")),
                                   "emploi": rec.get("emploi", "").strip(),
                                   "date_recrutement": date_txt(rec.get("date_recrutement", "")),
                                   "brut_mois": nombre(rec.get("brut_mois", "")),
                                   "salaire_base": nombre(rec.get("salaire_base", "")),
                                   "cotisation_936": mt, "src": src})
        # Ligne de total : montant égal à la somme des autres lignes de la feuille (à 1 % près)
        somme = sum(r["montant"] or 0 for r in feuille_pre)
        garder = []
        for r in feuille_pre:
            mt = r["montant"] or 0
            reste = somme - mt
            if mt > 0 and reste > 0 and abs(mt - reste) <= 0.01 * reste and len(feuille_pre) > 2:
                rejets.append({"src": src, "fichier": fichier, "motif": "ligne de total ou sans nom", "n": 1,
                               "montant": mt})
                continue
            garder.append(r)
        mats = {r["matricule"] for r in garder}
        lignes_pre += garder
        lignes_eff += [e for e in feuille_eff if e["matricule"] in mats]
        n_ok = len(garder)
        total = sum(r["montant"] or 0 for r in garder)
        inventaire.append({"src": src, "dossier": dossier, "fichier": fichier, "feuille": feuille, "poste": poste,
                           "sens": sens, "annee": an, "mois": mo, "lignes": n_ok, "montant": total,
                           "lecture": lecture})

    pre = pd.DataFrame(lignes_pre)
    # Montants saisis deux fois à la suite dans le fichier source (ex. 4792447924 pour 47924) :
    # corrigés seulement pour le poste décès 8Z6, où un précompte de plus d'un million est impossible.
    def double(v):
        if v is None or pd.isna(v) or v < 1e6:
            return None
        t = str(int(v))
        h = len(t) // 2
        return float(t[:h]) if len(t) % 2 == 0 and t[:h] == t[h:] else None
    pre["montant_source"] = pre["montant"]
    pre["corrige"] = False
    m = pre["poste"].eq("8Z6")
    corr = pre.loc[m, "montant"].map(double)
    idx = corr.dropna().index
    pre.loc[idx, "montant"] = corr.loc[idx]
    pre.loc[idx, "corrige"] = True
    inv = pd.DataFrame(inventaire)
    if len(idx):
        tot = pre.loc[idx].groupby(["src", "feuille", "annee", "mois"])
        for (sr, fe, a, mo), g in tot:
            k = (inv["src"] == sr) & (inv["feuille"] == fe) & (inv["annee"] == a) & (inv["mois"] == mo)
            inv.loc[k, "montant"] -= (g["montant_source"] - g["montant"]).sum()
    inv = inv[inv["lignes"] > 0].copy()
    # Une seule feuille par poste, sens et mois : celle qui a le plus de lignes. Les autres restent dans
    # sources_mois pour comparaison (doublons non identiques).
    inv = inv.sort_values(["poste", "sens", "annee", "mois", "lignes", "montant"],
                          ascending=[True, True, True, True, False, False])
    inv["retenu"] = ~inv.duplicated(["poste", "sens", "annee", "mois"])
    cle = set(map(tuple, inv.loc[inv["retenu"], ["src", "feuille", "poste", "sens", "annee", "mois"]].values))
    pre["retenu"] = [k in cle for k in zip(pre["src"], pre["feuille"], pre["poste"], pre["sens"],
                                           pre["annee"], pre["mois"])]
    for c in ("debut", "fin", "date_naissance"):
        if c in pre:
            pre[c] = pd.to_datetime(pre[c], errors="coerce")
    pre.to_parquet(os.path.join(dossier_sortie, "precomptes.parquet"), index=False)
    inv.to_parquet(os.path.join(dossier_sortie, "sources_mois.parquet"), index=False)
    eff = pd.DataFrame(lignes_eff)
    if not eff.empty:
        eff["date_recrutement"] = pd.to_datetime(eff["date_recrutement"], errors="coerce")
    eff.to_parquet(os.path.join(dossier_sortie, "effectifs_detail.parquet"), index=False)
    pd.DataFrame(rejets).to_parquet(os.path.join(dossier_sortie, "rejets.parquet"), index=False)
    return pre, inv, pd.DataFrame(rejets)


if __name__ == "__main__":
    entree = sys.argv[1] if len(sys.argv) > 1 else "donnees"
    sortie = sys.argv[2] if len(sys.argv) > 2 else "base"
    pre, inv, rej = construire(entree, sortie)
    print(f"Lignes de précomptes : {len(pre)} ; retenues : {int(pre['retenu'].sum())}")
    print(f"Feuilles-mois : {len(inv)} ; retenues : {int(inv['retenu'].sum())} ; rejets : {len(rej)}")
