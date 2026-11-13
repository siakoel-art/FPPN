# -*- coding: utf-8 -*-
"""
FPPN - Test des intérêts de l'épargne sur les policiers sortis (13/11/2026)
==========================================================================
Question : à la sortie, le FPPN rend-il plus que ce que le policier a versé (intérêts payés sans être comptabilisés)
ou exactement ce qu'il a versé (promesse d'intérêts non tenue) ?

Sources
- Reçu : débits nominatifs des comptes 41195100 (PEP) et 41195200 (EFC) des grands livres des comptes 2020 au 31/08/2026
  (gl_epargne_fppn.py). Les lignes du journal « RAN » sont exclues : elles recopient le détail de l'exercice précédent.
- Versé : précomptes retour de la Solde, poste 998 (PEP) et 870 (EFC) (base precomptes.parquet), plus les dépôts
  en espèces nominatifs crédités au même compte.

Règles
1. Rapprochement par le nom : le nom écrit après « / » dans le libellé est comparé au nom des fichiers de précomptes
   du même produit. Égalité exacte, sinon (libellé coupé à 35 caractères) nom de précompte unique qui commence par
   le nom du grand livre, sinon mêmes mots dans un autre ordre. Un nom qui correspond à plusieurs matricules est écarté.
2. Policier sorti : au moins un remboursement (REMB, REGLT, SOLDE, RESIL, acompte) et aucun précompte plus de deux mois
   après le dernier paiement. Un policier qui n'a fait que des rachats partiels n'est pas sorti.
3. Historique complet : le premier précompte est postérieur au premier mois couvert du produit (PEP 03/2001, EFC 01/2023)
   et aucun mois n'est manquant entre le premier précompte et la sortie. Sinon le « versé » est incomplet et le cas
   n'est pas concluant.
4. Égalité : écart (reçu - versé) inférieur ou égal à 1 000 FCFA ou à 0,5 % du versé.
5. Intérêts promis (pour comparer) : intérêt simple non capitalisé, PEP 3,5 % et EFC 5,5 % par an
   (« MANUEL DE PRODUIT D'EPARGNE FPPN VF VF.docx »), chaque précompte portant intérêt de la fin de son mois à la sortie.

Utilisation : python test_interets_epargne_fppn.py <dossier_base> <gl_epargne.parquet> <dossier_sortie>
Sorties : detail_nominatif.parquet (NE PAS diffuser), test_interets.json et test_interets.xlsx (agrégés).
"""
import json, os, re, sys, unicodedata
import pandas as pd

BASE, GLP, SORTIE = sys.argv[1:4]
PRODUITS = {"41195100": ("PEP", "998", 0.035, (2001, 3)), "41195200": ("EFC", "870", 0.055, (2023, 1))}
TOL_ABS, TOL_REL = 1000.0, 0.005


def norm(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().upper()
    s = re.sub(r"[^A-Z ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def mois_index(a, m):
    return int(a) * 12 + int(m) - 1


gl = pd.read_parquet(GLP)
gl = gl[gl.journal != "RAN"].copy()
gl["avant"] = gl.libelle.str.split("/").str[0].str.upper()
gl["nom_gl"] = gl.libelle.map(lambda l: norm(l.split("/", 1)[1]) if "/" in l else "")
gl["coupe"] = gl.libelle.str.len() >= 34
gl["date"] = pd.to_datetime(gl.date)
EST_SORTIE = re.compile(r"REMB(?! TROP)|REGLT|RGLT|SOLDE|RESIL|ACP|ACOMPTE")
EST_ANNUL = re.compile(r"ANNUL|REGUL|RETOUR|REJET|EXTOURNE")

pre = pd.read_parquet(os.path.join(BASE, "precomptes.parquet"),
                      columns=["poste", "sens", "annee", "mois", "matricule", "nom", "montant", "retenu"])
pre = pre[pre.retenu & (pre.sens == "retour") & pre.poste.isin(["998", "870"])].copy()
pre["nom_n"] = pre.nom.map(norm)
pre["mi"] = [mois_index(a, m) for a, m in zip(pre.annee, pre.mois)]

lignes, synthese = [], {}
for compte, (prod, poste, taux, debut) in PRODUITS.items():
    p = pre[pre.poste == poste]
    couverts = set(p.mi.unique())
    mi_debut = mois_index(*debut)
    # nom le plus fréquent par matricule, puis index nom -> matricules
    nom_mat = p.groupby("matricule").nom_n.agg(lambda s: s.value_counts().index[0])
    par_nom = nom_mat.reset_index().groupby("nom_n").matricule.apply(list).to_dict()
    tri = {}
    for n, ms in par_nom.items():
        tri.setdefault(" ".join(sorted(n.split())), []).extend(ms)

    def apparier(n, coupe):
        if not n:
            return None, "sans nom"
        ms = par_nom.get(n, [])
        if len(ms) == 1:
            return ms[0], "exact"
        if len(ms) > 1:
            return None, "homonymes"
        if coupe:
            c = [m for k, v in par_nom.items() if k.startswith(n) for m in v]
            if len(c) == 1:
                return c[0], "début du nom"
            if len(c) > 1:
                return None, "homonymes"
        c = tri.get(" ".join(sorted(n.split())), [])
        if len(c) == 1:
            return c[0], "mots dans un autre ordre"
        return None, "homonymes" if c else "non trouvé"

    g = gl[gl.compte == compte].copy()
    noms = g[["nom_gl", "coupe"]].drop_duplicates()
    app = {(n, c): apparier(n, c) for n, c in zip(noms.nom_gl, noms.coupe)}
    g["matricule"] = [app[(n, c)][0] for n, c in zip(g.nom_gl, g.coupe)]
    g["rapprochement"] = [app[(n, c)][1] for n, c in zip(g.nom_gl, g.coupe)]
    deb = g[g.debit > 0]
    stat_app = deb.groupby("rapprochement").agg(ecritures=("debit", "size"), montant=("debit", "sum")).reset_index()

    gm = g[g.matricule.notna()]
    pm = p.groupby("matricule")
    for mat, e in gm.groupby("matricule"):
        d = e[e.debit > 0]
        if d.empty or not d.avant.str.contains(EST_SORTIE).any():
            continue
        cr = e[e.credit > 0]
        annul = cr[cr.avant.str.contains(EST_ANNUL)].credit.sum()
        especes = cr[~cr.avant.str.contains(EST_ANNUL)].credit.sum()
        recu = d.debit.sum() - annul
        date_sortie = d.date.max()
        mi_sortie = mois_index(date_sortie.year, date_sortie.month)
        x = pm.get_group(mat) if mat in pm.groups else p.iloc[0:0]
        apres = x[x.mi > mi_sortie + 2]
        x = x[x.mi <= mi_sortie]
        verse_pre = x.montant.sum()
        verse = verse_pre + especes
        premier = x.mi.min() if len(x) else None
        trous = [] if premier is None else [m for m in range(premier, mi_sortie + 1) if m not in couverts and m <= mois_index(2026, 8)]
        if len(apres):
            statut = "non sorti (précomptes après le paiement)"
        elif premier is None or premier <= mi_debut:
            statut = "historique incomplet (versements avant le début des fichiers)"
        elif trous:
            statut = "historique incomplet (mois manquants)"
        else:
            statut = "historique complet"
        # intérêts promis : chaque précompte de la fin de son mois à la date de sortie
        fin_mois = pd.to_datetime([f"{a}-{m:02d}-01" for a, m in zip(x.annee, x.mois)]) + pd.offsets.MonthEnd(0)
        jours = (date_sortie - fin_mois).days.to_numpy().clip(min=0)
        interets = float((x.montant.values * taux * jours / 365).sum())
        ecart = recu - verse
        tol = max(TOL_ABS, TOL_REL * verse)
        if ecart > tol:
            sens = "reçu > versé"
        elif ecart < -tol:
            sens = "reçu < versé"
        else:
            sens = "reçu = versé"
        lignes.append({"produit": prod, "matricule": mat, "nom_gl": e.nom_gl.iloc[0], "rapprochement": e.rapprochement.iloc[0],
                       "date_sortie": date_sortie.date().isoformat(), "paiements": int(len(d)), "recu": recu,
                       "verse_precomptes": verse_pre, "verse_especes": especes, "verse": verse, "ecart": ecart,
                       "interets_promis": interets, "ecart_sur_interets": ecart / interets if interets else None,
                       "premier_precompte": None if premier is None else f"{premier % 12 + 1:02d}/{premier // 12}",
                       "mois_precomptes": int(x.mi.nunique()), "statut": statut, "resultat": sens})
    synthese[prod] = {"rapprochement_debits": stat_app.to_dict("records"),
                      "debits_total": float(deb.debit.sum()), "debits_rapproches": float(deb[deb.matricule.notna()].debit.sum())}

df = pd.DataFrame(lignes)
df.to_parquet(os.path.join(SORTIE, "detail_nominatif.parquet"), index=False)

# Agrégats (aucune donnée nominative)
agg = df.groupby(["produit", "statut", "resultat"]).agg(policiers=("recu", "size"), recu=("recu", "sum"), verse=("verse", "sum"),
                                                     ecart=("ecart", "sum"), interets_promis=("interets_promis", "sum")).reset_index()
comp = df[df.statut == "historique complet"]
ann = comp.assign(annee=comp.date_sortie.str[:4]).groupby(["produit", "annee", "resultat"]).agg(
    policiers=("recu", "size"), recu=("recu", "sum"), verse=("verse", "sum"), ecart=("ecart", "sum"),
    interets_promis=("interets_promis", "sum")).reset_index()
q = comp[comp.resultat == "reçu > versé"].ecart_sur_interets.dropna()
OUT = {"regles": {"tolerance_fcfa": TOL_ABS, "tolerance_relative": TOL_REL, "taux_pep": 0.035, "taux_efc": 0.055},
       "rapprochement": synthese, "par_statut": agg.to_dict("records"), "complets_par_annee": ann.to_dict("records"),
       "ecart_sur_interets_promis_quartiles": q.quantile([0.25, 0.5, 0.75]).to_dict() if len(q) else {}}
json.dump(OUT, open(os.path.join(SORTIE, "test_interets.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=float)
with pd.ExcelWriter(os.path.join(SORTIE, "test_interets.xlsx")) as xw:
    agg.to_excel(xw, sheet_name="par_statut", index=False)
    ann.to_excel(xw, sheet_name="complets_par_annee", index=False)
    pd.DataFrame([dict(produit=k, **r) for k, v in synthese.items() for r in v["rapprochement_debits"]]).to_excel(xw, sheet_name="rapprochement", index=False)
pd.set_option("display.width", 250)
print(pd.DataFrame([dict(produit=k, **r) for k, v in synthese.items() for r in v["rapprochement_debits"]]).to_string())
print(agg.to_string())
print(ann.to_string())
print("écart / intérêts promis (reçu > versé, complets) :", OUT["ecart_sur_interets_promis_quartiles"])
