# -*- coding: utf-8 -*-
"""
FPPN - Exemples chiffrés des onglets du classeur (rubrique « Exemple chiffré »)
==============================================================================
Chaque exemple refait un calcul de l'onglet sur un vrai chiffre, pour un lecteur non technique.
Les montants viennent des résultats des scripts de la mission (mêmes valeurs que le rapport et que les formules).
"""


def fcfa(x):
    return f"{round(x):,}".replace(",", " ") + " FCFA"


def pct(x):
    return f"{100 * x:.2f}".replace(".", ",") + " %"


def nb(x, d=2):
    return f"{x:,.{d}f}".replace(",", " ").replace(".", ",")


def exemples(R, C, Q):
    s = {x["exercice"]: x for x in R["sinistralite"]}
    b = {x["exercice"]: x for x in R["bilan"]}
    a = {int(k): v for k, v in Q["agregats"].items()}
    ms = {m["exercice"]: m for m in Q["marge_solvabilite"]}
    ar = R["aller_retour"][0]
    cv = R["couverture"][0]
    ef = R["effectifs_mensuels"][-1]
    gp = R["grenier_par"]
    rp = {x["exercice"]: x for x in R["rendement_placements"]}
    rt = {x["code"]: x for x in Q["ratios"]}
    tc = {t["indicateur"]: t for t in C["tcam"]}
    t0 = C["tcam"][0]
    mor = {m["table"]: m for m in C["mortalite"]}
    mo = {m["cloture"]: m for m in Q["mobilisation"]}
    pm = {x["exercice"]: x for x in C["provisions_mois"]}
    E = {
        "1_couverture": f"Poste {cv['poste']} (EFC), sens {cv['sens']} : {cv['mois_recus']} mois reçus, {nb(cv['lignes'], 0)} lignes, {fcfa(cv['montant'])} demandés au total.",
        "1_aller_retour": f"Poste {ar['poste']} (EFC) : {ar['mois_communs']} mois communs ; {fcfa(ar['montant_retour'])} précomptés / {fcfa(ar['montant_aller'])} demandés = {pct(ar['taux_recouvrement'])}.",
        "2_effectifs_mensuels": f"En {ef['libelle']} : {nb(ef['cotisants'], 0)} policiers précomptés au poste 936 pour {fcfa(ef['montant'])} ; {nb(ef['retraites'], 0)} retraités précomptés par la CGRAE pour {fcfa(ef['montant_cgrae'])}.",
        "2_cotisations": f"En 2025 : cotisations maladie comptabilisées (compte 7010) = {fcfa(s[2025]['cot_maladie'])} ; cotisations décès (7021) = {fcfa(s[2025]['cot_deces'])}.",
        "3_sinistralite_resultat": f"En 2025 : S/C maladie = prestations maladie {fcfa(s[2025]['prest_maladie'])} / cotisations maladie {fcfa(s[2025]['cot_maladie'])} = {pct(s[2025]['sc_maladie'])}. Taux de frais = {fcfa(s[2025]['frais_gestion'])} / ({fcfa(s[2025]['cot_maladie'])} + {fcfa(s[2025]['cot_deces'])}) = {pct(s[2025]['taux_frais'])}.",
        "3_prestations_maladie_nature": f"Le total des natures d'une année redonne les prestations maladie de l'onglet 3_sinistralite_resultat ({fcfa(s[2025]['prest_maladie'])} en 2025).",
        "4_bilan": f"Au 31/08/2026 : couverture de l'épargne = actifs réservés {fcfa(b[2026]['actifs_dedies_epargne'])} / épargne due {fcfa(b[2026]['epargne_totale'])} = {pct(b[2026]['couverture_epargne'])}.",
        "4_pep_versements": "Cumul d'une année = cumul de l'année précédente + versements de l'année (colonne D).",
        "4_efc_versements": "Variation du compte = solde du compte 41195200 de l'année - solde de l'année précédente.",
        "5_grenier_encours": f"Au 31/08/2026 : {nb(gp['prets'], 0)} prêts en cours pour un encours de {fcfa(gp['encours'])}.",
        "5_grenier_retards": f"Au 31/08/2026 : la part de l'encours en retard de plus de 90 jours vaut {pct(gp['par90'])}.",
        "5_grenier_par_historique": f"PAR30 au 31/08/2026 = encours des prêts en retard de plus de 30 jours / encours total = {pct(gp['par30'])}.",
        "6_cgrae": f"Cotisation moyenne d'un mois = montant précompté / retraités ; en {ef['libelle']} : {fcfa(ef['montant_cgrae'])} / {nb(ef['retraites'], 0)} = {fcfa(ef['montant_cgrae'] / ef['retraites'])}.",
        "7_rendement_placements": f"En 2025 : {fcfa(rp[2025]['produits_financiers'])} de produits financiers / {fcfa(rp[2025]['placements_moyens'])} de placements moyens = {pct(rp[2025]['rendement'])}.",
        "8_inflation_reel": "Valeur de 2025 en francs de 2020 = valeur de 2025 × 100 / indice des prix de 2025.",
        "8_tcam": f"{t0['indicateur']} : ({fcfa(t0['v1'])} / {fcfa(t0['v0'])}) puissance (1 / {nb(t0['annees'], 0)}) - 1 = {pct(t0['tcam'])} par an.",
        "8_tendances": f"Cotisants : pente de {nb(C['tendances'][0]['pente'], 0)}, soit environ {nb(C['tendances'][0]['pente'], 0)} cotisants de plus chaque mois en moyenne ; R² = {nb(C['tendances'][0]['r2'])} (la droite explique {pct(C['tendances'][0]['r2'])} des variations).",
        "8_sensibilite_retraits": f"Si 10 % de l'épargne était retirée : besoin = {fcfa(b[2026]['epargne_totale'])} × 10 % = {fcfa(b[2026]['epargne_totale'] * 0.1)}.",
        "8_sensibilite_pertes_grenier": "Provision = encours en retard de plus de 180 jours × taux de perte choisi (cellule orange).",
        "8_provisions_mois": f"En 2025 : provisions {fcfa(pm[2025]['provisions'])} / prestations d'un mois {fcfa(pm[2025]['prestations_mensuelles'])} = {nb(pm[2025]['mois_couverts'])} mois.",
        "8_tables_mortalite": f"Rente viagère à 60 ans (CIMA F, 3,5 %) = {nb(mor['CIMA F']['a60'])} : 100 000 FCFA par an versés à vie à partir de 60 ans valent {fcfa(100000 * mor['CIMA F']['a60'])} aujourd'hui.",
        "10_agregats": f"Cotisations acquises 2025 = moins le solde des comptes 70… - 7906 + 6906 = {fcfa(a[2025]['cotisations_acquises'])}.",
        "10_ratios": f"Taux de sinistralité 2025 = charges de prestations {fcfa(a[2025]['charges_prestations'])} / cotisations acquises {fcfa(a[2025]['cotisations_acquises'])} = {pct(rt['SIN']['valeurs']['2025'])}, hors de la fourchette de 70 % à 80 % : « Non conforme ».",
        "10_ratios_detail": f"SIN 2025 : numérateur {fcfa(a[2025]['charges_prestations'])}, dénominateur {fcfa(a[2025]['cotisations_acquises'])}, valeur {nb(rt['SIN']['valeurs']['2025'], 4)}.",
        "10_marge_solvabilite": f"2025 : marge exigée = le plus grand de 18 % des cotisations 2024 ({fcfa(ms[2025]['exigee_18pc_cotisations_n_1'])}) et de 26 % des prestations maladie moyennes ({fcfa(ms[2025]['exigee_26pc_prest_maladie_3ans'])}) = {fcfa(ms[2025]['marge_exigee'])}. Ratio comme le FPPN = {fcfa(ms[2025]['marge_disponible_comme_fppn'])} / {fcfa(ms[2025]['marge_exigee'])} = {nb(ms[2025]['ratio_comme_fppn'])} ; ratio net des pertes = {nb(ms[2025]['ratio_net'])}.",
        "10_actif_mobilisation": f"Au 31/08/2026 : actif mobilisable immédiatement {fcfa(mo['31/08/2026']['immediat'])} ; en moins d'un an {fcfa(mo['31/08/2026']['moins_1_an'])}.",
        "12_psap_rapprochement": f"En 2025 : provision maladie {fcfa(b[2025]['provision_maladie'])} / (prestations maladie {fcfa(s[2025]['prest_maladie'])} / 12) = {nb(b[2025]['provision_maladie'] / (s[2025]['prest_maladie'] / 12))} mois de prestations.",
        "12_sinistralite_variantes": "Taux sans MEDOC = numérateur / (dénominateur - cotisations MEDOC du compte 7022).",
    }
    return E
