# ---- v12, étape 2 : contenu (exécuté par patch_v12.py, variables s, rep, rep_re, tranche, bloc disponibles)
FLOW = "« FLOW_CHART_COTISATION_ET_PRODUITS_DEPARGNE_1.pptx » (reçu le 23/10/2026)"


rep('SRC_ENT14 = "Source : entretien du 19/10/2026 (14h)', 'FLOW = "' + FLOW.replace('"', '\\"') + '"\nSRC_ENT14 = "Source : entretien du 19/10/2026 (14h)')


def repall(a, b):
    global s
    assert s.count(a) >= 1, a[:100]
    s = s.replace(a, b)


# ------------------------------------------------------------------ « retraité » réservé aux policiers à la retraite ; lecture corrigée selon le PCMS
for a_, b_ in [("comme le FPPN, puis retraitée", "comme le FPPN, puis corrigée selon le PCMS"),
               ("Reclassements de la lecture retraitée", "Reclassements de la lecture corrigée selon le PCMS"),
               ('"Tiroir retraité"', '"Tiroir selon le PCMS"'),
               ("Le classement retraité ne modifie", "Le classement corrigé selon le PCMS ne modifie"),
               ("la lecture retraitée", "la lecture corrigée selon le PCMS"),
               ("Lecture retraitée :", "Lecture corrigée selon le PCMS :"),
               ('"lecture retraitée"', '"lecture corrigée selon le PCMS"'),
               ('"Liquidité retraitée"', '"Liquidité corrigée selon le PCMS"'),
               ('"Équilibre stable retraité"', '"Équilibre stable corrigé selon le PCMS"'),
               ("comme le FPPN et retraités", "comme le FPPN et corrigés selon le PCMS"),
               ("**Retraitée, la liquidité", "**Corrigée selon le PCMS, la liquidité"),
               ("**Retraité, l'équilibre stable", "**Corrigé selon le PCMS, l'équilibre stable"),
               ("2025, True)} retraitée ;", "2025, True)} corrigée selon le PCMS ;"),
               ("2025, True)} retraité ;", "2025, True)} corrigé selon le PCMS ;"),
               ("Retraité : on ajoute aux emplois", "Corrigé selon le PCMS : on ajoute aux emplois")]:
    repall(a_, b_)
rep('RT["REC"]["statut_global"] + " (retraités)"', 'RT["REC"]["statut_global"]')
rep("le recouvrement n'est mesurable que pour les retraités", "le recouvrement n'est mesurable que pour les retraités précomptés par la CGRAE")

# ------------------------------------------------------------------ résumé exécutif : constats datés du 23/10/2026, évolutions visibles
rep('tab("Principaux constats au 15/10/2026")', 'tab("Principaux constats au 23/10/2026")')
rep('"S/C " + pc("S/C maladie 2021", s[2021]["sc_maladie"]) + " (2021), " + pc("S/C maladie 2025", s[2025]["sc_maladie"]) + " (2025)"',
    '"S/C " + pc("S/C maladie 2021", s[2021]["sc_maladie"]) + " (2021), " + pc("S/C maladie 2025", s[2025]["sc_maladie"]) + " (2025), " + pc("S/C maladie 2026 (rés)", s[2026]["sc_maladie"]) + " (31/08/2026)"')
rep('F("Fonds propres 2025", b[2025]["fonds_propres"]) + " (2025)"',
    'F("Fonds propres 2025", b[2025]["fonds_propres"]) + " (2025), " + F("Fonds propres 2026 (rés)", b[2026]["fonds_propres"]) + " (31/08/2026)"')
rep('source="Source : analyses Adenka sur les documents reçus le 15/10/2026 et la documentation publique de l\'AIRMS")',
    'source="Source : analyses Adenka sur les documents reçus du 15/10/2026 au 23/10/2026 et les textes de l\'UEMOA")')
tranche('gab(49, {"titre": "Quatre indicateurs clés"', 'contenu({"titre": "Fonctionnement du FPPN"', bloc("bloc_indicateurs.py"))
rep('''def val_ratio(x, a, retraite=False):
    v = (x.get("valeurs_retraitees") if retraite else x["valeurs"]).get(a)
    if v is None:
        return "n.d."
    lab = f"{x['code']} {a}{' retraité' if retraite else ''}"
    if x["unite"] == "%":
        return pc(lab, v)
    if x["unite"] == "mois":
        return N(lab, v, 2) + " mois"
    if x["unite"] == "jours":
        return N(lab, v, 0) + " jours"
    if x["unite"] == "FCFA":
        return F(lab, v)
    return N(lab, v, 2)


contenu({"titre": "Conformité prudentielle en bref"''', 'contenu({"titre": "Conformité prudentielle en bref"')
tranche('contenu({"titre": "Conformité prudentielle en bref"', 'gab(107, {"titre": "Quatre taux à surveiller"', bloc("bloc_conformite_bref.py"))
rep('''     "effets": ["Coût réel des sinistres inconnu", "Engagements mal mesurés au bilan", "Manque de liquidités en cas de retraits", "Retards et pertes sur les prêts"]},''',
    '''     "effets": ["Coût réel des sinistres inconnu", "Engagements mal mesurés au bilan", "Manque de liquidités en cas de retraits", "Retards et pertes sur les prêts"],
     "lecture": "Chaque colonne relie une cause (en haut) à l'effet qu'elle produit (en bas). Les quatre causes sont établies par les données. Deux effets sont chiffrés en partie 9 : le manque de liquidités (adossement) et les pertes sur les prêts (dépréciation du Grenier). Les deux autres, le coût réel des sinistres et la mesure des engagements par âge, attendent la base des prestations datée et les dates de naissance des mutualistes."},''')

# ------------------------------------------------------------------ synthèse des constats : fait, norme ou référence, écart chiffré, effet, criticité
tranche("# ---- v11 : synthèse des constats (résumé exécutif)", '_syn_diapo(2, SYN2, "sinistralité, placements, données, gouvernance, Grenier, personnel")\n', bloc("bloc_definitions.py") + bloc("bloc_constats.py"))

# ------------------------------------------------------------------ partie 10 : ratios par garantie avant les ratios globaux, définitions avant chaque tableau de bord
tranche("# ---- v11 : ratios par garantie et définition du résultat technique", 'contenu({"titre": "Affectation des excédents"', "")
rep('contenu({"titre": "Tableau de bord (2/4) : risques longs"', '_defs_diapo(2, ("RL1", "RL2", "RL3", "RL4", "RL5", "RL6", "RL7"))\ncontenu({"titre": "Tableau de bord (2/4) : risques longs"')
rep('contenu({"titre": "Tableau de bord (3/4) : gestion"', '_defs_diapo(3, ("CHF", "CHB", "CUP", "CHP", "FID", "DPP", "FOR", "CRB", "REC"))\ncontenu({"titre": "Tableau de bord (3/4) : gestion"')
tranche("def ligne_bord(x, annees=(2022, 2023, 2024, 2025, 2026)):", 'contenu({"titre": "Tableau de bord (1/4) : solvabilité et réserves"', '''def ligne_bord(x, annees=(2020, 2021, 2022, 2023, 2024, 2025, 2026)):
    vals = []
    for a in annees:
        v_ = val_ratio(x, a) if a in x["valeurs"] else "n.d."
        vals.append(v_ if v_ == "n.d." else bv(v_, boule_statut(x["statuts"].get(a, ""))))
    nom = x["nom"].replace(" (risques longs)", "")
    return [str(x["priorite"]), nom, x["norme"].replace(" (tableau 1) ; 9 mois au § 1.2.2.5.4", "").replace(" (marge disponible au moins égale à la marge exigée)", ""), *vals, x["statut_global"]]


bord_ent = ["N°", "Ratio", "Norme", "2020", "2021", "2022", "2023", "2024", "2025", "31/08/2026", "Statut"]
bord_larg = [0.8, 6.3, 4.0, 2.56, 2.56, 2.56, 2.56, 2.56, 2.56, 2.56, 2.18]
LEG_BORD = [("🔴", "Non conforme", "la valeur ne respecte pas la norme"), ("🟢", "Conforme", "la valeur respecte la norme"),
            ("⚪", "À apprécier", "norme non chiffrée"), ("🟡", "En attente", "donnée du FPPN attendue")]
''')
rep('contenu({"titre": "Tableau de bord (1/4) : solvabilité et réserves"', bloc("bloc_garanties.py") + '\n_defs_diapo(1, ("SIN", "MSE", "ROB", "FEP", "LIQ", "EQS", "AUT", "EXN", "TFP"))\ncontenu({"titre": "Tableau de bord (1/4) : solvabilité et réserves"')
rep_re(r'"entetes": bord_ent, "largeurs": bord_larg, "statut_col": 8, "taille": 9,', '"entetes": bord_ent, "largeurs": bord_larg, "statut_col": 10, "taille": 8.5, "legende": LEG_BORD,', 3)
rep('''[[str(RT["REC"]["priorite"]), RT["REC"]["nom"], RT["REC"]["norme"], "", "", "", pc("REC CGRAE", Q["effectifs"]["recouvrement_cgrae"]), "", RT["REC"]["statut_global"]]]''',
    '''[[str(RT["REC"]["priorite"]), RT["REC"]["nom"] + " (retraités précomptés par la CGRAE, février 2024 à août 2026)", RT["REC"]["norme"], "", "", "", "", "", "", bv(pc("REC CGRAE", Q["effectifs"]["recouvrement_cgrae"]), "🟢"), RT["REC"]["statut_global"]]]''')
rep('"synthese": "Ratios de priorité actuarielle 1 à 9, lecture comme le FPPN ; 2026 est indicatif"',
    '"synthese": "Ratios de priorité actuarielle 1 à 9, lecture comme le FPPN ; 2020, 2021 et 2026 sont indicatifs"')

# ------------------------------------------------------------------ tableau 96 : valeur et boule chaque année
rep('contenu({"titre": "Ratios non conformes de 2020 au 31/08/2026"', '''def _t96(c, a):
    b_ = {"Conforme": "🟢", "Non conforme": "🔴"}.get(d13(c, "st", a), "")
    if c in RT and RT[c]["valeurs"].get(a) is not None:
        return bv(val_ratio(RT[c], a), b_)
    return b_ or "-"


contenu({"titre": "Ratios non conformes de 2020 au 31/08/2026"''')
rep('"lignes": [[NOM8[c], NORME8[c]] + [{"Conforme": "Oui", "Non conforme": "Non"}.get(d13(c, "st", a), "-") for a in A8] + [q13(c, 16)] for c in CODES8]},',
    '"legende": [("🔴", "Non conforme", "norme non respectée cette année-là"), ("🟢", "Conforme", "norme respectée cette année-là")],\n                 "lignes": [[NOM8[c], NORME8[c]] + [_t96(c, a) for a in A8] + [q13(c, 16)] for c in CODES8]},')
rep("« Oui » : conforme ; « Non » : non conforme ; « - » : non mesurable (pas d'exercice précédent).",
    "Chaque case donne la valeur du ratio et une boule : 🔴 non conforme, 🟢 conforme ; « - » : non mesurable (pas d'exercice précédent).")

# ------------------------------------------------------------------ garanties : barème de la CGRAE et barème tarifaire du FPPN
rep('["CGRAE", "Retraités", "Garantie santé, précompte sur pension", F("Moyen CGRAE", sum(c["moyen"] for c in CG) / len(CG))]]},',
    '["CGRAE", "Retraités", "Garantie santé, précompte sur pension : 2,5 % de (pension de base - MUGEFCI) + 5 000 FCFA (sous-officier), 8 000 FCFA (officier), 10 000 FCFA (commissaire de 2e classe à principal) ou 15 000 FCFA (commissaire divisionnaire à général)", F("Moyen CGRAE", sum(c["moyen"] for c in CG) / len(CG))]]},')
rep("Les montants moyens sont observés dans les fichiers de précomptes, pas tirés d'un barème (non reçu).",
    "Les montants moyens sont observés dans les fichiers de précomptes. La cotisation des retraités suit le barème de la CGRAE reçu le 23/10/2026.")
rep('''        analyse=["**Le FPPN couvre à la fois des risques courts et des risques longs.** La maladie se règle dans l'année ; l'EFC et le PEP sont versés des années après les cotisations.",''',
    '''        analyse=["**Le FPPN couvre à la fois des risques courts et des risques longs.** La maladie se règle dans l'année ; l'EFC et le PEP sont versés des années après les cotisations.",
                 "**Le barème tarifaire du FPPN fixe, pour chaque acte de la clinique (CMP), la part du FPPN à 80 % du tarif et celle du mutualiste à 20 %, ainsi qu'un plafond annuel de 1 500 000 FCFA par bénéficiaire et de 3 000 000 FCFA par famille.** Il donne aussi les tarifs proposés par le FPPN aux structures privées. [[Le document porte la mention « tarifications proposées » et n'est pas daté : sa date d'application reste à établir.]]",''')
rep('''        implication="Les risques longs (EFC, PEP) relèvent de règles prudentielles plus strictes : étude actuarielle tous les cinq ans, certification des comptes, réserve de trésorerie (partie 10).",
        source=SRC_PREC)''', '''        implication="Les risques longs (EFC, PEP) relèvent de règles prudentielles plus strictes : étude actuarielle tous les cinq ans, certification des comptes, réserve de trésorerie (partie 10).",
        source=SRC_PREC + " ; « Bar_me_tarif_du_FPPN.pdf » et « BAREME_CGRAE.xlsx » (reçus le 23/10/2026)")''')
rep("demandent les âges des retraités, leurs consommations et la cotisation appliquée.",
    "demandent les âges des retraités et leurs consommations. La cotisation, elle, est fixée par le barème de la CGRAE : 2,5 % de la pension de base nette de la MUGEFCI, plus 5 000 à 15 000 FCFA selon le grade.")
rep('["Retraités", "La garantie santé des retraités est-elle à vie ou limitée, et quelle est leur cotisation ?", _n(79), "Urgent"],',
    '["Retraités", "La garantie santé des retraités est-elle accordée à vie ou pour une durée limitée ? (Le barème de cotisation de la CGRAE a été reçu le 23/10/2026.)", _n(79), "Urgent"],')

# ------------------------------------------------------------------ circuit des précomptes : établi par le logigramme du FPPN
rep('''        analyse=["[[**Deux circuits coexistent.** Pour la maladie et le décès, la direction du personnel de la Police transmet directement le fichier aller à la Solde, sans plateforme de partage, et le FPPN ne reçoit que le retour. Pour les autres produits, le responsable des cotisations dépose le fichier aller sur la plateforme et l'entité concernée dépose le retour]] (" + ENT19 + ")."],''',
    '''        analyse=["**Deux circuits coexistent, et le logigramme du FPPN les confirme.** Pour la maladie et le décès, le service du personnel crée le fichier de précompte et calcule les cotisations, le Trésor prélève, et le FPPN reçoit le fichier retour, sur papier puis en numérique. Sans carte de santé FPPN, l'adhérent n'a pas de prestation. Pour l'EFC et le PEP, le service informatique du FPPN intègre la souscription dans la plateforme Easy money et transmet le fichier aller. Pour la CGRAE, le prélèvement porte sur la pension et la CGRAE vire les fonds selon son barème (" + ENT19 + " ; " + FLOW + ")."],''')
rep('''        source=SRC_ENT19 + " ; fichiers aller et retour reçus le 15/10/2026")
contenu({"titre": "Circuit des fichiers par produit"''', '''        source=SRC_ENT19 + " ; " + FLOW + " ; fichiers aller et retour reçus le 15/10/2026")
contenu({"titre": "Circuit des fichiers par produit"''')

# ------------------------------------------------------------------ Code CIMA : provision pour sinistres tardifs = art. 334-12 (circulaire n° 00230, reçue le 23/10/2026)
rep('["Code CIMA, art. 334-8 ; circulaire n° 00230/CIMA/CRCA/PDT/2005", "Provisions techniques ; méthode de la cadence des déclarations tardives", "Textes fournis le 19/10/2026"],',
    '["Code CIMA, art. 334-12 ; circulaire n° 00230/CIMA/CRCA/PDT/2005 du 24/10/2005", "Provision pour sinistres déclarés tardifs : méthode de la cadence des déclarations tardives", "Extrait du Code CIMA 2019 reçu le 23/10/2026 (pages 523 et 524)"],')
rep('"Code CIMA, art. 334-8 (référence technique)"', '"[[Code CIMA, art. 334-8]] (référence technique)"')
rep('"Source : CONVENTION DE PARTENARIAT ENTRE IPS-CGRAE ET LE FPPN.pdf ; Code CIMA, art. 334-8 ; "', '"Source : CONVENTION DE PARTENARIAT ENTRE IPS-CGRAE ET LE FPPN.pdf ; Code CIMA ; "')

# ------------------------------------------------------------------ « à apprécier » en blanc, comme dans les légendes des matrices
rep('("à apprécier", JAUNE_C, NOIR, "À apprécier", "norme non chiffrée"),', '("à apprécier", GRIS_C, NOIR, "À apprécier", "norme non chiffrée"),')

# ------------------------------------------------------------------ validité de la matrice « Conformité prudentielle en bref » (toutes les années)
rep('''_POS = {}
for i, x in enumerate(SL, 1):''', '''for x in SL:
    if x["type"] == "contenu" and x["entete"].get("titre") == "Conformité prudentielle en bref":
        x["valable"] = "balances de 2020 au 31/08/2026 (provisoire) ; statut global jugé de 2022 à 2025"
_POS = {}
for i, x in enumerate(SL, 1):''')
