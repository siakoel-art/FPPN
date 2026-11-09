# -*- coding: utf-8 -*-
"""Résumé exécutif Word du seul volet actuariel et statistique, pour les dirigeants du FPPN (8 à 12 pages).

Même plan que le résumé exécutif général de la mission : contexte et objectifs, approche et limites, résultats,
synthèse et plan d'actions priorisé, annexes courtes. Rédaction au standard Adenka : conclusion en gras, puis preuves
chiffrées, puis ce que cela veut dire concrètement. Contenu commun avec le résumé général (contenu_actuariel_fppn.py).

Usage : python3 resume_actuariel_dirigeants_fppn.py <sortie.docx>
"""
import os
import re
import sys

ICI = os.path.dirname(os.path.abspath(__file__))
SRC = open(os.path.join(ICI, "resume_executif_fppn.py"), encoding="utf-8").read()
# imports, constantes et constats du résumé exécutif existant
exec(SRC[:SRC.index("# ================================================================ DOCUMENT")], globals())
SORTIE = sys.argv[1]
import contenu_actuariel_fppn as CA  # noqa: E402

JAUNE = []


def H(texte):
    """Retire les balises [[…]] et retient le texte à surligner en jaune."""
    for m in re.findall(r"\[\[(.+?)\]\]", texte):
        JAUNE.append(m)
    return texte.replace("[[", "").replace("]]", "")


r = RapportAdenka(titre="Résumé exécutif du volet actuariel", client="FPPN", type_livrable="Rapport provisoire", date="Novembre 2026")
exec(SRC[SRC.index("def cellule("):SRC.index("def fiche(")], globals())
r.sigles({"AIRMS": "Agence ivoirienne de régulation de la mutualité sociale", "BCEAO": "Banque centrale des États de l'Afrique de l'Ouest",
          "CGRAE": "Caisse générale de retraite des agents de l'État", "CIMA": "Conférence interafricaine des marchés d'assurances",
          "DISA": "Déclaration individuelle des salaires annuels à la CNPS", "CNPS": "Caisse nationale de prévoyance sociale",
          "EFC": "Entraide fin de carrière (épargne versée au départ à la retraite)", "FPPN": "Fonds de prévoyance de la Police nationale",
          "PEP": "Plan Épargne Plus (épargne volontaire)", "UEMOA": "Union économique et monétaire ouest-africaine"})
r.abreviations({"art.": "Article", "FCFA": "Franc de la Communauté financière africaine", "n°": "Numéro",
                "P1 à P4": "Niveaux de priorité : P1 dans les 3 mois, P2 dans les 6 mois, P3 de 6 à 18 mois, P4 au-delà de 18 mois"})
RAP = CA.RAP

# ================================================================ messages clés
r.synthese("MESSAGES CLÉS")
r.para("**Le FPPN connaît ses dettes de court terme, mais il ne peut pas encore démontrer qu'il pourra rendre à temps l'épargne de tous les policiers.** "
       "L'épargne, environ 15,7 milliards de FCFA (15 743 516 391 FCFA) au 31/08/2026, a surtout financé des terrains, une clinique et des prêts, "
       "qui ne se transforment pas vite en argent. Les intérêts promis aux épargnants ne sont pas comptabilisés, et le FPPN ne respecte pas "
       "plusieurs règles de solidité fixées par la réglementation.")
r.tableau(["Question de la mission", "Réponse", "En clair"],
          [[q, f"{b} {v}", t] for q, b, v, t, _ in CA.VERDICTS],
          titre="Réponses aux quatre questions du volet actuariel", source=RAP + ", diapositives 4 à 8.", largeurs_cm=[4.6, 3.6, 10.2], taille=9)
r.para("**Les quatre constats les plus graves.**", gras_auto=False)
r.puces(["Les intérêts promis aux épargnants, environ 3,3 milliards de FCFA (3 278 913 238 FCFA), ne sont pas comptabilisés : le FPPN doit plus que ce que montrent ses comptes ;",
         "L'argent disponible dans l'année couvre 34,42 % de ce que le FPPN devra payer dans l'année, épargne comprise ;",
         "Les placements ne représentent que 5,53 % de l'épargne et rapportent moins que les intérêts promis ;",
         "14 des 36 règles de solidité ne sont pas respectées, ce qui expose le FPPN à un retrait d'agrément."])
r.para("**Huit actions sont à lancer dans les trois mois.** Elles portent sur le taux et les intérêts de l'épargne, le calcul des provisions, "
       "la réserve de liquidité, le plan de mise en conformité et la transmission des données (partie 4).", gras_auto=False)

# ================================================================ 1. contexte et objectifs
r.partie(1, "Contexte et objectifs", introduction="Cette partie rappelle pourquoi la mission a été lancée et ce que le volet actuariel devait établir.")
r.h2("1.1. CONTEXTE DE LA MISSION")
r.para("**Le FPPN protège les policiers et leur famille, et gère leur épargne.** C'est une mutuelle sociale régie par le règlement n° 07/2009/CM/UEMOA. "
       "Il couvre les soins, verse une allocation au décès, gère deux produits d'épargne (EFC et PEP) et finance des prêts aux policiers par le Grenier. "
       "Il compte 31 198 policiers cotisants en août 2026.")
r.para("**L'audit comptable et financier couvre la gestion de septembre 2020 à août 2026.** Le FPPN l'a confié à Adenka Partners après une consultation "
       "du 12 octobre 2026. Le volet actuariel et statistique en est une partie : il évalue ce que le FPPN doit à ses adhérents et sa capacité à le payer.")
r.h2("1.2. CE QUE LE VOLET ACTUARIEL DEVAIT ÉTABLIR")
r.para("**Le volet actuariel devait répondre à quatre questions.** Il a aussi recalculé les 36 indicateurs de solidité de l'UEMOA, nécessaires pour y répondre.", gras_auto=False)
r.tableau(["Question", "Ce qu'il fallait établir"], [[f"{k}. {q}", t] for k, (q, t) in enumerate(CA.QUESTIONS, 1)],
          titre="Les quatre questions du volet actuariel", source="Offre technique Adenka validée par le FPPN ; " + RAP + ", diapositives 4 et 42.", largeurs_cm=[4.6, 13.8], taille=9)

# ================================================================ 2. approche et limites
r.partie(2, "Approche méthodologique et limites", introduction="Cette partie décrit comment les réponses ont été obtenues, et ce qui en limite la précision.")
r.h2("2.1. APPROCHE ACTUARIELLE")
r.para(CA.APPROCHE_INTRO)
r.puces([e.rstrip(" ;.") + (" ;" if k < len(CA.APPROCHE_ETAPES) - 1 else ".") for k, e in enumerate(CA.APPROCHE_ETAPES)])
r.tableau(["Analyse", "Méthode retenue", "Pourquoi ce choix"], [[a, b, c] for a, b, c, _ in CA.METHODES],
          titre="Méthodes retenues et raison du choix", source=RAP + ", diapositives 47, 48 et 245 à 252.", largeurs_cm=[3.6, 8.0, 6.8], taille=8.5)
r.h2("2.2. LIMITES DES TRAVAUX")
r.para("**" + CA.LIMITES_INTRO.split(". ", 1)[0] + ".** " + CA.LIMITES_INTRO.split(". ", 1)[1], gras_auto=False)
r.puces([e.rstrip(" ;.") + (" ;" if k < len(CA.LIMITES) - 1 else ".") for k, e in enumerate(CA.LIMITES)])
r.para(CA.LIMITES_EFFET)

# ================================================================ 3. résultats
r.partie(3, "Résultats des travaux", introduction="Cette partie donne la réponse à chaque question, avec les chiffres qui la fondent et ce qu'elle signifie pour le FPPN.")
ANS = ["2020", "2021", "2022", "2023", "2024", "2025", "31/08/2026"]
for k, (titre_q, msg, paras, renvoi) in enumerate(CA.RESULTATS, 1):
    r.h2(f"3.{k}. {titre_q.upper()}")
    r.para(f"**{msg}** " + H(paras[0]), gras_auto=False)
    r.para(H(paras[1]))
    if k == 1:
        graphique(r, {"type": "barres", "categories": ANS, "format": "0.00", "titre_axe": "Milliards de FCFA",
                      "series": [{"nom": "Épargne due aux policiers", "valeurs": [round(md(v), 2) for v in [10828205640, 10147421276, 9364845293, 9121576576, 10785064432, 13814066964, 15743516391]], "couleur": "112A40"},
                                 {"nom": "Argent réservé à cette épargne", "valeurs": [round(md(v), 2) for v in [714042784, 1957211038, 888844233, 839212069, 699925787, 1608155525, 2562885661]], "couleur": "EFAC62"}]},
                  "Épargne due et argent réservé pour la rembourser (milliards de FCFA)", RAP + ", diapositives 13 et 102.", hauteur_cm=6.5)
    if k == 3:
        graphique(r, {"type": "barres", "categories": ["Paiements attendus dans l'année", "Argent disponible en moins d'un an"], "format": "0.00", "titre_axe": "Milliards de FCFA",
                      "series": [{"nom": "Au 31/08/2026", "valeurs": [round(md(10046113184), 2), round(md(3457866951), 2)], "couleur": "112A40"}]},
                  "Paiements de l'année et argent disponible au 31/08/2026 (milliards de FCFA)", RAP + ", diapositives 7 et 165.", largeur_cm=12.0, hauteur_cm=6.0)
    r.para(renvoi + ".", gras_auto=False)
r.h2("3.5. RATIOS PRUDENTIELS : TRAVAUX DU VOLET ACTUARIEL")
r.para(CA.RATIOS_INTRO, gras_auto=False)
for x in CA.RATIOS_PARAS:
    r.para(x)
r.tableau(["Indicateur", "Norme", "Valeur 2025", "Statut"], [[a, b, c, f"{s} {'Conforme' if s == '🟢' else 'Non conforme'}"] for a, b, c, s in CA.RATIOS_TABLE],
          titre="Principaux indicateurs de solidité en 2025 (Décision n° 05/2022)", source=RAP + ", diapositives 19 et 178 à 227.", largeurs_cm=[5.4, 4.6, 5.0, 3.4], taille=9)
r.h2("3.6. AUTRES ENSEIGNEMENTS")
r.para("**La garantie maladie ne s'équilibre plus, et des données manquent pour finir les calculs.**", gras_auto=False)
r.puces([e.rstrip(" ;.") + (" ;" if k < len(CA.AUTRES_CONSTATS) - 1 else ".") for k, e in enumerate(CA.AUTRES_CONSTATS)])
graphique(r, {"type": "lignes", "categories": ANS, "format": "0.00", "titre_axe": "% des cotisations maladie", "axe_min": 0, "axe_max": 110,
              "series": [{"nom": "Soins rapportés aux cotisations maladie", "valeurs": [12.54, 6.00, 6.47, 26.91, 48.13, 73.53, 97.40], "couleur": "112A40"}]},
          "Part des cotisations maladie consommée par les soins (%)", RAP + ", diapositives 87 et 88.", hauteur_cm=6.0)

# ================================================================ 4. synthèse et plan d'actions
r.partie(4, "Synthèse et plan d'actions priorisé", introduction="Cette partie classe les constats du plus critique au moins critique, puis présente les actions à mener, qui agit et dans quel délai.")
r.h2("4.1. CONSTATS CLASSÉS PAR CRITICITÉ")
r.para("**Quatre constats sont critiques : l'épargne, la liquidité, les placements et la conformité.** Chaque constat est formulé simplement, "
       "avec ce qu'il veut dire concrètement pour le FPPN et le chiffre qui le fonde.")
r.tableau(["Criticité", "Constat", "Ce que cela veut dire concrètement", "Chiffre clé", "Rapport"],
          [[f"{b} {CA.CRIT_LIB[b]}", f"**{code}** {txt}", sens, H(chif), f"Diap. {dt}"] for b, code, txt, sens, chif, dt in CA.CONSTATS],
          titre="Constats du volet actuariel, du plus critique au moins critique", source=RAP + ", diapositives 4 à 13 et 45 à 238.",
          largeurs_cm=[2.1, 4.2, 5.6, 4.4, 2.1], taille=8)
r.doc.add_page_break()
r.h2("4.2. PLAN D'ACTIONS PRIORISÉ")
r.para("**Huit actions à engager dans les trois mois conditionnent toutes les autres.** Le plan est organisé en cinq axes. Chaque action renvoie aux constats qu'elle traite.")
ORD = {"P1": 0, "P2": 1, "P3": 2, "P4": 3}
r.tableau(["N°", "Axe", "Action", "Constats", "Priorité", "Responsable", "Échéance"],
          [[str(k), axe, act, codes, f"{CA.BOULE_P[pr]} {pr}", qui, CA.ECHEANCE[pr]] for k, (axe, act, codes, pr, qui) in enumerate(sorted(CA.PLAN, key=lambda a: ORD[a[3]]), 1)],
          titre="Plan d'actions priorisé du volet actuariel", source=RAP + ", diapositives 25 à 27.", largeurs_cm=[0.8, 2.8, 5.6, 2.0, 1.6, 3.4, 2.2], taille=8)

# ================================================================ annexes
r.annexe(1, "Lexique")
r.tableau(["Terme", "Signification simple"],
          [["Provision", "Argent mis de côté pour une dépense déjà due mais pas encore payée (par exemple des soins faits mais pas encore facturés)."],
           ["Engagement", "Ce que le FPPN devra payer à ses adhérents : épargne et intérêts, allocations décès, soins."],
           ["Adossement actif-passif", "Le fait d'avoir l'argent disponible au moment où il faut payer. L'épargne, retirable à tout moment, doit être couverte par de l'argent vite disponible."],
           ["Sinistralité", "Part des cotisations consommée par les prestations (soins, allocations)."],
           ["Ratio de liquidité", "Argent disponible tout de suite divisé par les dettes à court terme. Il doit être au moins égal à 1."],
           ["Réserve obligatoire", "Réserve exprimée en mois de prestations. La réglementation exige 6 à 9 mois."],
           ["Fonds propres", "Ce qui reste au FPPN une fois toutes ses dettes payées. Ils servent de coussin en cas de choc."],
           ["Agrément", "Autorisation d'exercer délivrée par l'autorité de contrôle. Elle peut être retirée après deux exercices non conformes."],
           ["Méthode des unités de crédit projetées", "Méthode actuarielle de référence pour la retraite du personnel : elle compte la part déjà acquise de l'indemnité future, en tenant compte des départs, des hausses de salaire et du temps."]],
          titre="Lexique des termes utilisés", source="Adenka.", largeurs_cm=[4.4, 14.0], taille=9)
r.annexe(2, "Règles de classement")
r.tableau(["Niveau", "Signification"],
          [["🔴 Critique", "Non-conformité réglementaire, impact majeur, risque élevé"], ["🟠 Élevé", "Impact significatif sur les comptes ou les adhérents"],
           ["🟡 Moyen", "Amélioration souhaitable, impact ponctuel"], ["🟢 Faible", "Point maîtrisé ou de faible montant"],
           ["🔴 P1", "Action à engager dans les 3 mois"], ["🟠 P2", "Action à engager dans les 6 mois"], ["🟡 P3", "Action à engager de 6 à 18 mois"], ["🟢 P4", "Action à engager au-delà de 18 mois"]],
          titre="Échelles de criticité des constats et de priorité des actions", source=RAP + ", diapositives 10 et 25.", largeurs_cm=[4.4, 14.0], taille=9)
r.enregistrer(SORTIE)

# ---- retouches : séparateurs, boules colorées et surlignage jaune (repris du résumé exécutif existant)
_post = SRC[SRC.index("# ---- retouches : pages de séparation"):SRC.index("doc.save(SORTIE)")]
exec(_post, globals())
# surlignage jaune des passages à confirmer situés au milieu d'un paragraphe
import copy as _cp  # noqa: E402
for run in list(doc.element.body.iter(qn("w:r"))):
    ts = run.findall(qn("w:t"))
    if len(ts) != 1 or not ts[0].text:
        continue
    t0 = ts[0].text
    j = next((x for x in JAUNE if x in t0 and x != t0), None)
    if not j:
        continue
    for k_, morceau in enumerate(t0.split(j, 1)[0:1] + [j] + t0.split(j, 1)[1:]):
        if not morceau:
            continue
        nr = _cp.deepcopy(run)
        tt = nr.find(qn("w:t"))
        tt.text = morceau
        tt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        if morceau == j:
            rpr = nr.find(qn("w:rPr"))
            if rpr is None:
                rpr = nr.makeelement(qn("w:rPr"), {})
                nr.insert(0, rpr)
            h_ = rpr.makeelement(qn("w:highlight"), {qn("w:val"): "yellow"})
            apres_ = [e for e in rpr if e.tag in {qn("w:" + n_) for n_ in ("u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath")}]
            if apres_:
                apres_[0].addprevious(h_)
            else:
                rpr.append(h_)
        run.addprevious(nr)
    run.getparent().remove(run)
doc.save(SORTIE)
print(SORTIE, "constats :", len(CA.CONSTATS), "actions :", len(CA.PLAN), "à confirmer :", len(JAUNE))
