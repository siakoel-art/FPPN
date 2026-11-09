# -*- coding: utf-8 -*-
# v12, étape 6 (09/11/2026) : structure demandée par Paul Alexandre DAGRY et retouches de l'utilisateur sur le rapport du 23/10/2026.
#  - « Comprendre le FPPN » devient « Présentation du FPPN », « Contexte et démarche » devient « Contexte & Objectifs » ;
#  - nouvelle partie 3 « Méthodologie & Limites » : tableau des méthodes retenues (remplace les diapositives « Méthode : … »),
#    limites au modèle fourni, puis « Outils & Traçabilité » (ancienne partie 13) ;
#  - nouvelle partie 4 « Résultats de l'analyse actuarielle » : anciennes parties 3 à 12, numérotées 4.1 à 4.10 ;
#  - « Analyses en attente » devient la partie 5 ; renvois « partie N » renumérotés ;
#  - retouches de l'utilisateur : couverture, séparateurs, diapositives 13, 19, 20 et 23, listes, annexes, notes retirées.
# Exécuté par patch_v12.py.

# ---------------------------------------------------------------- intitulés
rep('separateur(1, "Comprendre le FPPN", "Cadre, garanties, circuits et notions de base",',
    'separateur(1, "Présentation du FPPN", "Cadre, garanties, circuits et notions de base",')
rep('separateur(2, "Contexte et démarche", "Mission, périmètre, chaîne de traitement")',
    'separateur(2, "Contexte & Objectifs", "Mission, périmètre, documents reçus, traitement des données")')
rep('separateur(13, "Outils et traçabilité", "Chaîne de traitement, extraits de code, contrôle des chiffres")',
    'separateur(13, "Outils & Traçabilité", "Comment chaque chiffre est produit, contrôlé et reproductible")')
rep('contenu({"titre": "Indicateurs clés de 2020 au 31/08/2026",', 'contenu({"titre": "Indicateurs clés de 2020 à 2026",')
s = s.replace('ref("Indicateurs clés de 2020 au 31/08/2026")', 'ref("Indicateurs clés de 2020 à 2026")')

# ---------------------------------------------------------------- diapositive 19 : deux lectures sur deux lignes
rep('''encadre=["**Lecture du FPPN** : comptes tels qu'ils sont tenus. **Lecture corrigée selon le PCMS** :''',
    '''encadre=["**Lecture du FPPN** : comptes tels qu'ils sont tenus.", "**Lecture corrigée selon le PCMS** :''')

# ---------------------------------------------------------------- diapositive 20 : « NA » au lieu de « — »
_a = s.index("_MAT_CF = [")
_b = s.index('contenu({"titre": "Matrice de conformité"', _a)
s = s[:_a] + s[_a:_b].replace('"—"', '"NA"') + s[_b:]
rep('textes applicables en colonnes ; « — » : le texte ne fixe pas de règle sur ce domaine"},',
    'textes applicables en colonnes. « NA » : le texte ne fixe pas de règle sur ce domaine"},')

# ---------------------------------------------------------------- couverture
rep('CJ[k] = {"Google Shape;138;p25": "Audit Comptable et financier FPPN", "Text Placeholder 4": "23 octobre 2026"}',
    'CJ[k] = {"Google Shape;138;p25": "Audit comptable et financier de la gestion du FPPN", "Text Placeholder 4": "09 novembre 2026"}')
rep("sh.left, sh.top, sh.width, sh.height = Emu(908861), Emu(2534894), Emu(5918659), Emu(1231106)",
    "sh.left, sh.top, sh.width, sh.height = Emu(390701), Emu(2372334), Emu(7087059), Emu(1410643)")
rep("st_ = sl.shapes.add_textbox(Emu(817421), Emu(5135034), Emu(6497320), Emu(461665))",
    "st_ = sl.shapes.add_textbox(Emu(299261), Emu(4972474), Emu(6497320), Emu(461665))")
rep('"assets", "logo_fppn.jpg"), Emu(938577), Emu(4017907), Emu(1844040), Emu(865220))',
    '"assets", "logo_fppn.jpg"), Emu(420417), Emu(3855347), Emu(1844040), Emu(865220))')
rep('''            if sh.name == "Text Placeholder 4":
                sh.left, sh.width = Emu(10506325), Emu(1187835)''',
    '''            if sh.name == "Text Placeholder 4":
                sh.left, sh.width = Emu(10506325), Emu(1187835)
            if sh.name == "ZoneTexte 13":
                sh.left, sh.top = Emu(310299), Emu(6285082)''')

# ---------------------------------------------------------------- séparateurs : cadres élargis (retouche de l'utilisateur)
rep('''            if sh.name in ("Rectangle 9", "Rectangle 5"):
                sh.width = Cm(10.6)''',
    '''            if sh.name in ("Rectangle 9", "Rectangle 5"):
                sh.width = Emu(4810927 if sh.name == "Rectangle 9" else 4377564)
            if sh.name == "Image 10":
                sh.left = Emu(8749481)
            if x["num"] == "R" and sh.name in globals().get("RETOUCHES_U", {}).get("sep_resume", {}):
                sh.left, sh.top, sh.width, sh.height = (Emu(v_) for v_ in RETOUCHES_U["sep_resume"][sh.name])''')

# ---------------------------------------------------------------- notes retirées (version de l'utilisateur sans notes)
rep('''    if x.get("notes"):
        sl.notes_slide.notes_text_frame.text = x["notes"]''',
    '''    if x.get("notes") and not globals().get("SANS_NOTES"):
        sl.notes_slide.notes_text_frame.text = x["notes"]''')

# ---------------------------------------------------------------- liste des tableaux sur deux diapositives
rep('SL.append({"type": "liste_tableaux", "notes": ""})',
    'SL.append({"type": "liste_tableaux", "notes": "", "part": (0, 2)})\nSL.append({"type": "liste_tableaux", "notes": "", "part": (1, 2)})')
rep('''        nom = "Figure" if x["type"] == "liste_figures" else "Tableau"''',
    '''        nom = "Figure" if x["type"] == "liste_figures" else "Tableau"
        if x.get("part"):
            _k, _n = x["part"]
            _per = (len(liste) + _n - 1) // _n
            liste = liste[_k * _per:(_k + 1) * _per]''')

# ---------------------------------------------------------------- titres en double admis pour les diapositives de la partie 3
rep('        assert t_ not in _POS, ("titre en double", t_)\n        _POS[t_] = i',
    '        if t_ in _POS and t_ == "Méthodologie & Limites":\n            continue\n        assert t_ not in _POS, ("titre en double", t_)\n        _POS[t_] = i')

# ---------------------------------------------------------------- nouvelle partie 3 et réorganisation (avant la résolution des renvois)
_NOUV = r'''
# ---- v13 : structure demandée par Paul Alexandre DAGRY (09/11/2026)
SANS_NOTES = True  # v13 : rapport sans notes du présentateur (version de l'utilisateur)
RETOUCHES_U = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "retouches_utilisateur_20261109.json"), encoding="utf-8"))
_n0 = len(SL)
_METH = [  # analyse, méthode retenue, pourquoi ce choix, données utilisées, diapositive de détail
    ("Contrôle des données", "Empreinte numérique de chaque fichier, comptage des lignes et cellules avant et après conversion, totaux imprimés retrouvés, rapprochement des fichiers aller et retour poste par poste et mois par mois",
     "Prouver que la base d'analyse reproduit fidèlement les fichiers reçus. Limite : la fidélité aux fichiers n'est pas l'exactitude des fichiers du FPPN", "4 287 fichiers, balances 2020 au 31/08/2026, 37 balances âgées du Grenier", "Couverture de la période"),
    ("Effectifs et taux de cotisation", "Effectif du mois = matricules distincts du fichier maladie 936. Entrées et sorties sur mois consécutifs. Taux = précompte du mois / salaire de base, agent par agent",
     "Seule source mensuelle nominative disponible. Limite : ni motif de sortie, ni âge, ni sexe", "Fichiers mensuels de 31 mois (janvier 2024 à août 2026)", "Effectifs cotisants"),
    ("Sinistralité", "S/C = prestations comptabilisées / cotisations comptabilisées, par garantie et par exercice, comparé à la norme de 70 % à 80 % (Décision n° 05/2022)",
     "Sans base des prestations datée, seule la lecture comptable est possible. Elle mélange les soins de l'année et des années antérieures", "Comptes 601, 6013, 6018, 60800007, 6021 et 6024, cotisations 7010 et 7021", "Sinistralité comptable"),
    ("Épargne et intérêts dus", "Couverture = actifs réservés (banques, dépôts à terme, fonds réservés) / épargne due (41195100, 41195200). Intérêts sur l'encours moyen aux taux du manuel",
     "Ni échéancier individuel ni relevé par épargnant : seule la mesure globale sur les soldes comptables est vérifiable", "Balances 2020 au 31/08/2026, manuel des produits d'épargne", "Adossement de l'épargne"),
    ("Prêts du Grenier", "Part de l'encours en retard de plus de 30 jours, en montant et en nombre (intervalle de Wilson à 95 %). Dépréciation par tranche de retard (40 %, 80 %, 100 %)",
     "Le retard est celui du logiciel de prêts. Les taux de dépréciation sont ceux de la politique de crédit du Grenier", "37 balances âgées, 2 959 prêts au 31/08/2026", "Encours du Grenier"),
    ("Bilan et fonds propres", "Bilan reconstitué compte par compte à partir des balances, confronté aux états financiers. Fonds propres = comptes 10, 11 et 12 et résultat",
     "Les balances donnent le détail que les états financiers ne donnent pas. La balance au 31/08/2026 est provisoire", "Balances et états financiers de 2020 au 31/08/2026", "Provisions techniques"),
    ("Provisions maladie et décès", "Maladie : soins d'une année facturés l'année suivante (déclarations tardives, Code CIMA, art. 334-12). Décès : états de fin d'année rapprochés du compte 41591000, décès déclarés après la clôture",
     "Les cliniques sont payées en quelques mois : Chain Ladder ne s'applique pas. Le FPPN calcule déjà sa dette décès dossier par dossier", "Journal des factures 2023 à 2025, états décès 2021 à 2025", "Provision maladie : montant nécessaire"),
    ("Provision du personnel", "Unités de crédit projetées, agent par agent, départ à la retraite à 60 ans (ordonnance n° 2012-03 du 11/01/2012)",
     "Méthode de référence : elle tient compte des départs avant la retraite, des hausses de salaire et du temps", "DISA 2021 à 2025, 79 agents en 2025", "Provision du personnel : méthode et paramètres"),
    ("Adossement actif-passif", "Actifs classés selon leur délai de mobilisation, rythme de sortie de l'épargne observé de 2021 à 2025, écarts de liquidité année par année, trois scénarios",
     "Sans échéancier de l'épargne ni données individuelles, la duration et les modèles stochastiques ne sont pas applicables", "Bilans au 31/12/2025 et au 31/08/2026, sorties d'épargne 2021 à 2025", "Choix des méthodes d'adossement"),
    ("Placements", "Encours de chaque placement lu jour par jour dans les grands livres, rendement comparé aux taux de la BCEAO et des titres de l'État, quatre conditions de l'art. 17",
     "L'inventaire signé des placements n'a pas été transmis : les grands livres sont la seule source complète", "Grands livres de 2021 au 31/08/2026", "Rendement des placements"),
    ("Conformité prudentielle", "36 indicateurs de la Décision n° 05/2022 à chaque date, en deux lectures (comptes tels que tenus, puis classés selon le PCMS), sensibilités et leviers chiffrés",
     "Le texte fixe les formules et les numéros de comptes : le calcul suit le texte à la lettre", "Balances 2020 au 31/08/2026, plan comptable des mutuelles sociales", "Textes applicables"),
    ("Tables de mortalité", "Comparaison des tables CIMA H, CIMA F, TD et TV au taux technique de 3,5 %, indicateurs de référence",
     "Sans dates de naissance, la mortalité observée du FPPN ne peut pas être mesurée : les tables réglementaires servent de référence", "Tables CIMA H, CIMA F, TD et TV", "Choix des tables"),
    ("Projection de l'équilibre technique", "Projection agrégée des cotisations, prestations et frais de 2026 à 2030, en deux scénarios fondés sur les tendances observées",
     "Une projection par garantie et par âge exige les données individuelles, non reçues", "Comptes 2023 à 2025, tendances observées", "Projection agrégée : deux scénarios"),
]
_LIG_M = [[a, b, c, d, ref(e)] for a, b, c, d, e in _METH]
for _k, (_lot, _ss) in enumerate(((_LIG_M[:7], "Méthodes retenues (1/2)"), (_LIG_M[7:], "Méthodes retenues (2/2)"))):
    contenu({"titre": "Méthodologie & Limites", "sous_titre": _ss,
             "synthese": "Pour chaque analyse : la méthode retenue, pourquoi elle a été choisie, les données utilisées et la diapositive où voir le détail"},
            "Les méthodes sont choisies selon les données disponibles : chacune est la plus précise que permettent les documents transmis.",
            tableau={"titre": None, "entetes": ["Analyse", "Méthode retenue", "Pourquoi ce choix", "Données utilisées", "Détail"], "largeurs": [4.2, 10.2, 8.4, 6.0, 2.4],
                     "taille": 8, "sans_code": True, "lignes": _lot},
            disposition="dessous",
            analyse=["**NB :** la liste des méthodes non retenues et la justification de chaque choix figurent en annexe (" + ref("Méthodes non retenues : vue d'ensemble") + " et suivantes)."],
            source="Source : Adenka, d'après les diapositives de détail citées")
_LIM = [  # diagnostic, limite, conséquence, documents à obtenir, priorité
    ("Données", "[[Base des prestations non datée]] : ni date des soins, ni date de dépôt, ni date de paiement des factures", "Provisions non calculables dossier par dossier, boni et mali non mesurables",
     "[[Base des prestations avec les dates des soins, du dépôt et du paiement]]", "P1"),
    ("Données", "Dates de naissance absentes des fichiers de cotisations, connues pour 105 retraités seulement", "Mortalité par âge, santé des retraités et projection par âge non calculables",
     "Fichier des mutualistes avec date de naissance et sexe ; fichier des ayants droit", "P1"),
    ("Données", "Fichiers maladie et décès de décembre 2024 vides ; fichier retour 9X4 de décembre 2024 manquant", "Effectifs et précomptes de fin 2024 non vérifiables",
     "Fichiers de décembre 2024", "P2"),
    ("Épargne", "[[Aucun taux servi arrêté par une décision signée]] ; aucun relevé individuel d'épargne", "Épargne due mesurée globalement ; intérêts dus estimés à deux taux",
     "Délibération fixant le taux servi ; relevés individuels des épargnants", "P1"),
    ("Placements et actifs", "Ni inventaire ni échéancier signés des placements", "Adossement par échéance établi sur les grands livres : borne favorable",
     "Inventaire et échéancier signés des placements", "P1"),
    ("Placements et actifs", "Terrains et clinique sans évaluation récente ; créance sur l'Hôpital de la Police sans accord de remboursement", "Valeur réalisable des actifs longs non démontrée",
     "Rapports d'évaluation ; accord écrit de l'État", "P2"),
    ("Comptes et gouvernance", "Balance au 31/08/2026 provisoire ; deux versions des comptes 2023", "Situation intermédiaire et excédent 2023 à confirmer",
     "Balance définitive ; comptes 2023 approuvés par l'Assemblée générale", "P2"),
    ("Comptes et gouvernance", "Méthode de la provision du personnel non transmise ; âge de départ appliqué non confirmé", "Recalcul à 60 ans sans comparaison avec la méthode du FPPN",
     "Note de calcul de la provision ; âge de départ appliqué", "P2"),
    ("Comptes et gouvernance", "Délibérations de placement et budget du projet de Cocody non transmis", "Gouvernance des placements non démontrée",
     "Procès-verbaux du conseil d'administration ; budget et plan de financement du projet", "P2"),
]
contenu({"titre": "Méthodologie & Limites", "sous_titre": "Limites",
         "synthese": "Plusieurs documents manquants ou incohérents empêchent de conclure sur certains points. Le tableau indique, pour chaque limite, sa conséquence et le document à obtenir"},
        "P1 : indispensable pour finaliser l'analyse. P2 : utile pour l'affiner.",
        tableau={"titre": None, "entetes": ["Diagnostic", "Limite de l'analyse", "Conséquence", "Documents à obtenir", "Priorité"], "largeurs": [4.2, 9.8, 7.4, 7.8, 2.0],
                 "taille": 8.5, "sans_code": True, "lignes": [list(l_) for l_ in _LIM]},
        source="Source : liste des documents demandés au FPPN ; " + "diapositives de détail du rapport", limites=True)
_NOUV_SL = SL[_n0:]
del SL[_n0:]
for _x in _NOUV_SL:
    _x["_partie"] = "3"
# blocs par partie
_pre, _blocs, _cur = [], {}, None
for _x in SL:
    if _x["type"] == "separateur":
        _cur = str(_x["num"])
        _blocs[_cur] = [_x]
    elif _cur is None:
        _pre.append(_x)
    else:
        _blocs[_cur].append(_x)
_RENUM = {"3": "4.1", "4": "4.2", "5": "4.3", "6": "4.4", "7": "4.5", "8": "4.6", "9": "4.7", "9.2": "4.7.2", "9.3": "4.7.3",
          "10": "4.8", "11": "4.9", "12": "4.10", "13": "3.1", "14": "5"}
for _k in ("3", "4", "5", "6", "7", "8", "9"):
    _blocs[_k] = [_x for _x in _blocs[_k] if not (_x["type"] == "gab" and str(_x.get("entete", {}).get("titre", "")).startswith("Méthode : "))]
_SEP_M = {"type": "separateur", "num": "3", "titre": "Méthodologie & Limites", "sous_titre": "Méthodes retenues, limites de l'analyse, outils et traçabilité", "notes": ""}
_SEP_R = {"type": "separateur", "num": "4", "titre": "Résultats de l'analyse actuarielle", "sous_titre": "Des données à la projection de l'équilibre technique", "notes": ""}
_ordre = _pre + _blocs["R"] + _blocs["1"] + _blocs["2"] + [_SEP_M] + _NOUV_SL + _blocs["13"] + [_SEP_R]
for _k in ("3", "4", "5", "6", "7", "8", "9", "9.2", "9.3", "10", "11", "12"):
    _ordre += _blocs[_k]
_ordre += _blocs["14"] + _blocs["A"]
_old = {id(_x): _i for _i, _x in enumerate(SL, 1)}
SL[:] = _ordre
_new = {id(_x): _i for _i, _x in enumerate(SL, 1)}
_o2n = {_old[_k]: _new[_k] for _k in _new if _k in _old}
_n_old = len(_old)
TRACE[:] = [_t for _t in TRACE if _t["diapositive"] > _n_old or _t["diapositive"] in _o2n]  # chiffres des diapositives « Méthode : » retirées
for _t in TRACE:
    if _t["diapositive"] <= _n_old:
        _t["diapositive"] = _o2n[_t["diapositive"]]
# figures et tableaux renumérotés dans l'ordre des diapositives
_MAPN = {}
for _lst, _mot in ((FIG, "Figure"), (TAB, "Tableau")):
    for _e in _lst:
        if _e[2] <= _n_old:
            _e[2] = _o2n.get(_e[2], _e[2])
    _ord = sorted(_lst, key=lambda e_: (e_[2], e_[0]))
    _MAPN[_mot] = {str(e_[0]): str(k_) for k_, e_ in enumerate(_ord, 1)}
    for k_, e_ in enumerate(_ord, 1):
        e_[0] = k_
    _lst.sort(key=lambda e_: e_[0])
for _x in SL:
    if _x["type"] == "separateur" and str(_x["num"]) in _RENUM and _x is not _SEP_M and _x is not _SEP_R:
        _x["num"] = _RENUM[str(_x["num"])]
    if "_partie" in _x and str(_x["_partie"]) in _RENUM and _x not in _NOUV_SL:
        _x["_partie"] = _RENUM[str(_x["_partie"])]


def _renvois_partie(v):
    """« partie 9 », « parties 5 à 10 » : numéros des anciennes parties remplacés par les nouveaux."""
    if isinstance(v, str):
        def _sub(m):
            nums = re.sub(r"\d+(?:\.\d+)?", lambda n: _RENUM.get(n.group(0), n.group(0)), m.group(2))
            return m.group(1) + " " + nums
        v = re.sub(r"\b(Figure|Tableau) (\d+) :", lambda m: m.group(1) + " " + _MAPN[m.group(1)].get(m.group(2), m.group(2)) + " :", v)
        return re.sub(r"\b([Pp]arties?)\s+(\d+(?:\.\d+)?(?:\s*(?:,|à|et)\s*\d+(?:\.\d+)?)*)", _sub, v)
    if isinstance(v, list):
        return [_renvois_partie(e) for e in v]
    if isinstance(v, tuple):
        return tuple(_renvois_partie(e) for e in v)
    if isinstance(v, dict):
        return {k: (_renvois_partie(e) if k not in ("num", "_partie", "type") else e) for k, e in v.items()}
    return v


for _i, _x in enumerate(SL):
    SL[_i] = _renvois_partie(_x)
'''
rep("_POS = {}\nfor i, x in enumerate(SL, 1):", _NOUV + "_POS = {}\nfor i, x in enumerate(SL, 1):")

# ---------------------------------------------------------------- retouches finales (avant l'enregistrement)
_FIN = r'''
# ---- v13 : retouches de l'utilisateur (09/11/2026) et mise en forme du tableau des limites
RETOUCHES_U = globals().get("RETOUCHES_U") or {"corrections": [], "formes": {}, "sep_resume": {}}
_corr13 = RETOUCHES_U["corrections"]
for sl in prs.slides:
    t_ = [sh.text_frame.text for sh in sl.shapes if sh.has_text_frame and sh.name == "Adenka_Titre"]
    for c_ in _corr13:
        if t_ and t_[0] == c_["titre"]:
            for f_ in c_["formes"]:
                for sh in sl.shapes:
                    if sh.name == f_["nom"]:
                        sh.left, sh.top, sh.width, sh.height = (Emu(v_) for v_ in f_["apres"])
                        break
    # encadrés retouchés : remplacés par la version de l'utilisateur
    if t_ and t_[0] in RETOUCHES_U["formes"]:
        from lxml import etree as _et
        for sh in list(sl.shapes):
            if sh.has_text_frame and sh.text_frame.text.startswith(("Deux lectures des comptes", "Échelles")):
                sh._element.getparent().remove(sh._element)
        arbre = sl.shapes._spTree
        for xml_ in RETOUCHES_U["formes"][t_[0]]:
            arbre.append(_et.fromstring(xml_))
    # tableau des limites : diagnostics fusionnés sur fond noir, priorités colorées
    for sh in sl.shapes:
        if sh.has_table and sh.table.cell(0, 0).text == "Diagnostic" and sh.table.cell(0, 4).text == "Priorité":
            tb_ = sh.table
            n_ = len(tb_.rows)
            i_ = 1
            while i_ < n_:
                j_ = i_
                while j_ + 1 < n_ and tb_.cell(j_ + 1, 0).text == tb_.cell(i_, 0).text:
                    j_ += 1
                lib_ = tb_.cell(i_, 0).text
                for k_ in range(i_ + 1, j_ + 1):
                    tb_.cell(k_, 0).text = ""
                if j_ > i_:
                    tb_.cell(i_, 0).merge(tb_.cell(j_, 0))
                c_ = tb_.cell(i_, 0)
                c_.text = ""
                c_.vertical_anchor = MSO_ANCHOR.MIDDLE
                runs_texte(c_.text_frame.paragraphs[0], lib_, 9, True, "FFFFFF")
                c_.fill.solid()
                c_.fill.fore_color.rgb = RGBColor.from_string(NOIR)
                i_ = j_ + 1
            for k_ in range(1, n_):
                c_ = tb_.cell(k_, 4)
                pr_ = c_.text.strip()
                c_.text = ""
                c_.vertical_anchor = MSO_ANCHOR.MIDDLE
                p_ = c_.text_frame.paragraphs[0]
                p_.alignment = PP_ALIGN.CENTER
                runs_texte(p_, pr_, 9, True, "FFFFFF")
                c_.fill.solid()
                c_.fill.fore_color.rgb = RGBColor.from_string("F4A261" if pr_ == "P1" else "595959")
                for p2_ in tb_.cell(k_, 2).text_frame.paragraphs:
                    for r2_ in p2_.runs:
                        r2_.font.italic = True
                for p2_ in tb_.cell(k_, 3).text_frame.paragraphs:
                    for r2_ in p2_.runs:
                        r2_.font.bold = True
            bas_ = (sh.top + sh.height) / 360000 + 0.15
            tb2 = sl.shapes.add_textbox(Cm(1.4), Cm(bas_), Cm(31.2), Cm(0.5))
            p_ = tb2.text_frame.paragraphs[0]
            for coul_, lib_ in (("F4A261", "P1 : indispensable pour finaliser l'analyse"), ("595959", "P2 : utile pour affiner l'analyse")):
                runs_texte(p_, "■ ", 9, True, coul_)
                runs_texte(p_, lib_ + "     ", 8.5, False, NOIR)
# annexes « Signes et abréviations » et « Références et sources » : sans pied de page ni ligne de sigles
for sl in prs.slides:
    t26 = [sh.text_frame.text for sh in sl.shapes if sh.has_text_frame and sh.name == "TextBox 26"]
    if t26 and t26[0] in ("Signes et abréviations", "Références et sources"):
        for sh in list(sl.shapes):
            if sh.has_text_frame and (sh.text_frame.text.startswith("Empreinte Africaine") or sh.text_frame.text.startswith("Sigles :")):
                sh._element.getparent().remove(sh._element)
            if t26[0] == "Signes et abréviations" and sh.has_table:
                sh.left, sh.top = Emu(480000), Emu(540000)
'''
rep("prs.save(SORTIE)\nfinaliser(SORTIE)\nos.remove(tmp)", _FIN + "prs.save(SORTIE)\nfinaliser(SORTIE)\nos.remove(tmp)")

# numéros de sous-partie longs (4.7.2, 4.10) : police réduite pour tenir dans le cadre du séparateur
rep("""                        r.font.size = Pt(r.font.size.pt * 0.7 if r.font.size else 40)""",
    """                        r.font.size = Pt((r.font.size.pt if r.font.size else 57) * (0.7 if len(str(x["num"])) <= 3 else 0.5))""")
