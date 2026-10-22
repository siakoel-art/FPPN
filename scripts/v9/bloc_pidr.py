# ---- v9 : provision pour indemnités de départ à la retraite recalculée agent par agent (partie 9)
PP, PSY, PIN, PSE = V9["pidr_parametres"], V9["pidr_synthese"], V9["pidr_individuel"], V9["pidr_sensibilites"]
SRC_DISA = "Source : DISA_256927_2021.xls, DISA_256927_2022.xls, DISA_256927_2023.xls, DISA_256927_2025.xls (déclarations CNPS) ; ETAT 301 & DISA 2024 FPPN.xls (RECAPSALAIRES 2024 CMP) ; ETAT 301 & DISA 2025 FPPN.xls (RECAPSALAIRES 2025 GRENIER)"
ENT9 = ("Direction Générale", "CMP", "Grenier", "Non documentée")
contenu({"titre": "Provision du personnel : méthode et paramètres", "sous_titre": "Chaque agent civil est évalué à partir de sa déclaration CNPS, avec des paramètres observés ou fixés par les textes",
         "synthese": "Méthode des unités de crédit projetées : indemnité projetée au départ, acquise au prorata de l'ancienneté, pondérée par la survie, la rotation et l'actualisation"},
        "L'indemnité de départ vaut un pourcentage du salaire mensuel moyen par année d'ancienneté (décret n° 2017-210, art. 4 et 5). La méthode des unités de crédit projetées, retenue par IAS 19 et utilisée comme référence de bonne pratique, ne provisionne que la part acquise à la date d'évaluation. Le calcul est fait par formules dans les onglets 14a à 14c du classeur.",
        tableau={"titre": tab("Paramètres de la provision du personnel au 31/12/2025"), "entetes": ["Paramètre", "Valeur", "Justification"], "largeurs": [6.4, 4.0, 20.8], "taille": 8.5,
                 "lignes": [["Âge de départ", N("PIDR âge", PP["age_depart"]) + " ans", "Convention collective interprofessionnelle du 19/07/1977, art. 40 : 55 ans, plus un an par enfant à charge (au plus 3) ; les enfants ne figurent pas dans les fichiers reçus."],
                            ["Barème", "30 % ; 35 % ; 40 %", "Décret n° 2017-210, art. 4 : jusqu'à 5 ans, de 6 à 10 ans, au-delà de 10 ans, fractions d'année au mois inférieur."],
                            ["Salaire", "Brut annuel / mois", "Salaire brut de la DISA divisé par la durée d'activité de l'année (salaire global mensuel moyen, art. 4)."],
                            ["Taux d'actualisation", pc("PIDR taux", PP["taux_actualisation"]), f"Moyenne de 2025 du rendement de l'obligation de l'État de Côte d'Ivoire à 10 ans, {N('PIDR cotations', PP['cotations_2025'])} cotations Investing.com (IAS 19, § 83)."],
                            ["Croissance des salaires", pc("PIDR croissance", PP["croissance_salaires"]), f"Médiane de la croissance annuelle observée sur les {N('PIDR agents croiss', PP['agents_croissance'])} agents présents douze mois en 2021 et en 2025."],
                            ["Rotation du personnel", pc("PIDR rotation", PP["rotation"]), "Moyenne des taux annuels de sortie observés entre les DISA 2021, 2022, 2023 et 2025."],
                            ["Survie", "Table CIMA F", "Table de survie, prudente pour une indemnité versée si l'agent est en vie au départ (principe de prudence, partie 11)."],
                            ["Entité", "DG, CMP, Grenier", "Listes nominatives des récapitulatifs de salaires 2024 (CMP) et 2025 (Grenier) ; agents absents de ces listes : rattachement non documenté."]]},
        disposition="dessous",
        analyse=["**Tous les paramètres viennent d'un texte ou d'une observation sur les fichiers du FPPN.** Seuls l'âge de départ, faute d'enfants déclarés, et le taux d'actualisation, faute de marché profond, relèvent d'un choix justifié."],
        implication="Les effets de l'âge de départ, du taux, de la rotation et des salaires sont mesurés plus loin, une hypothèse à la fois.",
        valable="personnel civil déclaré à la CNPS (DISA), hors travailleurs occasionnels et fonctionnaires de police détachés, évaluation au 31/12",
        source=SRC_DISA + " ; décret n° 2017-210 du 30/03/2017 ; convention collective du 19/07/1977, art. 40 ; IAS 19, § 67 et 83 ; Investing.com ; table CIMA F")
contenu({"titre": "Provision du personnel par entité de 2021 à 2025", "sous_titre": "La provision recalculée est inférieure au compte 196 ; les droits acquis en sont proches",
         "synthese": f"Au 31/12/2025 : provision recalculée {F('PIDR 2025 total', PSY[-1]['dbo'])}, droits acquis {F('PIDR 2025 acquis', PSY[-1]['acquis'])}, compte 196 {F('PIDR 2025 196', PSY[-1]['provision_196'])}"},
        "La provision recalculée suit la méthode des unités de crédit projetées. Les droits acquis sont l'indemnité que chaque agent recevrait s'il partait à la date d'évaluation, sans projection ni actualisation. Le rattachement à une entité suit les listes de 2024 et 2025, appliquées aussi aux années antérieures.",
        tableau={"titre": tab("Provision du personnel par entité, droits acquis et compte 196"), "entetes": ["Exercice", "Agents", "Direction Générale", "CMP", "Grenier", "Non documentée", "Provision recalculée", "Droits acquis", "Compte 196 (FPPN)"],
                 "largeurs": [2.4, 2.2, 3.8, 3.4, 3.2, 3.6, 4.2, 4.2, 4.2], "taille": 8.5,
                 "lignes": [[str(x["annee"]), N(f"PIDR agents {x['annee']}", x["agents"])] + [N(f"PIDR {e} {x['annee']}", x[e]["dbo"]) for e in ENT9]
                            + [N(f"PIDR tot {x['annee']}", x["dbo"]), N(f"PIDR acq {x['annee']}", x["acquis"]), N(f"PIDR 196 {x['annee']}", x["provision_196"])] for x in PSY]},
        disposition="dessous",
        analyse=[f"**Au 31/12/2025, la provision recalculée vaut {pc('PIDR ratio 196', PSY[-1]['dbo'] / PSY[-1]['provision_196'])} du compte 196 ; les droits acquis en valent {pc('PIDR ratio acquis', PSY[-1]['acquis'] / PSY[-1]['provision_196'])}.** Le montant du FPPN est donc proche d'un calcul sans projection ni actualisation, comme l'a décrit le chef comptable.",
                 f"**La Direction Générale porte {pc('PIDR part DG', PSY[-1]['Direction Générale']['dbo'] / PSY[-1]['dbo'])} de la provision 2025 avec {N('PIDR agents DG', PSY[-1]['Direction Générale']['agents'])} agents.** Le CMP compte {N('PIDR agents CMP', PSY[-1]['CMP']['agents'])} agents, dont {N('PIDR CMP depuis 2022', sum(1 for x in PIN if x['annee'] == 2025 and x['entite'] == 'CMP' and int(x['embauche'][-4:]) >= 2022))} recrutés depuis 2022 : leurs droits acquis restent faibles."],
        implication="Le compte 196 couvre la provision recalculée selon les deux approches. La méthode du FPPN (droits acquis) reste à confirmer par sa note de calcul (PBC, item 113).",
        valable="exercices 2021, 2022, 2023 et 2025 (pas de DISA exploitable en 2024), personnel civil déclaré à la CNPS ; montants en FCFA",
        source=SRC_DISA + " ; " + SRC_COMPTA[9:] + " ; onglets 14b et 14c du classeur v9")


def _lignes_pidr(liste):
    return [[x["id"], str(x["naissance"]), x["embauche"], N(f"PIDR âge {x['id']}", x["age"]), N(f"PIDR anc {x['id']}", x["anciennete"], 2),
             N(f"PIDR sal {x['id']}", x["salaire_mensuel"]), N(f"PIDR acq {x['id']}", x["indemnite_acquise"]), N(f"PIDR dbo {x['id']}", x["dbo"])] for x in liste]


_P25 = [x for x in PIN if x["annee"] == 2025]
_ENTETE_I = ["Agent", "Naissance", "Embauche", "Âge", "Ancienneté (ans)", "Salaire mensuel moyen", "Droits acquis", "Provision recalculée"]
_LARG_I = [3.0, 3.0, 3.6, 2.4, 3.6, 5.0, 5.0, 5.6]
def _part5(e):
    l_ = sorted((x["dbo"] for x in _P25 if x["entite"] == e), reverse=True)
    return sum(l_[:5]) / sum(l_)


_nd = [x for x in _P25 if x["entite"] == "Non documentée"]
for _ent, _sous, _morceaux in (("Direction Générale", "Les cinq provisions les plus élevées font " + pc("PIDR part top5 DG", _part5("Direction Générale")) + " de la provision de l'entité", [None]),
                               ("CMP", "Un personnel récent : les droits acquis restent faibles", [0, 1]),
                               ("Grenier", "Trois agents, dont un recruté en 2025", [None]),
                               ("Non documentée", N("PIDR nd n", len(_nd)) + " agents recrutés en " + "2025" + " ne figurent dans aucun récapitulatif de salaires par entité", [0, 1])):
    _l = sorted([x for x in _P25 if x["entite"] == _ent], key=lambda x: -x["dbo"])
    _np = -(-len(_l) // 16)          # au plus 16 agents par diapositive
    _taille_p = -(-len(_l) // _np)
    for _k in range(_np):
        _part = _l[_k * _taille_p:(_k + 1) * _taille_p]
        _suffixe = "" if _np == 1 else f" ({_k + 1}/{_np})"
        _top = max(_part, key=lambda x: x["dbo"])
        contenu({"titre": f"Provision du personnel agent par agent : {_ent if _ent != 'Non documentée' else 'rattachement non documenté'}{_suffixe}", "sous_titre": _sous,
                 "synthese": f"{N('PIDR n ' + _ent + _suffixe, len(_part))} agents ; provision recalculée {F('PIDR somme ' + _ent + _suffixe, sum(x['dbo'] for x in _part))} ; droits acquis {F('PIDR acq somme ' + _ent + _suffixe, sum(x['indemnite_acquise'] for x in _part))}"},
                "Agents identifiés par un numéro (année d'évaluation et rang dans la DISA 2025) ; les noms figurent dans l'onglet 14b du classeur, avec toutes les formules. Montants en FCFA. Âge calculé sur l'année de naissance."
                + (" Le rattachement de ces agents à la Direction Générale, au CMP ou au Grenier est à confirmer par le FPPN." if _ent == "Non documentée" else ""),
                tableau={"titre": tab(f"Provision du personnel au 31/12/2025 : {_ent}{_suffixe}"), "entetes": _ENTETE_I,
                         "largeurs": _LARG_I, "taille": 8,
                         "lignes": _lignes_pidr(_part)},
                disposition="dessous",
                analyse=[f"**L'agent {_top['id']} porte la provision la plus élevée de ce tableau ({F('PIDR top ' + _ent + _suffixe, _top['dbo'])})**, avec {N('PIDR top anc ' + _ent + _suffixe, _top['anciennete'], 2)} ans d'ancienneté."],
                implication="Ces montants s'additionnent dans la diapositive « Provision du personnel par entité ». Ils changent avec l'âge de départ, la date d'embauche et l'entité que le FPPN confirmera (PBC, item 113).",
                valable="agents déclarés dans la DISA 2025, évaluation au 31/12/2025 ; paramètres de l'onglet 14a",
                source=SRC_DISA + " ; onglet 14b du classeur v9")
_s0 = PSE[0]["dbo"]
_chg = V9["embauche_modifiee"]
_ag55 = [x for x in _P25 if x["age"] >= PP["age_depart"]]
contenu({"titre": "Provision du personnel : sensibilités et qualité des données", "sous_titre": "La rotation du personnel pèse plus que l'âge de départ ou le taux d'actualisation",
         "synthese": f"Sans rotation, la provision passerait de {F('PIDR sens base', _s0)} à {F('PIDR sens sans rot', next(x['dbo'] for x in PSE if 'rotation' in x['hypothese']))} ; {N('PIDR embauches modifiées', len(_chg))} dates d'embauche changent dans la DISA 2025"},
        "Chaque ligne modifie une seule hypothèse et garde les autres à leur valeur centrale (IAS 19, § 145). La seconde partie relève les incohérences des fichiers du personnel qui touchent directement le calcul.",
        tableau={"titre": tab("Sensibilité de la provision du personnel au 31/12/2025"), "entetes": ["Hypothèse", "Provision", "Écart à l'hypothèse centrale"], "largeurs": [12.0, 8.0, 8.0], "taille": 9,
                 "lignes": [[x["hypothese"], F(f"PIDR sens {x['hypothese']}", x["dbo"]), pc(f"PIDR écart {x['hypothese']}", x["dbo"] / _s0 - 1)] for x in PSE]},
        disposition="dessous",
        analyse=[f"**{N('PIDR âge 55 atteint', len(_ag55))} agents ont déjà {N('PIDR âge (t)', PP['age_depart'])} ans ou plus et sont toujours en poste en 2025, dont un de {N('PIDR âge max', max(x['age'] for x in _ag55))} ans.** Le FPPN n'applique donc pas l'âge de 55 ans ; l'âge retenu par le FPPN est à confirmer.",
                 f"**Pour {N('PIDR embauches (t)', len(_chg))} agents, la date d'embauche de la DISA 2025 diffère de celle des DISA 2021 à 2023**, dont {N('PIDR embauches 2013', sum(1 for c in _chg if c['dates'].get('2025') == '28/10/2013'))} remplacées par le 28/10/2013. Pour l'agent A25-01, embauché le 01/01/2006 selon les DISA 2021 à 2023, l'ancienneté retenue baisse de plus de sept ans.",
                 f"**L'indemnité projetée la plus élevée atteint {F('PIDR max indiv', V9['pidr_max_individuel_2025'])}** : le plafond de 25 fois le SMIG annuel (convention, art. 40) peut s'appliquer, mais le montant du SMIG ne figure dans aucun document reçu."],
        implication="Le FPPN doit confirmer l'âge de départ appliqué, les dates d'embauche, le rattachement des agents recrutés en 2025 et le nombre d'enfants à charge (PBC, item 113).",
        valable="personnel civil de la DISA 2025, évaluation au 31/12/2025 ; plafond conventionnel non appliqué",
        source=SRC_DISA + " ; IAS 19, § 145 ; onglets 14a à 14c du classeur v9")
