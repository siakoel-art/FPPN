# ---- v13 (13/11/2026) : passe avant la numérotation des diapositives
# 1. titres et sous-titres réécrits (titres_v13.py), en bleu ; 2. renvois mis à jour ; 3. diapositives nouvelles placées ;
# 4. tableaux en 12 points découpés en 1/2, 2/2 quand ils ne tiennent pas sur une diapositive.
import importlib.util as _iu13
import math as _m13
from collections import defaultdict as _dd13

for _i, _x in enumerate(SL, 1):
    _x["_p0"] = _i  # position avant la passe (numéro enregistré dans TRACE, FIG et TAB)
_sp13 = _iu13.spec_from_file_location("titres_v13", os.path.join(V13, "titres_v13.py"))
_tm13 = _iu13.module_from_spec(_sp13)
_sp13.loader.exec_module(_tm13)
_CLES13 = json.load(open(os.path.join(V13, "cles_v12.json"), encoding="utf-8"))
_SYN13 = getattr(_tm13, "SYN", {})
_par13 = _dd13(list)
for _n in sorted(_tm13.T):
    _par13[_CLES13[str(_n)]].append(_n)
_vus13, _REN13 = _dd13(int), {}
for _x in SL:
    if _x["type"] not in ("contenu", "gab") or _x.get("_dettes13"):
        continue
    _t0 = _x["entete"].get("titre")
    for _a, _b in (("de la chaîne de traitement", "du traitement des données"), ("la chaîne de traitement", "le traitement des données"),
                   ("Chaîne de traitement", "Traitement des données"), ("chaîne de traitement", "traitement des données")):
        _t0 = (_t0 or "").replace(_a, _b)
    if _t0 in _par13 and _vus13[_t0] < len(_par13[_t0]):
        _n = _par13[_t0][_vus13[_t0]]
        _vus13[_t0] += 1
        _nt, _ns = (fk13(v) for v in _tm13.T[_n])
        _x["entete"] = dict(_x["entete"], titre=_nt, sous_titre=_ns)
        if _n in _SYN13:
            _x["entete"]["synthese"] = fk13(_SYN13[_n])
        bl13(_nt)
        bl13(_ns)
        if len(_par13[_t0]) == 1:
            _REN13[_t0] = _nt
print("v13 : titres réécrits :", sum(_vus13.values()), "sur", len(_tm13.T))


def _rr13(v, table):
    if isinstance(v, str):
        if "§§" in v:
            for a_, b_ in table.items():
                v = v.replace("§§" + a_ + "§§", "§§" + b_ + "§§")
        return v
    if isinstance(v, list):
        return [_rr13(e, table) for e in v]
    if isinstance(v, tuple):
        return tuple(_rr13(e, table) for e in v)
    if isinstance(v, dict):
        return {k: _rr13(e, table) for k, e in v.items()}
    return v


# 3. placement des diapositives nouvelles
_aj13 = [x for x in SL if x.get("_apres")]
for _x in _aj13:
    SL.remove(_x)
_der13 = {}
for _x in _aj13:
    _cible = _REN13.get(_x["_apres"], _x["_apres"])
    _ref = _der13.get(_cible)
    if _ref is None:
        _ref = next(y for y in SL if y["type"] in ("contenu", "gab") and y["entete"].get("titre") == _cible)
    SL.insert(SL.index(_ref) + 1, _x)
    _der13[_cible] = _x


# 4. tableaux en 12 points : découpage
def _lignes13(txt, largeur):
    return sum(max(1, _m13.ceil(len(str(p)) / max(1.0, (largeur - 0.24) * 4.6))) for p in str(txt).split("\n"))


def _haut13(ligne, largeurs):
    return 0.16 + 0.5 * max(_lignes13(c, w) for c, w in zip(ligne, largeurs))


def _brut13(v):
    return re.sub(r"\*\*|\[\[|\]\]|§§[^§]*§§", lambda m: "diapositive 999" if m.group(0).startswith("§§") else "", str(v))


def _largeurs13(entetes, lignes, total, cw=1 / 4.9, pad=0.34):
    """Largeurs de colonnes selon le contenu : colonnes courtes (chiffres, codes) à leur largeur naturelle sur une ligne,
    reste de la place aux colonnes de texte en proportion de leur volume ; jamais sous le mot le plus long."""
    nc = len(entetes)
    cols = [[_brut13(entetes[j])] + [_brut13(l[j]) for l in lignes if j < len(l)] for j in range(nc)]
    nat = [max(len(p) for t in c for p in t.split("\n")) * cw + pad for c in cols]
    mini = [max(1.4, max((len(w) for t in c for w in t.split()), default=1) * cw + pad) for c in cols]
    if sum(nat) <= total:
        return [n + (total - sum(nat)) * n / sum(nat) for n in nat]
    courtes = [j for j in range(nc) if nat[j] <= 0.16 * total]
    longues = [j for j in range(nc) if j not in courtes]
    res = [0.0] * nc
    for j in courtes:
        res[j] = nat[j]
    reste = total - sum(res)
    if not longues or reste < sum(mini[j] for j in longues):
        # pas assez de place : chaque colonne courte redescend vers son minimum
        for j in courtes:
            res[j] = mini[j]
        reste = total - sum(res)
        longues = longues or courtes
    vol = {j: sum(len(t) for t in cols[j]) for j in longues}
    for j in longues:
        res[j] = max(mini[j], reste * vol[j] / max(1, sum(vol.values())))
    k = total / sum(res)
    return [round(w * k, 2) for w in res]


# largeurs de colonnes ajustées au contenu (tableaux des diapositives de contenu, largeur totale inchangée)
for _x in SL:
    _tb = _x.get("tableau") if _x["type"] in ("contenu", "annexe") else None
    if _tb and _tb.get("lignes") and len(_tb.get("largeurs", [])) == len(_tb["entetes"]) and not _tb.get("largeurs_fixes"):
        _tb["largeurs"] = _largeurs13(_tb["entetes"], _tb["lignes"], sum(_tb["largeurs"]))
# un sous-titre de deux lignes prend la place de la ligne de synthèse
for _x in SL:
    if _x["type"] in ("contenu", "gab") and len(str(_x["entete"].get("sous_titre") or "")) > 95:
        _x["entete"]["synthese"] = ""
_dec13 = {}


def _dispo13(x, tb, suite=False):
    """Hauteur disponible (cm) pour le tableau d'une diapositive de contenu ou d'annexe."""
    if x["type"] == "annexe":
        return 12.4
    d = BAS_CONTENU - 3.85 - (0.75 if tb.get("titre") else 0) - 0.7 - (0.62 if tb.get("legende") or tb.get("statut_col") is not None else 0)
    paras = [] if suite else (x.get("analyse") or [])
    if paras and x.get("disposition", "dessous") == "dessous":
        nl = sum(_m13.ceil(len(str(p)) / 150) for p in paras) + (1 if x.get("implication") else 0)
        d -= min(7.5, 0.75 * nl + 0.9)
    return d


def _decoupe13(x):
    """Renvoie la liste des diapositives (une ou plusieurs) qui portent le tableau de x en 12 points."""
    tb = x.get("tableau") if x["type"] in ("contenu", "annexe") else None
    if not tb or x.get("graphique") or x.get("graphiques") or x.get("_dettes13") or x.get("_partie") == "R":
        return [x]
    tb["taille"] = 12
    tb.pop("hauteur_ligne", None)
    larg = tb["largeurs"]
    dispo = _dispo13(x, tb)
    h_ent = _haut13(tb["entetes"], larg)
    blocs, cour, h = [], [], h_ent
    for i_, ligne in enumerate(tb["lignes"]):
        hl = _haut13(ligne, larg)
        if cour and h + hl > dispo:
            blocs.append(cour)
            cour, h = [], h_ent
            dispo = _dispo13(x, tb, suite=True)
        cour.append(i_)
        h += hl
    blocs.append(cour)
    # une ou deux lignes seules en fin de découpage remontent dans le bloc précédent (estimation prudente)
    if len(blocs) > 1 and len(blocs[-1]) <= 2:
        _hb = h_ent + sum(_haut13(tb["lignes"][i], larg) for i in blocs[-2] + blocs[-1])
        if _hb <= 1.10 * (_dispo13(x, tb, suite=len(blocs) > 2) if len(blocs) > 2 else _dispo13(x, tb)):
            _der = blocs.pop()
            blocs[-1] += _der
    if len(blocs) == 1:
        return [x]
    cle = "titre" if x["type"] == "annexe" else None
    t0 = x["titre"] if cle else x["entete"]["titre"]
    out = []
    for k_, idx in enumerate(blocs, 1):
        y = dict(x)
        _mm = re.search(r" \((\d+)/(\d+)\)$", t0)
        _tk = (t0 if k_ == 1 else t0[:-1] + f", suite {k_ - 1})") if _mm else f"{t0} ({k_}/{len(blocs)})"
        if cle:
            y["titre"] = _tk
        else:
            y["entete"] = dict(x["entete"], titre=_tk)
        if t0 in BLEU_T:
            BLEU_T.add(_tk)
        tb2 = dict(tb, lignes=[tb["lignes"][i] for i in idx])
        if tb.get("fonds"):
            tb2["fonds"] = {(idx.index(r), c): v for (r, c), v in tb["fonds"].items() if r in idx}
        if k_ > 1:
            tb2["titre"] = (tb["titre"] + " (suite)") if tb.get("titre") else None
            for c_ in ("analyse", "implication", "encadre"):
                y.pop(c_, None)
        y["tableau"] = tb2
        out.append(y)
    _t1 = t0 if re.search(r" \(\d+/\d+\)$", t0) else f"{t0} (1/{len(blocs)})"
    if _t1 != t0 and not cle:
        _dec13[t0] = _t1
    bl13(_t1)
    return out


def _listes13():
    """Listes des figures et des tableaux en 12 points : 32 lignes par diapositive (deux colonnes de 16)."""
    nouv = []
    for x in SL:
        if x["type"] == "liste_tableaux" and x.get("part") and x["part"][0] > 0:
            continue
        if x["type"] in ("liste_figures", "liste_tableaux"):
            n_ = max(1, _m13.ceil(len(FIG if x["type"] == "liste_figures" else TAB) / 32))
            nouv += [dict(x, part=(k_, n_)) for k_ in range(n_)]
        else:
            nouv.append(x)
    SL[:] = nouv


_listes13()
_NSL13 = []
for _x in SL:
    for _y in _decoupe13(_x):
        if _y["type"] == "annexe":
            _y["_v13"] = True
        _NSL13.append(_y)
SL[:] = _NSL13
print("v13 : tableaux découpés :", len(_dec13))
_REN13.update(_dec13)
for _i, _x in enumerate(SL):
    SL[_i] = _rr13(_x, _REN13)
    SL[_i] = _rr13(SL[_i], _dec13)

# 5. renumérotation des chiffres tracés, des figures et des tableaux (positions déplacées ou découpées)
_MAP13 = _dd13(list)
for _i, _x in enumerate(SL, 1):
    if "_p0" in _x:
        _MAP13[_x["_p0"]].append(_i)
for _t in TRACE:
    if _t["diapositive"] in _MAP13:
        _nv = _MAP13[_t["diapositive"]]
        _t["diapositive"], _t["_n13"] = _nv[0], len(_nv)
for _l in FIG + TAB:
    if _l[2] in _MAP13:
        _l[2] = _MAP13[_l[2]][0]
# 6. figures et tableaux renumérotés dans l'ordre des diapositives (diapositives nouvelles et découpées)
_MAP13N = {}
for _lst, _mot in ((FIG, "Figure"), (TAB, "Tableau")):
    _ord = sorted(_lst, key=lambda e_: (e_[2], e_[0]))
    _MAP13N[_mot] = {str(e_[0]): str(k_) for k_, e_ in enumerate(_ord, 1)}
    for k_, e_ in enumerate(_ord, 1):
        e_[0] = k_
    _lst.sort(key=lambda e_: e_[0])


def _num13(v):
    if isinstance(v, str):
        return re.sub(r"\b(Figure|Tableau) (\d+) :", lambda m: m.group(1) + " " + _MAP13N[m.group(1)].get(m.group(2), m.group(2)) + " :", v)
    if isinstance(v, list):
        return [_num13(e) for e in v]
    if isinstance(v, tuple):
        return tuple(_num13(e) for e in v)
    if isinstance(v, dict):
        return {k: (_num13(e) if k not in ("num", "_partie", "type") else e) for k, e in v.items()}
    return v


for _i, _x in enumerate(SL):
    SL[_i] = _num13(_x)
