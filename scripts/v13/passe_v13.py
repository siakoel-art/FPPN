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
    return sum(max(1, _m13.ceil(len(str(p)) / max(1.0, largeur * 4.2))) for p in str(txt).split("\n"))


def _haut13(ligne, largeurs):
    return 0.22 + 0.48 * max(_lignes13(c, w) for c, w in zip(ligne, largeurs))


_dec13 = {}
_NSL13 = []
for _x in SL:
    tb = _x.get("tableau") if _x["type"] == "contenu" else None
    if not tb or _x.get("graphique") or _x.get("graphiques") or _x.get("_dettes13"):
        _NSL13.append(_x)
        continue
    tb["taille"] = 12
    larg = tb["largeurs"]
    dispo = BAS_CONTENU - 3.85 - (0.75 if tb.get("titre") else 0) - 0.7 - (0.62 if tb.get("legende") or tb.get("statut_col") is not None else 0)
    paras = _x.get("analyse") or []
    if paras and _x.get("disposition", "dessous") == "dessous":
        nl = sum(_m13.ceil(len(str(p)) / 150) for p in paras) + (1 if _x.get("implication") else 0)
        dispo -= min(6.0, 0.6 * nl + 0.6)
    h_ent = _haut13(tb["entetes"], larg)
    blocs, cour, h = [], [], h_ent
    for i_, ligne in enumerate(tb["lignes"]):
        hl = _haut13(ligne, larg)
        if cour and h + hl > dispo:
            blocs.append(cour)
            cour, h = [], h_ent
            dispo = BAS_CONTENU - 3.85 - (0.75 if tb.get("titre") else 0) - 0.7 - (0.62 if tb.get("legende") or tb.get("statut_col") is not None else 0)
        cour.append(i_)
        h += hl
    blocs.append(cour)
    if len(blocs) == 1:
        _NSL13.append(_x)
        continue
    t0 = _x["entete"]["titre"]
    for k_, idx in enumerate(blocs, 1):
        y = dict(_x)
        y["entete"] = dict(_x["entete"], titre=f"{t0} ({k_}/{len(blocs)})")
        tb2 = dict(tb, lignes=[tb["lignes"][i] for i in idx])
        if tb.get("fonds"):
            tb2["fonds"] = {(idx.index(r), c): v for (r, c), v in tb["fonds"].items() if r in idx}
        if k_ > 1:
            tb2["titre"] = (tb["titre"] + " (suite)") if tb.get("titre") else None
            for cle in ("analyse", "implication", "encadre"):
                y.pop(cle, None)
        y["tableau"] = tb2
        _NSL13.append(y)
    _dec13[t0] = f"{t0} (1/{len(blocs)})"
    bl13(f"{t0} (1/{len(blocs)})")
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
