# ---- v10 : validité exacte par diapositive, renvois numérotés, retrait des mentions internes
from val_v10_fppn import VAL  # noqa: E402  (dictionnaire titre -> « Valable pour » exact)
for x in SL:
    if x["type"] == "contenu" and any(c in x for c in ("tableau", "graphique", "graphiques")):
        t_ = x["entete"]["titre"]
        if t_ in VAL:
            x["valable"] = VAL[t_]
_POS = {}
for i, x in enumerate(SL, 1):
    t_ = x["entete"].get("titre") if x["type"] == "contenu" else None
    if t_:
        assert t_ not in _POS, ("titre en double", t_)
        _POS[t_] = i
_NET = [(r"§§(.+?)§§", None), (r"(classeur des calculs|classeur) v\d+\b", "classeur des calculs"), (r"[Ee]mpreintes SHA-256", "empreintes numériques"), (r"empreinte SHA-256", "empreinte numérique"),
        (r"de la chaîne de traitement", "du traitement des données"), (r"la chaîne de traitement", "le traitement des données"), (r"Chaîne de traitement", "Traitement des données"),
        (r"chaîne de traitement", "traitement des données")]


def _nettoie(v):
    if isinstance(v, str):
        for motif, rempl in _NET:
            if rempl is None:
                def _r(m):
                    assert m.group(1) in _POS, ("renvoi introuvable", m.group(1))
                    return f"diapositive {_POS[m.group(1)]}"
                v = re.sub(motif, _r, v)
            else:
                v = re.sub(motif, rempl, v)
        return v
    if isinstance(v, list):
        return [_nettoie(e) for e in v]
    if isinstance(v, tuple):
        return tuple(_nettoie(e) for e in v)
    if isinstance(v, dict):
        return {k: _nettoie(e) for k, e in v.items()}
    return v


for i, x in enumerate(SL):
    SL[i] = _nettoie(x)
for i, x in enumerate(SL, 1):
    if len(x.get("source") or "") > 430:
        print("SOURCE LONGUE", i, len(x["source"]), (x.get("entete") or {}).get("titre", x.get("titre")))
