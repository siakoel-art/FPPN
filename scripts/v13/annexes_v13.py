# ---- v13 (13/11/2026) : annexes ajoutées après la passe (validité, sigles) : largeurs au contenu et découpage en 12 points
_NSL13b = []
for _x in SL:
    if _x["type"] == "annexe" and _x.get("tableau") and not _x.get("_v13"):
        _tb = _x["tableau"]
        if _tb.get("lignes") and len(_tb.get("largeurs", [])) == len(_tb["entetes"]):
            _tb["largeurs"] = _largeurs13(_tb["entetes"], _tb["lignes"], sum(_tb["largeurs"]))
        _parts = _decoupe13(_x)
        for _y in _parts:
            _y["_v13"] = True
        _NSL13b += _parts
    else:
        _NSL13b.append(_x)
SL[:] = _NSL13b
