# ---- v13 (13/11/2026) : le bleu ne marque plus que les nouveautés du 13/11/2026.
# Les marques des versions précédentes (bl, bll, blt) deviennent neutres ; bl13, bll13 et blt13 marquent les nouveautés.
_bl14, _bll14, _blt14 = bl, bll, blt


def bl(t):
    return t


def bll(t):
    return t


def blt(entete):
    return entete


bl13, bll13, blt13 = _bl14, _bll14, _blt14
