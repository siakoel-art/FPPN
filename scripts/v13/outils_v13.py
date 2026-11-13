# ---- v13 (13/11/2026) : outils communs (chiffres K13, mise en forme, ratios de contrôle)


def n13(x, dec=0):
    """Nombre au format français : espace des milliers, virgule décimale."""
    t = f"{abs(x):,.{dec}f}".replace(",", " ").replace(".", ",")
    return ("-" if x < 0 and round(abs(x), dec) != 0 else "") + t


def k13(cle):
    if cle.endswith("_abs"):
        return abs(K13[cle[:-4]])
    return K13[cle]


def fk13(texte):
    """Remplace {cle:p} (deux décimales) et {cle:f} (entier) par les chiffres de chiffres_v13.json."""
    def _r(m):
        v = k13(m.group(1))
        return n13(v, 2) if m.group(2) == "p" else n13(v, 0)
    return re.sub(r"\{(\w+):([pf])\}", _r, texte)


# ratios de contrôle : tableau 1 de la Décision n° 05/2022 (non-respect « entraîne des sanctions ») et section II (risques longs)
CTRL_T1 = ("FNG", "CHF", "CHB", "CHP", "SIN", "TTR", "OSA", "OSF", "AUT", "ROB", "EQS")
CTRL_RL = ("RL1", "RL2", "RL3", "RL4", "RL5")


def _g13(d, a):
    return d.get(a, d.get(str(a))) if isinstance(d, dict) else None


def _rv13(code, a):
    return _g13(RT[code]["valeurs"], a)


def _nc2(code):
    st = RT[code].get("statuts") or {}
    return all(_g13(st, a) == "Non conforme" for a in (2024, 2025))


NC_T1 = [c for c in CTRL_T1 if c in RT and _nc2(c)]
NC_RL = [c for c in CTRL_RL if c in RT and _nc2(c)]
NOMS_CTRL = {c: RT[c]["nom"] for c in NC_T1 + NC_RL}
