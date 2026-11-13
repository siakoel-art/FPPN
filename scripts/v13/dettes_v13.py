# ---- v13 (13/11/2026) : six diapositives sur les dettes (module diapos_dettes_v17_fppn) placées sur leurs emplacements
# Données : balances retraitées v13 (extourne comprise). Lectures A, B et C de l'argent disponible, intérêts selon la règle des entretiens.
import diapos_dettes_v17_fppn as _D17
from pptx.oxml.ns import qn as _qn17

_b17 = pd.read_parquet(os.path.join(BASE, "compta_balances.parquet"))
_b17 = _b17[_b17.nature == "comptes"].copy()
_b17["type"] = _b17.compte.astype(str).map(_D17.type_compte)
_b17 = _b17[_b17.type.notna()]
_b17["dette"] = (_b17.sol_c - _b17.sol_d).clip(lower=0)
_t17 = _b17.pivot_table(index="type", columns="exercice", values="dette", aggfunc="sum").fillna(0).loc[_D17.TYPES]

_titres17 = [[sh.text_frame.text for sh in sl.shapes if sh.has_text_frame and sh.name == "Adenka_Titre"] for sl in prs.slides]
_emp17 = [i for i, t in enumerate(_titres17) if t and t[0].startswith("Dettes ") and " : " in t[0] and t[0][7:8].isdigit()]
assert len(_emp17) == 6, _emp17
_fond17 = None
for sl in prs.slides:
    pics = [sh for sh in sl.shapes if sh.shape_type == 13 and abs(sh.left) < 50000 and abs(sh.top) < 50000 and sh.width >= prs.slide_width * 0.98]
    if pics and any(sh.has_text_frame and sh.name == "Adenka_SousTitre" for sh in sl.shapes):
        _fond17 = pics[0].image.blob
        _D17.MISE_EN_PAGE["layout"] = sl.slide_layout
        _D17.MISE_EN_PAGE["h_fond"] = pics[0].height
        break
_n0 = len(prs.slides)
_D17.ajouter(prs, _fond17, _t17)
# retrait des diapositives 5 et 6 du module (chiffres d'avant l'extourne), remplacées ci-dessous
_lst17 = prs.slides._sldIdLst
for _sid in list(_lst17)[-2:]:
    prs.part.drop_rel(_sid.get(_qn17("r:id")))
    _lst17.remove(_sid)
# titres et sous-titres du module : pas de sigle (« FPPN » devient « Fonds »)
for _k in range(4):
    for _sh in prs.slides[_n0 + _k].shapes:
        if _sh.has_text_frame and _sh.name in ("Adenka_Titre", "Adenka_SousTitre"):
            for _p in _sh.text_frame.paragraphs:
                for _r in _p.runs:
                    _r.text = _r.text.replace("Le FPPN", "Le Fonds").replace("le FPPN", "le Fonds").replace("du FPPN", "du Fonds")
# texte de la diapositive 2 : hausse mesurée après extourne
for _sh in prs.slides[_n0 + 1].shapes:
    if _sh.has_text_frame:
        for _p in _sh.text_frame.paragraphs:
            for _r in _p.runs:
                if "ont presque triplé en huit mois" in _r.text:
                    _r.text = _r.text.replace("ont presque triplé en huit mois", "ont augmenté en huit mois")

_nb17, _pc17 = _D17.nb, _D17.pc
_disp, _dct, _prov, _bes, _sor, _int = K13["actif_un_an"], K13["dettes_hors_epargne"], K13["provisions"], K13["besoin"], 3028376687, K13["interets_pep_dus"]
_A, _B = _dct + _prov + _sor, _dct + _bes + _sor
_C = _B + _int
_SRC5 = ["(1) « BALANCE FPPN DU 01-01-2026 AU 31-08-2026.xlsx », extourne de la provision de la clinique ajoutée par Adenka : actif disponible dans l'année et dettes de la classe 4 hors épargne.",
         "(2) Estimation Adenka des soins et décès survenus non payés (classeur, onglets 18 à 20). Le FPPN ne calcule pas ses provisions.",
         "(3) Épargne au 31/08/2026 (41195100 et 41195200) × taux moyen de retrait 2021 à 2025 (19,24 %), classeur, onglet 24a.",
         "(4) 3,5 % du solde du plan d'épargne au 31/08/2026, règle des entretiens (3,5 % une seule fois)."]
s5 = _D17.nouvelle(prs, _fond17, f"Il manque au moins {_nb17(_B - _disp)} FCFA pour payer l'année",
                   f"Pour 100 FCFA à payer dans l'année, le Fonds dispose de {_nb17(100 * _disp / _B, 2)} FCFA avec les provisions au besoin, {_nb17(100 * _disp / _A, 2)} selon son bilan", _SRC5)
_D17.tableau(s5, 1.4, 2.8, 31.2, ["", "A. Bilan tel quel", "B. Provisions au besoin", "C. B et intérêts du plan d'épargne"],
             [["Dettes de l'année, hors épargne (1)", _nb17(_dct), _nb17(_dct), _nb17(_dct)],
              ["Soins et décès en attente", f"{_nb17(_prov)} (1)", f"{_nb17(_bes)} (2)", f"{_nb17(_bes)} (2)"],
              ["Épargne retirée dans l'année (3)", _nb17(_sor), _nb17(_sor), _nb17(_sor)],
              ["Intérêts du plan d'épargne (4)", "", "", _nb17(_int)],
              ["**Total à payer dans l'année**", f"**{_nb17(_A)}**", f"**{_nb17(_B)}**", f"**{_nb17(_C)}**"],
              ["Argent disponible dans l'année (1)", _nb17(_disp), _nb17(_disp), _nb17(_disp)],
              ["**Pour 100 FCFA à payer**", f"**{_nb17(100 * _disp / _A, 2)}**", f"**{_nb17(100 * _disp / _B, 2)}**", f"**{_nb17(100 * _disp / _C, 2)}**"]],
             [9.6, 7.0, 7.0, 7.6], 12, 0.95)
_D17.texte(s5, 1.4, 11.3, 31.2, 4.0,
           [f"**La lecture B est la plus juste : il manque {_nb17(_B - _disp)} FCFA pour payer l'année.** La lecture A garde des provisions figées depuis 2020 qui ne correspondent à aucun calcul.",
            f"**La lecture C ajoute les intérêts du plan d'épargne s'ils étaient tous payés dans l'année : {_nb17(100 * _disp / _C, 2)} FCFA pour 100 FCFA à payer.** [[Les intérêts de l'entraide restent à ajouter quand sa règle sera connue.]]"], 13)
_SRC6 = ["(1) « MANUEL DE PRODUIT D'EPARGNE FPPN VF VF.docx », partie PEP. « CARTOGRAPHIE DES RISQUES COUVERTS ET DES PRESTATIONS OFFERTES.docx ». « CAHIER DE CHARGEEFC & PEP 2021.pdf ».",
         "(2) Grands livres 2020 à 2026, compte 41195100 : 4 987 remboursements sur 6 482 égaux au versé × 1,035, calcul Adenka.",
         "(3) Calcul Adenka sur les soldes moyens des comptes 41195100 et 41195200, du 01/01/2020 au 31/08/2026, sans capitalisation (classeur, onglet 21)."]
s6 = _D17.nouvelle(prs, _fond17, "Le Fonds paie des intérêts sans jamais les inscrire en charges",
                   f"Selon la règle des entretiens, 3,5 FCFA pour 100 FCFA versés, une seule fois, soit 3,5 % du montant versé : {_nb17(_int)} FCFA dus sur le plan d'épargne au 31/08/2026", _SRC6)
_D17.tableau(s6, 1.4, 2.8, 31.2, ["Règle", "Taux", "Intérêts dus au 31/08/2026"],
             [["Règle appliquée par le Fonds (entretiens, comptes) (2)", "Plan d'épargne : 3,5 % une seule fois", _nb17(_int)],
              ["Manuel des produits d'épargne (1)", "Plan d'épargne 3,5 % et entraide 5,5 % par an", "3 278 913 238 (3)"],
              ["Tout à 3,5 % par an", "Plan d'épargne et entraide", "2 541 799 542 (3)"],
              ["Cartographie des prestations (1)", "Entraide 3,5 % puis 5,5 % à partir de 2024", "2 954 084 955 (3)"],
              ["Cahier des charges de 2021 (1)", "Entraide : 14 %, période non dite", "Non calculé"]],
             [11.6, 11.0, 8.6], 12, 1.0)
_D17.texte(s6, 1.4, 9.4, 31.2, 6.0,
           [f"**Le Fonds rend 3,5 FCFA pour 100 FCFA versés, une seule fois, quelle que soit la durée d'épargne, soit 3,5 % du montant versé (cotisé) (2).** De 09/2020 à 08/2026, ces intérêts font {_nb17(K13['interets_payes_pep'])} FCFA, pris sur le compte d'épargne sans charge.",
            "**Le manuel promet 3,5 % « tous les ans » (1) :** la différence avec la règle appliquée doit être tranchée par les organes du Fonds, puis écrite et communiquée aux policiers. [[Texte fixant la règle à recevoir.]]",
            "**[[La règle de l'entraide de fin de carrière est à compléter]] :** ses règlements sont des montants ronds qui ne suivent pas les versements."], 13)

# déplacement sur les emplacements et retrait des emplacements
_ids17 = list(_lst17)
_nouv17 = _ids17[_n0:_n0 + 6]
for _k, _i in enumerate(_emp17):
    _ids17[_i].addprevious(_nouv17[_k])
for _el in [x for x in list(_lst17) if x in [_ids17[i] for i in _emp17]]:
    prs.part.drop_rel(_el.get(_qn17("r:id")))
    _lst17.remove(_el)
print("v13 : diapositives sur les dettes placées :", len(_nouv17))
