"""Résumé exécutif Word du volet actuariel et statistique, destiné au FPPN (lecteurs non spécialistes).

Contenu : travaux réalisés, réponses aux quatre questions, 28 constats en fiches (constat, enjeu, commentaire,
« ce que cela veut dire pour vous », plan d'actions, complexité, priorité), cartographie des risques (R1 à R10 du
rapport), matrice complexité × priorité, tableau récapitulatif, limites et prochaines étapes, lexique.
Tous les chiffres sont repris du rapport actuariel et statistique du 23/10/2026.

Usage : python3 resume_executif_fppn.py <dossier des figures> <sortie.docx>
"""
import os
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.patches import FancyBboxPatch, Rectangle  # noqa: E402

sys.path.insert(0, "/root/.claude/skills/synced/891d7168-f164-499a-9885-98883788a845_be017d1e-653f-49e6-8fd9-683fb0156b08/adenka-rapport-word/scripts")
import adenka_docx as AD  # noqa: E402
from adenka_docx import RapportAdenka  # noqa: E402
from docx import Document  # noqa: E402
from docx.enum.table import WD_TABLE_ALIGNMENT  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Cm, Pt  # noqa: E402

FIG, SORTIE = sys.argv[1:3]
os.makedirs(FIG, exist_ok=True)
NOIR, ORANGE, BLEU, GRIS = "#060606", "#EFAC62", "#112A40", "#BBBCBC"
ROUGE_R, JAUNE_R, VERT_R, ORANGE_R = "#E06666", "#FFD966", "#93C47D", "#F6B26B"
plt.rcParams.update({"font.family": "Liberation Sans", "font.size": 10, "axes.spines.top": False, "axes.spines.right": False})
ANS = ["2020", "2021", "2022", "2023", "2024", "2025", "31/08/2026"]
RAP = "Rapport actuariel et statistique du 23/10/2026"


def md(x):
    return x / 1e9


def enregistrer_fig(nom):
    chemin = os.path.join(FIG, nom)
    plt.savefig(chemin, dpi=300, bbox_inches="tight")
    plt.close()
    return chemin


# ================================================================ FIGURES
# 1. frise des travaux
fig, ax = plt.subplots(figsize=(10, 2.6))
etapes = [("Lancement", "13 et 14/10/2026\nréunions, liste\ndes documents"), ("Collecte", "15 au 23/10/2026\nsept envois\ndu FPPN"),
          ("Entretiens", "16 et 19/10/2026\nprestations, cotisations,\ncomptabilité"), ("Contrôles", "4 454 416 lignes\nde précomptes\n7 balances"),
          ("Analyses", "provisions, épargne,\nliquidité, placements,\n36 ratios"), ("Constats", "28 constats\n10 risques notés"), ("Plan d'actions", "actions classées\npar complexité\net priorité")]
for i, (t, d) in enumerate(etapes):
    x = i * 1.45
    ax.add_patch(FancyBboxPatch((x, 0.55), 1.25, 0.5, boxstyle="round,pad=0.02", fc=BLEU if i < 5 else ORANGE, ec="none"))
    ax.text(x + 0.625, 0.8, t, ha="center", va="center", color="white", fontsize=10, fontweight="bold")
    ax.text(x + 0.625, 0.25, d, ha="center", va="center", fontsize=8, color=NOIR)
    if i < len(etapes) - 1:
        ax.annotate("", xy=(x + 1.43, 0.8), xytext=(x + 1.27, 0.8), arrowprops=dict(arrowstyle="->", color=GRIS, lw=1.5))
ax.set_xlim(-0.05, len(etapes) * 1.45)
ax.set_ylim(-0.1, 1.15)
ax.axis("off")
F_FRISE = enregistrer_fig("frise_travaux.png")

# 2. épargne due et actifs réservés
ep = [10828205640, 10147421276, 9364845293, 9121576576, 10785064432, 13814066964, 15743516391]
res = [714042784, 1957211038, 888844233, 839212069, 699925787, 1608155525, 2562885661]
cov = ["6,59 %", "19,29 %", "9,49 %", "9,20 %", "6,49 %", "11,64 %", "16,28 %"]
fig, ax = plt.subplots(figsize=(9, 3.8))
xs = range(len(ANS))
ax.bar([x - 0.2 for x in xs], [md(v) for v in ep], 0.4, color=BLEU, label="Épargne due aux policiers")
ax.bar([x + 0.2 for x in xs], [md(v) for v in res], 0.4, color=ORANGE, label="Argent mis de côté pour cette épargne")
for x, v, c in zip(xs, res, cov):
    ax.text(x + 0.2, md(v) + 0.25, c, ha="center", fontsize=8)
ax.set_xticks(list(xs), ANS)
ax.set_ylabel("Milliards de FCFA")
ax.legend(frameon=False, loc="upper left")
F_EPARGNE = enregistrer_fig("epargne_actifs.png")

# 3. sinistralité
mal = [12.54, 6.00, 6.47, 26.91, 48.13, 73.53, 97.40]
dec = [82.23, 55.42, 15.21, 17.05, 21.81, 19.14, 21.76]
glob = [8.92, 20.77, 9.06, 24.56, 40.34, 57.48, 75.20]
fig, ax = plt.subplots(figsize=(9, 3.8))
ax.axhspan(70, 80, color=GRIS, alpha=0.35, label="Fourchette de l'AIRMS (70 % à 80 %)")
ax.plot(ANS, mal, color=BLEU, marker="o", lw=2, label="Maladie")
ax.plot(ANS, dec, color=ORANGE, marker="o", lw=2, label="Décès")
ax.plot(ANS, glob, color=NOIR, marker="s", lw=1.2, ls="--", label="Toutes garanties (taux réglementaire)")
for x, v in zip(ANS, mal):
    ax.text(x, v + 3, f"{v:.2f} %".replace(".", ","), ha="center", fontsize=7.5, color=BLEU)
ax.set_ylabel("% des cotisations")
ax.set_ylim(0, 110)
ax.legend(frameon=False, fontsize=8, loc="upper left")
F_SIN = enregistrer_fig("sinistralite.png")

# 4. fonds propres et résultat
fp = [-6743081733, -4659173531, -2598840854, -1668211752, -876142111, -438242200, -879498700]
rs = [1018579699, 2083908202, 2060332677, 930629102, 792069641, 437899911, -441256500]
fig, ax = plt.subplots(figsize=(9, 3.6))
ax.bar([x - 0.2 for x in xs], [md(v) for v in fp], 0.4, color=BLEU, label="Fonds propres (réserves moins pertes passées)")
ax.bar([x + 0.2 for x in xs], [md(v) for v in rs], 0.4, color=ORANGE, label="Résultat de l'année")
ax.axhline(0, color=NOIR, lw=0.8)
ax.set_xticks(list(xs), ANS)
ax.set_ylabel("Milliards de FCFA")
ax.legend(frameon=False, fontsize=8, loc="lower right")
F_FP = enregistrer_fig("fonds_propres.png")

# 5. provisions au bilan et besoins mesurés
fig, ax = plt.subplots(figsize=(9, 3.4))
lib = ["Maladie", "Décès (obsèques)", "Intérêts dus\nsur l'épargne", "Pertes sur les\nprêts du Grenier"]
bil = [2867602878, 1961695412, 0, 0]
bes = [12481392, 7100000, 3278913238, 14016392]
xs2 = range(len(lib))
ax.bar([x - 0.2 for x in xs2], [md(v) for v in bil], 0.4, color=BLEU, label="Montant au bilan du FPPN")
ax.bar([x + 0.2 for x in xs2], [md(v) for v in bes], 0.4, color=ORANGE, label="Besoin mesuré par Adenka")
for x, v in zip(xs2, bil):
    ax.text(x - 0.2, md(v) + 0.05, f"{md(v):.2f}".replace(".", ","), ha="center", fontsize=8)
for x, v in zip(xs2, bes):
    ax.text(x + 0.2, md(v) + 0.05, f"{md(v):.2f}".replace(".", ","), ha="center", fontsize=8)
ax.set_xticks(list(xs2), lib)
ax.set_ylabel("Milliards de FCFA")
ax.legend(frameon=False, fontsize=8)
F_PROV = enregistrer_fig("provisions.png")

# 6. écart de liquidité cumulé
cum = [-6588246233, -8709764501, -10360806802, -11631872515, -12596052822]
fig, ax = plt.subplots(figsize=(9, 3.2))
ax.bar(["Année 1", "Année 2", "Année 3", "Année 4", "Année 5"], [md(v) for v in cum], color=BLEU)
for i, v in enumerate(cum):
    ax.text(i, md(v) - 0.6, f"{md(v):.2f}".replace(".", ","), ha="center", fontsize=8, color=NOIR)
ax.axhline(0, color=NOIR, lw=0.8)
ax.set_ylabel("Manque cumulé (milliards de FCFA)")
ax.set_ylim(-14.5, 0.5)
F_LIQ = enregistrer_fig("ecart_liquidite.png")

# 7. matrice des risques (R1 à R10 du rapport, diapositive 23)
POS = {"R1": (4, 5), "R2": (4, 5), "R3": (5, 4), "R4": (5, 4), "R5": (5, 3), "R6": (5, 3), "R9": (5, 3), "R10": (4, 3), "R7": (5, 1), "R8": (5, 2)}
fig, ax = plt.subplots(figsize=(7.5, 5.2))
for p in range(1, 6):
    for g in range(1, 6):
        sc = p * g
        c = ROUGE_R if sc >= 15 else (JAUNE_R if sc >= 8 else VERT_R)
        ax.add_patch(Rectangle((g - 0.5, p - 0.5), 1, 1, fc=c, ec="white", lw=2))
        txt = ", ".join(k for k, v in POS.items() if v == (p, g))
        if txt:
            ax.text(g, p, txt, ha="center", va="center", fontsize=10, fontweight="bold")
ax.set_xlim(0.5, 5.5)
ax.set_ylim(0.5, 5.5)
ax.set_xticks(range(1, 6), ["1\nNégligeable", "2\nMineure", "3\nModérée", "4\nMajeure", "5\nCatastrophique"], fontsize=8)
ax.set_yticks(range(1, 6), ["1 Très peu\nprobable", "2 Peu\nprobable", "3 Possible", "4 Probable", "5 Très probable\n(déjà constaté)"], fontsize=8)
ax.set_xlabel("Gravité (impact en FCFA)")
ax.set_ylabel("Probabilité")
for s in ax.spines.values():
    s.set_visible(False)
F_RISQ = enregistrer_fig("matrice_risques.png")


# ================================================================ CONSTATS
def C(num, theme, titre, constat, pourquoi, commentaire, vous, actions, qui, cplx, prio, risque="", diapos="", confirmer=""):
    return dict(num=num, theme=theme, titre=titre, constat=constat, pourquoi=pourquoi, commentaire=commentaire, vous=vous,
                actions=actions, qui=qui, cplx=cplx, prio=prio, risque=risque, diapos=diapos, confirmer=confirmer)


TH = {"A": "Provisions et dettes envers les adhérents", "B": "Épargne des policiers (EFC et PEP)", "C": "Liquidité : l'argent sera-t-il disponible à temps ?",
      "D": "Conformité aux règles et solidité financière", "E": "Garantie maladie et cotisations", "F": "Placements et décisions",
      "G": "Prêts du Grenier du Policier", "H": "Données et contrôle interne"}
K = [
    C(1, "A", "Les provisions maladie et décès ne sont pas calculées",
      "Les provisions pour prestations à payer valent environ 4,8 milliards de FCFA (4 829 298 290 FCFA) : 2 867 602 878 FCFA en maladie et 1 961 695 412 FCFA en décès. Ce montant est le même depuis au moins le 01/09/2020. Le besoin réel mesuré n'est que de 19 581 392 FCFA. Les provisions dépassent donc le besoin d'environ 4,8 milliards de FCFA (4 809 716 898 FCFA).",
      "Une provision est l'argent mis de côté pour payer les soins et les décès déjà survenus mais pas encore payés. Les statuts (art. 89) et la réglementation (règlement d'exécution n° 004/2022, art. 10) imposent de la recalculer à chaque clôture.",
      "Le problème n'est pas un manque d'argent mis de côté, mais un argent mal rangé. Ces montants ne correspondent à aucun calcul : ils gonflent les dettes du bilan et font paraître le FPPN plus pauvre qu'il ne l'est sur ce poste. À l'inverse, rien n'est mis de côté pour les intérêts de l'épargne (constat 6). Le FPPN dispose déjà des fichiers nécessaires pour faire le calcul chaque année.",
      "L'argent existe, mais il n'est pas affecté au bon endroit. Bien rangé, il couvrirait mieux les vraies dettes envers les policiers.",
      ["Établir chaque année une note de calcul des provisions (méthode, fichiers, montant), signée par le directeur général et présentée à l'Assemblée générale.",
       "Faire décider par les organes le reclassement de l'excédent, après information de l'AIRMS."],
      "Service comptable, Direction générale, Assemblée générale", "Moyenne", "P1", "R4", "6, 123, 126, 144 et 145"),
    C(2, "A", "La dette envers les familles des policiers décédés augmente chaque année",
      "Les allocations décès dues et non payées passent de 370 503 332 FCFA fin 2021 à 489 419 669 FCFA au 31/08/2026. Chaque année, le FPPN inscrit plus d'allocations qu'il n'en paie. La date de dépôt de chaque demande n'est pas enregistrée.",
      "Ces sommes sont dues aux familles. Un dossier ancien non payé est une famille qui attend. Sans date de dépôt, le FPPN ne peut pas savoir quels dossiers sont encore valables (délai de 3 ans pour un membre et de 12 mois pour un ayant droit, règlement intérieur, art. 58 et 60).",
      "Le FPPN calcule bien sa dette décès chaque fin d'année, dossier par dossier, et l'inscrit en comptabilité au franc près en 2024 et 2025 : c'est une bonne pratique. Mais le stock grossit, ce qui montre des dossiers qui traînent depuis plusieurs années.",
      "Une partie des familles de policiers décédés attend encore son allocation.",
      ["Dresser la liste nominative des dossiers non payés, avec la date de dépôt et le motif du retard.", "Payer les dossiers dus et clôturer ceux qui sont prescrits, sur décision écrite."],
      "Service des prestations, service comptable", "Moyenne", "P2", "", "125 et 126"),
    C(3, "A", "Des allocations décès sont payées avec l'argent de l'épargne",
      "En 2024, 102 allocations décès ont été payées depuis les comptes bancaires de l'épargne : 93 paiements pour 49 000 000 FCFA depuis le compte PEP et 9 paiements pour 4 288 000 FCFA depuis le compte EFC, soit 53 288 000 FCFA.",
      "L'épargne appartient aux policiers épargnants. Elle ne doit servir qu'à leur rembourser leur épargne. La réglementation impose de séparer l'argent de l'épargne de celui des autres garanties (règlement d'exécution n° 004/2022, art. 15).",
      "Le montant est faible, mais le principe est important : mélanger les comptes empêche de savoir ce qui reste réellement aux épargnants.",
      "Votre épargne doit rester sur des comptes qui ne servent qu'à elle.",
      ["Payer les prestations uniquement depuis les comptes des garanties maladie et décès.", "Rembourser aux comptes de l'épargne les sommes utilisées."],
      "Direction générale, service comptable", "Faible", "P2", "", "126"),
    C(4, "A", "Les pertes probables sur les prêts du Grenier ne sont pas comptées",
      "Au 31/08/2026, 143 prêts ont plus de 3 mois de retard. La politique de crédit du Grenier impose de compter une perte de 40 %, 80 % ou 100 % selon l'ancienneté du retard, soit 14 016 392 FCFA. Aucune perte n'est comptée depuis 2022.",
      "Un prêt en retard depuis longtemps risque de ne jamais être remboursé. Ne pas le compter fait paraître les comptes meilleurs qu'ils ne sont.",
      "Le montant reste faible au regard de l'encours (environ 1 milliard de FCFA). Le vrai risque du Grenier n'est pas la perte sur les prêts, mais l'épargne qu'il immobilise (constat 25).",
      "Les prêts non remboursés finissent par coûter à tous les adhérents.",
      ["Comptabiliser chaque année la dépréciation des prêts en retard, selon la politique de crédit du Grenier."],
      "Grenier, service comptable", "Faible", "P3", "R7", "129"),
    C(5, "A", "La provision pour le départ à la retraite du personnel est suffisante",
      "Le FPPN a mis de côté 43 666 678 FCFA au 31/12/2025 pour les indemnités de départ à la retraite de ses salariés. Le recalcul agent par agent donne 24 446 566 FCFA.",
      "Le FPPN devra payer une indemnité à chaque salarié qui part à la retraite.",
      "C'est le seul poste où le montant du FPPN couvre largement le besoin. Le FPPN calcule ce montant comme si tous les salariés partaient aujourd'hui, ce qui est prudent. Seule la méthode n'est pas écrite.",
      "Ce poste ne présente pas de risque pour les adhérents.",
      ["Rédiger une note de calcul annuelle de cette provision (méthode, liste des salariés)."],
      "Service comptable", "Faible", "P4", "", "133"),
    C(6, "B", "Les intérêts promis aux épargnants ne sont pas comptabilisés",
      "Aucun intérêt n'a été inscrit sur l'épargne du 01/01/2020 au 31/08/2026. Aux taux du manuel d'épargne (3,50 % pour le PEP et 5,50 % pour l'EFC), les intérêts dus atteignent environ 3,3 milliards de FCFA (3 278 913 238 FCFA). L'épargne réellement due passe ainsi d'environ 15,7 milliards de FCFA (15 743 516 391 FCFA) à environ 19 milliards de FCFA (19 022 429 629 FCFA).",
      "C'est la plus grosse dette du FPPN : l'épargne et ses intérêts représentent 97,46 % de ce que le FPPN doit à ses adhérents.",
      "Le calcul retient le montant le plus bas possible : intérêts simples, sans intérêts sur les intérêts, et sans les intérêts antérieurs à 2020. La vraie dette est donc au moins de ce niveau. Même au taux le plus bas (3,50 %), les intérêts d'une seule année sur l'EFC (309 709 577 FCFA en 2025) dépassent ce que rapportent tous les placements du FPPN dans l'année.",
      "Les intérêts de votre épargne existent sur le papier, mais le FPPN ne les a pas encore enregistrés. Votre épargne vaut plus que ce que montrent les comptes.",
      ["Calculer et comptabiliser chaque année les intérêts dus, épargnant par épargnant, dès que le taux est arrêté (constat 7)."],
      "Service comptable, Conseil d'administration", "Moyenne", "P1", "R3", "5, 97 et 98",
      "Le montant exact dépend du taux qui sera arrêté : il serait de 2 541 799 542 FCFA si l'EFC était servie à 3,50 %."),
    C(7, "B", "Aucun taux d'intérêt de l'épargne n'est officiellement fixé",
      "Trois taux différents figurent dans trois documents pour l'EFC : 3,50 %, 5,50 % et 14 %. Aucun n'est fixé par une décision signée du conseil d'administration.",
      "Sans taux officiel, ni le FPPN ni l'épargnant ne peuvent savoir combien vaut l'épargne. Un épargnant peut contester le montant reçu à son départ.",
      "Le taux doit être réaliste : les placements du FPPN ont rapporté 4,90 % par an en moyenne de 2021 à 2025. Promettre plus que ce que rapporte l'argent crée une perte chaque année, que les cotisations maladie et décès finissent par payer.",
      "Le taux qui rémunère votre épargne doit être fixé clairement, par écrit, et être le même pour tous.",
      ["Faire adopter par le conseil d'administration une délibération fixant le taux servi sur le PEP et l'EFC, cohérent avec le rendement des placements.", "Communiquer ce taux à tous les épargnants."],
      "Conseil d'administration", "Faible", "P1", "R3", "95"),
    C(8, "B", "Le FPPN ne tient pas de compte individuel pour chaque épargnant",
      "Aucun relevé individuel d'épargne n'a été transmis : pas de solde d'ouverture, de versements, d'intérêts, de retraits et de solde de clôture par épargnant.",
      "Chaque policier doit pouvoir savoir à tout moment combien il a épargné et ce qu'il recevra.",
      "Sans comptes individuels, l'épargne n'est connue qu'en total. Le FPPN ne peut ni vérifier le total, ni répondre à un épargnant qui demande son solde, ni prévoir quand l'argent sortira.",
      "Vous devriez recevoir chaque année un relevé de votre épargne.",
      ["Mettre en place un compte individuel pour chaque épargnant dans le logiciel.", "Envoyer un relevé annuel à chaque épargnant."],
      "Direction générale, service informatique", "Élevée", "P2", "R9", "5 et 143"),
    C(9, "B", "Le rapport de gestion 2025 sous-estime l'épargne due",
      "Le rapport de gestion 2025 présente environ 4,4 milliards de FCFA (4 387 823 635 FCFA) comme « le montant des épargnes au 31/12/2025 ». C'est en réalité ce qui a été collecté dans l'année. L'épargne due à cette date est d'environ 13,8 milliards de FCFA (13 814 066 964 FCFA).",
      "L'Assemblée générale prend ses décisions sur la base de ce rapport.",
      "L'erreur porte sur la présentation, pas sur la comptabilité : la balance donne le bon montant. Mais l'Assemblée générale a reçu une information qui minore de plus des deux tiers la dette envers les épargnants.",
      "Les chiffres présentés en Assemblée générale doivent montrer tout ce que le FPPN vous doit.",
      ["Présenter dans le rapport de gestion le solde de l'épargne due, distinct de la collecte de l'année."],
      "Direction générale", "Faible", "P2", "", "99"),
    C(10, "B", "L'argent de l'épargne n'est pas séparé du reste",
      "Au 31/08/2026, seulement 16,28 % de l'épargne due est couverte par des comptes bancaires, des dépôts et des placements qui lui sont réservés, soit 2 562 885 661 FCFA sur 15 743 516 391 FCFA.",
      "La réglementation impose de séparer l'épargne (risque long) des garanties maladie et décès (risques courts) (règlement d'exécution n° 004/2022, art. 15).",
      "L'épargne augmente de 45,39 % depuis 2020, et l'épargne EFC a triplé. Les sommes réservées n'ont pas suivi : le reste de l'épargne a servi à financer la clinique, les terrains, la créance sur l'Hôpital de la Police et les prêts du Grenier.",
      "Sur 100 FCFA que vous avez épargnés, environ 16 FCFA sont aujourd'hui gardés à part pour vous les rendre.",
      ["Ouvrir des comptes réservés à l'épargne et y verser chaque mois l'épargne collectée.", "Fixer, dans le plan de redressement, un objectif de couverture de l'épargne par des actifs réservés."],
      "Direction générale, Conseil d'administration", "Moyenne", "P2", "R1", "90"),
    C(11, "C", "Le FPPN ne peut pas couvrir dans l'année tout ce qu'il devra payer",
      "Au 31/08/2026, l'argent disponible en moins d'un an vaut environ 3,5 milliards de FCFA (3 457 866 951 FCFA). Ce que le FPPN devra payer dans l'année, sorties d'épargne comprises, vaut environ 10 milliards de FCFA (10 046 113 184 FCFA). Il couvre donc 34,42 % des paiements de l'année. Sur cinq ans, le manque cumulé atteint environ 12,6 milliards de FCFA (12 596 052 822 FCFA).",
      "Chaque année, environ un cinquième de l'épargne sort (19,24 % en moyenne de 2021 à 2025, jusqu'à 27,55 % en 2023), surtout au moment des départs à la retraite.",
      "Ce calcul est favorable au FPPN : il compte les biens à leur valeur comptable et ne prend pas en compte les intérêts dus. Le manque réel est donc au moins de ce niveau. En pratique, le FPPN paie les sorties grâce aux nouvelles cotisations : il dépend de la poursuite des versements des policiers en activité.",
      "Si beaucoup de policiers demandent leur épargne en même temps, le FPPN ne peut pas tous les payer à temps.",
      ["Constituer une réserve de trésorerie au moins égale aux sorties d'épargne d'une année au taux le plus élevé observé.", "Établir l'échéancier des sorties de l'épargne (départs à la retraite).", "Arrêter de financer des biens durables avec l'épargne."],
      "Conseil d'administration, Direction générale", "Élevée", "P1", "R1", "7, 149 et 154 à 156"),
    C(12, "C", "L'essentiel des biens du FPPN ne peut pas être vendu rapidement",
      "Plus des trois quarts des biens du FPPN (17 170 238 905 FCFA sur 21 925 420 866 FCFA) sont à plus de cinq ans ou à délai incertain. La clinique (7 083 534 565 FCFA), les terrains (5 245 896 750 FCFA) et la créance sur l'Hôpital de la Police (3 407 519 850 FCFA, inchangée depuis 2020) n'ont pas été évalués récemment.",
      "Un bien ne protège les adhérents que si on connaît sa vraie valeur et qu'on peut le transformer en argent.",
      "Sur le papier, les biens couvrent 96,33 % des dettes connues. Sans les biens dont la valeur n'est pas établie, la couverture tombe à 27,19 %. La créance sur l'hôpital n'a reçu aucun remboursement depuis 2020.",
      "Les biens du FPPN sont surtout des bâtiments et des terrains : ils ne peuvent pas servir vite à vous payer.",
      ["Faire évaluer la clinique et les terrains par un expert.", "Obtenir de l'État un accord écrit sur le remboursement de la créance sur l'Hôpital de la Police."],
      "Direction générale, Conseil d'administration", "Élevée", "P2", "R1", "148 à 152"),
    C(13, "C", "Une partie de la trésorerie dort sans rapporter",
      "Au 31/08/2026, 2 356 399 123 FCFA restent sur des comptes courants qui ne rapportent rien.",
      "La réglementation demande aux mutuelles qui gèrent de l'épargne de limiter ces soldes non rémunérés.",
      "Une partie de cet argent peut être placée à court terme sans perdre sa disponibilité, ce qui aiderait à payer les intérêts de l'épargne.",
      "Une partie de l'argent peut travailler pour vous sans être bloquée.",
      ["Placer à court terme la part de trésorerie qui n'est pas nécessaire aux paiements du mois."],
      "Direction générale", "Faible", "P3", "", "165"),
    C(14, "D", "Le FPPN ne respecte pas plusieurs règles de solidité financière",
      "Sur les 36 indicateurs fixés par la réglementation (Décision n° 05/2022), 14 ne sont pas respectés de 2022 à 2025, 7 sont respectés et 15 n'ont pas pu être vérifiés faute de documents. En 2025, la réserve obligatoire couvre 5,01 mois de prestations pour un minimum de 6 mois, la liquidité vaut 0,13 et les fonds propres 0,22, pour un minimum de 1.",
      "Si un indicateur n'est pas respecté deux années de suite, l'AIRMS peut retirer l'agrément du FPPN (règlement n° 07/2009, art. 84). Ce n'est pas automatique.",
      "Deux points méritent d'être nuancés. D'abord, la liquidité dépend du classement de l'épargne : classée comme le prévoit le plan comptable des mutuelles, elle devient conforme en 2025 (1,15). Ensuite, le taux d'excédent dépasse le plafond de 5 % de 2020 à 2025 alors que les fonds propres sont négatifs : ces excédents servent à combler les pertes passées. Le vrai problème reste l'épargne et la liquidité, pas un seul ratio.",
      "C'est l'enjeu principal : la continuité du FPPN dépend du respect de ces règles.",
      ["Présenter à l'AIRMS un plan de mise en conformité avec des objectifs chiffrés par indicateur sur trois exercices, validé par le conseil d'administration.", "Suivre les indicateurs tous les six mois."],
      "Conseil d'administration, Direction générale", "Élevée", "P1", "R2", "10, 13, 178 et 184"),
    C(15, "D", "Les fonds propres du FPPN sont négatifs",
      "Les fonds propres sont négatifs à chaque clôture depuis 2020 : environ -6,7 milliards de FCFA fin 2020 (-6 743 081 733 FCFA), -438 242 200 FCFA fin 2025 et -879 498 700 FCFA au 31/08/2026.",
      "Les fonds propres sont le coussin qui absorbe un coup dur. Négatifs, ils signifient que les pertes passées dépassent les réserves.",
      "La tendance était bonne jusqu'en 2025 grâce aux excédents. Le déficit des huit premiers mois de 2026 (-441 256 500 FCFA) l'interrompt. À titre d'illustration, si les provisions étaient remises à leur juste niveau (constat 1) et si les intérêts dus et les pertes sur prêts étaient comptabilisés (constats 4 et 6), les fonds propres du 31/08/2026 deviendraient positifs d'environ 637 millions de FCFA (637 288 568 FCFA). Une partie du problème vient donc du rangement des comptes.",
      "Le FPPN n'a pas encore de réserve pour faire face à une mauvaise année.",
      ["Affecter les excédents futurs aux réserves.", "Réaliser les reclassements des constats 1, 4 et 6 après décision des organes."],
      "Assemblée générale, Conseil d'administration", "Moyenne", "P2", "R6", "160 et 161",
      "Le montant de 637 288 568 FCFA est une illustration : le reclassement exige une décision des organes et l'avis de l'AIRMS."),
    C(16, "D", "Les frais de fonctionnement dépassent la norme",
      "En 2025, les frais de fonctionnement représentent 36,34 % des cotisations, et 44,12 % sur les huit premiers mois de 2026, pour un maximum réglementaire de 20 %.",
      "Chaque franc de frais est un franc qui ne sert pas aux soins, aux allocations ou à l'épargne.",
      "Le taux baisse de 84,31 % en 2020 à 36,34 % en 2025, grâce à la hausse des cotisations. Il remonte en 2026. Les frais couvrent le personnel, les indemnités des organes, le fonctionnement et les amortissements.",
      "Une plus grande part de vos cotisations doit revenir en prestations.",
      ["Fixer un budget de fonctionnement avec un objectif de baisse vers 20 % des cotisations.", "Présenter chaque année l'écart entre budget et réalisé."],
      "Direction générale, Conseil d'administration", "Moyenne", "P3", "", "13 et 39"),
    C(17, "D", "La clinique n'a pas de comptabilité séparée",
      "Le bilan du FPPN porte 7 083 534 565 FCFA pour la clinique (CMP) : 5 803 025 455 FCFA d'investissements et 1 280 300 860 FCFA de charges de fonctionnement. Aucune comptabilité séparée de la clinique n'a été transmise.",
      "La réglementation interdit de couvrir les pertes d'une activité de soins avec les cotisations ou les réserves obligatoires (règlement d'exécution n° 004/2022, art. 27).",
      "Sans comptabilité séparée, personne ne sait si la clinique gagne ou perd de l'argent, ni combien d'épargne y est engagée.",
      "Vous devez savoir combien coûte et rapporte la clinique financée avec vos cotisations et votre épargne.",
      ["Tenir une comptabilité séparée de la clinique et présenter son résultat chaque année."],
      "Direction générale, service comptable", "Moyenne", "P2", "", "148, 164 et 201"),
    C(18, "D", "Des sommes durables sont rangées en comptes d'attente",
      "Au 31/08/2026, 11 603 989 675 FCFA d'avances à la clinique et au Grenier et de créance sur l'hôpital sont suivis en comptes d'attente ou de tiers, comme s'ils devaient être remboursés à court terme.",
      "Le plan comptable des mutuelles prévoit des comptes de prêts et de placements durables pour ces sommes.",
      "Ce rangement fausse la lecture des comptes : il présente comme de court terme de l'argent bloqué pour longtemps.",
      "Les comptes doivent montrer clairement où est placé l'argent des adhérents.",
      ["Reclasser ces sommes dans les comptes prévus par le plan comptable des mutuelles."],
      "Service comptable", "Faible", "P3", "", "200"),
    C(19, "E", "La garantie maladie perd de l'argent depuis 2025",
      "La part des cotisations maladie consommée par les soins passe de 6,00 % en 2021 à 73,53 % en 2025 et 97,40 % sur les huit premiers mois de 2026. Frais compris, la maladie coûte 109,87 % de ses cotisations en 2025 et 141,52 % en 2026. Elle perd 237 774 567 FCFA en 2025 et 701 679 552 FCFA sur les huit premiers mois de 2026.",
      "Une garantie qui coûte plus qu'elle ne rapporte doit être financée par autre chose. Au FPPN, c'est la garantie décès qui la finance.",
      "Le niveau bas de 2020 à 2022 n'était pas un signe de bonne santé : la couverture était limitée aux soins liés à une hospitalisation (règlement intérieur, art. 41 et 45) et le FPPN accumulait des factures impayées envers les cliniques (1 394 774 547 FCFA fin 2020). De même, le taux global de 75,20 % en 2026 est dans la fourchette de l'AIRMS (70 % à 80 %), mais il cache une maladie très déficitaire compensée par le décès : être dans la norme ne veut pas dire être à l'équilibre.",
      "Au rythme actuel, les soins coûteront plus que les cotisations maladie. Sans mesure, il faudra revoir les cotisations ou les remboursements.",
      ["Recalculer la cotisation maladie sur le coût complet de la garantie, frais compris.", "Revoir le barème des remboursements et les conventions avec les cliniques.", "Contrôler après paiement au moins 5 % des factures (Décision n° 05/2022)."],
      "Conseil d'administration, Assemblée générale, service des prestations", "Élevée", "P2", "R5", "11, 13, 75 et 195"),
    C(20, "E", "La part des cotisations entre maladie et décès n'est pas la même partout",
      "Le prélèvement total de 5 % du salaire de base est respecté. Mais la Solde prélève 3,77 % pour la maladie et 1,23 % pour le décès, alors que les statuts (art. 17) fixent 3,50 % et 1,50 %. La comptabilité 2025 retient encore une autre répartition (3,52 % et 1,48 %).",
      "La répartition décide de ce que chaque garantie reçoit, donc de son résultat.",
      "Le policier paie le bon total. Mais sans clé unique, le résultat de la maladie et celui du décès sont incertains, alors que la réglementation impose de les suivre séparément.",
      "Votre cotisation totale est correcte. C'est son partage entre maladie et décès qui doit être harmonisé.",
      ["Notifier à la Solde la répartition fixée par les statuts et l'appliquer en comptabilité."],
      "Direction générale, responsable des cotisations", "Faible", "P2", "", "9 et 65"),
    C(21, "E", "Le paiement des cliniques ralentit en 2026",
      "Le délai moyen de paiement des cliniques passe de 53,38 mois en 2020 à 2,62 mois en 2025, puis remonte à 5,00 mois au 31/08/2026. Les factures non payées passent de 242 140 869 FCFA fin 2025 à 690 234 451 FCFA au 31/08/2026.",
      "Le règlement intérieur (art. 52) fixe le paiement des cliniques à 3 mois. Des retards répétés peuvent pousser les cliniques à refuser les policiers.",
      "Le FPPN a fait un gros effort de rattrapage de 2020 à 2025. La remontée de 2026 coïncide avec le déficit de la garantie maladie (constat 19) : elle est à surveiller de près.",
      "Des retards de paiement peuvent rendre plus difficile votre accueil dans les cliniques conventionnées.",
      ["Suivre chaque mois le délai de paiement et les factures en attente.", "Enregistrer la date des soins et la date de dépôt de chaque facture."],
      "Service des prestations, service comptable", "Moyenne", "P2", "", "120 et 121"),
    C(22, "E", "Le coût des soins des policiers retraités n'est pas suivi",
      "Les retraités qui cotisent par la CGRAE passent de 16 en février 2024 à 610 en août 2026, avec une cotisation moyenne de 11 954 FCFA par mois. Leurs soins ne sont pas suivis à part, et leur date de naissance n'est connue que pour 105 d'entre eux.",
      "Un retraité a en général plus de besoins de soins qu'un actif, et sur une longue durée.",
      "Sans suivi des soins des retraités, le FPPN ne sait pas si leur cotisation couvre leur coût, ni combien il s'engage à payer s'il les couvre à vie.",
      "Si vous êtes ou serez retraité, le FPPN doit vérifier que la garantie santé pourra être tenue dans la durée.",
      ["Suivre séparément les soins des retraités.", "Fixer par écrit la durée de la garantie santé des retraités (à vie ou limitée).", "Obtenir de la CGRAE les dates de naissance des retraités."],
      "Service des prestations, Conseil d'administration", "Moyenne", "P3", "R9", "112"),
    C(23, "F", "Les placements sont trop faibles pour servir l'épargne promise",
      "En 2025, l'argent réellement placé (dépôts à terme et fonds commun de placement) vaut en moyenne 763 752 135 FCFA, soit 5,53 % de l'épargne due. Les placements ont rapporté 4,90 % par an en moyenne de 2021 à 2025. En 2025, les intérêts dus sur l'épargne dépassent les produits des placements de 470 608 168 FCFA.",
      "Les placements doivent être sûrs, rentables, disponibles et variés (règlement d'exécution n° 004/2022, art. 17).",
      "Le rendement lui-même n'est pas mauvais : il dépasse le taux directeur de la BCEAO (2,00 % à 3,50 %). Le problème est le volume : presque toute l'épargne est immobilisée dans des biens qui ne rapportent ni intérêts ni liquidités.",
      "Votre épargne doit être placée de façon à rapporter les intérêts promis et à rester disponible.",
      ["Faire voter une politique de placement : part minimale disponible, rendement minimal au regard du taux servi, limites sur l'immobilier."],
      "Conseil d'administration", "Moyenne", "P2", "R8", "8, 162, 165 et 166"),
    C(24, "F", "Les décisions de placement ne passent pas par le conseil d'administration",
      "[[Selon le chef comptable, le directeur général décide seul des placements.]] Aucune délibération du conseil d'administration sur les placements n'a été transmise, alors que les statuts l'exigent (art. 85). Le projet immobilier de Cocody n'a ni budget ni plan de financement transmis. La désignation du directeur général n'est documentée que pour 2020 à 2023.",
      "Les décisions qui engagent l'argent des adhérents doivent être prises par les organes élus et tracées par écrit.",
      "Il ne s'agit pas de mettre en cause une personne, mais de protéger le FPPN et ses dirigeants par des décisions collectives et écrites. Le projet de Cocody engagerait l'épargne dans un bien de long terme, alors que la liquidité ne couvre pas l'année qui vient (constat 11).",
      "Les grandes décisions sur votre argent doivent être prises collectivement par vos représentants.",
      ["Faire délibérer le conseil d'administration sur chaque placement.", "Présenter aux organes le budget, le financement et l'effet sur la liquidité du projet de Cocody avant tout engagement.", "Documenter le mandat du directeur général après 2023."],
      "Conseil d'administration", "Faible", "P2", "R10", "11, 163 et 164",
      "Le fonctionnement des décisions de placement repose sur une déclaration du chef comptable, à confirmer par les pièces."),
    C(25, "G", "Le Grenier grandit très vite avec l'argent de l'épargne",
      "L'encours des prêts du Grenier passe de 7 900 000 FCFA fin 2021 à 1 044 302 107 FCFA au 31/08/2026, soit 132 fois plus. Le FPPN finance 87,66 % de cet encours (915 477 260 FCFA).",
      "Les prêts sont un service apprécié, mais l'argent prêté n'est plus disponible pour rembourser l'épargne.",
      "Les prêts sont courts (10 mois en médiane) et rapportent (taux médian de 18,00 %). Le risque est la vitesse de croissance, financée par une épargne qui n'est déjà couverte qu'à 16,28 %.",
      "Le Grenier vous rend service, mais sa croissance ne doit pas mettre en danger le remboursement de votre épargne.",
      ["Fixer par décision un plafond de financement du Grenier par le FPPN.", "Suivre chaque mois les retards de remboursement."],
      "Conseil d'administration, Grenier", "Moyenne", "P3", "R7", "102 à 108"),
    C(26, "H", "Des données essentielles manquent pour finir les calculs",
      "Le FPPN n'a pas transmis de base des prestations avec les dates (soins, dépôt, paiement), ni les dates de naissance des policiers, ni le fichier des ayants droit. Les fichiers maladie et décès de décembre 2024 sont vides.",
      "Sans ces données, il est impossible de calculer les provisions dossier par dossier, la mortalité par âge et l'engagement envers les retraités.",
      "Ces manques expliquent pourquoi deux réponses sur quatre sont seulement « partiellement démontrées ». Ils ne changent pas le sens des conclusions, mais ils en limitent la précision.",
      "Des fichiers complets et à jour sont la condition pour bien protéger vos droits.",
      ["Transmettre les données demandées (liste des documents demandés).", "Enregistrer dans le logiciel la date des soins et la date de dépôt de chaque facture."],
      "Direction générale, service informatique", "Moyenne", "P1", "R9", "54 et 241"),
    C(27, "H", "Les fichiers de cotisations ne sont pas rapprochés",
      "Le fichier envoyé chaque mois à la Solde n'est pas comparé au fichier retour, et aucun fichier envoyé n'a été transmis sur la période de 80 mois.",
      "Ce rapprochement est le seul moyen de savoir si chaque cotisation demandée a bien été prélevée et reversée.",
      "Sans ce contrôle, une cotisation non prélevée ou un écart de montant peut passer inaperçu pendant des mois.",
      "Votre cotisation doit être vérifiée chaque mois entre ce qui est demandé et ce qui est reçu.",
      ["Rapprocher chaque mois le fichier envoyé et le fichier retour, avec un état des rejets signé."],
      "Responsable des cotisations", "Moyenne", "P2", "", "35 à 38 et 51"),
    C(28, "H", "Deux versions des comptes 2023 existent",
      "Les états financiers concordent avec les balances 5 années sur 7. En 2023, les états financiers scannés portent un excédent de 928 699 102 FCFA, contre 930 629 102 FCFA dans la balance et le rapport de gestion corrigé.",
      "Les comptes approuvés par l'Assemblée générale doivent être uniques.",
      "L'écart est faible. Il s'agit sans doute d'une version antérieure, mais la version approuvée doit être confirmée.",
      "Les comptes présentés aux adhérents doivent être une seule et même version.",
      ["Transmettre la version des comptes 2023 approuvée par l'Assemblée générale."],
      "Service comptable", "Faible", "P3", "", "56"),
]
assert [k["num"] for k in K] == list(range(1, 29))
BOULE = {"P1": "🔴", "P2": "🟠", "P3": "🟡", "P4": "🟢"}

# 8. matrice complexité × priorité
fig, ax = plt.subplots(figsize=(8, 3.8))
PR, CX = ["P1", "P2", "P3", "P4"], ["Faible", "Moyenne", "Élevée"]
for i, cx in enumerate(CX):
    for j, pr in enumerate(PR):
        c = {"P1": ROUGE_R, "P2": ORANGE_R, "P3": JAUNE_R, "P4": VERT_R}[pr]
        ax.add_patch(Rectangle((j, i), 1, 1, fc=c, ec="white", lw=2, alpha=1.0 if cx == "Faible" else 0.75 if cx == "Moyenne" else 0.5))
        nums = [str(k["num"]) for k in K if k["cplx"] == cx and k["prio"] == pr]
        lignes_ = [", ".join(nums[a:a + 5]) for a in range(0, len(nums), 5)]
        ax.text(j + 0.5, i + 0.5, "\n".join(lignes_), ha="center", va="center", fontsize=10, fontweight="bold")
ax.set_xlim(0, 4)
ax.set_ylim(0, 3)
ax.set_xticks([0.5, 1.5, 2.5, 3.5], ["P1\n3 mois", "P2\n6 mois", "P3\n6 à 18 mois", "P4\nau-delà"])
ax.set_yticks([0.5, 1.5, 2.5], ["Faible", "Moyenne", "Élevée"])
ax.set_xlabel("Priorité")
ax.set_ylabel("Complexité")
ax.text(0.05, -0.75, "En haut à gauche : actions urgentes et simples, à lancer tout de suite.", fontsize=8, transform=ax.transData)
for s in ax.spines.values():
    s.set_visible(False)
F_CP = enregistrer_fig("matrice_complexite_priorite.png")


# ================================================================ DOCUMENT
JAUNE = []
r = RapportAdenka(titre="Résumé exécutif du volet actuariel", client="FPPN", type_livrable="Rapport provisoire", date="Octobre 2026")
r.sigles({"AIRMS": "Agence ivoirienne de régulation de la mutualité sociale", "BCEAO": "Banque centrale des États de l'Afrique de l'Ouest",
          "CGRAE": "Caisse générale de retraite des agents de l'État", "CMP": "Clinique médicale du policier",
          "EFC": "Entraide fin de carrière (épargne versée au départ à la retraite)", "FCP": "Fonds commun de placement",
          "FPPN": "Fonds de prévoyance de la Police nationale", "PEP": "Plan Épargne Plus (épargne volontaire)",
          "UEMOA": "Union économique et monétaire ouest-africaine"})
r.abreviations({"art.": "Article", "FCFA": "Franc de la Communauté financière africaine", "n°": "Numéro", "P1 à P4": "Niveaux de priorité (voir l'annexe 2)"})


def cellule(c, texte, taille=9.5, fond=None, gras_tout=False):
    if fond:
        AD._fond(c, fond)
    p = c.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.line_spacing = 1.05
    for kind, txt in AD._tokens(texte):
        if kind == "f":
            continue
        surl = False
        for morceau in txt.split("[["):
            if "]]" in morceau:
                a, b = morceau.split("]]", 1)
                for t_, s_ in ((a, True), (b, False)):
                    if t_:
                        run = p.add_run(t_)
                        AD._police_run(run, taille, gras_tout or kind == "b" or None, italique=(kind == "i") or None)
                        if s_:
                            JAUNE.append(t_)
                surl = True
            elif morceau:
                run = p.add_run(morceau)
                AD._police_run(run, taille, gras_tout or kind == "b" or None, italique=(kind == "i") or None)
        del surl


def table_brute(lignes, largeurs, fonds=None):
    t = r.doc.add_table(rows=len(lignes), cols=len(largeurs))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    bords = OxmlElement("w:tblBorders")
    for cote in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement("w:" + cote)
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:color"), "BBBCBC")
        bords.append(b)
    tblPr = t._tbl.tblPr
    suivant = next((tblPr.find(qn(tag)) for tag in ("w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook") if tblPr.find(qn(tag)) is not None), None)
    if suivant is not None:
        suivant.addprevious(bords)
    else:
        tblPr.append(bords)
    for i, ligne in enumerate(lignes):
        for j, val in enumerate(ligne):
            c = t.rows[i].cells[j]
            c.width = Cm(largeurs[j])
            f = (fonds or {}).get((i, j))
            cellule(c, val, fond=f, gras_tout=(j == 0 and len(largeurs) > 1))
    r.doc.add_paragraph()
    return t


def encadre(texte, fond="FBE5C8"):
    table_brute([[texte]], [18.4], {(0, 0): fond})


def fiche(k):
    r.h3(f"Constat {k['num']} : {k['titre']}")
    act = "\n".join(f"• {a}" for a in k["actions"])
    lignes = [["Ce qui a été constaté", k["constat"]], ["Pourquoi c'est important", k["pourquoi"]], ["Commentaire", k["commentaire"]],
              ["Ce que cela veut dire pour vous", "**" + k["vous"] + "**"], ["Plan d'actions", act],
              ["Classement", f"Complexité : **{k['cplx']}**.   Priorité : {BOULE[k['prio']]} **{k['prio']}**.   Responsable : {k['qui']}."]]
    if k["confirmer"]:
        lignes.append(["À confirmer", k["confirmer"]])
    lignes.append(["Pour aller plus loin", f"Rapport du 23/10/2026, diapositives {k['diapos']}" + (f". Risque {k['risque']} de la cartographie." if k["risque"] else ".")])
    fonds = {(i, 0): "F2F2F2" for i in range(len(lignes))}
    fonds[(3, 1)] = "FBE5C8"
    return table_brute(lignes, [4.4, 14.0], fonds)


# ---------------------------------------------------------------- synthèse
r.synthese("L'ESSENTIEL")
r.para("**Le FPPN encaisse plus qu'il ne dépense, mais il ne peut pas encore démontrer qu'il pourra rembourser à temps l'épargne de tous les policiers.** "
       "Le FPPN a dégagé un excédent chaque année de 2020 à 2025. Mais l'épargne des policiers, environ 15,7 milliards de FCFA (15 743 516 391 FCFA) au 31/08/2026, "
       "a surtout financé des bâtiments, des terrains et des prêts, qui ne peuvent pas être transformés vite en argent. Les intérêts promis aux épargnants ne sont pas comptabilisés, "
       "et le FPPN ne respecte pas plusieurs règles de solidité fixées par la réglementation.")
r.para("**La mission devait répondre à quatre questions.** Les réponses sont données sur la base des documents reçus du FPPN jusqu'au 23/10/2026.", gras_auto=False)
r.tableau(["Question", "Réponse", "En clair"],
          [["1. Combien le FPPN doit-il à ses adhérents ?", "Partiellement démontré", "Environ 19,5 milliards de FCFA (19 518 949 298 FCFA), dont 97,46 % pour l'épargne et ses intérêts. Le détail par policier manque."],
           ["2. Les sommes mises de côté suffisent-elles ?", "Partiellement démontré", "Oui en total, mais elles sont mal réparties et ne sont pas calculées. Rien n'est prévu pour les intérêts de l'épargne."],
           ["3. L'argent sera-t-il disponible au bon moment ?", "Insuffisant", "Non : l'argent disponible dans l'année couvre 34,42 % de ce qu'il faudra payer dans l'année."],
           ["4. Les placements sont-ils adaptés ?", "Insuffisant", "Non : seulement 5,53 % de l'épargne est réellement placée, sans politique de placement votée."]],
          titre="Réponses aux quatre questions de la mission", source=RAP + ", diapositives 4 à 8.", largeurs_cm=[5.4, 3.6, 9.4])
r.para("**Les cinq constats les plus graves.**", gras_auto=False)
r.puces(["Les intérêts promis aux épargnants, environ 3,3 milliards de FCFA (3 278 913 238 FCFA), ne sont pas comptabilisés (constat 6).",
         "L'argent disponible dans l'année ne couvre qu'un tiers des paiements de l'année, épargne comprise (constat 11).",
         "14 des 36 règles de solidité de la réglementation ne sont pas respectées, ce qui expose le FPPN à un retrait d'agrément (constat 14).",
         "Les provisions maladie et décès, environ 4,8 milliards de FCFA, ne sont pas calculées et dépassent largement le besoin réel (constat 1).",
         "La garantie maladie perd de l'argent depuis 2025 : 701 679 552 FCFA de déficit sur les huit premiers mois de 2026 (constat 19)."])
r.para("**Les actions à lancer dans les trois mois (priorité P1).**", gras_auto=False)
r.puces(["Fixer par délibération le taux d'intérêt de l'épargne, puis comptabiliser les intérêts dus (constats 6 et 7).",
         "Calculer chaque année les provisions et faire décider le reclassement de l'excédent (constat 1).",
         "Constituer une réserve de trésorerie et établir l'échéancier des sorties d'épargne (constat 11).",
         "Présenter à l'AIRMS un plan de mise en conformité chiffré (constat 14).",
         "Transmettre les données qui manquent aux calculs (constat 26)."])
r.para("**Le document présente ensuite** les travaux réalisés, les 28 constats avec leur plan d'actions, la cartographie des risques et les prochaines étapes. "
       "Chaque action est classée par complexité (faible, moyenne, élevée) et par priorité (P1 à P4).", gras_auto=False)

# ---------------------------------------------------------------- partie 1
r.partie(1, "Les travaux réalisés", introduction="Cette partie décrit ce qui a été fait, avec quels documents et quels outils.")
r.h2("1.1. LE CADRE DE LA MISSION")
r.para("**Le volet actuariel et statistique vérifie si le FPPN pourra tenir ses promesses envers les policiers, aujourd'hui et demain.** "
       "L'audit comptable vérifie que les comptes sont justes. Le volet actuariel regarde vers l'avenir : ce que le FPPN devra payer, l'argent qu'il a mis de côté, "
       "et la disponibilité de cet argent au bon moment. La période étudiée va du 01/01/2020 au 31/08/2026.")
encadre("**Le mot « actuariel ».** L'actuariat est le métier qui mesure aujourd'hui des sommes qui seront payées plus tard, par exemple l'épargne remboursée au départ à la retraite ou les soins des retraités.")
r.h2("1.2. LES ÉTAPES DES TRAVAUX")
r.figure(F_FRISE, titre="Les sept étapes des travaux", source="Adenka.", largeur_cm=17)
r.tableau(["Étape", "Ce qui a été fait", "En chiffres"],
          [["Lancement", "Réunions de lancement des 13 et 14/10/2026 et liste des documents demandés au FPPN", "88 documents demandés, sans doublon"],
           ["Collecte", "Sept envois du FPPN du 15/10 au 23/10/2026, vérifiés un par un", "6 reçus, 39 en partie, 43 non reçus"],
           ["Entretiens", "Service des prestations, responsable des cotisations, chef comptable (16 et 19/10/2026)", "4 entretiens avec compte rendu"],
           ["Contrôles des données", "Vrai format des fichiers, lignes de total, montants doublés, fichiers vides, comparaison des fichiers envoyés et reçus", "4 454 416 lignes de précomptes, 506 feuilles, 7 balances comptables"],
           ["Analyses", "Effectifs, cotisations, soins, décès, épargne, prêts, retraités, provisions, liquidité, placements, 36 règles de solidité, projection", "14 thèmes analysés"],
           ["Vérification", "Chaque chiffre recalculé une seconde fois depuis les fichiers d'origine", "2 885 chiffres contrôlés, 0 écart"],
           ["Constats et actions", "Constats classés, risques notés, plan d'actions", "28 constats, 10 risques"]],
          titre="Ce qui a été fait", source=RAP + " ; liste des documents demandés au 26/10/2026.", largeurs_cm=[3.4, 9.0, 6.0])
r.h2("1.3. LES OUTILS UTILISÉS")
r.puces(["**Des programmes informatiques** pour lire et contrôler plus de quatre millions de lignes de fichiers, y compris des documents scannés.",
         "**Une empreinte numérique de chaque fichier reçu**, qui prouve que les fichiers analysés sont exactement ceux transmis par le FPPN.",
         "**Un classeur Excel de 99 onglets**, où chaque calcul est refait par une formule visible et expliquée.",
         "**Un contrôle automatique** qui retrouve chacun des 2 885 chiffres du rapport dans les calculs."])

# ---------------------------------------------------------------- partie 2
r.partie(2, "Les chiffres clés", introduction="Cette partie montre en images les grandes évolutions du FPPN de 2020 au 31/08/2026.")
r.h2("2.1. L'ÉPARGNE DES POLICIERS ET L'ARGENT MIS DE CÔTÉ")
r.para("**L'épargne due augmente beaucoup plus vite que l'argent réservé pour la rembourser.** L'épargne passe d'environ 10,8 milliards de FCFA en 2020 à environ 15,7 milliards de FCFA au 31/08/2026. "
       "L'argent réservé n'en couvre que 16,28 %.")
r.figure(F_EPARGNE, titre="Épargne due et argent réservé pour la rembourser (milliards de FCFA)", source=RAP + ", diapositives 13 et 90.", largeur_cm=16)
r.h2("2.2. LA CONSOMMATION DES COTISATIONS PAR LES SOINS ET LES DÉCÈS")
r.para("**Les soins consomment presque toutes les cotisations maladie en 2026, alors que le décès reste très excédentaire.** "
       "La bande grise montre la fourchette de 70 % à 80 % fixée par l'AIRMS pour l'ensemble des garanties.")
r.figure(F_SIN, titre="Part des cotisations consommée par les prestations (%)", source=RAP + ", diapositives 13 et 75.", largeur_cm=16)
r.h2("2.3. LES FONDS PROPRES")
r.para("**Les excédents de 2020 à 2025 réduisent les pertes passées sans les effacer, et 2026 marque un recul.**")
r.figure(F_FP, titre="Fonds propres et résultat de l'année (milliards de FCFA)", source=RAP + ", diapositive 161.", largeur_cm=16)
r.h2("2.4. LES SOMMES MISES DE CÔTÉ FACE AUX BESOINS")
r.para("**Trop d'argent est mis de côté pour la maladie et le décès, et rien pour les intérêts de l'épargne et les prêts en retard.**")
r.figure(F_PROV, titre="Montants au bilan et besoins mesurés (milliards de FCFA)", source=RAP + ", diapositives 144 et 145.", largeur_cm=16)
r.h2("2.5. LE MANQUE D'ARGENT DISPONIBLE")
r.para("**Si l'épargne continue de sortir au rythme observé, il manquerait environ 12,6 milliards de FCFA en cinq ans**, sans compter les nouvelles cotisations.")
r.figure(F_LIQ, titre="Manque d'argent disponible cumulé, année par année (milliards de FCFA)", source=RAP + ", diapositive 155.", largeur_cm=15)

# ---------------------------------------------------------------- partie 3
r.partie(3, "Les constats et le plan d'actions", introduction="Chaque constat est présenté dans une fiche : ce qui a été constaté, pourquoi c'est important, le commentaire d'Adenka, ce que cela veut dire pour les policiers, et les actions à mener.")
encadre("**Comment lire les fiches.** La complexité dit combien l'action est difficile à mener : faible (une décision ou un document suffit), moyenne (un travail interne de quelques mois), élevée (il faut un tiers ou un financement). "
        "La priorité dit quand agir : 🔴 P1 dans les 3 mois, 🟠 P2 dans les 6 mois, 🟡 P3 de 6 à 18 mois, 🟢 P4 au-delà de 18 mois.", fond="F2F2F2")
n_h2 = 0
for code, nom in TH.items():
    n_h2 += 1
    r.h2(f"3.{n_h2}. {nom.upper()}")
    for k in [k for k in K if k["theme"] == code]:
        fiche(k)

# ---------------------------------------------------------------- partie 4
r.partie(4, "La cartographie des risques", introduction="Cette partie situe les risques du FPPN selon leur probabilité et leur gravité, puis place les actions selon leur complexité et leur priorité.")
r.h2("4.1. LA MATRICE DES RISQUES")
r.para("**Sept risques sur dix sont dans la zone rouge, qui demande un traitement immédiat.** Chaque risque est noté de 1 à 5 en probabilité (5 : déjà constaté dans les comptes) "
       "et de 1 à 5 en gravité selon son impact en FCFA. Le score est le produit des deux notes.")
r.figure(F_RISQ, titre="Matrice des risques : probabilité et gravité", source=RAP + ", diapositives 23 et 24.", largeur_cm=13)
r.tableau(["N°", "Risque en clair", "Probabilité", "Gravité", "Score", "Constat"],
          [["R1", "Ne pas pouvoir payer les sorties d'épargne à temps", "4", "5 : manque de 12 596 052 822 FCFA à cinq ans", "🔴 20", "11"],
           ["R2", "Perdre l'agrément ou être sanctionné par l'AIRMS", "4", "5 : arrêt de l'activité", "🔴 20", "14"],
           ["R3", "Sous-estimer la dette d'épargne", "5", "4 : 3 278 913 238 FCFA d'intérêts non comptés", "🔴 20", "6 et 7"],
           ["R4", "Garder des provisions sans lien avec les risques", "5", "4 : 4 809 716 898 FCFA au-dessus du besoin", "🔴 20", "1"],
           ["R5", "Voir la garantie maladie s'enfoncer dans le déficit", "5", "3 : 701 679 552 FCFA de déficit sur huit mois de 2026", "🔴 15", "19"],
           ["R6", "N'avoir aucun coussin pour un coup dur", "5", "3 : fonds propres de -879 498 700 FCFA", "🔴 15", "15"],
           ["R9", "Ne pas pouvoir mesurer les engagements faute de données", "5", "3 : provisions, mortalité et retraités non calculables", "🔴 15", "26"],
           ["R10", "Engager l'argent sans décision des organes", "4", "3 : fonds commun de placement de 500 000 000 FCFA en 2025", "🟡 12", "24"],
           ["R8", "Ne pas gagner assez pour payer les intérêts promis", "5", "2 : 470 608 168 FCFA d'écart en 2025", "🟡 10", "23"],
           ["R7", "Perdre de l'argent sur les prêts du Grenier", "5", "1 : 14 016 392 FCFA", "🟢 5", "4 et 25"]],
          titre="Les dix risques, du plus grave au moins grave", source=RAP + ", diapositive 24.", largeurs_cm=[1.1, 5.3, 2.3, 6.0, 1.7, 2.0], taille=8.5)
r.h2("4.2. LES ACTIONS SELON LEUR COMPLEXITÉ ET LEUR PRIORITÉ")
r.para("**Six actions sont à lancer dans les trois mois, dont une seule est simple.** Le taux d'intérêt de l'épargne (constat 7) ne demande qu'une délibération. "
       "Les autres actions P1 demandent un travail interne ou l'accord d'un tiers : il faut les lancer tout de suite pour qu'elles aboutissent à temps.")
r.figure(F_CP, titre="Les 28 constats selon la complexité de l'action et sa priorité (numéros des constats)", source="Adenka.", largeur_cm=15)

# ---------------------------------------------------------------- partie 5
r.partie(5, "Récapitulatif, limites et suite", introduction="Cette partie récapitule tous les constats, dit ce qui n'a pas encore pu être vérifié et présente les prochaines étapes.")
r.h2("5.1. TABLEAU RÉCAPITULATIF DES CONSTATS ET DES ACTIONS")
ordre = sorted(K, key=lambda k: (k["prio"], ["Faible", "Moyenne", "Élevée"].index(k["cplx"]), k["num"]))
r.tableau(["N°", "Constat", "Action principale", "Complexité", "Priorité"],
          [[str(k["num"]), k["titre"], k["actions"][0], k["cplx"], f"{BOULE[k['prio']]} {k['prio']}"] for k in ordre],
          titre="Les 28 constats et leur action principale, par priorité", source=RAP + " ; analyses Adenka.", largeurs_cm=[1.0, 5.6, 7.8, 2.0, 2.0], taille=8)
r.h2("5.2. CE QUI N'A PAS ENCORE PU ÊTRE VÉRIFIÉ, ET POURQUOI")
r.para("**Les limites ci-dessous changent la précision des montants, pas le sens des conclusions.**")
r.tableau(["Ce qui manque", "Ce que cela empêche de vérifier", "Document attendu du FPPN"],
          [["Dates des soins, du dépôt et du paiement de chaque facture", "Les provisions maladie dossier par dossier, et la comparaison des provisions passées avec les paiements réels", "Base des prestations datée"],
           ["Dates de naissance des policiers et des retraités", "La mortalité par âge, l'engagement santé des retraités, la projection par âge", "Fichier des mutualistes, fichiers de la CGRAE"],
           ["Comptes individuels d'épargne et taux officiel", "L'épargne due à chaque policier et le montant exact des intérêts", "Relevés individuels, délibération sur le taux"],
           ["Valeur de la clinique et des terrains", "La couverture réelle des dettes par les biens du FPPN", "Rapports d'évaluation"],
           ["Inventaire signé des placements", "L'adossement année par année des placements à l'épargne", "Inventaire et échéancier des placements"],
           ["Déclarations non documentées", "Certains faits de gouvernance, surlignés en jaune dans ce document", "Délibérations et procès-verbaux"]],
          titre="Les vérifications restant à faire", source=RAP + ", diapositives 241 et 242 ; liste des documents demandés au 26/10/2026.", largeurs_cm=[5.0, 8.0, 5.4])
r.h2("5.3. LES PROCHAINES ÉTAPES")
r.tableau(["Étape", "Qui", "Quand"],
          [["Transmettre les données et documents encore attendus", "FPPN", "Dès que possible"],
           ["Calculer les provisions dossier par dossier et l'épargne par policier, puis projeter l'équilibre des garanties", "Adenka", "À réception des données"],
           ["Restitution du rapport provisoire", "Adenka et FPPN", "23 et 24/11/2026"],
           ["Intégration des commentaires du FPPN", "Adenka", "Du 25 au 27/11/2026"],
           ["Remise du rapport définitif", "Adenka", "30/11/2026"]],
          titre="Prochaines étapes", source="Adenka.", largeurs_cm=[10.4, 4.0, 4.0])

# ---------------------------------------------------------------- annexes
r.annexe(1, "Lexique")
r.tableau(["Mot", "Ce qu'il veut dire"],
          [["Actuariat", "Métier qui mesure aujourd'hui des sommes qui seront payées plus tard."],
           ["Adossement", "Le fait d'avoir l'argent disponible au moment où il faut payer."],
           ["Agrément", "Autorisation de l'AIRMS sans laquelle le FPPN ne peut pas exercer."],
           ["Ayant droit", "Conjoint ou enfant d'un policier, couvert par les garanties du FPPN."],
           ["Bilan", "Photographie, à une date, de ce que le FPPN possède et de ce qu'il doit."],
           ["Cotisation", "Somme prélevée chaque mois sur la solde ou la pension."],
           ["Dépréciation", "Perte probable comptée sur un prêt qui risque de ne pas être remboursé."],
           ["Engagement", "Somme que le FPPN devra payer à ses adhérents, maintenant ou plus tard."],
           ["Fonds propres", "Réserves du FPPN moins ses pertes passées : le coussin pour les coups durs."],
           ["Liquidité", "Capacité à payer à la date prévue avec l'argent disponible."],
           ["Placement", "Argent confié à une banque ou à un fonds pour qu'il rapporte des intérêts."],
           ["Précompte", "Prélèvement de la cotisation directement sur la solde (par la Solde) ou sur la pension (par la CGRAE)."],
           ["Provision", "Argent mis de côté à la clôture pour payer des soins ou des décès déjà survenus mais pas encore payés."],
           ["Sinistralité", "Part des cotisations consommée par les prestations (soins, allocations)."],
           ["Taux servi", "Taux d'intérêt promis aux épargnants."]],
          titre="Les mots techniques utilisés", source="Adenka.", largeurs_cm=[4.0, 14.4])
r.annexe(2, "Règles de classement")
r.tableau(["Classement", "Niveau", "Signification"],
          [["Complexité", "Faible", "Une décision des organes ou un document suffit."],
           ["Complexité", "Moyenne", "Un travail interne de quelques mois est nécessaire."],
           ["Complexité", "Élevée", "L'action dépend d'un tiers (État, AIRMS, expert, actuaire) ou d'un financement."],
           ["Priorité", "🔴 P1", "Risque critique : à lancer dans les 3 mois."],
           ["Priorité", "🟠 P2", "Risque élevé : à mener dans les 6 mois."],
           ["Priorité", "🟡 P3", "Risque moyen : à mener de 6 à 18 mois."],
           ["Priorité", "🟢 P4", "Amélioration : au-delà de 18 mois."],
           ["Gravité d'un risque", "1 à 5", "Moins de 50 000 000 FCFA (1), jusqu'à 500 000 000 FCFA (2), jusqu'à 2 000 000 000 FCFA (3), jusqu'à 5 000 000 000 FCFA (4), au-delà ou arrêt de l'activité (5)."],
           ["Probabilité d'un risque", "1 à 5", "De très peu probable (1) à très probable, c'est-à-dire déjà constaté dans les comptes ou les fichiers (5)."],
           ["Score d'un risque", "1 à 25", "Probabilité multipliée par gravité. 🔴 15 à 25 : traitement immédiat. 🟡 8 à 12 : plan d'action. 🟢 1 à 6 : surveillance."]],
          titre="Règles de classement des actions et des risques", source="Adenka ; " + RAP + ", diapositive 23.", largeurs_cm=[3.8, 2.4, 12.2])
r.enregistrer(SORTIE)

# ---- retouches : pages de séparation (cadre élargi, sans habillage), surlignage jaune
doc = Document(SORTIE)
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
for dr in doc.element.body.iter(qn("w:drawing")):
    txt = "".join(t.text or "" for t in dr.iter(qn("w:t")))
    if txt.startswith("PARTIE") or txt.startswith("ANNEXE"):
        for e in list(dr.iter(WP + "extent")) + list(dr.iter(A + "ext")):
            if e.get("cx") == "1691640":
                e.set("cx", "4320000")
        for w_ in list(dr.iter(WP + "wrapSquare")):
            w_.addprevious(w_.makeelement(WP + "wrapNone", {}))
            w_.getparent().remove(w_)
for run in doc.element.body.iter(qn("w:r")):
    txt = "".join(t.text or "" for t in run.iter(qn("w:t")))
    if txt and any(txt == j for j in JAUNE):
        rpr = run.find(qn("w:rPr"))
        if rpr is None:
            rpr = run.makeelement(qn("w:rPr"), {})
            run.insert(0, rpr)
        h = rpr.makeelement(qn("w:highlight"), {qn("w:val"): "yellow"})
        apres = [rpr.find(qn("w:" + n_)) for n_ in ("u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath")]
        apres = [e for e in apres if e is not None]
        if apres:
            min(apres, key=list(rpr).index).addprevious(h)
        else:
            rpr.append(h)
doc.save(SORTIE)
print(SORTIE, "constats :", len(K), {p: sum(1 for k in K if k["prio"] == p) for p in ("P1", "P2", "P3", "P4")})
