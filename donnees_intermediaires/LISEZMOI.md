# Données intermédiaires FPPN reconstituées sur le PC (13/11/2026)

Ce dossier remplace le dossier de travail temporaire de la session cloud, qui a été effacé.
**Seul ce fichier est suivi par Git.** Tout le reste (bases parquet avec matricules et dates de naissance, JSON, classeurs) est exclu par `.gitignore`.

## Organisation

| Dossier | Contenu |
|---|---|
| `base/` | 9 bases parquet (précomptes, sources par mois, effectifs, rejets, balances comptables, Grenier) |
| `resultats/` | JSON et classeurs produits par les scripts amont, lus par `rapport_v12_fppn.py` |
| `pbc/` | Correspondance des numéros d'items PBC (`correspondance_pbc.json`, argument `PBCJ` du rapport) |
| `sources_externes/` | Tables de mortalité CIMA et série d'inflation de la Banque mondiale, fournies par l'utilisateur le 13/11/2026 |
| `controle/` | Copie du classeur du 11/11/2026 recalculée par Excel et résultat de la comparaison |
| `logs/` | Sortie écran de chaque script |

## Entrées

| Entrée | Chemin sur le PC |
|---|---|
| Extraction de la première dataroom (18/10/2026) | `C:\Dev\Missions\05-FPPN\05-Analyses\Extraction_FPPN_20261018_1939\donnees` (empreinte de `journal_fichiers.csv` : `6fd43e6c667350d8…`) |
| Lots des documents nouveaux (LOT1, LOT4) | `C:\Dev\Missions\05-FPPN\04-Documents reçus du client` (sous-dossiers `02-` à `08-`) |
| Grands livres (dataroom du 09/11/2026) | `D:\03-Adenka-Audit\02-Documents FPPN\Balance générale et Grand livre` |
| Tables CIMA | `Tables de mortalité CIMA.zip` (Téléchargements), fichiers renommés avec `_` à la place des espaces (`CIMA_F.xlsx`…), comme les attendent les scripts |
| Inflation | `API_FP.CPI.TOTL.ZG_DS2_fr_excel_v2_1775.xls` (Banque mondiale, mise à jour du 08/10/2026), série CSV jointe pour mémoire |
| IHPC ANStat, rendements obligataires | `scripts/assets/ihpc_anstat`, `scripts/assets/investing` (dans le dépôt) |

## Commandes exactes

Depuis la racine du dépôt (`06-Depot-Adenka`), Git Bash, Python 3.14.7 (pandas 3.1.1, pyarrow 26.0.1, python-calamine, openpyxl, statsmodels 0.15.0), `PYTHONIOENCODING=utf-8`.

```bash
DI=donnees_intermediaires; B=$DI/base; R=$DI/resultats; SE=$DI/sources_externes
EX="/c/Dev/Missions/05-FPPN/05-Analyses/Extraction_FPPN_20261018_1939/donnees"
LOT="/c/Dev/Missions/05-FPPN/04-Documents reçus du client"
G="/d/03-Adenka-Audit/02-Documents FPPN/Balance générale et Grand livre"

# 1. Bases parquet (environ 6 minutes)
python scripts/base_fppn.py "$EX" $B
python scripts/compta_fppn.py "$EX" $B
python scripts/grenier_fppn.py "$EX" $B

# 2. Correspondance PBC et placements lus dans les grands livres
python scripts/maj_pbc_23_fppn.py FPPN_PBC_complete_pointee_20261022.xlsx $DI/pbc/FPPN_PBC_complete_pointee_20261023_regen.xlsx $DI/pbc/correspondance_pbc.json
python scripts/placements_gl_fppn.py $R/placements_gl.json \
  "2021=$G/Nouveaux fichiers/GRAND LIVRE 2021.xlsx" \
  "2022=$G/2022/Nouveau/GRAND LIVRE DES COMPTES 2022.xlsx" \
  "2023=$G/2023/GRAND LIVRE DES COMPTES FPPN 2023.xlsx" \
  "2024=$G/Nouveaux fichiers/GRAND LIVRE FPPN 2024.xlsx" \
  "2025=$G/2025/GRAND LIVRE DES COMPTES FPPN 2025.xlsx" \
  "2026=$G/2026/GRAND LIVRE DES COMPTES FPPN DU 01-01-2026 AU 31-08-2026.xlsx"

# 3. Résultats amont
python scripts/analyses_fppn.py $B $R
python scripts/provisions_fppn.py $B $R
python scripts/ratios_airms_fppn.py $B $R/resultats.json $R
python scripts/qualite_bases_fppn.py $B $R/qualite.json
python scripts/verif_precomptes_fppn.py "$EX/tableaux" $B $R/verif_precomptes.json
python scripts/analyses_complements_fppn.py $B $R/resultats.json $SE/inflation/API_FP.CPI.TOTL.ZG_DS2_fr_excel_v2_1775.xls $SE/tables_cima "$EX" $R
python scripts/sensibilites_fppn.py $R/ratios.json $R/resultats.json scripts/assets/ihpc_anstat $R/sensibilites.xlsx scripts/assets/investing
python scripts/donnees_v9_fppn.py "$LOT" $B $R/resultats.json $R/provisions.json $SE/tables_cima scripts/assets/investing $R/donnees_v9.json $R/donnees_v9.xlsx
python scripts/donnees_v10_fppn.py "$LOT" "$LOT" $B $R/resultats.json $R/provisions.json $R/ratios.json $R/donnees_v9.json $R/donnees_v10.json $R/donnees_v10.xlsx
python scripts/donnees_v11_fppn.py $R/placements_gl.json $R/donnees_v9.json $R/ratios.json $R/resultats.json $R/donnees_v11.json
```

Non relancé volontairement : `sinistralite_gl_fppn.py` (il réécrit `scripts/assets/sinistralite_v15.json`, déjà dans le dépôt) et `analyses_sinistres_fppn.py` (hors de la chaîne du rapport, ses fichiers d'entrée préparés n'existent pas).

## Contrôle

Le classeur livré `rapport/FPPN_Calculs_rapport_20261111.xlsx` contient des formules sans valeur en cache. Une copie a été recalculée par Excel (`controle/FPPN_Calculs_rapport_20261111_recalcule.xlsx`, le fichier livré n'est pas modifié), puis comparée par :

```bash
python scripts/controle_donnees_intermediaires_fppn.py $DI/controle/FPPN_Calculs_rapport_20261111_recalcule.xlsx $R
```

Résultat (`controle/comparaison_brute.txt`) : 60 onglets comparés, 58 identiques à 100 % (nombres arrondis à l'unité). Les deux autres s'expliquent :
- `7_rendement_placements` : le classeur du 11/11 remplace les encours de clôture par les encours moyens lus dans les grands livres. Ces encours sont identiques, à l'unité près, à ceux de `resultats/donnees_v11.json` (2021 : 701 924 182 FCFA … 2025 : 763 752 135 FCFA).
- `10_agregats` : seuls les en-têtes d'années (2020 à 2026) diffèrent, écrits en texte dans le classeur livré.

Les provisions figées de 4 829 298 290 FCFA sont retrouvées (`complements.json`, `ratios.json`, `resultats.json`). Le besoin estimé de 19 581 392 FCFA n'est pas une donnée amont : il est calculé à l'étape du rapport (onglet `9_tracabilite`, lignes « SY besoin » et « CP besoin »).

## Empreintes des fichiers produits (SHA-256, 16 premiers caractères)

| Fichier | Octets | SHA-256 |
|---|---|---|
| `base/compta_balances.parquet` | 97274 | `a76a0c14ff659a0c` |
| `base/compta_controle.parquet` | 7937 | `6f35f05a54a2b593` |
| `base/effectifs_detail.parquet` | 2536623 | `85ad1642887eae4e` |
| `base/grenier_balances.parquet` | 896403 | `dc85ddd95bb7b24a` |
| `base/grenier_controle.parquet` | 9991 | `691be9519532192a` |
| `base/grenier_production.parquet` | 432503 | `b6795c5829a47a38` |
| `base/precomptes.parquet` | 37284396 | `ddcec14359e50a77` |
| `base/rejets.parquet` | 9014 | `6ec2e5603b38e93b` |
| `base/sources_mois.parquet` | 37468 | `8532ed272a98c268` |
| `pbc/FPPN_PBC_complete_pointee_20261023_regen.xlsx` | 110505 | `adb9e1f9354b30d0` |
| `pbc/correspondance_pbc.json` | 2118 | `f82cb392fb806ed8` |
| `resultats/FPPN_Calculs_complements.xlsx` | 11722 | `ea2812beaf94d444` |
| `resultats/FPPN_Calculs_provisions.xlsx` | 11356 | `f6de5fb76488074c` |
| `resultats/FPPN_Calculs_rapport.xlsx` | 30004 | `846bf7717e915ed4` |
| `resultats/FPPN_Calculs_ratios.xlsx` | 51735 | `06c5d35a3f49c732` |
| `resultats/complements.json` | 21043 | `adf369206b07f3f8` |
| `resultats/donnees_v10.json` | 33341 | `f7195283c9395c70` |
| `resultats/donnees_v10.xlsx` | 42759 | `1d02fb123670b7b3` |
| `resultats/donnees_v11.json` | 9208 | `c5563e34d832e12f` |
| `resultats/donnees_v9.json` | 180274 | `1a458955ad9cf4af` |
| `resultats/donnees_v9.xlsx` | 35842 | `88d28777f748fc3b` |
| `resultats/placements_gl.json` | 38869 | `cf88c0f5c5b9d032` |
| `resultats/provisions.json` | 13413 | `d281421a6af0e020` |
| `resultats/qualite.json` | 7168 | `bc551e225f298457` |
| `resultats/ratios.json` | 72136 | `78f99247a802f656` |
| `resultats/resultats.json` | 47846 | `faf7639404579df1` |
| `resultats/sensibilites.xlsx` | 112204 | `fb7bb09a49ad650e` |
| `resultats/sensibilites_carte.json` | 15605 | `fed901ec06c0d2df` |
| `resultats/verif_precomptes.json` | 143655 | `09c445ef76fced02` |

## Test des intérêts de l'épargne (13/11/2026)

Dossier `test_interets/`. Les fichiers `*_nominatif.parquet` contiennent des noms et des matricules : ils ne doivent pas être diffusés.

```bash
T=$DI/test_interets; mkdir -p $T
python scripts/gl_epargne_fppn.py "D:/03-Adenka-Audit/02-Documents FPPN/Balance générale et Grand livre" $T/gl_epargne.parquet
python scripts/test_interets_epargne_fppn.py $B $T/gl_epargne.parquet $T
python scripts/regle_interets_pep_fppn.py $T/gl_epargne.parquet $T/detail_nominatif.parquet $T
```

Résultats (`regle_interets_pep.json`) :
- **Règle appliquée par le FPPN au PEP : reçu = total versé × 1,035, une seule fois, quelle que soit la durée.** Sur les 6 482 remboursements PEP de 09/2020 à 08/2026, 4 987 (76,9 %) suivent cette règle à 1 FCFA près (2026 : 2 528 sur 2 633). Les acomptes et rachats partiels sont payés sans intérêt (93,9 % et 96,5 %).
- Intérêts ainsi payés : 198 040 675 FCFA, prélevés sur le compte d'épargne 41195100, aucun compte de charges ne porte d'intérêts sur l'épargne des policiers (recherche dans les grands livres 2023 au 31/08/2026).
- Manuel : « le PEP est rémunéré au taux de 3,5% tous les ans » (« MANUEL DE PRODUIT D'EPARGNE FPPN VF VF.docx », partie PEP). Sur 1 190 policiers sortis payés selon la règle (historique complet, durée médiane 3,20 ans) : 19 715 675 FCFA payés en plus du versé, contre 43 468 398 FCFA d'intérêts promis (intérêt simple de chaque précompte jusqu'à la sortie). Manque : 23 752 723 FCFA.
- 277 policiers : le FPPN rembourse sur une base inférieure au versé de 1 à 3 mois de précomptes (4 576 001 FCFA versés non repris dans le calcul), à rapprocher des trop-perçus remboursés ensuite.
- Non concluants : 54 dossiers payés au-delà de 1,035 (sorties plus anciennes, durée médiane 9,77 ans), 86 dossiers avec plus de 3 mois d'écart (paiements partiels ou homonymes), dossiers non rapprochés par le nom (1 521 écritures PEP, 1 668 741 919 FCFA).
- EFC : non concluant. Les précomptes EFC ne sont connus que depuis 01/2023 et les règlements EFC sont des montants ronds (1 800 000, 1 900 000, 2 000 000 FCFA…) qui ne suivent ni le versé ni le versé × 1,055. Il faut le règlement de l'Entraide de Fin de Carrière pour le tester.

## Vérification de la revue indépendante (13/11/2026)

Dossier `revue/` : texte du rapport du 11/11/2026 par diapositive, texte de la Décision n° 05/2022, recalculs (`verif_revue.json`).

```bash
python scripts/verif_revue_fppn.py $B $DI/controle/FPPN_Calculs_rapport_20261111_recalcule.xlsx $DI/revue/rapport_20261111_texte.json "D:/03-Adenka-Audit/02-Documents FPPN/Balance générale et Grand livre" $DI/revue/verif_revue.json
python scripts/registre_revue_fppn.py revue_independante/FPPN_Verification_revue_independante_20261113.xlsx
```

Registre : 85 lignes couvrant les 155 constats des cinq réviseurs (doublons regroupés). 67 confirmés, 7 confirmés en partie, 11 non revérifiés faute de pièce. Décisions proposées : 66 à corriger, 7 à ajouter, 12 à demander au FPPN.

## Ce qui reste pour régénérer le rapport lui-même

`rapport_v12_fppn.py` attend aussi les modules des skills Adenka (`pptx_tools`, `adenka_docx`, gabarit `Timesaver_Adenka.pptx`) et le dossier d'images de code (`images_code_fppn.py`, qui demande `pygments`). Ils ne sont pas installés sur le PC à ce jour.
