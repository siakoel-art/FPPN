# Guide pas à pas : extraire et m'envoyer les données FPPN

Durée : 10 minutes de préparation, puis quelques minutes à une demi-heure d'exécution.

## Ce que fait le script

- **Excel (.xlsx, vrais .xls) :** chaque feuille devient un fichier CSV brut, avec toutes les lignes et toutes les cellules. Rien n'est trié, filtré ni retiré.
- **CSV, TXT, .doc, exports « page web » du Grenier :** copiés tels quels, octet pour octet.
- **Fichiers que le script ne sait pas lire :** copiés tels quels, pour que je tente de les réparer.
- **Laissés de côté :**
  - les scans de pièces individuelles (pièces CGRAE, contrats MEDOC, bordereaux, transmissions, résiliations, précomptes en double ou indus) ;
  - les PDF qui doublonnent un fichier Excel, sauf 5 PDF utiles (balances et grand livre dont l'Excel est endommagé, rapports d'audit informatique) ;
  - les doublons exacts.

  Ils sont tous listés dans le journal, avec leur taille.
- **Journal de contrôle :** pour chaque fichier, une empreinte SHA-256 (sa « signature » numérique), l'action réalisée, le nombre de lignes et de cellules, et la somme des nombres. Je recalcule ces chiffres à réception pour prouver que rien n'est perdu.
- **Archive :** le tout est compressé dans une archive ZIP, découpée en parties de 20 Mo à m'envoyer.

Le script ne modifie, ne déplace ni ne supprime aucun fichier du dossier source.

Une limite à connaître : pour une cellule Excel qui contient une formule, le script garde le résultat affiché, pas la formule. Les fichiers du FPPN sont des exports de logiciels, donc ce cas est rare.

## Étape 1 : placer le script

1. Copier `extraction_fppn.py` dans `C:\Dev\Missions\05-FPPN\05-Analyses\scripts` (le dossier du script d'inventaire).
2. Dans VS Code : `Fichier > Ouvrir le dossier` puis choisir `05-Analyses\scripts`.
3. Ouvrir `extraction_fppn.py`.

**Contrôle :** le fichier s'affiche avec des couleurs (syntaxe Python), et en bas à droite apparaît l'interpréteur Python (par exemple « Python 3.11 »).

## Étape 2 : vérifier les paramètres

En haut du fichier, partie **PARAMÈTRES** :

| Paramètre | Valeur attendue |
|---|---|
| `DOSSIER_SOURCE` | Le même dossier que pour l'inventaire : `C:\Dev\Missions\05-FPPN\04-Documents reçus du client\01-download` |
| `DOSSIER_SORTIE` | `C:\Dev\Missions\05-FPPN\05-Analyses` |
| `TAILLE_PARTIE_MO` | `20` |
| `SIMULATION` | `True` pour le premier lancement (étape 3) |

Mettre `SIMULATION = True`, puis enregistrer (`Ctrl + S`).

## Étape 3 : simulation (rien n'est écrit)

1. Cliquer sur le triangle « Run Python File » en haut à droite.
2. Au premier lancement, le script installe deux bibliothèques (`python-calamine` et `openpyxl`). Cela prend moins d'une minute.
3. Le terminal affiche `[1/4280] ...` puis une synthèse.

**Contrôle :**
- La ligne `Fichiers trouvés :` donne environ 4 280.
- Le tableau de synthèse donne, par action, le nombre de fichiers et leur volume.
- La dernière ligne indique « SIMULATION : rien n'a été écrit ».

**Copie-moi la synthèse affichée** (du titre « Extraction FPPN » jusqu'à « Volume écrit »). Je vérifie que tout est conforme avant l'extraction réelle.

## Étape 4 : extraction réelle

1. Remettre `SIMULATION = False`, puis enregistrer (`Ctrl + S`).
2. Relancer avec « Run Python File ».
3. À la fin, le terminal affiche :
   - « Terminé en … minutes » ;
   - la taille de l'archive et le nombre de parties ;
   - le chemin du dossier `A_ENVOYER`.

**Si le PC se met en veille, le script s'arrête.** Désactiver la mise en veille pendant l'exécution.

**Contrôle :** dans `05-Analyses\Extraction_FPPN_AAAAMMJJ_HHMM\` on trouve :

| Élément | Contenu |
|---|---|
| `donnees\tableaux\` | Les CSV (un par feuille Excel) |
| `donnees\bruts\` | Les fichiers copiés tels quels |
| `donnees\journal_fichiers.csv` | Une ligne par fichier source |
| `donnees\journal_feuilles.csv` | Une ligne par feuille Excel convertie |
| `donnees\synthese.txt` | La synthèse affichée à l'écran |
| `A_ENVOYER\` | Les parties `.zip.001`, `.zip.002`… et `LISEZMOI_parties.txt` |

## Étape 5 : m'envoyer les fichiers

1. Ouvrir le dossier `A_ENVOYER`.
2. M'envoyer **d'abord** `LISEZMOI_parties.txt`. Il contient la liste des parties et leurs empreintes.
3. M'envoyer ensuite les parties, dans l'ordre (`.001`, `.002`, …). On peut en joindre plusieurs par message si la limite de Claude Pro le permet. Sinon, une par message.
4. N'ouvre pas et ne renomme pas les parties : elles ne sont lisibles qu'une fois réassemblées. Je les réassemble et je vérifie chaque empreinte.

Si un envoi est refusé parce que le fichier est trop gros, mettre `TAILLE_PARTIE_MO = 10`, puis relancer l'étape 4.

## En cas de problème

| Message | Solution |
|---|---|
| `ERREUR : le dossier source est introuvable` | Vérifier `DOSSIER_SOURCE` (copier le chemin depuis l'explorateur Windows). |
| Installation refusée (`pip`) | Dans le terminal de VS Code : `python -m pip install --user python-calamine openpyxl`, puis relancer. |
| `PermissionError` | Fermer les fichiers Excel ouverts, puis relancer. |
| `MemoryError` | Fermer les autres applications, puis relancer. |
| Le script semble bloqué sur un fichier | Attendre : les grands livres comptables sont longs à lire. |
