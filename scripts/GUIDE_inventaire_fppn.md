# Guide pas à pas : lancer l'inventaire des documents FPPN

Durée : 15 minutes d'installation, puis le temps d'exécution (de quelques minutes à plus d'une heure selon le nombre de scans).

## Étape 1 : installer Tesseract (lecture des scans)

1. Télécharger l'installateur Windows de Tesseract (version « UB Mannheim ») :
   https://github.com/UB-Mannheim/tesseract/wiki
2. Lancer l'installateur.
3. À l'écran « Choose Components », déplier « Additional language data » et cocher **French**.
4. Garder le dossier d'installation proposé : `C:\Program Files\Tesseract-OCR`.
5. Terminer l'installation.

Si tu ne peux pas installer Tesseract (droits administrateur), le script fonctionne quand même. Les scans seront seulement signalés « à lire à la main ».

## Étape 2 : placer le script

1. Créer le dossier `...\Missions\FPPN\05-Analyses\scripts`.
2. Y copier le fichier `inventaire_fppn.py`.

## Étape 3 : ouvrir et vérifier le script dans VS Code

1. Ouvrir VS Code, puis `Fichier > Ouvrir le dossier` et choisir `05-Analyses\scripts`.
2. Ouvrir `inventaire_fppn.py`.
3. En haut du fichier, partie **PARAMÈTRES**, vérifier :
   - `DOSSIER_SOURCE` : le dossier où tu as décompressé les documents du client ;
   - `DOSSIER_SORTIE` : le dossier où seront écrits les résultats (`05-Analyses`) ;
   - `TESSERACT_EXE` : à changer seulement si tu as installé Tesseract ailleurs.
4. En bas à droite de VS Code, vérifier que l'interpréteur Python est sélectionné (par exemple « Python 3.11 »). Sinon, cliquer dessus et en choisir un.

## Étape 4 : lancer le script

1. Cliquer sur le triangle « Run Python File » en haut à droite.
   Ou, dans le terminal de VS Code (`Terminal > Nouveau terminal`) : `python inventaire_fppn.py`
2. Au premier lancement, le script installe les bibliothèques manquantes. Cela peut prendre quelques minutes.
3. Le terminal affiche la progression : `[12/850] nom du fichier`.
4. À la fin, il affiche le chemin des deux fichiers produits.

**Si le PC se met en veille, le script s'arrête.** Pour un long traitement, désactive la mise en veille le temps de l'exécution.

## Étape 5 : lire le résultat

Deux fichiers sont créés dans `05-Analyses` :

| Fichier | Usage |
|---|---|
| `Inventaire_FPPN_complet_AAAAMMJJ_HHMM.xlsx` | Pour toi. Contient des extraits de données : ne pas diffuser. |
| `Inventaire_FPPN_masque_AAAAMMJJ_HHMM.xlsx` | Sans extraits de données ni noms de colonnes. À m'envoyer. |

Onglets :

| Onglet | Ce qu'il contient |
|---|---|
| Synthèse | Nombre de fichiers, taille, répartition par type et par pertinence |
| Inventaire | Une ligne par fichier : type, taille, feuilles, colonnes, années, thèmes, pertinence, doublons, erreurs |
| Résumés | Une phrase par fichier qui décrit son contenu |
| Par thème | Les fichiers regroupés par thème du périmètre actuariel |
| Priorités | Les fichiers à analyser en premier |

Lecture des colonnes clés :

- **Pertinence** : « Forte » = fichier de données utile à tes calculs ; « Moyenne » = utile pour comprendre (statuts, états financiers) ; « Faible » = peu de lien avec ton périmètre.
- **Thèmes détectés** : le chiffre entre parenthèses est le nombre de mots-clés trouvés. Plus il est élevé, plus le fichier traite du thème.
- **Doublon de** : le même fichier existe ailleurs ; seul l'original est à analyser.
- **Remarque / erreur** : à lire en priorité ; elle signale les fichiers non lus et leur cause.

La pertinence est calculée automatiquement à partir de mots-clés. Elle oriente, elle ne remplace pas ta lecture.

## Étape 6 : m'envoyer le résultat

Envoie-moi le fichier **masqué**. Je t'aiderai à :
1. confirmer le contenu des fichiers utiles ;
2. pointer ta liste de documents (reçu, partiel, manquant) ;
3. écrire les scripts d'analyse de ton périmètre.

## En cas de problème

| Message | Solution |
|---|---|
| `ERREUR : le dossier source est introuvable` | Vérifier `DOSSIER_SOURCE` (copier le chemin depuis l'explorateur Windows). |
| `pip` ou installation refusée | Lancer dans le terminal : `python -m pip install --user pandas openpyxl xlrd pymupdf python-docx python-pptx pytesseract Pillow` |
| `Reconnaissance de texte (Tesseract) : non` | Tesseract n'est pas trouvé : vérifier `TESSERACT_EXE`. |
| `la langue 'fra' n'est pas installée` | Réinstaller Tesseract en cochant **French** (étape 1). |
| Fichier `.doc`, `.ppt` ou `.xls` ancien en erreur | L'ouvrir et le réenregistrer au format récent (`.docx`, `.pptx`, `.xlsx`). |
| ZIP dans un ZIP | Extraire le ZIP intérieur dans le dossier source, puis relancer. |
| Script trop long | Mettre `OCR_PAGES_MAX = 1`, ou `OCR_ACTIF = False` pour une première passe rapide. |
