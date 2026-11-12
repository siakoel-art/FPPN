# Reprise du travail FPPN dans une session Claude Code sur le PC (12/11/2026)

Lis ce fichier en entier, puis `REGLES_REDACTION_FPPN.md`, avant toute action.

## Contexte
Audit du FPPN (Fonds de prévoyance de la Police nationale, Côte d'Ivoire), volet actuariel et statistique, 2020 au 31/08/2026, pour Adenka.
Le travail a été fait dans une session cloud. Son dossier de travail temporaire (données extraites, bases parquet, JSON de résultats,
documents dézippés) a été effacé. Le dépôt contient tout le code, les livrables et le classeur des calculs, mais pas les données intermédiaires.
Cette session locale a accès aux archives d'origine sur le PC.

## Où sont les données sur le PC
- `C:\Dev\Missions\05-FPPN\04-Documents reçus du client\` : documents reçus du FPPN, dont `01-download` (première dataroom)
  et `11-Documentations au 11112026` (envoi du 11/11/2026, 32 fichiers nouveaux).
- `C:\Dev\Missions\05-FPPN\05-Analyses\` : sorties du script d'extraction (`scripts/extraction_fppn.py`, voir `scripts/GUIDE_extraction_fppn.md`)
  et scripts d'inventaire.
- Dataroom du 09/11/2026 : dossiers `03-Adenka-Audit` et `05-Adenka Audit Juridique et Fiscal` (chemins d'origine `D:\...`, à retrouver).
- Listes des fichiers déjà reçus et empreintes : `dataroom/`.

## État au 12/11/2026
- Dernière version livrée : `rapport/FPPN_Rapport_actuariel_statistique_20261111.pptx` (300 diapositives) et livrables `*_20261111.*`.
- Revue indépendante du rapport par cinq réviseurs : constats bruts dans `revue_independante/constats_A.md` à `constats_E.md`
  (non encore vérifiés un par un). Point le plus lourd, confirmé par deux réviseurs : le passif à un an contient les provisions
  figées de 4 829 298 290 FCFA ; avec le besoin estimé (19 581 392 FCFA), la couverture de l'année passe de 34,42 % à 66,04 %.
- Aperçu des 6 nouvelles diapositives (dettes par type, lectures A et B, intérêts promis) : `rapport/FPPN_Apercu_dettes_interets_20261112.pptx`,
  produit par `scripts/diapos_dettes_v17_fppn.py` (graphiques natifs).
- Chaîne de génération du rapport : `scripts/v11/patch_v11.py`, puis `scripts/v12/patch_v12.py` (exécute `etape2.py` à `etape10.py`),
  qui écrit `scripts/rapport_v12_fppn.py`, puis ce dernier avec ses arguments (résultats JSON, bases parquet, sorties v8 à v11).
  Les scripts amont qui produisent ces entrées sont dans `scripts/` (extraction_fppn, analyses_complements_fppn, ratios_airms_fppn,
  provisions_fppn, donnees_v9/v10/v11_fppn, sensibilites_fppn, sinistralite_gl_fppn…).

## Travaux à faire, dans cet ordre
1. **Reconstituer les données intermédiaires** à partir des archives du PC, en relançant les scripts amont. Les ranger dans un dossier
   `donnees_intermediaires/` (hors des documents du client) et noter les commandes exactes dans `donnees_intermediaires/LISEZMOI.md`.
2. **Test prioritaire sur les intérêts de l'épargne** : pour les policiers sortis (retraite, décès, retrait), comparer le total versé
   (précomptes EFC et PEP) et le montant reçu à la sortie. Reçu > versé : des intérêts sont payés sans être comptabilisés.
   Reçu = versé : la promesse d'intérêts n'est pas tenue (risque de réclamation). Formulation adoucie validée par l'utilisateur.
3. **Vérifier chaque constat de la revue indépendante** sur les données, et dresser la liste : à corriger, à ajouter, confirmé.
   La soumettre à l'utilisateur avant de modifier le rapport.
4. **Réécrire le rapport selon `REGLES_REDACTION_FPPN.md`** : titres-messages, langage simple, tableaux en 12 points découpés,
   graphiques natifs, sources numérotées en bas de page, résumé exécutif en 11 diapositives dans l'ordre validé, ajouts en bleu.
   Intégrer les diapositives de `diapos_dettes_v17_fppn.py` et les lectures A, B (et C avec intérêts).
5. **Construire la version dirigeant** (30 à 40 diapositives) avec les mêmes règles.
6. Plus tard : sinistralité réelle sur les cartes actives, quand l'utilisateur fournira la liste des mutualistes et des cartes actives.

## Règles Git
- Travailler sur la branche `claude/great-gates-u1k9b5`, committer et pousser régulièrement.
- Ne jamais committer les documents du client ni de données nominatives. Les données intermédiaires agrégées peuvent l'être
  si l'utilisateur l'accepte.
