# Retrait des mentions internes (versions, lots, journaux, empreintes) dans le texte destiné au client.
# Exécuté par patch_v10.py (exec) avec rep() et s déjà définis.
rep('["Documents reçus le 22/10/2026 (lots 1 à 3, 116 fichiers)", "États financiers, rapports de gestion, cahiers des charges, manuel d\'épargne, étude des ratios, DISA, grands livres, rapprochements bancaires, conventions", "Dossier 04-Documents reçus du client ; journaux journal_lot1.csv à journal_lot3.csv"],',
    '["Documents reçus le 22/10/2026 (124 fichiers)", "États financiers, rapports de gestion, cahiers des charges, manuel d\'épargne, étude des ratios, DISA, grands livres, rapprochements bancaires, conventions, états des allocations décès, arrêté du CMP, dérogation de Cocody, désignation du directeur général", "Noms exacts en annexe « Fichiers sources »"],')
rep('["Classeur des calculs v9", "Tableaux, ratios, provisions, formules (onglet 11_formules), sensibilités (onglets 13), provision du personnel (14), mortalité (15), EFC (16), contrôles croisés (17), traçabilité et validité", "FPPN_Calculs_rapport_20261015_v9.xlsx"]]},',
    '["Classeur des calculs", "Tous les calculs du rapport, expliqués onglet par onglet : tableaux, ratios, provisions, formules, sensibilités, provision du personnel, mortalité, EFC, contrôles croisés, revue actuarielle (onglets 18 à 25), traçabilité et validité", "FPPN_Calculs_rapport_20261022.xlsx"]]},')
rep("implication=\"Les réponses permettront de lever les réserves de la version 9 sur l'épargne et la provision du personnel.\",",
    "implication=\"Les réponses permettront de lever les réserves du rapport sur l'épargne et la provision du personnel.\",")
rep('"Rapprochement avec les agrégats de la version 2"]', '"Rapprochement avec les agrégats du premier traitement"]')
rep("Chaque fichier reçoit une empreinte SHA-256 qui prouve qu'il n'a pas changé.", "Chaque fichier reçoit une empreinte numérique qui prouve qu'il n'a pas changé.")
rep("    (\"SHA-256\", r\"SHA-256\", \"algorithme d'empreinte numérique\"),\n", "")
rep("    (\"SHA-256\", r\"SHA-256\", \"algorithme d'empreinte numérique servant à vérifier qu'un fichier n'a pas été modifié\"),\n", "")
rep("\"notes\": \"Rapport d'étape au 15/10/2026, version 10 (réponse aux quatre demandes, revue actuarielle complète, adossement actif-passif ; version 9 : documents du 22/10/2026,",
    "\"notes\": \"Rapport d'étape (réponse aux quatre demandes, revue actuarielle complète, adossement actif-passif ; documents du 22/10/2026,")
