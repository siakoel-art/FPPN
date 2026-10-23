# -*- coding: utf-8 -*-
# Étape 2 du correctif v11 (exécutée dans l'espace de noms de patch_v11.py) : formulations établies, déclarations surlignées,
# corrections de la vérification de cohérence, nouveaux blocs.

# ------------------------------------------------------------------ tableaux : plusieurs paragraphes par cellule (« \n »)
rep('            n = max(1, math.ceil(len(_brut(v)) / cpl))',
    '            n = sum(max(1, math.ceil(len(_brut(seg)) / cpl)) for seg in str(v).split("\\n"))')
rep('''            for morceau, gras, surl in morceaux_balises(v):
                r = p.add_run()
                r.text = morceau
                r.font.name = "Helvetica"
                r.font.size = Pt(taille)
                r.font.bold = gras_c or gras
                r.font.color.rgb = RGBColor.from_string(coul)
                if surl:
                    surligner(r)''', '''            for li_, ligne_ in enumerate(str(v).split("\\n")):
                if li_:
                    p = c.text_frame.add_paragraph()
                for morceau, gras, surl in morceaux_balises(ligne_):
                    r = p.add_run()
                    r.text = morceau
                    r.font.name = "Helvetica"
                    r.font.size = Pt(taille)
                    r.font.bold = gras_c or gras
                    r.font.color.rgb = RGBColor.from_string(coul)
                    if surl:
                        surligner(r)''')

# ------------------------------------------------------------------ écarts de liquidité calculés avant le résumé exécutif
_d = s.index("# écarts de liquidité année par année au 31/08/2026")
_f = s.index('contenu({"titre": "Adossement : écarts de liquidité année par année"')
_gap = s[_d:_f]
s = s[:_d] + s[_f:]
rep('CHOC_MAL = s[2025]["prest_maladie"] * PIC\n', 'CHOC_MAL = s[2025]["prest_maladie"] * PIC\n' + _gap)

# ------------------------------------------------------------------ résumé exécutif : synthèse des constats, recommandations
rep('_omis(_snap(), {"titre": "Réponse aux quatre questions de la mission"', bloc("bloc_synthese_constats.py") + '\n_omis(_snap(), {"titre": "Réponse aux quatre questions de la mission"')
rep('ENT19 = "déclaration du responsable des cotisations, entretien du 19/10/2026, à confirmer par document"', 'ENT19 = "déclaration du responsable des cotisations, entretien du 19/10/2026"')
rep('SRC_ENT19 = "Source : entretien du 19/10/2026 avec le responsable des cotisations (déclarations à confirmer par document)"', 'SRC_ENT19 = "Source : entretien du 19/10/2026 avec le responsable des cotisations"')
rep('"synthese": "Les causes relèvent de la gestion des données et des choix d\'emploi des ressources ; les effets restent à chiffrer"},',
    '"synthese": "Les causes sont établies par les données ; le manque de liquidité et les pertes sur prêts sont chiffrés en partie 9"},')
rep("Les causes sont démontrées par les données ; les effets sont des risques, c'est-à-dire des conséquences possibles, à confirmer par les analyses en attente.",
    "Les causes sont démontrées par les données. Le manque de liquidité et la dépréciation des prêts sont chiffrés (partie 9). Le coût réel des sinistres et la mesure des engagements par âge attendent la base des prestations et les dates de naissance.")
tranche('contenu({"titre": "Recommandations prioritaires"', 'gab(39, {"titre": "Feuille de route proposée"', bloc("bloc_recommandations.py") + "\n")
rep('("Calculer", "Adenka calcule les provisions (PSAP, Chain Ladder), la mortalité observée et les engagements d\'épargne."),',
    '("Calculer", "Adenka précise la provision maladie (factures datées), la provision décès (dossiers du 41591000), la mortalité observée et les engagements par épargnant."),')

# ------------------------------------------------------------------ partie 1
rep('à un « cahier des charges du régime des prestations sociales » et à un « tableau des allocations », non transmis (PBC, items 100 et 101).',
    'à un « cahier des charges du régime des prestations sociales » et à un « tableau des allocations », reçus le 22/10/2026 (« CAHIER DE CHARGE  MALADIES ET INVAL-DECES 2021.pdf », « CARTOGRAPHIE DES RISQUES COUVERTS ET DES PRESTATIONS OFFERTES.docx »).')
rep('"Le barème officiel est établi ; la répartition de la Solde reste à expliquer."',
    '"Le barème officiel est établi. [[La Solde applique une autre répartition, sans document de notification reçu]] (PBC, item 56)."')
rep('"Complément pharmacie, le policier paie 20 % (déclaration du 19/10/2026)"', '"[[Complément pharmacie, le policier paie 20 % (déclaration du 19/10/2026)]]"')
rep('"**Plusieurs produits n\'ont pas de document qui fixe leurs règles.** Le barème, le règlement de l\'EFC et du PEP, et la documentation de MEDOC n\'ont pas été transmis."',
    '"**MEDOC est le seul produit sans aucun document.** L\'EFC et le PEP sont décrits par le cahier des charges de 2021 et le manuel des produits d\'épargne, sans décision d\'adoption signée (partie 6)."')
rep("Le circuit est celui décrit par le responsable des cotisations le 19/10/2026 ; il reste à confirmer par document. Les formes sont modifiables dans PowerPoint.",
    "Le circuit est celui décrit par le responsable des cotisations le 19/10/2026 ; la procédure écrite est demandée au FPPN. Les formes sont modifiables dans PowerPoint.")
rep('analyse=["**Deux circuits coexistent.** Pour la maladie et le décès, la direction du personnel de la Police transmet directement le fichier aller à la Solde, sans plateforme de partage, et le FPPN ne reçoit que le retour. Pour les autres produits, le responsable des cotisations dépose le fichier aller sur la plateforme et l\'entité concernée dépose le retour (" + ENT19 + ")."],',
    'analyse=["[[**Deux circuits coexistent.** Pour la maladie et le décès, la direction du personnel de la Police transmet directement le fichier aller à la Solde, sans plateforme de partage, et le FPPN ne reçoit que le retour. Pour les autres produits, le responsable des cotisations dépose le fichier aller sur la plateforme et l\'entité concernée dépose le retour]] (" + ENT19 + ")."],')
rep('"Reçu directement par le FPPN, canal à préciser", "Non"]', '"Reçu par le FPPN : « …_effectif_police_936_….xlsx » et « FPPN_ALLOC_DECES_8Z6_….xlsx »", "Non"]')
rep('analyse=["**Ce circuit explique l\'absence totale des fichiers aller maladie et décès dans l\'envoi du 15/10/2026.** Le FPPN ne les produit pas (" + ENT19 + ")."],',
    'analyse=["**Les fichiers aller maladie et décès sont absents de tous les envois reçus.** [[Le FPPN ne les produit pas]] (" + ENT19 + ")."],')
rep('implication="Les fichiers aller maladie et décès sont à demander à la direction du personnel de la Police ou à la Solde (PBC, item 61). Le canal de réception des retours doit être précisé (PBC, item 62).",',
    'implication="Les fichiers aller maladie et décès sont demandés à la direction du personnel de la Police ou à la Solde (PBC, item 61). Les retours vides ou manquants (décembre 2024, décès de mai à octobre 2024 et d\'avril 2025) sont demandés au FPPN (PBC, item 62).",')

# ------------------------------------------------------------------ partie 3 : fichiers endommagés renvoyés lisibles, qualité des bases
rep('''contenu({"titre": "Fichiers vides ou endommagés", "sous_titre": "Onze fichiers sont vides, illisibles ou mal classés et doivent être renvoyés",
         "synthese": "Les fichiers comptables endommagés ont tous un double lisible, sauf le grand livre des comptes 2022"},''',
    '''contenu({"titre": "Fichiers vides ou endommagés", "sous_titre": "Quatre défauts restent ouverts : décembre 2024, MEDOC et deux fichiers mal classés",
         "synthese": "Tous les fichiers comptables endommagés ont été renvoyés lisibles le 22/10/2026, dont le grand livre des comptes 2022"},''')
rep('''                 "lignes": [["Décès 8Z6 et maladie 936, décembre 2024", "Vides (0 octet)", "Le mois de décembre 2024 n'a aucune donnée."],''',
    '''                 "lignes": [["« FPPN_ALLOC_DECES_8Z6_DECEMBRE_2024.xlsx » et « 122024_effectif_police_936_DECEMBRE 2024.xlsx »", "Vides (0 octet)", "Le mois de décembre 2024 n'a aucune donnée. Aucun envoi suivant ne les contient."],''')
rep('''                            ["Balances des tiers 2021 et 2025 (Excel)", "Endommagées", "Adenka a utilisé les PDF lisibles."],
                            ["Grands livres des comptes 2021 et 2024 (Excel)", "Endommagés", "Adenka a utilisé les PDF lisibles."],
                            ["Grand livre des comptes 2022 (Excel et PDF)", "Endommagés", "Adenka a récupéré 417 pages sur 657 ; la balance 2022 est intacte."],
                            ["Grands livres des tiers 2023 et 2025 (Excel)", "Endommagés", "Adenka a utilisé les PDF lisibles."],''',
    '''                            ["Balances des tiers 2021 et 2025, grands livres des comptes 2021, 2022 et 2024, grands livres des tiers 2021, 2023, 2024 et 2025 (Excel)", "Endommagés à la première réception", "Renvoyés lisibles le 22/10/2026 (dossier « 02-Nouveaux_GL et Balances »). « GRAND LIVRE DES COMPTES 2022.xlsx » couvre l'année entière (192 comptes)."],''')
rep('implication="Le renvoi de ces fichiers complètera décembre 2024 et le détail comptable 2022, sans changer les totaux annuels déjà utilisés.",',
    'implication="Le renvoi des fichiers de décembre 2024 complètera l\'année 2024, sans changer les totaux annuels déjà utilisés.",')
rep('retenir("données", (', bloc("bloc_qualite.py") + '\nretenir("données", (')
rep('"La base reproduit les fichiers du FPPN sans perte, mais il manque les données nécessaires aux calculs actuariels."),',
    '"La base reproduit les fichiers du FPPN, à 26 lignes près sans matricule valide (737 890 FCFA, 0,0017 % des précomptes), mais il manque les données nécessaires aux calculs actuariels."),')

# ------------------------------------------------------------------ partie 4 : règles des produits, recrues
rep('"Colonne « règle déclarée » : déclaration du responsable des cotisations, entretien du 19/10/2026, à confirmer par document.', '"Colonne « règle déclarée » : déclaration du responsable des cotisations, entretien du 19/10/2026.')
rep('"entetes": ["Produit", "Règle déclarée (19/10/2026)", "Ce que montrent les fichiers", "À confirmer"]', '"entetes": ["Produit", "Règle déclarée (19/10/2026)", "Ce que montrent les fichiers", "Pièce demandée"]')
rep('"Aucun document formel selon l\'interlocuteur (PBC, item 58)."', '"Décision de 2021 (PBC, item 58)."')
rep('"À partir de 10 000 FCFA, par paliers de 5 000 FCFA. Durée à vérifier."', '"À partir de 10 000 FCFA, par paliers de 5 000 FCFA. Durée d\'au moins trois ans (règlement intérieur, art. 66)."')
rep('["EFC et PEP", "Présentés aux élèves policiers comme obligatoires dès l\'école, sans texte.", "Non mesurable dans les fichiers.", "Aucun texte selon l\'interlocuteur."],',
    '["EFC et PEP", "[[Présentés aux élèves policiers comme obligatoires dès l\'école, sans texte.]]", "Non mesurable dans les fichiers.", "Texte qui fonde l\'obligation (aucun reçu)."],')
rep('["Prestations", "Au moins trois mois de cotisation avant de pouvoir bénéficier des prestations.", "Non mesurable sans base des prestations.",',
    '["Prestations", "[[Au moins trois mois de cotisation avant de pouvoir bénéficier des prestations.]]", "Non mesurable sans base des prestations.",')
rep("tel que décrit par le responsable des cotisations le 19/10/2026 (à confirmer par document). Cadre rouge", "tel que décrit par le responsable des cotisations le 19/10/2026. Cadre rouge")
rep('analyse=["**Les nouvelles recrues souscrivent avant d\'avoir un mécano.** Les élèves policiers n\'ont qu\'un matricule. Le FPPN attend l\'attribution des mécanos, puis reporte chaque mécano sur les fiches de souscription avant de les inscrire au fichier aller. Le premier précompte peut être décalé d\'un mois (" + ENT19 + ").",',
    'analyse=["[[**Les nouvelles recrues souscrivent avant d\'avoir un mécano.** Les élèves policiers n\'ont qu\'un matricule. Le FPPN attend l\'attribution des mécanos, puis reporte chaque mécano sur les fiches de souscription avant de les inscrire au fichier aller. Le premier précompte peut être décalé d\'un mois]] (" + ENT19 + ").",')

# ------------------------------------------------------------------ partie 5 : sinistralité
rep('"**Les ratios très bas de 2020 à 2022 sont suspects.** Des factures de ces années ont probablement été comptabilisées plus tard, ce que seule la base des prestations permettra de vérifier.",',
    '"**Les ratios de 2020 à 2022 sont bas et s\'expliquent par les textes et les comptes.** Le règlement intérieur limite la couverture aux soins liés à une hospitalisation (art. 41 et 45). Sur 2023 à 2025, seules années mesurables, moins de 1 % des factures d\'une année sont enregistrées l\'année suivante (" + ref("Provision maladie : factures arrivées en retard") + ").",')
rep('''implication="Pour 2022 à 2025, la garantie décès consomme moins d'un quart de ses cotisations selon les deux sources. L'écart de 2021 entre les deux sources est à expliquer par le FPPN.",
        valable="garantie décès-invalidité, exercices 2020 à 2025",''',
    '''implication="Pour 2022 à 2025, la garantie décès consomme moins d'un quart de ses cotisations selon les deux sources. L'écart de 2021 est demandé au FPPN (questions au FPPN, partie 14).",
        valable="garantie décès-invalidité, exercices 2020 à 2025 ; prestations 2020 des comptes = obsèques " + F("D20 obs", V11["deces_2020"]["obseques"]) + " + prime d'assurance décès " + F("D20 prime", V11["deces_2020"]["prime_assurance"]) + " (compte 65180000)",''')
rep('implication="Le décès subventionne la maladie : si le barème officiel confirme une clé de 70 % pour la maladie, la cotisation maladie ne couvre plus son coût complet et doit être revue.",',
    'implication="Le décès subventionne la maladie : avec la clé des statuts (art. 17 : 3,5 % sur 5 %, soit 70 % pour la maladie), la cotisation maladie ne couvre plus son coût complet depuis 2025 et doit être revue.",')
rep('"Excédent d\'un risque = cotisations du risque - prestations du risque - frais attribués au prorata des cotisations.',
    '"Excédent d\'un risque = cotisations du risque - prestations du risque - charges de fonctionnement (612, 613, 618, 62 à 68) attribuées au prorata des cotisations.')
rep('contenu({"titre": "Affectation des excédents"', bloc("bloc_ratios_garantie.py") + '\ncontenu({"titre": "Affectation des excédents"')

# ------------------------------------------------------------------ partie 6 : épargne
rep('("Limite", "Le règlement des produits n\'a pas été transmis : la part de l\'épargne exigible à court terme et l\'affectation du fonds commun de placement restent à confirmer.")],',
    '("Limite", "Le manuel des produits d\'épargne ne fixe pas d\'échéancier individuel. Le FCP est compté comme actif réservé, aucun document ne l\'affectant à un autre emploi.")],')

# ------------------------------------------------------------------ partie 8 : retraités et annexes
rep("devra être comparé à leur cotisation dès réception des factures. La convention avec la CGRAE n'a pas été reçue.",
    "devra être comparé à leur cotisation dès réception des factures. La convention IPS-CGRAE, reçue le 22/10/2026, ne fixe pas la cotisation.")
rep('"synthese": "Flotte automobile (9W3) et MEDOC (8Z7) : nature, garanties et comptabilisation à préciser ; projets immobiliers suivis en compte d\'attente"},',
    '"synthese": "Flotte automobile (9W3) : convention avec GNA non reçue ; MEDOC (8Z7) : aucun document ; projets immobiliers suivis en compte d\'attente"},')
rep('"La convention avec GNA, les primes reversées, les sinistres et les soldes par adhérent : selon le responsable des cotisations, GNA met mal à jour les soldes (entretien du 19/10/2026)."',
    '"La convention avec GNA, les primes reversées, les sinistres et les soldes par adhérent. [[GNA met mal à jour les soldes (responsable des cotisations, entretien du 19/10/2026)]]."')
rep('"La période de fonctionnement et la décision de suspension : produit décrit comme un complément pharmacie, en pause depuis environ deux mois (entretien du 19/10/2026)."',
    '"La période de fonctionnement et la décision de suspension. [[Produit décrit comme un complément pharmacie, en pause depuis environ deux mois (entretien du 19/10/2026)]]."')

# ------------------------------------------------------------------ partie 9 : bilan, provisions, personnel, adossement, placements
rep('''            ("Limite", "Le bilan n'a pas été transmis, et le 31/08/2026 est provisoire : il ne comprend ni amortissements ni écritures de fin d'année.")],
            "Le bilan n'a pas été transmis : il est reconstitué à partir des soldes des balances.''',
    '''            ("Limite", "Le bilan est reconstitué compte par compte, puis confronté aux états financiers de 2020 au 31/08/2026 reçus le 22/10/2026 (partie 3). La balance au 31/08/2026 est provisoire.")],
            "Le bilan est reconstitué à partir des soldes des balances et confronté aux états financiers reçus le 22/10/2026.''')
rep('"**Le calcul de la provision (PSAP, Chain Ladder) suppose les bases de prestations demandées** (partie 12)."],',
    '"**La provision nécessaire est mesurée sur les factures arrivées après la clôture** (" + ref("Provision maladie : montant nécessaire") + ") : quelques millions de FCFA, très loin du montant au bilan."],')
rep('contenu({"titre": "Provision maladie : délais de paiement des cliniques"', '''contenu({"titre": "Circuit d'une facture de soins", "sous_titre": "La date des soins n'est saisie nulle part : la comptabilité ne connaît que la date d'enregistrement de la facture",
         "synthese": "Le règlement intérieur fixe trois délais : paiement des cliniques en trois mois, demande de remboursement en trois mois, remboursement en six mois"},
        "Logigramme modifiable dans PowerPoint, tiré du règlement intérieur (art. 52, 57 et 60) et du journal SIN des grands livres des tiers 2023 à 2025. Cadre rouge : point de contrôle manquant.",
        logigramme="facture",
        analyse=["**La facture est enregistrée au compte 40110000 à sa réception, avec son numéro et sa date d'enregistrement.** La date des soins n'apparaît que dans le numéro de facture, et seulement pour une partie des factures.",
                 "**La provision maladie se mesure donc sur les factures d'une année enregistrées l'année suivante** (diapositives suivantes). Une base des prestations datée donnerait la date des soins de chaque facture."],
        implication="Saisir la date des soins et la date de dépôt de chaque facture permet de calculer chaque année la provision maladie et le délai réel de paiement des cliniques.",
        source="Source : " + SRC_RI + ", art. 52, 57 et 60 ; grands livres des tiers 2023 à 2025 (journal SIN)")
contenu({"titre": "Provision maladie : délais de paiement des cliniques"''')
rep('"**La règle des trois mois est donc approchée, pas garantie.** Elle limite la durée pendant laquelle une prestation reste à payer, ce qui rend inutile une méthode sur plusieurs années (voir l\'annexe des méthodes non retenues)."],',
    '"**Depuis 2024, le délai moyen reste proche de la règle des trois mois (art. 52).** Cette règle limite la durée pendant laquelle une prestation reste à payer, ce qui rend inutile une méthode sur plusieurs années (annexe des méthodes non retenues)."],')
rep('"**En 2022 et 2023, les écarts restent à expliquer ; en 2021, le premier crédit reprend aussi des dossiers antérieurs à l\'ouverture du compte.**',
    '"**En 2022 et 2023, les crédits du compte diffèrent des états ; en 2021, le premier crédit reprend aussi des dossiers antérieurs à l\'ouverture du compte.**')
rep('source="Source : " + REX + ", art. 10 ; Code CIMA (méthode de référence, article à citer à réception du texte)")', 'source="Source : " + REX + ", art. 10 ; circulaire n° 00230/CIMA/CRCA/PDT/2005")')
rep('"Selon le chef comptable et le responsable des cotisations (entretiens du 19/10/2026, à confirmer par la convention), les primes vont à GNA',
    '"Selon le chef comptable et le responsable des cotisations (entretiens du 19/10/2026), les primes vont à GNA')
rep('analyse=["**En 2024, le débit du compte GNA est égal aux produits des activités annexes** : le compte enregistre les commissions dues par GNA. Ses crédits correspondraient aux versements de GNA (à confirmer).",',
    'analyse=["**En 2024, le débit du compte GNA est égal aux produits des activités annexes** : le compte enregistre les commissions dues par GNA. [[Ses crédits sont les versements de GNA]].",')
rep("L'assiette (primes précomptées et primes payées autrement) est à préciser.\"],", "[[La commission porte aussi sur des primes payées hors précompte]] (convention demandée, PBC, item 89).\"],")
rep('source="Source : entretiens du 19/10/2026 (à confirmer par document) ; " + SRC_COMPTA[9:]', 'source="Source : entretiens du 19/10/2026 ; " + SRC_COMPTA[9:]')
rep('implication="Le compte 196 couvre la provision recalculée selon les deux approches. La méthode du FPPN (droits acquis) reste à confirmer par sa note de calcul (PBC, item 113).",',
    'implication="Le compte 196 couvre la provision recalculée selon les deux approches. [[La méthode du FPPN est celle des droits acquis]] : sa note de calcul est demandée (PBC, item 113).",')
rep('+ (" Le rattachement de ces agents à la Direction Générale, au CMP ou au Grenier est à confirmer par le FPPN." if _ent == "Non documentée" else ""),',
    '+ (" Le rattachement de ces agents à une entité est demandé au FPPN." if _ent == "Non documentée" else ""),')
rep('implication="Ces montants s\'additionnent dans la diapositive « Provision du personnel par entité ». Ils changent avec l\'âge de départ, la date d\'embauche et l\'entité que le FPPN confirmera (PBC, item 113).",',
    'implication="Ces montants s\'additionnent dans la diapositive « Provision du personnel par entité ». L\'âge de départ, la date d\'embauche et l\'entité de chaque agent sont demandés au FPPN (PBC, item 113).",')
rep("Le FPPN n'applique donc pas l'âge de 55 ans ; l'âge retenu par le FPPN est à confirmer.\",",
    "[[Le FPPN n'applique donc pas l'âge de 55 ans]] : la sensibilité ci-dessus chiffre un départ à 58 et à 60 ans.\",")
rep('f"**L\'indemnité projetée la plus élevée atteint {F(\'PIDR max indiv\', V9[\'pidr_max_individuel_2025\'])}** : le plafond de 25 fois le SMIG annuel (convention, art. 40) peut s\'appliquer, mais le montant du SMIG ne figure dans aucun document reçu."],',
    'f"**Le plafond de la convention collective (art. 40 : 25 fois le SMIG annuel de 2 080 heures) vaut {F(\'PIDR plafond\', V11[\'pidr_plafond\'][\'plafond\'])}** avec le SMIG de 75 000 FCFA par mois (décret n° 2022-986 du 21/12/2022). {N(\'PIDR au-dessus\', V11[\'pidr_plafond\'][\'agents_au_dessus\'])} agents le dépassent : la provision passe à {F(\'PIDR plafonnée\', V11[\'pidr_plafond\'][\'dbo_avec_plafond\'])}, toujours sous le compte 196. Le décret n° 2017-210 (art. 4 et 5), plus récent, ne prévoit pas de plafond : le calcul central ne l\'applique pas."],')
rep('''implication="Le FPPN doit confirmer l'âge de départ appliqué, les dates d'embauche, le rattachement des agents recrutés en 2025 et le nombre d'enfants à charge (PBC, item 113).",
        valable="personnel civil de la DISA 2025, évaluation au 31/12/2025 ; plafond conventionnel non appliqué",''',
    '''implication="L'âge de départ appliqué, les dates d'embauche, le rattachement des agents recrutés en 2025 et le nombre d'enfants à charge sont demandés au FPPN (PBC, item 113). Aucune de ces données ne peut porter la provision au-dessus du compte 196.",
        valable="personnel civil de la DISA 2025, évaluation au 31/12/2025 ; plafond conventionnel chiffré à part",''')
rep('source=SRC_DISA + " ; IAS 19, § 145 ; onglets 14a à 14c du classeur v9")',
    'source=SRC_DISA + " ; IAS 19, § 145 ; convention collective interprofessionnelle du 19/07/1977, art. 40 ; décret n° 2022-986 du 21/12/2022 (SMIG) ; onglets 14a à 14c du classeur v9")')
rep("Les engagements maladie et décès sont des dettes de quelques mois, réglées en principe dans les délais du règlement intérieur (art. 52, 57 et 58).\"],",
    "Les engagements maladie et décès sont des dettes de quelques mois : le délai moyen de paiement des cliniques est de \" + N('CE délai 25', DELAI[2025], 2) + \" mois en 2025 (\" + ref('Provision maladie : délais de paiement des cliniques') + \"), et la dette décès s'élève à \" + F('CE dette déc', DETTE_DEC26) + \" au 31/08/2026 (\" + ref('Allocations décès : dette réelle et provision obsèques') + \").\"],")
rep('implication="Une note de calcul annuelle des provisions, signée et présentée à l\'Assemblée générale, permettrait de démontrer leur suffisance chaque année (statuts, art. 85 et 89).",',
    'implication="Chaque année, le service comptable établit une note de calcul qui reprend ce tableau : méthode, fichiers utilisés, montant de chaque provision et comparaison au bilan. Le directeur général la signe, et elle est présentée à l\'Assemblée générale avec les comptes (statuts, art. 85 et 89). Elle démontre la suffisance des provisions à l\'AIRMS et au commissaire aux comptes.",')
rep("; la signification du sigle CMP reste à confirmer par le FPPN.", " ; CMP désigne la Clinique médicale du policier, autorisée par l'arrêté n° 0027/MSHP/DGS/DEPS du 10/02/2020.")
rep("Le FPPN finance des engagements courts avec des actifs longs : c'est le risque de liquidité central, à confirmer avec l'échéancier de l'épargne.",
    "Le FPPN finance des engagements courts avec des actifs longs : c'est le risque de liquidité central, chiffré dans les écarts année par année.")
rep('"Hausse probable (sinistres tardifs, allocations restant à payer)"', '"Baisse : besoins mesurés très inférieurs au bilan (partie 9)"')
rep('"Baisse ou hausse, à établir", REX + ", art. 20"]', '"Sens inconnu sans évaluation", REX + ", art. 20"]')
rep('"Décision fixant le taux (non reçue)"]', '"Manuel des produits d\'épargne (taux de 3,5 % et 5,5 %)"]')
rep('"PCMS ; barème BCEAO des SFD"]', '"PCMS ; politique de crédit du Grenier"]')
rep('"**Aucune correction attendue ne peut améliorer l\'adossement, sauf une réévaluation des terrains.** Les engagements non provisionnés ajouteront du passif, et les actifs dont la valeur n\'est pas démontrée réduiront l\'actif."],',
    '"**Les intérêts dus ajoutent du passif, les actifs non évalués peuvent réduire l\'actif.** Seule la baisse des provisions maladie et décès allège le passif : elle est trop faible pour changer la conclusion (manque cumulé de " + F("AB gap", -GAP[4][5]) + " à cinq ans)."],')
rep('implication="L\'écart mesuré en partie 9 est un minimum. Le rapport peut donc affirmer que le FPPN n\'est pas adossé, même dans l\'hypothèse la plus favorable ; seule l\'ampleur de l\'écart reste à préciser.",',
    'implication="Le FPPN n\'est pas adossé, même dans l\'hypothèse la plus favorable. L\'écart est chiffré : manque cumulé de " + F("AB gap2", -GAP[4][5]) + " à cinq ans (" + ref("Adossement : écarts de liquidité année par année") + ") et couverture de " + pc("AB couv", couverture("31/08/2026", T_MOY)[2]) + " des paiements de l\'année dans le scénario central (" + ref("Adossement : tests de résistance") + ").",')
rep('analyse=["**Le non-adossement est démontré avec les données du FPPN**, quel que soit le scénario retenu ; seule l\'ampleur exacte dépend de l\'échéancier de l\'épargne."],',
    'analyse=["**Le non-adossement est démontré avec les données du FPPN**, quel que soit le scénario retenu. Un échéancier individuel de l\'épargne remplacerait le taux de sortie observé, sans changer la conclusion."],')
# placements : tableau tiré des grands livres
tranche('contenu({"titre": "Placements du FPPN"', 'contenu({"titre": "Décision de placement"', bloc("bloc_placements_v11.py"))
rep("Le circuit est celui décrit par le chef comptable le 19/10/2026 ; il reste à confirmer par document. Cadre rouge : contrôle manquant.", "Le circuit est celui décrit par le chef comptable le 19/10/2026. Cadre rouge : contrôle manquant.")
rep('analyse=["**Le choix des placements ne suit aucune politique validée par les organes du FPPN.** La banque propose, le directeur général décide après un point de trésorerie, et la comptabilité enregistre sur la base des relevés."],',
    'analyse=["**Le choix des placements ne suit aucune politique validée par les organes du FPPN.** [[La banque propose, le directeur général décide après un point de trésorerie, et la comptabilité enregistre sur la base des relevés.]]"],')
rep('"Rendement apparent = produits financiers (classe 77) / encours moyen des placements (52 à 53). Il est irrégulier : l\'essentiel des intérêts est comptabilisé en 2025 à la clôture d\'un dépôt à terme.',
    '"Rendement apparent = produits financiers (classe 77) / encours moyen des placements lu dans les grands livres (DAT et FCP, moyenne pondérée par les jours). Il est irrégulier : les intérêts du DAT PEP, constatés chaque année puis annulés, ne sont encaissés qu\'à sa clôture le 30/10/2025.')
rep("pour un encours moyen de placements d'environ {F('Placements moyens', sum(r['placements_moyens'] for r in rp) / len(rp))}.\",",
    "pour un encours moyen de placements de {F('Placements moyens', sum(r['placements_moyens'] for r in rp) / len(rp))} (grands livres).\",")
rep('"**Le rendement moyen dépasse le taux directeur, mais il est irrégulier.** L\'essentiel des intérêts est comptabilisé en 2025, à la clôture d\'un dépôt à terme.",',
    'f"**Le rendement moyen dépasse le taux directeur, mais il est irrégulier.** En 2025, {F(\'RDP int pep\', V11[\'dat_pep\'][\'interets_2025\'])} d\'intérêts du DAT PEP, accumulés depuis 2020, sont enregistrés en une fois.",')
# rendement comparé au marché : rendement lu dans les grands livres
rep('''                 "lignes": [["Rendement brut des placements du FPPN"] + [pc(f"S8 RL6 {a}", _ri("RL6", 3 + j)) for j, a in enumerate(_A5)] + ["Partie 10"],''',
    '''                 "lignes": [["Rendement brut des placements du FPPN"] + [pc(f"S8 RL6 {a}", _RGL[a]) for a in _A5] + ["Grands livres"],''')
rep('''                            ["Conforme face aux dépôts à terme ?", "-", "-", "-", _OK[_ri("RL6ST", 6)], _OK[_ri("RL6ST", 7)], "Onglet 13i"],
                            ["Conforme face à l'État à 5 ans ?"] + [_OK[_ri("RL6INV", 3 + j)] for j in range(5)] + ["Onglet 13i"],''',
    '''                            ["Conforme face aux dépôts à terme ?", "-", "-", "-", "Oui" if _RGL[2024] >= _tx(_DAT, 11) else "Non", "Oui" if _RGL[2025] >= _tx(_DAT, 12) else "Non", "Grands livres, BCEAO"],
                            ["Conforme face à l'État à 5 ans ?"] + ["Oui" if _RGL[a] >= _ri("INV5", 3 + j) else "Non" for j, a in enumerate(_A5)] + ["Grands livres, Investing.com"],''')
rep('_A5 = list(range(2021, 2026))\n', '_A5 = list(range(2021, 2026))\n_RGL = {x["exercice"]: x["rendement"] for x in V11["rendement_gl"]}\n')
rep('"**De 2022 à 2024, le rendement des placements est inférieur à toutes les références de marché disponibles.** Celui de 2025 les dépasse seulement parce que les intérêts d\'un dépôt à terme sont comptabilisés à sa clôture.",',
    '"**De 2022 à 2024, le rendement des placements est inférieur à toutes les références de marché disponibles.** Celui de 2025 les dépasse seulement parce que les intérêts du DAT PEP accumulés depuis 2020 sont enregistrés à sa clôture, le 30/10/2025.",')
rep("taux servi sur l'épargne : 3,50 % selon le chef comptable, 5,50 % selon le manuel d'épargne\"},", "taux servi sur l'épargne : [[3,50 % selon le chef comptable]], 5,50 % selon le manuel d'épargne\"},")
rep("Les échéances doivent suivre l'échéancier de l'épargne, à établir.\",", "Les échéances doivent suivre l'échéancier de l'épargne, que le FPPN ne tient pas encore (recommandation 2).\",")

# ------------------------------------------------------------------ partie 10
rep('"Risque long, à confirmer", "Mêmes obligations que l\'EFC si la durée du plan dépasse un an."]', '"Risque long (plan d\'au moins trois ans, règlement intérieur, art. 66)", "Mêmes obligations que l\'EFC."]')
rep("Le rapport de réévaluation reste à transmettre (PBC, item 86).", "Le rapport de réévaluation est demandé au FPPN (PBC, item 86).")
rep('"synthese": "Chaque obligation est comparée aux pièces disponibles ; plusieurs restent à documenter par le FPPN"},', '"synthese": "Chaque obligation est comparée aux pièces reçues ; une obligation sans pièce est non démontrée"},')
rep('"La décision ministérielle qui fixe ces conditions n\'a pas été transmise."]', '"Aucune politique de placement. Aucune des quatre conditions n\'est démontrée (partie 9.3)."]')
rep('contenu({"titre": "Contrôle du plan comptable"', bloc("bloc_2aci_disp.py") + '\ncontenu({"titre": "Contrôle du plan comptable"')
rep('separateur("9.2", "Adossement actif-passif"', bloc("bloc_2aci_prov.py") + '\nseparateur("9.2", "Adossement actif-passif"')
rep("La série ANStat après octobre 2024 et les pondérations officielles restent à obtenir.", "Après octobre 2024, seule la série de l'UEMOA publiée par la BCEAO est disponible : elle ne montre pas de hausse des prix de la santé en 2025.")

# ------------------------------------------------------------------ partie 14
tranche('contenu({"titre": "Analyses en attente", "sous_titre": "Les analyses actuarielles centrales dépendent de trois bases non reçues"', 'separateur("A", "Annexes"', bloc("bloc_partie14.py"))

# ------------------------------------------------------------------ annexes : formules et méthodes
rep('elif x["logigramme"] == "deces":\n                LG.provision_deces(sl, 4.0)\n                y_ = 10.6',
    'elif x["logigramme"] == "deces":\n                LG.provision_deces(sl, 4.0)\n                y_ = 10.6\n            elif x["logigramme"] == "facture":\n                LG.circuit_facture(sl, 4.0)\n                y_ = 10.6')

# ------------------------------------------------------------------ page de garde : date de remise
rep('"Text Placeholder 4": "14 octobre 2026"', '"Text Placeholder 4": "23 octobre 2026"')

