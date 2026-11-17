# -*- coding: utf-8 -*-
"""Propositions rédigées à la main pour les taux « pour 100 FCFA » de la v6 sans « soit X % » (17/11/2026).
Clé : (diapo, début de la phrase actuelle). Valeur : phrase proposée, ou consigne si la ligne est un titre de tableau."""
P = {
 (4, "La maladie dépense plus"): "Dans le tableau : « En 2025, pour 100 FCFA de cotisations maladie reçus, 103,90 FCFA sont sortis en soins et en frais, soit 103,90 % des cotisations maladie » | D91",
 (6, "Il doit aussi"): "Il doit aussi les intérêts du plan d'épargne : 3,5 FCFA pour 100 FCFA versés, une seule fois, soit 3,5 % du montant versé, ce qui fait 183 023 590 FCFA.",
 (7, "Sans l'épargne"): "Sans l'épargne, les autres dettes de l'année sont couvertes : 177,05 FCFA disponibles pour 100 FCFA à payer, soit 177,05 % des paiements de l'année.",
 (8, "En 2025, la maladie a dépensé"): "En 2025, la maladie a dépensé 103,90 FCFA pour 100 FCFA reçus, soit 103,90 % des cotisations maladie",
 (8, "Le décès ne dépense"): "Le décès ne dépense que 19,16 FCFA pour 100 FCFA reçus, soit 19,16 % des cotisations décès : son surplus bouche le trou de la maladie",
 (8, "Tableau 5"): "Titre de tableau, sans chiffre : ajouter une colonne « soit en % des cotisations » à côté de chaque taux.",
 (8, "En 2025, pour 100 FCFA"): "En 2025, pour 100 FCFA de cotisations maladie, 103,90 FCFA sont sortis en soins et en frais, soit 103,90 % des cotisations maladie, une fois rangées dans les frais les dépenses qui ne sont pas des soins.",
 (8, "Après correction"): "Après correction, le taux de 2026 est de 64,54 FCFA pour 100 FCFA de cotisations, soit 64,54 % des cotisations, sous la norme de 70 % à 80 %.",
 (13, "Les frais dépassent"): "Les frais dépassent la limite de 20 FCFA pour 100 FCFA de cotisations, soit 20 % des cotisations.",
 (13, "En 2025, la maladie dépense"): "En 2025, la maladie dépense 103,90 FCFA de soins et de frais pour 100 FCFA reçus, soit 103,90 % des cotisations maladie",
 (15, "Les retenues paient"): "Les retenues paient 76 à 85 FCFA pour 100 FCFA dus, soit 76 % à 85 % des sommes dues.",
 (19, "Les ressources propres"): "Les ressources propres couvrent 22 FCFA pour 100 FCFA de dettes, soit 22 % des dettes.",
 (19, "Les frais prennent"): "Les frais prennent 36,34 FCFA pour 100 FCFA de cotisations, soit 36,34 % des cotisations.",
 (19, "Réserve à 5,01"): "Réserve à 5,01 mois pour 6 demandés, frais à 36,34 FCFA pour 100 FCFA de cotisations, soit 36,34 % des cotisations, pour une limite de 20 %",
 (55, "La Solde retient"): "La Solde retient 97 à 100 FCFA pour 100 FCFA demandés, soit 97 % à 100 % des sommes demandées",
 (56, "En janvier 2025"): "En janvier 2025, 76,96 FCFA sont retenus pour 100 FCFA demandés, soit 76,96 % des sommes demandées.",
 (68, "Maladie et décès"): "Maladie et décès prennent 5 FCFA pour 100 FCFA de salaire de base, soit 5 % du salaire de base",
 (70, "En 2025, la maladie dépense"): "En 2025, la maladie dépense 73,53 FCFA pour 100 FCFA reçus selon les comptes, soit 73,53 % des cotisations, et 68,75 FCFA selon la Solde, soit 68,75 %",
 (78, "Soins payés pour 100"): "Titre de figure, sans chiffre : ajouter « (en FCFA pour 100 FCFA, soit en %) » à l'axe.",
 (78, "En 2025, la maladie a dépensé"): "En 2025, la maladie a dépensé 73,53 FCFA de soins pour 100 FCFA reçus, soit 73,53 % des cotisations maladie",
 (79, "6,00 FCFA"): "6,00 FCFA pour 100 FCFA de cotisations en 2021, soit 6,00 %, 73,53 en 2025, soit 73,53 %, et 82,31 à fin août 2026 après correction, soit 82,31 % des cotisations",
 (80, "Après correction des 255"): "Après correction des 255 000 000 FCFA, la maladie dépense 82,31 FCFA pour 100 FCFA de cotisations maladie, soit 82,31 % des cotisations, et non 97,40",
 (80, "Prestations maladie pour 100"): "Ligne de tableau : ajouter « soit 97,40 % » et « soit 82,31 % » sous les valeurs, ou une ligne « en % des cotisations ».",
 (80, "Prestations pour 100 FCFA"): "Ligne de tableau : ajouter « soit 75,20 % » et « soit 64,54 % » sous les valeurs (norme 70 % à 80 %).",
 (82, "En 2025, 57,48"): "En 2025, 57,48 FCFA de soins et d'aides pour 100 FCFA de cotisations, soit 57,48 % des cotisations",
 (89, "En 2020 et 2022"): "En 2020 et 2022, le Fonds rendait moins de 10 FCFA de soins pour 100 FCFA reçus, soit moins de 10 % des cotisations",
 (89, "Le bon repère"): "Le bon repère : soins et frais sous 100 FCFA pour 100 FCFA reçus, soit moins de 100 % des cotisations",
 (89, "Un taux qui approche"): "Remplacer la parenthèse : « … moins de 10 FCFA de prestations pour 100 FCFA de cotisations, soit moins de 10 % des cotisations (20,77 FCFA en 2021, soit 20,77 %) … »",
 (91, "En 2025, 103,90"): "En 2025, 103,90 FCFA de soins et de frais pour 100 FCFA reçus, soit 103,90 % des cotisations maladie : la maladie perd [[montant exact en FCFA à reprendre du tableau]]",
 (96, "Le Fonds dépense 36,34"): "Le Fonds dépense 36,34 FCFA de frais pour 100 FCFA de cotisations, soit 36,34 % des cotisations",
 (103, "16,28 FCFA en banque"): "16,28 FCFA en banque pour 100 FCFA d'épargne à rendre, soit 16,28 % de l'épargne : un retrait de 15,46 % de l'épargne vide la caisse disponible",
 (106, "Le Fonds rend 3,5"): "Le Fonds rend 3,5 FCFA pour 100 FCFA versés, une seule fois, quelle que soit la durée d'épargne, soit 3,5 % du montant versé",
 (115, "Les retenues remboursent"): "Les retenues remboursent 76 à 85 FCFA pour 100 FCFA dus, soit 76 % à 85 % des sommes dues",
 (167, "La lecture C"): "La lecture C ajoute les intérêts du plan d'épargne s'ils étaient tous payés dans l'année : 66,96 FCFA pour 100 FCFA à payer, soit 66,96 % des paiements de l'année.",
 (171, "122,63 FCFA"): "122,63 FCFA de biens pour 100 FCFA de dettes, soit 122,63 % des dettes, avant de retirer les biens dont la valeur n'est pas prouvée",
 (180, "69,42 FCFA"): "69,42 FCFA pour 100 FCFA à payer, soit 69,42 % des paiements de l'année.",
 (187, "Le Fonds place 5,53"): "Le Fonds place 5,53 FCFA pour 100 FCFA d'épargne, soit 5,53 % de l'épargne",
 (190, "Les placements ont"): "Les placements ont rapporté 4,90 FCFA pour 100 FCFA placés par an, soit 4,90 % des sommes placées",
 (212, "36,34 FCFA de frais"): "36,34 FCFA de frais pour 100 FCFA de cotisations, soit 36,34 % des cotisations, pour une limite de 20 % : 1,82 fois la limite",
 (213, "36,34 FCFA de frais"): "36,34 FCFA de frais pour 100 FCFA de cotisations, soit 36,34 % des cotisations, pour une limite de 20 % : 1,82 fois la limite",
 (218, "Soins et aides pour"): "Titre de figure : ajouter « (en FCFA pour 100 FCFA, soit en %) ».",
 (218, "En 2025, 57,48"): "En 2025, 57,48 FCFA de soins et d'aides pour 100 FCFA de cotisations, soit 57,48 % des cotisations",
 (221, "Les ressources propres"): "Les ressources propres couvrent 22 FCFA pour 100 FCFA de dettes, soit 22 % des dettes",
 (222, "La trésorerie couvre"): "La trésorerie couvre 13 FCFA pour 100 FCFA de dettes à court terme, soit 13 % des dettes à payer dans l'année",
 (224, "En 2025, le surplus"): "En 2025, le surplus fait 12,81 FCFA pour 100 FCFA de cotisations, soit 12,81 % des cotisations",
 (225, "Les frais prennent"): "Les frais prennent 36,34 FCFA pour 100 FCFA de cotisations, soit 36,34 % des cotisations",
 (225, "Le FPPN dépense 63"): "Le FPPN dépense 63 FCFA de fonctionnement pour 100 FCFA de prestations, soit 63 % des prestations.",
 (278, "Le Fonds place 5,53"): "Le Fonds place 5,53 FCFA pour 100 FCFA d'épargne, soit 5,53 % de l'épargne, et n'a pas de taux écrit : la méthode ne s'applique pas",
 (300, "5 |"): "Ligne d'un tableau d'annexe : ajouter une colonne « soit en % des cotisations ».",
}


def proposition(n, phrase):
    for (d, debut), v in P.items():
        if d == n and phrase.strip().startswith(debut):
            return v
    return "[à rédiger]"
