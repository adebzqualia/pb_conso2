# Vérification du reporting détaillé — 22 octobre 2026

## Livraison

- Deux modes conservés : synthèse initiale, détaillé ; `--report both` produit les deux.
- Exemple détaillé : 19 onglets, dont 10 onglets de présentation (11 pages pour le périmètre fourni), 14 graphiques Excel natifs.
- Six pays : Austria, Italy, South Africa, Denmark, Norway, Sweden.
- Original `Synthese_2026Q2.xlsx` conservé octet pour octet : SHA-256 `221B97A3FEFFA4DEF6E8DDE57F2DB2623A2EF55F8A5926CE99D0D94EF1DA438A`.
- Aucune écriture dans les fichiers pays, Taux.xlsx ou l'ancienne synthèse manuelle.

## Calculs et tests

- 25 tests réussis : 13 tests de consolidation initiale, 7 tests du modèle détaillé et 5 tests du classeur détaillé. Les tests ajoutés couvrent la structure native des graphiques, les données manquantes, les totaux, un périmètre étendu à 13 pays, les commentaires longs et les deux modes CLI.
- Recalcul indépendant avec Artifact Tool : les 1 701 cellules contenant des formules concordent avec les résultats Python enregistrés, à 0,00001 unité près. Aucun résultat d'erreur inattendu dans le balayage des cellules.
- L'API de ce moteur peut renvoyer un objet Date pour une formule numérique dépendant d'un critère DATE(). La vérification normalise cet objet en numéro de série Excel avant comparaison ; le fichier livré conserve ses formats numériques et monétaires.
- Modification temporaire d'un stock source de +100 EUR dans le moteur indépendant : propagation de +100 EUR dans la vue Global, puis restauration de la valeur. La comparaison Q1 absente reste `n.d.` après recalcul.
- Contrôle de la structure OOXML : graphiques combinés empilés et courbes, pourcentages sur un axe secondaire droit, plages GraphData internes, aucune liaison externe. Les sources de graphiques disposent de caches enregistrés.
- Vérification des répartitions KPI_3, des catégories Nordics et des efficacités pondérées à partir des fichiers pays. Les libellés commençant par `<` ou `>` sont comparés littéralement, sans être interprétés comme des opérateurs de filtre.
- CLI testé en répertoire temporaire : génération des deux versions, puis génération du seul détaillé sans changer l'empreinte de la synthèse ; aucune duplication du trimestre dans l'historique.

## Présentation

Les 11 pages de présentation, GraphData, Presentation inputs et le guide modifié ont été rendus et inspectés. Les mises en page des tableaux sources inchangés reprennent la vérification de la première version. Les intitulés, devises, tranches, totaux, lignes de comparaison et espaces de commentaires restent lisibles. Les longues listes de pays et de commentaires ont aussi été vérifiées pour les chevauchements de cellules et les zones d'impression.

Le rendu indépendant affiche parfois l'axe secondaire à gauche et lisse les courbes, malgré les paramètres natifs droits et non lissés du fichier. La structure Excel a donc été contrôlée séparément. Microsoft Excel et PowerPoint eux-mêmes n'ont pas été automatisés pour cette vérification ; le classeur n'a pas été réexporté par le moteur de prévisualisation.

## Interprétation des chiffres

- Le stock des tableaux pays est celui à traiter fourni par KPI_3. Il ne doit pas être assimilé au stock identifié incluant les clients injoignables.
- Pour Italy : KPI_3 à traiter = 787 426,24 EUR ; stock identifié KPI_2 = 3 127 145,95 EUR. Le chiffre photographié comporte une autre classification ; aucune correction non validée n'est appliquée.
- Les pays sont affichés en devise locale, les groupes en EUR. Une conversion de la page pays en EUR ne modifie pas les seuils de tranches de la source ; une note l'indique.
- Les mois de South Africa sont janvier à juin 2026 dans le fichier, même si la photographie porte d'autres libellés.
- L'historique commence à Q2 2026. Q1, les comptes CST et les détails absents restent indisponibles. Les formats particuliers Romania/Brazil, Germany/UK et la reprise de l'historique manuel restent à traiter séparément.
- Les 29 observations d'import sont conservées dans Controles. Aucune donnée photographiée ni commentaire métier non validé n'a été repris comme source de calcul.
