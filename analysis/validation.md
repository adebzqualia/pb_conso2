# Validation de la première version — 22 octobre 2026

- 13 tests `unittest` réussis sur les fichiers exemples et des copies temporaires.
- Six entités extraites : AT, IT, ZA, DK, NO et SE.
- 36 cohortes mensuelles et 21 lignes de stock KPI_2 pour le snapshot Q2 2026.
- Rapprochement indépendant des montants, comptes et efficacités avec les cellules des fichiers pays.
- Les empreintes des sources restent inchangées par l'exécution.
- Les erreurs de date, cache de formule absent, taux manquant, fichier inconnu, doublon et écrasement de source sont testées.
- Les données manquantes restent inconnues ; les vrais zéros restent numériques.
- Le remplacement d'un trimestre et l'ajout d'un second snapshot sont testés sans doublon ni changement du taux d'un trimestre antérieur.
- Six tableaux structurés Excel, sélecteur de période, caches numériques et formules présents ; aucun lien externe Excel.
- Recalcul par le moteur indépendant Artifact Tool : aucun résultat d'erreur Excel détecté ; 36 indicateurs clés rapprochés avec les valeurs attendues.
- Modification de contrôle : +100 EUR sur le stock initial Austria donne +100 EUR sur le stock à traiter consolidé.
- Fixture de contrôle à deux trimestres : le taux ZAR doublé en Q3 divise par deux le montant EUR de Q3 sélectionné, sans changer Q2. Le fichier fictif est supprimé après vérification.
- Les neuf onglets ont été rendus pour revue visuelle. La présentation sépare clairement les données indisponibles des zéros.

La livraison contient 29 observations d'import, essentiellement des millésimes historiques vides, deux fichiers exclus, les CST manquants et des incohérences source décrites dans le guide. Ces observations ne sont pas des erreurs de formule du classeur généré.

Microsoft Excel lui-même n'a pas été exécuté pour cette validation. Les caches source déjà présents peuvent être anciens sans que le programme puisse le détecter. L'historique antérieur à Q2 2026 et les exceptions métier ne sont pas migrés.
