# Positive Balance — consolidation trimestrielle

Ce projet transforme un dossier trimestriel de fichiers pays en un classeur de synthèse. Les fichiers entrants restent inchangés. La consolidation fonctionne en Python, sans installation d'Excel, avec des noms de fichiers, devises, feuilles et coordonnées configurables.

## Démarrage

Python 3.10 ou supérieur est nécessaire. Depuis ce dossier :

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe reporting.py Q2_2026
```

Le programme produit `outputs/Synthese_2026Q2.xlsx` et conserve les données dans `data/history.json`. L'exemple livré se trouve dans `outputs/positive_balance/Synthese_2026Q2.xlsx`.

Le trimestre est déduit du nom `Q1_2026` à `Q4_2026`, et se généralise aux années suivantes. Les clés internes sont `2026Q1`, etc. Les dates dans les fichiers doivent correspondre au trimestre indiqué.

```powershell
# Choisir les chemins de sortie et d'historique
python reporting.py Q2_2026 --output outputs/Synthese_2026Q2.xlsx --history data/history.json

# Vérifier les fichiers sans écrire de sortie ni d'historique
python reporting.py Q2_2026 --validate-only

# Utiliser une configuration distincte
python reporting.py Q2_2026 --config config

# Dossier sans nom de trimestre : préciser la période
python reporting.py dossier_recu --period 2026Q2

# Refuser aussi les avertissements, notamment les données manquantes et exclusions
python reporting.py Q2_2026 --strict
```

Les chemins relatifs des arguments sont résolus depuis le dossier courant. Le chemin de configuration par défaut est celui du projet. Le code de sortie est `0` en cas de génération/validation réussie, `2` en cas d'erreur. Un résultat avec avertissements peut être généré hors mode strict : lire l'onglet `Controles`.

## Périmètre de cette première version

La configuration fournie consolide six entités : Austria, Italy, South Africa (RCS), Denmark, Norway et Sweden. Les trois pays Nordics sont extraits de leurs feuilles respectives dans un seul fichier, sans ajouter le total Nordics une deuxième fois.

Le dossier d'exemple contient également Romania et Brazil. Leurs fichiers sont explicitement exclus car leurs formats de run-off demandent un adaptateur distinct. Le classeur fourni ne représente donc **pas tous les pays de l'ancienne synthèse**. Les corrections particulières Germany et UK, la collecte/saisie du compte CST Nordics et la migration des anciennes synthèses sont différées.

Les feuilles KPI_1 sont contrôlées pour identifier le pays, la devise et la date. Les mesures financières sont extraites de KPI_2 et les répartitions de stock de KPI_3. Les catégories étendues de Nordics sont conservées sans les renommer artificiellement en « faux positifs » ou « clients injoignables ».

## Configuration

Deux fichiers JSON suffisent pour ajouter un pays au format connu ou adapter une correspondance :

| Fichier | Rôle |
| --- | --- |
| `config/countries.json` | Entités, fichiers, noms attendus, devises, préfixes de feuilles, unités, exclusions |
| `config/layouts.json` | Profils de mise en page, cellules, lignes, colonnes, contrôles de libellés, taux de change |

Exemple d'entité au format standard :

```json
{
  "id": "AT",
  "name": "Austria",
  "file": "AUSTRIA.xlsx",
  "currency": "EUR",
  "country_aliases": ["Austria"],
  "profile": "standard",
  "sheet_prefix": "",
  "units": 1,
  "required": true,
  "enabled": true
}
```

- `id` est une clé stable et unique. Il doit contenir uniquement lettres, chiffres, tirets ou underscores.
- `file` correspond au nom harmonisé manuellement dans le dossier. Un fichier Excel inconnu déclenche une erreur, sauf exclusion ou motif d'ignorance explicite.
- `sheet_prefix` vaut par exemple `DEN_` pour `DEN_KPI_1`, `DEN_KPI_2`, `DEN_KPI_3`.
- `units` multiplie les **montants**, jamais les nombres de contrats. Utiliser `1000` uniquement si le fichier exprime réellement ses montants en milliers.
- `country_aliases` et `currency_aliases` permettent de contrôler les libellés fournis sans modifier les sources.
- `required: true` bloque la génération si le fichier manque. `false` produit une observation et exclut cette entité absente.
- `enabled: false` désactive une entité. `excluded_files` documente les fichiers volontairement exclus.
- `requires_cst: true` conserve le compte CST comme inconnu et rend le stock total avec CST indisponible. Cette version ne fournit pas encore de fichier d'ajustements manuels.

Les profils `standard` et `nordics` définissent les catégories effectivement traitées via `processed_fields`. Une catégorie inexistante n'est pas transformée en zéro. Ajouter un nouveau modèle exige un profil cohérent et des contrôles de libellés correspondants. La structure des fichiers reçus n'est pas modifiée.

La version de base accepte les taux de **clôture**, exprimés en **devise locale pour 1 EUR** : colonne E du fichier fourni. Les paramètres descriptifs `rate_type` et `quote` sont contrôlés ; les changer ne suffit pas à introduire une autre convention. Une nouvelle convention nécessiterait une évolution du code et de ses tests.

## Lecture du classeur

| Onglet | Utilisation |
| --- | --- |
| `Synthese` | Choisir le trimestre en C4. Stocks, flux du semestre et historique de stock réunis dans une feuille |
| `Evolution flux` | Comparer les trimestres de gauche à droite, avec les différentes définitions d'efficacité |
| `Consolidation` | Une ligne calculée par trimestre de reporting et pays |
| `Flux` | Une ligne par trimestre de reporting, pays et cohorte mensuelle |
| `Stock` | Données KPI_2 et calculs par millésime |
| `Repartition stock` | Répartitions KPI_3 par tranche et par millésime, à filtrer séparément |
| `Taux` | Taux propres à chaque trimestre et liste des périodes disponibles |
| `Controles` | Observations d'import, anomalies et empreintes des fichiers sources/configurations |
| `Mode emploi` | Définitions et instructions intégrées au classeur |

Les onglets de données utilisent de vrais **tableaux structurés Excel** avec filtres. Le programme agrandit leurs plages à chaque exécution. Les synthèses utilisent des formules sur les clés pays/période, un sélecteur et des vues chronologiques. Il ne s'agit pas de tableaux croisés dynamiques Excel natifs. L'ajout d'un trimestre se fait en relançant Python, sans insertion manuelle de colonnes.

Les formules sont livrées avec des résultats calculés en Python pour permettre une lecture immédiate, et le classeur demande le recalcul automatique à Excel. Les modifications directes dans le classeur de sortie ne sont pas réimportées : corriger les sources ou la configuration puis relancer. Les contrôles d'import sont un constat à la génération et ne changent pas automatiquement après des modifications manuelles de la sortie.

## Définitions et choix de calcul

**Stock à traiter** = stock initial KPI_2 − somme des catégories traitées applicables. Le stock KPI_3 doit correspondre à ce restant. Ses répartitions par tranche et par millésime décrivent le même stock : elles ne doivent pas être additionnées.

**Stock identifié** = stock à traiter + clients injoignables dans le modèle standard. Pour Nordics, il correspond au stock à traiter disponible dans le fichier. Le CST reçu par mail n'est pas fourni : le **stock avec CST** reste `n.d.`. Une absence nécessaire rend le total concerné indisponible, sans masquer les autres mesures disponibles.

**Flux du semestre glissant** = six cohortes mensuelles observées à la clôture. Q2 2026 contient janvier à juin, mesurés au 30 juin. Il ne s'agit pas uniquement des flux d'avril à juin. La mesure des trois cohortes du trimestre est également présentée séparément.

**Efficacité pondérée** = total traité / total entrant, séparément pour montants et nombres de contrats. L'ancienne « efficacité trimestrielle » reprend en réalité la cohorte la plus ancienne, M-5. Cette mesure est préservée avec un libellé explicite, à côté de l'efficacité pondérée. Les ratios ne sont pas moyennés arbitrairement. Les taux d'efficacité sont indisponibles lorsque le dénominateur vaut zéro.

**Conversion** = montant local / taux de clôture, pour les stocks et les flux, conformément à l'ancienne synthèse. Aucun montant n'est arrondi avant calcul ; les formats Excel ne changent que l'affichage.

**Données manquantes** : `n.d.` distingue les données non disponibles ou catégories non applicables d'un vrai zéro numérique. Les cellules vides et tirets de certaines anciennes répartitions sont conservés comme inconnus.

## Historique et prochain trimestre

```powershell
python reporting.py Q3_2026 --history data/history.json --output outputs/Synthese_2026Q3.xlsx
```

Le fichier JSON contient les données sources normalisées de chaque trimestre, les taux, les contrôles et les empreintes SHA-256. Une nouvelle exécution du même trimestre le remplace complètement, sans doublon. Les autres trimestres sont conservés et ne sont pas reconvertis aux taux actuels. Conserver et sauvegarder ce fichier avec le reporting.

Les cohortes sont identifiées par **trimestre de reporting + pays + mois** : janvier observé au 31 mars et janvier observé au 30 juin sont deux observations distinctes. Ne pas additionner les semestres de plusieurs snapshots.

Le remplacement d'un trimestre refuse un changement de liste des entités. L'historique refuse également le changement de devise d'une même entité. Utiliser un historique séparé pour un changement de périmètre, ou définir une migration contrôlée ultérieure. Les pays peuvent être ajoutés à un nouveau trimestre ; les comparaisons affichent alors les absences des périodes précédentes.

Les écritures utilisent des fichiers temporaires et des remplacements atomiques par fichier. Un verrou empêche deux processus d'écrire le même historique simultanément. Après une interruption anormale, supprimer le fichier `.lock` uniquement après avoir vérifié qu'aucun processus de consolidation ne tourne. Le classeur et le JSON sont deux fichiers distincts : conserver les deux et relancer la commande si une écriture est interrompue.

L'historique livré commence à Q2 2026. Les anciennes colonnes du classeur manuel n'ont pas été reprises automatiquement.

## Contrôles et limites

Le programme bloque les mauvaises dates de trimestre, devises inattendues, feuilles manquantes, libellés déplacés, mois de cohorte incorrects, clés dupliquées, fichiers non configurés, taux manquants/non positifs et formules source sans résultat enregistré. Les totaux et ventilations incohérents sont consignés dans `Controles`.

**Les classeurs source doivent avoir été recalculés puis enregistrés dans Excel.** `openpyxl` lit le cache des formules sans les exécuter. Un cache absent est détecté ; un cache ancien mais présent ne peut pas être identifié de façon fiable. Enregistrer avec une bibliothèque Python qui efface les caches nécessite ensuite un recalcul dans Excel.

Les observations de l'exemple incluent un total KPI_3 autrichien arrondi, des millésimes vides, un libellé EUR incohérent dans RCS et une répartition RCS concentrée dans la première tranche. Les sources sont conservées telles quelles et les anomalies restent visibles.

L'Italie présente un reclassement de 103 901,32 EUR dans la synthèse manuelle par rapport au fichier pays. Cette version suit le **fichier pays**. Certains totaux historiques de la synthèse manuelle omettent l'Afrique du Sud ; les nouveaux totaux portent sur toutes les entités du périmètre configuré. Le détail de l'analyse se trouve dans [analysis/report_findings.md](analysis/report_findings.md).

## Vérification

```powershell
python -m unittest discover -s tests -v
```

Les tests utilisent les exemples locaux et des copies temporaires. Ils vérifient les montants attendus indépendants, les taux, les définitions d'efficacité, l'historique, les entrées invalides, les données manquantes, les formules et caches du classeur, les tables et la protection des sources. Les tests ne modifient pas les fichiers d'origine.

Une vérification complémentaire de la livraison utilise un moteur de calcul indépendant pour recalculer les formules, tester la propagation d'une modification et rendre les onglets. Le programme Python n'a pas besoin de ce moteur. Le comportement dans Microsoft Excel lui-même n'a pas été testé automatiquement dans cette session.

## Deux versions : synthèse et tableaux pour présentation

La version initiale reste disponible avec la commande habituelle, équivalente à `--report summary`. La nouvelle version détaillée reprend l'organisation des exemples de slides : graphique mensuel empilé et courbes d'efficacité à gauche, tableaux de stock à droite et espace de commentaires. Ses graphiques sont des graphiques Excel natifs modifiables, destinés à être copiés dans PowerPoint. Elle ne crée pas de fichier PowerPoint ni de copie des chiffres photographiés.

```powershell
# Générer uniquement la version détaillée, sans régénérer la synthèse initiale
python reporting.py Q2_2026 --report detailed

# Générer les deux versions à partir des mêmes données
python reporting.py Q2_2026 --report both

# Choisir les deux chemins
python reporting.py Q2_2026 --report both --output outputs/Synthese_2026Q2.xlsx --detailed-output outputs/Synthese_detaillee_2026Q2.xlsx

# Utiliser un autre fichier de présentation
python reporting.py Q2_2026 --report detailed --presentation-config config/presentation.json
```

Avec `--report both`, `--output` désigne la synthèse initiale ; le fichier détaillé est créé à côté, sous le nom `Synthese_detaillee_2026Q2.xlsx`, sauf choix explicite de `--detailed-output`. Avec `--report detailed`, `--output` ou `--detailed-output` désigne le seul classeur produit ; les deux options ne peuvent pas être fournies simultanément. Tous les classeurs demandés sont générés en fichiers temporaires avant le remplacement des sorties et la mise à jour de l'historique. Chaque remplacement est atomique par fichier, sans transaction globale entre fichiers.

La présentation détaillée contient une vue par pays et des vues globales, Core et Non Core. Ces regroupements décrivent le périmètre importé : ils ne complètent pas les pays absents et ne créent pas de regroupement Germany + Austria. Les tableaux globaux séparent le stock identifié du stock à traiter, proposent la comparaison au trimestre précédent et répartissent le stock entre les trois millésimes `<année précédente`, `année précédente`, `année courante`. Les données tabulaires et contrôles restent disponibles pour vérifier les graphiques.

Les vues pays utilisent la **devise locale**, les vues globales l'**EUR**. Les six mois sont affichés de gauche à droite d'après les dates des fichiers source. L'indicateur « après six mois » reprend la cohorte **M-5** observée à la clôture ; les efficacités globales sont calculées à partir des numérateurs et dénominateurs consolidés, jamais comme une moyenne simple des pourcentages pays.

Les tableaux pays présentent la répartition **KPI_3 du stock à traiter**. Ils ne représentent pas automatiquement le stock identifié incluant les clients injoignables : cette distinction explique notamment pourquoi l'Italie ne peut pas reproduire les montants du tableau photographié avec le seul KPI_3 fourni. Les écarts de sources restent visibles. Lorsque la ventilation est complète, ses sommes servent de totaux et de dénominateurs ; lorsqu'elle est incomplète, le total KPI_3 déclaré sert de dénominateur, sans transformer les lignes inconnues en zéros. Les libellés des tranches restent ceux de la source, même si les montants d'une vue sont convertis.

La comparaison utilise uniquement le **trimestre calendaire immédiatement précédent**, avec ses propres taux de change. Une période plus ancienne ne le remplace pas. En l'absence d'historique ou de données d'un pays, la comparaison reste `n.d.` ; le programme ne reconstitue pas des valeurs à partir des captures. Les stocks Nordics restent hors CST tant que ce compte n'est pas fourni.

Le fichier `config/presentation.json`, requis uniquement pour les versions détaillée et combinée, configure :

- `groups` et `country_order` : regroupement et ordre des entités ; un pays nouveau non explicitement classé rejoint `default_group`.
- `country_currency` : `local` ou `EUR` ; `country_options.<ID>.currency` permet un choix par pays.
- `country_options.<ID>.show_stock_bins` : affichage des tranches de stock. Il est désactivé pour RCS dans l'exemple, dont la source concentre tout le stock dans la première tranche ; les données et l'observation sont conservées.
- `style` et `category_labels` : couleurs et libellés des graphiques.
- `comments` : commentaires métier par trimestre et entité, indépendants des calculs. Les clés sont les identifiants pays, `GLOBAL`, `core` et `noncore`.

Exemple de commentaires à ajouter dans `comments` :

```json
{
  "2026Q2": {
    "IT": ["Commentaire métier validé pour l'Italie."],
    "GLOBAL": ["Commentaire consolidé validé pour le trimestre."]
  }
}
```

Les commentaires par défaut sont vides. Les actions, explications et décisions affichées dans les photos ne sont pas supposées valides pour les prochaines périodes. Modifier la configuration puis relancer permet de conserver ces commentaires dans le processus d'automatisation ; une modification directe de la sortie Excel sera remplacée au prochain lancement.

Les pages se lisent dans cet ordre : `Global` contient les deux vues stock et flux ; `CORE Stock`, `CORE Flows` et `NON CORE Stock` reprennent les vues de groupe ; les six onglets pays reprennent le graphique combiné, les tableaux KPI_3 et les indicateurs commentés. Les zones d'impression sont configurées en paysage. Copier les graphiques ou les plages dans PowerPoint permet de préparer les slides. Les libellés et couleurs suivent les captures ; aucun logo n'est reconstitué.

Les pages s'allongent lorsque le nombre de pays ou de millésimes augmente. Les commentaires dépassant la place disponible sont conservés dans `Presentation inputs`, avec un renvoi sur la page concernée. `GraphData` contient les données liées par formules aux tableaux sources ; ses cellules vides représentent des séries indisponibles, pas des zéros.

L'exemple détaillé livré est `outputs/positive_balance/Synthese_detaillee_2026Q2.xlsx`. L'archive `positive_balance_python_v2.zip` regroupe le code, les configurations, la documentation et les tests ; les fichiers pays et l'historique ne sont pas inclus. Pour exécuter les tests de rapprochement, conserver le dossier local `Q2_2026`. Le compte rendu de vérification est dans `analysis/validation_detailed.md`.
