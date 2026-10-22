# Analyse du reporting Positive Balance — Q2 2026

Cette analyse repose sur les fichiers sources fournis dans `Q2_2026`, le classeur de synthèse existant et `Taux.xlsx`. Les chiffres ci-dessous ont été lus dans les valeurs calculées enregistrées dans les fichiers ; les formules ont également été examinées. Aucun fichier source n'a été modifié.

## Mesures à préserver

La synthèse existante distingue deux stocks, qu'il faut conserver avec des libellés explicites :

- **Stock restant à traiter** : KPI_2 ligne 17, soit total des soldes générés (ligne 11) moins contrats traités (ligne 12). Les montants traités comprennent remboursements clients (13), clients injoignables (14) et faux soldes positifs (15). Les lignes sont différentes dans le format Nordics étendu.
- **Stock incluant les clients injoignables** : restant à traiter + ligne 14. C'est la mesure de la partie gauche de `Synthèse Q2 2026` et de `Evolution Stock`. Par exemple, `Synthèse Q2 2026!D8 = Italy!C17+Italy!C14`, tandis que `K8 = Italy!C17`.
- **Ventilation KPI_3** : détail du stock restant à traiter par tranche et ancienneté, en montant et nombre de contrats. Ce stock est distinct du stock incluant les clients injoignables.

Les cellules jaunes standard copiées manuellement sont `G4`, `C11:E11`, `C13:E15`, `C27:H28`, `C31:H31` et `C33:H35`. Les totaux et ratios doivent être recalculés depuis ces données, et non lus comme des saisies indépendantes.

## Fenêtre des flux et efficacité

Dans le format standard, `C26:H26` correspond à juin, mai, avril, mars, février et janvier 2026. Les six mois **incluent le mois de reporting**, malgré la note contradictoire « Excluding current period ». Les dates Excel dérivent jusqu'au 28 janvier par chaînage de `EDATE` ; utiliser une clé année-mois ou un vrai dernier jour de mois évite cet artefact.

| Mesure | Formule standard |
| --- | --- |
| Flux généré sur 6 mois | `SUM(C27:H27)` |
| Montant traité sur 6 mois | `SUM(C30:H30)` ; chaque cellule ligne 30 est la somme des lignes 33 à 35 |
| Nombre de contrats générés | `SUM(C28:H28)` |
| Nombre de contrats traités | `SUM(C31:H31)` |
| Moyenne mensuelle des montants | Total 6 mois / 6 |
| Efficacité pondérée en montant | Total traité 6 mois / total généré 6 mois |
| Efficacité pondérée en nombre | Total contrats traités 6 mois / total contrats générés 6 mois |
| Efficacité d'une cohorte mensuelle | Montant ou nombre traité de cette cohorte / montant ou nombre généré de cette cohorte |

**Les lignes 46 et suivantes de `Evolution flux` ne calculent pas une efficacité moyenne.** Pour Q2 2026, elles reprennent la cohorte la plus ancienne du semestre, janvier 2026 : `G50=H34` pour l'Italie en montant et `G62=Italy!H37` en nombre. Il faut distinguer cette mesure, « Efficacité cohorte M-5 », de l'efficacité pondérée sur six mois. Par exemple, l'Italie affiche 98,44 % pour la cohorte M-5, contre 90,26 % pour le semestre pondéré.

L'historique des flux doit porter deux dates : **trimestre d'observation** et **mois de création du solde**. Un mois peut apparaître dans plusieurs reportings avec un niveau de traitement actualisé. Il ne faut ni additionner ces observations successives ni écraser une ancienne observation en utilisant seulement la clé pays/mois.

## Taux de change

Feuille `Table des taux M` : codes devises en colonne A, libellés en B, date de table en `B2` = 1 juin 2026. Cette date désigne le mois de clôture ; elle n'est pas le dernier jour du trimestre.

| Colonne | Signification |
| --- | --- |
| D | Taux d'ouverture |
| E | Taux de clôture |
| F | Taux moyen |
| G | Taux moyen N-1 |
| H | Taux moyen périodique |
| I | Taux de changement de méthode |

La convention du classeur existant est **montant EUR = montant local / taux E**. Le taux exprime donc des unités de devise locale pour un euro. Exemple : `South Africa!B6=1/VLOOKUP($B$5,'Table des taux M'!A4:E184,5,FALSE)`.

Le reporting existant applique le même **taux de clôture aux stocks et aux flux**. Aucune application d'un taux moyen aux flux n'a été trouvée. Garder ce choix configurable, avec clôture comme réglage correspondant au reporting actuel.

| Devise | Cellule du taux de clôture | Unités locales par EUR |
| --- | --- | ---: |
| EUR | E53 | 1 |
| ZAR | E182 | 18.72282849 |
| DKK | E45 | 7.47467107 |
| NOK | E122 | 11.311264655 |
| SEK | E144 | 11.078026565 |
| GBP | E59 | 0.861647587 |
| RON | E136 | 5.23835783 |
| BRL | E26 | 5.90311809 |

Contrôles utiles : mois et année de `B2` cohérents avec le trimestre demandé ; code devise unique ; taux numérique, fini et strictement positif. Les valeurs historiques doivent conserver le taux de leur période, sans être reconverties au taux du dernier trimestre.

## Chiffres de référence issus des sources

Montants EUR ; arrondis uniquement pour présentation. Les valeurs complètes figurent dans `analysis/reconciliation_expected.json`.

| Pays | Stock à traiter | Stock incluant injoignables | Contrats KPI_3 | Flux 6 mois | Traité 6 mois | Efficacité montant 6 mois | Efficacité nombre 6 mois |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| Austria | 11 992,29 | 11 992,29 | 74 | 54 745,46 | 49 360,14 | 90,1630 % | 75,3788 % |
| Italy | 787 426,24 | 3 127 145,95 | 5 595 | 6 678 519,98 | 6 028 261,37 | 90,2634 % | 85,4547 % |
| South Africa | 1 665 285,95 | 1 905 080,93 | 71 052 | 1 175 844,64 | 532 101,49 | 45,2527 % | 66,4063 % |
| Denmark | 24 388,74 | 24 388,74 | 692 | 115 178,67 | 90 789,93 | 78,8253 % | 68,2860 % |
| Norway | 10 807,64 | 10 807,64 | 461 | 61 465,81 | 50 658,17 | 82,4168 % | 58,9127 % |
| Sweden | 14 140,07 | 14 140,07 | 487 | 170 419,97 | 156 279,90 | 91,7028 % | 86,0897 % |

Pour les Nordics, les montants locaux restants sont respectivement 182 297,81 DKK, 122 248,06 NOK et 156 644,06 SEK. Pour l'Afrique du Sud, le montant restant est 31 178 863,21 ZAR et le montant incluant les injoignables est 35 668 503,49 ZAR.

## Écarts et défauts observés

1. **Italie : reclassement manuel dans l'ancienne synthèse.** `ITALY.xlsx!KPI_2!C13` vaut 114 923 062,57, contre 115 026 963,89 dans `Italy!C13` de la synthèse. `KPI_2!C14` vaut 1 852 194,77, contre 1 748 293,45 dans `Italy!C14`. Le total traité et le restant sont inchangés, mais le stock incluant les injoignables diffère de **103 901,32 EUR** : sources 3 127 145,95 ; synthèse 3 023 244,63. Cet écart ne doit pas être corrigé silencieusement par le programme.
2. **Autriche : total KPI_3 arrondi/saisi.** `KPI_3!D11=11992`, alors que `SUM(D6:D10)=11992.29` et `D23=11992.29`. Le détail permet une consolidation exacte ; conserver le contrôle de l'écart de 0,29 EUR.
3. **Totaux incomplets dans la synthèse existante.** `Synthèse Q2 2026!C4=SUM(C5:C11)` exclut la ligne 12 South Africa ; la partie droite `J4=SUM(J5:J12)` l'inclut. Même exclusion de South Africa dans les totaux des flux `Evolution flux!C6=SUM(C7:C13)` et des montants traités ligne 18. Une consolidation par tables doit totaliser toutes les entités explicitement retenues.
4. **Formules cassées anciennes.** Les cellules `Synthèse Q2 2026!M15:M17` des Nordics contiennent `#REF!`. Des historiques stock contiennent aussi `#REF!`, `#DIV/0!` et des cellules vides interprétées comme zéro. Ne pas reproduire ces défauts.
5. **Libellés de devises incohérents.** Des lignes disent « EUR » dans les feuilles pays non-EUR ; la devise de `KPI_2!B5` et les notes du fichier doivent guider la conversion. L'Afrique du Sud met tout le stock dans la première tranche de KPI_3 ; c'est une ventilation source non informative, à conserver et signaler plutôt qu'à redistribuer.

## Format Nordics

Le fichier fourni a un format KPI_2 **étendu**, contrairement à l'hypothèse de format strictement identique. Il comporte les colonnes de millésime `C:F` (`<2024`, `2024`, `2025`, `2026`) et des lignes supplémentaires de traitement : ligne 14 remboursement autorités, ligne 15 sans remboursement, puis détails 16:23. Le restant est en ligne 25. Les flux sont en 35, les nombres en 36, les montants traités en 38, les nombres traités en 39, et les composantes de traitement en 41:43. Les détails 44:51 sont inclus dans la ligne 43 et ne doivent pas être ajoutés une seconde fois.

Les feuilles `DEN_`, `NOR_` et `SWE_` doivent être consolidées individuellement. Les feuilles `KPI_1`, `KPI_2` et `KPI_3` sans préfixe sont un agrégat Nordics et ne doivent pas s'ajouter aux trois pays.

La synthèse actuelle range la ligne 15 « Without refund » Nordics sous « False positive balance » et met les injoignables à zéro. Le résultat restant est cohérent, mais ces libellés ne sont pas équivalents : préserver une catégorie distincte « traité sans remboursement » dans les données normalisées. Le compte CST reçu par courriel est également une mesure distincte et ne doit pas être inventé depuis les KPI.

## Organisation proposée et périmètre initial

Une synthèse principale peut réunir les deux stocks, le volume de contrats, les flux six mois, les deux efficacités pondérées et l'évolution trimestrielle du stock. Les données détaillées doivent prendre la forme de tables Excel filtrables avec colonnes période, pays, devise et mesure. Une vue chronologique doit placer les périodes de gauche à droite. Les tableaux croisés dynamiques restent possibles à partir de ces tables, sans déplacer manuellement six colonnes à chaque trimestre.

L'historique du nouveau programme peut commencer à la première exécution et accumuler les périodes enregistrées ; une migration de l'ancienne synthèse serait un chantier séparé, car ses plages, taux et totaux présentent des incohérences. Réexécuter un trimestre doit remplacer ses observations sans les dupliquer.

Les exceptions métier suivantes restent hors du socle et doivent être visibles comme non traitées ou incomplètes : Brésil `.xlsb` et format spécifique ; Roumanie avec stock seul ; corrections UK ; comptes CST Nordics reçus par courriel ; ajustements Germany. La note Germany fournie est en fait `AE50` : « Faire la somme des segments 100-500 et 500-1000 ». Aucun fichier Germany ou UK source n'est fourni ici pour valider ces règles.
