# Delta Datacube — Pack d'analyses

## Fichier source
`260511_Delta_Datacube.xlsx` — onglet `ResultatFinancement` (1 ligne = 1 client).

## Comment générer le fichier d'analyses

```bash
pip install openpyxl
python build_delta_datacube_analyses.py \
    "260511_Delta_Datacube.xlsx" \
    "260511_Delta_Datacube_with_Analyses.xlsx"
```

Le script :
1. Ouvre le workbook source
2. Détecte la ligne d'entêtes (scan des 10 premières lignes)
3. Mappe chaque colonne attendue (`Id`, `MRR en cours HT`, `CA_2025`, `NbDocs_YYYY_MM`, etc.) à sa lettre de colonne réelle dans le fichier
4. Ajoute 9 onglets d'analyses, **uniquement des formules** référençant `ResultatFinancement` — aucune valeur en dur
5. Sauve le résultat dans un nouveau fichier (l'original n'est pas modifié)

À l'ouverture dans Excel 365, toutes les formules s'évaluent automatiquement.

## Pré-requis Excel
Le script utilise des formules dynamiques **Excel 365** : `UNIQUE`, `FILTER`, `SORTBY`, `SEQUENCE`. Si tu es sur une version antérieure d'Excel, les listes dynamiques (top NAF, top Sources, etc.) ne fonctionneront pas — il faudra remplacer ces formules par des listes statiques ou utiliser un TCD.

## Onglets générés

| Onglet | Contenu | Formules clés |
|---|---|---|
| **01_KPIs** | Dashboard une page : # clients, MRR, ARR, ARPU, CA 2025/2026, croissance, docs émis, utilisateurs | `COUNTA`, `COUNTIF`, `SUM`, `AVERAGE` |
| **02_MRR_Subscriptions** | Décomposition par TypeAbo, TypeDuree, StatusOffre : # clients, MRR, ARR, ARPU, %MRR, CA | `UNIQUE+FILTER`, `SUMIFS`, `COUNTIFS` |
| **03_Cohorts** | Cohortes par année d'inscription (DateCreation) : # clients, % actifs, MRR, CA, CA moyen + matrice cohorte × mois (NbDocs) | `SUMPRODUCT` avec `LEFT(date,4)` pour extraire l'année |
| **04_Engagement_Monthly** | Pour chaque mois `NbDocs_YYYY_MM` : total docs, # clients ≥1/≥10/≥50 docs, docs moyens/client actif, max | `SUM`, `COUNTIF`, `MAX` |
| **05_Segment_Sector** | Top 20 codes NAF, top 20 secteurs Henrri, répartition forme juridique, tranches d'effectif (1, 2-5, 6-10, 11-50, 51+) | `SORTBY(UNIQUE, -COUNTIF)`, `SUMIFS`, `COUNTIFS` |
| **06_Segment_Geo** | Par département (01-95 + DOM 971-976) : # clients, MRR, CA, % du CA national | `SUMPRODUCT(LEFT(TEXT(CP,"00000"),2)=...)` |
| **07_Acquisition** | Par SourceCreation : # clients, # payants, taux de conversion, MRR, CA + matrice conversion par année de cohorte | `SUMIFS`, `COUNTIFS`, `SUMPRODUCT` |
| **08_Pareto** | Concentration CA 2025 : % top 10/50/100/500 clients + liste détaillée des 50 premiers (ID, ville, NAF, CA, MRR) | `LARGE` + `INDEX/MATCH` |
| **09_Churn_Activity** | Segmentation par activité 2025/2026 (basée sur somme NbDocs) : actifs des deux côtés, churn présumé, réactivés, inactifs | `SUMPRODUCT` sur sommes de plages mensuelles |

## Notes sur les colonnes attendues

Le script attend ces colonnes (détectées dynamiquement, peu importe leur position) :
- Identification : `Id`, `CodePostal`, `LibelleCommune`, `ActivitePrincipaleEntreprise`, `Secteur d'activité récolté dans Henrri`, `forme juridique`, `effectif`
- Acquisition / lifecycle : `SourceCreation`, `DateCreation`, `DatePassageModeOfficiel`, `DatePremierAchat`, `client actif`
- Souscription : `TypeAbo`, `TypeDuree`, `DateAchatEnCours`, `DateFinEnCours`, `MRR en cours HT`, `StatusOffre`, `Nb utlisateurs du compte` (note : typo "utlisateurs" préservée)
- Usage mensuel : `NbDocs_2025_01` à `NbDocs_2026_12` (24 mois max ; seuls ceux présents sont utilisés)
- Revenu : `CA_2025`, `CA_2026`

Si une colonne attendue manque, le script affiche un WARNING mais continue ; seuls les onglets dépendant de cette colonne seront vides.

## Points de vigilance

- **Dates en texte** : `DateCreation` et `DatePremierAchat` sont stockées comme chaînes ISO ("2019-01-16 09:24:..."). Les formules de cohorte utilisent `LEFT(...,4)` pour extraire l'année — pas `YEAR()`.
- **CodePostal en texte ou nombre** : `TEXT(CP,"00000")` gère les deux cas et préserve le 0 initial pour les départements 01-09.
- **Performance** : références "colonne entière" (`$B:$B`) → recalcul peut être lent sur un gros volume. Si trop lent, restreindre à `$B$3:$B$N` avec N = dernière ligne de données.
- **Pareto** : `INDEX/MATCH` sur `LARGE` ne gère pas parfaitement les ex aequo (renvoie le premier client trouvé). Pour M&A, c'est généralement OK.

## Adapter / étendre

Chaque onglet est généré par une fonction `add_*` dans `build_delta_datacube_analyses.py`. Pour ajouter une analyse :
1. Créer une nouvelle fonction `add_my_analysis(wb, cm)` (où `cm` est le dict `header_label -> column_letter`)
2. L'appeler dans `main()`

Utiliser le helper `col_ref('X')` qui retourne `'ResultatFinancement'!$X:$X`.
