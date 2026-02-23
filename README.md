# Dashboard Contact Log — Automated Deal Flow Pipeline

Système automatisé de traitement des Contact Logs pour banquiers d'investissement.
Chaque note d'appel brute est analysée par OpenAI GPT-4o et transformée en données structurées, prêtes pour Looker Studio.

```
Raw_Logs (Google Sheet)
       │  nouvelle ligne détectée (col E vide)
       ▼
Google Apps Script
       │  appel API (ligne unique — privacy-first)
       ▼
OpenAI API (gpt-4o)
       │  JSON structuré retourné
       ▼
Clean_Data (Google Sheet)
       │
       ▼
Looker Studio Dashboard
```

---

## Schéma des données

### Onglet `Raw_Logs` — saisie manuelle

| Col | Champ | Description |
|-----|-------|-------------|
| A | Timestamp | Horodatage de la saisie |
| B | Nom du Banquier | Prénom + Nom |
| C | Nom de l'Investisseur | Fonds / LP / Family Office |
| D | Notes du call | Texte libre, brouillon accepté |
| E | Statut de traitement | Vide → `Traité` ou `Erreur` (écrit par le script) |

### Onglet `Clean_Data` — généré par l'IA

| Col | Champ | Valeurs autorisées |
|-----|-------|--------------------|
| A | Timestamp | copié |
| B | Nom du Banquier | copié |
| C | Nom de l'Investisseur | copié |
| D | Statut du Deal | `Approché` · `Teaser Envoyé` · `NDA Signé` · `Management Pres` · `En Due Diligence` · `Offre Soumise` · `Pass/Dropped` |
| E | Drop Reason | `Valuation` · `Sector Mismatch` · `Fund Lifecycle` · `Conflict of Interest` · `Other` · *(vide si non applicable)* |
| F | Next Steps | Résumé concis de la prochaine action |
| G | Date Prochaine Action | Format `YYYY-MM-DD` ou vide |

---

## Déploiement — Guide pas à pas

### Étape 1 — Préparer le Google Sheet

1. Ouvrez (ou créez) le Google Sheet cible.
2. Dans le menu, allez dans **Extensions → Apps Script**.

### Étape 2 — Déployer le script

1. Supprimez le contenu du fichier `Code.gs` par défaut.
2. Copiez-collez l'intégralité du fichier [`src/ContactLogProcessor.gs`](src/ContactLogProcessor.gs).
3. Sauvegardez (`Ctrl+S` / `Cmd+S`).

### Étape 3 — Stocker la clé API de façon sécurisée

> **Règle absolue :** la clé ne doit jamais apparaître dans le code source versionné.

1. Dans le fichier `ContactLogProcessor.gs`, localisez la fonction `setApiKey()`.
2. Remplacez `'YOUR_KEY_HERE'` par votre clé OpenAI réelle (commence par `sk-...`).
3. **Exécutez `setApiKey` une seule fois** depuis l'éditeur Apps Script (bouton ▶).
4. **Effacez immédiatement** la valeur de la clé dans le code (remettez `'YOUR_KEY_HERE'`).
5. Sauvegardez à nouveau.

La clé est désormais stockée dans **Script Properties** (chiffrement AES-256 côté Google) — elle n'est plus jamais visible dans le code.

Pour vérifier : **Paramètres du projet → Propriétés du script** → vous devriez voir `OPENAI_API_KEY`.

### Étape 4 — Initialiser les onglets

Exécutez la fonction `initializeSheets()` une fois.
Elle crée les onglets `Raw_Logs` et `Clean_Data` avec les en-têtes corrects si nécessaire.

### Étape 5 — Créer le déclencheur automatique

Exécutez la fonction `createTimeTrigger()` une fois.
Le script s'exécutera alors **toutes les heures** et traitera automatiquement toute nouvelle ligne.

Pour exécuter manuellement à tout moment : lancez `processContactLogs()`.

---

## Sécurité & Privacy des données

| Mesure | Implémentation |
|--------|---------------|
| Clé API chiffrée | `PropertiesService.getScriptProperties()` — jamais dans le code |
| Exposition minimale | Seule la ligne en cours est envoyée à l'API (pas le sheet entier) |
| Données sensibles | Le contexte envoyé se limite à : nom du banquier, nom de l'investisseur, notes du call |
| Pas de rétention | L'API OpenAI (sans cache activé) ne conserve pas les données d'inférence entre appels |
| Marquage des erreurs | Les lignes en erreur sont taguées `Erreur` — elles ne sont jamais renvoyées en boucle |

---

## Gestion des erreurs

| Scénario | Comportement |
|----------|-------------|
| Clé API manquante | Exception levée, ligne marquée `Erreur`, log descriptif |
| HTTP 429 / Rate limit | Exception, ligne marquée `Erreur`, traitement des suivantes non bloqué |
| JSON invalide retourné | Exception, ligne marquée `Erreur` |
| Valeur hors nomenclature | `drop_reason` inconnu → `Other` ; `next_action_date` invalide → `null` |
| Onglet manquant | Log d'erreur, arrêt propre sans crasher |

Pour relancer les lignes en erreur après correction : exécutez `retryErrorRows()`.

---

## Prompt système exact envoyé à l'API

```
You are a precision data-extraction engine for an investment bank's M&A deal pipeline.
Your ONLY task: read banker call notes and return a single, raw JSON object.

OUTPUT RULES — ABSOLUTE:
1. Return ONLY the JSON object. Zero text before or after it.
2. No markdown, no code fences, no explanations, no apologies.
3. The JSON must have EXACTLY these four keys:
   - "deal_status"       : string — exactly one of the allowed values listed below
   - "drop_reason"       : string or null
   - "next_steps"        : string or null
   - "next_action_date"  : string (YYYY-MM-DD) or null

ALLOWED VALUES:
deal_status  → "Approché" | "Teaser Envoyé" | "NDA Signé" | "Management Pres" |
               "En Due Diligence" | "Offre Soumise" | "Pass/Dropped"
drop_reason  → "Valuation" | "Sector Mismatch" | "Fund Lifecycle" |
               "Conflict of Interest" | "Other"
               Use null when deal_status ≠ "Pass/Dropped".
next_action_date → Extract from notes if present. Format as YYYY-MM-DD.
                   Use null if no date is mentioned.

INFERENCE RULES:
- If status is ambiguous, choose the most recent confirmed stage (conservative).
- If the investor declined or passed, set deal_status = "Pass/Dropped".
- Never fabricate information not present in the notes.
- For relative dates (e.g. "next Tuesday"), resolve to an absolute date only
  if you can be certain; otherwise use null.

EXAMPLE OUTPUT:
{"deal_status":"NDA Signé","drop_reason":null,"next_steps":"Send the Information Memorandum draft for legal review before circulating to investor.","next_action_date":"2025-07-15"}
```

---

## Looker Studio — Guide de connexion et dashboards

### Étape 1 — Connecter l'onglet `Clean_Data`

1. Ouvrez [Looker Studio](https://lookerstudio.google.com) et créez un nouveau rapport.
2. Cliquez sur **Ajouter des données → Google Sheets**.
3. Sélectionnez votre Google Sheet, puis l'onglet **`Clean_Data`**.
4. Cochez **"Utiliser la première ligne comme en-tête"**.
5. Cliquez **Ajouter** → les 7 colonnes apparaissent comme dimensions/métriques.

> **Conseil :** configurez le type de la colonne `Date Prochaine Action` sur **Date (AAAA-MM-JJ)** dans le panneau des champs pour activer les filtres temporels.

---

### Étape 2 — Graphique en entonnoir (Funnel) — Statut du Deal

Visualise la progression des deals à travers les étapes du pipeline.

1. Insérez un **graphique à barres empilées horizontal** (le natif "Funnel" de Looker Studio n'existe pas ; une barre horizontale triée est l'équivalent standard).
2. Configuration :
   - **Dimension** : `Statut du Deal`
   - **Métrique** : `Nombre d'enregistrements` (Record Count)
   - **Tri** : Manuel dans l'ordre du pipeline (voir ci-dessous)
3. Pour forcer l'ordre du pipeline, créez un **Champ calculé** :

```
CASE
  WHEN Statut du Deal = "Approché"         THEN 1
  WHEN Statut du Deal = "Teaser Envoyé"    THEN 2
  WHEN Statut du Deal = "NDA Signé"        THEN 3
  WHEN Statut du Deal = "Management Pres"  THEN 4
  WHEN Statut du Deal = "En Due Diligence" THEN 5
  WHEN Statut du Deal = "Offre Soumise"    THEN 6
  WHEN Statut du Deal = "Pass/Dropped"     THEN 7
  ELSE 8
END
```

4. Triez le graphique par ce champ calculé (ordre croissant).
5. Appliquez un **filtre** : `Statut du Deal ≠ "Pass/Dropped"` pour n'afficher que le pipeline actif, ou gardez-le pour voir l'attrition complète.

---

### Étape 3 — Graphique circulaire (Pie Chart) — Drop Reason

Visualise la répartition des raisons d'abandon.

1. Insérez un **graphique en anneau** (Donut chart) ou **Pie chart**.
2. Configuration :
   - **Dimension** : `Drop Reason`
   - **Métrique** : `Nombre d'enregistrements` (Record Count)
3. Ajoutez un **filtre de graphique** :
   - Condition : `Statut du Deal = "Pass/Dropped"`
   - Cela garantit que le pie chart ne montre que les deals abandonnés.
4. Activez l'option **"Afficher les étiquettes de données"** avec valeur + pourcentage.

> **Conseil :** ajoutez une **plage de dates** comme contrôle de filtre global (basé sur `Timestamp`) pour filtrer les deux graphiques simultanément par période.

---

## Référence rapide des fonctions

| Fonction | Usage | Fréquence |
|----------|-------|-----------|
| `initializeSheets()` | Crée les onglets et en-têtes | Une seule fois |
| `setApiKey()` | Stocke la clé API dans PropertiesService | Une seule fois |
| `createTimeTrigger()` | Active l'exécution horaire automatique | Une seule fois |
| `processContactLogs()` | Traite les nouvelles lignes | Automatique (trigger) ou manuel |
| `retryErrorRows()` | Remet les lignes `Erreur` en file d'attente | Au besoin |

---

## Stack technique

| Couche | Technologie |
|--------|-------------|
| Script | Google Apps Script (V8 runtime) |
| Source de données | Google Sheets |
| LLM | OpenAI (`gpt-4o`) |
| Dashboard | Looker Studio |
| Sécurité clé | Google PropertiesService |
| Déclencheur | Apps Script Time-based Trigger (toutes les heures) |
