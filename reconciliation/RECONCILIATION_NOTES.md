# Rapprochement Grand Livre → P&L consolidé — Notes

**Source de vérité :** `P_L_consolidé_par_unité_-_2025_06.xlsx`, onglet `CB035-C` (P&L par destination).
**Reconstruit depuis :** `Grand_Livre_Consolide_2024-2025_vGB.csv` (charges par nature × entité, ISO‑8859‑1, format anglo‑saxon).
**Périmètre du rapprochement :** charges + IS uniquement — le GL n'a **pas** de chiffre d'affaires.

Tous les chiffres ci‑dessous sont produits par `build_reconciliation.py` et formulés dans
`reconciliation_GL_vs_PnL.xlsx` (SUMIF/VLOOKUP, traçables jusqu'aux postes GL). Les mappings
sont externalisés (`mapping_entites.csv`, `mapping_postes.csv`, `mappings.json`) : les modifier
et relancer le script régénère le classeur.

---

## 1. Normalisation des données

- **CSV** : parenthèses → négatif `(275,361)` = −275 361 ; virgule = séparateur de milliers ;
  `-` = 0 ; encodage ISO‑8859‑1. **Contrôle** : la somme recalculée de chaque colonne reproduit
  la ligne `Total` du CSV à ±2 € près (arrondis propres au GL — ex. BM EST −10 726 792 vs
  −10 726 790).
- **Excel** : valeurs `data_only` de `CB035-C`, déjà signées.
- Le poste GL littéral **`#N/A`** est interprété par Excel comme une *valeur d'erreur*. Il est
  ré‑étiqueté `Poste #N/A (non mappe GL)` côté classeur (même poste, toujours traçable) pour ne
  pas casser le VLOOKUP. Les CSV de mapping conservent le libellé source `#N/A`.

## 2. Mapping des entités (traceur = Charges de personnel)

Le traceur « Charges de personnel » mappe 1:1 sans agrégation. Appariement par proximité :

| Entité GL | → Entité P&L | Personnel GL | Personnel P&L | Écart |
|-----------|--------------|-------------:|--------------:|------:|
| H2R | SARL H2R | −34 944 | −35 546 | −1,7 % |
| BM EST | SARL BM EST | −2 481 037 | −2 516 306 | −1,4 % |
| RIVAGROUPE | RIVAGROUPE SAS | −219 277 | −221 792 | −1,1 % |
| VALPOLIS | SARL VALPOLIS | −360 393 | −359 684 | +0,2 % |
| ALLIANCE | SAS GROUPE RIVALIS | −252 782 | −238 132 | **+6,2 %** |
| **KCOLD** | **— (ORPHELIN)** | 0 | — | — |

- 4 entités s'apparient à < 2 %. **ALLIANCE → SAS GROUPE RIVALIS** est le seul appariement
  restant (6,2 %) ; c'est aussi l'entité qui porte les écarts résiduels les plus marqués
  (cf. §4), cohérent avec un écart de source.
- **KCOLD = entité orpheline** : 0 charge de personnel, pas de CA, **total −22 057** (Honoraires
  −5 653, Charges financières −12 896, Dotations −2 210, Locations −600, Services bancaires −698).
  Elle est **isolée** (colonne dédiée dans `GL_retraite`) et **exclue du consolidé** des 5 entités.

## 3. Mapping des postes → lignes P&L (1 ligne P&L = N postes GL)

29 postes GL mappés vers 12 lignes P&L (charges + IS). Détail dans `mapping_postes.csv`. Points
d'arbitrage tranchés :

- **`#N/A` → Charges de personnel.** Découverte traceur : `Charges de personnel` GL + `#N/A`
  reconcilie **exactement** (écart 0) le personnel P&L pour **4 entités sur 5** (BM EST, H2R,
  RIVAGROUPE, VALPOLIS). Seul ALLIANCE reste à −10 534 (écart de source, voir §4). Le `#N/A`
  est donc un **complément de paie** non mappé en nature, pas un poste à jeter.
- **Convention PCG** (arbitrée) : Énergie + Fournitures → *Achats consommés* ; Services bancaires
  + Commissions sur encaissement → *Charges externes* ; Pertes sur créances → *Autres charges
  d'exploitation*.
- **Lignes `[obsolète]`** rattachées à leur nature : `[obsolète] Charges financières` → *Charges
  et produits financiers* ; `[obsolète] Charges exceptionnelles` → *Charges et produits
  exceptionnels* (montant : BM EST −61 085, VALPOLIS −272).
- **Management fees / intra‑groupe** → *Opérations d'exploitation Intra‑Groupe*.
- Le GL ne contient **aucun** détail intra‑groupe financier/exceptionnel : les lignes P&L
  *Opérations financières / exceptionnelles Intra‑Groupe* restent à 0 côté GL (P&L = 0 au total).

## 4. Bridge consolidé (5 entités mappées)

| Ligne P&L | GL recalculé | P&L cible | Écart € | Type |
|-----------|-------------:|----------:|--------:|------|
| Achats consommés | −51 289 | −383 120 | +331 831 | DONNÉES |
| Charges externes | −5 034 946 | −3 067 800 | **−1 967 146** | DONNÉES |
| Charges de personnel | −3 381 993 | −3 371 460 | −10 533 | MAPPING ok |
| Autres charges d'exploitation | −1 009 116 | −1 009 117 | +1 | MAPPING ok |
| Impôts et taxes | −43 735 | −75 042 | +31 307 | DONNÉES |
| Variations amort. & dépréc. | −3 071 861 | −1 392 479 | **−1 679 382** | DONNÉES |
| Opérations d'exploitation Intra‑Groupe | −4 314 875 | 0 | **−4 314 875** | ÉLIMINATION |
| Charges et produits financiers | −374 365 | −79 112 | −295 253 | DONNÉES |
| Charges et produits exceptionnels | −61 357 | −58 309 | −3 048 | MAPPING ok |
| Impôt sur les bénéfices | −69 988 | +32 986 | −102 974 | DONNÉES |
| **TOTAL charges + IS** | **−17 413 525** | **−9 403 453** | **−8 010 072** | |

**Contrôles :** GL 5 entités −17 413 525 **+** KCOLD orphelin −22 057 **=** −17 435 582 (somme
des 6 colonnes du GL). ✔

## 5. Diagnostic des écarts — MAPPING vs DONNÉES

**Écarts de MAPPING (corrigeables) → résiduels quasi nuls.** Personnel (−10 533, soit le seul
ALLIANCE), Autres charges (+1 €), Exceptionnels (−3 048). Le mapping est sain : là où le GL et le
P&L parlent de la même chose, ça tombe juste.

**Écarts de DONNÉES / consolidation (hors de notre contrôle) — l'essentiel du −8,0 M€ :**

1. **Intra‑groupe −4 314 875 (élimination).** Le GL ne porte que la **jambe charge** des
   management fees ; en consolidé la contrepartie produit (dans le CA, absent du GL) l'élimine →
   la ligne P&L nette à **0**. Ce n'est pas un écart de mapping mais une **élimination de conso**.
2. **Variations d'amortissements −1 679 382.** Dotations GL −3 074 071 ≫ P&L −1 392 479. Écart
   concentré sur ALLIANCE (−1 300 993) et BM EST (−1 729 962) : retraitements de consolidation
   (provisions/écarts d'acquisition traités sur une ligne dédiée du P&L, ici vide).
3. **Charges externes −1 967 146.** Nature « externe » du GL > destination P&L. Principal
   suspect : **Honoraires ALLIANCE −1 077 471** (probables frais d'acquisition capitalisés en
   écart d'acquisition, donc retirés des charges en consolidé), + reclassements par destination.
4. **Achats consommés +331 831.** Le P&L « par destination » loge en *Achats consommés* des coûts
   (sous‑traitance/coût des ventes) non identifiables comme « achats » dans la nomenclature par
   nature du GL → reclassement interne. Achats + Charges externes **combinés** : résiduel ramené
   à −1 635 315.
5. **Financiers −295 253, IS −102 974, Impôts & taxes +31 307.** Reclassements / différences de
   version entre le GL statutaire et le P&L consolidé (signe, crédits d'impôt).

## 6. Le consolidé GL annoncé (−6 191 160) — NON réconcilié

La valeur `Consolidé = (6,191,160)` du CSV **ne se relie ni** à la somme des 6 entités
(−17 435 582) **ni** au total charges+IS du P&L (−9 403 453), et **aucun sous‑ensemble propre
d'éliminations** ne la reconstitue (recherche exhaustive jusqu'à 6 postes : aucune combinaison ne
somme à l'écart de 11 244 422 requis). À traiter comme un **total de contrôle externe à vérifier
auprès du producteur du GL** (figuré en rouge dans `Bridge_conso`), distinct du travail de mapping.

## 7. Taux de rapprochement atteint

- **Mapping** : tous les postes sont mappés (0 non affecté), `#N/A` et `[obsolète]` résolus,
  orphelin (KCOLD) isolé. Sur les lignes où GL et P&L sont comparables « pomme à pomme »
  (personnel, autres charges, exceptionnel), l'écart est **< 0,3 %**.
- **Données** : après neutralisation de l'élimination intra‑groupe (structurellement 0 en
  consolidé), le GL reste **−3 695 197** plus négatif que le P&L →
  **taux de rapprochement net ≈ 60,7 %**. En agrégeant Achats + Charges externes (reclassement
  interne par destination) : **≈ 60 %** sur la somme des |résiduels|.
- Le ~40 % non rapproché est **intégralement d'origine DONNÉES** (consolidation : éliminations,
  retraitements d'amortissements/écarts d'acquisition, reclassements par destination), **pas un
  défaut de mapping**. Le GL est pré‑consolidation ; le P&L est post‑consolidation.

## 8. Postes « À ARBITRER » restants

Aucun poste laissé en suspens : tous les arbitrages (PCG, `#N/A`, `[obsolète]`, KCOLD) ont été
validés. **Points à confirmer côté métier** (n'affectent pas le total, < 2 % chacun) :
le rattachement *Services bancaires* / *Commissions sur encaissement* à *Charges externes*
(vs financier) et *Énergie/Fournitures* à *Achats consommés* (vs charges externes) — éditer
`mapping_postes.csv` et relancer pour tester une autre hypothèse.
