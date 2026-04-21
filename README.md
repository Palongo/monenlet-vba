# MONENLET-VBA 💰

> **Montant en Lettres — Excel VBA** | Malgache · Français · Anglais

Bibliothèque VBA open source pour convertir un montant numérique en lettres selon les règles linguistiques officielles de trois langues, prête pour l'usage bancaire et juridique sous Excel.

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
![Language: VBA](https://img.shields.io/badge/Language-VBA-green)
![Excel: 2010+](https://img.shields.io/badge/Excel-2010%2B-green)
![Langues: MG · FR · EN](https://img.shields.io/badge/Langues-MG%20%C2%B7%20FR%20%C2%B7%20EN-orange)

---

## Fonctions disponibles

| Fonction | Langue | Exemple de résultat |
|---|---|---|
| `MONENLET_MG` | 🇲🇬 Malgache | `Sivy amby valopolo sy eninjato ... Ariary` |
| `MONENLET_FR` | 🇫🇷 Français | `MILLE DEUX CENT CINQUANTE ARIARYS` |
| `MONENLET_EN` | 🇬🇧 Anglais  | `ONE THOUSAND TWO HUNDRED AND FIFTY ARIARYS` |

---

## Installation

### 1 — Télécharger les fichiers source

Cloner le dépôt ou télécharger les 4 fichiers `.bas` du dossier `src/` :

```
src/
├── ModuleRegistration.bas   ← à importer EN DERNIER
├── MontantMalagasy.bas
├── MontantFrancaise.bas
└── MontantAnglaise.bas
```

### 2 — Importer dans Excel

1. Ouvrir Excel → créer ou ouvrir un classeur
2. Appuyer sur **`Alt + F11`** (ouvre l'éditeur VBA)
3. Menu **Fichier → Importer un fichier...** (`Ctrl + M`)
4. Importer dans cet ordre :
   - `MontantMalagasy.bas`
   - `MontantFrancaise.bas`
   - `MontantAnglaise.bas`
   - `ModuleRegistration.bas` ← **en dernier**
5. Sauvegarder le classeur au format **`.xlsm`** (pas `.xlsx`)
6. Fermer et rouvrir le fichier — les 3 fonctions sont enregistrées

> ⚠️ **Important** : Le fichier doit être sauvegardé en `.xlsm` pour que les macros se conservent. Si Excel demande d'activer les macros à l'ouverture, cliquer sur **Activer**.

---

## Syntaxe

### MONENLET_MG — Malgache

```
=MONENLET_MG(Montant; [Devise])
```

| Paramètre | Type | Défaut | Description |
|---|---|---|---|
| `Montant` | Double | requis | Montant numérique à convertir |
| `Devise` | Texte | `"Ariary"` | Libellé de devise affiché |

**Exemples :**

| Formule | Résultat |
|---|---|
| `=MONENLET_MG(1352689.60)` | `Sivy amby valopolo sy eninjato sy roa arivo sy dimy alina sy telo hetsy sy iray tapitrisa Ariary faingo enimpolo` |
| `=MONENLET_MG(11471786)` | `Enina amby valopolo sy fitonjato sy arivo sy fito alina sy efatra hetsy sy iraika amby folo tapitrisa Ariary` |
| `=MONENLET_MG(1000)` | `Arivo Ariary` |
| `=MONENLET_MG(10000)` | `Iray alina Ariary` |

---

### MONENLET_FR — Français

```
=MONENLET_FR(Valeur; [NbDecimales]; [Devise])
```

| Paramètre | Type | Défaut | Description |
|---|---|---|---|
| `Valeur` | Double | requis | Montant numérique à convertir |
| `NbDecimales` | Entier | `2` | Nombre de décimales (0 = ignorer centimes) |
| `Devise` | Texte | `"ARIARY"` | Devise (`"ARIARY"` gère le pluriel automatique) |

**Exemples :**

| Formule | Résultat |
|---|---|
| `=MONENLET_FR(1250.25)` | `MILLE DEUX CENT CINQUANTE ARIARYS ET VINGT-CINQ CENTIMES` |
| `=MONENLET_FR(80; 0)` | `QUATRE-VINGTS ARIARYS` |
| `=MONENLET_FR(200; 0)` | `DEUX CENTS ARIARYS` |
| `=MONENLET_FR(201; 0)` | `DEUX CENT UN ARIARYS` |
| `=MONENLET_FR(71; 0)` | `SOIXANTE ET ONZE ARIARYS` |

---

### MONENLET_EN — Anglais (British standard)

```
=MONENLET_EN(Valeur; [NbDecimales]; [Devise])
```

| Paramètre | Type | Défaut | Description |
|---|---|---|---|
| `Valeur` | Double | requis | Amount to convert |
| `NbDecimales` | Integer | `2` | Decimal digits (0 = ignore cents) |
| `Devise` | Text | `"ARIARY"` | Currency (`"ARIARY"` auto-pluralises) |

**Examples:**

| Formula | Result |
|---|---|
| `=MONENLET_EN(1250.25)` | `ONE THOUSAND TWO HUNDRED AND FIFTY ARIARYS AND TWENTY-FIVE CENTS` |
| `=MONENLET_EN(1001; 0)` | `ONE THOUSAND AND ONE ARIARYS` |
| `=MONENLET_EN(1100; 0)` | `ONE THOUSAND ONE HUNDRED ARIARYS` |
| `=MONENLET_EN(200; 0)` | `TWO HUNDRED ARIARYS` |

---

## Règles linguistiques

### 🇲🇬 Malgache — Règles clés

| Règle | Explication | Exemple |
|---|---|---|
| Ordre ascendant | Lecture de gauche à droite : petit → grand | `enina amby valopolo sy eninjato...` |
| `amby` | Connecteur unité + dizaine | `9 amby valopolo` = 89 |
| `sy` | Connecteur groupe + groupe supérieur | `eninjato sy roa arivo` |
| `faingo` | Virgule décimale | `Ariary faingo enimpolo` |
| `iraika` | Forme de "1" avant `amby` | `iraika amby folo` = 11 |
| `arivo` seul | 1 000 exact sans "iray" | `arivo` ≠ `iray arivo` |
| `iray alina` | 10 000 avec "iray" obligatoire | `iray alina` ≠ `alina` |

### 🇫🇷 Français — Règles clés

| Règle | Exemple |
|---|---|
| `ET` pour 21, 31, 41, 51, 61, 71 | VINGT ET UN, SOIXANTE ET ONZE |
| Tiret dizaine-unité | VINGT-DEUX, QUATRE-VINGT-DIX |
| `QUATRE-VINGTS` avec S pour 80 exact | QUATRE-VINGTS ≠ QUATRE-VINGT-UN |
| `CENT` prend S uniquement si final | DEUX CENTS, DEUX CENT UN |
| `MILLE` sans UN | MILLE ≠ UN MILLE |
| `MILLIONS`/`MILLIARDS` prennent le pluriel | DEUX MILLIONS |

### 🇬🇧 Anglais — Règles clés (British standard)

| Règle | Exemple |
|---|---|
| Tiret dizaine-unité | TWENTY-ONE, NINETY-NINE |
| `AND` après HUNDRED | ONE HUNDRED AND ONE |
| `AND` après THOUSAND si reste < 100 | ONE THOUSAND AND ONE |
| Pas de `AND` si reste ≥ 100 | ONE THOUSAND ONE HUNDRED |
| `ONE THOUSAND` (contrairement au FR) | ONE THOUSAND ≠ FR MILLE |
| Jamais de S sur multiplicateurs | TWO HUNDRED, TWO MILLION |

---

## Architecture du projet

```
monenlet-vba/
├── src/
│   ├── ModuleRegistration.bas   Auto_Open centralisé (1 seul par classeur)
│   ├── MontantMalagasy.bas      MONENLET_MG + RegisterMONENLET_MG
│   ├── MontantFrancaise.bas     MONENLET_FR + RegisterMONENLET_FR
│   └── MontantAnglaise.bas      MONENLET_EN + RegisterMONENLET_EN
├── docs/
│   └── NOMENCLATURE.md          Tables de référence complètes
├── examples/                    (répertoire pour futurs exemples)
├── .gitignore
├── LICENSE                      MIT
└── README.md
```

### Pourquoi `ModuleRegistration.bas` ?

VBA n'autorise qu'**un seul** `Auto_Open` par classeur. Avoir trois modules avec chacun leur `Auto_Open` provoque l'erreur `Ambiguous name detected` qui empêche la compilation du projet entier. Ce module centralise l'enregistrement des info-bulles Excel pour les trois fonctions.

---

## Capacité et limites

| Critère | Valeur |
|---|---|
| Valeur maximale | 999 999 999 999 (999 milliards / billion) |
| Décimales | 0 à 2 chiffres |
| Valeurs négatives | Converties automatiquement via `Abs()` |
| Excel minimum | 2010 (Windows et Mac) |
| `Option Explicit` | Activé — typage strict |

---

## Contribuer

Les contributions sont les bienvenues :

1. **Fork** le dépôt
2. Créer une branche : `git checkout -b feature/ma-contribution`
3. Commit : `git commit -m "feat: description"`
4. Push : `git push origin feature/ma-contribution`
5. Ouvrir une **Pull Request**

### Idées d'améliorations

- [ ] Support d'autres devises avec accord de genre (ex : euro/euros)
- [ ] Version `MONENLET_AR` pour l'Arabe
- [ ] Version `MONENLET_SW` pour le Swahili
- [ ] Fichier `.xlsm` de démonstration prêt à l'emploi

---

## Auteur

**Justin FARALAHY**
- 🏢 [MAAS — Managn'Asa](https://managnasa.co) · Nosy Be, Madagascar
- 📊 [Karamako.mg](https://karamako.mg) — Transparence salariale Madagascar
- 🎨 [ARS - Madagascar](https://www.facebook.com/artregnerstudio/) — Art'Régner Studio (ARS)

---

## API/Module

*   "API REST disponible → [monenlet-api](https://monenlet-api.vercel.app/)"
*   "Module Excel/VBA → [monenlet-vba](https://github.com/Palongo/monenlet-vba.git)"
---

## Licence

Ce projet est distribué sous licence **MIT** — voir [LICENSE](LICENSE).

Libre d'utilisation, de modification et de distribution, y compris pour un usage commercial, sous réserve de conserver la mention de copyright.
