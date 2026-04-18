# Changelog

Toutes les modifications notables de ce projet sont documentées ici.

Format basé sur [Keep a Changelog](https://keepachangelog.com/fr/1.0.0/).

---

## [1.0.0] — 2026-04-18 — Publication initiale

### Ajouté
- `MONENLET_MG` v2.0 — Montant en lettres Malgache (lecture ascendante)
- `MONENLET_FR` v2.1 — Montant en lettres Français (normes bancaires)
- `MONENLET_EN` v1.1 — Montant en lettres Anglais (British banking standard)
- `ModuleRegistration.bas` — Auto_Open centralisé (résout le conflit de noms)
- `docs/NOMENCLATURE.md` — Tables de référence complètes (3 langues)

### Corrigé (MONENLET_MG)
- Règle `arivo` : 1 000 exact → `arivo` seul (jamais `iray arivo`)
- Renommage `MONENLETMA` → `MONENLET_MG`

### Corrigé (MONENLET_FR)
- Suppression de `Attribute VB_Name` (erreur de compilation)
- `Decimale As Long` : ajout de `CLng(Round())` pour la précision flottante
- Milliards : `Mod` remplacé par soustraction explicite (imprecision Double > 2^31)
- Valeurs négatives : `Abs()` ajouté en entrée
- Tableaux `Unite`/`Dizaine` déplacés au niveau module (optimisation récursion)
- Renommage `MONENLET_PRO` → `MONENLET_FR`

### Corrigé (MONENLET_EN)
- `Currency` (mot-clé réservé VBA) → renommé `Devise` (lignes rouges corrigées)
- `Auto_Open` → renommé `RegisterMONENLET_EN` (conflit multi-modules résolu)
