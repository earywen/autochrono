# ROADMAP.md

> **Current Phase**: Phase 1: Planning (In Progress)
> **Milestone**: v2.0 (Unified OutlookToolGen)

## Must-Haves (from SPEC)

- [ ] GUI HTML/Python unifiée, sans choix séparés pour AR et Chrono.
- [ ] Moteur de template python fusionnant tout en un seul string pour `ThisOutlookSession`.
- [ ] Exécutable "idiot-proof" générant un code Presse-papier contenant 100% du VBA requis.

## Phases

### Phase 1: Planning and Architecture Revision
**Status**: 🚧 In Progress
**Objective**: Redéfinir l'architecture dans les fichiers `.gsd`.
**Requirements**: Refactor de `SPEC.md` et `ROADMAP.md`.

---

### Phase 2: Backend Refactoring `vba_generator.py`
**Status**: ⬜ Not Started
**Objective**: Simplifier le backend pour qu'il ne propose plus qu'une seule fonction de génération (la fusionnée étendue).
**Requirements**:
- Mettre à jour `ChronoCreatorGenerator` (ou `VBAGenerator`).
- Rédiger un mega-template VBA qui inclut le code Session, les macros Autonomes (`NouveauChrono` et `AccuseReception`) afin de tout livrer d'un coup.

---

### Phase 3: Frontend Simplification `index.html` & `main.py`
**Status**: ⬜ Not Started
**Objective**: Rendre l'interface utilisateur à l'épreuve des balles.
**Requirements**:
- Supprimer les onglets (Chrono vs AR) dans `index.html`.
- Afficher un seul grand formulaire.
- Retirer la fonctionnalité complexe de "sauvegarde de fichier .bas" pour forcer la copie dans le presse-papier (instructions: "Collez le tout dans ThisOutlookSession").
- Mettre à jour API bridge Python (`main.py`).

---

### Phase 4: Integration & Tests
**Status**: ⬜ Not Started
**Objective**: Valider que le code VBA fusionné compilé est valide dans Outlook et que le bouton Python fonctionne parfaitement.
**Deliverables**:
- Lancement de `main.py`.
- Copie du code.
- Vérification visuelle (ou syntaxe) du VBA.
