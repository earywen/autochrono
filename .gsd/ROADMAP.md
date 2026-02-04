# ROADMAP.md

> **Current Phase**: Not started
> **Milestone**: v1.0

## Must-Haves (from SPEC)

- [ ] GUI Python avec champs configurables
- [ ] Génération fichier .bas personnalisé
- [ ] VBA détection pattern mail + archivage automatique

## Phases

### Phase 1: Project Setup
**Status**: ⬜ Not Started
**Objective**: Initialiser la structure du projet Python avec les dépendances
**Deliverables**:
- Structure de dossiers
- requirements.txt
- Configuration PyInstaller

---

### Phase 2: GUI Application
**Status**: ⬜ Not Started
**Objective**: Créer l'interface graphique de configuration
**Requirements**: GUI-01 à GUI-07
**Deliverables**:
- Fenêtre principale avec formulaire
- Navigation fichiers/dossiers
- Validation des entrées
- Génération du fichier .bas

---

### Phase 3: VBA Module Template
**Status**: ⬜ Not Started
**Objective**: Créer le template VBA avec placeholders
**Requirements**: VBA-01 à VBA-09
**Deliverables**:
- Template .bas avec variables à remplacer
- Logique de détection du pattern
- Logique d'archivage du mail

---

### Phase 4: Integration & Packaging
**Status**: ⬜ Not Started
**Objective**: Intégrer GUI + template, générer l'exécutable
**Deliverables**:
- Génération dynamique du .bas
- Build PyInstaller
- Instructions d'import Outlook
- Test manuel end-to-end
