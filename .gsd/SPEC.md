# SPEC.md — Project Specification

> **Status**: `FINALIZED`

## Vision

OutlookToolGen est un outil de productivité unifié pour automatiser l'envoi et la gestion de mails (Archivage Chrono + Accusé de Réception) selon la procédure qualité de l'entreprise. L'objectif est de fournir une interface "idiot-proof" générant un unique code VBA (`ThisOutlookSession`) fusionnant toutes les fonctionnalités sans risque de conflit.

## Goals

1. **Interface Unifiée** — L'utilisateur configure en une seule fois son Trigramme, le nom du Chef de Projet/Téléphone et les chemins des dossiers. Plus besoin de naviguer entre plusieurs outils.
2. **Génération Full-Session VBA** — Produire un seul fichier texte/code contenant "ThisOutlookSession" gérant simultanément :
    - La détection et le classement automatique des mails "Chrono".
    - Le suivi et le classement des "Accusés de Réception".
3. **Zéro Conflit** — Garantir qu'Outlook n'a qu'un seul événement `Application_ItemSend` correctement écrit pour gérer toutes les macros.

## Non-Goals (Out of Scope)

- Modification automatique des macros Outlook en arrière-plan (nécessite droits admin).
- Mise à jour automatique des fichiers de l'entreprise (Excel, etc.) en écriture, sauf ajout d'une ligne pour les nouveaux chronos.
- Applications séparées pour Chrono et AR (Elles sont désormais fusionnées).

## Users

**Employés du secrétariat/bureau et Chefs de Projet** qui souhaitent respecter la procédure qualité d'archivage des mails de manière transparente et sans compétences informatiques.

## Constraints

- **Pas de droits admin** — L'import VBA se fait via copier/coller manuel instruit par l'application.
- **Conflits d'événements Outlook** — Outlook ne supportant pas plusieurs `ItemSend`, le code généré *doit* tout encapsuler.
- **Environnement entreprise partagé** — Fichiers Excel partagés (lecture de préférence).

## Technical Decisions

| Décision | Choix | Justification |
|----------|-------|---------------|
| Langage | Python 3.x | Simple, portable |
| GUI | PyWebView + HTML/Tailwind | Interface web moderne, intuitive, look professionnel |
| Packaging | PyInstaller | Génère un `.exe` standalone portable |
| Output | Code presse-papier | Directement collable dans `ThisOutlookSession`, supprimant le besoin d'installer des `.bas` multiples |

## Functional Requirements

### Application GUI (Python/HTML)

| ID | Requirement |
|----|-------------|
| GUI-01 | Afficher un unique formulaire regroupant : Trigramme, Nom, Tél, Dossier Chrono, Fichier Excel Chrono |
| GUI-02 | Validation stricte des données avant génération |
| GUI-03 | Un seul bouton "Générer le Code Outlook" |
| GUI-04 | Remplacer la vue par onglets par une vue simple de bout en bout |
| GUI-05 | Afficher clairement les 3 étapes d'installation (Copier -> Alt+F11 -> Coller) après génération |

### Module VBA Généré (Fusionné)

| ID | Requirement |
|----|-------------|
| VBA-01 | Encapsuler toute la logique dans un seul bloc `Private Sub Application_ItemSend` |
| VBA-02 | *Chrono* : Détecter le pattern `REF...N°` et proposer l'archivage |
| VBA-03 | *Chrono* : Lire le fichier Excel Excel défini et générer/trouver les dossiers Chrono |
| VBA-04 | *AR* : Détecter le flag `ActionAR` et proposer l'archivage avec `BrowseForFolder` |
| VBA-05 | Fournir la macro autonome `NouveauChrono` en complément dans le même texte pour l'assignation clavier |
| VBA-06 | Fournir la macro autonome `AccuseReception` dans ce même texte |

## Success Criteria

- L'interface ne présente plus de choix entre Chrono et AR. Elle demande toutes les variables nécessaires d'un coup.
- Le Python génère un code VBA robuste qui s'exécute sans erreur de compilation dans Outlook.
- Le code fusionné gère correctement l'interception au moment de l'envoi du mail (ItemSend) pour les deux cas d'usage simultanément.
