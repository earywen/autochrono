# SPEC.md — Project Specification

> **Status**: `FINALIZED`

## Vision

AutoChrono est un outil de productivité pour automatiser le rangement des mails envoyés aux clients selon la procédure qualité de l'entreprise. Il génère un module VBA personnalisé pour Outlook qui détecte automatiquement les mails contenant une référence Chrono et propose de les archiver dans le dossier approprié sur le serveur.

## Goals

1. **Simplifier la configuration** — Une interface graphique permet à chaque utilisateur de configurer ses paramètres (trigramme, chemins) sans modifier de code
2. **Générer du VBA personnalisé** — Produire un fichier `.bas` importable dans Outlook avec les valeurs personnalisées intégrées
3. **Automatiser l'archivage** — Le VBA détecte les mails Chrono à l'envoi et les sauvegarde automatiquement dans le bon dossier

## Non-Goals (Out of Scope)

- Modification automatique des macros Outlook (nécessite droits admin)
- Mise à jour du fichier Excel Chrono
- Gestion multi-années automatique
- Notifications après rangement

## Users

**Employés du secrétariat/bureau** qui envoient régulièrement des offres et rapports aux clients et doivent respecter la procédure qualité d'archivage des mails.

## Constraints

- **Pas de droits admin** — L'import VBA doit être manuel avec instructions claires
- **Environnement entreprise** — L'exécutable doit être "safe" et ne pas déclencher d'alertes antivirus
- **Fichier Excel partagé** — Doit pouvoir lire le fichier même s'il est ouvert par quelqu'un d'autre (lecture seule)
- **Colonnes Excel inconnues** — Les positions des colonnes doivent être configurables

## Technical Decisions

| Décision | Choix | Justification |
|----------|-------|---------------|
| Langage | Python 3.x | Simple, portable, tkinter inclus nativement |
| GUI | Tkinter | Pas de dépendances externes, look natif Windows |
| Packaging | PyInstaller | Génère un .exe standalone sans installation Python |
| Output | Fichier .bas | Format standard pour import VBA dans Outlook |

## Functional Requirements

### Application GUI (Python)

| ID | Requirement |
|----|-------------|
| GUI-01 | Afficher un formulaire avec champs : Trigramme, Dossier Chrono, Fichier Excel Chrono |
| GUI-02 | Permettre la navigation fichier/dossier avec boutons "Parcourir" |
| GUI-03 | Afficher les champs pour les numéros de colonnes Excel (Chrono, Client, Trigramme) |
| GUI-04 | Valider les entrées (champs non vides, chemins existants) |
| GUI-05 | Générer le fichier .bas avec les valeurs saisies |
| GUI-06 | Afficher un message de succès avec le chemin du fichier généré |
| GUI-07 | Inclure un bouton pour afficher les instructions d'import Outlook |

### Module VBA Généré

| ID | Requirement |
|----|-------------|
| VBA-01 | S'accrocher à l'événement `Application.ItemSend` |
| VBA-02 | Détecter le pattern `REF : ... - N°XXXXX` dans les 100 premiers caractères du corps du mail |
| VBA-03 | Parser le pattern avec tolérance aux espaces variables autour des tirets |
| VBA-04 | Afficher une modale "Voulez-vous ranger ce mail dans le dossier Chrono ?" |
| VBA-05 | Si Oui : ouvrir le fichier Excel en lecture seule |
| VBA-06 | Trouver la dernière ligne remplie et extraire : numéro chrono, client, trigramme |
| VBA-07 | Créer le dossier `{chrono} - {client} ({trigramme})` s'il n'existe pas |
| VBA-08 | Sauvegarder le mail en `.msg` dans ce dossier |
| VBA-09 | Envoyer le mail normalement (que l'utilisateur ait dit Oui ou Non) |

## Success Criteria

- [ ] L'exécutable génère un fichier .bas valide avec les paramètres personnalisés
- [ ] Le fichier .bas s'importe correctement dans Outlook VBA
- [ ] À l'envoi d'un mail avec "REF : ... - N°XXXXX", une modale apparaît
- [ ] Si accepté, le mail est sauvegardé dans le bon dossier avec le bon nom
- [ ] Le mail est envoyé normalement dans tous les cas
