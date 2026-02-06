# Installation ChronoCreator

## Installation rapide (3 minutes)

### Etape 1 : Importer le module principal
1. Outlook > **Alt+F11** (editeur VBA)
2. **Fichier > Importer fichier** > `ChronoCreator.bas`
3. Configurer les 3 constantes en haut du module

### Etape 2 : Activer l'archivage automatique
1. Double-cliquer sur **ThisOutlookSession** (panneau gauche)
2. Ouvrir `ThisOutlookSession.txt` et copier tout le code
3. Coller dans ThisOutlookSession
4. Configurer `CHRONO_FOLDER`

### Etape 3 : Sauvegarder et redemarrer
1. **Ctrl+S** pour sauvegarder
2. Fermer Outlook completement
3. Relancer Outlook

---

## Configuration

```vba
' Dans ChronoCreator :
Private Const CHRONO_FILE As String = "\\serveur\Chrono 2026.xlsx"
Private Const CHRONO_FOLDER As String = "\\serveur\Chrono"
Private Const USER_TRIGRAM As String = "ABC"

' Dans ThisOutlookSession :
Private Const CHRONO_FOLDER As String = "\\serveur\Chrono"
```

---

## Utilisation

### Creer un nouveau Chrono
- Lancer la macro `NouveauChrono`
- Remplir les 4 champs
- REF copiee automatiquement > Ctrl+V dans le mail

### Archivage automatique
- A l'envoi du mail, si REF detectee dans le corps
- Le .msg est sauvegarde dans le dossier Chrono

### Archivage manuel
- Selectionner un mail
- Lancer la macro `ArchiverMailSelectionne`
