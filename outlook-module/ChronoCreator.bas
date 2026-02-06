'===============================================================================
' ChronoCreator - Creation de nouveaux numeros Chrono depuis Outlook
' Import unique : Alt+F11 > Fichier > Importer ce fichier > C'est pret !
'===============================================================================

Option Explicit

'===============================================================================
' CONFIGURATION - MODIFIER CES VALEURS AVANT UTILISATION
'===============================================================================
Private Const CHRONO_FILE As String = "\\serveur\partage\Chrono 2026.xlsx"
Private Const CHRONO_FOLDER As String = "\\serveur\partage\Chrono"
Private Const USER_TRIGRAM As String = "XXX"  ' Votre trigramme (ex: ARME)

'===============================================================================
' API Windows pour le presse-papier
'===============================================================================
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As String, ByVal Length As Long)

Private Const CF_TEXT = 1
Private Const GMEM_MOVEABLE = &H2

'===============================================================================
' MACRO PRINCIPALE - Assigner a un bouton ou raccourci clavier
'===============================================================================
Public Sub NouveauChrono()
    On Error GoTo ErrorHandler
    
    Dim nextChrono As Long
    Dim societe As String
    Dim destinataire As String
    Dim typeDoc As String
    Dim numRef As String
    Dim refText As String
    Dim newChrono As Long
    Dim folderPath As String
    Dim mailItem As Object
    
    ' Recuperer le prochain numero
    nextChrono = GetNextChronoNumber()
    If nextChrono = 0 Then
        MsgBox "Impossible de lire le fichier Excel." & vbCrLf & _
               "Verifiez le chemin : " & CHRONO_FILE, vbExclamation, "Erreur"
        Exit Sub
    End If
    
    ' Saisie des informations
    societe = InputBox("Nom de la societe :" & vbCrLf & vbCrLf & _
                       "(Prochain N" & Chr(176) & " Chrono : " & nextChrono & ")", _
                       "Nouveau Chrono - Etape 1/4")
    If societe = "" Then Exit Sub
    
    destinataire = InputBox("Nom du destinataire :" & vbCrLf & _
                            "(Prenom NOM)", _
                            "Nouveau Chrono - Etape 2/4")
    
    typeDoc = InputBox("Type de document :" & vbCrLf & vbCrLf & _
                       "Tapez P pour Proposition" & vbCrLf & _
                       "Tapez R pour Rapport", _
                       "Nouveau Chrono - Etape 3/4", "P")
    If typeDoc = "" Then Exit Sub
    If UCase(Left(typeDoc, 1)) = "R" Then
        typeDoc = "Rapport"
    Else
        typeDoc = "Proposition"
    End If
    
    numRef = InputBox("Numero de reference :" & vbCrLf & _
                      "(ex: NO60.P.0733)", _
                      "Nouveau Chrono - Etape 4/4")
    If numRef = "" Then Exit Sub
    
    ' Confirmation
    If MsgBox("Creer ce Chrono ?" & vbCrLf & vbCrLf & _
              "N" & Chr(176) & " : " & nextChrono & vbCrLf & _
              "Societe : " & societe & vbCrLf & _
              "Destinataire : " & destinataire & vbCrLf & _
              "Type : " & typeDoc & vbCrLf & _
              "Reference : " & numRef, _
              vbYesNo + vbQuestion, "Confirmation") = vbNo Then
        Exit Sub
    End If
    
    ' Creer l'entree Excel
    newChrono = AddChronoEntry(societe, destinataire, typeDoc, numRef)
    
    If newChrono > 0 Then
        ' Creer le dossier Chrono
        folderPath = CreateChronoFolder(newChrono, societe)
        
        ' Generer le texte REF
        refText = "REF : " & USER_TRIGRAM & " - " & numRef & " - N" & Chr(176) & newChrono
        
        ' Copier dans le presse-papier
        CopyTextToClipboard refText
        
        MsgBox "Chrono N" & Chr(176) & newChrono & " cree avec succes !" & vbCrLf & vbCrLf & _
               "Ligne ajoutee dans Excel" & vbCrLf & _
               "Dossier cree : " & folderPath & vbCrLf & vbCrLf & _
               "REF copiee : " & refText & vbCrLf & vbCrLf & _
               "Collez dans votre mail (Ctrl+V)", vbInformation, "Succes"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur : " & Err.Description & vbCrLf & _
           "Ligne : " & Erl, vbExclamation, "ChronoCreator"
End Sub

'===============================================================================
' Recupere le prochain numero de Chrono disponible
'===============================================================================
Private Function GetNextChronoNumber() As Long
    On Error GoTo ErrorHandler
    
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
    Dim row As Long
    Dim chronoNum As Long
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    
    Set xlWb = xlApp.Workbooks.Open(CHRONO_FILE, ReadOnly:=True)
    Set xlWs = xlWb.Sheets(1)
    
    ' Trouver la premiere ligne sans date (colonne B vide)
    row = 2
    Do While Len(Trim(CStr(xlWs.Cells(row, "B").Value & ""))) > 0
        row = row + 1
        If row > 15000 Then Exit Do
    Loop
    
    chronoNum = CLng(xlWs.Cells(row, "A").Value)
    
    xlWb.Close False
    xlApp.Quit
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    
    GetNextChronoNumber = chronoNum
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    GetNextChronoNumber = 0
End Function

'===============================================================================
' Ajoute une nouvelle entree dans le fichier Excel
'===============================================================================
Private Function AddChronoEntry(ByVal societe As String, _
                                ByVal destinataire As String, _
                                ByVal typeDoc As String, _
                                ByVal numRef As String) As Long
    On Error GoTo ErrorHandler
    
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
    Dim newRow As Long
    Dim newChrono As Long
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    
    ' Ouvrir en ecriture
    Set xlWb = xlApp.Workbooks.Open(CHRONO_FILE, ReadOnly:=False)
    Set xlWs = xlWb.Sheets(1)
    
    ' Trouver la premiere ligne sans date
    newRow = 2
    Do While Len(Trim(CStr(xlWs.Cells(newRow, "B").Value & ""))) > 0
        newRow = newRow + 1
        If newRow > 15000 Then Exit Do
    Loop
    
    newChrono = CLng(xlWs.Cells(newRow, "A").Value)
    
    ' Ecrire les donnees
    xlWs.Cells(newRow, "B").Value = Date
    xlWs.Cells(newRow, "C").Value = societe
    xlWs.Cells(newRow, "D").Value = destinataire
    xlWs.Cells(newRow, "E").Value = "Mail"
    xlWs.Cells(newRow, "F").Value = typeDoc
    xlWs.Cells(newRow, "G").Value = numRef
    xlWs.Cells(newRow, "J").Value = USER_TRIGRAM
    
    ' Sauvegarder
    xlWb.Save
    xlWb.Close False
    xlApp.Quit
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    
    AddChronoEntry = newChrono
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    MsgBox "Erreur Excel : " & Err.Description & vbCrLf & vbCrLf & _
           "Le fichier est peut-etre ouvert en ecriture par quelqu'un.", vbExclamation, "Erreur"
    AddChronoEntry = 0
End Function

'===============================================================================
' Cree le dossier Chrono sur le serveur
'===============================================================================
Private Function CreateChronoFolder(ByVal chronoNum As Long, ByVal societe As String) As String
    On Error Resume Next
    
    Dim fso As Object
    Dim folderName As String
    Dim fullPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Format : 11069 - SOCIETE (XXX)
    folderName = chronoNum & " - " & SanitizeName(societe) & " (" & USER_TRIGRAM & ")"
    fullPath = CHRONO_FOLDER & "\" & folderName
    
    If Not fso.FolderExists(fullPath) Then
        fso.CreateFolder fullPath
    End If
    
    Set fso = Nothing
    CreateChronoFolder = fullPath
End Function

'===============================================================================
' Nettoie le nom pour le systeme de fichiers
'===============================================================================
Private Function SanitizeName(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("<", ">", ":", """", "/", "\", "|", "?", "*")
        s = Replace(s, c, "_")
    Next c
    If Len(s) > 50 Then s = Left(s, 50)
    SanitizeName = s
End Function

'===============================================================================
' Copie le texte dans le presse-papier (API Windows)
'===============================================================================
Private Sub CopyTextToClipboard(ByVal text As String)
    Dim hMem As LongPtr
    Dim pMem As LongPtr
    Dim textLen As Long
    
    textLen = Len(text) + 1
    hMem = GlobalAlloc(GMEM_MOVEABLE, textLen)
    If hMem = 0 Then Exit Sub
    
    pMem = GlobalLock(hMem)
    If pMem = 0 Then Exit Sub
    
    CopyMemory pMem, text, textLen
    GlobalUnlock hMem
    
    If OpenClipboard(0) <> 0 Then
        EmptyClipboard
        SetClipboardData CF_TEXT, hMem
        CloseClipboard
    End If
End Sub

'===============================================================================
' CODE A COPIER DANS ThisOutlookSession (Alt+F11 > ThisOutlookSession)
' Ceci intercepte l'envoi des mails et archive automatiquement
'===============================================================================

'-------------------------------------------------------------------------------
' COPIER TOUT CE QUI SUIT DANS ThisOutlookSession :
'-------------------------------------------------------------------------------
'
' Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
'     On Error Resume Next
'     
'     If TypeName(Item) <> "MailItem" Then Exit Sub
'     
'     Dim mailBody As String
'     Dim chronoNum As String
'     Dim folderPath As String
'     
'     mailBody = Left(Item.Body, 200)
'     
'     ' Chercher le pattern REF avec numero Chrono
'     chronoNum = ExtractChronoFromRef(mailBody)
'     
'     If Len(chronoNum) > 0 Then
'         folderPath = FindChronoFolder(chronoNum)
'         If Len(folderPath) > 0 Then
'             SaveMailToFolder Item, folderPath
'         End If
'     End If
' End Sub
'
'-------------------------------------------------------------------------------

'===============================================================================
' Extrait le numero Chrono depuis le texte REF
' Cherche "N" suivi du symbole degre et de chiffres
'===============================================================================
Public Function ExtractChronoFromRef(ByVal text As String) As String
    Dim pos As Long
    Dim i As Long
    Dim numStr As String
    
    ' Chercher "N" suivi de caracteres puis chiffres
    pos = InStr(1, text, "N" & Chr(176), vbTextCompare)
    If pos = 0 Then pos = InStr(1, text, "N ", vbTextCompare)
    If pos = 0 Then Exit Function
    
    i = pos + 2
    
    ' Sauter les espaces
    Do While i <= Len(text) And Mid(text, i, 1) = " "
        i = i + 1
    Loop
    
    ' Extraire les chiffres
    numStr = ""
    Do While i <= Len(text) And IsNumeric(Mid(text, i, 1))
        numStr = numStr & Mid(text, i, 1)
        i = i + 1
    Loop
    
    If Len(numStr) >= 4 Then
        ExtractChronoFromRef = numStr
    End If
End Function

'===============================================================================
' Trouve le dossier Chrono correspondant au numero
'===============================================================================
Public Function FindChronoFolder(ByVal chronoNum As String) As String
    On Error Resume Next
    
    Dim fso As Object
    Dim folder As Object
    Dim subfolder As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(CHRONO_FOLDER) Then Exit Function
    
    Set folder = fso.GetFolder(CHRONO_FOLDER)
    
    ' Chercher un dossier commencant par le numero
    For Each subfolder In folder.SubFolders
        If Left(subfolder.Name, Len(chronoNum)) = chronoNum Then
            FindChronoFolder = subfolder.Path
            Exit Function
        End If
    Next
    
    Set fso = Nothing
End Function

'===============================================================================
' Sauvegarde le mail dans le dossier
'===============================================================================
Public Sub SaveMailToFolder(ByVal mailItem As Object, ByVal folderPath As String)
    On Error Resume Next
    
    Dim fileName As String
    Dim fullPath As String
    
    fileName = SanitizeName(mailItem.Subject)
    If Len(fileName) = 0 Then fileName = "mail"
    
    fullPath = folderPath & "\" & fileName & ".msg"
    
    ' Sauvegarder en format MSG
    mailItem.SaveAs fullPath, 3  ' olMSG = 3
    
    ' Message de confirmation (optionnel - commenter si non desire)
    ' MsgBox "Mail archive dans :" & vbCrLf & fullPath, vbInformation, "ChronoCreator"
End Sub

'===============================================================================
' Macro pour archiver manuellement le mail selectionne
'===============================================================================
Public Sub ArchiverMailSelectionne()
    On Error GoTo ErrorHandler
    
    Dim mailItem As Object
    Dim chronoNum As String
    Dim folderPath As String
    
    Set mailItem = Application.ActiveExplorer.Selection.Item(1)
    
    If TypeName(mailItem) <> "MailItem" Then
        MsgBox "Veuillez selectionner un mail.", vbExclamation
        Exit Sub
    End If
    
    chronoNum = ExtractChronoFromRef(Left(mailItem.Body, 200))
    
    If Len(chronoNum) = 0 Then
        chronoNum = InputBox("Numero Chrono non detecte." & vbCrLf & _
                             "Entrez le numero manuellement :", "Archiver Mail")
        If Len(chronoNum) = 0 Then Exit Sub
    End If
    
    folderPath = FindChronoFolder(chronoNum)
    
    If Len(folderPath) = 0 Then
        MsgBox "Dossier Chrono " & chronoNum & " non trouve.", vbExclamation
        Exit Sub
    End If
    
    SaveMailToFolder mailItem, folderPath
    MsgBox "Mail archive dans :" & vbCrLf & folderPath, vbInformation, "Succes"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur : " & Err.Description, vbExclamation
End Sub

