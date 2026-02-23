"""
BUMP - Unified VBA Generator
"""

class UnifiedVBAGenerator:
    """Genere le module VBA principal (ThisOutlookSession) unifie."""
    
    def __init__(self, trigram, chrono_file, chrono_folder):
        self.trigram = trigram
        # Replace slashes for VBA path compatibility
        self.chrono_file = chrono_file.replace("/", "\\") if chrono_file else ""
        self.chrono_folder = chrono_folder.replace("/", "\\") if chrono_folder else ""
    
    def get_unified_session_module(self):
        """Retourne le code complet et fusionne pour ThisOutlookSession."""
        template = self._get_unified_template()
        return template.replace(
            "{{CHRONO_FILE}}", self.chrono_file
        ).replace(
            "{{CHRONO_FOLDER}}", self.chrono_folder
        ).replace(
            "{{USER_TRIGRAM}}", self.trigram
        )
    
    def _get_unified_template(self):
        """Template UNIQUE et global pour ThisOutlookSession (Chrono + AR)."""
        return """'===============================================================================
' OUTLOOK TOOL GEN - MODULE UNIFIE (Chrono + AR)
'
' INSTRUCTIONS D'INSTALLATION :
' 1. Copiez l'integralite de ce code.
' 2. Dans Outlook, faites Alt+F11 pour ouvrir l'editeur VBA.
' 3. A gauche, double-cliquez sur "ThisOutlookSession".
' 4. Collez tout ce code dans la fenetre blanche (remplacez tout ce qui existe).
' 5. Sauvegardez (Ctrl+S) et fermez la fenetre.
'===============================================================================

Option Explicit

'===============================================================================
' CONFIGURATION GLOBALE (Generee automatiquement)
'===============================================================================
Private Const CHRONO_FILE As String = "{{CHRONO_FILE}}"
Private Const CHRONO_FOLDER_ROOT As String = "{{CHRONO_FOLDER}}"
Private Const USER_TRIGRAM As String = "{{USER_TRIGRAM}}"

'===============================================================================
' API WINDOWS (Pour le presse-papier)
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
' EVENEMENT PRINCIPAL - DECLENCHE A CHAQUE ENVOI DE MAIL
' Gere a la fois l'archivage Chrono et l'archivage des Accuses de Reception
'===============================================================================
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    On Error Resume Next
    If TypeName(Item) <> "MailItem" Then Exit Sub
    
    ' --- 1. LOGIQUE MODULE ACCUSE RECEPTION (Archivage AR) ---
    Dim prop As Object
    Set prop = Item.UserProperties.Find("ActionAR")
    
    If Not prop Is Nothing Then
        If prop.Value = "True" Then
            ' C'est un AR, on propose l'archivage manuel du dossier
            If MsgBox("Voulez-vous enregistrer une copie (.msg) de cet Accuse de Reception sur le serveur ?", vbQuestion + vbYesNo, "Archivage AR") = vbYes Then
                Dim savePath As String
                savePath = BrowseForFolder()
                If savePath <> "" Then
                    SaveMail Item, savePath, "AR"
                    MsgBox "Accuse de reception archive avec succes.", vbInformation
                End If
            End If
            ' On sort pour ne pas archiver le mail 2 fois s'il contient aussi un numero Chrono
            Exit Sub
        End If
    End If
    
    ' --- 2. LOGIQUE MODULE CHRONO (Classement Auto) ---
    Dim mailBody As String, chronoNum As String, folderPath As String
    
    ' Analyse le debut du mail pour trouver un N° Chrono
    mailBody = Left(Item.Body, 300)
    chronoNum = ExtractChrono(mailBody)
    
    If Len(chronoNum) > 0 Then
        folderPath = FindChronoFolder(chronoNum)
        If Len(folderPath) > 0 Then
            ' Proposer d'archiver (Comportement optionnel selon preference utilisateur, remis ici par defaut)
            If MsgBox("Archiver ce mail dans le dossier Chrono (" & chronoNum & ") ?" & vbCrLf & folderPath, vbQuestion + vbYesNo, "AutoChrono") = vbYes Then
                SaveMail Item, folderPath, "CHRONO"
            End If
        End If
    End If
End Sub


'===============================================================================
' MACROS MANUELLES (A associer a des boutons dans le ruban Outlook)
'===============================================================================

'-------------------------------------------------------------------------------
' 1. MACRO : NOUVEAU CHRONO
'-------------------------------------------------------------------------------
Public Sub NouveauChrono()
    On Error GoTo ErrorHandler
    
    Dim nextChrono As Long
    Dim societe As String, destinataire As String, typeMode As String
    Dim typeDoc As String, numRef As String
    Dim refText As String, folderPath As String
    Dim newChrono As Long
    
    ' 1. Verifier acces Excel
    nextChrono = GetNextChronoNumber()
    If nextChrono = 0 Then
        MsgBox "Impossible de lire le fichier Excel." & vbCrLf & _
               "Chemin : " & CHRONO_FILE, vbExclamation, "Erreur ChronoCreator"
        Exit Sub
    End If
    
    ' 2. Saisie des infos
    societe = InputBox("Nom de la Societe :" & vbCrLf & vbCrLf & _
                       "(Prochain N" & Chr(176) & " : " & nextChrono & ")", "Nouveau Chrono - 1/4")
    If societe = "" Then Exit Sub
    
    destinataire = InputBox("Nom du Destinataire :" & vbCrLf & "(Optionnel)", "Nouveau Chrono - 2/4")
    
    typeMode = InputBox("Type de Document :" & vbCrLf & vbCrLf & _
                        "[P] Proposition" & vbCrLf & _
                        "[R] Rapport", "Nouveau Chrono - 3/4", "P")
    If typeMode = "" Then Exit Sub
    
    If UCase(Left(typeMode, 1)) = "R" Then
        typeDoc = "Rapport"
    Else
        typeDoc = "Proposition"
    End If
    
    numRef = InputBox("Reference du dossier :" & vbCrLf & "(ex: NO60.P.0733)", "Nouveau Chrono - 4/4")
    If numRef = "" Then Exit Sub
    
    ' 3. Confirmation
    If MsgBox("Confirmez-vous la creation ?" & vbCrLf & vbCrLf & _
              "N" & Chr(176) & " : " & nextChrono & vbCrLf & _
              "Societe : " & societe & vbCrLf & _
              "Dest : " & destinataire & vbCrLf & _
              "Type : " & typeDoc & vbCrLf & _
              "Ref : " & numRef, vbQuestion + vbYesNo, "Validation") = vbNo Then Exit Sub
              
    ' 4. Creation de la ligne Excel
    newChrono = AddChronoEntry(societe, destinataire, typeDoc, numRef)
    
    If newChrono > 0 Then
        ' Creation du dossier Reseau
        folderPath = CreateChronoFolder(newChrono, societe)
        
        ' 5. Presse-papier
        refText = "REF : " & USER_TRIGRAM & " - " & numRef & " - N" & Chr(176) & newChrono
        CopyTextToClipboard refText
        
        MsgBox "Succes !" & vbCrLf & vbCrLf & _
               "Dossier cree : " & folderPath & vbCrLf & _
               "Ligne Excel ajoutee." & vbCrLf & _
               "REF copiee dans le presse-papier.", vbInformation, "Chrono N" & Chr(176) & newChrono
    Else
        MsgBox "Erreur lors de l'ecriture dans Excel.", vbCritical
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Erreur inattendue : " & Err.Description, vbCritical
End Sub

'-------------------------------------------------------------------------------
' 2. MACRO : ACCUSE DE RECEPTION AUTOMATIQUE
'-------------------------------------------------------------------------------
Public Sub AccuseReception()
    On Error Resume Next
    
    Dim olItem As Object
    Dim olMail As MailItem
    Dim olReply As MailItem
    
    Dim cpNom As String, cpTel As String
    Dim supNom As String, supTel As String
    Dim dateTravaux As String, dateRendu As String
    Dim sousTraitant As String
    Dim bodyHTML As String
    
    ' 1. Verifier la selection
    Dim objWindow As Object
    Set objWindow = Application.ActiveWindow
    
    If TypeName(objWindow) = "Inspector" Then
        Set olItem = objWindow.CurrentItem
    Else
        If Application.ActiveExplorer.Selection.Count = 0 Then
            MsgBox "Veuillez selectionner un email.", vbExclamation
            Exit Sub
        End If
        Set olItem = Application.ActiveExplorer.Selection.Item(1)
    End If
    
    If TypeName(olItem) <> "MailItem" Then
        MsgBox "L'element selectionne n'est pas un email.", vbExclamation
        Exit Sub
    End If
    Set olMail = olItem
    
    ' 2. Saisie des informations simplifiee
    supNom = InputBox("Nom du Superviseur (laisser vide si aucun) :", "AR Commande - 1/3")
    supTel = "" ' A completer manuellement si besoin
    
    dateTravaux = InputBox("Planning Travaux :" & vbCrLf & "(ex: fin semaine 12)", "AR Commande - 2/3")
    dateRendu = InputBox("Planning Rendu :" & vbCrLf & "(ex: debut fevrier)", "AR Commande - 3/3")
    
    sousTraitant = "" 
    
    ' 3. Generation de la reponse
    If olMail.Sent = False Then
        Set olReply = olMail
    Else
        Set olReply = olMail.Reply
    End If
    
        bodyHTML = "<p>Madame, monsieur,</p>" & _
        "<p>Nous avons bien re&ccedil;u votre commande concernant la mission cit&eacute;e en objet et nous vous en remercions.<br>" & _
        "J'aurai le plaisir d'&ecirc;tre le chef de votre projet.</p>" & _
        "<p>Les principaux intervenants seront les suivants :</p>" & _
        "<table border='1' cellspacing='0' cellpadding='5' style='border-collapse:collapse;'>" & _
        "<tr><td bgcolor='#EEE'><b>Fonction</b></td><td bgcolor='#EEE'><b>Pr&eacute;nom - Nom</b></td><td bgcolor='#EEE'><b>T&eacute;l&eacute;phone</b></td><td bgcolor='#EEE'><b>Courriel</b></td></tr>" & _
        "<tr><td>Superviseur (Supervision generale)</td><td>" & supNom & "</td><td>" & supTel & "</td><td>@groupeginger.com</td></tr>" & _
        "</table>" & _
        "<p><i>Si &eacute;cart par rapport &agrave; l'&eacute;quipe de l'offre justifier pourquoi on change</i></p>" & _
        "<p>Le planning de notre intervention est le suivant :" & _
        "<ul>" & _
        "<li>Nous lan&ccedil;ons d&egrave;s &agrave; pr&eacute;sent les DICT et les recherches historiques.</li>" & _
        "<li>Les travaux sur site seront r&eacute;alis&eacute;s <b>" & dateTravaux & "</b>.</li>" & _
        "<li>Les r&eacute;sultats de l'&eacute;tude vous seront remis <b>" & dateRendu & "</b>.</li>" & _
        "</ul></p>"
        
    If sousTraitant <> "" Then
        bodyHTML = bodyHTML & "<p style='background-color:yellow'>Les sondages (ou) forages (ou) (&agrave; pr&eacute;ciser) .... seront r&eacute;alis&eacute;s par l'entreprise <b>" & sousTraitant & "</b>.</p>"
    End If
    
    bodyHTML = bodyHTML & "<p>Comme pr&eacute;cis&eacute; dans notre offre nous vous remercions de nous adresser en retour :" & _
        "<ul>" & _
        "<li>Le Relev&eacute; Technique Amiante avant Travaux (RAT) ;</li>" & _
        "<li>Le plan des r&eacute;seaux internes.</li>" & _
        "</ul></p>" & _
        "<p>Cordialement,</p>"
    
    olReply.HTMLBody = bodyHTML & olReply.HTMLBody
    
    ' Propriete cachee qui dira au ItemSend de proposer l'archivage AR
    Dim prop As Object
    Set prop = olReply.UserProperties.Add("ActionAR", 1) ' 1 = olText
    prop.Value = "True"
    
    olReply.Save
    olReply.Display
End Sub


'===============================================================================
' FONCTIONS UTILITAIRES INTERNES (Recherche, Archivage, Dossiers)
'===============================================================================

Private Sub SaveMail(ByVal mailItem As Object, ByVal folderPath As String, Optional Context As String = "")
    On Error Resume Next
    Dim fileName As String, fullPath As String, c As Variant
    
    fileName = mailItem.Subject
    If Context = "AR" Then fileName = "AR_" & fileName
    
    ' Nettoyage du nom de fichier
    For Each c In Array("<", ">", ":", Chr(34), "/", "\\", "|", "?", "*")
        fileName = Replace(fileName, c, "_")
    Next c
    
    If Len(fileName) > 100 Then fileName = Left(fileName, 100)
    If Len(fileName) = 0 Then fileName = "mail"
    
    fullPath = folderPath & "\\" & fileName & ".msg"
    mailItem.SaveAs fullPath, 3 ' olMSG
End Sub

Private Function ExtractChrono(ByVal text As String) As String
    Dim pos As Long, i As Long, numStr As String
    pos = InStr(1, text, "N" & Chr(176), vbTextCompare)
    If pos = 0 Then pos = InStr(1, text, "N ", vbTextCompare)
    If pos = 0 Then Exit Function
    
    i = pos + 2
    Do While i <= Len(text) And Mid(text, i, 1) = " ": i = i + 1: Loop
    
    numStr = ""
    Do While i <= Len(text) And IsNumeric(Mid(text, i, 1))
        numStr = numStr & Mid(text, i, 1): i = i + 1
    Loop
    
    If Len(numStr) >= 4 Then ExtractChrono = numStr
End Function

Private Function FindChronoFolder(ByVal chronoNum As String) As String
    On Error Resume Next
    Dim fso As Object, folder As Object, subfolder As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(CHRONO_FOLDER_ROOT) Then Exit Function
    
    Set folder = fso.GetFolder(CHRONO_FOLDER_ROOT)
    For Each subfolder In folder.SubFolders
        If Left(subfolder.Name, Len(chronoNum)) = chronoNum Then
            FindChronoFolder = subfolder.Path
            Exit Function
        End If
    Next
    Set fso = Nothing
End Function

Private Function BrowseForFolder() As String
    On Error Resume Next
    Dim objShell As Object, objFolder As Object
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(0, "Selectionnez le dossier d'archivage :", 0, 0)
    If Not objFolder Is Nothing Then
        BrowseForFolder = objFolder.Self.Path
    End If
    Set objFolder = Nothing
    Set objShell = Nothing
End Function

Public Function GetNextChronoNumber() As Long
    On Error GoTo ErrorHandler
    Dim xlApp As Object, xlWb As Object, xlWs As Object
    Dim row As Long, chronoNum As Long
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False: xlApp.DisplayAlerts = False
    Set xlWb = xlApp.Workbooks.Open(CHRONO_FILE, ReadOnly:=True)
    Set xlWs = xlWb.Sheets(1)
    
    row = 2
    Do While Len(Trim(CStr(xlWs.Cells(row, "B").Value & ""))) > 0
        row = row + 1
        If row > 15000 Then Exit Do
    Loop
    chronoNum = CLng(xlWs.Cells(row, "A").Value)
    
    xlWb.Close False: xlApp.Quit
    Set xlWs = Nothing: Set xlWb = Nothing: Set xlApp = Nothing
    GetNextChronoNumber = chronoNum
    Exit Function
ErrorHandler:
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    GetNextChronoNumber = 0
End Function

Public Function AddChronoEntry(ByVal societe As String, ByVal destinataire As String, ByVal typeDoc As String, ByVal numRef As String) As Long
    On Error GoTo ErrorHandler
    Dim xlApp As Object, xlWb As Object, xlWs As Object
    Dim newRow As Long, newChrono As Long
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False: xlApp.DisplayAlerts = False
    Set xlWb = xlApp.Workbooks.Open(CHRONO_FILE, ReadOnly:=False)
    Set xlWs = xlWb.Sheets(1)
    
    newRow = 2
    Do While Len(Trim(CStr(xlWs.Cells(newRow, "B").Value & ""))) > 0
        newRow = newRow + 1
        If newRow > 15000 Then Exit Do
    Loop
    newChrono = CLng(xlWs.Cells(newRow, "A").Value)
    
    With xlWs
        .Cells(newRow, "B").Value = Date
        .Cells(newRow, "C").Value = societe
        .Cells(newRow, "D").Value = destinataire
        .Cells(newRow, "E").Value = "Mail"
        .Cells(newRow, "F").Value = typeDoc
        .Cells(newRow, "G").Value = numRef
        .Cells(newRow, "J").Value = USER_TRIGRAM
    End With
    
    xlWb.Save
    xlWb.Close False: xlApp.Quit
    Set xlWs = Nothing: Set xlWb = Nothing: Set xlApp = Nothing
    AddChronoEntry = newChrono
    Exit Function
ErrorHandler:
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    AddChronoEntry = 0
End Function

Public Function CreateChronoFolder(ByVal chronoNum As Long, ByVal societe As String) As String
    On Error Resume Next
    Dim fso As Object, folderName As String, fullPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderName = chronoNum & " - " & SanitizeName(societe) & " (" & USER_TRIGRAM & ")"
    fullPath = CHRONO_FOLDER_ROOT & "\\" & folderName
    If Not fso.FolderExists(fullPath) Then fso.CreateFolder fullPath
    Set fso = Nothing
    CreateChronoFolder = fullPath
End Function

Public Function SanitizeName(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("<", ">", ":", Chr(34), "/", "\\", "|", "?", "*")
        s = Replace(s, c, "_")
    Next c
    If Len(s) > 50 Then s = Left(s, 50)
    SanitizeName = s
End Function

Public Sub CopyTextToClipboard(ByVal text As String)
    Dim hMem As LongPtr, pMem As LongPtr, textLen As Long
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
"""
