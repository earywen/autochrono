"""
ChronoCreator Generator - Generateur de module VBA pour Outlook
"""

import os


class ChronoCreatorGenerator:
    """Genere le module VBA personnalise selon le mode choisi."""
    
    def __init__(self, trigram, chrono_file, chrono_folder, user_name="", user_phone=""):
        self.trigram = trigram
        # Replace slashes for VBA path compatibility
        self.chrono_file = chrono_file.replace("/", "\\") if chrono_file else ""
        self.chrono_folder = chrono_folder.replace("/", "\\") if chrono_folder else ""
        self.user_name = user_name
        self.user_phone = user_phone
    
    def get_chrono_module(self):
        """Retourne le code pour le module ChronoCreator (standalone)."""
        return self._get_chrono_template().replace(
            "{{CHRONO_FILE}}", self.chrono_file
        ).replace(
            "{{CHRONO_FOLDER}}", self.chrono_folder
        ).replace(
            "{{USER_TRIGRAM}}", self.trigram
        )

    def get_ar_module(self):
        """Retourne le code pour le module AccuseReception (standalone)."""
        return self._get_ar_template().replace(
            "{{USER_NAME}}", self.user_name
        ).replace(
            "{{USER_PHONE}}", self.user_phone
        )
    
    def get_session_module(self):
        """Retourne le code pour ThisOutlookSession."""
        return self._get_session_template().replace(
            "{{CHRONO_FOLDER}}", self.chrono_folder
        )
    
    def _get_chrono_template(self):
        """Template pour ChronoCreator uniquement."""
        return """'===============================================================================
' ChronoCreator - Creation de nouveaux numeros Chrono
'===============================================================================
Option Explicit

' CONFIGURATION
Public Const CHRONO_FILE As String = "{{CHRONO_FILE}}"
Public Const CHRONO_FOLDER As String = "{{CHRONO_FOLDER}}"
Public Const USER_TRIGRAM As String = "{{USER_TRIGRAM}}"

' API PRESSE-PAPIER
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

' MACRO : NOUVEAU CHRONO
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
               "Chemin : " & CHRONO_FILE, vbExclamation, "Erreur"
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
              
    ' 4. Creation
    newChrono = AddChronoEntry(societe, destinataire, typeDoc, numRef)
    
    If newChrono > 0 Then
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

' FONCTIONS UTILITAIRES
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
    fullPath = CHRONO_FOLDER & "\\" & folderName
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

    def _get_ar_template(self):
        """Template pour Accuse Reception uniquement."""
        return """'===============================================================================
' OutlookTools - Accuse de Reception
'===============================================================================
Option Explicit

Public Const DEFAULT_USER_NAME As String = "{{USER_NAME}}"
Public Const DEFAULT_USER_PHONE As String = "{{USER_PHONE}}"

' MACRO : ACCUSE DE RECEPTION
Public Sub AccuseReception()
    On Error Resume Next
    
    Dim olItem As Object
    Dim olMail As MailItem
    Dim olReply As MailItem
    Dim destFolder As Folder
    
    ' Variables du modele
    Dim cpNom As String, cpTel As String
    Dim supNom As String, supTel As String
    Dim dateTravaux As String, dateRendu As String
    Dim sousTraitant As String
    Dim bodyHTML As String
    
    ' 1. Verifier la selection
    Set olItem = Application.ActiveExplorer.Selection.Item(1)
    If olItem.Class <> olMail Then
        MsgBox "Veuillez selectionner un email.", vbExclamation
        Exit Sub
    End If
    Set olMail = olItem
    
    ' 2. Saisie des informations
    cpNom = InputBox("Nom du Chef de Projet :", "AR Commande - 1/5", DEFAULT_USER_NAME)
    If cpNom = "" Then Exit Sub
    
    cpTel = InputBox("Telephone ligne directe :", "AR Commande - 2/5", DEFAULT_USER_PHONE)
    
    supNom = InputBox("Nom du Superviseur :", "AR Commande - 3/5")
    If supNom = "" Then supNom = "[A DEFINIR]"
    
    supTel = InputBox("Telephone Standard/Superviseur :", "AR Commande - 4/5")
    
    dateTravaux = InputBox("Planning Travaux :" & vbCrLf & "(ex: fin semaine 12)", "AR Commande - 5/5")
    dateRendu = InputBox("Planning Rendu :" & vbCrLf & "(ex: debut fevrier)", "AR Commande - 5/5")
    sousTraitant = InputBox("Sous-traitant / Intervenant terrain (Optionnel) :" & vbCrLf & "(ex: l'entreprise de sondage XXXX)", "AR Commande")
    
    ' 3. Rangement automatique
    If MsgBox("Voulez-vous definir le dossier de rangement maintenant ?" & vbCrLf & _
              "Le mail envoye (AR) y sera range automatiquement.", vbQuestion + vbYesNo, "Rangement") = vbYes Then
        Set destFolder = Application.Session.PickFolder
    End If
    
    ' 4. Generation de la reponse
    Set olReply = olMail.Reply
    If Not destFolder Is Nothing Then
        Set olReply.SaveSentSentMessageFolder = destFolder
    End If
    
    ' Construction du corps HTML
    bodyHTML = "<p>Madame, monsieur,</p>" & _
        "<p>Nous avons bien recu votre commande concernant la mission citee en objet et nous vous en remercions.<br>" & _
        "J'aurai le plaisir d'etre le chef de votre projet.</p>" & _
        "<p>Les principaux intervenants seront les suivants :</p>" & _
        "<table border='1' cellspacing='0' cellpadding='5' style='border-collapse:collapse;'>" & _
        "<tr><td bgcolor='#EEE'><b>Fonction</b></td><td bgcolor='#EEE'><b>Prenom - Nom</b></td><td bgcolor='#EEE'><b>Telephone</b></td><td bgcolor='#EEE'><b>Courriel</b></td></tr>" & _
        "<tr><td>Chef du projet</td><td>" & cpNom & "</td><td>" & cpTel & "</td><td>@groupeginger.com</td></tr>" & _
        "<tr><td>Superviseur</td><td>" & supNom & "</td><td>" & supTel & "</td><td>@groupeginger.com</td></tr>" & _
        "</table>" & _
        "<p><i>Si ecart par rapport a l'equipe de l'offre justifier pourquoi on change</i></p>" & _
        "<p>Le planning de notre intervention est le suivant :" & _
        "<ul>" & _
        "<li>Nous lancons des a present les DICT et les recherches historiques.</li>" & _
        "<li>Les travaux sur site seront realises <b>" & dateTravaux & "</b>.</li>" & _
        "<li>Les resultats de l'etude vous seront remis <b>" & dateRendu & "</b>.</li>" & _
        "</ul></p>"
        
    If sousTraitant <> "" Then
        bodyHTML = bodyHTML & "<p style='background-color:yellow'>Les sondages (ou) forages (ou) (a preciser) .... seront realises par l'entreprise <b>" & sousTraitant & "</b>.</p>"
    End If
    
    bodyHTML = bodyHTML & "<p>Comme precise dans notre offre nous vous remercions de nous adresser en retour :" & _
        "<ul>" & _
        "<li>Le Releve Technique Amiante avant Travaux (RAT) ;</li>" & _
        "<li>Le plan des reseaux internes.</li>" & _
        "</ul></p>" & _
        "<p>Cordialement,</p>"
    
    olReply.HTMLBody = bodyHTML & olReply.HTMLBody
    olReply.Display
    
End Sub
    End If
    chronoNum = ExtractChronoFromRef(Left(mailItem.Body, 200))
    If Len(chronoNum) = 0 Then
        chronoNum = InputBox("Numero Chrono :", "Archiver Mail")
        If Len(chronoNum) = 0 Then Exit Sub
    End If
    folderPath = FindChronoFolder(chronoNum)
    If Len(folderPath) = 0 Then
        MsgBox "Dossier Chrono " & chronoNum & " non trouve.", vbExclamation: Exit Sub
    End If
    SaveMailToFolder mailItem, folderPath
    MsgBox "Mail archive dans :" & vbCrLf & folderPath, vbInformation, "Succes"
    Exit Sub
ErrorHandler:
    MsgBox "Erreur : " & Err.Description, vbExclamation
End Sub
"""
    
    def _get_session_template(self):
        """Template ThisOutlookSession"""
        return """'===============================================================================
' CODE POUR ThisOutlookSession
' Alt+F11 > Double-cliquer sur ThisOutlookSession > Coller ce code
'===============================================================================

Option Explicit

Private Const CHRONO_FOLDER As String = "{{CHRONO_FOLDER}}"

Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    On Error Resume Next
    If TypeName(Item) <> "MailItem" Then Exit Sub
    
    Dim mailBody As String, chronoNum As String, folderPath As String
    mailBody = Left(Item.Body, 300)
    chronoNum = ExtractChrono(mailBody)
    
    If Len(chronoNum) > 0 Then
        folderPath = FindFolder(chronoNum)
        If Len(folderPath) > 0 Then SaveMail Item, folderPath
    End If
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

Private Function FindFolder(ByVal chronoNum As String) As String
    On Error Resume Next
    Dim fso As Object, folder As Object, subfolder As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(CHRONO_FOLDER) Then Exit Function
    Set folder = fso.GetFolder(CHRONO_FOLDER)
    For Each subfolder In folder.SubFolders
        If Left(subfolder.Name, Len(chronoNum)) = chronoNum Then
            FindFolder = subfolder.Path: Exit Function
        End If
    Next
    Set fso = Nothing
End Function

Private Sub SaveMail(ByVal mailItem As Object, ByVal folderPath As String)
    On Error Resume Next
    Dim fileName As String, fullPath As String, c As Variant
    fileName = mailItem.Subject
    For Each c In Array("<", ">", ":", Chr(34), "/", "\\", "|", "?", "*")
        fileName = Replace(fileName, c, "_")
    Next c
    If Len(fileName) > 50 Then fileName = Left(fileName, 50)
    If Len(fileName) = 0 Then fileName = "mail"
    fullPath = folderPath & "\\" & fileName & ".msg"
    mailItem.SaveAs fullPath, 3
End Sub
"""
