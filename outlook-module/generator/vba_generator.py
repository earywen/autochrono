"""
ChronoCreator Generator - Generateur de module VBA pour Outlook
"""

import os


class ChronoCreatorGenerator:
    """Genere le module VBA ChronoCreator personnalise."""
    
    def __init__(self, trigram, chrono_file, chrono_folder):
        self.trigram = trigram
        self.chrono_file = chrono_file.replace("/", "\\")
        self.chrono_folder = chrono_folder.replace("/", "\\")
    
    def get_main_module(self):
        """Retourne le code du module principal ChronoCreator.bas"""
        return self._get_main_template().replace(
            "{{CHRONO_FILE}}", self.chrono_file
        ).replace(
            "{{CHRONO_FOLDER}}", self.chrono_folder
        ).replace(
            "{{USER_TRIGRAM}}", self.trigram
        )
    
    def get_session_module(self):
        """Retourne le code pour ThisOutlookSession"""
        return self._get_session_template().replace(
            "{{CHRONO_FOLDER}}", self.chrono_folder
        )
    
    def _get_main_template(self):
        """Template ChronoCreator.bas"""
        return """'===============================================================================
' ChronoCreator - Creation de nouveaux numeros Chrono depuis Outlook
' Import : Alt+F11 > Fichier > Importer ce fichier
'===============================================================================

Option Explicit

'===============================================================================
' CONFIGURATION
'===============================================================================
Private Const CHRONO_FILE As String = "{{CHRONO_FILE}}"
Private Const CHRONO_FOLDER As String = "{{CHRONO_FOLDER}}"
Private Const USER_TRIGRAM As String = "{{USER_TRIGRAM}}"

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
' MACRO PRINCIPALE
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
    
    nextChrono = GetNextChronoNumber()
    If nextChrono = 0 Then
        MsgBox "Impossible de lire le fichier Excel." & vbCrLf & _
               "Verifiez le chemin : " & CHRONO_FILE, vbExclamation, "Erreur"
        Exit Sub
    End If
    
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
    
    If MsgBox("Creer ce Chrono ?" & vbCrLf & vbCrLf & _
              "N" & Chr(176) & " : " & nextChrono & vbCrLf & _
              "Societe : " & societe & vbCrLf & _
              "Destinataire : " & destinataire & vbCrLf & _
              "Type : " & typeDoc & vbCrLf & _
              "Reference : " & numRef, _
              vbYesNo + vbQuestion, "Confirmation") = vbNo Then
        Exit Sub
    End If
    
    newChrono = AddChronoEntry(societe, destinataire, typeDoc, numRef)
    
    If newChrono > 0 Then
        folderPath = CreateChronoFolder(newChrono, societe)
        refText = "REF : " & USER_TRIGRAM & " - " & numRef & " - N" & Chr(176) & newChrono
        CopyTextToClipboard refText
        
        MsgBox "Chrono N" & Chr(176) & newChrono & " cree !" & vbCrLf & vbCrLf & _
               "Ligne ajoutee dans Excel" & vbCrLf & _
               "Dossier cree : " & folderPath & vbCrLf & vbCrLf & _
               "REF copiee : " & refText, vbInformation, "Succes"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur : " & Err.Description, vbExclamation, "ChronoCreator"
End Sub

Private Function GetNextChronoNumber() As Long
    On Error GoTo ErrorHandler
    
    Dim xlApp As Object, xlWb As Object, xlWs As Object
    Dim row As Long, chronoNum As Long
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    
    Set xlWb = xlApp.Workbooks.Open(CHRONO_FILE, ReadOnly:=True)
    Set xlWs = xlWb.Sheets(1)
    
    row = 2
    Do While Len(Trim(CStr(xlWs.Cells(row, "B").Value & ""))) > 0
        row = row + 1
        If row > 15000 Then Exit Do
    Loop
    
    chronoNum = CLng(xlWs.Cells(row, "A").Value)
    
    xlWb.Close False
    xlApp.Quit
    Set xlWs = Nothing: Set xlWb = Nothing: Set xlApp = Nothing
    
    GetNextChronoNumber = chronoNum
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    GetNextChronoNumber = 0
End Function

Private Function AddChronoEntry(ByVal societe As String, ByVal destinataire As String, ByVal typeDoc As String, ByVal numRef As String) As Long
    On Error GoTo ErrorHandler
    
    Dim xlApp As Object, xlWb As Object, xlWs As Object
    Dim newRow As Long, newChrono As Long
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    
    Set xlWb = xlApp.Workbooks.Open(CHRONO_FILE, ReadOnly:=False)
    Set xlWs = xlWb.Sheets(1)
    
    newRow = 2
    Do While Len(Trim(CStr(xlWs.Cells(newRow, "B").Value & ""))) > 0
        newRow = newRow + 1
        If newRow > 15000 Then Exit Do
    Loop
    
    newChrono = CLng(xlWs.Cells(newRow, "A").Value)
    
    xlWs.Cells(newRow, "B").Value = Date
    xlWs.Cells(newRow, "C").Value = societe
    xlWs.Cells(newRow, "D").Value = destinataire
    xlWs.Cells(newRow, "E").Value = "Mail"
    xlWs.Cells(newRow, "F").Value = typeDoc
    xlWs.Cells(newRow, "G").Value = numRef
    xlWs.Cells(newRow, "J").Value = USER_TRIGRAM
    
    xlWb.Save
    xlWb.Close False
    xlApp.Quit
    Set xlWs = Nothing: Set xlWb = Nothing: Set xlApp = Nothing
    
    AddChronoEntry = newChrono
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    MsgBox "Erreur Excel : " & Err.Description, vbExclamation, "Erreur"
    AddChronoEntry = 0
End Function

Private Function CreateChronoFolder(ByVal chronoNum As Long, ByVal societe As String) As String
    On Error Resume Next
    Dim fso As Object, folderName As String, fullPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderName = chronoNum & " - " & SanitizeName(societe) & " (" & USER_TRIGRAM & ")"
    fullPath = CHRONO_FOLDER & "\\" & folderName
    If Not fso.FolderExists(fullPath) Then fso.CreateFolder fullPath
    Set fso = Nothing
    CreateChronoFolder = fullPath
End Function

Private Function SanitizeName(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("<", ">", ":", Chr(34), "/", "\\", "|", "?", "*")
        s = Replace(s, c, "_")
    Next c
    If Len(s) > 50 Then s = Left(s, 50)
    SanitizeName = s
End Function

Private Sub CopyTextToClipboard(ByVal text As String)
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

Public Function ExtractChronoFromRef(ByVal text As String) As String
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
    If Len(numStr) >= 4 Then ExtractChronoFromRef = numStr
End Function

Public Function FindChronoFolder(ByVal chronoNum As String) As String
    On Error Resume Next
    Dim fso As Object, folder As Object, subfolder As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(CHRONO_FOLDER) Then Exit Function
    Set folder = fso.GetFolder(CHRONO_FOLDER)
    For Each subfolder In folder.SubFolders
        If Left(subfolder.Name, Len(chronoNum)) = chronoNum Then
            FindChronoFolder = subfolder.Path: Exit Function
        End If
    Next
    Set fso = Nothing
End Function

Public Sub SaveMailToFolder(ByVal mailItem As Object, ByVal folderPath As String)
    On Error Resume Next
    Dim fileName As String, fullPath As String
    fileName = SanitizeName(mailItem.Subject)
    If Len(fileName) = 0 Then fileName = "mail"
    fullPath = folderPath & "\\" & fileName & ".msg"
    mailItem.SaveAs fullPath, 3
End Sub

Public Sub ArchiverMailSelectionne()
    On Error GoTo ErrorHandler
    Dim mailItem As Object, chronoNum As String, folderPath As String
    Set mailItem = Application.ActiveExplorer.Selection.Item(1)
    If TypeName(mailItem) <> "MailItem" Then
        MsgBox "Veuillez selectionner un mail.", vbExclamation: Exit Sub
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
