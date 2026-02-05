"""
AutoChrono - Generateur de code VBA
"""

import os


class VBAGenerator:
    """Genere le module VBA personnalise pour Outlook."""
    
    def __init__(self, trigram, chrono_folder, chrono_file, col_chrono, col_client, col_trigram):
        self.trigram = trigram
        self.chrono_folder = chrono_folder.replace("/", "\\")
        self.chrono_file = chrono_file.replace("/", "\\")
        self.col_chrono = col_chrono
        self.col_client = col_client
        self.col_trigram = col_trigram
    
    def generate(self, output_path):
        """Genere le fichier .bas avec les valeurs personnalisees."""
        vba_code = self.get_code()
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(vba_code)
    
    def get_code(self):
        """Retourne le code VBA formate (pour copier dans le presse-papier)."""
        vba_code = self._get_vba_template()
        vba_code = vba_code.replace("{{TRIGRAM}}", self.trigram)
        vba_code = vba_code.replace("{{CHRONO_FOLDER}}", self.chrono_folder)
        vba_code = vba_code.replace("{{CHRONO_FILE}}", self.chrono_file)
        vba_code = vba_code.replace("{{COL_CHRONO}}", self.col_chrono)
        vba_code = vba_code.replace("{{COL_CLIENT}}", self.col_client)
        vba_code = vba_code.replace("{{COL_TRIGRAM}}", self.col_trigram)
        return vba_code
    
    def _get_vba_template(self):
        """Retourne le template VBA (ASCII pur pour eviter problemes encodage)."""
        return """'===============================================================================
' AutoChrono - Archivage automatique des mails Chrono
' Collez ce code dans ThisOutlookSession (Alt+F11)
'===============================================================================

Option Explicit

' Configuration
Private Const USER_TRIGRAM As String = "{{TRIGRAM}}"
Private Const CHRONO_FOLDER As String = "{{CHRONO_FOLDER}}"
Private Const CHRONO_FILE As String = "{{CHRONO_FILE}}"
Private Const COL_CHRONO As String = "{{COL_CHRONO}}"
Private Const COL_CLIENT As String = "{{COL_CLIENT}}"
Private Const COL_TRIGRAM As String = "{{COL_TRIGRAM}}"

'===============================================================================
' Evenement avant l'envoi d'un mail
'===============================================================================
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    On Error GoTo ErrorHandler
    
    Dim mailBody As String
    Dim chronoNumber As String
    
    If TypeName(Item) <> "MailItem" Then Exit Sub
    
    mailBody = Left(Item.Body, 150)
    
    ' Chercher REF dans le corps du mail
    If InStr(1, mailBody, "REF", vbTextCompare) > 0 Then
        chronoNumber = ExtractChronoNumber(mailBody)
        
        If Len(chronoNumber) > 0 Then
            If MsgBox("Archiver ce mail dans le dossier Chrono ?" & vbCrLf & vbCrLf & _
                      "Chrono d" & Chr(233) & "tect" & Chr(233) & " : " & chronoNumber, _
                      vbYesNo + vbQuestion, "AutoChrono") = vbYes Then
                Call ArchiveMail(Item, chronoNumber)
            End If
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur AutoChrono : " & Err.Description, vbExclamation, "AutoChrono"
End Sub

'===============================================================================
' Extrait le numero de Chrono (cherche N suivi de 4+ chiffres)
'===============================================================================
Private Function ExtractChronoNumber(ByVal text As String) As String
    Dim pos As Long
    Dim numStr As String
    Dim i As Long
    
    pos = InStr(1, text, "N", vbTextCompare)
    
    Do While pos > 0
        i = pos + 1
        ' Sauter 1-2 caracteres non numeriques (comme le symbole degre)
        If i <= Len(text) And Not IsNumeric(Mid(text, i, 1)) Then i = i + 1
        If i <= Len(text) And Not IsNumeric(Mid(text, i, 1)) Then i = i + 1
        
        numStr = ""
        Do While i <= Len(text) And IsNumeric(Mid(text, i, 1))
            numStr = numStr & Mid(text, i, 1)
            i = i + 1
        Loop
        
        ' Chrono = au moins 4 chiffres
        If Len(numStr) >= 4 Then
            ExtractChronoNumber = numStr
            Exit Function
        End If
        
        pos = InStr(pos + 1, text, "N", vbTextCompare)
    Loop
    
    ExtractChronoNumber = ""
End Function

'===============================================================================
' Archive le mail dans le dossier Chrono
'===============================================================================
Private Sub ArchiveMail(ByVal mailItem As Object, ByVal chronoNumber As String)
    On Error GoTo ErrorHandler
    
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
    Dim foundCell As Object
    Dim targetRow As Long
    Dim clientName As String
    Dim trigramme As String
    Dim folderName As String
    Dim folderPath As String
    Dim fso As Object
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    
    Set xlWb = xlApp.Workbooks.Open(CHRONO_FILE, ReadOnly:=True)
    Set xlWs = xlWb.Sheets(1)
    
    ' Recherche exacte du numero
    Set foundCell = xlWs.Columns(COL_CHRONO).Find(What:=chronoNumber, LookIn:=-4163, LookAt:=1)
    
    If foundCell Is Nothing Then
        MsgBox "Chrono " & chronoNumber & " non trouv" & Chr(233) & " dans Excel.", vbExclamation, "AutoChrono"
        GoTo CleanExit
    End If
    
    targetRow = foundCell.Row
    clientName = Trim(CStr(xlWs.Range(COL_CLIENT & targetRow).Value))
    trigramme = Trim(CStr(xlWs.Range(COL_TRIGRAM & targetRow).Value))
    
    If clientName = "" Or clientName = "0" Then clientName = "Client_Inconnu"
    If trigramme = "" Or trigramme = "0" Then trigramme = "Inc"
    
    xlWb.Close False
    xlApp.Quit
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    
    folderName = chronoNumber & " " & clientName & " (" & trigramme & ")"
    folderPath = CHRONO_FOLDER & "\\" & folderName
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    
    mailItem.SaveAs folderPath & "\\" & SanitizeFileName(mailItem.Subject) & ".msg", 3
    
    MsgBox "Mail archiv" & Chr(233) & " !" & vbCrLf & "Dossier : " & folderName, vbInformation, "AutoChrono"
    
    Set fso = Nothing
    Exit Sub
    
CleanExit:
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    Exit Sub

ErrorHandler:
    MsgBox "Erreur archivage : " & Err.Description, vbExclamation, "AutoChrono"
    GoTo CleanExit
End Sub

'===============================================================================
' Nettoie le nom de fichier
'===============================================================================
Private Function SanitizeFileName(ByVal fn As String) As String
    Dim c As Variant
    For Each c In Array("<", ">", ":", Chr(34), "/", "\\", "|", "?", "*")
        fn = Replace(fn, c, "_")
    Next c
    If Len(fn) > 100 Then fn = Left(fn, 100)
    SanitizeFileName = fn
End Function
"""
