"""
AutoChrono - Générateur de code VBA
"""

import os


class VBAGenerator:
    """Génère le module VBA personnalisé pour Outlook."""
    
    def __init__(self, trigram, chrono_folder, chrono_file, col_chrono, col_client, col_trigram):
        self.trigram = trigram
        self.chrono_folder = chrono_folder.replace("/", "\\")
        self.chrono_file = chrono_file.replace("/", "\\")
        self.col_chrono = col_chrono
        self.col_client = col_client
        self.col_trigram = col_trigram
    
    def generate(self, output_path):
        """Génère le fichier .bas avec les valeurs personnalisées."""
        vba_code = self._get_vba_template()
        
        # Remplacer les placeholders
        vba_code = vba_code.replace("{{TRIGRAM}}", self.trigram)
        vba_code = vba_code.replace("{{CHRONO_FOLDER}}", self.chrono_folder)
        vba_code = vba_code.replace("{{CHRONO_FILE}}", self.chrono_file)
        vba_code = vba_code.replace("{{COL_CHRONO}}", self.col_chrono)
        vba_code = vba_code.replace("{{COL_CLIENT}}", self.col_client)
        vba_code = vba_code.replace("{{COL_TRIGRAM}}", self.col_trigram)
        
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(vba_code)
    
    def _get_vba_template(self):
        """Retourne le template VBA."""
        return '''Attribute VB_Name = "AutoChrono"
'===============================================================================
' AutoChrono - Module d'archivage automatique des mails Chrono
' Généré automatiquement - Ne pas modifier les constantes ci-dessous
'===============================================================================

Option Explicit

' Configuration personnalisée
Private Const USER_TRIGRAM As String = "{{TRIGRAM}}"
Private Const CHRONO_FOLDER As String = "{{CHRONO_FOLDER}}"
Private Const CHRONO_FILE As String = "{{CHRONO_FILE}}"
Private Const COL_CHRONO As String = "{{COL_CHRONO}}"
Private Const COL_CLIENT As String = "{{COL_CLIENT}}"
Private Const COL_TRIGRAM As String = "{{COL_TRIGRAM}}"

'===============================================================================
' Événement déclenché avant l'envoi d'un mail
'===============================================================================
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    On Error GoTo ErrorHandler
    
    Dim mailBody As String
    Dim chronoNumber As String
    
    ' Vérifier si c'est un mail
    If TypeName(Item) <> "MailItem" Then Exit Sub
    
    ' Récupérer les 100 premiers caractères du corps
    mailBody = Left(Item.Body, 100)
    
    ' Chercher le pattern "REF" et "N°"
    If InStr(1, mailBody, "REF", vbTextCompare) > 0 And InStr(1, mailBody, "N°", vbTextCompare) > 0 Then
        ' Extraire le numéro de chrono
        chronoNumber = ExtractChronoNumber(mailBody)
        
        If Len(chronoNumber) > 0 Then
            ' Demander confirmation à l'utilisateur
            If MsgBox("Voulez-vous ranger ce mail dans le dossier Chrono ?" & vbCrLf & vbCrLf & _
                      "N° Chrono détecté : " & chronoNumber, _
                      vbYesNo + vbQuestion, "AutoChrono") = vbYes Then
                
                ' Archiver le mail
                Call ArchiveMail(Item, chronoNumber)
            End If
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur AutoChrono : " & Err.Description, vbExclamation, "AutoChrono"
End Sub

'===============================================================================
' Extrait le numéro de Chrono du texte
'===============================================================================
Private Function ExtractChronoNumber(ByVal text As String) As String
    Dim regex As Object
    Dim matches As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.Pattern = "N°\\s*(\\d+)"
    
    If regex.Test(text) Then
        Set matches = regex.Execute(text)
        ExtractChronoNumber = matches(0).SubMatches(0)
    Else
        ExtractChronoNumber = ""
    End If
End Function

'===============================================================================
' Archive le mail dans le dossier Chrono (CORRIGÉ - Recherche par numéro)
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
    
    ' Ouvrir Excel
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    
    ' Ouvrir le fichier en lecture seule
    Set xlWb = xlApp.Workbooks.Open(CHRONO_FILE, ReadOnly:=True)
    Set xlWs = xlWb.Sheets(1)
    
    ' CORRECTION : Recherche du numéro de chrono dans la colonne A
    ' LookIn:=-4163 (xlValues), LookAt:=1 (xlWhole) pour correspondance exacte
    Set foundCell = xlWs.Columns(COL_CHRONO).Find(What:=chronoNumber, LookIn:=-4163, LookAt:=1)
    
    If foundCell Is Nothing Then
        ' Si le numéro n'existe pas dans Excel
        MsgBox "Le numéro de Chrono " & chronoNumber & " n'a pas été trouvé dans le fichier Excel." & vbCrLf & _
               "Le mail ne sera pas archivé automatiquement.", vbExclamation, "AutoChrono"
        GoTo CleanExit
    End If
    
    targetRow = foundCell.Row
    
    ' Récupérer les informations de la ligne trouvée
    clientName = Trim(CStr(xlWs.Range(COL_CLIENT & targetRow).Value))
    trigramme = Trim(CStr(xlWs.Range(COL_TRIGRAM & targetRow).Value))
    
    ' Sécurité : remplacer les valeurs vides ou "0" par des valeurs par défaut
    If clientName = "" Or clientName = "0" Then clientName = "Client_Inconnu"
    If trigramme = "" Or trigramme = "0" Then trigramme = "Inc"
    
    ' Fermer Excel proprement
    xlWb.Close False
    xlApp.Quit
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    
    ' Créer le nom du dossier
    folderName = chronoNumber & " - " & clientName & " (" & trigramme & ")"
    folderPath = CHRONO_FOLDER & "\\" & folderName
    
    ' Créer le dossier s'il n'existe pas
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    
    ' Sauvegarder le mail
    mailItem.SaveAs folderPath & "\\" & SanitizeFileName(mailItem.Subject) & ".msg", 3 ' olMSG = 3
    
    MsgBox "Mail archivé avec succès !" & vbCrLf & vbCrLf & _
           "Dossier : " & folderName, vbInformation, "AutoChrono"
    
    Set fso = Nothing
    Exit Sub
    
CleanExit:
    ' Nettoyage de sécurité
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set fso = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Erreur lors de l'archivage : " & Err.Description, vbExclamation, "AutoChrono"
    GoTo CleanExit
End Sub

'===============================================================================
' Nettoie le nom de fichier des caractères invalides
'===============================================================================
Private Function SanitizeFileName(ByVal fileName As String) As String
    Dim invalidChars As Variant
    Dim i As Integer
    
    invalidChars = Array("<", ">", ":", """", "/", "\\", "|", "?", "*")
    
    For i = LBound(invalidChars) To UBound(invalidChars)
        fileName = Replace(fileName, invalidChars(i), "_")
    Next i
    
    ' Limiter la longueur
    If Len(fileName) > 100 Then
        fileName = Left(fileName, 100)
    End If
    
    SanitizeFileName = fileName
End Function
'''
