Attribute VB_Name = "FPPN_navigation"
' FPPN - Classeur des calculs : boutons de navigation (facultatifs)
' Mode d'emploi (2 minutes) :
'   1. Alt + F11 (éditeur VBA) ; menu Fichier > Importer un fichier... ; choisir FPPN_navigation.bas ; fermer l'éditeur.
'   2. Alt + F8 ; choisir InstallerBoutons ; Exécuter.
'   3. Fichier > Enregistrer sous > type « Classeur Excel (prenant en charge les macros) (*.xlsm) ».
' Les liens bleus de la ligne 1 de chaque onglet marchent sans ces macros.
Option Explicit

Public Sub AllerAuGuide()
    On Error Resume Next
    ThisWorkbook.Worksheets("0_guide").Activate
    ThisWorkbook.Worksheets("0_guide").Range("A1").Select
End Sub

Public Sub AllerAuControle()
    On Error Resume Next
    ThisWorkbook.Worksheets("0_controle").Activate
End Sub

Public Sub OngletSuivant()
    Dim i As Long
    i = ActiveSheet.Index
    Do While i < ThisWorkbook.Sheets.Count
        i = i + 1
        If ThisWorkbook.Sheets(i).Visible = xlSheetVisible Then ThisWorkbook.Sheets(i).Activate: Exit Sub
    Loop
End Sub

Public Sub OngletPrecedent()
    Dim i As Long
    i = ActiveSheet.Index
    Do While i > 1
        i = i - 1
        If ThisWorkbook.Sheets(i).Visible = xlSheetVisible Then ThisWorkbook.Sheets(i).Activate: Exit Sub
    Loop
End Sub

Public Sub ListeDesOnglets()
    ' Affiche la liste de tous les onglets (clic droit sur les flèches en bas à gauche, version Windows)
    On Error Resume Next
    Application.CommandBars("Workbook tabs").ShowPopup
    If Err.Number <> 0 Then AllerAuGuide
End Sub

Public Sub InstallerBoutons()
    Dim ws As Worksheet, b As Object, x As Double, k As Long
    Dim libelles As Variant, macros As Variant
    libelles = Array("Guide", "< Précédent", "Suivant >", "Tous les onglets", "Contrôle")
    macros = Array("AllerAuGuide", "OngletPrecedent", "OngletSuivant", "ListeDesOnglets", "AllerAuControle")
    For Each ws In ThisWorkbook.Worksheets
        For Each b In ws.Buttons
            If Left(b.Name, 4) = "nav_" Then b.Delete
        Next b
        x = ws.Range("L1").Left
        For k = 0 To UBound(libelles)
            Set b = ws.Buttons.Add(x + k * 92, ws.Range("A1").Top + 1, 88, 17)
            b.Name = "nav_" & k
            b.Caption = libelles(k)
            b.OnAction = macros(k)
            b.Font.Size = 8
        Next k
    Next ws
    MsgBox "Boutons installés sur " & ThisWorkbook.Worksheets.Count & " onglets. Enregistrez le classeur au format .xlsm.", vbInformation
End Sub

Public Sub SupprimerBoutons()
    Dim ws As Worksheet, b As Object
    For Each ws In ThisWorkbook.Worksheets
        For Each b In ws.Buttons
            If Left(b.Name, 4) = "nav_" Then b.Delete
        Next b
    Next ws
End Sub
