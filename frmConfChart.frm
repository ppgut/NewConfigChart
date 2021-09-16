VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConfChart 
   Caption         =   " "
   ClientHeight    =   3930
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   1785
   OleObjectBlob   =   "frmConfChart.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConfChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    m2ProcessNewConfigChart.SplitRowOneToOne
    AppActivate Application.Caption
End Sub

Private Sub CommandButton2_Click()
    m2ProcessNewConfigChart.SplitRowAnyToAny
    AppActivate Application.Caption
End Sub

Private Sub CommandButton4_Click()
    m3CreateMPL.PrepareMPL
    ThisWorkbook.Worksheets("SB mods to upload").Activate
End Sub

Private Sub CommandButton5_Click()
    m2ProcessNewConfigChart.CopyRow
    AppActivate Application.Caption
End Sub

Private Sub CommandButton3_Click()
    m2ProcessNewConfigChart.CheckProgressions
    AppActivate Application.Caption
End Sub

Private Sub CommandButton6_Click()
    m2ProcessNewConfigChart.AddMPLData
    AppActivate Application.Caption
End Sub

Private Sub CommandButton7_Click()
    Dim rCell As Range
    Dim rLastRow As Integer
    
    Dim rAfter As Range
    Dim rRangeToLook As Range
    
    With Sheet2
        rLastRow = .Cells(65000, colPrePN).End(xlUp).Row
        Set rAfter = .Cells(Application.Min(rLastRow, Selection.Item(1).Row), colPrePN)
        Set rRangeToLook = Range(.Cells(2, colPrePN), .Cells(rLastRow, colPrePN))
        
        If Not rAfter Is Nothing And Not rRangeToLook Is Nothing Then
            Set rCell = rRangeToLook.Find(vbLf, rAfter, , xlPart)

            If rCell Is Nothing Then
                Set rAfter = .Cells(Application.Min(rLastRow, Selection.Item(1).Row), colPostPN)
                Set rRangeToLook = Range(.Cells(2, colPostPN), .Cells(rLastRow, colPostPN))
                
                If Not rAfter Is Nothing And Not rRangeToLook Is Nothing Then
                    Set rCell = rRangeToLook.Find(vbLf, rAfter, , xlPart)
                End If
            End If
        End If
        
        If Not rCell Is Nothing Then rCell.Select
        
    End With
End Sub

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.93 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.8 * Application.Height) - (0.5 * Me.Height)
    AppActivate Application.Caption
End Sub
