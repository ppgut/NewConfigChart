Attribute VB_Name = "m3CreateMPL"
Option Explicit

Sub PrepareMPL()

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim wsNewConfChart As Worksheet
    Dim wsMPL As Worksheet
    Dim wsSBConfChart As Worksheet
    Dim sSuperior As String
    Dim sSB As String
    Dim rng As Range
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim n As Integer
    Dim iLastRow As Integer
    Dim bPsbToCreate As Boolean
    Dim bReCreated As Boolean
    Dim bRwCreated As Boolean
    Dim tempStatus As String
    
    Set wsNewConfChart = ThisWorkbook.Worksheets("New Conf. Chart")
    Set wsMPL = ThisWorkbook.Worksheets("SB mods to upload")
    Set wsSBConfChart = ThisWorkbook.Worksheets("SB Conf. Chart")
    
    With wsMPL
        .Rows("2:" & .Rows.Count).Clear
        .Rows("2:" & .Rows.Count).Cells.Font.Color = vbBlack
        .Columns.HorizontalAlignment = xlCenter
        .Columns.VerticalAlignment = xlCenter
        .Columns.EntireColumn.NumberFormat = "@"
        Union(.Columns(colMPLDocNo), .Columns(colMPLActionType)).HorizontalAlignment = xlLeft
        .Rows(1).HorizontalAlignment = xlCenter
    End With
    
    m2ProcessNewConfigChart.CheckProgressions
    
    iLastRow = wsNewConfChart.Cells(1000000, colProgressionCheck).End(xlUp).Row
    bReCreated = False
    bRwCreated = False
    bPsbToCreate = False
    
    With wsMPL
        j = 2
        For i = 3 To iLastRow
            If wsNewConfChart.Cells(i, colProgressionCheck).Value = "ok" Then
            
                tempStatus = wsNewConfChart.Cells(i, colOpCode)
                
                If tempStatus = "RE/RW" Or tempStatus = "RW/RE" Then
                    If Not bReCreated Then
                        tempStatus = "RE"
                        bReCreated = True
                    ElseIf Not bRwCreated Then
                        tempStatus = "RW"
                        bRwCreated = True
                    End If
                End If
                
                If bPsbToCreate Then
                    tempStatus = "RW"
                End If

                .Cells(j, colMPLPrePn) = wsNewConfChart.Cells(i, colPrePN)
                .Cells(j, colMPLPreFID) = wsNewConfChart.Cells(i, colPreFID)
                .Cells(j, colMPLPreVar) = wsNewConfChart.Cells(i, colPreVariant)
                .Cells(j, colMPLPreQty) = wsNewConfChart.Cells(i, colPreQTY)
                
                .Cells(j, colMPLPostPn) = wsNewConfChart.Cells(i, colPostPN)
                .Cells(j, colMPLPostFID) = wsNewConfChart.Cells(i, colPostFID)
                .Cells(j, colMPLPostVar) = wsNewConfChart.Cells(i, colPostVariant)
                .Cells(j, colMPLPostQty) = wsNewConfChart.Cells(i, colPostQTY)
                
                .Cells(j, colMPLDocPart) = "000"
                .Cells(j, colMPLDocVer) = Application.WorksheetFunction.VLookup(wsNewConfChart.Cells(i, colSBNo), wsSBConfChart.Range("A:B"), 2, 0)
                If Not .Cells(j, colMPLDocVer) Like "??" Then .Cells(j, colMPLDocVer).Clear
                
                .Cells(j, colMPLOpCode) = StatusFullName(tempStatus)
                .Cells(j, colMPLActionType) = ModDescription(tempStatus)
                .Cells(j, colMPLChangeCode) = Application.WorksheetFunction.text(CStr(Left(wsNewConfChart.Cells(i, colChangeCode), 1)), "00")
                
                    'set sSuperior
                    If wsNewConfChart.Cells(i, colPreSuperior).Value <> "" Then
                        sSuperior = Right(wsNewConfChart.Cells(i, colPreSuperior).Value, Len(wsNewConfChart.Cells(i, colPreSuperior)) - 3)
                    Else
                        sSuperior = Right(wsNewConfChart.Cells(i, colPostSuperior).Value, Len(wsNewConfChart.Cells(i, colPostSuperior)) - 3)
                    End If
                    
                    'set sSB
                    If wsNewConfChart.Cells(i, colSBNo).Value Like "SB ??-????" Then
                        sSB = Right(wsNewConfChart.Cells(i, colSBNo).Value, Len(wsNewConfChart.Cells(i, colSBNo)) - 3)
                    ElseIf wsNewConfChart.Cells(i, colSBNo).Value Like "??-????" Then
                        sSB = wsNewConfChart.Cells(i, colSBNo).Value
                    End If
                
                Select Case bPsbToCreate
                Case False
                    .Cells(j, colMPLDocType) = "SSB"
                    .Cells(j, colMPLDocNo) = "2X_" & sSB & "_" & sSuperior
                Case True:
                    .Cells(j, colMPLDocType) = "PSB"
                    .Cells(j, colMPLDocNo) = "2X_" & sSB & "_" & sSuperior & "PSB"
                    bPsbToCreate = False
                End Select
 
                If .Cells(j, colMPLDocType) = "SSB" And .Cells(j, colMPLOpCode) = "REWORK" Then
                    bPsbToCreate = True
                End If
                
                If (bReCreated And Not bRwCreated) Or bPsbToCreate Then
                    i = i - 1
                End If
                
                If bReCreated And bRwCreated And Not bPsbToCreate Then
                    bReCreated = False
                    bRwCreated = False
                End If
                
                j = j + 1
            End If
        Next i
        
        For Each rng In .Range("A1").CurrentRegion.Cells
            If rng.Value = "--" Or rng.Value = "-" Then rng.Value = ""
        Next rng
    
        RemoveDuplicatedMods
        SortMPL
        AddCounter
        UseDocNameOnlyOnce
        
        With .Range("A1").CurrentRegion
            Union(.Columns(colMPLPrePn), .Columns(colMPLPostPn), .Columns(colMPLOpCode)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        End With
        
        For k = 1 To colMPLLast
            .Columns(k).EntireColumn.ColumnWidth = 100
            .Columns(k).EntireColumn.AutoFit
            .Columns(k).EntireColumn.ColumnWidth = .Columns(k).EntireColumn.ColumnWidth * 1.1
        Next k
    
    End With
    
    Set rng = Nothing
    Set wsNewConfChart = Nothing
    Set wsMPL = Nothing
    Set wsSBConfChart = Nothing
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub

Private Sub RemoveDuplicatedMods()
    
    Dim wsMPL As Worksheet
    Set wsMPL = ThisWorkbook.Worksheets("SB mods to upload")
    
    wsMPL.Range("A1").CurrentRegion.RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15), Header:=xlYes

    Set wsMPL = Nothing
    
End Sub

Public Sub SortMPL()

    Dim wsMPL As Worksheet
    Set wsMPL = ThisWorkbook.Worksheets("SB mods to upload")

    With wsMPL.Range("A1").CurrentRegion
    
        .Cells.Sort Key1:=.Columns(colMPLDocNo), Order1:=xlAscending, _
                    Key2:=.Columns(colMPLDocType), Order2:=xlDescending, _
                    Key3:=.Columns(colMPLPreFID), Order3:=xlAscending, _
                    Orientation:=xlTopToBottom, Header:=xlYes
        
    End With
    
    Set wsMPL = Nothing
    
End Sub

Private Sub AddCounter()
    
    Dim wsMPL As Worksheet
    Set wsMPL = ThisWorkbook.Worksheets("SB mods to upload")
    
    Dim i As Integer
    Dim iLastRow As Integer
    
    With wsMPL
        .Columns.ClearOutline
        iLastRow = .Cells(1000000, colMPLDocType).End(xlUp).Row
        
        If iLastRow >= 2 Then
            For i = 2 To iLastRow
            
                If i = 2 Or .Cells(i, colMPLDocType) <> .Cells(i, colMPLDocType).Offset(-1, 0) Or .Cells(i, colMPLDocNo) <> .Cells(i, colMPLDocNo).Offset(-1, 0) Then
                    wsMPL.Cells(i, colMPLCounter) = 1
                    .Range(.Cells(i, 1), .Cells(i, colMPLLast)).Borders(xlEdgeTop).LineStyle = xlContinuous
                Else
                    wsMPL.Cells(i, colMPLCounter) = wsMPL.Cells(i, colMPLCounter).Offset(-1, 0) + 1
                End If
            
            Next i
        End If

    End With
    
    Set wsMPL = Nothing
    
End Sub

Private Function StatusFullName(ByVal sStatus As String)
    Select Case sStatus
    Case "RM": StatusFullName = "REPLACE"
    Case "RE": StatusFullName = "REPLACE"
    Case "RW": StatusFullName = "REWORK"
    Case "RI": StatusFullName = "REWORK"
    Case "QTC": StatusFullName = "REPLACE"
    Case "AD": StatusFullName = ""
    Case "DE": StatusFullName = ""
    Case "RE/RW": StatusFullName = "RE/RW"
    Case "RW/RE": StatusFullName = "RE/RW"
    Case Else: StatusFullName = "#?"
    End Select
End Function

Private Function ModDescription(ByVal sStatus As String)
    Select Case sStatus
    Case "RM": ModDescription = "Parts Replaced"
    Case "RE": ModDescription = "Parts Replaced"
    Case "RW": ModDescription = "Parts Replaced"
    Case "RI": ModDescription = "Parts Replaced"
    Case "QTC": ModDescription = "Parts Replaced"
    Case "AD": ModDescription = "New Node Added"
    Case "DE": ModDescription = "Node Deleted"
    Case Else: ModDescription = "#?"
    End Select
End Function

Public Sub UseDocNameOnlyOnce()
    
    Dim wsMPL As Worksheet
    Set wsMPL = ThisWorkbook.Worksheets("SB mods to upload")
    
    Dim i As Integer
    Dim iLastRow As Integer
    
    With wsMPL
        iLastRow = .Cells(1000000, colMPLCounter).End(xlUp).Row
        
        If 3 <= iLastRow Then
            For i = iLastRow To 3 Step -1
            
                If .Cells(i, colMPLDocNo) = .Cells(i - 1, colMPLDocNo) Then
                    
                    .Cells(i, colMPLDocType).Clear
                    .Cells(i, colMPLDocNo).Clear
                    .Cells(i, colMPLDocVer).Clear
                    .Cells(i, colMPLDocPart).Clear
                    
                End If
            
            Next i
        End If
        
    End With
        
    Set wsMPL = Nothing
    
End Sub

Public Sub UseDocNameForEachLine()
    
    Dim wsMPL As Worksheet
    Set wsMPL = ThisWorkbook.Worksheets("SB mods to upload")
    
    Dim i As Integer
    Dim iLastRow As Integer
    
    With wsMPL
        iLastRow = .Cells(1000000, colMPLCounter).End(xlUp).Row
        
        If 3 <= iLastRow Then
            For i = 3 To iLastRow
            
                If .Cells(i, colMPLDocNo) = "" Then
                    
                    .Cells(i, colMPLDocType).Value = .Cells(i, colMPLDocType).Offset(-1, 0).Value
                    .Cells(i, colMPLDocNo) = .Cells(i, colMPLDocNo).Offset(-1, 0)
                    .Cells(i, colMPLDocVer) = .Cells(i, colMPLDocVer).Offset(-1, 0)
                    .Cells(i, colMPLDocPart) = .Cells(i, colMPLDocPart).Offset(-1, 0)
                    
                End If
            
            Next i
            Union(.Columns(colMPLDocType), .Columns(colMPLDocVer), .Columns(colMPLDocPart)).HorizontalAlignment = xlCenter
        End If
        
    End With
    
    Set wsMPL = Nothing
        
End Sub
