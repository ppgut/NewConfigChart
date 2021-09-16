Attribute VB_Name = "m8LoadSBFile"
Const sSBFolderName As String = "\SB Config Charts\"


Sub LoadSBConfChart()

    Dim diaFile As FileDialog
    Dim sPath() As Variant
    Dim i       As Integer
    Dim n       As Long
    Dim bLoadFromList As Boolean
    Dim sSBConfigChartFilePath As String
    bLoadFromList = False
    Dim bNoConfChart As Boolean
    Dim bSpare As Boolean
    
    If Sheet1.Range("S2") <> "" Then
        If MsgBox("Load SBs from the list?", vbYesNo) = vbYes Then
            bLoadFromList = True
            Sheet1.Columns(21).Cells.Clear
            Sheet1.Columns(21).Cells.VerticalAlignment = xlVAlignCenter
            Sheet1.Columns(21).ColumnWidth = 25
        End If
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    If Not bLoadFromList Then
        Set diaFile = Application.FileDialog(msoFileDialogFilePicker)
        diaFile.AllowMultiSelect = True
        If Not diaFile.Show Then Exit Sub
        
        n = diaFile.SelectedItems.Count
        
        ReDim sPath(1 To n)
        For i = 1 To n
            sPath(i) = diaFile.SelectedItems(i)
        Next i
    Else
        If Sheet1.Range("S3") = "" Then
            n = 1
        Else
            n = Sheet1.Range(Cells(2, 19), Cells(2, 19).End(xlDown)).Count
        End If
        
        FindLatestRev n
        
        ReDim sPath(1 To n)
        For i = 1 To n
            sPath(i) = FilePath(i, Sheet1.Cells(i + 1, 19).Value & " R" & Format(Sheet1.Cells(i + 1, 20).Value, "00"))
        Next i
    End If
    
    For i = 1 To n

        If Len(Dir(sPath(i))) > 0 Then
            bNoConfChart = False
            bSpare = False
            
            If bLoadFromList Then
                If InStr(Sheet1.Cells(i + 1, 21), "no Config Chart") > 0 Then bNoConfChart = True
                If InStr(Sheet1.Cells(i + 1, 21), "Spare") > 0 Then bSpare = True
            Else
                If InStr(sPath(i), "no Config Chart") > 0 Then bNoConfChart = True
                If InStr(sPath(i), "Spare") > 0 Then bSpare = True
            End If
            
            If Not bNoConfChart Then
                If Sheet1.OBOnlySP Then
                    If bSpare Then
                        Load (sPath(i))
                    ElseIf bLoadFromList Then
                        Sheet1.Cells(i + 1, 21).Value = "not Spare Part - not loaded"
                    End If
                ElseIf Sheet1.OBNoSP Then
                    If Not bSpare Then
                        Load (sPath(i))
                    ElseIf bLoadFromList Then
                        Sheet1.Cells(i + 1, 21).Value = "Spare Part - not loaded"
                    End If
                Else
                    Load (sPath(i))
                End If
            End If
        Else
            Sheet1.Cells(i + 1, 21).Value = "File not found"
        End If
    Next i

    Application.CutCopyMode = False
    
    Call ChangeSignsToStandard
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub

Sub Load(sPath As String)
    
    Dim wsSource As Worksheet
    Dim wsConfChart As Worksheet
    Dim rPasteCell As Range
    Dim rLastCell As Range
    Dim sSBNo As String
    Dim sSBVer As String
    
    Set wsConfChart = ThisWorkbook.Worksheets("SB Conf. Chart")
    Set wsSource = Workbooks.Open(sPath).Worksheets(1)
    
    sSBNo = Left(wsSource.Parent.Name, 7)
    sSBVer = Mid(wsSource.Parent.Name, 10, 2)
    
    Set rPasteCell = wsConfChart.Range("A1000000").End(xlUp)
    If rPasteCell.Row <> 1 Then Set rPasteCell = rPasteCell.Offset(1, 0)

    rPasteCell.Value = sSBNo
    rPasteCell.Offset(0, 1).Value = sSBVer

    Set rLastCell = Range("A1000000").End(xlUp)
    wsSource.Range(Range("A1"), rLastCell.Offset(rLastCell.MergeArea.Rows.Count - 1, 6)).Copy rPasteCell.Offset(1, 0)
   
    wsSource.Parent.Close SaveChanges:=False
    
End Sub

Private Sub FindLatestRev(SBOnTheListCount As Long)

    Dim oFSO            As Object
    Dim oFile           As Object
    Dim oFolder         As Object
    Dim i               As Single
    Dim RevNo           As Single
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.getfolder(ThisWorkbook.Path & sSBFolderName)
    
    With Sheet1
        
        For i = 1 To SBOnTheListCount
        
            If .Cells(i + 1, 20).Font.Color = vbBlue Then
                .Cells(i + 1, 20).Clear
            End If
            If .Cells(i + 1, 20).Value = "" Then
            
                For Each oFile In oFolder.Files
                
                    If Left(oFile.Name, 7) = .Cells(i + 1, 19) Then
                        
                        RevNo = CSng(Mid(oFile.Name, 10, 2))
                        If .Cells(i + 1, 20).Value = "" Or (.Cells(i + 1, 20).Font.Color = vbBlue And RevNo > CSng(.Cells(i + 1, 20).Value)) Then
                            .Cells(i + 1, 20).Value = Format(RevNo, "00")
                            .Cells(i + 1, 20).Font.Color = vbBlue
                        End If

                    End If
                
                Next oFile
            
            End If
        
        Next i

        .Columns(20).EntireColumn.HorizontalAlignment = xlCenter
        .Columns(20).EntireColumn.VerticalAlignment = xlVAlignCenter
        
    End With
    
    Set oFSO = Nothing
    Set oFile = Nothing
    Set oFolder = Nothing
    
End Sub

Private Function FilePath(iRow As Integer, PathToCheck As String) As String

    Dim oFSO            As Object
    Dim oFile           As Object
    Dim oFolder         As Object
    
    FilePath = ThisWorkbook.Path & sSBFolderName & PathToCheck
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.getfolder(ThisWorkbook.Path & sSBFolderName)
    
    For Each oFile In oFolder.Files
    
        If Left(oFile.Name, 11) = PathToCheck Then
            
            If Len(Left(oFile.Name, InStr(oFile.Name, ".") - 1)) > 11 Then
                Sheet1.Cells(iRow + 1, 21) = Mid(oFile.Name, 13, InStr(oFile.Name, ".") - 13)
            End If
            
            FilePath = ThisWorkbook.Path & sSBFolderName & oFile.Name
            Set oFSO = Nothing
            Set oFile = Nothing
            Set oFolder = Nothing
            Exit Function
        
        End If
    
    Next oFile

    Set oFSO = Nothing
    Set oFile = Nothing
    Set oFolder = Nothing

End Function

Sub ChangeSignsToStandard()

    Dim wsConfChart As Worksheet
    Dim rConfChartAreaToCheck As Range
    Dim rng As Range
    
    Set wsConfChart = ThisWorkbook.Worksheets("SB Conf. Chart")
    With wsConfChart.Range(Cells(1, 1), Cells(1000000, 7).End(xlUp))
        Set rConfChartAreaToCheck = Union(.Columns(1), .Columns(4))
    End With
    
    For Each rng In rConfChartAreaToCheck
        If rng.Value <> "" Then
            rng.Value = Replace(rng.Value, ChrW(8229), "..")
            rng.Value = Replace(rng.Value, Chr(133), "...")
        End If
    Next rng
    
    Set rng = Nothing
    Set wsConfChart = Nothing
    
End Sub

Public Sub LoadSSBMPL()

    Dim diaFile As FileDialog
    Dim sPath() As Variant
    Dim i       As Integer
    Dim n       As Long
    
    Set diaFile = Application.FileDialog(msoFileDialogFilePicker)
    diaFile.AllowMultiSelect = True
    If Not diaFile.Show Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    m3CreateMPL.UseDocNameForEachLine
    
    n = diaFile.SelectedItems.Count
    
    ReDim sPath(1 To n)
    
    For i = 1 To n
        sPath(i) = diaFile.SelectedItems(i)
    Next i
    
    For i = 1 To n
        LoadSSBFile (sPath(i))
    Next i
    
    m3CreateMPL.SortMPL
    m3CreateMPL.UseDocNameOnlyOnce
    
    Dim wsMPL As Worksheet
    Set wsMPL = ThisWorkbook.Worksheets("SB mods to upload")
    
    Dim l As Integer
    Dim lLastRow As Integer
    
    With wsMPL
    
        .Cells.Borders.LineStyle = xlNone
        With .Range("A1").CurrentRegion
            Union(.Columns(colMPLPrePn), .Columns(colMPLPostPn), .Columns(colMPLOpCode)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Cells.NumberFormat = "@"
            .Cells.HorizontalAlignment = xlCenter
            Union(.Columns(colMPLDocNo), .Columns(colMPLActionType)).HorizontalAlignment = xlLeft
        End With
        
        lLastRow = .Cells(1000000, colMPLCounter).End(xlUp).Row
        
        If lLastRow >= 2 Then
            For l = 2 To lLastRow
            
                If .Cells(l, colMPLDocType) <> "" Then
                    Range(.Cells(l, 1), .Cells(l, colMPLLast)).Borders(xlEdgeTop).LineStyle = xlContinuous
                End If
            
            Next l
        End If
    End With
    
    Call AddLegend
    
    Set wsMPL = Nothing
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub

Private Sub LoadSSBFile(sPath As String)

    Dim wsMPL As Worksheet
    Dim wsSAPMPL As Worksheet
    Dim sDocName As String
    Dim sDocVer As String
    
    Set wsMPL = ThisWorkbook.Worksheets("SB mods to upload")
    Set wsSAPMPL = Workbooks.Open(sPath).Worksheets("Sheet1")
    sDocName = Left(wsSAPMPL.Parent.Name, InStr(wsSAPMPL.Parent.Name, " ") - 1)
    sDocVer = Mid(wsSAPMPL.Parent.Name, InStr(wsSAPMPL.Parent.Name, " ") + 2, 2)
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim iStart As Long
    Dim jLast As Long
    Dim kStart As Long
    Dim kEnd As Long
    
    i = wsMPL.Cells(65000, colMPLCounter).End(xlUp).Row
    iStart = i + 1                                                      'first empty row of MPL (for pasting imported mods)
    jLast = wsSAPMPL.Cells(65000, colSAPMPLCounter).End(xlUp).Row       'last mod row in SAPMPL file
    kStart = SSBFirstRow(wsMPL, sDocName)                               'first row of SSB area in Created MPL
    kEnd = SSBLastRow(wsMPL, sDocName, kStart)                          'last row of SSB area in Created MPL
       
    If jLast = 1 Then Exit Sub
    
    With wsMPL
    
        .Columns(colMPLDocPart).EntireColumn.NumberFormat = "@"
        
        'delete previous comparison of this SSB (delete blue and red lines and change green to black)
        For l = 2 To i
            If .Cells(l, colMPLDocNo).Value = sDocName Then
                If .Cells(l, colMPLCounter).Font.Color = vbBlue Or .Cells(l, colMPLCounter).Font.Color = vbRed Then
                    .Cells(l, colMPLCounter).EntireRow.Delete
                    l = l - 1
                    i = i - 1
                    iStart = iStart - 1
                    kEnd = kEnd - 1
                ElseIf .Cells(l, colMPLCounter).Font.Color = 5287936 Then
                    .Cells(l, colMPLCounter).EntireRow.Font.Color = vbBlack
                End If
            End If
        Next l
        
        For j = 2 To jLast
            
            k = 0
            If kStart <> 0 And kEnd <> 0 And kEnd >= kStart Then
                k = MatchingModRow(wsMPL, wsSAPMPL, j, kStart, kEnd) 'check if extracted modification exists on created MPL list
            End If
            
            If k > 0 Then
                
                Range(.Cells(k, colMPLCounter), .Cells(k, colMPLChangeCode)).Font.Color = 5287936 'if the modification is in created MPL, make MPL line green
                        
            Else

                kEnd = kEnd + 1
                .Rows(kEnd).Insert
                'i = i + 1
                If k = -1 Then
                    'if the modification is in created MPL and the line is already green or blue, make the new line red to mark duplicated mod
                    Range(.Cells(kEnd, colMPLCounter), .Cells(kEnd, colMPLChangeCode)).Font.Color = vbRed
                Else
                    'if the line from extracted file is not in created MPL, make new line blue
                    Range(.Cells(kEnd, colMPLCounter), .Cells(kEnd, colMPLChangeCode)).Font.Color = vbBlue
                End If
                
                .Cells(kEnd, colMPLDocType) = "SSB"
                .Cells(kEnd, colMPLDocNo) = sDocName
                .Cells(kEnd, colMPLDocVer) = sDocVer
                .Cells(kEnd, colMPLDocPart) = "000"
                .Cells(kEnd, colMPLCounter) = CLng(wsSAPMPL.Cells(j, colSAPMPLCounter))
                .Cells(kEnd, colMPLPrePn) = Application.Run("PNshort", wsSAPMPL.Cells(j, colSAPMPLPrePN))
                .Cells(kEnd, colMPLPreFID) = wsSAPMPL.Cells(j, colSAPMPLPreFID)
                .Cells(kEnd, colMPLPreVar) = wsSAPMPL.Cells(j, colSAPMPLPreVar)
                .Cells(kEnd, colMPLPreQty) = wsSAPMPL.Cells(j, colSAPMPLPreQty)
                .Cells(kEnd, colMPLPostPn) = Application.Run("PNshort", wsSAPMPL.Cells(j, colSAPMPLPostPN))
                .Cells(kEnd, colMPLPostFID) = wsSAPMPL.Cells(j, colSAPMPLPostFID)
                .Cells(kEnd, colMPLPostVar) = wsSAPMPL.Cells(j, colSAPMPLPostVar)
                .Cells(kEnd, colMPLPostQty) = wsSAPMPL.Cells(j, colSAPMPLPostQty)
                .Cells(kEnd, colMPLOpCode) = wsSAPMPL.Cells(j, colSAPMPLStatus)
                .Cells(kEnd, colMPLActionType) = wsSAPMPL.Cells(j, colSAPMPLActionType)
                .Cells(kEnd, colMPLChangeCode) = wsSAPMPL.Cells(j, colSAPMPLChangeCode)
                
            End If
            
        Next j
        
        wsSAPMPL.Parent.Close SaveChanges:=False
        
    End With
    
    Set wsMPL = Nothing
    
End Sub

Private Function SSBFirstRow(wsDest As Worksheet, sDocName As String) As Long

    Dim rRangeToCheck As Range
    Set rRangeToCheck = Range(wsDest.Cells(1, colMPLDocNo), wsDest.Cells(1000000, colMPLDocNo).End(xlUp))
    
    Dim rFoundCell As Range
    Set rFoundCell = rRangeToCheck.Find(sDocName, lookat:=xlWhole)
    If Not rFoundCell Is Nothing Then
        SSBFirstRow = rFoundCell.Row
    Else
        SSBFirstRow = 0
    End If

End Function

Private Function SSBLastRow(wsDest As Worksheet, sDocName As String, kStart As Long) As Long

    SSBLastRow = 0
    Dim rng As Range
    
    If kStart = 0 Then Exit Function
    
    For Each rng In Range(wsDest.Cells(kStart, colMPLDocNo), wsDest.Cells(65000, colMPLDocNo).End(xlUp))
        
        If rng.Value = sDocName And rng.Offset(1, 0).Value <> sDocName Then
            SSBLastRow = rng.Row
            Exit Function
        End If

    Next rng

End Function

Private Function MatchingModRow(ByVal wsDest As Worksheet, ByVal wsSrc As Worksheet, ByVal jSrc As Long, ByVal kDestStart As Long, ByVal kDestEnd As Long) As Long

    MatchingModRow = 0
    Dim k As Long
    
    With wsDest
    
        For k = kDestStart To kDestEnd
            If _
                .Cells(k, colMPLPrePn) = Application.Run("PNshort", wsSrc.Cells(jSrc, colSAPMPLPrePN)) And _
                .Cells(k, colMPLPreFID) = wsSrc.Cells(jSrc, colSAPMPLPreFID) And _
                .Cells(k, colMPLPreVar) = wsSrc.Cells(jSrc, colSAPMPLPreVar) And _
                (.Cells(k, colMPLPreQty) = wsSrc.Cells(jSrc, colSAPMPLPreQty) Or (.Cells(k, colMPLPreQty) = "" And wsSrc.Cells(jSrc, colSAPMPLPreQty) = 0)) And _
                .Cells(k, colMPLPostPn) = Application.Run("PNshort", wsSrc.Cells(jSrc, colSAPMPLPostPN)) And _
                .Cells(k, colMPLPostFID) = wsSrc.Cells(jSrc, colSAPMPLPostFID) And _
                .Cells(k, colMPLPostVar) = wsSrc.Cells(jSrc, colSAPMPLPostVar) And _
                (.Cells(k, colMPLPostQty) = wsSrc.Cells(jSrc, colSAPMPLPostQty) Or (.Cells(k, colMPLPostQty) = "" And wsSrc.Cells(jSrc, colSAPMPLPostQty) = 0)) And _
                .Cells(k, colMPLOpCode) = wsSrc.Cells(jSrc, colSAPMPLStatus) Then
                
                If .Cells(k, colMPLCounter).Font.Color = 5287936 Or .Cells(k, colMPLCounter).Font.Color = vbBlue Then
                    MatchingModRow = -1 'row already exists
                Else
                    MatchingModRow = k
                End If
                Exit Function
                
            End If
        Next k
        
    End With
    
End Function

Private Sub AddLegend()

    With ThisWorkbook.Worksheets("SB mods to upload")

        .Range("R12").Font.Color = vbBlack
        .Range("R13").Font.Color = 5287936
        .Range("R14").Font.Color = vbBlue
        .Range("R15").Font.Color = vbRed
        
        .Range("R12").Value = "mod created from New Config Chart, not defined in the system"
        .Range("R13").Value = "mod created from New Config Chart, defined in the system"
        .Range("R14").Value = "mod doesn't created from New Config Chart, defined in the system"
        .Range("R15").Value = "mod duplicated in the system"
        
        .Range("R12:R15").HorizontalAlignment = xlLeft
        
    End With

End Sub
