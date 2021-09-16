Attribute VB_Name = "m1CreateNewConfigChart"
Option Explicit

Public NewCCArr() As Variant
Public LineType() As Variant

'--------------------------------------------------------------------------------
'----------------------------- Create New Config Chart --------------------------
'--------------------------------------------------------------------------------

Sub CreateNewConfigChart()

    Dim dTime As Date
    
    Dim rng As Range
    Dim i As Long
    Dim iStart As Long
    Dim iEnd As Long
    Dim o As Long
    Dim m As Long
    Dim j As Integer
    Dim k As Integer
    Dim q As Integer
    Dim kLastRow As Integer
    Dim bAdd As Boolean
    Dim NewCCarrTemp() As Variant
    
    Dim wsNewCC As Worksheet
    Dim wsOldCC As Worksheet
    Set wsNewCC = ThisWorkbook.Worksheets("New Conf. Chart")
    Set wsOldCC = ThisWorkbook.Worksheets("SB Conf. Chart")
    
    'Ask if data in New Conf. Chart worskheet should be erased
    If wsNewCC.Range("A3").Value <> "" Then
        If MsgBox("Clear existing data?", vbYesNo) <> vbYes Then bAdd = True
    End If
    
    dTime = Time
    Debug.Print "NewCC"
    Debug.Print Format(dTime, "hh:mm:ss")
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    With wsNewCC
    
        'existing data to be appended
        If bAdd Then
            .Columns.ClearOutline
            iStart = .Cells(65000, 1).End(xlUp).Row + 1
            
        'existing data to be erased
        Else
            PrepareEmptyChart
            iStart = 2
        End If
    End With
    
    Erase NewCCArr
    Erase LineType
    i = 0
    k = 1
    
    With wsOldCC
    
        kLastRow = .Cells(.Rows.Count, 5).End(xlUp).Row 'column 5 choosed cause pre qty cell is always merged for whole individual modification
    
        Do While k <= kLastRow
            
            If .Cells(k, 5) <> "" Then
                Set rng = .Cells(k, 5)
                
            'if cells(k, 5) = "" it means whole row of old configuration chart is merged into one cell
            Else
                Set rng = .Cells(k, 1)
            End If
            
            i = i + 1
            ReDim Preserve NewCCarrTemp(1 To colLast, 1 To i)
            ReDim Preserve LineType(1 To i)
            
            'LineType = 0 - standard line
            'LineType = 1 - merged line in Old Config Chart with value other than "OR" or "Deleted" or line containing SB number
            'LineType = 2 - merged line in Old Config Chart with SB number
            'LineType = 3 - line containing at least one cell for additional formatting
            
            'LineType = 0 - no additional formatting
            'LineType = 1 - .Cells(i, colName).Font.Bold = True,
            'LineType = 1 - .Range(.Cells(i, 1), .Cells(i, colLast)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            'LineType = 2 - LineType 1 formatting + .Cells(i, colName).Font.Color = vbRed
            'LineType = 3 - values "OR", "Deleted", "--" to be centered; qty="X" to be changed to "1" and font color to blue
            
            'condition for merged line in Old Config Chart
            If rng.Column = 1 Then
            
                NewCCarrTemp(colName, i) = rng.Value
                LineType(i) = 1
                
                If NewCCarrTemp(colSBNo, i) = "" And i > 1 Then
                    NewCCarrTemp(colSBNo, i) = NewCCarrTemp(colSBNo, i - 1)
                End If
                
                If NewCCarrTemp(colName, i) = "OR" Or NewCCarrTemp(colName, i) = "Deleted" Then
                    LineType(i) = 3
                ElseIf NewCCarrTemp(colName, i) Like "SB *" Or NewCCarrTemp(colName, i) Like "??-????" Then
                    LineType(i) = 2
                    NewCCarrTemp(colSBNo, i) = rng.Value
                End If
                
            'all other rows
            Else
                LineType(i) = 0
                If NewCCarrTemp(colSBNo, i) = "" And i > 1 Then NewCCarrTemp(colSBNo, i) = NewCCarrTemp(colSBNo, i - 1)
                NewCCarrTemp(colName, i) = Name(rng)
                NewCCarrTemp(colSIN, i) = PartSIN(rng)
                NewCCarrTemp(colPrePN, i) = OldPartNumber(rng)
                NewCCarrTemp(colPreATA, i) = OldATA(rng)
                NewCCarrTemp(colPreQTY, i) = OldQty(rng)
    
                NewCCarrTemp(colPostPN, i) = NewPartNumber(rng)
                NewCCarrTemp(colPostATA, i) = NewATA(rng)
                NewCCarrTemp(colPostQTY, i) = NewQty(rng)
    
                NewCCarrTemp(colOpCode, i) = OpCode(rng)
                NewCCarrTemp(colChangeCode, i) = ChangeCode(rng)
                If IsNumeric(NewCCarrTemp(colPreQTY, i)) Then NewCCarrTemp(colPreQTY, i) = CInt(NewCCarrTemp(colPreQTY, i))
                If IsNumeric(NewCCarrTemp(colPostQTY, i)) Then NewCCarrTemp(colPostQTY, i) = CInt(NewCCarrTemp(colPostQTY, i))
                
                'additional formatting
                If NewCCarrTemp(colPrePN, i) = "--" Or _
                    NewCCarrTemp(colPreATA, i) = "--" Or _
                    NewCCarrTemp(colPostPN, i) = "--" Or _
                    NewCCarrTemp(colPostATA, i) = "--" Or _
                    NewCCarrTemp(colPreQTY, i) = "X" Or _
                    NewCCarrTemp(colPostQTY, i) = "X" Then
                    LineType(i) = 3
                End If
            End If
            
            If UCase(Trim(NewCCarrTemp(colPrePN, i))) Like "*OLDPARTNUMBER*" Or NewCCarrTemp(colName, i) = "" Or NewCCarrTemp(colName, i) = "Deleted" Then
                For q = 1 To colLast
                    NewCCarrTemp(q, i) = ""
                Next q
                i = i - 1
            End If
            
            DoEvents
            If .Cells(k, 5) <> "" And .Cells(k, 5).MergeCells Then
                k = k + .Cells(k, 5).MergeArea.Rows.Count
            Else
                k = k + 1
            End If
        Loop
        iEnd = iStart - 1 + i
        
    End With
    
    Call CorrectPNsInNewConfigChart(NewCCarrTemp)
    Call MakeIndentions(NewCCarrTemp)
    Call SplitEasyOnes(NewCCarrTemp, LineType, iEnd)
    
    ReDim NewCCArr(1 To UBound(NewCCarrTemp, 2), 1 To UBound(NewCCarrTemp, 1))
    For m = 1 To UBound(NewCCarrTemp, 2)
        For o = 1 To UBound(NewCCarrTemp, 1)
            NewCCArr(m, o) = NewCCarrTemp(o, m)
        Next o
    Next m
           
    wsNewCC.Cells(iStart, 1).Resize(iEnd - iStart + 1, colLast).Value = NewCCArr
    wsNewCC.Select
    
    
    Call FormatNewConfigChart                               'General formatting for New Conf. Chart worksheet
    Call FormatNewEntries(iStart, iEnd)                     'Cells formatting for newly created entries based on their values and type of entry
        
    Set rng = Nothing

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    frmConfChart.Show 0
    
    ThisWorkbook.Worksheets("New Conf. Chart").Range("B2").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    
    Debug.Print Format(Time, "hh:mm:ss")
    Debug.Print Format(Time - dTime, "hh:mm:ss")
    
End Sub

Private Sub PrepareEmptyChart()
    'Clear the New Conf. Chart worksheet and recreates whole layout from scratch

    With ThisWorkbook.Worksheets("New Conf. Chart")
        If .AutoFilterMode Then .AutoFilterMode = False
        .Columns.ClearOutline
        .Cells.Clear
        .Columns.UseStandardWidth = True
        .Rows.UseStandardHeight = True
        .Cells.Font.Bold = False
        .Cells.Font.Color = vbBlack
        .Cells.Borders.LineStyle = xlNone
        .Cells.HorizontalAlignment = xlLeft
        .Cells.VerticalAlignment = xlCenter
        .Range(.Cells(1, 1), .Cells(1, colLast)).EntireColumn.NumberFormat = "@"
        .Range(.Cells(1, colPreQTY), .Cells(1, colPrePPEQTY)).EntireColumn.NumberFormat = "General"
        .Range(.Cells(1, colPostQTY), .Cells(1, colPostPPEQTY)).EntireColumn.NumberFormat = "General"
        
        .Cells(1, colSBNo) = "SB"
        .Cells(1, colName) = "Name"
        .Cells(1, colSIN) = "SIN"
        
        .Cells(1, colPrePN) = "Pre" & vbLf & "PN"
        .Cells(1, colPreATA) = "Pre" & vbLf & "ATA"
        .Cells(1, colPreQTY) = "Pre" & vbLf & "Qty"
        
        .Cells(1, colPreFIDNo) = " "
        .Cells(1, colPreSuperiorNo) = " "
        .Cells(1, colPreVariantNo) = " "
        
        .Cells(1, colPreFID) = "FID"
        .Cells(1, colPreSuperior) = "Superior"
        .Cells(1, colPreVariant) = "Variant"
        .Cells(1, colPreObjDep) = "Obj" & vbLf & "Dep"
        .Cells(1, colPrePPEQTY) = "PPE" & vbLf & "Qty"
        
        .Cells(1, colPostPN) = "Post" & vbLf & "PN"
        .Cells(1, colPostATA) = "Post" & vbLf & "ATA"
        .Cells(1, colPostQTY) = "Post" & vbLf & "Qty"
        
        .Cells(1, colPostFIDNo) = " "
        .Cells(1, colPostSuperiorNo) = " "
        .Cells(1, colPostVariantNo) = " "
        
        .Cells(1, colPostFID) = "FID"
        .Cells(1, colPostSuperior) = "Superior"
        .Cells(1, colPostVariant) = "Variant"
        .Cells(1, colPostObjDep) = "Obj" & vbLf & "Dep"
        .Cells(1, colPostPPEQTY) = "PPE" & vbLf & "Qty"
        
        .Cells(1, colOpCode) = "Op" & vbLf & "Code"
        .Cells(1, colChangeCode) = "Change" & vbLf & "Code"
        .Cells(1, colProgressionCheck) = "Check"
    End With

End Sub

'--------------------------------------------------------------------------------
'---------------------------- Formatuj nowa tabele ------------------------------
'--------------------------------------------------------------------------------

Sub FormatNewEntries(iStart As Long, ByRef iEnd As Long)
    'formats newly created data set based on global variable LineType()
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    'LineType = 0 - standard line
    'LineType = 1 - merged line in Old Config Chart with value other than "OR" or "Deleted"
    'LineType = 2 - merged line in Old Config Chart with SB number
    'LineType = 3 - line containing at least one cell for additional formatting
    
    'LineType = 0 - no additional formatting
    'LineType = 1 - .Cells(i, colName).Font.Bold = True,
    'LineType = 1 - .Range(.Cells(i, 1), .Cells(i, colLast)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    'LineType = 2 - LineType 1 formatting + .Cells(i, colName).Font.Color = vbRed
    'LineType = 3 - values "OR", "Deleted", "--" to be centered; qty="X" to be changed to 1 and font color to blue
    
    With Sheet2
        For i = 1 To iEnd - iStart + 1
        
            Select Case LineType(i)
            Case 1
                j = iStart + i - 1
                .Cells(j, colName).Font.Bold = True
                .Range(.Cells(j, 1), .Cells(j, colLast)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Case 2
                j = iStart + i - 1
                .Cells(j, colName).Font.Bold = True
                .Range(.Cells(j, 1), .Cells(j, colLast)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Cells(j, colName).Font.Color = vbRed
            Case 3
                j = iStart + i - 1
                For k = 1 To colLast
                    Select Case NewCCArr(i, k)
                    Case "OR", "Deleted", "--"
                        .Cells(j, k).HorizontalAlignment = xlCenter
                    Case "X"
                        If k = colPreQTY Or k = colPostQTY Then
                            .Cells(j, k) = 1
                            .Cells(j, k).Font.Color = vbBlue
                        End If
                    End Select
                Next k
            Case 4
                .Cells(j, colName).EntireRow.Delete
                i = i - 1
                iEnd = iEnd - 1
            End Select
        Next i
    End With

End Sub

Sub FormatNewConfigChart()
    'additional formatting after New config chart is created and loaded to spreadsheet

    Dim rng As Range
    Dim j As Byte
    
    With Sheet2.Range("A1").CurrentRegion
    
        .Columns.HorizontalAlignment = xlCenter
        
        .Columns(colName).HorizontalAlignment = xlLeft
        .Columns(colPreATA).HorizontalAlignment = xlLeft
        .Columns(colPostATA).HorizontalAlignment = xlLeft
        
        .Rows(1).HorizontalAlignment = xlCenter
        .Rows(1).Font.Bold = True
        
        .Columns(colPrePN).EntireColumn.Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Columns(colPostPN).EntireColumn.Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Columns(colOpCode).EntireColumn.Borders(xlEdgeLeft).LineStyle = xlContinuous
        
        .Columns(colPrePN).Font.Color = 10498160
        .Columns(colPostPN).Font.Color = 10498160
        .Columns(colPreATA).Font.Color = 11892015
        .Columns(colPostATA).Font.Color = 11892015
        .Columns(colPreQTY).Font.Color = 5287936
        .Columns(colPostQTY).Font.Color = 5287936
        Union(.Columns(colPreFIDNo), .Columns(colPreSuperiorNo), .Columns(colPreVariantNo), _
            .Columns(colPostFIDNo), .Columns(colPostSuperiorNo), .Columns(colPostVariantNo)).Font.Color = 10921638
        
        'delete autofit
        For j = 1 To colLast
            .Columns(j).EntireColumn.ColumnWidth = 100
            .Columns(j).EntireColumn.AutoFit
            'If .Columns(j).EntireColumn.ColumnWidth > 100 Then .Columns(j).EntireColumn.ColumnWidth = 100
            .Columns(j).EntireColumn.ColumnWidth = .Columns(j).EntireColumn.ColumnWidth * 1.1
        Next j

        'add width for each column
        .Columns(colName).EntireColumn.ColumnWidth = 35
        .Columns(colSIN).EntireColumn.ColumnWidth = 10
        .Columns(colPrePN).EntireColumn.ColumnWidth = 20
        .Columns(colPostPN).EntireColumn.ColumnWidth = 20
        .Columns(colPreATA).EntireColumn.ColumnWidth = 20
        .Columns(colPostATA).EntireColumn.ColumnWidth = 20
        .Columns(colPreQTY).EntireColumn.ColumnWidth = 5
        .Columns(colPostQTY).EntireColumn.ColumnWidth = 5
        .Columns(colOpCode).EntireColumn.ColumnWidth = 10
        .Columns(colChangeCode).EntireColumn.ColumnWidth = 10
        .Columns(colProgressionCheck).EntireColumn.ColumnWidth = 15
        
'-------- add conditional formatting --------------------
        Dim condition1 As FormatCondition
        Dim condition1b As FormatCondition
        Dim condition2 As FormatCondition
        Dim condition3 As FormatCondition
        Dim condition4 As FormatCondition
        Dim condition5 As FormatCondition
        Dim condition6 As FormatCondition
        
        Dim CondFormula1 As String
        Dim CondFormula1b As String
        Dim CondFormula2 As String
        Dim CondFormula3 As String
        Dim CondFormula4 As String
        Dim CondFormula56 As String
        
'       semicolon delimited formulas (can be commented out in dependence on system configuration)
'       some API would need to be loaded to get this information from registry
        
        'FID ok
        CondFormula1 = """2X_"" & MID(E1;1;2) & MID(E1;4;2) & MID(E1;7;2) & MID(E1;10;2) & MID(E1;13;2) & 0"
        'dummy FID for clamps ok
        CondFormula1b = """2X_"" & MID(E1;1;2) & MID(E1;4;2) & MID(E1;7;2) & MID(E1;10;2) & MID(E1;13;2) & ""0_"""
        'Variant ok
        CondFormula2 = """2X"" & MID(E1;11;1) & MID(E1;13;LEN(E1)-12)"
        'Qty ok
        CondFormula3 = "text(F1;""#"")"
        'Qty ok
        CondFormula4 = "F1"
        'FID; variant; qty ok
        CondFormula56 = "AND(" & .Cells(1, colPreFID).Address(False, False) & "=" & CondFormula1 & ";" _
                                & .Cells(1, colPreVariant).Address(False, False) & "=" & CondFormula2 & ";" _
                                & "OR(" & .Cells(1, colPrePPEQTY).Address(False, False) & "=" & CondFormula3 & ";" _
                                & .Cells(1, colPrePPEQTY).Address(False, False) & "=" & CondFormula4 & "))"
                                
'       comma delimited formulas (can be commented out in dependence on system configuration)
'       some API would need to be loaded to get this information from registry
'
'        'FID ok
'        CondFormula1 = """2X_"" & MID(E1,1,2) & MID(E1,4,2) & MID(E1,7,2) & MID(E1,10,2) & MID(E1,13,2) & 0"
'        'dummy FID for clamps ok
'        CondFormula1b = """2X_"" & MID(E1,1,2) & MID(E1,4,2) & MID(E1,7,2) & MID(E1,10,2) & MID(E1,13,2) & ""0_"""
'        'Variant ok
'        CondFormula2 = """2X"" & MID(E1,11,1) & MID(E1,13,LEN(E1)-12)"
'        'Qty ok
'        CondFormula3 = "text(F1,""#"")"
'        'Qty ok
'        CondFormula4 = "F1"
'        'FID, variant, qty ok
'        CondFormula56 = "AND(" & .Cells(1, colPreFID).Address(False, False) & "=" & CondFormula1 & "," _
'                                & .Cells(1, colPreVariant).Address(False, False) & "=" & CondFormula2 & "," _
'                                & "OR(" & .Cells(1, colPrePPEQTY).Address(False, False) & "=" & CondFormula3 & "," _
'                                & .Cells(1, colPrePPEQTY).Address(False, False) & "=" & CondFormula4 & "))"
        
        .FormatConditions.Delete
        
        Set condition1 = Union(.Columns(colPreFID), .Columns(colPostFID)).Cells.FormatConditions.Add _
            (xlCellValue, xlEqual, "=" & CondFormula1)
        Set condition1b = Union(.Columns(colPreFID), .Columns(colPostFID)).Cells.FormatConditions.Add _
            (xlCellValue, xlEqual, "=" & CondFormula1b)
        Set condition2 = Union(.Columns(colPreVariant), .Columns(colPostVariant)).Cells.FormatConditions.Add _
            (xlCellValue, xlEqual, "=" & CondFormula2)
        Set condition3 = Union(.Columns(colPrePPEQTY), .Columns(colPostPPEQTY)).Cells.FormatConditions.Add _
            (xlCellValue, xlEqual, "=" & CondFormula3)
        Set condition4 = Union(.Columns(colPrePPEQTY), .Columns(colPostPPEQTY)).Cells.FormatConditions.Add _
            (xlCellValue, xlEqual, "=" & CondFormula4)
        Set condition5 = Union(.Columns(colPreSuperior), .Columns(colPostSuperior)).Cells.FormatConditions.Add _
            (xlExpression, xlEqual, "=" & CondFormula56)
        Set condition6 = Union(.Columns(colPreObjDep), .Columns(colPostObjDep)).Cells.FormatConditions.Add _
            (xlExpression, xlEqual, "=" & CondFormula56)

        Dim condColor As String
        condColor = 10921638
        
        condition1.Font.Color = condColor
        condition1b.Font.Color = condColor
        condition2.Font.Color = condColor
        condition3.Font.Color = condColor
        condition4.Font.Color = condColor
        condition5.Font.Color = condColor
        condition6.Font.Color = condColor
        
'-------- grouping the columns ---------------------
        .Rows.AutoFit
        .Columns(colSBNo).EntireColumn.Group
        .Range(.Columns(colPreFIDNo), .Columns(colPrePPEQTY)).EntireColumn.Group
        .Range(.Columns(colPostFIDNo), .Columns(colPostPPEQTY)).EntireColumn.Group
        .Parent.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
        
    End With
    
    Set rng = Nothing
    Set condition1 = Nothing
    Set condition1b = Nothing
    Set condition2 = Nothing
    Set condition3 = Nothing
    Set condition4 = Nothing
    Set condition5 = Nothing
    Set condition6 = Nothing
    
End Sub

Sub MakeIndentions(arr() As Variant)
    'makes indentions in 'Name' column to show the straucture levels - based on dots from PN column in oryginal document
    '(dots are deleted in 'New Conf. Chart' spreadsheet after this step)

    Dim lIndention As Byte 'number of spaces for single indention
    lIndention = 4
    
    Dim i As Long
    Dim j As Long
    Dim PNcol As Long
    Dim Lvl As Byte
    Dim TempPN As String
    
    For i = 1 To UBound(arr, 2)
        
        If arr(colPrePN, i) <> "" Then
            
            'if there is no PN in 'PrePN' column check 'PostPN' column
            If arr(colPrePN, i) <> "--" Then
                PNcol = colPrePN
            Else
                PNcol = colPostPN
            End If
            
            'to extract first PN from list of PNs - in case there is more than one PN in single cell
            TempPN = CStr(arr(PNcol, i))
            If InStr(TempPN, vbLf) > 0 Then
                TempPN = Left(TempPN, InStr(TempPN, vbLf) - 1)
            End If
            
            'find the level of indention
            Lvl = CountSignInText(".", TempPN)
            
            'make indention in the 'Name' column
            arr(colName, i) = String(Lvl * lIndention, " ") & arr(colName, i)
            arr(colName, i) = Replace(arr(colName, i), vbLf, vbLf & String(Lvl * lIndention, " "))
            arr(colPrePN, i) = Replace(arr(colPrePN, i), ".", "")
            arr(colPostPN, i) = Replace(arr(colPostPN, i), ".", "")
        End If
    Next i

End Sub

Private Sub CorrectPNsInNewConfigChart(ByRef arr() As Variant)

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim sPN As String
    Dim PNcol As Long
    Dim ATAcol As Long
    
    For k = 1 To 2
    
        If k = 1 Then PNcol = colPrePN: ATAcol = colPreATA
        If k = 2 Then PNcol = colPostPN: ATAcol = colPostATA
        
        For i = 1 To UBound(arr, 2)
            
            sPN = arr(PNcol, i)
            
            'split if PN connected with ATA chapter
            If sPN Like "*??-??-??-??*" Then
                If InStr(sPN, vbLf) = 0 And Len(sPN) > 18 And arr(ATAcol, i) = "--" Then
                    For j = 1 To Len(sPN) - 10
                        If Mid(sPN, j, 11) Like "??-??-??-??" Then Exit For
                    Next j
                    If j < Len(sPN) - 10 Then
                        arr(PNcol, i) = Left(sPN, j - 1)
                        arr(ATAcol, i) = Right(sPN, Len(sPN) - j + 1)
                    End If
                End If
            End If
            
            'split if connected with alternative PN in brackets
            If sPN Like "*(*)" Then
                arr(PNcol, i) = Left(sPN, InStr(sPN, "(") - 1) & vbLf & Mid(sPN, InStr(sPN, "(") + 1, InStr(sPN, ")") - InStr(sPN, "(") - 1)
            End If
            
        Next i
    Next k
    
End Sub

'-----------------------------------------------------------------------------------------------------------------
'------ Union of top cells from each modification line (mod line can consist of multiple merged cells) -----------
'-----------------------------------------------------------------------------------------------------------------

Private Function ItemListRange() As Range
    
    Dim i As Integer
    Dim iLastRow As Single
    Dim TargetCell As Range

    With Sheet1
    
        i = 1
        iLastRow = .Cells(1048576, 5).End(xlUp).Row 'column 5 choosed cause pre qty cell is always merged for each individual modification line
        
        Do While i <= iLastRow
            
            If .Cells(i, 5) <> "" Then
                Set TargetCell = .Cells(i, 5)
            Else
                Set TargetCell = .Cells(i, 1)
            End If
            
            If ItemListRange Is Nothing Then
                Set ItemListRange = TargetCell
            Else
                On Error GoTo Er
                Set ItemListRange = Union(ItemListRange, TargetCell)
                On Error GoTo 0
            End If
            
            If Cells(i, 5) <> "" And Cells(i, 5).MergeCells Then
                i = i + Cells(i, 5).MergeArea.Rows.Count
            Else
                i = i + 1
            End If
        Loop
        
    End With
    
    Set TargetCell = Nothing
    
Exit Function
Er:
MsgBox "Error in row: " & i
End Function

'----------------------------------------------------------------------------------------------
'-------- Function to get the information from certain fields in old coniguration chart -------
'----------------------------------------------------------------------------------------------

Private Function OldPartNumber(ByVal r As Range) As String
    
    Dim i As Integer
    Dim rng As Range
    Dim Gmin As Integer
    Dim Gmax As Integer
    Dim TempPN As String
    Dim temp As String
    
    OldPartNumber = Replace(OldPartNumberRange(r).Item(1).Value, " ", "")
    TempPN = Replace(OldPartNumber, " ", "")
    TempPN = Replace(TempPN, vbLf, "")
    
    If Right(OldPartNumber, 1) = "/" Then OldPartNumber = Left(OldPartNumber, Len(OldPartNumber) - 1)
    
    If TempPN Like "*GEnx-2B67*G??/*" Or TempPN Like "*GENX-2B67*G??/*" Then
        temp = Right(TempPN, Len(TempPN) - InStr(InStr(TempPN, "2B67") + 5, TempPN, "/") + 1)
        TempPN = Left(TempPN, InStr(InStr(TempPN, "2B67") + 5, TempPN, "/") - 1)
        Do While temp Like "G??*" Or temp Like "/G??*"
            If temp Like "/G??*" Then temp = Right(temp, Len(temp) - 1)
            TempPN = TempPN & vbLf & Left(TempPN, InStr(InStr(TempPN, "2B67"), TempPN, "G") - 1) & Left(temp, 3)
            temp = Right(temp, Len(temp) - 3)
        Loop
        OldPartNumber = TempPN
        Exit Function
    End If
    
    If OldPartNumber Like "*????M??G??/*" Or OldPartNumber Like "*????M??P??/*" Then
        temp = Right(OldPartNumber, Len(OldPartNumber) - InStr(OldPartNumber, "/") + 1)
        OldPartNumber = Left(OldPartNumber, InStr(OldPartNumber, "/") - 1)
        Do While temp Like "G??*" Or temp Like "/G??*"
            If temp Like "/G??*" Then temp = Right(temp, Len(temp) - 1)
            OldPartNumber = OldPartNumber & vbLf & Left(OldPartNumber, InStr(OldPartNumber, "G") - 1) & Left(temp, 3)
            temp = Right(temp, Len(temp) - 3)
        Loop
        Do While temp Like "P??*" Or temp Like "/P??*"
            If temp Like "/P??*" Then temp = Right(temp, Len(temp) - 1)
            OldPartNumber = OldPartNumber & vbLf & Left(OldPartNumber, InStr(OldPartNumber, "P") - 1) & Left(temp, 3)
            temp = Right(temp, Len(temp) - 3)
        Loop
    End If
    
    If OldPartNumber Like "*????M??G??-G??" Then
        Gmin = CInt(Mid(OldPartNumber, InStr(OldPartNumber, "-") - 2, 2))
        Gmax = CInt(Mid(OldPartNumber, InStr(OldPartNumber, "-") + 2, 2))
        temp = Mid(OldPartNumber, InStr(OldPartNumber, "-") - 10, 8)
        OldPartNumber = ""
        For i = Gmin To Gmax
            OldPartNumber = OldPartNumber & temp & Format(i, "00")
            If i <> Gmax Then OldPartNumber = OldPartNumber & vbLf
        Next i
        Exit Function
    End If
    
    If r.MergeCells Then
        If OldPartNumber Like "*????M??G??*" Or OldPartNumber Like "*????M??P??*" Then
            For i = 2 To r.MergeArea.Rows.Count
                temp = Replace(OldPartNumberRange(r).Item(i).Value, " ", "")
                Do While temp Like "G??*" Or temp Like "/G??*"
                    If temp Like "/G??*" Then temp = Right(temp, Len(temp) - 1)
                    OldPartNumber = OldPartNumber & vbLf & Left(OldPartNumber, InStr(OldPartNumber, "G") - 1) & Left(temp, 3)
                    temp = Right(temp, Len(temp) - 3)
                Loop
                Do While temp Like "P??*" Or temp Like "/P??*"
                    If temp Like "/P??*" Then temp = Right(temp, Len(temp) - 1)
                    OldPartNumber = OldPartNumber & vbLf & Left(OldPartNumber, InStr(OldPartNumber, "P") - 1) & Left(temp, 3)
                    temp = Right(temp, Len(temp) - 3)
                Loop
            Next i
        Else
            For Each rng In OldPartNumberRange(r)
                If rng.Value Like "*????M??P??*" Or rng.Value Like "*????M??G??*" Then
                    temp = Replace(rng.Value, " ", "")
                    If Left(temp, 1) = "(" Then temp = Right(temp, Len(temp) - 1)
                    If Right(temp, 1) = ")" Then temp = Left(temp, Len(temp) - 1)
                    OldPartNumber = OldPartNumber & vbLf & temp
                End If
            Next rng
        End If
    End If
    
    OldPartNumber = Replace(OldPartNumber, vbCrLf, vbLf)
    OldPartNumber = Replace(OldPartNumber, Chr(42), "")
    OldPartNumber = Replace(OldPartNumber, "0.", ".")

    Set rng = Nothing
    
End Function

Private Function NewPartNumber(ByVal r As Range) As String
    
    Dim i As Integer
    Dim rng As Range
    Dim Gmin As Integer
    Dim Gmax As Integer
    Dim TempPN As String
    Dim temp As String
    
    NewPartNumber = Replace(NewPartNumberRange(r).Item(1).Value, " ", "")
    TempPN = Replace(NewPartNumber, " ", "")
    TempPN = Replace(TempPN, vbLf, "")
    
    If Right(NewPartNumber, 1) = "/" Then NewPartNumber = Left(NewPartNumber, Len(NewPartNumber) - 1)
    
    If TempPN Like "*GEnx-2B67*G??/*" Or TempPN Like "*GENX-2B67*G??/*" Then
        temp = Right(TempPN, Len(TempPN) - InStr(InStr(TempPN, "2B67") + 5, TempPN, "/") + 1)
        TempPN = Left(TempPN, InStr(InStr(TempPN, "2B67") + 5, TempPN, "/") - 1)
        Do While temp Like "G??*" Or temp Like "/G??*"
            If temp Like "/G??*" Then temp = Right(temp, Len(temp) - 1)
            TempPN = TempPN & vbLf & Left(TempPN, InStr(InStr(TempPN, "2B67"), TempPN, "G") - 1) & Left(temp, 3)
            temp = Right(temp, Len(temp) - 3)
        Loop
        NewPartNumber = TempPN
        Exit Function
    End If
    
    If NewPartNumber Like "*????M??G??/*" Or NewPartNumber Like "*????M??P??/*" Then
        temp = Right(NewPartNumber, Len(NewPartNumber) - InStr(NewPartNumber, "/") + 1)
        NewPartNumber = Left(NewPartNumber, InStr(NewPartNumber, "/") - 1)
        Do While temp Like "G??*" Or temp Like "/G??*"
            If temp Like "/G??*" Then temp = Right(temp, Len(temp) - 1)
            NewPartNumber = NewPartNumber & vbLf & Left(NewPartNumber, InStr(NewPartNumber, "G") - 1) & Left(temp, 3)
            temp = Right(temp, Len(temp) - 3)
        Loop
        Do While temp Like "P??*" Or temp Like "/P??*"
            If temp Like "/P??*" Then temp = Right(temp, Len(temp) - 1)
            NewPartNumber = NewPartNumber & vbLf & Left(NewPartNumber, InStr(NewPartNumber, "P") - 1) & Left(temp, 3)
            temp = Right(temp, Len(temp) - 3)
        Loop
    End If
    
    If NewPartNumber Like "*????M??G??-G??" Then
        Gmin = CInt(Mid(NewPartNumber, InStr(NewPartNumber, "-") - 2, 2))
        Gmax = CInt(Mid(NewPartNumber, InStr(NewPartNumber, "-") + 2, 2))
        temp = Mid(NewPartNumber, InStr(NewPartNumber, "-") - 10, 8)
        NewPartNumber = ""
        For i = Gmin To Gmax
            NewPartNumber = NewPartNumber & temp & Format(i, "00")
            If i <> Gmax Then NewPartNumber = NewPartNumber & vbLf
        Next i
        Exit Function
    End If
    
    If r.MergeCells Then
        If NewPartNumber Like "*????M??G??*" Or NewPartNumber Like "*????M??P??*" Then
            For i = 2 To r.MergeArea.Rows.Count
                temp = Replace(NewPartNumberRange(r).Item(i).Value, " ", "")
                Do While temp Like "G??*" Or temp Like "/G??*"
                    If temp Like "/G??*" Then temp = Right(temp, Len(temp) - 1)
                    NewPartNumber = NewPartNumber & vbLf & Left(NewPartNumber, InStr(NewPartNumber, "G") - 1) & Left(temp, 3)
                    temp = Right(temp, Len(temp) - 3)
                Loop
                Do While temp Like "P??*" Or temp Like "/P??*"
                    If temp Like "/P??*" Then temp = Right(temp, Len(temp) - 1)
                    NewPartNumber = NewPartNumber & vbLf & Left(NewPartNumber, InStr(NewPartNumber, "P") - 1) & Left(temp, 3)
                    temp = Right(temp, Len(temp) - 3)
                Loop
            Next i
        Else
            For Each rng In NewPartNumberRange(r)
                If rng.Value Like "*????M??P??*" Or rng.Value Like "*????M??G??*" Then
                    temp = Replace(rng.Value, " ", "")
                    If Left(temp, 1) = "(" Then temp = Right(temp, Len(temp) - 1)
                    If Right(temp, 1) = ")" Then temp = Left(temp, Len(temp) - 1)
                    NewPartNumber = NewPartNumber & vbLf & temp
                End If
            Next rng
        End If
    End If
    
    NewPartNumber = Replace(NewPartNumber, vbCrLf, vbLf)
    NewPartNumber = Replace(NewPartNumber, Chr(42), "")
    NewPartNumber = Replace(NewPartNumber, "0.", ".")
    
    Set rng = Nothing
    
End Function

Private Function OldATA(ByVal r As Range) As String
    
    Dim rng As Range
    Dim temp As String
    
    OldATA = "--"
    
    If r.MergeCells Then
        For Each rng In OldPartNumberRange(r)
            temp = Replace(rng.Value, " ", "")
            temp = Replace(rng.Value, Chr(160), "")
            If temp Like "??-??-??*" Then
                If OldATA = "--" Then
                    OldATA = temp
                Else
                    OldATA = OldATA & vbLf & temp
                End If
            ElseIf temp Like "??-??*,??-??-??" Then
                If OldATA = "--" Then
                    OldATA = Right(temp, 8) & "-" & Left(temp, InStr(temp, ",") - 1)
                Else
                    OldATA = OldATA & vbLf & Right(temp, 8) & "-" & Left(temp, InStr(temp, ",") - 1)
                End If
            End If
        Next rng
    End If
    
    OldATA = Replace(OldATA, vbCrLf, vbLf)
    
    Set rng = Nothing
    
End Function

Private Function NewATA(ByVal r As Range) As String
    
    Dim rng As Range
    Dim temp As String
    
    NewATA = "--"
    
    If r.MergeCells Then
        For Each rng In NewPartNumberRange(r)
            temp = Replace(rng.Value, " ", "")
            temp = Replace(rng.Value, Chr(160), "")
            If temp Like "??-??-??*" Then
                If NewATA = "--" Then
                    NewATA = temp
                Else
                    NewATA = NewATA & vbLf & temp
                End If
            ElseIf temp Like "??-??*,??-??-??" Then
                If NewATA = "--" Then
                    NewATA = Right(temp, 8) & "-" & Left(temp, InStr(temp, ",") - 1)
                Else
                    NewATA = NewATA & vbLf & Right(temp, 8) & "-" & Left(temp, InStr(temp, ",") - 1)
                End If
            End If
        Next rng
    End If
    
    NewATA = Replace(NewATA, vbCrLf, vbLf)
    
    Set rng = Nothing
    
End Function

Private Function OldQty(ByVal r As Range) As String
    
    If Left(OldQtyRange(r).Item(1).Value, 1) = "-" Then
        OldQty = Right(OldQtyRange(r).Item(1).Value, Len(OldQtyRange(r).Item(1).Value) - 1)
    Else
        OldQty = OldQtyRange(r).Item(1).Value
        If Left(OldQty, 1) = "(" Then OldQty = Right(OldQty, Len(OldQty) - 1)
        If Right(OldQty, 1) = ")" Then OldQty = Left(OldQty, Len(OldQty) - 1)
    End If
    
End Function

Private Function NewQty(ByVal r As Range) As String
    
    If Left(NewQtyRange(r).Item(1).Value, 1) = "-" Then
        NewQty = Right(NewQtyRange(r).Item(1).Value, Len(NewQtyRange(r).Item(1).Value) - 1)
    Else
        NewQty = NewQtyRange(r).Item(1).Value
        If Left(NewQty, 1) = "(" Then NewQty = Right(NewQty, Len(NewQty) - 1)
        If Right(NewQty, 1) = ")" Then NewQty = Left(NewQty, Len(NewQty) - 1)
    End If
    
End Function

Private Function Name(ByVal r As Range) As String
    
    Name = Replace(PartNameRange(r).Item(1).Value, vbCrLf, vbLf)
    
    If Not PartNameRange(r).Find("New Name") Is Nothing Then
        If PartNameRange(r).Find("New Name").Offset(-1, 0) <> "" Then
            Name = Name & " (pre)" & vbLf & PartNameRange(r).Find("New Name").Offset(-1, 0) & " (post)"
        Else
            Name = Name & " (pre)" & vbLf & PartNameRange(r).Find("New Name").Offset(-2, 0) & " (post)"
        End If
    End If
    
End Function

Private Function PartSIN(ByVal r As Range) As String
    
    Dim rng As Range
    
    PartSIN = ""
    
    If r.MergeCells Then
        For Each rng In PartNameRange(r)
            If rng.Value Like "(SIN*)" Then
                PartSIN = Mid(rng.Value, 6, Len(rng.Value) - 6)
                Exit For
            End If
        Next rng
    End If
    
    Set rng = Nothing
    
End Function

Private Function OpCode(ByVal r As Range) As String
    
    Dim rng As Range
    
    OpCode = ""
    
    For Each rng In OpCodeRange(r)
        OpCode = OpCode & rng.Value
    Next rng
    
    Set rng = Nothing
    
End Function

Private Function ChangeCode(ByVal r As Range) As String

    ChangeCode = ChangeCodeRange(r).Item(1).Value
    
End Function

'--------------------------------------------------------------------------------------
'------ Functions to get references to certain fields in old configuration chart ------
'--------------------------------------------------------------------------------------

Private Function OldPartNumberRange(ByVal r As Range) As Range
    If r.MergeCells Then
        Set OldPartNumberRange = Range(r.MergeArea.Item(1).Offset(0, -1).Address & ":" & r.MergeArea.Item(r.MergeArea.Rows.Count).Offset(0, -1).Address)
    Else
        Set OldPartNumberRange = r.Offset(0, -1)
    End If
End Function

Private Function NewPartNumberRange(ByVal r As Range) As Range
    If r.MergeCells Then
        Set NewPartNumberRange = Range(r.MergeArea.Item(1).Offset(0, -4).Address & ":" & r.MergeArea.Item(r.MergeArea.Rows.Count).Offset(0, -4).Address)
    Else
        Set NewPartNumberRange = r.Offset(0, -4)
    End If
End Function

Private Function PartNameRange(ByVal r As Range) As Range
    If r.MergeCells Then
        Set PartNameRange = Range(r.MergeArea.Item(1).Offset(0, -2).Address & ":" & r.MergeArea.Item(r.MergeArea.Rows.Count).Offset(0, -2).Address)
    Else
        Set PartNameRange = r.Offset(0, -2)
    End If
End Function

Private Function OldQtyRange(ByVal r As Range) As Range
    If r.MergeCells Then
        Set OldQtyRange = Range(r.MergeArea.Address)
    Else
        Set OldQtyRange = r
    End If
End Function

Private Function NewQtyRange(ByVal r As Range) As Range
    If r.MergeCells Then
        Set NewQtyRange = Range(r.MergeArea.Item(1).Offset(0, -3).Address & ":" & r.MergeArea.Item(r.MergeArea.Rows.Count).Offset(0, -3).Address)
    Else
        Set NewQtyRange = r.Offset(0, -3)
    End If
End Function

Private Function OpCodeRange(ByVal r As Range) As Range
    If r.MergeCells Then
        Set OpCodeRange = Range(r.MergeArea.Item(1).Offset(0, 1).Address & ":" & r.MergeArea.Item(r.MergeArea.Rows.Count).Offset(0, 1).Address)
    Else
        Set OpCodeRange = r.Offset(0, 1)
    End If
End Function

Private Function ChangeCodeRange(ByVal r As Range) As Range
    If r.MergeCells Then
        Set ChangeCodeRange = Range(r.MergeArea.Item(1).Offset(0, 2).Address & ":" & r.MergeArea.Item(r.MergeArea.Rows.Count).Offset(0, 2).Address)
    Else
        Set ChangeCodeRange = r.Offset(0, 2)
    End If
End Function

Private Function Compare(ByVal r As Range, ByVal ValueToCompare As String) As Boolean
        If r.Value Like ValueToCompare Then Compare = True
End Function
