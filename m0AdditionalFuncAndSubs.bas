Attribute VB_Name = "m0AdditionalFuncAndSubs"
'===============================================================================================
'============================ Clear SB Configuration Chart Area ================================
'===============================================================================================
Sub ClearSBConfChartArea()

    Dim i As Integer
    Dim iLast As Integer

    With ThisWorkbook.Worksheets("SB Conf. Chart")
        .Columns("A:G").Cells.Clear
        .Columns("U:U").Cells.Clear
        .Columns.UseStandardWidth = True
        .Rows.UseStandardHeight = True
        .Columns("A:G").ColumnWidth = 15
        .Columns("A:G").Cells.NumberFormat = "@"
        .Columns("G:G").Borders(xlEdgeRight).LineStyle = xlContinuous

        .Cells(1, 1).Borders().LineStyle = xlContinuous
        '.Cells(1, 1).NumberFormat = "@"
        .Cells(1, 1).Value = "SB no"
        .Cells(1, 1).HorizontalAlignment = xlCenter
        .Cells(1, 2).Borders().LineStyle = xlContinuous
        '.Cells(1, 2).NumberFormat = "@"
        .Cells(1, 2).Value = "rev"
        .Cells(1, 2).HorizontalAlignment = xlCenter

        i = 2
        iLast = .Cells(1000, 20).End(xlUp).Row

        Do While i <= iLast
            If .Cells(i, 20).Font.Color = vbBlue Then .Cells(i, 20).Clear
            i = i + 1
        Loop

        .Columns("T:T").EntireColumn.Cells.VerticalAlignment = xlVAlignCenter
        .Columns("T:T").EntireColumn.Cells.HorizontalAlignment = xlCenter
        .Columns("U:U").EntireColumn.Cells.VerticalAlignment = xlVAlignCenter

    End With

End Sub

'===============================================================================================
'============ Appending arrays, Splitting New Config Chart lines to single-PN ones =============
'===============================================================================================

Sub AddRowToArray2D(ByRef arr() As Variant, ByVal lAfter As Long, ByVal lNoOfRowsToAdd As Long)
    'only second dimension of an array can be enlarged by this sub
    
    Dim m As Long
    Dim o As Long
    
    'add lNoOfRowsToAdd number of rows to the second dimension of an array
    ReDim Preserve arr(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2) + lNoOfRowsToAdd)
    
    For m = UBound(arr, 2) To lAfter + lNoOfRowsToAdd + 1 Step -1       'move the values to the end of an array
        For o = LBound(arr, 1) To UBound(arr, 1)                        'and delete them from old location
            arr(o, m) = arr(o, m - lNoOfRowsToAdd)                      'to make a 'space' for new entries
            arr(o, m - lNoOfRowsToAdd) = ""
        Next o
    Next m
End Sub

Sub AddRowToArray1D(ByRef arr() As Variant, ByVal lAfter As Long, ByVal lNoOfRowsToAdd As Long)
    'only one dimension array can be enlarged by this sub

    Dim m As Long
    
    'add lNoOfRowsToAdd number of elements to the first dimension of an array
    ReDim Preserve arr(LBound(arr) To UBound(arr) + lNoOfRowsToAdd)
    
    For m = UBound(arr) To lAfter + lNoOfRowsToAdd + 1 Step -1          'move the values to the end of an array
        arr(m) = arr(m - lNoOfRowsToAdd)                                'and delete them from old location
        arr(m - lNoOfRowsToAdd) = ""                                    'to make a 'space' for new entries
    Next m
End Sub

Sub SplitEasyOnes(ByRef NewCCArr() As Variant, ByRef LineTypeArr() As Variant, ByRef iEnd As Long)
    'Splits part numbers in 'Pre PN' and 'Post PN' columns cells in 'New Conf. Chart' worksheet in case there is more than one in single cell
    ' and case can be recognized by algorithm (in some specific cases man's decision is required)
    'Creates separate line in 'New Conf. Chart' worksheet for each Pre-Post PN pair in a way there is always only one PN in single cell
    'Works on 'NewCCArrTemp' array from m1CreateNewConfigChart module
    'Adds proper Line Type for newly created lines to LineType array from m1CreateNewConfigChart module
    'NewCCArr is pasted later on to New Conf. Chart worksheet and cells are formatted acc to LineType array

    Dim i As Long
    Dim j As Long
    Dim uboundInitial As Long
    uboundInitial = UBound(NewCCArr, 2)
    
    Dim PNcol1 As Long
    Dim PNcol2 As Long

    For j = 1 To 2
        
        'check PrePN cells vs PostPN cells first
        'then vice-versa
        If j = 1 Then
            PNcol1 = colPrePN
            PNcol2 = colPostPN
        ElseIf j = 2 Then
            PNcol1 = colPostPN
            PNcol2 = colPrePN
        End If
        
        'NewCCArr and LineTypeArr dimension 'i' can be enlarged by SplitArrLineOneToOne and/or SplitArrLineAnyToAny subs
        i = 1
        Do While i <= UBound(NewCCArr, 2)
        
            'if there is more than one PN in a cell...
            If InStr(NewCCArr(PNcol1, i), vbLf) > 0 Then
                
                'if pre and post configurations are the same (remain progression)
                If NewCCArr(PNcol1, i) = NewCCArr(PNcol2, i) And (NewCCArr(colOpCode, i) = "" Or NewCCArr(colOpCode, i) = "RM") Then
                    Call SplitArrLineOneToOne(NewCCArr, LineTypeArr, i)
                    
                'if there are two PNs in a cell from which one is a VIN number
                ElseIf CountSignInText(vbLf, NewCCArr(PNcol1, i)) = 1 And CountSignInText("VIN", NewCCArr(PNcol1, i)) = 1 Then
                    
                    'if in second column there are also only two PNs from which one is a VIN number
                    If CountSignInText(vbLf, NewCCArr(PNcol2, i)) = 1 And CountSignInText("VIN", NewCCArr(PNcol2, i)) = 1 Then
                        Call SplitArrLineOneToOne(NewCCArr, LineTypeArr, i)
                        
                    'if in second column PN is not specified
                    ElseIf NewCCArr(PNcol2, i) = "--" Then
                        Call SplitArrLineAnyToAny(NewCCArr, LineTypeArr, i)
                    End If
                
                'if in second column there is only one PN
                ElseIf InStr(NewCCArr(PNcol2, i), vbLf) = 0 Then
                    Call SplitArrLineAnyToAny(NewCCArr, LineTypeArr, i)
                End If
                
            End If
            i = i + 1
        Loop
    Next j
    
    'remove 'VIN' prefix from PNs in cells
    For i = 1 To UBound(NewCCArr, 2)
        NewCCArr(colPrePN, i) = Replace(NewCCArr(colPrePN, i), "VIN ", "")
        NewCCArr(colPrePN, i) = Replace(NewCCArr(colPrePN, i), "VIN", "")
        NewCCArr(colPostPN, i) = Replace(NewCCArr(colPostPN, i), "VIN ", "")
        NewCCArr(colPostPN, i) = Replace(NewCCArr(colPostPN, i), "VIN", "")
    Next i
    
    'calculate new ending row
    iEnd = iEnd + (UBound(NewCCArr, 2) - uboundInitial)
    
End Sub

Sub SplitArrLineOneToOne(ByRef NewCCArr() As Variant, ByRef LineTypeArr() As Variant, ByRef i As Long)
    'works in case of the same number of PNs in prePN column cell and postPN column cell
    'adds additional rows for distinguished data
    'ATA data is being adjusted accordingly
    '
    ' |A/B| - |C/D|  ->   |A| - |C|
    '                ->   |B| - |D|
    '
    
    Dim j As Long
    Dim k As Long

    Dim prePNLfNumber As Long
    Dim postPNLfNumber As Long
    Dim preATALfNumber As Long
    Dim postATALfNumber As Long

    'count line feed signs
    'lines where of all those would be equal to 0 are excluded by calling procedure
    prePNLfNumber = CountSignInText(vbLf, NewCCArr(colPrePN, i))
    postPNLfNumber = CountSignInText(vbLf, NewCCArr(colPostPN, i))
    preATALfNumber = CountSignInText(vbLf, NewCCArr(colPreATA, i))
    postATALfNumber = CountSignInText(vbLf, NewCCArr(colPostATA, i))
    
    'works if ATA chapters amount matches PN amount or if there is only one ATA chapter given (then it is assign to all PNs)
    If (preATALfNumber = 0 Or preATALfNumber = prePNLfNumber) And (postATALfNumber = 0 Or postATALfNumber = postPNLfNumber) Then
        
        Call AddRowToArray2D(NewCCArr, i, prePNLfNumber)
        Call AddRowToArray1D(LineTypeArr, i, prePNLfNumber)
        
        For j = i + 1 To i + prePNLfNumber
        
            For k = 1 To colLast
            
                Select Case k
                Case colPrePN, colPostPN
                    NewCCArr(k, j) = Right(NewCCArr(k, j - 1), Len(NewCCArr(k, j - 1)) - (InStr(NewCCArr(k, j - 1), vbLf) - 1) - Len(vbLf))
                    NewCCArr(k, j - 1) = Left(NewCCArr(k, j - 1), InStr(NewCCArr(k, j - 1), vbLf) - 1)
                Case colPreATA, colPostATA
                    If k = colPreATA And preATALfNumber = 0 Or k = colPostATA And postATALfNumber = 0 Then
                        NewCCArr(k, j) = NewCCArr(k, j - 1)
                    Else
                        NewCCArr(k, j) = Right(NewCCArr(k, j - 1), Len(NewCCArr(k, j - 1)) - (InStr(NewCCArr(k, j - 1), vbLf) - 1) - Len(vbLf))
                        NewCCArr(k, j - 1) = Left(NewCCArr(k, j - 1), InStr(NewCCArr(k, j - 1), vbLf) - 1)
                    End If
                Case Else
                    NewCCArr(k, j) = NewCCArr(k, j - 1)
                End Select
            
            Next k
            
            If NewCCArr(colPrePN, j) = "--" Or NewCCArr(colPostPN, j) = "--" Or NewCCArr(colName, j) = "OR" Or _
                NewCCArr(colName, j) = "Deleted" Or NewCCArr(colPreQTY, j) = "X" Or NewCCArr(colPostQTY, j) = "X" Then
                LineTypeArr(j) = 3
            Else
                LineTypeArr(j) = 0
            End If
            
        Next j
        
        i = i + prePNLfNumber
    End If

End Sub

Sub SplitArrLineAnyToAny(ByRef NewCCArr() As Variant, ByRef LineTypeArr() As Variant, ByRef i As Long)
    
    Dim j As Long
    Dim k As Long
    Dim m As Long
    Dim o As Long
    Dim iInit As Long
    iInit = i

    Dim prePNLfNumber As Long
    Dim postPNLfNumber As Long
    Dim preATALfNumber As Long
    Dim postATALfNumber As Long

    prePNLfNumber = CountSignInText(vbLf, NewCCArr(colPrePN, i))
    postPNLfNumber = CountSignInText(vbLf, NewCCArr(colPostPN, i))
    preATALfNumber = CountSignInText(vbLf, NewCCArr(colPreATA, i))
    postATALfNumber = CountSignInText(vbLf, NewCCArr(colPostATA, i))
    
    Dim sVal As String
    Dim PrePN() As Variant
    Dim PostPN() As Variant
    Dim PreATA() As Variant
    Dim PostATA() As Variant
    
    If (preATALfNumber = 0 Or preATALfNumber = prePNLfNumber) And (postATALfNumber = 0 Or postATALfNumber = postPNLfNumber) Then
        
        Call AddRowToArray2D(NewCCArr, i, (prePNLfNumber + 1) * (postPNLfNumber + 1) - 1)
        Call AddRowToArray1D(LineTypeArr, i, (prePNLfNumber + 1) * (postPNLfNumber + 1) - 1)
        
        ReDim PrePN(1 To prePNLfNumber + 1)
        ReDim PostPN(1 To postPNLfNumber + 1)
        ReDim PreATA(1 To prePNLfNumber + 1)
        ReDim PostATA(1 To postPNLfNumber + 1)
        
        sVal = NewCCArr(colPrePN, i) & vbLf
        For j = 1 To prePNLfNumber + 1
            PrePN(j) = Left(sVal, InStr(sVal, vbLf) - 1)
            sVal = Right(sVal, Len(sVal) - (InStr(sVal, vbLf) - 1) - Len(vbLf))
        Next j
        
        sVal = NewCCArr(colPostPN, i) & vbLf
        For j = 1 To postPNLfNumber + 1
            PostPN(j) = Left(sVal, InStr(sVal, vbLf) - 1)
            sVal = Right(sVal, Len(sVal) - (InStr(sVal, vbLf) - 1) - Len(vbLf))
        Next j
        
        If preATALfNumber = 0 Then
            For j = 1 To prePNLfNumber + 1
                PreATA(j) = NewCCArr(colPreATA, i)
            Next j
        Else
            sVal = NewCCArr(colPreATA, i) & vbLf
            For j = 1 To prePNLfNumber + 1
                PreATA(j) = Left(sVal, InStr(sVal, vbLf) - 1)
                sVal = Right(sVal, Len(sVal) - (InStr(sVal, vbLf) - 1) - Len(vbLf))
            Next j
        End If
        
        If postATALfNumber = 0 Then
            For j = 1 To postPNLfNumber + 1
                PostATA(j) = NewCCArr(colPostATA, i)
            Next j
        Else
            sVal = NewCCArr(colPostATA, i) & vbLf
            For j = 1 To postPNLfNumber + 1
                PostATA(j) = Left(sVal, InStr(sVal, vbLf) - 1)
                sVal = Right(sVal, Len(sVal) - (InStr(sVal, vbLf) - 1) - Len(vbLf))
            Next j
        End If
        
        For j = 1 To prePNLfNumber + 1
            For k = 1 To postPNLfNumber + 1
            
                If i > iInit Then
                    For o = 1 To colLast
                        NewCCArr(o, i) = NewCCArr(o, i - 1)
                    Next o
                End If
                
                NewCCArr(colPrePN, i) = PrePN(j)
                NewCCArr(colPostPN, i) = PostPN(k)
                NewCCArr(colPreATA, i) = PreATA(j)
                NewCCArr(colPostATA, i) = PostATA(k)

                If NewCCArr(colPrePN, i) = "--" Or NewCCArr(colPostPN, i) = "--" Or NewCCArr(colName, i) = "OR" Or _
                    NewCCArr(colName, i) = "Deleted" Or NewCCArr(colPreQTY, i) = "X" Or NewCCArr(colPostQTY, i) = "X" Then
                    LineTypeArr(i) = 3
                Else
                    LineTypeArr(i) = 0
                End If
                
                i = i + 1
            Next k
        Next j
        
        i = i - 1
    End If
    
End Sub

Public Function PNlong(ByVal sPN As String) As String
    Dim wsMM As Worksheet
    Set wsMM = ThisWorkbook.Worksheets("MM data")
    
    Dim VLookUpRange As Range
    Set VLookUpRange = Range(wsMM.Cells(2, 1), wsMM.Cells(2, 2).End(xlDown))
    
    With Application.WorksheetFunction
        If .CountIf(VLookUpRange.Columns(1), sPN) = 0 Then PNlong = "": Exit Function
        PNlong = .IfError(.VLookup(sPN, VLookUpRange, 2, 0), "")
    End With
    
    Set wsMM = Nothing
    Set VLookUpRange = Nothing
    
End Function

Function CountSignInText(ByVal sSign As String, ByVal sText As String) As Integer

    Dim i As Integer
    If sSign = "" Or sText = "" Then Exit Function
    
    CountSignInText = (Len(sText) - Len(Replace(sText, sSign, ""))) / Len(sSign)
        
End Function
