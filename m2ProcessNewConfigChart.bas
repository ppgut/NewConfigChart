Attribute VB_Name = "m2ProcessNewConfigChart"
Option Explicit

Sub MoveToNewFile()

    Dim wbNew As Workbook

    Application.Calculation = xlCalculationManual
    
    With ThisWorkbook.Worksheets("New Conf. Chart")
        If .AutoFilterMode Then .AutoFilterMode = False
        .Columns.ClearOutline
        Set wbNew = Application.Workbooks.Add
        .Range("A1").CurrentRegion.Copy
        With wbNew.Worksheets(1)
            .Range("A1").PasteSpecial xlPasteColumnWidths
            .Range("A1").PasteSpecial xlPasteFormats
            .Range("A1").PasteSpecial xlPasteValues
            .Columns(colSBNo).EntireColumn.Group
            .Range(.Columns(colPreFIDNo), .Columns(colPrePPEQTY)).EntireColumn.Group
            .Range(.Columns(colPostFIDNo), .Columns(colPostPPEQTY)).EntireColumn.Group
            .Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
            .Range("A1").Select
        End With
        .Columns(colSBNo).EntireColumn.Group
        .Range(.Columns(colPreFIDNo), .Columns(colPrePPEQTY)).EntireColumn.Group
        .Range(.Columns(colPostFIDNo), .Columns(colPostPPEQTY)).EntireColumn.Group
        .Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
        Application.CutCopyMode = False
    End With
    
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Sub SplitRowOneToOne()

    If Selection.Rows.Count > 1 Then
        MsgBox "Select only one row"
        Exit Sub
    End If

    Dim iRow As Integer
    Dim i As Byte
    Dim j As Byte
    Dim n As Byte
    
    Dim PrePNarr() As Variant
    Dim PreATAarr() As Variant
    Dim PostPNarr() As Variant
    Dim PostATAarr() As Variant
    
    iRow = Selection.Row

    PrePNarr = Split(Cells(iRow, colPrePN).Value)
    PreATAarr = Split(Cells(iRow, colPreATA))
    PostPNarr = Split(Cells(iRow, colPostPN))
    PostATAarr = Split(Cells(iRow, colPostATA))
    
    'jesli jest 1 PrePN i kilka PreATA - skopiuj PrePN do ilosci PreATA
    'jesli jest 1 PreATA i kilka PrePN - skopiuj PreATA do ilosci PrePN
    If UBound(PrePNarr) <> UBound(PreATAarr) Then
        If UBound(PrePNarr) = 0 Then
            Do While UBound(PreATAarr) <> UBound(PrePNarr)
                ReDim Preserve PrePNarr(UBound(PrePNarr) + 1)
                PrePNarr(UBound(PrePNarr)) = PrePNarr(UBound(PrePNarr) - 1)
            Loop
        ElseIf UBound(PreATAarr) = 0 Then
            Do While UBound(PreATAarr) <> UBound(PrePNarr)
                ReDim Preserve PreATAarr(UBound(PreATAarr) + 1)
                PreATAarr(UBound(PreATAarr)) = PreATAarr(UBound(PreATAarr) - 1)
            Loop
        Else
            MsgBox "Different number of PrePN and PreATA"
            Exit Sub
        End If
    End If
    
    'jesli jest 1 PostPN i kilka PostATA - skopiuj PostPN do ilosci PostATA
    'jesli jest 1 PostATA i kilka PostPN - skopiuj PostATA do ilosci PostPN
    If UBound(PostPNarr) <> UBound(PostATAarr) Then
        If UBound(PostPNarr) = 0 Then
            Do While UBound(PostPNarr) <> UBound(PostATAarr)
                ReDim Preserve PostPNarr(UBound(PostPNarr) + 1)
                PostPNarr(UBound(PostPNarr)) = PostPNarr(UBound(PostPNarr) - 1)
            Loop
        ElseIf UBound(PostATAarr) = 0 Then
            Do While UBound(PostATAarr) <> UBound(PostPNarr)
                ReDim Preserve PostATAarr(UBound(PostATAarr) + 1)
                PostATAarr(UBound(PostATAarr)) = PostATAarr(UBound(PostATAarr) - 1)
            Loop
        Else
            MsgBox "Different number of PostPN and PostATA"
            Exit Sub
        End If
    End If
    
    If UBound(PrePNarr) <> UBound(PostPNarr) Then
        MsgBox "Different number of Pre PN and Post PN"
        Exit Sub
    End If
    
    If UBound(PrePNarr) = 0 Then
        MsgBox "Nothing to split"
        Exit Sub
    End If
    
    n = 0
    For i = 0 To UBound(PostPNarr)
        If n = 0 Then
            Cells(iRow, colPrePN).Value = PrePNarr(i)
            Cells(iRow, colPreATA).Value = PreATAarr(i)
            Cells(iRow, colPostPN).Value = PostPNarr(i)
            Cells(iRow, colPostATA).Value = PostATAarr(i)
        Else
            Rows(iRow + n).Insert shift:=xlDown
            
            Cells(iRow + n, colPrePN).Value = PrePNarr(i)
            Cells(iRow + n, colPreATA).Value = PreATAarr(i)
            Cells(iRow + n, colPostPN).Value = PostPNarr(i)
            Cells(iRow + n, colPostATA).Value = PostATAarr(i)
            
            Cells(iRow + n, colSBNo).Value = Cells(iRow, colSBNo).Value
            Cells(iRow + n, colName).Value = Cells(iRow, colName).Value
            Cells(iRow + n, colSIN).Value = Cells(iRow, colSIN).Value
            Cells(iRow + n, colPreQTY).Value = Cells(iRow, colPreQTY).Value
            Cells(iRow + n, colPostQTY).Value = Cells(iRow, colPostQTY).Value
            Cells(iRow + n, colOpCode).Value = Cells(iRow, colOpCode).Value
            Cells(iRow + n, colChangeCode).Value = Cells(iRow, colChangeCode).Value
        End If
        n = n + 1
    Next i
    
    n = n - 1
    'Range(Cells(iRow, colName), Cells(iRow + n, colName)).Font.Color = vbRed
    'Range(Cells(iRow, colSIN), Cells(iRow + n, colSIN)).Font.Color = vbRed
    Range(Cells(iRow, colOpCode), Cells(iRow + n, colOpCode)).Font.Color = vbRed
    Range(Cells(iRow, colChangeCode), Cells(iRow + n, colChangeCode)).Font.Color = vbRed
    Cells(iRow + n, Selection.Column).Select
    
End Sub

Sub SplitRowAnyToAny()

    If Selection.Rows.Count > 1 Then
        MsgBox "Select only one row"
        Exit Sub
    End If

    Dim iRow As Integer
    Dim i As Byte
    Dim j As Byte
    Dim n As Byte
    
    Dim PrePNarr() As Variant
    Dim PreATAarr() As Variant
    Dim PostPNarr() As Variant
    Dim PostATAarr() As Variant
    
    iRow = Selection.Row
    
    PrePNarr = Split(Cells(iRow, colPrePN))
    PreATAarr = Split(Cells(iRow, colPreATA))
    PostPNarr = Split(Cells(iRow, colPostPN))
    PostATAarr = Split(Cells(iRow, colPostATA))
    
    'jesli jest 1 PrePN i kilka PreATA - skopiuj PrePN do ilosci PreATA
    'jesli jest 1 PreATA i kilka PrePN - skopiuj PreATA do ilosci PrePN
    If UBound(PrePNarr) <> UBound(PreATAarr) Then
        If UBound(PrePNarr) = 0 Then
            Do While UBound(PreATAarr) <> UBound(PrePNarr)
                ReDim Preserve PrePNarr(UBound(PrePNarr) + 1)
                PrePNarr(UBound(PrePNarr)) = PrePNarr(UBound(PrePNarr) - 1)
            Loop
        ElseIf UBound(PreATAarr) = 0 Then
            Do While UBound(PreATAarr) <> UBound(PrePNarr)
                ReDim Preserve PreATAarr(UBound(PreATAarr) + 1)
                PreATAarr(UBound(PreATAarr)) = PreATAarr(UBound(PreATAarr) - 1)
            Loop
        Else
            MsgBox "Different number of PrePN and PreATA"
            Exit Sub
        End If
    End If
    
    'jesli jest 1 PostPN i kilka PostATA - skopiuj PostPN do ilosci PostATA
    'jesli jest 1 PostATA i kilka PostPN - skopiuj PostATA do ilosci PostPN
    If UBound(PostPNarr) <> UBound(PostATAarr) Then
        If UBound(PostPNarr) = 0 Then
            Do While UBound(PostPNarr) <> UBound(PostATAarr)
                ReDim Preserve PostPNarr(UBound(PostPNarr) + 1)
                PostPNarr(UBound(PostPNarr)) = PostPNarr(UBound(PostPNarr) - 1)
            Loop
        ElseIf UBound(PostATAarr) = 0 Then
            Do While UBound(PostATAarr) <> UBound(PostPNarr)
                ReDim Preserve PostATAarr(UBound(PostATAarr) + 1)
                PostATAarr(UBound(PostATAarr)) = PostATAarr(UBound(PostATAarr) - 1)
            Loop
        Else
            MsgBox "Different number of PostPN and PostATA"
            Exit Sub
        End If
    End If
    
    If UBound(PrePNarr) = 0 And UBound(PostPNarr) = 0 Then
        MsgBox "Nothing to split"
        Exit Sub
    End If
    
    If UBound(PrePNarr) * UBound(PostPNarr) >= 20 Then
    
        If MsgBox("It will generate " & UBound(PrePNarr) * UBound(PostPNarr) & "new lines. Continue?", vbYesNo) = vbYes Then
        
        Else
            Exit Sub
        End If
    
    End If
    
    n = 0
    For i = 0 To UBound(PrePNarr)
        For j = 0 To UBound(PostPNarr)
            If n = 0 Then
                Cells(iRow, colPrePN).Value = PrePNarr(i)
                Cells(iRow, colPreATA).Value = PreATAarr(i)
                Cells(iRow, colPostPN).Value = PostPNarr(j)
                Cells(iRow, colPostATA).Value = PostATAarr(j)
            Else
                Rows(iRow + n).Insert shift:=xlDown
                
                Cells(iRow + n, colPrePN).Value = PrePNarr(i)
                Cells(iRow + n, colPreATA).Value = PreATAarr(i)
                Cells(iRow + n, colPostPN).Value = PostPNarr(j)
                Cells(iRow + n, colPostATA).Value = PostATAarr(j)
                
                Cells(iRow + n, colSBNo).Value = Cells(iRow, colSBNo).Value
                Cells(iRow + n, colName).Value = Cells(iRow, colName).Value
                Cells(iRow + n, colSIN).Value = Cells(iRow, colSIN).Value
                Cells(iRow + n, colPreQTY).Value = Cells(iRow, colPreQTY).Value
                Cells(iRow + n, colPostQTY).Value = Cells(iRow, colPostQTY).Value
                Cells(iRow + n, colOpCode).Value = Cells(iRow, colOpCode).Value
                Cells(iRow + n, colChangeCode).Value = Cells(iRow, colChangeCode).Value
            End If
            n = n + 1
        Next j
    Next i
    
    n = n - 1
    'Range(Cells(iRow, colName), Cells(iRow + n, colName)).Font.Color = vbRed
    'Range(Cells(iRow, colSIN), Cells(iRow + n, colSIN)).Font.Color = vbRed
    Range(Cells(iRow, colOpCode), Cells(iRow + n, colOpCode)).Font.Color = vbRed
    Range(Cells(iRow, colChangeCode), Cells(iRow + n, colChangeCode)).Font.Color = vbRed
    Cells(iRow + n, Selection.Column).Select
    
End Sub

Private Function Split(ByVal sValue As String) As Variant()

    Dim arr() As Variant
    ReDim arr(0) As Variant
    Dim i As Long
    
    i = 0
    Do While InStr(sValue, vbLf) > 0
        If i = 0 Then
            arr(0) = Left(sValue, InStr(sValue, vbLf) - 1)
            sValue = Right(sValue, Len(sValue) - Len(arr(0)) - 1)
            i = 1
        Else
            ReDim Preserve arr(UBound(arr) + 1) As Variant
            arr(UBound(arr)) = Left(sValue, InStr(sValue, vbLf) - 1)
            sValue = Right(sValue, Len(sValue) - Len(arr(UBound(arr))) - 1)
        End If
    Loop
    If i = 0 Then
        ReDim arr(0) As Variant
    Else
        ReDim Preserve arr(UBound(arr) + 1) As Variant
    End If
    arr(UBound(arr)) = sValue
    
    Split = arr()
    
End Function

Sub CopyRow()

    If Selection.Rows.Count > 1 Then
        MsgBox "Select only one row"
        Exit Sub
    End If
    
    Dim j As Byte

    Rows(Selection.Row + 1).Insert shift:=xlDown
    
    For j = 1 To colLast
        Cells(Selection.Row + 1, j).Value = Cells(Selection.Row, j)
    Next j
    Cells(Selection.Row, colName).Font.Color = vbRed
    Cells(Selection.Row, colSIN).Font.Color = vbRed
    Cells(Selection.Row, colOpCode).Font.Color = vbRed
    Cells(Selection.Row, colChangeCode).Font.Color = vbRed
    Cells(Selection.Row + 1, colName).Font.Color = vbRed
    Cells(Selection.Row + 1, colSIN).Font.Color = vbRed
    Cells(Selection.Row + 1, colOpCode).Font.Color = vbRed
    Cells(Selection.Row + 1, colChangeCode).Font.Color = vbRed
    Selection.Offset(1, 0).Select

End Sub

Sub AddMPLData()

    If Not AddIns("PPE demo").Installed Then
        If MsgBox("This function requires 'PPE demo' add-in loaded" & vbLf _
                & "It can be found in this file location / PPEadd-in demo folder" & vbLf & vbLf _
                & "Do You want to load it now? (it will be unloaded on file exit)", vbYesNo) = vbYes Then
                
            On Error GoTo ErrHandler
                AddIns.Add Filename:=ThisWorkbook.Path & "\PPEadd-in demo\PPEadd-in demo.xlam"
                AddIns("PPE demo").Installed = True
            On Error GoTo 0
        End If
        
        MsgBox "Pleaser re-run this function"
        Exit Sub
    End If

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim r As Range
    Dim j As Byte
    Dim i As Single
    Dim iMax As Single
    
    Dim delimiter As String
    Dim sampleFormula As String
    sampleFormula = ThisWorkbook.Worksheets("MM data").Range("D1").Formula
    delimiter = Mid(sampleFormula, InStr(sampleFormula, "(") + 1, 1)

    With Sheet2
        If .AutoFilterMode Then .AutoFilterMode = False
        .Columns(colProgressionCheck).ClearContents
        .Cells(1, colProgressionCheck).Value = "Check"
    
        i = 1
        iMax = Range(.Cells(2, colPrePN), .Cells(1000000, colPrePN).End(xlUp)).Rows.Count
    
        For Each r In Range(.Cells(2, colPrePN), .Cells(1000000, colPrePN).End(xlUp))
            If r.Value <> "" And r.Value <> "--" Then
                If IsError(Application.Run("Fid_Pn", CStr(r.Value))) Then
                    .Cells(r.Row, colPreFID).Value = "no PPE data"
                    .Cells(r.Row, colPreSuperior).Value = ""
                    .Cells(r.Row, colPreVariant).Value = ""
                    .Cells(r.Row, colPreObjDep).Value = ""
                    .Cells(r.Row, colPrePPEQTY).Value = ""
                Else
                    .Cells(r.Row, colPreFID).Value = Application.Run("Fid_Pn", CStr(r.Value))
                    .Cells(r.Row, colPreSuperior).Value = Application.Run("Superior_Fid", CStr(Cells(r.Row, colPreFID).Value))
                    .Cells(r.Row, colPreVariant).Value = Application.Run("Variant_Pn_Fid", CStr(r.Value), CStr(Cells(r.Row, colPreFID).Value))
                    .Cells(r.Row, colPreObjDep).Value = Application.Run("ObjDep_Fid_Var", CStr(Cells(r.Row, colPreFID).Value), CStr(Cells(r.Row, colPreVariant).Value))
                    .Cells(r.Row, colPrePPEQTY).Value = Application.Run("Qty_Fid_Var", CStr(Cells(r.Row, colPreFID).Value), CStr(Cells(r.Row, colPreVariant).Value))
                    
                    If Left(.Cells(r.Row, colPreFID).Value, 4) = "#M/R" Then
                        .Cells(r.Row, colPreFID).Formula = _
                            "=fid_pn(" & .Cells(r.Row, colPrePN).Address & delimiter & .Cells(r.Row, colPreFIDNo).Address & ")"
                        .Cells(r.Row, colPreSuperior).Formula = _
                            "=superior_fid(" & .Cells(r.Row, colPreFID).Address & delimiter & .Cells(r.Row, colPreSuperiorNo).Address & ")"
                        .Cells(r.Row, colPreVariant).Formula = _
                            "=variant_pn_fid(" & .Cells(r.Row, colPrePN).Address & delimiter & .Cells(r.Row, colPreFID).Address & delimiter & Cells(r.Row, colPreVariantNo).Address & ")"
                        .Cells(r.Row, colPreObjDep).Formula = _
                            "=objdep_fid_var(" & .Cells(r.Row, colPreFID).Address & delimiter & .Cells(r.Row, colPreVariant).Address & ")"
                        .Cells(r.Row, colPrePPEQTY).Formula = _
                            "=qty_fid_var(" & .Cells(r.Row, colPreFID).Address & delimiter & .Cells(r.Row, colPreVariant).Address & ")"
                    ElseIf Left(.Cells(r.Row, colPreSuperior).Value, 4) = "#M/R" Then
                        .Cells(r.Row, colPreSuperior).Formula = _
                            "=superior_fid(" & .Cells(r.Row, colPreFID).Address & delimiter & .Cells(r.Row, colPreSuperiorNo).Address & ")"
                    ElseIf Left(Cells(r.Row, colPreVariant).Value, 4) = "#M/R" Then
                        .Cells(r.Row, colPreVariant).Formula = _
                            "=variant_pn_fid(" & .Cells(r.Row, colPrePN).Address & delimiter & .Cells(r.Row, colPreFID).Address & delimiter & Cells(r.Row, colPreVariantNo).Address & ")"
                        .Cells(r.Row, colPreObjDep).Formula = _
                            "=objdep_fid_var(" & .Cells(r.Row, colPreFID).Address & delimiter & .Cells(r.Row, colPreVariant).Address & ")"
                        .Cells(r.Row, colPrePPEQTY).Formula = _
                            "=qty_fid_var(" & .Cells(r.Row, colPreFID).Address & delimiter & .Cells(r.Row, colPreVariant).Address & ")"
                    End If
                    
                End If
            End If
            
            'Application.Wait DateAdd("s", 0.5, Now)
            Application.StatusBar = "Processing... " & Format(i / (2 * iMax), "0%")
            If i Mod 200 = 0 Then DoEvents 'to increase performance of procedure call DoEvents intermittently
            i = i + 1
            
        Next r
        
        For Each r In Range(.Cells(2, colPostPN), .Cells(1000000, colPostPN).End(xlUp))
            If r.Value <> "" And r.Value <> "--" Then
                If IsError(Application.Run("Fid_Pn", CStr(r.Value))) Then
                    .Cells(r.Row, colPostFID).Value = "no PPE data"
                    .Cells(r.Row, colPostSuperior).Value = ""
                    .Cells(r.Row, colPostVariant).Value = ""
                    .Cells(r.Row, colPostObjDep).Value = ""
                    .Cells(r.Row, colPostPPEQTY).Value = ""
                Else
                    .Cells(r.Row, colPostFID).Value = Application.Run("Fid_Pn", CStr(r.Value))
                    .Cells(r.Row, colPostSuperior).Value = Application.Run("Superior_Fid", CStr(Cells(r.Row, colPostFID).Value))
                    .Cells(r.Row, colPostVariant).Value = Application.Run("Variant_Pn_Fid", CStr(r.Value), CStr(Cells(r.Row, colPostFID).Value))
                    .Cells(r.Row, colPostObjDep).Value = Application.Run("ObjDep_Fid_Var", CStr(Cells(r.Row, colPostFID).Value), CStr(Cells(r.Row, colPostVariant).Value))
                    .Cells(r.Row, colPostPPEQTY).Value = Application.Run("Qty_Fid_Var", CStr(Cells(r.Row, colPostFID).Value), CStr(Cells(r.Row, colPostVariant).Value))
                
                    If Left(.Cells(r.Row, colPostFID).Value, 4) = "#M/R" Then
                        .Cells(r.Row, colPostFID).Formula = _
                            "=fid_pn(" & .Cells(r.Row, colPostPN).Address & "," & .Cells(r.Row, colPostFIDNo).Address & ")"
                        .Cells(r.Row, colPostSuperior).Formula = _
                            "=superior_fid(" & .Cells(r.Row, colPostFID).Address & "," & .Cells(r.Row, colPostSuperiorNo).Address & ")"
                        .Cells(r.Row, colPostVariant).Formula = _
                            "=variant_pn_fid(" & .Cells(r.Row, colPostPN).Address & "," & .Cells(r.Row, colPostFID).Address & "," & .Cells(r.Row, colPostVariantNo).Address & ")"
                        .Cells(r.Row, colPostObjDep).Formula = _
                            "=objdep_fid_var(" & .Cells(r.Row, colPostFID).Address & "," & .Cells(r.Row, colPostVariant).Address & ")"
                        .Cells(r.Row, colPostPPEQTY).Formula = _
                            "=qty_fid_var(" & .Cells(r.Row, colPostFID).Address & "," & .Cells(r.Row, colPostVariant).Address & ")"
                    ElseIf Left(.Cells(r.Row, colPostSuperior).Value, 4) = "#M/R" Then
                        .Cells(r.Row, colPostSuperior).Formula = _
                            "=superior_fid(" & .Cells(r.Row, colPostFID).Address & "," & .Cells(r.Row, colPostSuperiorNo).Address & ")"
                    ElseIf Left(.Cells(r.Row, colPostVariant).Value, 4) = "#M/R" Then
                        .Cells(r.Row, colPostVariant).Formula = _
                            "=variant_pn_fid(" & .Cells(r.Row, colPostPN).Address & "," & .Cells(r.Row, colPostFID).Address & "," & .Cells(r.Row, colPostVariantNo).Address & ")"
                        .Cells(r.Row, colPostObjDep).Formula = _
                            "=objdep_fid_var(" & .Cells(r.Row, colPostFID).Address & "," & .Cells(r.Row, colPostVariant).Address & ")"
                        .Cells(r.Row, colPostPPEQTY).Formula = _
                            "=qty_fid_var(" & .Cells(r.Row, colPostFID).Address & "," & .Cells(r.Row, colPostVariant).Address & ")"
                    End If
                
                End If
            End If
            
            'Application.Wait DateAdd("s", 0.5, Now)
            Application.StatusBar = "Processing... " & Format(i / (2 * iMax), "0%")
            If i Mod 200 = 0 Then DoEvents 'to increase performance of procedure call DoEvents intermittently
            i = i + 1
            
        Next r
        
        For j = 1 To colLast
            .Columns(j).EntireColumn.ColumnWidth = 100
            .Columns(j).EntireColumn.AutoFit
            On Error Resume Next
            .Columns(j).EntireColumn.ColumnWidth = Columns(j).EntireColumn.ColumnWidth * 1.1
            On Error GoTo 0
        Next j
        
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
        
        .Columns(colSBNo).Hidden = True
        
    End With
    
    Set r = Nothing
    
    Application.Calculation = xlCalculationAutomatic
    
    Call FindCorrectResult(colPreFID)
    Call FindCorrectResult(colPreVariant)
    Call FindCorrectResult(colPostFID)
    Call FindCorrectResult(colPostVariant)
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True

Exit Sub
ErrHandler:
    MsgBox "Unable to load 'PPE demo' add-in :("
    On Error GoTo 0
End Sub

Sub FindCorrectResult(col As Integer)
    
    Dim rng As Range
    Dim i As Single
    Dim n As Single
    
    With Sheet2
        For Each rng In Range(.Cells(2, col), .Cells(65000, col).End(xlUp))
            
            If Not IsError(rng.Value) Then
                If rng.HasFormula And Left(rng, 3) = "#M/" Then
                    If rng.DisplayFormat.Font.Color <> 10921638 Then
                        n = CSng(Right(rng.Value, Len(rng.Value) - 5))
                        'If n > 10 Then n = 10
                        
                        For i = 1 To n
                            
                            rng.Offset(0, -1) = i
                            If rng.DisplayFormat.Font.Color = 10921638 Then
                                rng.Value = rng.Value
                                Exit For
                            End If
                            If i = n Then rng.Offset(0, -1) = ""
                            
                        Next i
                    End If
                End If
            End If
        Next rng
    End With

End Sub
Sub CheckProgressions()

    Dim i As Integer
    Dim wsMM As Worksheet
    Dim bPrePNexpendable As Boolean
    Dim bPostPNexpendable As Boolean
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Set wsMM = ThisWorkbook.Worksheets("MM data")
    
    With Sheet2
        .Columns(colProgressionCheck).ClearContents
        .Cells(1, colProgressionCheck).Value = "Check"
        
        For i = 2 To Sheet2.Range("A1").CurrentRegion.Rows.Count
            bPrePNexpendable = False
            bPostPNexpendable = False
            
            If .Cells(i, colPrePN).Value <> "" Then
            
                If Application.CountIf(Range(wsMM.Range("C1"), wsMM.Range("C1").End(xlDown)), .Cells(i, colPrePN).Value) > 0 Then bPrePNexpendable = True
                If Application.CountIf(Range(wsMM.Range("C1"), wsMM.Range("C1").End(xlDown)), .Cells(i, colPostPN).Value) > 0 Then bPostPNexpendable = True
            
                If (.Cells(i, colPrePN).Value <> "--" And Application.CountIf(Range(wsMM.Range("A1"), wsMM.Range("A1").End(xlDown)), .Cells(i, colPrePN).Value) = 0) Or _
                    (.Cells(i, colPostPN).Value <> "--" And Application.CountIf(Range(wsMM.Range("A1"), wsMM.Range("A1").End(xlDown)), .Cells(i, colPostPN).Value) = 0) Then
                    Sheet2.Cells(i, colProgressionCheck).Value = "no MM data"
                ElseIf .Cells(i, colPrePN).Value <> "--" And .Cells(i, colPostPN).Value <> "--" And bPrePNexpendable And bPostPNexpendable Then
                    Sheet2.Cells(i, colProgressionCheck).Value = "expendable"
                ElseIf .Cells(i, colPrePN).Value <> "--" And .Cells(i, colPostPN).Value <> "--" And (bPrePNexpendable Or bPostPNexpendable) Then
                    Sheet2.Cells(i, colProgressionCheck).Value = "exp/rot"
                ElseIf (.Cells(i, colPrePN).Value <> "--" Or .Cells(i, colPostPN).Value <> "--") And (bPrePNexpendable Or bPostPNexpendable) Then
                    Sheet2.Cells(i, colProgressionCheck).Value = "expendable"
                ElseIf ProgressionCorrect(i) Then
                    Sheet2.Cells(i, colProgressionCheck).Value = "ok"
                Else
                    Sheet2.Cells(i, colProgressionCheck).Value = "to check"
                End If
            End If
        
        Next i
        
        Call CheckIfModule
        
        .Columns(colProgressionCheck).AutoFit
        .Columns(colProgressionCheck).ColumnWidth = .Columns(colProgressionCheck).ColumnWidth * 1.1
    End With

    Set wsMM = Nothing
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub

Sub CheckIfModule()

    Dim wsModules As Worksheet
    Dim rModules As Range
    Dim r As Range
    
    Set wsModules = ThisWorkbook.Worksheets("Modules")
    Set rModules = Range(wsModules.Cells(1, 1), wsModules.Cells(100, 1).End(xlUp))
    
    With Sheet2
        
        For Each r In Range(.Cells(1, colProgressionCheck), .Cells(65000, colProgressionCheck).End(xlUp))
        
            If Application.CountIf(rModules, .Cells(r.Row, colPreFID).Value) > 0 Or Application.CountIf(rModules, .Cells(r.Row, colPostFID).Value) > 0 Then
                'r.Value = r.Value & " /module"
                If r.Value <> "no MM data" Then
                    r.Value = "module"
                End If
            End If
        
        Next r
    
    End With
    
End Sub

Private Function ProgressionCorrect(ByVal iRow As Integer) As Byte

    ProgressionCorrect = True
    
    With Sheet2
        
'-------SAP MPL data exists
        If Not IsError(.Cells(iRow, colPreFID)) And _
            Not IsError(.Cells(iRow, colPreSuperior)) And _
            Not IsError(.Cells(iRow, colPreVariant)) And _
            Not IsError(.Cells(iRow, colPrePPEQTY)) And _
            Not IsError(.Cells(iRow, colPostFID)) And _
            Not IsError(.Cells(iRow, colPostSuperior)) And _
            Not IsError(.Cells(iRow, colPostVariant)) And _
            Not IsError(.Cells(iRow, colPostPPEQTY)) Then
            'ok
        Else
            ProgressionCorrect = False
            Exit Function
        End If
        
'-------Op Code in correct form
        If .Cells(iRow, colOpCode) = "RM" Or _
            .Cells(iRow, colOpCode) = "RE" Or _
            .Cells(iRow, colOpCode) = "RW" Or _
            .Cells(iRow, colOpCode) = "RI" Or _
            .Cells(iRow, colOpCode) = "QTC" Or _
            .Cells(iRow, colOpCode) = "AD" Or _
            .Cells(iRow, colOpCode) = "DE" Or _
            .Cells(iRow, colOpCode) = "RE/RW" Or _
            .Cells(iRow, colOpCode) = "RW/RE" Then
            'ok
        Else
            ProgressionCorrect = False
            Exit Function
        End If

'-------single PrePN and single PostPN (no returns)
        If InStr(.Cells(iRow, colPrePN), vbLf) = 0 And InStr(.Cells(iRow, colPostPN), vbLf) = 0 Then
            'ok
        Else
            ProgressionCorrect = False
            Exit Function
        End If
        
'-------Pre PN SAP MPL data cells populated
        
        If .Cells(iRow, colPrePN) = "--" Or _
            (.Cells(iRow, colPrePN) <> "--" And _
            .Cells(iRow, colPreFID) <> "" And _
            .Cells(iRow, colPreSuperior) <> "" And _
            .Cells(iRow, colPreVariant) <> "" And _
            .Cells(iRow, colPrePPEQTY) <> "") Then
            'ok
        Else
            ProgressionCorrect = False
            Exit Function
        End If
                    
'-------Post PN SAP MPL data cells populated
        If .Cells(iRow, colPostPN) = "--" Or _
            (.Cells(iRow, colPostPN) <> "--" And _
            .Cells(iRow, colPostFID) <> "" And _
            .Cells(iRow, colPostSuperior) <> "" And _
            .Cells(iRow, colPostVariant) <> "" And _
            .Cells(iRow, colPostPPEQTY) <> "") Then
            'ok
        Else
            ProgressionCorrect = False
            Exit Function
        End If
        
'-------single FID/Superior/Variant/PPEQty assigned
        If InStr(.Cells(iRow, colPreFID), "#") = 0 And _
            InStr(.Cells(iRow, colPreSuperior), "#") = 0 And _
            InStr(.Cells(iRow, colPreVariant), "#") = 0 And _
            InStr(.Cells(iRow, colPrePPEQTY), "#") = 0 And _
            InStr(.Cells(iRow, colPostFID), "#") = 0 And _
            InStr(.Cells(iRow, colPostSuperior), "#") = 0 And _
            InStr(.Cells(iRow, colPostVariant), "#") = 0 And _
            InStr(.Cells(iRow, colPostPPEQTY), "#") = 0 Then
            'ok
        Else
            ProgressionCorrect = False
            Exit Function
        End If
        
'-------PreQTY is numerical value and equal to PPE QTY
        If .Cells(iRow, colPreQTY) <> "-" Then
            If IsNumeric(.Cells(iRow, colPreQTY)) And Format(.Cells(iRow, colPreQTY), "@") = Format(.Cells(iRow, colPrePPEQTY), "@") Then
                'ok
            Else
                ProgressionCorrect = False
                Exit Function
            End If
        End If
            
'-------PostQTY is numerical value and equal to PPE QTY
        If .Cells(iRow, colPostQTY) <> "-" Then
            If IsNumeric(.Cells(iRow, colPostQTY)) And Format(.Cells(iRow, colPostQTY), "@") = Format(.Cells(iRow, colPostPPEQTY), "@") Then
                'ok
            Else
                ProgressionCorrect = False
                Exit Function
            End If
        End If
            
    End With

End Function

