Attribute VB_Name = "m7CreateTemplate"
Sub CreateTemplate()

    If MsgBox("This template can be uploaded only by SAP Consultant. Continue?", vbYesNo) <> vbYes Then Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim wsSBConfChart As Worksheet
    Dim wsMPL As Worksheet
    Dim wsTemplate As Worksheet
    Dim wsNewTemplate As Worksheet
    
    Set wsSBConfChart = ThisWorkbook.Worksheets("SB Conf. Chart")
    Set wsMPL = ThisWorkbook.Worksheets("SB mods to upload")
    Set wsTemplate = ThisWorkbook.Worksheets("template")
    
    If wsMPL.Cells(2, colMPLDocNo) = "" Then Exit Sub
    
    Set wsNewTemplate = Application.Workbooks.Add.Sheets(1)
    
    wsNewTemplate.Name = wsSBConfChart.Cells(1, 1) & " R" & wsSBConfChart.Cells(1, 2)
    wsNewTemplate.Range(Cells(1, 1), Cells(1, colTempLast)).EntireColumn.HorizontalAlignment = xlCenter
    
    Dim i As Long
    Dim j As Long
    Dim iLast As Long

    iLast = Application.WorksheetFunction.Min(wsMPL.Cells(1, colMPLCounter).End(xlDown).Row, 50000)
    
    With wsNewTemplate
        
        For j = 1 To colTempLast
            .Cells(1, j) = wsTemplate.Cells(1, j)
        Next j
        
        For i = 2 To iLast
        
            .Cells(i, colTempPreFID) = wsMPL.Cells(i, colMPLPreFID)
            .Cells(i, colTempPreVar) = wsMPL.Cells(i, colMPLPreVar)
            .Cells(i, colTempCounter) = wsMPL.Cells(i, colMPLCounter)
            .Cells(i, colTempPrePN) = PNlong(CStr(wsMPL.Cells(i, colMPLPrePn).Value))
            .Cells(i, colTempPreQty) = wsMPL.Cells(i, colMPLPreQty)
            .Cells(i, colTempPreUnit) = "EA"
            .Cells(i, colTempPostPN) = PNlong(CStr(wsMPL.Cells(i, colMPLPostPn).Value))
            .Cells(i, colTempIcCode) = ""
            .Cells(i, colTempIcDescr) = ""
            .Cells(i, colTempPostQty) = wsMPL.Cells(i, colMPLPostQty)
            .Cells(i, colTempPostUnit) = "EA"
            .Cells(i, colTempPostFID) = wsMPL.Cells(i, colMPLPostFID)
            .Cells(i, colTempPostVar) = wsMPL.Cells(i, colMPLPostVar)
            .Cells(i, colTempStatus) = wsMPL.Cells(i, colMPLOpCode)
            .Cells(i, colTempAction) = wsMPL.Cells(i, colMPLActionType)
            
            If wsMPL.Cells(i, colMPLDocType) <> "" Then
                .Cells(i, colTempDocNo) = wsMPL.Cells(i, colMPLDocNo)
                .Cells(i, colTempDocType) = wsMPL.Cells(i, colMPLDocType)
                .Cells(i, colTempDocPart) = wsMPL.Cells(i, colMPLDocPart)
                .Cells(i, colTempDocVer) = wsMPL.Cells(i, colMPLDocVer)
            Else
                .Cells(i, colTempDocNo) = wsMPL.Cells(i, colMPLDocNo).End(xlUp)
                .Cells(i, colTempDocType) = wsMPL.Cells(i, colMPLDocType).End(xlUp)
                .Cells(i, colTempDocPart) = wsMPL.Cells(i, colMPLDocPart).End(xlUp)
                .Cells(i, colTempDocVer) = wsMPL.Cells(i, colMPLDocVer).End(xlUp)
            End If

        Next i
    End With
    
    wsNewTemplate.Range(Cells(1, 1), Cells(1, colTempLast)).EntireColumn.AutoFit
    wsNewTemplate.Activate
    
    Set wsSBConfChart = Nothing
    Set wsMPL = Nothing
    Set wsTemplate = Nothing
    Set wsNewTemplate = Nothing
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub

Sub CreateSSBTemplate()

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim wsMPL As Worksheet
    Dim wsSSBTemplate As Worksheet
    Dim wsNewTemplate As Worksheet
    
    Set wsMPL = ThisWorkbook.Worksheets("SB mods to upload")
    Set wsSSBTemplate = ThisWorkbook.Worksheets("SSB upl template")
    
    If wsMPL.Cells(2, colMPLDocNo) = "" Then Exit Sub
    
    Set wsNewTemplate = Application.Workbooks.Add.Sheets(1)
    
    wsNewTemplate.Range(Cells(1, 1), Cells(1, colSSBTempLast)).EntireColumn.HorizontalAlignment = xlCenter
    
    Dim i As Long
    Dim j As Long
    Dim iLast As Long

    iLast = Application.WorksheetFunction.Min(wsMPL.Cells(1, colMPLCounter).End(xlDown).Row, 50000)
    
    With wsNewTemplate
        
        For j = 1 To colSSBTempLast
            .Cells(1, j) = wsSSBTemplate.Cells(1, j)
        Next j
        
        j = 1
        For i = 2 To iLast
            If wsMPL.Cells(i, colMPLDocType) = "SSB" Or (wsMPL.Cells(i, colMPLDocType) = "" And wsMPL.Cells(i, colMPLDocType).End(xlUp) = "SSB") Then
                j = j + 1
                .Cells(j, colSSBTempPreFID) = wsMPL.Cells(i, colMPLPreFID)
                
                If wsMPL.Cells(i, colMPLActionType) <> "Node Deleted" Then
                    .Cells(j, colSSBTempPreVar) = wsMPL.Cells(i, colMPLPreVar)
                End If
                
                .Cells(j, colSSBTempPostFID) = wsMPL.Cells(i, colMPLPostFID)
                .Cells(j, colSSBTempPostVar) = wsMPL.Cells(i, colMPLPostVar)
                .Cells(j, colSSBTempStatus) = wsMPL.Cells(i, colMPLOpCode)
                .Cells(j, colSSBTempChangeCode) = wsMPL.Cells(i, colMPLChangeCode)
       
                If wsMPL.Cells(i, colMPLDocType) <> "" Then
                    .Cells(j, colSSBTempDocNo) = wsMPL.Cells(i, colMPLDocNo)
                    .Cells(j, colSSBTempDocType) = wsMPL.Cells(i, colMPLDocType)
                    .Cells(j, colSSBTempDocPart) = wsMPL.Cells(i, colMPLDocPart)
                    .Cells(j, colSSBTempDocVer) = wsMPL.Cells(i, colMPLDocVer)
                Else
                    .Cells(j, colSSBTempDocNo) = wsMPL.Cells(i, colMPLDocNo).End(xlUp)
                    .Cells(j, colSSBTempDocType) = wsMPL.Cells(i, colMPLDocType).End(xlUp)
                    .Cells(j, colSSBTempDocPart) = wsMPL.Cells(i, colMPLDocPart).End(xlUp)
                    .Cells(j, colSSBTempDocVer) = wsMPL.Cells(i, colMPLDocVer).End(xlUp)
                End If
            End If
        Next i
    End With
    
    wsNewTemplate.Range(Cells(1, 1), Cells(1, colSSBTempLast)).EntireColumn.AutoFit
    wsNewTemplate.Activate
    
    Set wsMPL = Nothing
    Set wsSSBTemplate = Nothing
    Set wsNewTemplate = Nothing
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub

Sub CreatePSBTemplate()

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim wsMPL As Worksheet
    Dim wsPSBTemplate As Worksheet
    Dim wsNewTemplate As Worksheet
    
    Set wsMPL = ThisWorkbook.Worksheets("SB mods to upload")
    Set wsPSBTemplate = ThisWorkbook.Worksheets("PSB upl template")
    
    If wsMPL.Cells(2, colMPLDocNo) = "" Then Exit Sub
    
    Set wsNewTemplate = Application.Workbooks.Add.Sheets(1)
    
    wsNewTemplate.Range(Cells(1, 1), Cells(1, colPSBTempLast)).EntireColumn.HorizontalAlignment = xlCenter
    
    Dim i As Long
    Dim j As Long
    Dim iLast As Long

    iLast = Application.WorksheetFunction.Min(wsMPL.Cells(1, colMPLCounter).End(xlDown).Row, 50000)
    
    With wsNewTemplate
        
        For j = 1 To colPSBTempLast
            .Cells(1, j) = wsPSBTemplate.Cells(1, j)
        Next j
        
        j = 1
        For i = 2 To iLast
            If wsMPL.Cells(i, colMPLDocType) = "PSB" Or (wsMPL.Cells(i, colMPLDocType) = "" And wsMPL.Cells(i, colMPLDocType).End(xlUp) = "PSB") Then
                
                j = j + 1
                .Cells(j, colPSBTempPrePN) = PNlong(CStr(wsMPL.Cells(i, colMPLPrePn).Value))
                .Cells(j, colPSBTempPostPN) = PNlong(CStr(wsMPL.Cells(i, colMPLPostPn).Value))
       
                If wsMPL.Cells(i, colMPLDocType) <> "" Then
                    .Cells(j, colPSBTempDocNo) = wsMPL.Cells(i, colMPLDocNo)
                    .Cells(j, colPSBTempDocType) = wsMPL.Cells(i, colMPLDocType)
                    .Cells(j, colPSBTempDocPart) = wsMPL.Cells(i, colMPLDocPart)
                    .Cells(j, colPSBTempDocVer) = wsMPL.Cells(i, colMPLDocVer)
                Else
                    .Cells(j, colPSBTempDocNo) = wsMPL.Cells(i, colMPLDocNo).End(xlUp)
                    .Cells(j, colPSBTempDocType) = wsMPL.Cells(i, colMPLDocType).End(xlUp)
                    .Cells(j, colPSBTempDocPart) = wsMPL.Cells(i, colMPLDocPart).End(xlUp)
                    .Cells(j, colPSBTempDocVer) = wsMPL.Cells(i, colMPLDocVer).End(xlUp)
                End If
            End If
        Next i
    End With
    
    wsNewTemplate.Range(Cells(1, 1), Cells(1, colPSBTempLast)).EntireColumn.AutoFit
    wsNewTemplate.Activate
    
    Set wsMPL = Nothing
    Set wsPSBTemplate = Nothing
    Set wsNewTemplate = Nothing
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub


