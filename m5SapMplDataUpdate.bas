Attribute VB_Name = "m5SapMplDataUpdate"
Const sFilesFolder As String = "..."
Dim sLastModificationDate1 As String
Dim sLastModificationDate2 As String

Sub MM_Update()

    MsgBox "Function disabled"

'    Application.ScreenUpdating = False
'    UpdateMMDataFromFiles
'    Application.ScreenUpdating = True
    
End Sub


Private Sub UpdateMMDataFromFiles()

    Dim wsMM As Worksheet
    Dim wsMMSourceData As Worksheet
    Dim wsExpSourceData As Worksheet
    
    Dim sMMDataFileName As String  'name of file with structrure
    Dim sExpDataFileName As String  'name of file with nodes details
    
    Dim r As Range
    
    sMMDataFileName = Dir(sFilesFolder & "SAP export ???????? - MM all.xlsx")
    sExpDataFileName = Dir(sFilesFolder & "SAP export ???????? - MM exp.xlsx")
    
    sLastModificationDate1 = FileDateTime(sFilesFolder & sMMDataFileName)
    sLastModificationDate2 = FileDateTime(sFilesFolder & sExpDataFileName)
    
    If sMMDataFileName = "" Or sExpDataFileName = "" Then
        MsgBox "One of data files not found"
        Exit Sub
    End If
    
    Set wsMM = ThisWorkbook.Worksheets("MM data")
    Set wsMMSourceData = Workbooks.Open(sFilesFolder & sMMDataFileName).Worksheets("Sheet1")
    Set wsExpSourceData = Workbooks.Open(sFilesFolder & sExpDataFileName).Worksheets("Sheet1")
    
    wsMM.Columns(1).Cells.Clear
    wsMM.Columns(2).Cells.Clear
    wsMM.Columns(3).Cells.Clear
    wsMM.Columns(1).EntireColumn.NumberFormat = "@"
    wsMM.Columns(2).EntireColumn.NumberFormat = "@"
    wsMM.Columns(3).EntireColumn.NumberFormat = "@"
    wsMMSourceData.Range(wsMMSourceData.Cells(1, 1), wsMMSourceData.Cells(1, 1).End(xlDown)).Copy
    wsMM.Cells(1, 2).PasteSpecial xlValues
    wsExpSourceData.Range(wsExpSourceData.Cells(1, 1), wsExpSourceData.Cells(1, 1).End(xlDown)).Copy
    wsMM.Cells(1, 3).PasteSpecial xlValues
    Application.CutCopyMode = False
    
    wsMM.Range("A1").Value = "PN"
    wsMM.Range("B1").Value = "PN:Cage code"
    wsMM.Range("C1").Value = "Expendable"
    
    For Each r In wsMM.Range(wsMM.Cells(2, 2), wsMM.Cells(1, 2).End(xlDown))
        r.Offset(0, -1) = Application.Run("PNshort", CStr(r.Value))
    Next r
    For Each r In wsMM.Range(wsMM.Cells(2, 3), wsMM.Cells(1, 3).End(xlDown))
        r.Value = Application.Run("PNshort", CStr(r.Value))
    Next r
    
    wsMM.Range("J2").Value = "Last modification of source file1: " & sLastModificationDate1
    wsMM.Range("J3").Value = "Last modification of source file2: " & sLastModificationDate2
    wsMM.Range("J2").EntireColumn.AutoFit
    
    wsMMSourceData.Parent.Close SaveChanges:=False
    wsExpSourceData.Parent.Close SaveChanges:=False
    wsMM.Range("A1").Select
    
End Sub


