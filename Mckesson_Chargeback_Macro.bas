Attribute VB_Name = "WC_Report"
Public WCRaw As Workbook
Public BasePath As String
Sub WCReport()
Attribute WCReport.VB_ProcData.VB_Invoke_Func = "R\n14"

Application.DisplayAlerts = False
Call ReadFilesWithSpecificWord
Call MergeVPFile
'Application.ScreenUpdating = False
BasePath = Application.ActiveWorkbook.Path
Set WCRaw = Workbooks.Open(Application.ActiveWorkbook.Path & "\Support_Files\FinalFiles\Export\RW_Raw_Report.XLSX")

Sheets("Sheet1").Activate
For i = 1 To 4

    With ActiveSheet.Columns(i)
        .NumberFormat = "0"
        .Value = .Value
    End With
Next i

Sheets.Add(After:=Sheets("Sheet1")).Name = "VP"
Sheets.Add(After:=Sheets("VP")).Name = "IDOC"

Sheets(1).Activate
Range("Y1").Value = "Today"
Range("Z1").Value = "Resub Days"
Range("X1").Copy
Range("Y1:Z1").PasteSpecial xlPasteFormats
Range("Y2").Value = Date
Range("Z2").Value = "=Y2-J2"
Range("Y2:Z2").Copy
LastRow = Range("A1").End(xlDown).Row
Range("Y2:Z" & LastRow).PasteSpecial xlPasteAll
Range("Y2:Z" & LastRow).Copy
Range("Y2:Z" & LastRow).PasteSpecial xlPasteValues

Range("U1").EntireColumn.Delete
Range("T1").EntireColumn.Delete

ActiveSheet.Range("A1:X1").AutoFilter Field:=24, Criteria1:="<15"

Dim Rng As Range
Set Rng = Range("A1").CurrentRegion
Rng.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
ActiveSheet.AutoFilterMode = False

Call DataFromVPCFile
Call DataFromIDOCFile
Call SortingRaW
Call TemplateFile

MsgBox "Compeleted", vbInformation, "Success"
End Sub
Function DataFromVPCFile()

Dim VPFile As Workbook
'VPFilePath = "C:\Users\eo5v4x3\Desktop\WIP\CB Macro\Vendor Parameters.xlsx"

Sheets(1).Activate
Set VPFile = Workbooks.Open(Application.ActiveWorkbook.Path & "\VendorParaMeter.xlsx")

LastRow = Range("A1").End(xlDown).Row

Set Cols = New ArrayList
    Cols.Add "A"
    Cols.Add "D"
    Cols.Add "AO"
    
Dim COl2 As ArrayList
Set COl2 = New ArrayList
    COl2.Add "A"
    COl2.Add "B"
    COl2.Add "C"

For i = 0 To 2
    VPFile.Activate
    Sheets(1).Activate
    Range(Cols(i) & "1:" & Cols(i) & LastRow).Copy
    WCRaw.Activate
    Sheets("VP").Activate
    Range(COl2(i) & "1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
Next i
For i = 1 To 3
    If i = 2 Then
        i = 3
    End If
    With ActiveSheet.Columns(i)
        .NumberFormat = "0"
        .Value = .Value
    End With
    If i = 3 Then
        Exit For
    End If
Next i
VPFile.Close

End Function
Function DataFromIDOCFile()

Dim IDocFile As Workbook
'IDocFilePath = "C:\Users\eo5v4x3\Desktop\WIP\CB Macro\IDOC_RawData_06292023.xlsx"
On Error GoTo Blink
    Set IDocFile = Workbooks.Open(Application.ActiveWorkbook.Path & "\Support_Files\FinalFiles\Export\IDOC_RawData.xlsx")
Blink:
    Set IDocFile = Workbooks.Open(BasePath & "\IDOC_RawData.xlsx")

For i = 2 To 4
    IDocFile.Activate
    Sheets(i).Activate
    LastRow = Range("A1").End(xlDown).Row
    
    If i > 2 Then
        Range("A2:A" & LastRow).Copy
        WCRaw.Activate
        Sheets("IDOC").Activate
        LastRow = Range("A1").End(xlDown).Row
        LastRow = LastRow + 1
        Range("A" & LastRow).PasteSpecial xlPasteAll
    Else
        Range("A1:A" & LastRow).Copy
        WCRaw.Activate
        Sheets("IDOC").Activate
        Range("A1").PasteSpecial xlPasteAll
        
    End If
    
    
    IDocFile.Activate
    Sheets(i).Activate
    If i > 2 Then
        LastRow = Range("P1").End(xlDown).Row
        Range("P2:Q" & LastRow).Copy
        WCRaw.Activate
        Sheets("IDOC").Activate
        LastRow = Range("B1").End(xlDown).Row
        LastRow = LastRow + 1
        Range("B" & LastRow).PasteSpecial xlPasteAll
    Else
        Range("P1:Q" & LastRow).Copy
        WCRaw.Activate
        Sheets("IDOC").Activate
        Range("B1").PasteSpecial xlPasteAll
    End If
    
    
Next i
Range("B:B").Cut
Range("A:A").Insert
Application.CutCopyMode = False
IDocFile.Close

End Function

Function SortingRaW()
WCRaw.Activate
Sheets(1).Activate

ActiveSheet.Range("A1:G1").AutoFilter Field:=7, Criteria1:="BAD*"

Dim Rng As Range
Set Rng = Range("A1").CurrentRegion
Rng.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
ActiveSheet.AutoFilterMode = False

Range("W2").Value = "=VLOOKUP(D2,'VP'!A:C,3,0)"

Dim FiblFile As Workbook
Dim FiblPath As String
FiblPath = Application.ActiveWorkbook.Path & "\FIBL_AgingZeroto999.XLSX_0.xls"

Set FiblFile = Workbooks.Open(FiblPath)
FiblFile.Activate

sh = Sheets(1).Name
    If FiblPath Like "/" Then
        WBName = GetFileName(FiblPath)
    Else
        WBName = FiblPath
    End If

WCRaw.Activate
Sheets(1).Activate
VlookupFormula = "=VLOOKUP(G2,FIBL_AgingZeroto999.XLSX_0.xls!$D:$K,7,0)"
Range("V2").Value = VlookupFormula

LastRow = Range("V1").End(xlDown).Row
Range("V2:V2").Copy
Range("V3:V" & LastRow).PasteSpecial xlPasteAll
Range("V2:V" & LastRow).Copy
Range("V2:V" & LastRow).PasteSpecial xlPasteValues

LastRow = Range("X1").End(xlDown).Row

ActiveSheet.Range("A1:X1").AutoFilter Field:=22, Criteria1:="#N/A"

Dim Rng1 As Range
Set Rng1 = Range("A1").CurrentRegion
Rng.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
ActiveSheet.AutoFilterMode = False
FiblFile.Close


End Function

Function GetFileName(myPath As String) As String

    Dim FileName As String
    Dim FSO As New FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")

    'Get File Name
    FileName = FSO.GetFileName(myPath)
    GetFileName = FileName
    
End Function
Function TemplateFile()

Dim TempFile As Workbook
'Dim TempFilePath As String
'TempFilePath = "C:\Users\eo5v4x3\Desktop\WIP\CB Macro\WC Report_06.13.2023.xlsx"
Set TempFile = Workbooks.Open(BasePath & "\Template_File\RW Template Report.xlsx")

'Unhide All Sheets
Dim ws As Worksheet
For Each ws In Worksheets
    ws.Visible = True
Next

'Copying IDoc To TempFile
Sheets("IDOC Errors").Activate
LastRow = Range("A1").End(xlDown).Row
Range("A2:C" & LastRow).Clear

WCRaw.Activate
Sheets("IDOC").Activate
LastRow = Range("A1").End(xlDown).Row
Range("A2:C" & LastRow).Copy

TempFile.Activate
Sheets("IDOC Errors").Activate
Range("A2").PasteSpecial xlPasteAll

'Copying VP To TempFile
TempFile.Activate
Sheets("Vendor Info").Activate
LastRow = Range("A1").End(xlDown).Row
Range("A2:C" & LastRow).Clear

WCRaw.Activate
Sheets("VP").Activate
LastRow = Range("A1").End(xlDown).Row
Range("A2:C" & LastRow).Copy

TempFile.Activate
Sheets("Vendor Info").Activate
Range("A2").PasteSpecial xlPasteAll

'Clear LastWeek and MoveNew Data into sheet
TempFile.Activate
Sheets("Last week").Activate
LastRow = Range("A1").End(xlDown).Row
Range("A2:AB" & LastRow).Clear

Sheets(6).Activate
LastRow = Range("A1").End(xlDown).Row
Range("A2:AB" & LastRow).Copy

Sheets("Last week").Activate
Range("A2").PasteSpecial xlPasteAll

'Clearing LW sheets
Sheets(6).Activate
LastRow = Range("A1").End(xlDown).Row
Range("A2:AB" & LastRow).Clear

'Pulling Data from Raw File to WC Report Sheet
Sheets(3).Activate
LastRow = Range("A1").End(xlDown).Row
Range("A3:AB" & LastRow).Clear

WCRaw.Activate
Sheets("Sheet1").Activate
LastRow = Range("A1").End(xlDown).Row
Range("A2:X" & LastRow).Copy

TempFile.Activate
Sheets(3).Activate
Range("A3").PasteSpecial xlPasteAll

With ActiveSheet.Columns(21)
        .NumberFormat = "0"
        .Value = .Value
End With

LastRow = Range("A1").End(xlDown).Row
Range("Y2:AB2").Copy
Range("Y3:AB" & LastRow).PasteSpecial xlPasteAll
Range("A2:AB2").Copy
Range("A3:AB" & LastRow).PasteSpecial xlPasteFormats

Rows(2).EntireRow.Delete

'Copying to LW file
LastRow = Range("A1").End(xlDown).Row
Range("A2:AB" & LastRow).Copy
Sheets(6).Activate
Range("A2").PasteSpecial xlPasteAll

'Copying Zero Amount Data
Sheets("Zero Balance in FI").Activate
LastRow = Range("A1").End(xlDown).Row
Range("A2:AB" & LastRow).Clear

Sheets(3).Activate
ActiveSheet.Range("A1:AB1").AutoFilter Field:=22, Criteria1:="=0.00"

Dim Rng As Range
Set Rng = Range("A1").CurrentRegion
Rng.Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy

Sheets("Zero Balance in FI").Activate
Range("A2").PasteSpecial xlPasteAll
Application.CutCopyMode = False
Sheets(3).Activate
Range("A1").Select

Sheets(3).Activate
Set Rng = Range("A1").CurrentRegion
Rng.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
ActiveSheet.AutoFilterMode = False

WCRaw.Save
WCRaw.Close

'Refreshing Pivots
Sheets(3).Activate
LastCell = Range("A1").End(xlDown).Row

Call Sorting

LastRow = Range("M1").End(xlDown).Row

For i = 2 To LastRow

    CheckVal = Range("M" & i).Value
    If CheckVal < 30 Then
        LastCell = Range("M" & i).Row
        LastCell = LastCell - 1
        Exit For
    End If
Next i

Sheets("No Deduct 30+ Summary").Activate
Set pt = ActiveSheet.PivotTables("PivotTable5")
ActiveSheet.PivotTables("PivotTable5").PivotFields("Resub Age"). _
        CurrentPage = "(All)"
ActiveSheet.PivotTables("PivotTable5").PivotFields("Deduct/No Deduct"). _
        CurrentPage = "No Deduct"

Dim ShName As String
ShName = Sheets(3).Name
newSource = "'" & ShName & "'!$A$1:$AB$" & LastCell
pt.ChangePivotCache ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=newSource)
ActiveSheet.PivotTables("PivotTable5").RefreshTable


Sheets("Summary").Activate
Set pt = ActiveSheet.PivotTables("PivotTable4")
newSource = "'" & ShName & "'!$A$1:$AB$" & LastRow
pt.ChangePivotCache ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=newSource)
ActiveSheet.PivotTables("PivotTable4").RefreshTable

'Final

Tday = Format(Date, "MM.dd.yyyy")
Sheets(3).Name = "RW Report " & Tday
Sheets(6).Name = "LW_" & Tday

For i = 4 To 8
    Sheets(i).Visible = False
Next i

Sheets("No Deduct 30+ Summary").Activate
Range("B2").Select

Sheets(3).Activate
'TempFile.Save

End Function

Function Sorting()
    Range("A1").Select
   Dim AllData As Range
   Dim SpecCOl As Range
   
   Set AllData = Range("A:AB")
   Set SpecCOl = Range("M:M")
   AllData.Sort Key1:=SpecCOl, Order1:=xlDescending, Header:=xlYes
   Range("M1").Select
   

End Function
Function MergeVPFile()

Dim wb As Workbook
Dim WB2 As Workbook

WB1Path = ActiveWorkbook.Path & "\Support_Files\FinalFiles\Export\AgingAbove999.XLSX"
WB2Path = ActiveWorkbook.Path & "\Support_Files\FinalFiles\Export\AgingZeroto999.XLSX"

Set wb = Workbooks.Open(WB1Path)
wb.Activate
LastRow = Range("A1").End(xlDown).Row
Rows(LastRow).EntireRow.Delete
LastRow = Range("A1").End(xlDown).Row
Range("A2:X" & LastRow).Copy

Set WB2 = Workbooks.Open(WB2Path)
WB2.Activate
LastRow = Range("A1").End(xlDown).Row
Range("A" & LastRow).PasteSpecial xlPasteAll

Dim WCRawWB As Workbook

WB2.Save
WB2.Close
wb.Close

Name Application.ActiveWorkbook.Path & "\Support_Files\FinalFiles\Export\AgingZeroto999.XLSX" As _
    Application.ActiveWorkbook.Path & "\Support_Files\FinalFiles\Export\RW_Raw_Report.XLSX"
    
End Function
Function ReadFilesWithSpecificWord()
    Dim FolderPath As String
    Dim FileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim SpecificWord As String
    Dim Wbs As Workbook
    
    ' Define the folder path
    
    ThisWorkbook.Activate
    FolderPath = Application.ActiveWorkbook.Path & "\Support_Files\FinalFiles\Export\"
    
    Set Wbs = Workbooks.Open(FolderPath & "FIBL_1_AgingZeroto999.XLSX")
    
    ' Define the specific word to look for in file names
    SpecificWord = "FIBL"
    
    ' Loop through files in the folder
    FileName = Dir(FolderPath & "*.xls")
    Do While FileName <> ""
        If FileName = "FIBL_1_AgingZeroto999.XLSX" Then
            'Jump
        Else
            ' Check if the specific word is present in the file name
            If InStr(1, FileName, SpecificWord, vbTextCompare) > 0 Then
                ' Open the workbook
                Set wb = Workbooks.Open(FolderPath & FileName)
                
                ' Perform operations on the workbook
                wb.Activate
                Sheets(1).Activate
                LastRow = Range("B7").End(xlDown).Row
                Range("B7:AH" & LastRow).Copy
                ' Your code to work with the worksheet goes here
                Wbs.Activate
                Sheets(1).Activate
                LastRow = Range("B7").End(xlDown).Row
                LastRow = LastRow + 1
                Range("B" & LastRow).PasteSpecial xlPasteValues
                ' Close the workbook without saving changes
                wb.Close SaveChanges:=False
            End If
        End If
            FileName = Dir
    Loop
    Wbs.Save
    Wbs.Close
End Function
Sub Pivot()
Dim pvtCache As PivotCache
Dim pvtTable As PivotTable

Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Sheet6!$A$1:$B$5")
Set ws = Worksheets.Add
'ws.Name = "Pivot"

With pvtCache.CreatePivotTable(Sheets(1).Range("A3"), "PivotTable2")
End With

ActiveWorkbook.ShowPivotTableFieldList = True

Set pvtfld = pvtTable.PageFields("Name")
pvtfld.Orientation = xlRowField


End Sub
