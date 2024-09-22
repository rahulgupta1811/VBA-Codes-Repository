Attribute VB_Name = "Module1"
Sub CPH_WF()

Dim uRow As Integer
Dim LRow As Integer
Dim LastRow As Integer
Dim WF As Workbook

Dim WF_File As String
WF_File = UserForm1.TextBox1.Value 'ThisWorkbook.Path & "\Open WF Analyst_Full Data_data.csv"

'Call Savebackup

Application.ScreenUpdating = False
'removing Formula from WF Sheet

Sheets("Open_WF_Mgr_Full_Data_data").Activate
Sheets("Open_WF_Mgr_Full_Data_data").AutoFilterMode = False
LRow = Range("A2").End(xlDown).Row
Range("AT3:BD" & LRow).Copy
Range("AT3:BD" & LRow).PasteSpecial xlPasteValuesAndNumberFormats
Application.CutCopyMode = False
Range("A2").Activate

'clearing Last Day Dump Sheet
Worksheets("Last Day Dump").Visible = True
Sheets("Last Day Dump").Activate
Sheets("Last Day Dump").AutoFilterMode = False

LastRow = Range("BE2").End(xlDown).Row
Range("A4:BE" & LastRow).Clear
Application.CutCopyMode = False

'copying WF status in Last Day Dump
Sheets("Open_WF_Mgr_Full_Data_data").Activate
LastRow = Range("A2").End(xlDown).Row
Range("A3:BD" & LastRow).Copy
Sheets("Last Day Dump").Activate
Range("A4").PasteSpecial xlPasteValuesAndNumberFormats

'Formatting Last Day Dump Sheet
Rows("3:3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Dim LastRow2 As Integer
    LastRow2 = Range("BE2").End(xlDown).Row
    LastRow2 = LastRow2 - 1
    Rows("4:" & LastRow2).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

LRow = Range("A2").End(xlDown).Row
Range("BE3").Copy
'LRow = LRow + 1
Range("BE4:BE" & LRow).PasteSpecial xlPasteFormulas

Sheets("Last Day Dump").Activate
Sheets("Last Day Dump").AutoFilterMode = False
Rows(3).EntireRow.Delete

'Putting Current WF data
ThisWorkbook.Activate
Sheets("Open_WF_Mgr_Full_Data_data").Activate
Range("A3:AS" & LastRow).Delete Shift:=xlUp

Set WF = Workbooks.Open(WF_File)
WF.Activate
Sheets(1).Activate
LastRow = Range("A1").End(xlDown).Row
Range("A2:AS" & LastRow).Copy

ThisWorkbook.Activate
Sheets("Open_WF_Mgr_Full_Data_data").Activate
Range("A3").PasteSpecial xlPasteValuesAndNumberFormats
Application.ScreenUpdating = True
Application.DisplayAlerts = False
WF.Close
Application.DisplayAlerts = True

'Getting Data from Last DayDump
ThisWorkbook.Activate
Sheets("Last Day Dump").Activate
Range("A2:BE2").AutoFilter , Field:=57, Criteria1:="#N/A"
Range("A2").CurrentRegion.SpecialCells(xlCellTypeVisible).Copy

Sheets("Open_WF_Mgr_Full_Data_data").Activate
LastRow = LastRow + 2

Range("A" & LastRow).Activate
Range("A" & LastRow).PasteSpecial xlPasteValuesAndNumberFormats

Rows(LastRow).EntireRow.Delete

'Formula Setup
Sheets("Open_WF_Mgr_Full_Data_data").Activate
Range("A1").Activate
NewLastRow = Range("A3").End(xlDown).Row
'NewLastRow = NewLastRow - 2
Range("AT1:BD1").Copy
Range("AT3:BD" & NewLastRow).PasteSpecial xlPasteFormulasAndNumberFormats
Range("AT3:BD" & NewLastRow).PasteSpecial xlPasteFormats
Range("BD3:BD" & NewLastRow).Clear
'
''Setting up Blanks
'Sheets("Open_WF_Mgr_Full_Data_data").Activate
'LRow = Range("A:A").SpecialCells(xlCellTypeLastCell).Row
'LRow = LRow - 1
'Dim Rng As Range
'Dim CellAdd As String
'
'For Each Rng In Range("AT3:AT" & LRow)
'
'    If Rng.Value = "" Then
'        Rng.Value = "Not Available"
'        CellAdd = Rng.Address
'        CellAdd = Replace(CellAdd, "$AT$", "")
'        'MsgBox CellAdd
'        Range("AU" & CellAdd).Value = "=G" & CellAdd
'    End If
'Next Rng


'Completion
'Sheets("Last Day Dump").Activate
'Sheets("Last Day Dump").AutoFilterMode = False
'Rows(3).EntireRow.Delete

NewLRow = Range("A:A").SpecialCells(xlCellTypeLastCell).Row
'MsgBox NewLRow
NewLRow = NewLRow - 1
'Rows(NewLRow).EntireRow.Delete
Range("A3").Activate

Sheets("Open_WF_Mgr_Full_Data_data").Activate
uRow = Range("BD:BD").SpecialCells(xlCellTypeLastCell).Row
'MsgBox uRow
uRow = uRow - 1
Range("BE3:BE" & uRow).Clear

Range("AT1:BC1").Copy
Range("AT" & uRow & ":BC" & uRow).PasteSpecial xlPasteFormats

Range("A3").Activate
Range("A3").Select
Application.CutCopyMode = False
uRow = 0
LRow = 0
LastRow = 0

Worksheets("Last Day Dump").Visible = False

Sheets("Open_WF_Mgr_Full_Data_data").Activate

Call Additional

'Pasting Values only
LastCell = Range("A3").End(xlDown).Row
Range("AT3:AZ" & LastCell).Copy
'Range("AT3:AZ" & LastCell).PasteSpecial xlPasteValues
Application.CutCopyMode = False
Range("A2").Activate

'Finished
MsgBox "Completed", vbInformation, "Success"

End Sub
Sub Savebackup()
    Application.DisplayAlerts = False
    Dim Tday As String
    Tday = Format(Date, "MMDDYYYY")


    Dim fso As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    MsgBox ThisWorkbook.FullName
    Call fso.CopyFile(ThisWorkbook.FullName, ThisWorkbook.Path & "\Backup_WF Status_ ISMC.xlsm")
    Application.DisplayAlerts = True
End Sub
Sub Additional()
    TLastCell = Range("A2").End(xlDown).Row
    Range("A2:BD2").AutoFilter , Field:=46, Criteria1:=""
    With Worksheets("Open_WF_Mgr_Full_Data_data").AutoFilter.Range
        'Setting Not Available in Blanks
        Range("AT" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
        CurrentCell = ActiveCell.Address
        RCount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
        If RCount > 1 Then
            Range(CurrentCell & ":AT" & TLastCell).SpecialCells(xlCellTypeVisible).SpecialCells(xlCellTypeVisible).Value = "Not Available"
        End If

'        Setting Values From G col
        If RCount > 1 Then
            Range("AU" & .Offset(1, 0).SpecialCells(xlCellTypeVisible).Row).Select
            'ActiveCell.Offset(1, 0).SpecialCells(xlCellTypeVisible).Select
            Cfrom = ActiveCell.Row

            Addr = Range("AT" & Rows.Count).End(xlUp).Row
            
            'Range("G" & Cfrom & ":G" & Addr).SpecialCells(xlCellTypeVisible).Copy
            'Range("AU" & Cfrom).Value = "=G" & Cfrom
            'Range("AU" & Cfrom).Copy
            'Range("AU" & Cfrom & ":AU" & Addr).SpecialCells(xlCellTypeVisible).PasteSpecial xlPasteAll
            'Range("AT" & Cfrom & ":AZ" & Addr).SpecialCells(xlCellTypeVisible).Copy
            'Range("AT" & Cfrom & ":AZ" & Addr).PasteSpecial xlPasteValues
            Application.CutCopyMode = False
            
            
        End If
    
    End With
End Sub
