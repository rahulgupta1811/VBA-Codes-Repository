Attribute VB_Name = "Cheat_Sheet"
Sub Start_Activity()

Dim DumpFilePath As String
Dim CheatSheetPath As String
Dim TemplatePath As String
Dim ParentPath As String
Dim DumpFile As Workbook
Dim CheatSheet As Workbook
Dim Template As Workbook
Dim DataElement As ArrayList
Set DataElement = New ArrayList

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False

'Setting up globla path for ParentPath variable
User = Environ("USERNAME")
ParentPath = ThisWorkbook.Path
ParentPath = Replace(ParentPath, "https://mydrive.merck.com/personal/" & User & "_merck_com/Documents", "C:\Users\" & User & "\OneDrive - Merck Sharp & Dohme, Corp")
ParentPath = Replace(ParentPath, "/", "\")

'Declaring File Path in variables
'DumpFilePath = ParentPath & "\AH DE JEs.xlsx"
DumpFilePath = UserForm1.TextBox1.Value
CheatSheetPath = ParentPath & "\Template\AH Germany_Austria_NL_CH_ATR_CI_Cheat Sheet.xlsx"
TemplatePath = ParentPath & "\Template\ATLAS JE Validation Template.xlsm"

'Copying Template File to Reports Folder
FolderPath = ParentPath & "\Reports\"
If Dir(FolderPath, vbDirectory) = "" Then
    MkDir FolderPath
End If
'Dim FSO As Object
'Set FSO = CreateObject("Scripting.FileSystemObject")
'FSO.CopyFile TemplatePath, FolderPath & "ATLAS JE Validation Template.xlsm"

'Dump File Activity
Set DumpFile = Workbooks.Open(DumpFilePath)
Sheets(1).Activate
ALastCell = Range("A" & Range("A:A").Rows.Count).End(xlUp).Row

'Adding Company Codes to a list
Dim Comp_Code As ArrayList
Set Comp_Code = New ArrayList
Dim Rng As Range
Set Rng = Range("A2:A" & ALastCell)
For Each cell In Rng
    If Not Comp_Code.Contains(cell.Value) Then
        If IsNumeric(cell.Value) Then
            Comp_Code.Add cell.Value
        End If
    End If
    
Next cell
CompanyCount = Comp_Code.Count

For comp = 0 To Comp_Code.Count - 1
    
    DumpFile.Activate
    Range("A1:U1").AutoFilter Field:=4, Criteria1:="ZR"
    Range("A1:U1").AutoFilter Field:=1, Criteria1:=Comp_Code(comp)
    FilCount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
    Range("D1").Offset(1, 0).Select
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.CopyFile TemplatePath, FolderPath & "ATLAS JE Validation Template_" & Comp_Code(comp) & ".xlsm"
    

    'Adding data elemnet from dump file to DataElement list
    For i = 2 To ALastCell
        DumpFile.Activate
        If Not ActiveCell.EntireRow.Hidden Then
            BaseCell = ActiveCell.Address
            Set DataElement = New ArrayList
            DataElement.Add Range(BaseCell).Offset(0, -3).Value 'Company_Code
            DataElement.Add Range(BaseCell).Offset(0, -2).Value 'Account
            DataElement.Add Range(BaseCell).Offset(0, 3).Value 'Value_Date
            DataElement.Add Range(BaseCell).Offset(0, 4).Value 'Posting KEy
            DataElement.Add Range(BaseCell).Offset(0, 5).Value 'Amount
            DataElement.Add Range(BaseCell).Offset(0, 6).Value 'Doc_Currency
            DataElement.Add Range(BaseCell).Offset(0, 11).Value 'Text
            DataElement.Add Range(BaseCell).Offset(0, 12).Value 'Assignment
            Text_val = Range(BaseCell).Offset(0, 17).Value
            If Range(BaseCell).Row > ALastCell Then
                Exit For
            End If
            ActiveCell.Offset(1, 0).Select
            'Calling GetGL function to get GL code from Cheat Sheet
            CheatGL = GetGL(Text_val, CheatSheetPath, DataElement(0))
            'Calling Generate Template to generate final JE file
            FinalJE = Generate_Template(FolderPath & "ATLAS JE Validation Template_" & Comp_Code(comp) & ".xlsm", DataElement, CheatGL)
            Set DataElement = Nothing
            
        Else
            ActiveCell.Offset(1, 0).Select
        End If
        
    Next i
    Workbooks(FinalJE).Save
    Workbooks(FinalJE).Close
Next comp

'Saving and closing
DumpFile.Close
UserForm1.Hide
MsgBox "Completed", vbInformation, "Success"

End Sub
Private Function Generate_Template(Final_FilePath As String, DataElement As ArrayList, CheatGL) As String

Dim FinalFile As Workbook
Set FinalFile = Workbooks.Open(Final_FilePath)
Sheets("Journal Entry").Activate

'Entering Data into Headers Field
Range("A3").Value = DataElement(0)
Range("A5").Value = "SA"
Range("D5").Value = "AH DE JEs"
Range("D9").Value = "BSC EMEA"
Range("G7").Value = UserForm1.TextBox2.Value
Range("F9").Value = DataElement(5)

CurrMon = Format(DateAdd("M", 0, Date), "M")
CurrDay = Format(DateAdd("M", 0, Date), "dd")
CurrDate = Format(DateAdd("M", 0, Date), "MM/dd/yyyy")

If CurrDay = 1 And CurDay <= 3 Then
    CurrDate = Format(DateAdd("M", 0, DateSerial(Year(Date), Month(Date), 0)), "MM/dd/yyyy")
    CurrMon = Format(DateAdd("M", -1, Date), "M")
End If

Range("A7").Value = CurrDate
Range("A9").Value = CurrDate
Range("D7").Value = CurrMon

'Entering Field Values data
Dim FieldCol As ArrayList
Set FieldCol = New ArrayList

'Adding cols to field cols list
FieldCol.Add "U" 'Company Code
FieldCol.Add "O" 'Acount
FieldCol.Add "AB" 'Value Date
FieldCol.Add "P" 'Posting Key
FieldCol.Add "Q" 'For Skipping
FieldCol.Add "Q" 'Amount
FieldCol.Add "V" 'Text
FieldCol.Add "AG" 'Assignment

'Adding Data to rows
LCell = Range("O" & Range("O:O").Rows.Count).End(xlUp).Row
    For i = 0 To DataElement.Count - 1
        If i = 5 Then
            GoTo ex
        End If
        Range(FieldCol(i) & LCell + 1 & ":" & Range(FieldCol(i) & LCell + 1).Offset(1, 0).Address).Value = DataElement(i)
ex:
    Next i

'Overwritting posting key and account number
pstkey = Range("P" & LCell + 1).Value
If pstkey = 40 Or pstkey = "40" Then
    Range("P" & LCell + 1).Value = 50
End If
If pstkey = 50 Or pstkey = "50" Then
    Range("P" & LCell + 1).Value = 40
End If

'Adding GL code extracted from cheat sheet to account column
Range("O" & LCell + 2).Value = CheatGL

'Removing Negative Sign from amounts
For n = 12 To Range("Q" & Range("Q:Q").Rows.Count).End(xlUp).Row
    If InStr(Range("Q" & n).Value, "-") Then
        Range("Q" & n).Value = Replace(Range("Q" & n).Value, "-", "")
    End If
Next n

Generate_Template = FinalFile.Name
End Function
Private Function GetGL(Text, CheatSheetPath As String, Company_Code)

'Opening Cheatsheet file to get the GL code
Dim cheatWb As Workbook
Set cheatWb = Workbooks.Open(CheatSheetPath)
Sheets(Company_Code).Activate
ActiveSheet.AutoFilterMode = False

'Looping through all cell of B column to match the text and getting the GL col from next column
For i = 6 To Range("B" & Range("B:B").Rows.Count).End(xlUp).Row
    If InStr(Range("B" & i).Value, Text) > 0 Then
        GetGL = Range("B" & i).Offset(0, 1).Value
        cheatWb.Close
        Exit For
    End If
Next i

End Function

