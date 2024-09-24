Attribute VB_Name = "GL_Activity"
Sub Start_activity()

Dim DumpFilePath As String
Dim GL_FilePath As String
Dim GLFile As Workbook
Dim DumpFile As Workbook
Dim GLNameList As ArrayList

Application.AskToUpdateLinks = False
Application.DisplayAlerts = False
Application.ScreenUpdating = False

'Changing Parent Path
Usr = Environ("USERNAME")
Usr = LCase(Usr)
ParentPath = ThisWorkbook.Path
ParentPath = Replace(ParentPath, "https://mydrive.merck.com/personal/" & Usr & "_merck_com/Documents", "C:\Users\" & Usr & "\OneDrive - Merck Sharp & Dohme LLC")
'ParentPath = Replace(ParentPath, "https://mydrive.merck.com/personal/" & Usr & "_merck_com/Documents", "C:\Users\" & Usr & "\OneDrive - Merck Sharp & Dohme, Corp")
ParentPath = Replace(ParentPath, "/", "\")
'DumpFilePath = ParentPath & "\1367_Open Item_22.08.2024.xlsx"
DumpFilePath = UserForm1.TextBox1.Value

'Decalring ArrayList and assigning Values
Set GLNameList = New ArrayList
GLNameList.Add "IVA PERCEPCION"
GLNameList.Add "Retention Imposed"
GLNameList.Add "Vat 21%"
GLNameList.Add "Bank Fee"
GLNameList.Add "AFIP"
GLNameList.Add "COELSA"
GLNameList.Add "TAX LAW"

Set DumpFile = Workbooks.Open(DumpFilePath)

'C:\Users\chousari\OneDrive - Merck Sharp & Dohme LLC\JE\New folder\Desk Top\Other Important file\Important File\APAC\2024 Framework\QPA Coding\QPA Rahul\1367_Open Item_22.08.2024.xlsx
'C:\Users\chousari\OneDrive - Merck Sharp & Dohme, Corp\JE\New folder\Desk Top\Other Important file\Important File\APAC\2024 Framework\QPA Coding\QPA Rahul\1367_Open Item_22.08.2024.xlsx

Sheets(1).Activate
ActiveSheet.AutoFilterMode = False
LastRow = Range("A1").End(xlDown).Row

FolderPath = ParentPath & "\Reports\"

For i = 0 To GLNameList.Count - 1
    If GLNameList(i) = "IVA PERCEPCION" Then
        FinalFile = FolderPath & GLNameList(i) & "\Weekly_CI_6_TAX _IVA PERCEPCION_T-code_FB41 " & Format(DateAdd("M", 0, Date), "DD_mm") & ".xlsm"
        TemplateFile = ParentPath & "\Template\" & GLNameList(i) & "\Weekly_CI_6_TAX _IVA PERCEPCION_T-code_FB41.xlsm"
        Call IVA_PERCEPCION(DumpFile, ParentPath, GLNameList(i), GLNameList(i), FinalFile, TemplateFile)
    End If
    
    If GLNameList(i) = "Retention Imposed" Then
        FinalFile = FolderPath & GLNameList(i) & "\Weekly_CI#6_Retention Imposed_T-code_FB41_" & Format(DateAdd("M", 0, Date), "DD_mm") & ".xlsm"
        TemplateFile = ParentPath & "\Template\" & GLNameList(i) & "\Retention Imposed_Template.xlsm"
        Call Retention_Imposed(DumpFile, ParentPath, GLNameList(i), GLNameList(i), FinalFile, TemplateFile)
    End If
    
    If GLNameList(i) = "Vat 21%" Then
        FinalFile = FolderPath & GLNameList(i) & "\Weekly_CI#6_TAX _VAT_21% _T-code_FB41_" & Format(DateAdd("M", 0, Date), "DD_mm") & ".xlsm"
        TemplateFile = ParentPath & "\Template\" & "\VAT\Weekly_CI_TAX _VAT_Template.xlsm"
        Call IVA_PERCEPCION(DumpFile, ParentPath, GLNameList(i), GLNameList(i), FinalFile, TemplateFile)
    End If
    
    If GLNameList(i) = "AFIP" Then
        FinalFile = FolderPath & GLNameList(i) & "\1367 CI#6_AFIP_" & Format(DateAdd("M", 0, Date), "mmmm yyyy") & ".csv"
        TemplateFile = ParentPath & "\Template\" & "\AFIP\1367_AFIP_Template.csv"
        Call AFIP(DumpFile, ParentPath, GLNameList(i), TemplateFile, FinalFile)
    End If
    If GLNameList(i) = "COELSA" Then
        FinalFile = FolderPath & GLNameList(i) & "\CI#3_COELSA_" & Format(DateAdd("M", 0, Date), "dd.mm") & ".csv"
        TemplateFile = ParentPath & "\Template\" & "\COELSA\CI#3_COELSA.csv"
        Call COELSA(DumpFile, ParentPath, GLNameList(i), TemplateFile, FinalFile)
    End If
    If GLNameList(i) = "TAX LAW" Then
        FinalFile = FolderPath & GLNameList(i) & "\Weekly_CI#6_TAX_25413_T-code_FB41_" & Format(DateAdd("M", 0, Date), "dd.mm") & ".csv"
        TemplateFile = ParentPath & "\Template\" & "\TAX BY DEBIT LAW 25413\Weekly_TAX_Template.xlsm"
        Call Tax_Law(DumpFile, ParentPath, GLNameList(i), TemplateFile, FinalFile)
    End If
    
    If GLNameList(i) = "Bank Fee" Then
        FinalFile = FolderPath & GLNameList(i) & "\Weekly_CI#3_Bank Fees_ARS_" & Format(DateAdd("M", 0, Date), "dd.mm") & ".csv"
        TemplateFile = ParentPath & "\Template\" & "\Bank Fee\Weekly_Bank Fees_Template.csv"
        Call Bank_Fee(DumpFile, ParentPath, GLNameList(i), TemplateFile, FinalFile)
    End If
Next

DumpFile.Close
UserForm1.Hide
MsgBox "Completed", vbInformation, "Success"

End Sub

Private Sub IVA_PERCEPCION(DumpFile As Workbook, ParentPath, ReportFolder As String, GLName As String, FinalFile, TemplateFile)
'ReportFolder = "IVA PERCEPCION"
'Creating Reports flders if it doesn't exists

FolderPath = ParentPath & "\Reports\" & ReportFolder
If Dir(FolderPath, vbDirectory) = "" Then
    MkDir FolderPath
End If

'Copying Template_File to Report Folder
Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.CopyFile TemplateFile, FinalFile

'C:\Users\CHATSUMO\Downloads\QPA Rahul 2\QPA Rahul\Template\IVA PERCEPCION\IVA PERCEPCION\Weekly_CI_6_TAX _IVA PERCEPCION_T-code_FB41.xlsm

'Opening GL File and clearing it
Set GLFile = Workbooks.Open(FinalFile)
Sheets("Template").Activate
GLLastRow = Range("C23").End(xlDown).Row
Range("C23:AF" & GLLastRow).Clear
Sheets("Support").Activate
ActiveSheet.AutoFilterMode = False
Range("A2:T" & Range("T1").End(xlDown).Row).Clear

'Filtering Data from Dump File and Copying it
DumpFile.Activate
LastRow = Range("A1").End(xlDown).Row
Range("A1:U1").AutoFilter field:=13, Criteria1:="*" & GLName & "*"
Range("A1:U1").AutoFilter field:=4, Criteria1:="ZR"
With ActiveSheet.AutoFilter.Range
    Range("A2:U" & LastRow).SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
End With

GLFile.Activate
Range("A2").PasteSpecial xlPasteAll
Application.CutCopyMode = False

'Copying Posting keys
HlastRow = Range("H1").End(xlDown).Row
Range("H2:H" & HlastRow).Copy
Sheets("Template").Activate
Range("J23").PasteSpecial xlPasteValues
Range("J23:J" & Range("J23").End(xlDown).Row).Value = 50
JNewLastRow = Range("J23").End(xlDown).Row + 1


'Copying Amount keys
Sheets("Support").Activate
Range("I2:I" & HlastRow).Copy
Sheets("Template").Activate
Range("O23").PasteSpecial xlPasteValues

'Copying Text
Sheets("Support").Activate
Range("M2:M" & HlastRow).Copy
Sheets("Template").Activate
Range("S23").PasteSpecial xlPasteValues

'Copying Posting key and amount as 40 posting key
Range("J23:S" & Range("O23").End(xlDown).Row).Copy

Range("J" & JNewLastRow).PasteSpecial xlPasteValues
Range("J" & JNewLastRow & ":J" & Range("J23").End(xlDown).Row).Value = 40


Curr_ency = Sheets("Support").Range("J2").Value
Company_Code = Sheets("Support").Range("A2").Value
Posting_Date_DocDate = Format(DateAdd("M", 0, Date), "dd.MM.yyyy")
Doc_Header = "CI#6_Taxes_" & Format(DateAdd("M", 0, Date), "dd.MM")
txt = Sheets("Support").Range("M2").Value
ValueDate = Format(Sheets("Support").Range("F2").Value, "dd.MM.yyyy")
Reference = "BSC AMERICAS"

JlastRow = Range("J23").End(xlDown).Row
Range("C23:C" & JlastRow).Value = 1
Range("D23:D" & JlastRow).Value = Company_Code
Range("E23:E" & JlastRow).Value = UserForm1.TextBox2.Value
Range("F23:F" & JlastRow).Value = Doc_Header
Range("G23:G" & JlastRow).Value = Posting_Date_DocDate
Range("H23:H" & JlastRow).Value = Posting_Date_DocDate
Range("I23:I" & JlastRow).Value = Curr_ency
Range("AD23").Value = Reference


'Setting account numbers on the basis of Posting key
For i = 23 To JlastRow

    PostingKey = Range("J" & i).Value
    If PostingKey = 50 Then
        Range("Y" & i).Value = ValueDate
        Range("K" & i).Value = 2530037

    Else
        Range("K" & i).Value = 2203002
        If GLName = "Vat 21%" Then
            Range("K" & i).Value = 2203027
        End If
    End If

Next i
Application.CutCopyMode = False
Range("C23").Select

GLFile.Save
GLFile.Close

End Sub

Private Sub Retention_Imposed(DumpFile As Workbook, ParentPath, ReportFolder As String, GLName As String, FinalFile, TemplateFile)

'Creating Reports flders if it doesn't exists

FolderPath = ParentPath & "\Reports\" & ReportFolder
If Dir(FolderPath, vbDirectory) = "" Then
    MkDir FolderPath
End If

'Copying Template_File to Report Folder
Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.CopyFile TemplateFile, FinalFile

'Opening GL File and clearing it
Set GLFile = Workbooks.Open(FinalFile)
GLFile.Activate
Sheets("Template").Activate
GLLastRow = Range("C23").End(xlDown).Row
Range("C23:AF" & GLLastRow).Clear
Sheets("Support").Activate
ActiveSheet.AutoFilterMode = False
Range("A2:T" & Range("T1").End(xlDown).Row).Clear

'Filtering Data from Dump File and Copying it
DumpFile.Activate
LastRow = Range("A1").End(xlDown).Row
Range("A1:U1").AutoFilter field:=13, Criteria1:="*" & GLName & "*"
Range("A1:U1").AutoFilter field:=4, Criteria1:="ZR"
With ActiveSheet.AutoFilter.Range
    Range("A2:U" & LastRow).SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
End With

GLFile.Activate
Range("A2").PasteSpecial xlPasteAll
Application.CutCopyMode = False

'Copying Posting keys
HlastRow = Range("H1").End(xlDown).Row
Range("H2:H" & HlastRow).Copy
Sheets("Template").Activate
Range("J23").PasteSpecial xlPasteValues
'Range("J23:J" & Range("J23").End(xlDown).Row).Value = 50
'JNewLastRow = Range("J23").End(xlDown).Row + 1


'Copying Amount keys
Sheets("Support").Activate
Range("I2:I" & HlastRow).Copy
Sheets("Template").Activate
Range("O23").PasteSpecial xlPasteValues

'Copying Text
Sheets("Support").Activate
Range("M2:M" & HlastRow).Copy
Sheets("Template").Activate
Range("S23").PasteSpecial xlPasteValues

'Copying Value Date
Sheets("Support").Activate
Range("G2:G" & HlastRow).Copy
Sheets("Template").Activate
Range("Y23").PasteSpecial xlPasteValues

'Copying Account Numbers
Sheets("Support").Activate
Range("Y2:Y" & HlastRow).Copy
Sheets("Template").Activate
Range("K23").PasteSpecial xlPasteValues

'Copying Posting key and Reversing posting keys
NewLastRow = Range("O23").End(xlDown).Row
Range("J23:Y" & NewLastRow).Copy
Range("J" & NewLastRow + 1).PasteSpecial xlPasteValues

For i = NewLastRow + 1 To Range("J23").End(xlDown).Row
    If Range("J" & NewLastRow + 1).Value = 40 Then
        Range("J" & NewLastRow + 1).Value = 50
        GoTo nx
    End If
    If Range("J" & NewLastRow + 1).Value = 50 Then
        Range("J" & NewLastRow + 1).Value = 40
        GoTo nx
    End If
nx:
Next i

'Removing minus
For i = 23 To Range("J23").End(xlDown).Row
    Amt = Range("O" & i).Value
    If InStr(Amt, "-") Then
       Amt = Replace(Amt, "-", "")
       Range("O" & i).Value = Amt
    End If
Next i



Curr_ency = Sheets("Support").Range("J2").Value
Company_Code = Sheets("Support").Range("A2").Value
Posting_Date_DocDate = Format(DateAdd("M", 0, Date), "dd.MM.yyyy")
Doc_Header = "CI#6_Taxes_" & Format(DateAdd("M", 0, Date), "dd.MM")
txt = Sheets("Support").Range("M2").Value
ValueDate = Format(Sheets("Support").Range("F2").Value, "dd.MM.yyyy")
Reference = "BSC AMERICAS"

JlastRow = Range("J23").End(xlDown).Row
Range("C23:C" & JlastRow).Value = 1
Range("D23:D" & JlastRow).Value = Company_Code
Range("E23:E" & JlastRow).Value = UserForm1.TextBox2.Value
Range("F23:F" & JlastRow).Value = Doc_Header
Range("G23:G" & JlastRow).Value = Posting_Date_DocDate
Range("H23:H" & JlastRow).Value = Posting_Date_DocDate
Range("I23:I" & JlastRow).Value = Curr_ency
Range("AD23").Value = Reference
Range("K" & NewLastRow + 1 & ":K" & JlastRow).Value = 2203040

Application.CutCopyMode = False
Range("C23").Select

GLFile.Save
GLFile.Close

End Sub

Private Sub AFIP(DumpFile As Workbook, ParentPath, ReportFolder, TemplateFile, FinalFile)

FolderPath = ParentPath & "\Reports\" & ReportFolder
If Dir(FolderPath, vbDirectory) = "" Then
    MkDir FolderPath
End If

'Copying Template_File to Report Folder
Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.CopyFile TemplateFile, FinalFile

'Opening GL File and clearing it
Set GLFile = Workbooks.Open(FinalFile)
GLFile.Activate
Sheets(1).Activate
GLLastRow = Range("C1").End(xlDown).Row
Range("A2:AR" & GLLastRow).Clear

'Filtering Data from Dump File and Copying it
DumpFile.Activate
LastRow = Range("A1").End(xlDown).Row
Range("A1:U1").AutoFilter field:=13, Criteria1:="*" & "TRANSFERENCE INTERBANKING" & "*"
Range("A1:U1").AutoFilter field:=4, Criteria1:="ZR"
With ActiveSheet.AutoFilter.Range
    Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).Copy
    GLFile.Activate
    Range("B2").PasteSpecial xlPasteAll
    DumpFile.Activate
    Range("I2:I" & LastRow).SpecialCells(xlCellTypeVisible).Copy
    GLFile.Activate
    Range("P2").PasteSpecial xlPasteAll
    DumpFile.Activate
    Range("G2:G" & LastRow).SpecialCells(xlCellTypeVisible).Copy
    GLFile.Activate
    Range("X2").PasteSpecial xlPasteAll
    'ProfitCenter = Range("R2:R" & LastRow).SpecialCells(xlCellTypeVisible).Value
    DumpFile.Activate
    Range("B2:B" & LastRow).SpecialCells(xlCellTypeVisible).Copy
    GLFile.Activate
    Range("N2").PasteSpecial xlPasteAll
    DumpFile.Activate
    Range("M2:M" & LastRow).SpecialCells(xlCellTypeVisible).Copy
    GLFile.Activate
    Range("U2").PasteSpecial xlPasteAll
    DumpFile.Activate
    Range("H2:H" & LastRow).SpecialCells(xlCellTypeVisible).Copy
    GLFile.Activate
    Range("O2").PasteSpecial xlPasteAll
    
End With

'Adding data to Final File
Range("A2:AR" & Range("P1").End(xlDown).Row).Copy
NewRow = Range("P1").End(xlDown).Row + 1
Range("A" & NewRow).PasteSpecial xlPasteAll
Range("O2:O" & NewRow - 1).Value = 50


'Current Month Dates
CurrDate = DateAdd("M", 0, Date)
LasttDay = DateSerial(Year(CurrDate), Month(CurrDate) + 1, 0)
LastDay = Format(CurrDate, "mmm yy")
DDate = Format(LasttDay, "dd-mmm-yy")
Period = Format(LasttDay, "M")
'setting up fixed column value

Dim Fixed_Data As ArrayList
Set Fixed_Data = New ArrayList

Fixed_Data.Add "1" 'Col A
Fixed_Data.Add DDate 'Col C
Fixed_Data.Add "SA" 'Col D
Fixed_Data.Add DDate 'Col E
Fixed_Data.Add Period ' 'Col F
Fixed_Data.Add "ARS" 'Col G
Fixed_Data.Add UserForm1.TextBox2.Value 'I
Fixed_Data.Add "CI#6_AFIP" 'Col J
'Fixed_Data.Add Account  'Col N
Fixed_Data.Add CompanyCode 'Col T
Fixed_Data.Add "BSC AMERICAS" 'Col AJ


Dim Fixed_Col As ArrayList
Set Fixed_Col = New ArrayList

Last_Row = Range("P1").End(xlDown).Row

Fixed_Col.Add "A2" 'Document Code
Fixed_Col.Add "C2:C" & Last_Row 'Document Code
Fixed_Col.Add "D2:D" & Last_Row 'SA'
Fixed_Col.Add "E2:E" & Last_Row 'Posting Date - DDate
Fixed_Col.Add "F2:F" & Last_Row 'Period
Fixed_Col.Add "G2:G" & Last_Row 'Currency
Fixed_Col.Add "I2:I" & Last_Row 'Refernce
Fixed_Col.Add "J2:J" & Last_Row 'Doc Header
'Fixed_Col.Add "N2:N" & Last_Row 'Account Number
Fixed_Col.Add "T2:T" & Last_Row 'Compnay Code
Fixed_Col.Add "AJ2" 'Refernce 1



For m = 0 To Fixed_Col.Count - 1
    Range(Fixed_Col(m)).Value = Fixed_Data(m)
Next m
ActiveSheet.Columns.AutoFit

Range("N" & NewRow & ":N" & Range("N1").End(xlDown).Row).Value = 2211017
Range("U" & NewRow & ":U" & Range("U1").End(xlDown).Row).Value = "24000PESVP007109003"
Range("AA" & NewRow & ":AA" & Range("U1").End(xlDown).Row).Value = "BS_A_INVEN"
Range("X" & NewRow & ":X" & Range("U1").End(xlDown).Row).Clear

For i = 2 To Range("P1").End(xlDown).Row
    Range("M" & i).Value = i - 1
Next i

GLFile.Save
GLFile.Close

End Sub
Private Sub COELSA(DumpFile As Workbook, ParentPath, ReportFolder, TemplateFile, FinalFile)

FolderPath = ParentPath & "\Reports\" & ReportFolder
If Dir(FolderPath, vbDirectory) = "" Then
    MkDir FolderPath
End If

'Copying Template_File to Report Folder
Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.CopyFile TemplateFile, FinalFile

'Opening GL File and clearing it
Set GLFile = Workbooks.Open(FinalFile)
GLFile.Activate
Sheets(1).Activate
GLLastRow = Range("C1").End(xlDown).Row
Range("A2:AR" & GLLastRow).Clear

'Filtering Data from Dump File and Copying it
DumpFile.Activate
LastRow = Range("A1").End(xlDown).Row
Range("A1:U1").AutoFilter field:=13, Criteria1:="*" & "COELSA" & "*"
Range("A1:U1").AutoFilter field:=4, Criteria1:="ZR"
With ActiveSheet.AutoFilter.Range
    Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).Copy
    GLFile.Activate
    Range("B2").PasteSpecial xlPasteAll
    DumpFile.Activate
    Range("I2:I" & LastRow).SpecialCells(xlCellTypeVisible).Copy
    GLFile.Activate
    Range("P2").PasteSpecial xlPasteAll
    DumpFile.Activate
    Range("G2:G" & LastRow).SpecialCells(xlCellTypeVisible).Copy
    GLFile.Activate
    Range("X2").PasteSpecial xlPasteAll
    'ProfitCenter = Range("R2:R" & LastRow).SpecialCells(xlCellTypeVisible).Value
    DumpFile.Activate
    Range("B2:B" & LastRow).SpecialCells(xlCellTypeVisible).Copy
    GLFile.Activate
    Range("N2").PasteSpecial xlPasteAll
    DumpFile.Activate
    Range("M2:M" & LastRow).SpecialCells(xlCellTypeVisible).Copy
    GLFile.Activate
    Range("U2").PasteSpecial xlPasteAll
    DumpFile.Activate
    Range("H2:H" & LastRow).SpecialCells(xlCellTypeVisible).Copy
    GLFile.Activate
    Range("O2").PasteSpecial xlPasteAll
    
End With

'Adding data to Final File
Range("A2:AR" & Range("P1").End(xlDown).Row).Copy
NewRow = Range("P1").End(xlDown).Row + 1
Range("A" & NewRow).PasteSpecial xlPasteAll
Range("O" & NewRow & ":O" & Range("P1").End(xlDown).Row).Value = 50


'Current Month Dates
CurrDate = DateAdd("M", 0, Date)
LasttDay = DateSerial(Year(CurrDate), Month(CurrDate) + 1, 0)
LastDay = Format(CurrDate, "mmm yy")
DDate = Format(LasttDay, "dd-mmm-yy")
Period = Format(LasttDay, "M")
'setting up fixed column value

Dim Fixed_Data As ArrayList
Set Fixed_Data = New ArrayList

Fixed_Data.Add "1" 'Col A
Fixed_Data.Add DDate 'Col C
Fixed_Data.Add "SA" 'Col D
Fixed_Data.Add DDate 'Col E
Fixed_Data.Add Period ' 'Col F
Fixed_Data.Add "ARS" 'Col G
Fixed_Data.Add UserForm1.TextBox2.Value 'I
Fixed_Data.Add "CI#3_Bank Fees_" & Format(DateAdd("M", 0, Date), "dd.mm") 'Col J
'Fixed_Data.Add Account  'Col N
Fixed_Data.Add CompanyCode 'Col T
Fixed_Data.Add "BSC AMERICAS" 'Col AJ


Dim Fixed_Col As ArrayList
Set Fixed_Col = New ArrayList

Last_Row = Range("P1").End(xlDown).Row

Fixed_Col.Add "A2" 'Document Code
Fixed_Col.Add "C2:C" & Last_Row 'Document Code
Fixed_Col.Add "D2:D" & Last_Row 'SA'
Fixed_Col.Add "E2:E" & Last_Row 'Posting Date - DDate
Fixed_Col.Add "F2:F" & Last_Row 'Period
Fixed_Col.Add "G2:G" & Last_Row 'Currency
Fixed_Col.Add "I2:I" & Last_Row 'Refernce
Fixed_Col.Add "J2:J" & Last_Row 'Doc Header
'Fixed_Col.Add "N2:N" & Last_Row 'Account Number
Fixed_Col.Add "T2:T" & Last_Row 'Compnay Code
Fixed_Col.Add "AJ2" 'Refernce 1


'Setting-up Serial number
For m = 0 To Fixed_Col.Count - 1
    Range(Fixed_Col(m)).Value = Fixed_Data(m)
Next m

Range("N" & NewRow & ":N" & Range("N1").End(xlDown).Row).Value = 4513000
'Range("U" & NewRow & ":U" & Range("U1").End(xlDown).Row).Value = "24000PESVP007109003"
Range("V" & NewRow & ":V" & Range("U1").End(xlDown).Row).Value = 886100
Range("X" & NewRow & ":X" & Range("U1").End(xlDown).Row).Clear

For i = 2 To Range("P1").End(xlDown).Row
    Range("M" & i).Value = i - 1
Next i
Range("X:X").NumberFormat = "d-mmm-yyyy"
ActiveSheet.Columns.AutoFit

'Creating new xlsx format file and putting data in it
ActiveSheet.UsedRange.Select
Selection.Copy
Set NewXlsx = Workbooks.Add

Sheets(1).Activate
Range("A1").PasteSpecial xlPasteAll
Range("A1:L1").Interior.ColorIndex = 35
Range("AJ2").Interior.ColorIndex = 6
Range("AM2").Interior.Color = RGB(255, 199, 127)
Range("I2:I" & Range("I1").End(xlDown).Row).Interior.ColorIndex = 6
ActiveSheet.Columns.AutoFit

NewXlsx.SaveAs Filename:=FolderPath & "\CI#3_COELSA_" & Format(DateAdd("M", 0, Date), "dd.mm") & ".xlsx", FileFormat:=xlOpenXMLWorkbook
NewXlsx.Close

GLFile.Save
GLFile.Close

End Sub

Private Sub Tax_Law(DumpFile As Workbook, ParentPath, ReportFolder, TemplateFile, FinalFile)

FolderPath = ParentPath & "\Reports\" & ReportFolder
If Dir(FolderPath, vbDirectory) = "" Then
    MkDir FolderPath
End If

'Copying Template_File to Report Folder
Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.CopyFile TemplateFile, FinalFile

'Opening GL File and clearing it
Set GLFile = Workbooks.Open(FinalFile)
GLFile.Activate
Sheets("WORKING").Activate
ActiveSheet.AutoFilterMode = False
GLLastRow = Range("C1").End(xlDown).Row
Range("A2:T" & GLLastRow).Clear

Sheets("JE").Activate
ActiveSheet.AutoFilterMode = False
Range("A4:AD4" & GLLastRow).Clear

'Filtering Data from Dump File and Copying it
DumpFile.Activate
LastRow = Range("A1").End(xlDown).Row
Range("A1:U1").AutoFilter field:=13, Criteria1:="*" & "TAX BY" & "*"
Range("A1:U1").AutoFilter field:=4, Criteria1:="ZR"
With ActiveSheet.AutoFilter.Range
    Range("A2:U" & LastRow).SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
End With

'Inserting data from Dump file to final file's working sheet
GLFile.Activate
Sheets("WORKING").Activate
Range("A2").PasteSpecial xlPasteAll
Application.CutCopyMode = False
NewLastRow = Range("N1").End(xlDown).Row
Range("U2").Value = "LAW"
Range("U2:AD2").Copy
Range("U2:AD" & NewLastRow).PasteSpecial xlPasteAll

'Setting Data in JE Sheet

'For ABS
Call Copier("V", "M") 'Copying amounts
Call Copier("B", "I") 'Copying Account Number
Call Copier("W", "H") 'Copying Posting Key
Call Copier("AD", "V") 'Copying Assignment Key
Call Copier("G", "W") 'Copying Value Date

'For 67%
Call Copier("Z", "M") 'Copying amounts
Call Copier("X", "I") 'Copying Account Number
Call Copier("Y", "H") 'Copying Posting Key
Call Copier("AD", "V") 'Copying Assignment Key
'Call Copier("G2", "W") 'Copying Value Date

'For 33%
Call Copier("AC", "M") 'Copying amounts
Call Copier("AA", "I") 'Copying Account Number
Call Copier("AB", "H") 'Copying Posting Key
Call Copier("AD", "V") 'Copying Assignment Key


'Adding Manual Data
Curr_ency = "ARS"
Company_Code = 1367
Posting_Date_DocDate = Format(DateAdd("M", 0, Date), "dd.MM.yyyy")
Doc_Header = "CI#6_Taxes -" & Format(DateAdd("M", 0, Date), "dd/MM")
'txt = Sheets("Support").Range("M2").Value
ValueDate = Format(Sheets("Support").Range("F2").Value, "dd.MM.yyyy")
Reference = "BSC AMERICAS"

JlastRow = Range("M4").End(xlDown).Row
Range("A4:A" & JlastRow).Value = 1
Range("B4:B" & JlastRow).Value = Company_Code
Range("C4:C" & JlastRow).Value = UserForm1.TextBox2.Value
Range("D4:D" & JlastRow).Value = Doc_Header
Range("E4:G" & JlastRow).Value = Posting_Date_DocDate
Range("F4:H" & JlastRow).Value = Posting_Date_DocDate
Range("G4:G" & JlastRow).Value = Curr_ency
Range("AB4").Value = Reference
Range("W4:W" & Range("W1").End(xlDown).Row).NumberFormat = "MM.dd.yyyy"

'Coloring Rows
For m = 4 To Range("I1").End(xlDown).Row
    If Range("I" & m).Value = 4513002 Then
        Range("A" & m & ":AD" & m).Interior.ColorIndex = 35
    End If
    If Range("I" & m).Value = 2400001 Then
        Range("A" & m & ":AD" & m).Interior.ColorIndex = 6
    End If
    
Next m
Columns.AutoFit

GLFile.Save
GLFile.Close

End Sub

Private Function Copier(CopyRange, PasteRange)

Sheets("WORKING").Activate
LastCellRow = Range("N1").End(xlDown).Row
Range(CopyRange & "2:" & CopyRange & LastCellRow).Copy

Sheets("JE").Activate
LastCellRow2 = Range(PasteRange & "1").End(xlDown).Row
Range(PasteRange & LastCellRow2 + 1).PasteSpecial xlPasteValues

End Function
Private Sub Bank_Fee(DumpFile As Workbook, ParentPath, ReportFolder, TemplateFile, FinalFile)

FolderPath = ParentPath & "\Reports\" & ReportFolder
If Dir(FolderPath, vbDirectory) = "" Then
    MkDir FolderPath
End If

'Copying Template_File to Report Folder
Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.CopyFile TemplateFile, FinalFile

'Opening GL File and clearing it
Set GLFile = Workbooks.Open(FinalFile)
GLFile.Activate
Sheets(1).Activate
GLLastRow = Range("C1").End(xlDown).Row
Range("A2:AR" & GLLastRow).Clear

'Filtering Data from Dump File and Copying it
DumpFile.Activate
LastRow = Range("A1").End(xlDown).Row
FilterValue = Array("/BAI/699/FEES ON RETURNED CHECKES /PT/DE/EI/COMISI", "/BAI/699/INVOICING INTERBANKING /PT/DE/EI/FACTURAC", "/BAI/699/A/C MAINTENANCE FEES /PT/DE/EI/C.MANT.C.C", "/BAI/699/PAYLINK-FEES /PT/DE/EI/COMISION PAYLINK B", "/BAI/699/LEGAL ANALYSIS FEES /PT/DE/EI/COMISION AN")
Range("A1:U1").AutoFilter field:=13, Criteria1:=FilterValue, Operator:=xlFilterValues
Range("A1:U1").AutoFilter field:=4, Criteria1:="ZR"
With ActiveSheet.AutoFilter.Range
    Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).Copy
    GLFile.Activate
    Range("B2").PasteSpecial xlPasteAll
    DumpFile.Activate
    Range("I2:I" & LastRow).SpecialCells(xlCellTypeVisible).Copy
    GLFile.Activate
    Range("P2").PasteSpecial xlPasteAll
    DumpFile.Activate
    Range("G2:G" & LastRow).SpecialCells(xlCellTypeVisible).Copy
    GLFile.Activate
    Range("X2").PasteSpecial xlPasteAll
    'ProfitCenter = Range("R2:R" & LastRow).SpecialCells(xlCellTypeVisible).Value
    DumpFile.Activate
    Range("B2:B" & LastRow).SpecialCells(xlCellTypeVisible).Copy
    GLFile.Activate
    Range("N2").PasteSpecial xlPasteAll
    DumpFile.Activate
    Range("M2:M" & LastRow).SpecialCells(xlCellTypeVisible).Copy
    GLFile.Activate
    Range("U2").PasteSpecial xlPasteAll
    DumpFile.Activate
    Range("H2:H" & LastRow).SpecialCells(xlCellTypeVisible).Copy
    GLFile.Activate
    Range("O2").PasteSpecial xlPasteAll
    
End With

'Adding data to Final File
Range("A2:AR" & Range("P1").End(xlDown).Row).Copy
NewRow = Range("P1").End(xlDown).Row + 1
Range("A" & NewRow).PasteSpecial xlPasteAll
Range("O2:O" & NewRow).Value = 50


'Current Month Dates
CurrDate = DateAdd("M", 0, Date)
LasttDay = DateSerial(Year(CurrDate), Month(CurrDate) + 1, 0)
LastDay = Format(CurrDate, "mmm yy")
DDate = Format(LasttDay, "dd-mmm-yy")
Period = Format(LasttDay, "M")
'setting up fixed column value

Dim Fixed_Data As ArrayList
Set Fixed_Data = New ArrayList

Fixed_Data.Add "1" 'Col A
Fixed_Data.Add DDate 'Col C
Fixed_Data.Add "SA" 'Col D
Fixed_Data.Add DDate 'Col E
Fixed_Data.Add Period ' 'Col F
Fixed_Data.Add "ARS" 'Col G
Fixed_Data.Add UserForm1.TextBox2.Value 'I
Fixed_Data.Add "CI#3_Bank Fees_ARS_" & Format(DateAdd("M", 0, Date), "dd.mm") 'Col J
'Fixed_Data.Add Account  'Col N
Fixed_Data.Add CompanyCode 'Col T
Fixed_Data.Add "BSC AMERICAS" 'Col AJ
Fixed_Data.Add "Bank Fee"


Dim Fixed_Col As ArrayList
Set Fixed_Col = New ArrayList

Last_Row = Range("P1").End(xlDown).Row

Fixed_Col.Add "A2" 'Document Code
Fixed_Col.Add "C2:C" & Last_Row 'Document Code
Fixed_Col.Add "D2:D" & Last_Row 'SA'
Fixed_Col.Add "E2:E" & Last_Row 'Posting Date - DDate
Fixed_Col.Add "F2:F" & Last_Row 'Period
Fixed_Col.Add "G2:G" & Last_Row 'Currency
Fixed_Col.Add "I2:I" & Last_Row 'Refernce
Fixed_Col.Add "J2:J" & Last_Row 'Doc Header
'Fixed_Col.Add "N2:N" & Last_Row 'Account Number
Fixed_Col.Add "T2:T" & Last_Row 'Compnay Code
Fixed_Col.Add "AJ2" 'Refernce 1
Fixed_Col.Add "AB2:AB" & Last_Row


'Setting-up Serial number
For m = 0 To Fixed_Col.Count - 1
    Range(Fixed_Col(m)).Value = Fixed_Data(m)
Next m

Range("N" & NewRow & ":N" & Range("N1").End(xlDown).Row).Value = 4513000
'Range("U" & NewRow & ":U" & Range("U1").End(xlDown).Row).Value = "24000PESVP007109003"
Range("V" & NewRow & ":V" & Range("U1").End(xlDown).Row).Value = 2530037
Range("X" & NewRow & ":X" & Range("U1").End(xlDown).Row).Clear

For i = 2 To Range("P1").End(xlDown).Row
    Range("M" & i).Value = i - 1
Next i
Range("X:X").NumberFormat = "d-mmm-yyyy"
ActiveSheet.Columns.AutoFit

'Creating new xlsx format file and putting data in it
ActiveSheet.UsedRange.Select
Selection.Copy
Set NewXlsx = Workbooks.Add


Sheets(1).Activate
Range("A1").PasteSpecial xlPasteAll
Range("A1:L1").Interior.ColorIndex = 35
Range("AJ2").Interior.ColorIndex = 6
Range("AM2").Interior.Color = RGB(255, 199, 127)
Range("I2:I" & Range("I1").End(xlDown).Row).Interior.ColorIndex = 6
ActiveSheet.Columns.AutoFit

NewXlsx.SaveAs Filename:=FolderPath & "\Weekly_CI#3_Bank Fees_ARS_" & Format(DateAdd("M", 0, Date), "dd.mm") & ".xlsx", FileFormat:=xlOpenXMLWorkbook

NewXlsx.Close

GLFile.Save
GLFile.Close

End Sub

