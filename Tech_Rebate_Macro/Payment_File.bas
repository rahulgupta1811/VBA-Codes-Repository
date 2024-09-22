Attribute VB_Name = "Payment_File"
Public Segfile As Workbook
Public SegFilelocation As String
Public APCI As String
Public APCIPPA As String
Public IPC As String
Public IPCPBA As String
Public Liberty As String
Public Reliant As String
Public APSC As String
Public AlastRow As Long
Public lastRow As Long
Public DateLM As Date
Public Confile As String
Public confilewb As Workbook
Public myDate As Variant
Sub UpdateFile()
Dim ParentPath As String

'ParentPath = "\\ddcf2015\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\Payment Files"
ParentPath = ActiveWorkbook.Path

CurrYear = Year(Date)
Currmonth = Format(DateAdd("M", -2, Now), "MM")
lastmonth = CurrYear & Currmonth

'Initializing Payment files locations
APCI = ParentPath & "\Payment Files\APCI\APCI Tech Payment_" & lastmonth & " Working file.xlsx"
APCIPPA = ParentPath & "\Payment Files\APCI New Non Compliant TR (Working File)_New.xlsx"
IPC = ParentPath & "\Payment Files\IPC\IPC Payment Summary " & lastmonth & "_Working File.xlsx"
IPCPBA = ParentPath & "\Payment Files\Unaf_PBA\UNAF_PBA Tech File " & lastmonth & ".xlsx"
Liberty = ParentPath & "\Liberty\Liberty " & lastmonth & " Payments.xlsx"
Reliant = ParentPath & "\Payment Files\Reliant\Reliant Tech Rebate Payment - " & lastmonth & ".xlsx"
APSC = ParentPath & "\Payment Files\APSC\APSC Tech Payment Summary " & lastmonth & " - Working File.xlsx"
Confile = ParentPath & "\Payment Files\Tech Rebate Payments_Consolidated WC.xlsx"

'Disabling Alerts for smooth rendering
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'Segregated File intializing
SegFilelocation = ParentPath & "Payment Files\Tech Rebate Payment Files_Latest from Apr'20 Onwards.xlsx"
Set Segfile = Workbooks.Open(SegFilelocation)

'Updating APCI Payment Payments
Call APCI_Update

'Updating APCI PPA Payments
Call APCI_PPA

'Updating APSC Payments
Call APSC_Update

'Updating Reliant Payments
Call Reliant_Update

'Updating IPC Payments
Call IPC_Update

'Updating IPC PBA Payments
Call IPCPBA_Update

'Saving and Closing
Call Save_Close

End Sub
Sub APCI_Update()
'Opening APCI Payment File
Dim APCIFile As Workbook
Set APCIFile = Workbooks.Open(APCI)

'Coying Customer Name Column
APCIFile.Activate
Sheets("Payment Upload").Activate
AlastRow = Range("A6").End(xlDown).Row
Range("A" & AlastRow).Select
Range("A6:A" & AlastRow).Copy

'Pasting Customer Name of APCI
Segfile.Activate
Sheets("APCI").Activate
lastRow = Range("A1").End(xlDown).Row
NewRow = lastRow + 1
Range("C" & NewRow).PasteSpecial xlPasteValues

'Coying Account Number Column
APCIFile.Activate
Sheets("Payment Upload").Activate
AlastRow = Range("A6").End(xlDown).Row
Range("B" & AlastRow).Select
Range("B6:B" & AlastRow).Copy

'Pasting Account Number of APCI
Segfile.Activate
Sheets("APCI").Activate
lastRow = Range("A1").End(xlDown).Row
NewRow = lastRow + 1
Range("D" & NewRow).PasteSpecial xlPasteValues

'Coying OVERPAID AMT, TECH REBATE, Supplemental, FINAL REBATE PAID,
APCIFile.Activate
Sheets("Payment Upload").Activate
AlastRow = Range("A6").End(xlDown).Row
Range("J" & AlastRow).Select
Range("I6:K" & AlastRow).Copy

'Pasting OVERPAID AMT, TECH REBATE, Supplemental, FINAL REBATE PAID,
Segfile.Activate
Sheets("APCI").Activate
lastRow = Range("A1").End(xlDown).Row
NewRow = lastRow + 1
Range("E" & NewRow).PasteSpecial xlPasteValues

'Copying Rebate Month and Paid Month
APCIFile.Activate
Sheets("Payment Upload").Activate
Range("L6:M6").Copy

'Pasting Month and Paid Month
Segfile.Activate
Sheets("APCI").Activate
lastRow = Range("A1").End(xlDown).Row
NewRow = lastRow + 1
Range("H" & NewRow).PasteSpecial xlPasteAll
RebMonth = Range("H" & NewRow)

'Dim myDate As Variant
myDate = Range("H" & NewRow)
NewM = Right(myDate, 2)
NewY = Left(myDate, 4)
''NewDate As String
NewDate = NewM & "/01/" & NewY
Range("H" & NewRow).value = NewDate
myDate = Range("I" & NewRow)
NewM = Right(myDate, 2)
NewY = Left(myDate, 4)
NewDate = NewM & "/01/" & NewY
Range("I" & NewRow).value = NewDate

Range("H" & NewRow & ":" & "I" & NewRow).NumberFormat = "mmm-yy"
ClastRow = Range("C" & NewRow).End(xlDown).Row
Range("H" & NewRow & ":" & "I" & NewRow).Copy
Range("H" & NewRow & ":" & "I" & ClastRow).PasteSpecial xlPasteAll

'Coying Compliance
APCIFile.Activate
Sheets("Payment Upload").Activate
AlastRow = Range("A6").End(xlDown).Row
Range("N" & AlastRow).Select
Range("N6:N" & AlastRow).Copy

'Pasting Compliance
Segfile.Activate
Sheets("APCI").Activate
lastRow = Range("A1").End(xlDown).Row
NewRow = lastRow + 1
Range("K" & NewRow).PasteSpecial xlPasteValues

'Coying Month Comments
APCIFile.Activate
Sheets("Payment Upload").Activate
AlastRow = Range("A6").End(xlDown).Row
Range("AC" & AlastRow).Select
Range("AC6:AC" & AlastRow).Copy

'Pasting Month Comments
Segfile.Activate
Sheets("APCI").Activate
lastRow = Range("A1").End(xlDown).Row
NewRow = lastRow + 1
Range("L" & NewRow).PasteSpecial xlPasteValues

'Writing Manual Values for APCI
NewLrow = Range("C" & NewRow).End(xlDown).Row
Range("A" & NewRow & ":A" & NewLrow).value = "APCI"

'Range("A" & lastRow & ":M" & lastRow).Copy
'Range("A" & NewRow & ":M" & NewLrow).PasteSpecial xlPasteFormats

'Putting Data in Consolidated File
Segfile.Activate
Sheets("APCI").Activate
Range("A" & NewRow & ":" & "N" & ClastRow).Copy

Set confilewb = Workbooks.Open(Confile)
LastConRow = Range("A:A").End(xlDown).Row
nLastConRow = LastConRow + 1
Range("A" & nLastConRow).PasteSpecial xlPasteValues

confilewb.Save
confilewb.Close
APCIFile.Close
End Sub
Sub APCI_PPA()

Dim APCIPBAfile As Workbook
Set APCIPBAfile = Workbooks.Open(APCIPPA)

APCIPBAfile.Activate
Sheets("APCI New ").Activate
updLastRow = Range("H6").End(xlDown).value
Datfilter = Format(updLastRow, "mmm-yy")
With Range("A6:AD6")
    .AutoFilter Field:=1, Criteria1:="<>"
    .AutoFilter Field:=8, Criteria1:=Datfilter
End With

'Copying Customer Name
Range("A6").Activate
ActiveSheet.AutoFilter.Range.Offset(1, 1).SpecialCells(xlCellTypeVisible).Cells(1, 1).Select
filterfirstr = ActiveCell.Address
updLastRow = Range("H6").End(xlDown).Row
updLastRow = updLastRow + 1
UpdatedRow = "B" & updLastRow
Range(filterfirstr & ":" & UpdatedRow).SpecialCells(xlCellTypeVisible).Copy

'Pasting Customer Name
Segfile.Sheets("APCI").Activate
lastRow = Range("A1").End(xlDown).Row
NewLastRow = lastRow + 1
Range("C" & NewLastRow).PasteSpecial xlPasteValues

'Copying Customer Number
APCIPBAfile.Activate
Sheets("APCI New ").Activate
Range("C6").Select
ActiveSheet.AutoFilter.Range.Offset(1, 2).SpecialCells(xlCellTypeVisible).Cells(1, 1).Select
filterfirstr = ActiveCell.Address
updLastRow = Range("H6").End(xlDown).Row
updLastRow = updLastRow + 1
UpdatedRow = "C" & updLastRow
Range(filterfirstr & ":" & UpdatedRow).SpecialCells(xlCellTypeVisible).Copy

'Pasting Customer Name
Segfile.Sheets("APCI").Activate
lastRow = Range("A1").End(xlDown).Row
NewLastRow = lastRow + 1
Range("D" & NewLastRow).PasteSpecial xlPasteValues

'Copying  Base Rebate     Supplemental Rebate     TOTAL Tech Rebate
APCIPBAfile.Activate
Sheets("APCI New ").Activate
Range("G6").Select
ActiveSheet.AutoFilter.Range.Offset(1, 6).SpecialCells(xlCellTypeVisible).Cells(1, 1).Select
filterfirstr = ActiveCell.Address
updLastRow = Range("H6").End(xlDown).Row
updLastRow = updLastRow + 1
UpdatedRow = "G" & updLastRow
Range(filterfirstr & ":" & UpdatedRow).SpecialCells(xlCellTypeVisible).Copy

'Pasting  Base Rebate     Supplemental Rebate     TOTAL Tech Rebate
Segfile.Sheets("APCI").Activate
lastRow = Range("A1").End(xlDown).Row
NewLastRow = lastRow + 1
Range("F" & NewLastRow).PasteSpecial xlPasteValues
Range("G" & NewLastRow).PasteSpecial xlPasteValues

'Copying Rebate Month
APCIPBAfile.Activate
Sheets("APCI New ").Activate
Range("H6").Select
ActiveSheet.AutoFilter.Range.Offset(1, 7).SpecialCells(xlCellTypeVisible).Cells(1, 1).Select
filterfirstr = ActiveCell.Address
updLastRow = Range("H6").End(xlDown).Row
updLastRow = updLastRow + 1
UpdatedRow = "H" & updLastRow
Range(filterfirstr & ":" & UpdatedRow).SpecialCells(xlCellTypeVisible).Copy

'Pasting Rebate Month
Segfile.Activate
Sheets("APCI").Activate
lastRow = Range("A1").End(xlDown).Row
NewRow = lastRow + 1
Range("H" & NewRow).PasteSpecial xlPasteValues

'Copying Paid Month
APCIPBAfile.Activate
Sheets("APCI New ").Activate
Range("I6").Select
ActiveSheet.AutoFilter.Range.Offset(1, 8).SpecialCells(xlCellTypeVisible).Cells(1, 1).Select
filterfirstr = ActiveCell.Address
updLastRow = Range("H6").End(xlDown).Row
updLastRow = updLastRow + 1
UpdatedRow = "I" & updLastRow
Range(filterfirstr & ":" & UpdatedRow).SpecialCells(xlCellTypeVisible).Copy

'Pasting Paid Month
Segfile.Activate
Sheets("APCI").Activate
lastRow = Range("A1").End(xlDown).Row
NewRow = lastRow + 1
Range("I" & NewRow).PasteSpecial xlPasteValues

'Copying Compliance
APCIPBAfile.Activate
Sheets("APCI New ").Activate
Range("N6").Select
ActiveSheet.AutoFilter.Range.Offset(1, 10).SpecialCells(xlCellTypeVisible).Cells(1, 1).Select
filterfirstr = ActiveCell.Address
updLastRow = Range("H6").End(xlDown).Row
updLastRow = updLastRow + 1
UpdatedRow = "N" & updLastRow
Range(filterfirstr & ":" & UpdatedRow).SpecialCells(xlCellTypeVisible).Copy

'Pasting Compliance
Segfile.Activate
Sheets("APCI").Activate
lastRow = Range("A1").End(xlDown).Row
NewRow = lastRow + 1
Range("K" & NewRow).PasteSpecial xlPasteValues

'Copying Comments
APCIPBAfile.Activate
Sheets("APCI New ").Activate
Range("X6").Select
ActiveSheet.AutoFilter.Range.Offset(1, 23).SpecialCells(xlCellTypeVisible).Cells(1, 1).Select
filterfirstr = ActiveCell.Address
updLastRow = Range("H6").End(xlDown).Row
updLastRow = updLastRow + 1
UpdatedRow = "X" & updLastRow
Range(filterfirstr & ":" & UpdatedRow).SpecialCells(xlCellTypeVisible).Copy

'Pasting Comments
Segfile.Activate
Sheets("APCI").Activate
lastRow = Range("A1").End(xlDown).Row
NewRow = lastRow + 1
Range("L" & NewRow).PasteSpecial xlPasteValues

'Setting Cell Values
LPrevRow = Range("A1").End(xlDown).Row
LPrevRow = LPrevRow + 1
NLRow = Range("C1").End(xlDown).Row
Range("A" & LPrevRow & ":A" & NLRow).value = "APCI"
Range("J" & LPrevRow & ":J" & NLRow).value = "APCI PPA"
'Range("N" & NewRow & ":N" & NewLrow).Copy
'Range("K" & NewRow).PasteSpecial xlPasteValues
'Range("N" & NewRow & ":N" & NewLrow).Clear

''Putting Data in Consolidated File
Segfile.Activate
Sheets("APCI").Activate
Range("A" & LPrevRow & ":" & "N" & NLRow).Copy


Set confilewb = Workbooks.Open(Confile)
LastConRow = Range("A:A").End(xlDown).Row
nLastConRow = LastConRow + 1
Range("A" & nLastConRow).PasteSpecial xlPasteValues

confilewb.Save
confilewb.Close
APCIPBAfile.Close

End Sub
Sub APSC_Update()

Segfile.Activate
Sheets("APSC").Activate

Dim APSEfile As Workbook
Set APSEfile = Workbooks.Open(APSC)

'Coying Customer Name and Number Column
APSEfile.Activate
Sheets("Payment File").Activate
AlastRow = Range("A6").End(xlDown).Row
Range("B6:C" & AlastRow).Copy

'Pasting Name and Number
Segfile.Activate
Sheets("APSC").Activate
lastRow = Range("A1").End(xlDown).Row
NewRow = lastRow + 1
Range("C" & NewRow).PasteSpecial xlPasteValues

'Coying OVERPAID AMT, TECH REBATE, Supplemental, FINAL REBATE PAID,
APSEfile.Activate
Sheets("Payment File").Activate
AlastRow = Range("A6").End(xlDown).Row
Range("J" & AlastRow).Select
Range("H6:J" & AlastRow).Copy

'Pasting OVERPAID AMT, TECH REBATE, Supplemental, FINAL REBATE PAID,
Segfile.Activate
Sheets("APSC").Activate
lastRow = Range("A1").End(xlDown).Row
NewRow = lastRow + 1
Range("E" & NewRow).PasteSpecial xlPasteValues

'Copying Rebate Month and Paid Month
APSEfile.Activate
Sheets("Payment File").Activate
Range("K6:L6").Copy


'Pasting Month and Paid Month
Segfile.Activate
Sheets("APSC").Activate
lastRow = Range("A1").End(xlDown).Row
NewRow = lastRow + 1
Range("H" & NewRow).PasteSpecial xlPasteAll
RebMonth = Range("H" & NewRow)

'Dim myDate As Variant
myDate = Range("H" & NewRow)
NewM = Right(myDate, 2)
NewY = Left(myDate, 4)
'NewDate As String
NewDate = NewM & "/01/" & NewY
Range("H" & NewRow).value = NewDate
myDate = Range("I" & NewRow)
NewM = Right(myDate, 2)
NewY = Left(myDate, 4)
NewDate = NewM & "/01/" & NewY
Range("I" & NewRow).value = NewDate

Range("H" & NewRow & ":" & "I" & NewRow).NumberFormat = "mmm-yy"
ClastRow = Range("C" & NewRow).End(xlDown).Row
Range("H" & NewRow & ":" & "I" & NewRow).Copy
Range("H" & NewRow & ":" & "I" & ClastRow).PasteSpecial xlPasteValues
Range("H" & NewRow & ":" & "I" & ClastRow).NumberFormat = "mmm-yy"

'Copying Comments and Aniversary date
APSEfile.Activate
Sheets("Payment File").Activate
AlastRow = Range("A6").End(xlDown).Row
Range("AD" & AlastRow).Select
Range("R6:S" & AlastRow).Copy


'Pasting Comments and Aniversary date
Segfile.Activate
Sheets("APSC").Activate
lastRow = Range("A1").End(xlDown).Row
NewRow = lastRow + 1
Range("L" & NewRow).PasteSpecial xlPasteValues

'Setting Cell Values
NewLrow = Range("C" & NewRow).End(xlDown).Row
Range("A" & NewRow & ":A" & NewLrow).value = "APSC"
AtLastRow = Range("A:A").End(xlDown).Row
Range("A" & lastRow & ":M" & lastRow).Copy
Range("A" & lastRow & ":M" & AtLastRow).PasteSpecial xlPasteFormats

'Putting Data in Consolidated File
Segfile.Activate
Sheets("APSC").Activate
Range("A" & NewRow & ":" & "N" & NewLrow).Copy


Set confilewb = Workbooks.Open(Confile)
LastConRow = Range("A:A").End(xlDown).Row
nLastConRow = LastConRow + 1
Range("A" & nLastConRow).PasteSpecial xlPasteValues
Range("H:I").NumberFormat = "mmm-yy"

confilewb.Save
confilewb.Close
APSEfile.Close

End Sub
Sub Reliant_Update()

Dim ReliantFile As Workbook
Set ReliantFile = Workbooks.Open(Reliant)
Sheets("Validation").Activate

'Copying Customer

lastRow = Range("A3").End(xlDown).Row
Range("A4:B" & lastRow).Copy

'Pasting Customer
Segfile.Activate
Sheets("Reliant").Activate

nlastRow = Range("A:A").End(xlDown).Row
nlastRow = nlastRow + 1
Range("C" & nlastRow).PasteSpecial xlPasteValues

'Copying Rebate Amount
ReliantFile.Activate
Sheets("Validation").Activate
Range("P4:P" & lastRow).Copy

'Pasting Rebate Amount
Segfile.Activate
Sheets("Reliant").Activate
Range("G" & nlastRow).PasteSpecial xlPasteValues

'Copying Comments
ReliantFile.Activate
Sheets("Validation").Activate
Range("Q4:Q" & lastRow).Copy

'Pasting Rebate Amount
Segfile.Activate
Sheets("Reliant").Activate
Range("L" & nlastRow).PasteSpecial xlPasteValues

'Copying Dates
ReliantFile.Activate
Sheets("Validation").Activate
Range("G4:H" & lastRow).Copy

'Pasting Dates
Segfile.Activate
Sheets("Reliant").Activate
lastRow = Range("A1").End(xlDown).Row
NewRow = lastRow + 1
Range("H" & NewRow).PasteSpecial xlPasteAll
RebMonth = Range("H" & NewRow)

'Dim myDate As Variant
myDate = Range("H" & NewRow)
NewM = Right(myDate, 2)
NewY = Left(myDate, 4)
'NewDate As String
NewDate = NewM & "/01/" & NewY
Range("H" & NewRow).value = NewDate
myDate = Range("I" & NewRow)
NewM = Right(myDate, 2)
NewY = Left(myDate, 4)
NewDate = NewM & "/01/" & NewY
Range("I" & NewRow).value = NewDate

Range("H" & NewRow & ":" & "I" & NewRow).NumberFormat = "mmm-yy"
ClastRow = Range("C" & NewRow).End(xlDown).Row
Range("H" & NewRow & ":" & "I" & NewRow).Copy
Range("H" & NewRow & ":" & "I" & ClastRow).PasteSpecial xlPasteAll

NewLrow = Range("C" & nlastRow).End(xlDown).Row
Range("A" & nlastRow & ":A" & NewLrow).value = "Reliant"

Range("G" & lastRow & ":I" & lastRow).Copy
Range("G" & lastRow & ":I" & NewLrow).PasteSpecial xlPasteFormats

'Putting Data in Consolidated File
Segfile.Activate
Sheets("Reliant").Activate
Range("A" & nlastRow & ":" & "N" & NewLrow).Copy


Set confilewb = Workbooks.Open(Confile)
LastConRow = Range("A:A").End(xlDown).Row
nLastConRow = LastConRow + 1
Range("A" & nLastConRow).PasteSpecial xlPasteValues
Range("H:I").NumberFormat = "mmm-yy"
Range("I:I").Font.Bold = False
Range("I1").Font.Bold = True

confilewb.Save
confilewb.Close
ReliantFile.Close

End Sub
Sub IPC_Update()

Dim IPCfile As Workbook
Set IPCfile = Workbooks.Open(IPC)

'Copying Old/New
IPCfile.Activate
Sheets(1).Activate
lastRow = Range("A6").End(xlDown).Row
Range("A7:A" & lastRow).Copy

'Pasting Old/New
Segfile.Activate
Sheets("IPC&PBA").Activate

SegLastRow = Range("A1").End(xlDown).Row
NewLastRow = SegLastRow + 1
Range("B" & NewLastRow).PasteSpecial xlPasteValues

'Copying IPC Text for A col
IPCfile.Activate
Sheets(1).Activate
Range("B7:B" & lastRow).Copy

'Pasting IPC Text for A col
Segfile.Activate
Sheets("IPC&PBA").Activate
Range("A" & NewLastRow).PasteSpecial xlPasteValues

'Copying Customer Number and Name
IPCfile.Activate
Sheets(1).Activate
Range("C7:D" & lastRow).Copy

'Pasting IPC Text for A col
Segfile.Activate
Sheets("IPC&PBA").Activate
Range("C" & NewLastRow).PasteSpecial xlPasteValues

'Copying Rebates
IPCfile.Activate
Sheets(1).Activate
Range("I7:K" & lastRow).Copy

'Pasting Rebates
Segfile.Activate
Sheets("IPC&PBA").Activate
Range("E" & NewLastRow).PasteSpecial xlPasteValues

'Copying Compliance
IPCfile.Activate
Sheets(1).Activate
Range("Q7:Q" & lastRow).Copy

'Pasting Compliance
Segfile.Activate
Sheets("IPC&PBA").Activate
Range("K" & NewLastRow).PasteSpecial xlPasteValues

'Copying Comments
IPCfile.Activate
Sheets(1).Activate
Range("Y7:Y" & lastRow).Copy

'Pasting Comments
Segfile.Activate
Sheets("IPC&PBA").Activate
Range("L" & NewLastRow).PasteSpecial xlPasteValues

'Copying Anniversary Date
IPCfile.Activate
Sheets(1).Activate
Range("AD7:AD" & lastRow).Copy

'Pasting Anniversary Date
Segfile.Activate
Sheets("IPC&PBA").Activate
Range("M" & NewLastRow).PasteSpecial xlPasteValues

'Copying Rebate Dates
IPCfile.Activate
Sheets(1).Activate
Range("O7:P" & lastRow).Copy

'Pasting Rebate Date
Segfile.Activate
Sheets("IPC&PBA").Activate
Range("H" & NewLastRow).PasteSpecial xlPasteAll

'Dim myDate As Variant
myDate = Range("H" & NewLastRow).value
NewM = Right(myDate, 2)
NewY = Left(myDate, 4)
'NewDate As String
NewDate = NewM & "/01/" & NewY
Dim NDate As String
NDate = Format(NewDate, "MMM-YY")
Range("H" & NewLastRow).value = NDate
myDate = Range("I" & NewLastRow).value
NewM = Right(myDate, 2)
NewY = Left(myDate, 4)
NewDate = NewM & "/01/" & NewY
Range("I" & NewLastRow).value = NewDate

NDate = Format(NewDate, "MMM-YY")
Range("I" & NewLastRow).value = NDate
ClastRow = Range("C1").End(xlDown).Row
Range("H" & NewLastRow & ":" & "I" & NewLastRow).Copy
Range("H" & NewLastRow & ":" & "I" & ClastRow).PasteSpecial xlPasteAll
Range("H" & NewLastRow & ":" & "I" & ClastRow).NumberFormat = "mmm-yy"

'Setting Formats
Segfile.Activate
Sheets("IPC&PBA").Activate
Range("A" & lastRow & ":M" & lastRow).Copy
AtLastRow = Range("A" & lastRow).End(xlDown).Row
Range("A" & NewLastRow & ":M" & AtLastRow).PasteSpecial xlPasteFormats

'Putting Data in Consolidated File
Segfile.Activate
Sheets("IPC&PBA").Activate
NewLrow = Range("A" & NewLastRow).End(xlDown).Row
Range("A" & NewLastRow & ":" & "N" & NewLrow).Copy


Set confilewb = Workbooks.Open(Confile)
LastConRow = Range("A:A").End(xlDown).Row
nLastConRow = LastConRow + 1
Range("A" & nLastConRow).PasteSpecial xlPasteValues

confilewb.Save
confilewb.Close
IPCfile.Close

End Sub
Sub IPCPBA_Update()

Dim IPCPBAfile As Workbook
Set IPCPBAfile = Workbooks.Open(IPCPBA)

Sheets("PBA").Activate
lastRow = Range("A6").End(xlDown).Row
ActiveSheet.AutoFilterMode = False
'Copying New/Old
IPCPBAfile.Activate
Sheets("PBA").Activate
Range("A6:A" & lastRow).Copy

'Pasting New/Old
Segfile.Activate
Sheets("IPC&PBA").Activate
SegfileLastRow = Range("A1").End(xlDown).Row
NewRow = SegfileLastRow + 1
Range("B" & NewRow).PasteSpecial xlPasteValues

'Copying IPC text to COl A
IPCPBAfile.Activate
Sheets("PBA").Activate
Range("B6:B" & lastRow).Copy

'Pasting IPC text to COl A
Segfile.Activate
Sheets("IPC&PBA").Activate
Range("A" & NewRow).PasteSpecial xlPasteValues

'Copying Customer Number and Name
IPCPBAfile.Activate
Sheets("PBA").Activate
Range("C6:D" & lastRow).Copy

'Pasting Customer Number and Name
Segfile.Activate
Sheets("IPC&PBA").Activate
Range("C" & NewRow).PasteSpecial xlPasteValues

'Copying Rebate Values
IPCPBAfile.Activate
Sheets("PBA").Activate
Range("I6:K" & lastRow).Copy

'Pasting Rebate Values
Segfile.Activate
Sheets("IPC&PBA").Activate
Range("E" & NewRow).PasteSpecial xlPasteValues

'Copying Compliance
IPCPBAfile.Activate
Sheets("PBA").Activate
Range("N6:N" & lastRow).Copy

'Pasting Compliance
Segfile.Activate
Sheets("IPC&PBA").Activate
Range("K" & NewRow).PasteSpecial xlPasteValues

'Copying Comments
IPCPBAfile.Activate
Sheets("PBA").Activate
Range("U6:U" & lastRow).Copy

'Pasting Comments
Segfile.Activate
Sheets("IPC&PBA").Activate
Range("L" & NewRow).PasteSpecial xlPasteValues

'Copying Anniversary Dates
IPCPBAfile.Activate
Sheets("PBA").Activate
Range("W6:W" & lastRow).Copy
    
'Pasting Anniversary Dates
Segfile.Activate
Sheets("IPC&PBA").Activate
Range("M" & NewRow).PasteSpecial xlPasteValues

'Copying Rebate Date
IPCPBAfile.Activate
Sheets("PBA").Activate
ActiveSheet.AutoFilterMode = False
Range("L6:M" & lastRow).Copy

'Pasting Rebate Date
Segfile.Activate
Sheets("IPC&PBA").Activate
Range("H" & NewRow).PasteSpecial xlPasteValues

'Dim myDate As Variant
myDate = Range("H" & NewRow)
NewM = Right(myDate, 2)
NewY = Left(myDate, 4)
'NewDate As String
NewDate = NewM & "/01/" & NewY
Range("H" & NewRow).value = NewDate
myDate = Range("I" & NewRow)
NewM = Right(myDate, 2)
NewY = Left(myDate, 4)
NewDate = NewM & "/01/" & NewY
Range("I" & NewRow).value = NewDate

Range("H" & NewRow & ":" & "I" & NewRow).NumberFormat = "mmm-yy"
ClastRow = Range("C1").End(xlDown).Row
Range("H" & NewRow & ":" & "I" & NewRow).Copy
Range("H" & NewRow & ":" & "I" & ClastRow).PasteSpecial xlPasteAll

'Changing Anniversary Date Format
nlastRow = Range("A1").End(xlDown).Row
Range("M" & NewRow).Select
Range("M" & NewRow & ":M" & nlastRow).Font.FontStyle = "Calibri"
Range("M" & NewRow & ":M" & nlastRow).Font.Size = 11

'Setting Formats
Range("J" & NewRow & ":J" & nlastRow).value = "PBA"
Range("A" & lastRow & ":M" & lastRow).Copy
Range("A" & NewRow & ":M" & nlastRow).PasteSpecial xlPasteFormats


'Putting Data in Consolidated File
Segfile.Activate
Sheets("IPC&PBA").Activate
NewLrow = Range("A" & NewRow).End(xlDown).Row
Range("A" & NewRow & ":" & "N" & NewLrow).Copy

Set confilewb = Workbooks.Open(Confile)
Sheets("Tech Payments").Activate
LastConRow = Range("A:A").End(xlDown).Row
nLastConRow = LastConRow + 1
Range("A" & nLastConRow).PasteSpecial xlPasteValues
Nlast = Range("H2").End(xlDown).Row
Range("H2:I2").Copy
Range("H2:I" & Nlast).PasteSpecial xlPasteFormats
Range("H:I").NumberFormat = "mmm-yy"

confilewb.Save
confilewb.Close
IPCPBAfile.Close
End Sub
Sub Save_Close()
'Saving and Closing

Segfile.Save
Segfile.Close
MsgBox "Done", vbInformation, "Success"

End Sub
