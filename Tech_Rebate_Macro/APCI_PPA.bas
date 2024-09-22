Attribute VB_Name = "APCI_PPA"
Sub APCIPPA()

Dim APCIFile As Workbook
Dim LastTwoMonth
LastToMonth = DateAdd("M", -2, Date)
'Date on File Name
LastTwoMonth = Format(LastToMonth, "YYYYMM")
LastTMonth = Format(LastToMonth, "(mmm'yy Rbt\s)")

LastOMonth = DateAdd("M", -1, Date)
'Date on File Name
LastOneMonth = Format(LastOMonth, "MM mmmm'yy ")
LYear = Format(LastOMonth, "yyyy")

'Openning Last 2nd Month APCI Payment file to get non compliant Data
APCIFilePath = "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Payment Files\" & LYear & "\" & LastOneMonth & LastTMonth & "\APCI\APCI Tech Payment_" & LastTwoMonth & " Working file.xlsx"
Set APCIFile = Workbooks.Open(APCIFilePath)

'Openning PPA working file and removing active Filter
Dim PPAfile As Workbook
Set PPAfile = Workbooks.Open("\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\Payment Files\APCI New Non Compliant TR (Working File)_New.xlsx")
PPAfile.Activate
Sheets("APCI New ").Activate
ActiveSheet.Range("A6:AD6").AutoFilter ' Removing Filter
PPALastRow = Range("B6").End(xlDown).Row
PPALastRow = PPALastRow + 1

'Activating APCI Payment
APCIFile.Activate
Sheets("Payment Upload").Activate
ActiveSheet.Range("A5:AE5").AutoFilter Field:=14, Criteria1:="N" ' Setting filter on Non Compliant Customers
R = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
Rcount = Range("A6").SpecialCells(xlCellTypeVisible).End(xlDown).Row ' Last Cell Number on filtered Cells

With Sheets("Payment Upload").AutoFilter.Range
    'Copying Customer Name and Number
    APCIFile.Activate
    Range("A" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    Range(ActiveCell.Address & ":B" & Rcount).Copy
    'Pasting Customers Name and Number PPA file
    PPAfile.Activate
    Range("B" & PPALastRow).PasteSpecial xlPasteValues
    'Copying Customer NCPDCP, Chain and DC Number
    APCIFile.Activate
    Range("D" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    Range(ActiveCell.Address & ":F" & Rcount).Copy
    'Pasting Copying Customer NCPDCP, Chain and DC Number PPA file
    PPAfile.Activate
    Range("D" & PPALastRow).PasteSpecial xlPasteValues
    'Copying Dates from APCI Payment File
    APCIFile.Activate
    Range("L" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    Range(ActiveCell.Address & ":M" & Rcount).Copy
    'Pasting Dates in PPA File
    PPAfile.Activate
    Range("H" & PPALastRow).PasteSpecial xlPasteValues
    'Copying Compliant Column Data
    APCIFile.Activate
    Range("N" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    Range(ActiveCell.Address & ":N" & Rcount).Copy
    'Pasting Compliance data in PPA File
    PPAfile.Activate
    Range("N" & PPALastRow).PasteSpecial xlPasteValues
    'Copying NTE Amount
    APCIFile.Activate
    Range("S" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    Range(ActiveCell.Address & ":S" & Rcount).Copy
    'Pasting NTE in PPA File
    PPAfile.Activate
    Range("T" & PPALastRow).PasteSpecial xlPasteValues
    'Copying System Cost Amount
    APCIFile.Activate
    Range("X" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    Range(ActiveCell.Address & ":X" & Rcount).Copy
    'Pasting System Cost in PPA File
    PPAfile.Activate
    Range("U" & PPALastRow).PasteSpecial xlPasteValues
    'Copying BPR,GPR, GCR and HM
    APCIFile.Activate
    Range("Y" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    Range(ActiveCell.Address & ":AB" & Rcount).Copy
    'Pasting BPR,GPR, GCR and HM
    PPAfile.Activate
    Range("Y" & PPALastRow).PasteSpecial xlPasteValues
    'Copying OS Sales and NP
    APCIFile.Activate
    Range("T" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    Range(ActiveCell.Address & ":U" & Rcount).Copy
    'Pasting OS Sales and NP
    PPAfile.Activate
    Range("V" & PPALastRow).PasteSpecial xlPasteValues
    'Copying District
    APCIFile.Activate
    Range("Q" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    Range(ActiveCell.Address & ":Q" & Rcount).Copy
    'Pasting District
    PPAfile.Activate
    Range("AC" & PPALastRow).PasteSpecial xlPasteValues
    'Copying CarryOver Cost
    APCIFile.Activate
    Range("AE" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    Range(ActiveCell.Address & ":AE" & Rcount).Copy
    'Pasting CarryOver Cost
    PPAfile.Activate
    Range("AD" & PPALastRow).PasteSpecial xlPasteValues
    
End With

'Closing APCI Payment File
APCIFile.Activate
ActiveSheet.Range("A5:AE5").AutoFilter
Application.DisplayAlerts = False
APCIFile.Close
Application.DisplayAlerts = True

'Copying Previous Row Format on BPR GPR GCR HM on New Data
PPAfile.Activate
PPALastRowMinus1 = PPALastRow - 1
NewLastRow = Range("B6").End(xlDown).Row
Application.CutCopyMode = False
Range("B" & PPALastRowMinus1 & ":AD" & PPALastRowMinus1).Copy
Range("B" & PPALastRowMinus1 & ":AD" & NewLastRow).PasteSpecial xlPasteFormats
Application.CutCopyMode = False

'Setting up Dates
Range("I" & PPALastRowMinus1).Select
ActiveCell.Copy
ActiveCell.Offset(1, -1).PasteSpecial xlPasteAll
ActiveCell.Offset(0, 1).Select
LeftCellAdd = ActiveCell.Offset(0, -1).Address
ActiveCell.value = "=" & LeftCellAdd & "+31"
Range("H" & PPALastRow & ":I" & PPALastRow).Copy
Range("H" & PPALastRow & ":I" & NewLastRow).PasteSpecial xlPasteAll

ActiveSheet.Range("A6:AD6").AutoFilter Field:=10, Criteria1:="<>" & ""
Rcount = Range("A6").SpecialCells(xlCellTypeVisible).End(xlDown).Row
With ActiveSheet.AutoFilter.Range
    Range("J" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    Range(ActiveCell.Address & ":J" & Rcount).Select
    Selection.Clear
End With

ActiveSheet.Range("A6:AD6").AutoFilter

'Waivers Setup
'Calling Waivers Function to Create Waivers File
Call WaiversSetup

Dim WaiverFile As Workbook
Set WaiverFile = Workbooks.Open("C:\Users\eo5v4x3\Downloads\Waivers.xls")
WaiverFile.Activate
Sheets("Sheet2").Activate

WLastRow = Range("A1").End(xlDown).Row

For i = 2 To WLastRow
    
    WaiverFile.Activate
    Sheets("Sheet2").Activate
    
    Customer = Range("A" & i).value
    EffDate = Range("c" & i).value
    Waiver = Range("b" & i).value
    
    PPAfile.Activate
    Lrow = Range("B6").End(xlDown).Row
    ActiveSheet.Range("A6:AD6").AutoFilter Field:=3, Criteria1:=Customer
    ActiveSheet.Range("A6:AD6").AutoFilter Field:=8, Operator:= _
        xlFilterValues, Criteria2:=Array(1, EffDate)
    VCount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    If VCount >= 1 Then
        Rcount = Range("A6").SpecialCells(xlCellTypeVisible).End(xlDown).Row
        With ActiveSheet.AutoFilter.Range
            Range("A" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
            RebateAmt = ActiveCell.Offset(0, 6).value
            If ActiveCell.value <> Date And RebateAmt > 0 Then
                'Already Paid
            ElseIf Len(ActiveCell.value) = 0 Or ActiveCell.value = Date Then
                ActiveCell.value = Date
                ExistingValue = ActiveCell.Offset(0, 9).value
                If Len(ExistingValue) > 0 Then
                    ActiveCell.Offset(0, 9).value = ExistingValue & ", " & Waiver
                Else
                    ActiveCell.Offset(0, 9).value = Waiver
                End If
            End If
        End With
    End If
    ActiveSheet.Range("A6:AD6").AutoFilter
Next i

WaiverFile.Activate
Application.DisplayAlerts = False
Sheets("Sheet2").Delete
Sheets("Sheet1").Delete
Application.DisplayAlerts = True
WaiverFile.Save
WaiverFile.Close

End Sub
Function WaiversSetup()

Dim WaiversFile As Workbook
WaiversFilePath = "C:\Users\eo5v4x3\Downloads\Waivers.xls"
Set WaiversFile = Workbooks.Open(WaiversFilePath)
Sheets(1).Activate
ActiveSheet.Range("A1:L1").AutoFilter
Sheets.Add(After:=Sheets(1)).Name = "Sheet1"
Sheets("Sheet1").Activate
Range("A1").value = "Account Number"
Range("B1").value = "Waiver Type"
Range("C1").value = "Effective Date"
Range("D1").value = "Expiration Date"
Range("A1:D1").Font.Bold = True
Sheets(1).Activate
ActiveSheet.Range("A1:L1").AutoFilter Field:=2, Criteria1:="Accepted"
ActiveSheet.Range("A1:L1").AutoFilter Field:=3, Criteria1:="Waiver Approved"

R = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
Rcount = Range("F6").SpecialCells(xlCellTypeVisible).End(xlDown).Row ' Last Cell Number on filtered Cells

With Sheets(1).AutoFilter.Range
    'Copying Customer Number
    Range("F" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    Range(ActiveCell.Address & ":F" & Rcount).Copy
    'Pasting Customers Name
    Sheets("Sheet1").Activate
    Range("A2").PasteSpecial xlPasteAll
    'Copying Waivers and Dates
    Sheets(1).Activate
    Range("J" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    Range(ActiveCell.Address & ":L" & Rcount).Copy
    'Pasting Customers Name and Number PPA file
    Sheets("Sheet1").Activate
    Range("B2").PasteSpecial xlPasteAll
    
End With

'Changing CUstomer Number formats to Geneera (Numbers format)
Lrow = Range("A1").End(xlDown).Row
With Range("A2:A" & Lrow)
    .NumberFormat = "General"
    .value = .value
End With

'Removing Duplicate Values
Columns("A:D").Select
ActiveSheet.Range("$A$1:$D$" & Lrow).RemoveDuplicates Columns:=Array(1, 2, 3, 4), _
        Header:=xlYes
        
'Autofiting all data
ActiveSheet.UsedRange.EntireColumn.AutoFit

'Removing Values More than Two Years
Dim currentDate As Date
Dim lastDay As Date

' Get the current date
currentDate = Date

' Calculate the first day of the current month
Dim firstDayOfCurrentMonth As Date
firstDayOfCurrentMonth = DateSerial(Year(currentDate), Month(currentDate), 1)

' Calculate the last day of the second-to-last month
lastDay = firstDayOfCurrentMonth - 1

' Calculate the first day of the second-to-last month
Dim firstDayOfLast2ndMonth As Date
firstDayOfLast2ndMonth = DateSerial(Year(lastDay), Month(lastDay) - 1, 1)

' Calculate the last day of the second-to-last month
LastDayOfLast2ndMonth = firstDayOfLast2ndMonth - 0
LWDate = DateSerial(Year(firstDayOfLast2ndMonth), Month(firstDayOfLast2ndMonth) + 1, 0)


JDate = "01/01/"
YearDate = DateAdd("yyyy", -2, Date)
YDate = Format(YearDate, "yyyy")
JanDateofLast2Years = JDate & YDate


    ActiveSheet.Range("$A$1:$D$833").AutoFilter Field:=3, Criteria1:= _
        ">=" & JanDateofLast2Years, Operator:=xlAnd, Criteria2:="<=" & LWDate

Sheets.Add(After:=Sheets("Sheet1")).Name = "Sheet2"
Sheets("Sheet1").Activate
        
RnCount = Range("F6").SpecialCells(xlCellTypeVisible).End(xlDown).Row ' Last Cell Number on filtered Cells

With Sheets(1).AutoFilter.Range
    'Copying Customer Number
    Range("A1").Select
    Range(ActiveCell.Address & ":D" & RnCount).Copy
    Sheets("Sheet2").Activate
    Range("A1").PasteSpecial xlPasteAll
    ActiveSheet.UsedRange.EntireColumn.AutoFit
    
End With
        
'VlookUp from PPAFile on Customers
Sheets("Sheet2").Activate
LMRow = Range("A1").End(xlDown).Row
Range("E2").Select
Range("E2").value = "=VLOOKUP(A2,'[APCI New Non Compliant TR (Working File)_New.xlsx]APCI New '!$C1:$C$65536,1,0)"
Range("E2").Copy
Range("E2:E" & LMRow).PasteSpecial xlPasteAll
Range("E2:E" & LMRow).Copy
Range("E2:E" & LMRow).PasteSpecial xlPasteValues

ActiveSheet.Range("A1:E1").AutoFilter Field:=5, Criteria1:="#N/A"
RnCount = Range("A1").SpecialCells(xlCellTypeVisible).End(xlDown).Row

With ActiveSheet.AutoFilter.Range
    Range("A" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    Range(ActiveCell.Address & ":A" & RnCount).EntireRow.Delete
End With
ActiveSheet.Range("A1:D1").AutoFilter
Range("E:E").Clear

WaiversFile.Save
WaiversFile.Close

End Function
