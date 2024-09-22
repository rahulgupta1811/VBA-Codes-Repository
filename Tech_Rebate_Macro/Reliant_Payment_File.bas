Attribute VB_Name = "Reliant"
'Global Variables
Public CostFile As String
Public BWFile As String
Public ReliantFile As Workbook
Public RenamedFile As String
Sub ReliantPaymentFile()
Dim SourceFile As String
Dim DestinationFile As String

'Copying Relaint Previous Month Payment File
Tday = DateAdd("M", -1, Date)
CurrentYear = Format(Tday, "YYYY")
CurrentYear2 = Right(CurrentYear, 2)

LMonthYear3 = DateAdd("M", -2, Date)
CurrentYear3 = Format(LMonthYear3, "YYYY")
CurrentYear3 = Right(CurrentYear3, 2)

CurrentMonth = DateAdd("M", -1, Date)
lastmonth = Format(CurrentMonth, "mmmm")
LastMonthNum = Format(CurrentMonth, "MM")
LastMonthFolder = LastMonthNum & " " & lastmonth & "'" & CurrentYear3

CurrentMonth2 = DateAdd("M", -2, Date)
lastmonth2 = Format(CurrentMonth2, "mmmm")
LMonth = Format(CurrentMonth2, "MM")
lastmonth2 = Left(lastmonth2, 3)
LastMonthNum2 = Month(CurrentMonth2)
LastMonthFolder2 = " (" & lastmonth2 & "'" & CurrentYear3 & " " & "Rbts)"

WorkFileName = "20" & CurrentYear3 & LMonth

Dim CostFile As Object
Set CostFile = CreateObject("Scripting.Filesystemobject")
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False
CostFileDestinations = "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\Payment Files\Reliant\"
SourceFileForReliant = "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Payment Files\" & CurrentYear & "\" & LastMonthFolder & LastMonthFolder2 & "\Reliant\Reliant Tech Rebate Payment - " & WorkFileName & ".xlsx"
CostFile.CopyFile SourceFileForReliant, CostFileDestinations, True

'Update BW Data into Working File
Lmo = DateAdd("M", -2, Date)
LMon = Format(Lmo, "MM")
Yr = Format(Lmo, "YYYY")
Dim BWFile As Workbook
Dim ReliantFile As Workbook
Set ReliantFile = Workbooks.Open("\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\Payment Files\Reliant\Reliant Tech Rebate Payment - " & Yr & LMon & ".xlsx")

'Clearing BW File Data from Working File
ReliantFile.Activate
Sheets("BW-Compliance Data").Activate
LastCell = Range("A2").End(xlDown).Row
Range("A2:DH" & LastCell).Clear

'Putting Data from BW to WF
Set BWFile = Workbooks.Open("\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\BW Queries\Reliant.xlsx")
Sheets("Table").Activate
LastCell = Range("G16").End(xlDown).Row
Range("G16:DN" & LastCell).Copy
ReliantFile.Activate
Sheets("BW-Compliance Data").Activate
Range("A2").PasteSpecial xlPasteAll
Application.CutCopyMode = False

LastRowin = Range("D2").End(xlDown).Row
With Range("D2:D" & LastRowin)
    .NumberFormat = "General"
    .value = .value
End With
Range("A1").Select
' Calling function to sort data by Totoal Purchases in Decensding order
Call APSC.Sorting
BWFile.Close

'Setting Up CarryOver Cost
ReliantFile.Activate
Sheets("Carryover Cost").Activate

LastColumnCell = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Address
Range(LastColumnCell).Select
ActiveCell.Offset(0, -3).Select
CopyfromCol = ActiveCell.Address

Range(CopyfromCol & ":" & LastColumnCell).Copy
Range(LastColumnCell).Offset(0, 1).Select
ActiveCell.PasteSpecial xlPasteAll

'Updating Headers of New CarryOver Columns
LastReb = Range(LastColumnCell).Offset(0, 1).value
UptdCarry = Range(LastColumnCell).Offset(0, 2).value
CostFrCurr = Range(LastColumnCell).Offset(0, 3).value
CurrCarryCost = Range(LastColumnCell).Offset(0, 4).value

'Putting Data from Validation to CarryOver Sheet
Sheets("Validation").Activate
Range("K4:K6").Copy

Sheets("Carryover Cost").Activate
Range(LastColumnCell).Offset(1, 3).PasteSpecial xlPasteValues

'Last Payment
Sheets("Validation").Activate
Range("P4:P6").Copy

Sheets("Carryover Cost").Activate
Range(LastColumnCell).Offset(1, 1).PasteSpecial xlPasteValues

'Copying Carryiver Fomula
Sheets("Carryover Cost").Activate
LastCell = Range(LastColumnCell).End(xlDown).Address
Range(LastColumnCell).Offset(1, 0).Select
Range(ActiveCell.Address & ":" & LastCell).Copy
Range(LastColumnCell).Offset(1, 4).PasteSpecial xlPasteFormulas

Range(LastColumnCell).Offset(1, -2).Select
LastCell = Range(ActiveCell.Address).End(xlDown).Address
Range(ActiveCell.Address & ":" & LastCell).Copy
Range(ActiveCell.Address).Offset(0, 4).PasteSpecial xlPasteFormulas

'Copying Format
Lastsel = Range(LastColumnCell).End(xlDown).Address
StarCell = Range(LastColumnCell).Offset(0, -3).Address

Range(Lastsel & ":" & StarCell).Copy
Range(LastColumnCell).Offset(0, 1).PasteSpecial xlPasteFormats

ActiveSheet.UsedRange.EntireColumn.AutoFit
Application.CutCopyMode = False

'Change Carryover Header for Current Month
Dim CurrYr As String
LMonth = DateAdd("M", -2, Date)
CurrMo = DateAdd("M", -1, Date)
CurMo = Format(CurrMo, "mmmm")
CurrYr = Right(CurrMo, 2)
Lmonthe = Format(LMonth, "mmmm")
Range(LastColumnCell).Offset(0, 1).value = "Rebate Paid in " & Lmonthe
Range(LastColumnCell).Offset(0, 3).value = "Cost for " & CurMo & "'" & CurrYr & " Eval Period"
Range(LastColumnCell).Offset(0, 4).value = "Carry Over Cost for " & CurMo & "'" & CurrYr & " Eval Period"
Range(LastColumnCell).Select

'Evaluations ---->
'Rebate Formula Setup
Sheets("Validation").Activate
Range("P4").value = "=IF(M4>=1000,N4,M4)"
EndCell = Range("P4").End(xlDown).Row
Range("P4").Copy
Range("P4:P" & EndCell).PasteSpecial xlPasteAll

'Getting New Carryover cost
Sheets("Carryover Cost").Activate
LastColumnRow = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
LastColumnRow = LastColumnRow - 1
LastColumn = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Address

Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
Regex.Global = True
Regex.Pattern = "[0-9]"
OutStr = Regex.Replace(LastColumn, "")
OutStr = Replace(OutStr, "$", "")

Sheets("Validation").Activate
Range("O4").value = "=VLOOKUP(B4,'Carryover Cost'!B:" & OutStr & "," & LastColumnRow & ",0)"
EndCell = Range("O4").End(xlDown).Row

Range("O4").Copy
Range("O4:O" & EndCell).PasteSpecial xlPasteAll
Range("O4:O" & EndCell).Copy
Range("O4:O" & EndCell).PasteSpecial xlPasteValues
Application.CutCopyMode = False


'Getting New Purchase amount From BW sheet
Sheets("Validation").Activate
Range("L4").value = "=VLOOKUP(B4,'BW-Compliance Data'!D:BF,55,0)"
EndCell = Range("L4").End(xlDown).Row

Range("L4").Copy
Range("L4:L" & EndCell).PasteSpecial xlPasteAll
Range("L4:L" & EndCell).Copy
Range("L4:L" & EndCell).PasteSpecial xlPasteValues
Application.CutCopyMode = False

'Setting up Dates
Sheets("Validation").Activate
EndCell = Range("H4").End(xlDown).Row
Range("H4").Copy
Range("G4:G" & EndCell).PasteSpecial xlPasteAll

CurrYear = Year(Date)
Currmonth = Format(DateAdd("M", 0, Now), "MM")
CurrentMonth = CurrYear & Currmonth
Range("H4:H" & EndCell).value = CurrentMonth

'Setting Comments
For i = 4 To 6
    Range("M" & i).Select
    NP = ActiveCell.value
    cost = ActiveCell.Offset(0, -2).value
    Rebate = ActiveCell.Offset(0, 3).value
    
    If Rebate < 1000 Then
        ActiveCell.Offset(0, 4).value = "Paid on NP"
    ElseIf cost < 1000 And NP >= 1000 And Rebate >= 1000 Then
        ActiveCell.Offset(0, 4).value = "Paid on NTE Using carry over Cost"
    Else
        ActiveCell.Offset(0, 4).value = "Paid on NTE"
    End If
    
Next i
    
'Moving Data to Final Sheet
Sheets("Validation").Activate
Range("A4:C" & EndCell).Copy
Sheets("Final List").Activate
Range("A6").PasteSpecial xlPasteValues

'Dates
Sheets("Validation").Activate
Range("G4:H" & EndCell).Copy
Sheets("Final List").Activate
Range("D6").PasteSpecial xlPasteValues

'Rebate Amount
Sheets("Validation").Activate
Range("P4:P" & EndCell).Copy
Sheets("Final List").Activate
Range("F6").PasteSpecial xlPasteValues

Range("A2").value = CurrMo
Range("A1").Select

ReliantFile.Save
ReliantFile.Close

Mr = Format(CurrMo, "MM")
NYyear = DateAdd("M", -1, Date)
NY = Format(NYyear, "YYYY")

ExistFile = "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\Payment Files\Reliant\Reliant Tech Rebate Payment - " & Yr & LMon & ".xlsx"

Name ExistFile As _
   "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\Payment Files\Reliant\Reliant Tech Rebate Payment - " & NY & Mr & ".xlsx"


MsgBox "Completed", vbInformation, "Success"

End Sub
