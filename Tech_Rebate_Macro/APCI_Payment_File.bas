Attribute VB_Name = "APCI"
Public APSCworkingFile As Workbook
Sub APCIPaymentFile()
'Copying APCI File

LMonthYear = DateAdd("M", -1, Date)
CurrentYear = Format(LMonthYear, "YYYY")
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
CostFileDestination = "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\Payment Files\APCI\"
SourceFileForIPC = "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Payment Files\" & CurrentYear & "\" & LastMonthFolder & LastMonthFolder2 & "\APCI\APCI Tech Payment_" & WorkFileName & " Working File.xlsx"
CostFile.CopyFile SourceFileForIPC, CostFileDestination, True

'BW Data
Dim WorkingFile As String
WorkingFile = "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\Payment Files\APCI\APCI Tech Payment_" & WorkFileName & " Working File.xlsx"
Set APCIWorkingfile = Workbooks.Open(WorkingFile)
APCIWorkingfile.Activate

'Clearing BW File Data from Working File
Sheets("BW-Compliance Data").Activate
ActiveSheet.Range("A1:DH1").AutoFilter
LastCell = Range("A2").End(xlDown).Row
Range("A2:DH" & LastCell).Clear

'Getting BW data from BW File
Dim BWFile As Workbook
BWFileLoc = "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\BW Queries\APCI.xlsx"
Set BWFile = Workbooks.Open(BWFileLoc)
BWFile.Activate
LastCell = Range("G16").End(xlDown).Row
Range("G16:DN" & LastCell).Copy
APCIWorkingfile.Activate
Sheets("BW-Compliance Data").Activate
Range("A2").PasteSpecial xlPasteAll
Application.CutCopyMode = False
Range("A1").Select
LastRowin = Range("D2").End(xlDown).Row
With Range("D2:D" & LastRowin)
    .NumberFormat = "General"
    .value = .value
End With
' Calling function to sort data by Totoal Purchases in Decensding order
Call APSC.Sorting
BWFile.Close

'Setting CarryOver Cost
APCIWorkingfile.Activate
Sheets("Carryover").Activate

LastColumn = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Address
Range(LastColumn).Select
NewLastCol = ActiveCell.Offset(0, -2).Address
Range(LastColumn & ":" & NewLastCol).Copy
Range(LastColumn).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.PasteSpecial xlPasteAll
lastRow = Range("A2").End(xlDown).Row
    
'Sheets(1).Activate
'LastCell = Range("K6").End(xlDown).Row
'LastCell = LastCell - 1
'Range("K6:K" & LastCell).Copy
Sheets("Carryover").Activate

Range(LastColumn).Select
Lrowadd = ActiveCell.End(xlDown).Offset(0, 1).Address
NewLastCol = ActiveCell.Offset(1, 1).Address
Range(NewLastCol).value = "=VLOOKUP(A2,'Payment Upload'!B:K,10,0)"
Range(NewLastCol).Copy
Range(NewLastCol & ":" & Lrowadd).PasteSpecial xlPasteAll
Range(NewLastCol & ":" & Lrowadd).Copy
Range(NewLastCol & ":" & Lrowadd).PasteSpecial xlPasteValues

'Copying Formula of CarryOver Cost From Preivous Month val
Sheets("Carryover").Activate
NLastCell = Range("A3").End(xlDown).Row

Range(LastColumn).Select
NewLastCol = ActiveCell.Offset(1, 0).Address
Range(NewLastCol).Copy
Range(LastColumn).Select
NewLastCol2 = ActiveCell.Offset(1, 3).Address
Range(NewLastCol2).PasteSpecial xlPasteAll

Dim ComRow As String
ComRow = Range("A3").End(xlDown).Row
CarryLastColumn = Range(LastColumn).Offset(ComRow - 2, 3).Address

Range(NewLastCol2).Copy
Range(NewLastCol2 & ":" & CarryLastColumn).PasteSpecial xlPasteAll

'Setting headers of carryover
Last2dmonth = DateAdd("M", -2, Date)
Last2month = Format(Last2dmonth, "mmm")

lastdmonth = DateAdd("M", -1, Date)
lastmonth = Format(lastdmonth, "mmm")
yr1 = Format(Last2dmonth, "YYYY")
yr1 = Right(yr1, 2)
yr2 = Format(lastdmonth, "YYYY")
yr2 = Right(yr2, 2)
Range(LastColumn).Offset(0, 1).value = "FINAL REBATE PAID-" & Last2month & "'" & yr1
Range(LastColumn).Offset(0, 2).value = "Cost of " & lastmonth & "'" & yr2
Range(LastColumn).Offset(0, 3).value = "Carry-over cost-" & lastmonth & "'" & yr2
Application.CutCopyMode = False

ActiveSheet.UsedRange.EntireColumn.AutoFit

'Setting up Payment sheet
Sheets("Payment Upload").Activate
Lcell = Range("A5").End(xlDown).Row
Range("I6:I" & Lcell).value = ""
Range("L6:N" & Lcell).value = ""
Range("O6:O" & Lcell).value = ""
Range("T6:T" & Lcell).value = ""
Range("Y6:AB" & Lcell).value = ""
Range("AE6:AE" & Lcell).value = ""
Range("AC6:AC" & Lcell).Copy
Range("AD6").PasteSpecial xlPasteValues
Range("AC6:AC" & Lcell).value = ""
Application.CutCopyMode = False

lastdmonth = DateAdd("M", -1, Date)
lastmonth = Format(lastdmonth, "YYYYmm")

Range("A3").value = lastmonth

'Setting  Rebate Month and Paid Month
Range("L6:L" & Lcell).value = lastmonth
lastdmonth = DateAdd("M", 0, Date)
lastmonth = Format(lastdmonth, "YYYYmm")
Range("M6:M" & Lcell).value = lastmonth

'Preparing System Cost
Sheets("Payment Upload").Activate

For n = 1 To 3
    Columns(25).EntireColumn.Insert
Next n

Range("Y5").value = "MPS"
Range("Z5").value = "Parata"
Range("AA5").value = "PW"

'Getting Cost from Cost File
CurrYear = Year(Date)
Currmonth = Format(DateAdd("M", -1, Now), "MM") 'Change -1 to 0
CurrentMonth = CurrentYear & Currmonth

Dim CstFile As Workbook
Set CstFile = Workbooks.Open("\\ddcf2015\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\System Cost\CostFiles_Template\Cost File Template_ " & CurrentMonth & ".xlsx")

CstFile.Activate
Sheets("Sheet1").Activate
APCIWorkingfile.Activate
Sheets("Payment Upload").Activate
Range("Y6").value = "=VLOOKUP(D6,'[Cost File Template_ " & CurrentMonth & ".xlsx]Sheet1'!$A:$B,2,0)"
Range("Z6").value = "=VLOOKUP(B6,'[Cost File Template_ " & CurrentMonth & ".xlsx]Parata '!$B:$C,2,0)"
Range("AA6").value = "=VLOOKUP(B6,'[Cost File Template_ " & CurrentMonth & ".xlsx]Prescribed Wellness '!$B:$C,2,0)"

LastCell = Range("A5").End(xlDown).Row
Range("Y6:AA6").Copy
Range("Y6:AA" & LastCell).PasteSpecial xlPasteAll
Range("Y6:AA" & LastCell).Copy


CstFile.Close

'Replacing N/A Value to 0
Range("Y6:AA" & LastCell).Copy
Range("Y6:AA" & LastCell).PasteSpecial xlPasteValues
Range("Y6:AA" & LastCell).Replace What:="#N/A", Replacement:="0", MatchCase:=True

Dim F As Integer
Dim Cellval As Integer
Cellval = 6
Dim NewCost
lastRow = LastCell - 2
For F = 6 To lastRow
    MPS = Range("Y" & Cellval).value
    Parata = Range("Z" & Cellval).value
    PW = Range("AA" & Cellval).value
    NewCost = MPS + Parata + PW
    Range("X" & Cellval).value = NewCost
    Cellval = Cellval + 1
Next F

'Deleting Extra Created Columns
For n = 1 To 3
    Columns(25).EntireColumn.Delete
Next n

Sheets("Carryover").Activate

Lcell = Range(LastColumn).End(xlDown).Offset(0, 2).Address
Range(LastColumn).Offset(1, 2).value = "=VLOOKUP(A2,'Payment Upload'!B:X,23,0)"
Range(LastColumn).Offset(1, 2).Select
ActiveCell.Copy
FCell = ActiveCell.Address

Range(FCell & ":" & Lcell).PasteSpecial xlPasteAll
Range(FCell & ":" & Lcell).Copy
Range(FCell & ":" & Lcell).PasteSpecial xlPasteValues
ForCop1 = Range(LastColumn).Offset(1, 0).Address
Forcop2 = Range(LastColumn).Offset(1, -2).Address

Fcol = Range(ForCop1 & ":" & Forcop2).Copy
Range(LastColumn).Offset(1, 1).PasteSpecial xlPasteFormats
ForCop1 = Range(LastColumn).Offset(1, 1).Address
Forcop2 = Range(LastColumn).Offset(1, 2).Address
Range(ForCop1 & ":" & Forcop2).Copy
NCell = Range(FCell).Offset(0, -1).Address
NCell2 = Range(FCell).Offset(0, 1).End(xlDown).Address
Range(NCell & ":" & NCell2).PasteSpecial xlPasteFormats
Range(NCell & ":" & NCell2).Replace What:="#N/A", Replacement:="0", MatchCase:=True
ActiveSheet.UsedRange.EntireColumn.AutoFit

'Getting Fresh Data on Payment Upload Sheet
Sheets("Payment Upload").Activate

'Getting Data from BW Sheet
Sheets("BW-Compliance Data").Activate
LastRowin = Range("D2").End(xlDown).Row
With Range("D2:D" & LastRowin)
    .NumberFormat = "General"
    .value = .value
End With

Sheets("Payment Upload").Activate
lastRow = Range("A5").End(xlDown).Row

'Calling Waivers Function
Call WaiversSetup

Sheets("BW-Compliance Data").Activate
ActiveSheet.Range("A1:DH1").AutoFilter

Sheets("Payment Upload").Activate
Range("O6").value = "=VLOOKUP(B6,Sheet1!A:B,2,0)"
Range("O6").Copy
Range("O6:O" & lastRow).PasteSpecial xlPasteAll
Range("O6:O" & lastRow).Copy
Range("O6:O" & lastRow).PasteSpecial xlPasteValues
Sheets("Sheet1").Delete
Sheets("Payment Upload").Activate
Range("O6:O" & lastRow).Replace What:="#N/A", Replacement:="#"

Range("Y6").value = "=CONCATENATE(VLOOKUP(B6,'BW-Compliance Data'!D:XFD,48,0),$AB$1)"
Range("Y6").Copy
Range("Y6:Y" & lastRow).PasteSpecial xlPasteAll
Range("Y6:Y" & lastRow).Copy
Range("Y6:Y" & lastRow).PasteSpecial xlPasteValues

Range("Z6").value = "=CONCATENATE(VLOOKUP(B6,'BW-Compliance Data'!D:XFD,52,0),$AB$1)"
Range("Z6").Copy
Range("Z6:Z" & lastRow).PasteSpecial xlPasteAll
Range("Z6:Z" & lastRow).Copy
Range("Z6:Z" & lastRow).PasteSpecial xlPasteValues

Range("AA6").value = "=CONCATENATE(VLOOKUP(B6,'BW-Compliance Data'!D:XFD,49,0),$AB$1)"
Range("AA6").Copy
Range("AA6:AA" & lastRow).PasteSpecial xlPasteAll
Range("AA6:AA" & lastRow).Copy
Range("AA6:AA" & lastRow).PasteSpecial xlPasteValues


Range("AB6").value = "=VLOOKUP(B6,'BW-Compliance Data'!D:XFD,15,0)"
Range("AB6").Copy
Range("AB6:AB" & lastRow).PasteSpecial xlPasteAll
Range("AB6:AB" & lastRow).Copy
Range("AB6:AB" & lastRow).PasteSpecial xlPasteValues


Range("T6").value = "=VLOOKUP(B6,'BW-Compliance Data'!D:XFD,55,0)"
Range("T6").Copy
Range("T6:T" & lastRow).PasteSpecial xlPasteAll
Range("T6:T" & lastRow).Copy
Range("T6:T" & lastRow).PasteSpecial xlPasteValues

'Removing N/A
Range("T6:T" & lastRow).Replace What:="#N/A", Replacement:="0"
Range("Y6:AA" & lastRow).Replace What:="#N/A", Replacement:="0"
Range("AB6:AB" & lastRow).Replace What:="#N/A", Replacement:=""

For i = 25 To 27
    With ActiveSheet.Columns(i)
        .NumberFormat = "0"
        .value = .value
    End With
Next i

Range("Y6:AA" & lastRow).Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"

'Get Carryover Cost from CarryoverSheet into Payment Upload

Sheets("Carryover").Activate
LastCol = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
Sheets("Payment Upload").Activate
lastRow = Range("A5").End(xlDown).Row
Range("AE6").value = "=VLOOKUP(B6,'Carryover'!A:XFD," & LastCol & ",0)"
Range("AE6").Copy
Range("AE6:AE" & lastRow).PasteSpecial xlPasteAll
Range("AE6:AE" & lastRow).Copy
Range("AE6:AE" & lastRow).PasteSpecial xlPasteValues
Range("AE6:AE" & lastRow).Replace What:="#N/A", Replacement:="0"
Range("O6:O" & lastRow).Replace What:="#N/A", Replacement:="#"
'Evaluation Start Here---->
'Compliance Check
'District 3 and 7
' Filtering District 3 and 7
Sheets("Payment Upload").Activate
LastCell = Range("A5").End(xlDown).Row
For i = 6 To LastCell
    BPR = Range("Y" & i).value
    GCR = Range("Z" & i).value
    GPR = Range("AA" & i).value
    HM = Range("AB" & i).value
    Dist = Range("Q" & i).value
    Chain = Range("E" & i).value
    
    If VarType(Chain) <> vbInteger Then
        Chain = 0
    End If
    
    
    'Dist 3 and 7
    If Dist = 701100 Or Dist = 301100 Or Dist = 311100 Then
        If GCR >= 0.24 And HM = "Y" Then
            'Range("AC" & i) = "Compliant"
            Range("I" & i) = "0.00"
            Range("N" & i) = "Y"
        ElseIf GCR < 0.24 And HM = "Y" Then
            Range("AC" & i) = "Non Compliant. Missing GCR"
            Range("I" & i) = "0.00"
            Range("N" & i) = "N"
        ElseIf GCR >= 0.24 And HM = "N" Or HM = "" Then
            Range("AC" & i) = "Non Compliant. Missing HM"
            Range("I" & i) = "0.00"
            Range("N" & i) = "N"
        ElseIf GCR < 0.24 And HM = "N" Or HM = "" Then
            Range("AC" & i) = "Non Compliant. Missing GCR and HM"
            Range("I" & i) = "0.00"
            Range("N" & i) = "N"
        ElseIf GCR = 0 And BPR = 0 And GPR = 0 And HM = "N" Or HM = "" Then
            Range("AC" & i) = "No Data on BW"
            Range("I" & i) = "0.00"
            Range("N" & i) = "N"
        End If
        
    'District 1
    Else
        If Left(Chain, Len("4")) = "4" Then
            If BPR >= 0.85 And GPR >= 0.9 And HM = "Y" Then
                Range("N" & i) = "Y"
            ElseIf BPR < 0.85 And GPR >= 0.9 And HM = "Y" Then
                Range("N" & i) = "N"
                Range("I" & i) = "0.00"
                Range("AC" & i) = "Non Compliant. Missing BPR"
            ElseIf BPR >= 0.85 And GPR < 0.9 And HM = "Y" Then
                Range("N" & i) = "N"
                Range("I" & i) = "0.00"
                Range("AC" & i) = "Non Compliant. Missing GPR"
            ElseIf BPR < 0.85 And GPR < 0.9 And HM = "Y" Then
                Range("N" & i) = "N"
                Range("I" & i) = "0.00"
                Range("AC" & i) = "Non Compliant. Missing BPR and GPR"
            ElseIf BPR < 0.85 And GPR < 0.9 And HM = "N" Then
                Range("N" & i) = "N"
                Range("I" & i) = "0.00"
                Range("AC" & i) = "Non Compliant. Missing BPR, GPR and HM"
            ElseIf BPR < 0.85 And GPR >= 0.9 And HM = "N" Then
                Range("N" & i) = "N"
                Range("I" & i) = "0.00"
                Range("AC" & i) = "Non Compliant. Missing BPR and HM"
            ElseIf BPR >= 0.85 And GPR < 0.9 And HM = "N" Then
                Range("N" & i) = "N"
                Range("I" & i) = "0.00"
                Range("AC" & i) = "Non Compliant. Missing GPR and HM"
            ElseIf BPR >= 0.85 And GPR >= 0.9 And HM = "N" Then
                Range("N" & i) = "N"
                Range("I" & i) = "0.00"
                Range("AC" & i) = "Non Compliant. Missing HM"
            ElseIf GCR = 0 And BPR = 0 And GPR = 0 And HM = "N" Or HM = "" Then
                Range("AC" & i) = "No Data on BW"
                Range("I" & i) = "0.00"
                Range("N" & i) = "N"
            End If
        
            
        Else
            If BPR >= 0.9 And GPR >= 0.9 And HM = "Y" Then
                Range("N" & i) = "Y"
            ElseIf BPR < 0.9 And GPR >= 0.9 And HM = "Y" Then
                Range("N" & i) = "N"
                Range("I" & i) = "0.00"
                Range("AC" & i) = "Non Compliant. Missing BPR"
            ElseIf BPR >= 0.9 And GPR < 0.9 And HM = "Y" Then
                Range("N" & i) = "N"
                Range("I" & i) = "0.00"
                Range("AC" & i) = "Non Compliant. Missing GPR"
            ElseIf BPR < 0.9 And GPR < 0.9 And HM = "Y" Then
                Range("N" & i) = "N"
                Range("I" & i) = "0.00"
                Range("AC" & i) = "Non Compliant. Missing BPR and GPR"
            ElseIf BPR < 0.9 And GPR < 0.9 And HM = "N" Then
                Range("N" & i) = "N"
                Range("I" & i) = "0.00"
                Range("AC" & i) = "Non Compliant. Missing BPR, GPR and HM"
            ElseIf BPR < 0.9 And GPR >= 0.9 And HM = "N" Then
                Range("N" & i) = "N"
                Range("I" & i) = "0.00"
                Range("AC" & i) = "Non Compliant. Missing BPR and HM"
            ElseIf BPR >= 0.9 And GPR < 0.9 And HM = "N" Then
                Range("N" & i) = "N"
                Range("I" & i) = "0.00"
                Range("AC" & i) = "Non Compliant. Missing GPR and HM"
            ElseIf BPR >= 0.9 And GPR >= 0.9 And HM = "N" Then
                Range("N" & i) = "N"
                Range("I" & i) = "0.00"
                Range("AC" & i) = "Non Compliant. Missing HM"
            ElseIf GCR = 0 And BPR = 0 And GPR = 0 And HM = "N" Or HM = "" Then
                Range("AC" & i) = "No Data on BW"
                Range("I" & i) = "0.00"
                Range("N" & i) = "N"
        End If
            End If
        End If
    
Next i

' Waivers Override Non Compliance
ActiveSheet.Range("A5:AE5").AutoFilter Field:=14, Criteria1:="N"
ActiveSheet.Range("A5:AE5").AutoFilter Field:=15, Criteria1:="<>" & "#"

TLastCell = Range("A6").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    For i = 1 To Rcount
       With Sheets("Payment Upload").AutoFilter.Range
           'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
           If Rcount >= 1 Then
                Range("O" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
                
                Waivers = ActiveCell.value
                BPR = ActiveCell.Offset(0, 10).value
                GCR = ActiveCell.Offset(0, 11).value
                GPR = ActiveCell.Offset(0, 12).value
                HM = ActiveCell.Offset(0, 13).value
                Chain = ActiveCell.Offset(0, -10).value
                NP = ActiveCell.Offset(0, 6).value
                NTE = ActiveCell.Offset(0, 4).value
                SystemCost = ActiveCell.Offset(0, 9).value
                CarryOverCost = ActiveCell.Offset(0, 16).value
                
                If Left(Chain, Len("4")) = "4" Then
                    If BPR < 0.85 And InStr(1, Waivers, "BPR") > 0 Then
                        ActiveCell.Offset(0, -1).value = "Y"
                        If NP > NTE And SystemCost > NTE And NP > 0 Then
                            ActiveCell.Offset(0, -6).value = NTE
                            ActiveCell.Offset(0, 14).value = "Made by BPR Waiver. Paid on NTE"
                        ElseIf SystemCost < NTE And CarryOverCost > NTE * 2 And NP > NTE Then
                            ActiveCell.Offset(0, -6).value = NTE
                            ActiveCell.Offset(0, 14).value = "Made by BPR Waiver. TU to NTE using CarryoverCost"
                        ElseIf SystemCost > NP And NP < NTE And NP > 0 And CarryOverCost > NP Then
                            ActiveCell.Offset(0, -6).value = NP
                            ActiveCell.Offset(0, 14).value = "Made by BPR Waiver. Paid on NP using CarryoverCost"
                        End If
                    ElseIf GPR < 0.85 And InStr(1, Waivers, "GPR") > 0 Then
                        Range("N" & i) = "Y"
                        If NP > NTE And SystemCost > NTE And NP > 0 Then
                            ActiveCell.Offset(0, -6).value = NTE
                            ActiveCell.Offset(0, 14).value = "Made by GPR Waiver. Paid on NTE"
                        ElseIf SystemCost < NTE And CarryOverCost > NTE * 2 And NP > NTE Then
                            ActiveCell.Offset(0, -6).value = NTE
                            ActiveCell.Offset(0, 14).value = "Made by GPR Waiver. TU to NTE using CarryoverCost"
                        ElseIf SystemCost > NP And NP < NTE And NP > 0 And CarryOverCost > NP Then
                            ActiveCell.Offset(0, -6).value = NP
                            ActiveCell.Offset(0, 14).value = "Made by GPR Waiver. Paid on NP using CarryoverCost"
                        End If
                    ElseIf HM <> "Y" And InStr(1, Waivers, "HEALTHMART") > 0 Then
                        Range("N" & i) = "Y"
                        If NP > NTE And SystemCost > NTE And NP > 0 Then
                            ActiveCell.Offset(0, -6).value = NTE
                            ActiveCell.Offset(0, 14).value = "Made by HM Waiver. Paid on NTE"
                        ElseIf SystemCost < NTE And CarryOverCost > NTE * 2 And NP > NTE Then
                            ActiveCell.Offset(0, -6).value = NTE
                            ActiveCell.Offset(0, 14).value = "Made by HM Waiver. TU to NTE using CarryoverCost"
                        ElseIf SystemCost > NP And NP < NTE And NP > 0 And CarryOverCost > NP Then
                            ActiveCell.Offset(0, -6).value = NP
                            ActiveCell.Offset(0, 14).value = "Made by HM Waiver. Paid on NP using CarryoverCost"
                        End If
                    End If
                ElseIf Left(Chain, Len("3")) = "3" Then
                    If BPR < 0.9 And InStr(1, Waivers, "BPR") > 0 Then
                        ActiveCell.Offset(0, -1).value = "Y"
                        If NP > NTE And SystemCost > NTE And NP > 0 Then
                            ActiveCell.Offset(0, -6).value = NTE
                            ActiveCell.Offset(0, 14).value = "Made by BPR Waiver. Paid on NTE"
                        ElseIf SystemCost < NTE And CarryOverCost > NTE * 2 And NP > NTE Then
                            ActiveCell.Offset(0, -6).value = NTE
                            ActiveCell.Offset(0, 14).value = "Made by BPR Waiver. TU to NTE using CarryoverCost"
                        ElseIf SystemCost > NP And NP < NTE And NP > 0 And CarryOverCost > NP Then
                            ActiveCell.Offset(0, -6).value = NP
                            ActiveCell.Offset(0, 14).value = "Made by BPR Waiver. Paid on NP using CarryoverCost"
                        End If
                    ElseIf GPR < 0.9 And InStr(1, Waivers, "GPR") > 0 Then
                        Range("N" & i) = "Y"
                        If NP > NTE And SystemCost > NTE And NP > 0 Then
                            ActiveCell.Offset(0, -6).value = NTE
                            ActiveCell.Offset(0, 14).value = "Made by GPR Waiver. Paid on NTE"
                        ElseIf SystemCost < NTE And CarryOverCost > NTE * 2 And NP > NTE Then
                            ActiveCell.Offset(0, -6).value = NTE
                            ActiveCell.Offset(0, 14).value = "Made by GPR Waiver. TU to NTE using CarryoverCost"
                        ElseIf SystemCost > NP And NP < NTE And NP > 0 And CarryOverCost > NP Then
                            ActiveCell.Offset(0, -6).value = NP
                            ActiveCell.Offset(0, 14).value = "Made by GPR Waiver. Paid on NP using CarryoverCost"
                        End If
                    ElseIf HM <> "Y" And InStr(1, Waivers, "HEALTHMART") > 0 Then
                        Range("N" & i) = "Y"
                        If NP > NTE And SystemCost > NTE And NP > 0 Then
                            ActiveCell.Offset(0, -6).value = NTE
                            ActiveCell.Offset(0, 14).value = "Made by HM Waiver. Paid on NTE"
                        ElseIf SystemCost < NTE And CarryOverCost > NTE * 2 And NP > NTE Then
                            ActiveCell.Offset(0, -6).value = NTE
                            ActiveCell.Offset(0, 14).value = "Made by HM Waiver. TU to NTE using CarryoverCost"
                        ElseIf SystemCost > NP And NP < NTE And NP > 0 And CarryOverCost > NP Then
                            ActiveCell.Offset(0, -6).value = NP
                            ActiveCell.Offset(0, 14).value = "Made by HM Waiver. Paid on NP using CarryoverCost"
                        End If
                    End If
                End If
            End If
        
            ActiveSheet.Range("A5:AE5").AutoFilter Field:=14, Criteria1:="N"
        End With
    Next i
ActiveSheet.Range("A5:AE5").AutoFilter
'Moved to Liberty Accounts
ActiveSheet.Range("A5:AE5").AutoFilter Field:=30, Criteria1:="Moved to Liberty. Hence No rebate Paid"
Tcell = Range("A5").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
With Sheets("Payment Upload").AutoFilter.Range
    Range("AC" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    Range(ActiveCell.Address & ":AC" & Tcell).SpecialCells(xlCellTypeVisible).value = "Moved to Liberty. Hence No rebate Paid"
    Range("I" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    Range(ActiveCell.Address & ":I" & Tcell).SpecialCells(xlCellTypeVisible).value = "0.00"
End With
    ActiveSheet.Range("A5:U5").AutoFilter

'No System No Rebate Paid
ActiveSheet.Range("A5:AE5").AutoFilter Field:=14, Criteria1:="Y"
ActiveSheet.Range("A5:AE5").AutoFilter Field:=29, Criteria1:=""
ActiveSheet.Range("A5:AE5").AutoFilter Field:=24, Criteria1:="<0.01"
ActiveSheet.Range("A5:AE5").AutoFilter Field:=30, Criteria1:="No System Cost & No/low Carry Over Cost. No rebate paid"
Tcell = Range("A5").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
With Sheets("Payment Upload").AutoFilter.Range
    Range("AC" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    Range(ActiveCell.Address & ":AC" & Tcell).SpecialCells(xlCellTypeVisible).value = "No System Cost & No/low Carry Over Cost. No rebate paid"
    Range("I" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    Range(ActiveCell.Address & ":I" & Tcell).SpecialCells(xlCellTypeVisible).value = "0.00"
End With

ActiveSheet.Range("A5:U5").AutoFilter
    
'TU NTE When Cost is 0
ActiveSheet.Range("A5:AE5").AutoFilter Field:=14, Criteria1:="Y"
ActiveSheet.Range("A5:AE5").AutoFilter Field:=24, Criteria1:=0
ActiveSheet.Range("A5:AE5").AutoFilter Field:=29, Criteria1:=""
        
Tcell = Range("A5").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1

For i = 1 To Rcount
    With Sheets("Payment Upload").AutoFilter.Range
        Range("AC" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
        SystemCost = ActiveCell.Offset(0, -5).value
        RebMax = ActiveCell.Offset(0, -10).value
        NP = ActiveCell.Offset(0, -8).value
        CarryOverCost = ActiveCell.Offset(0, 2).value
        
        'TU to NTE using carry over cost
        If NP > RebMax And CarryOverCost > RebMax Then
            ActiveCell.value = "TU to NTE using carry over cost"
            ActiveCell.Offset(0, -20).value = RebMax
        End If
        'Paid on NP
        If NP < RebMax And CarryOverCost > NP And NP > 0 Then
            ActiveCell.value = "Paid on NP"
            ActiveCell.Offset(0, -20).value = NP
        End If
        'No System Cost & No/low Carry Over Cost. No rebate paid
        If NP > RebMax And CarryOverCost < RebMax Then
            ActiveCell.value = "No System Cost & No/low Carry Over Cost. No rebate paid"
            ActiveCell.Offset(0, -20).value = "0.00"
        End If
        'No System Cost & No/low Carry Over Cost. No rebate paid when Carryover Cost is Zero
        If NP > RebMax And CarryOverCost < 0 Then
            ActiveCell.value = "No System Cost & No/low Carry Over Cost. No rebate paid"
            ActiveCell.Offset(0, -20).value = "0.00"
        End If
        
        ActiveSheet.Range("A5:AE5").AutoFilter Field:=29, Criteria1:=""
    End With
Next i
    ActiveSheet.Range("A5:U5").AutoFilter

'Actual Payments
Sheets("Payment Upload").Activate
ActiveSheet.Range("A5:AE5").AutoFilter Field:=14, Criteria1:="Y"
ActiveSheet.Range("A5:AE5").AutoFilter Field:=29, Criteria1:=""

Tcell = Range("A5").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1

For i = 1 To Rcount
    With Sheets("Payment Upload").AutoFilter.Range
        Range("AC" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
        SystemCost = ActiveCell.Offset(0, -5).value
        RebMax = ActiveCell.Offset(0, -10).value
        NP = ActiveCell.Offset(0, -8).value
        CarryOverCost = ActiveCell.Offset(0, 2).value
        
        'Paid on NTE
        If RebMax < SystemCost And NP > RebMax Then
            ActiveCell.value = "Paid on NTE"
            ActiveCell.Offset(0, -20).value = RebMax
        End If
        'Paid on NP
        If NP < SystemCost Or NP < RebMax And NP < CarryOverCost And NP > 0 Then
            ActiveCell.value = "Paid on NP"
            ActiveCell.Offset(0, -20).value = NP
        End If
        'TU to NTE using Carryover Cost
        If NP > RebMax And SystemCost < RebMax And CarryOverCost > SystemCost And CarryOverCost > RebMax Then
            ActiveCell.value = "TU to NTE using carry over cost"
            ActiveCell.Offset(0, -20).value = RebMax
        End If
        'Paid on SystemCost Low or No System Cost
        If SystemCost < RebMax And CarryOverCost < SystemCost And NP > RebMax And CarryOverCost < RebMax Then
            ActiveCell.value = "Paid on System Cost Low or No System Cost"
            ActiveCell.Offset(0, -20).value = SystemCost
        End If
        
        'Paid on SystemCost Low or No System Cost - Twice
        If SystemCost < RebMax And CarryOverCost < SystemCost And NP > RebMax And CarryOverCost < RebMax Then
            ActiveCell.value = "Paid on System Cost Low or No System Cost"
            ActiveCell.Offset(0, -20).value = SystemCost
        End If
        
        'Paid on SystemCost Low or No System Cost - if Carryover cost is greater than System cost but less them NTE
        If SystemCost < RebMax And NP > RebMax And CarryOverCost < RebMax Then
            ActiveCell.value = "Paid on System Cost Low or No System Cost"
            ActiveCell.Offset(0, -20).value = SystemCost
        End If
        
        'Negative NP. No Rebate
        If NP < 0 Then
            ActiveCell.value = "Negative NP. No Rebate"
            ActiveCell.Offset(0, -20).value = "0.00"
        End If
        
        'Default
        If ActiveCell.value = "" Then
            ActiveCell.value = "N/A"
            ActiveCell.Offset(0, -20).value = "0.00"
        End If
        
        ActiveSheet.Range("A5:AE5").AutoFilter Field:=29, Criteria1:=""
    End With
Next i
ActiveSheet.Range("A5:U5").AutoFilter

'Nick Patel Commment Override
ActiveSheet.Range("A5:AE5").AutoFilter Field:=30, Criteria1:="*" & "Nick Patel Exception" & "*"
Tcell = Range("A5").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1

For i = 1 To Rcount
    With Sheets("Payment Upload").AutoFilter.Range
        Range("AC" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
        'CommentOverride
        CurrComment = ActiveCell.value
        OverrideCom = "Nick Patel Exception. " & CurrComment
        ActiveCell.value = OverrideCom
        ActiveSheet.Range("A5:AE5").AutoFilter Field:=29, Criteria1:="<>" & "*" & "Nick Patel Exception" & "*"
        
    End With
Next i
ActiveSheet.Range("A5:AE5").AutoFilter

APCIWorkingfile.Save
APCIWorkingfile.Close

LMonth = LMonth + 1
WorkFileName2 = CurrentYear & "0" & LMonth

Name "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\Payment Files\APCI\APCI Tech Payment_" & WorkFileName & " Working File.xlsx" As _
   "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\Payment Files\APCI\APCI Tech Payment_" & WorkFileName2 & " Working File.xlsx"


MsgBox "Completed", vbInformation, "Success"

End Sub
Function WaiversSetup()

Sheets.Add.Name = "Sheet1"
Sheets("BW-Compliance Data").Activate
ActiveSheet.Range("A1:DH1").AutoFilter , Field:=20, Criteria1:="<>" & "#"
Rt = Range("D1").SpecialCells(xlCellTypeVisible).End(xlDown).Row
Range("D1:D" & Rt).Copy
Sheets("Sheet1").Activate
Range("A1").PasteSpecial xlPasteValues
Sheets("BW-Compliance Data").Activate
Range("T1:T" & Rt).Copy
Sheets("Sheet1").Activate
Range("B1").PasteSpecial xlPasteValues
ShLRow = Range("A1").End(xlDown).Row

Columns("A:B").Select
    Application.CutCopyMode = False
    ActiveSheet.Range("$A$1:$B$" & ShLRow).RemoveDuplicates Columns:=Array(1, 2), Header _
        :=xlYes

Sheets("Sheet1").Activate
Lrow = Range("A1").End(xlDown).Row
Range("C2").value = "=COUNTIF(A:A,A2)"
Range("C2").Copy
Range("C2:C" & Lrow).PasteSpecial xlPasteAll
'Range("C2:C" & LRow).Copy
'Range("C2:C" & LRow).PasteSpecial xlPasteValues

Range("D2").value = "=LEN(B2)"
Range("D2").Copy
Range("D2:D" & Lrow).PasteSpecial xlPasteAll
'Range("D2:D" & LRow).Copy
'Range("D2:D" & LRow).PasteSpecial xlPasteValues

For G = 1 To Lrow

    ActiveSheet.Range("A1:D1").AutoFilter , Field:=3, Criteria1:=">=2", _
            Operator:=xlAnd
        LStoper = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
        
        If LStoper = 0 Then
            ActiveSheet.Range("A1:D1").AutoFilter
            Exit For
        End If
        With Sheets("Sheet1").AutoFilter.Range
            Range("A" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
            'CommentOverride
            Cust = ActiveCell.value
            ActiveSheet.Range("A1:D1").AutoFilter Field:=1, Criteria1:=Cust
            ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add2 Key:= _
                Range("D1:D" & Lrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
                xlSortNormal
            With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
            Rcount2 = Range("D1").SpecialCells(xlCellTypeVisible).End(xlDown).Address
            HighText = Range(Rcount2).Offset(0, 3).value
            For n = 1 To Rcount
                Range("D" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
                If ActiveCell.value <> HighText Then
                    ActiveCell.EntireRow.Delete
                End If
            Next n
            ActiveSheet.Range("A1:D1").AutoFilter
            
    End With
Next G

End Function
