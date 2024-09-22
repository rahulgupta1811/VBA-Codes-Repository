Attribute VB_Name = "IPC_PBA"
Public WorkingFile As Workbook
Public CostFile As Workbook
Public PaymentFile As Workbook
Public DestinationFolderName As String
Public CostFileDesti As String
Public SourceFile As String
Public ParentPath As String
Sub PBAPaymentFile()
Dim DestLocation As String
Dim Subject As ArrayList
Dim lastmonth As String
Dim PrevMonth As String
Dim MSearch As String

Application.AskToUpdateLinks = False

Dim d As Date
d = DateAdd("m", -1, Date)
lastmonth = DateSerial(Year(d), Month(d), "01")
MO = Format(lastmonth, "MM")
Yr = Right(lastmonth, 4)

'Check IPC working file exists
Dim strFileName As String
Dim strFileExists As String

    strFileName = "\\ddcf2015\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\Payment Files\IPC\IPC Payment Summary " & Yr & MO & "_Working File.xlsx"
    strFileExists = Dir(strFileName)

   If strFileExists = "" Then
        MsgBox "IPC Working file does not exists. Please Process IPC first", vbCritical, "Stop"
        Exit Sub
    End If

Dim BWFile As Workbook
Dim WorkingFile As Workbook
Set WorkingFile = Workbooks.Open(strFileName)


'Putting Data from BW to WF
Set BWFile = Workbooks.Open("\\ddcf2015\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\BW Queries\PBA.xlsx") 'Change BW Path
Sheets("Table").Activate
LastCell = Range("G16").End(xlDown).Row
Range("G16:DN" & LastCell).Copy
WorkingFile.Activate
Sheets("BW-Compliance Data").Activate
LastCell = Range("A1").End(xlDown).Row
LastCell = LastCell + 1
Range("A" & LastCell).PasteSpecial xlPasteAll
Application.CutCopyMode = False
' Calling function to sort data by Totoal Purchases in Decensding order
Call APSC.Sorting
BWFile.Close

Sheets("PBA").Activate
LastCell = Range("A3").End(xlDown).Row
Range("K3:K" & LastCell).Copy

Sheets("Carryover cost").Activate
LastColumn = ActiveSheet.Cells(2, ActiveSheet.Columns.Count).End(xlToLeft).Address
Lrow = Range(LastColumn).Offset(0, -2).Address
LastRowAdd = Range(Lrow).End(xlDown).Address
Range(LastRowAdd).Offset(1, 0).PasteSpecial xlPasteValues

'Setting up PBA sheet
Sheets("PBA").Activate
Lcell = Range("A2").End(xlDown).Row
Range("I3:I" & Lcell).value = ""
Range("L3:N" & Lcell).value = ""
Range("R3:R" & Lcell).value = ""
Range("T3:T" & Lcell).value = ""
Range("X3:X" & Lcell).value = ""
Range("U3:U" & Lcell).Copy
Range("V3").PasteSpecial xlPasteValues
Range("U3:U" & Lcell).value = ""

'Preparing System Cost
Sheets("PBA").Activate

Dim n As Integer
For n = 1 To 3
    Columns(18).EntireColumn.Insert
Next n

Range("R2").value = "MPS"
Range("S2").value = "Parata"
Range("T2").value = "PW"

'Getting Cost from Cost File
CurrYear = Year(Date)
Currmonth = Format(DateAdd("M", -1, Now), "MM") 'Change -1 to 0
CurrentMonth = CurrYear & Currmonth

Dim CstFile As Workbook
Set CstFile = Workbooks.Open("\\ddcf2015\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\System Cost\CostFiles_Template\Cost File Template_ " & Yr & MO & ".xlsx")

CstFile.Activate
Sheets("Sheet1").Activate
WorkingFile.Activate
Sheets("PBA").Activate
Range("R3").value = "=VLOOKUP(F3,'[Cost File Template_ " & Yr & MO & ".xlsx]Sheet1'!$A:$B,2,0)"
Range("S3").value = "=VLOOKUP(D3,'[Cost File Template_ " & Yr & MO & ".xlsx]Parata '!$B:$C,2,0)"
Range("T3").value = "=VLOOKUP(D3,'[Cost File Template_ " & Yr & MO & ".xlsx]Prescribed Wellness '!$B:$C,2,0)"

LastCell = Range("A3").End(xlDown).Row
Range("R3:T3").Copy
Range("R4:T" & LastCell).PasteSpecial xlPasteAll
Range("R3:T" & LastCell).Copy
Range("R3").PasteSpecial xlPasteValues

CstFile.Close

'Replacing N/A Value to 0
Range("R3:T" & LastCell).Replace What:="#N/A", Replacement:="0", MatchCase:=True

Dim F As Integer
Dim Cellval As Integer
Cellval = 3
Dim NewCost
lastRow = LastCell - 2
For F = 3 To lastRow
    MPS = Range("R" & Cellval).value
    Parata = Range("S" & Cellval).value
    PW = Range("T" & Cellval).value
    NewCost = MPS + Parata + PW
    Range("Q" & Cellval).value = NewCost
    Cellval = Cellval + 1
Next F

'Deleting Extra Created Columns
For n = 1 To 3
    Columns(18).EntireColumn.Delete
Next n

'Copying System Cots into Carryover Sheet
LastCell = Range("Q3").End(xlDown).Row
Range("Q3:Q" & LastCell).Copy

Sheets("Carryover cost").Activate

Set Regex = CreateObject("VBScript.RegExp")
Regex.Global = True
Regex.Pattern = "[0-9]"
OutStr = Regex.Replace(LastColumn, "")
OutStr = Replace(OutStr, "$", "")
Range(LastRowAdd).Offset(1, 1).Select
ActiveCell.PasteSpecial xlPasteValues

ActiveSheet.UsedRange.EntireColumn.AutoFit
Application.CutCopyMode = False

'BW Customer Number Format Change
Sheets("BW-Compliance Data").Activate
LastRowin = Range("D2").End(xlDown).Row

With Range("D2:D" & LastRowin)
    .NumberFormat = "General"
    .value = .value
End With

Sheets("PBA").Activate
Range("L3:L" & LastCell).value = Yr & MO
MoCurr = DateAdd("M", 0, Now)
YrCurr = DateAdd("Y", 0, Now)
MCurr = Format(MoCurr, "mm")
YCurr = Format(YrCurr, "yyyy")
Range("M3:M" & LastCell).value = YCurr & MCurr

'Putting Values Using Vlookup
Range("R3").value = "=VLOOKUP(D3,'BW-Compliance Data'!D:BF,55,0)"
Range("T3").value = "=CONCATENATE(VLOOKUP(D3,'BW-Compliance Data'!D:BC,52,0),$N$1)"
Range("R3").Copy
Range("R3:R" & LastCell).PasteSpecial xlPasteAll
Range("R3:R" & LastCell).Copy
Range("R3:R" & LastCell).PasteSpecial xlPasteValues
Range("R3:R" & LastCell).Replace What:="#N/A", Replacement:="0", MatchCase:=True

Range("T3").Copy
Range("T3:T" & LastCell).PasteSpecial xlPasteAll
Range("T3:T" & LastCell).Copy
Range("T3:T" & LastCell).PasteSpecial xlPasteValues
Range("T3:T" & LastCell).Replace What:="#N/A", Replacement:="0", MatchCase:=True

With Range("T3:T" & LastCell)
    .NumberFormat = "General"
    .value = .value
End With

'Getting Latest Carry OverCost in PBA Shet
Sheets("Carryover cost").Activate
Range("A2").Select
LastCol = ActiveSheet.Cells(2, ActiveSheet.Columns.Count).End(xlToLeft).Column
Sheets("PBA").Activate
Range("X3").value = "=VLOOKUP(D3,'Carryover cost'!A:XFD," & LastCol & ",0)"
Range("X3").Copy
Range("X3:X" & LastCell).PasteSpecial xlPasteAll
Range("X3:X" & LastCell).Copy
Range("X3:X" & LastCell).PasteSpecial xlPasteValues

Application.CutCopyMode = False

'Evalution
Sheets("PBA").Activate

Dim CR As Integer
Dim ir As Integer
Dim LastnRow As Integer
LastnRow = Range("A3").End(xlDown).Row
CR = 3
For ir = 3 To LastnRow

    Dim GCR
    GCR = Range("T" & CR).value
    If GCR >= 0.16 Then
        Range("N" & CR).value = "Y"
    ElseIf GCR = "#N/A" Then
        Range("N" & CR).value = "N"
        Range("T" & CR).value = "0"
    Else
        Range("N" & CR).value = "N"
    End If
    CR = CR + 1
Next ir

'1) No System Cost No Rebate

ActiveSheet.Range("A2:X2").AutoFilter Field:=17, Criteria1:="$0.00"
TLastCell = Range("A2").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    For i = 1 To Rcount
       With Sheets("PBA").AutoFilter.Range
           'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
           If Rcount >= 1 Then
                Range("U" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
                ActiveCell.value = "No System Cost as verified against Cost File; hence no rebate earned"
                Range(ActiveCell.Address).Offset(0, -12).value = 0
                ActiveSheet.Range("A2:X2").AutoFilter Field:=21, Criteria1:=""
            End If

        End With
    Next i
    ActiveSheet.Range("A2:X2").AutoFilter

'2) Non Compliant
ActiveSheet.Range("A2:X2").AutoFilter Field:=14, Criteria1:="N"
ActiveSheet.Range("A2:X2").AutoFilter Field:=21, Criteria1:=""
TLastCell = Range("A2").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    For i = 1 To Rcount
       With Sheets("PBA").AutoFilter.Range
           'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
           If Rcount >= 1 Then
                Range("U" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
                ActiveCell.value = "Non compliant. Missing GCR"
                Range(ActiveCell.Address).Offset(0, -12).value = 0
                ActiveSheet.Range("A2:X2").AutoFilter Field:=21, Criteria1:=""
            End If

        End With
    Next i
    ActiveSheet.Range("A5:U5").AutoFilter

'3) Paid on system cost on following trend
FilterValue = "Paid on System Cost following trend"
ActiveSheet.Range("A2:X2").AutoFilter Field:=14, Criteria1:="Y"
ActiveSheet.Range("A2:X2").AutoFilter Field:=21, Criteria1:=""
ActiveSheet.Range("A2:X2").AutoFilter Field:=22, Criteria1:="*" & FilterValue & "*", Operator:=xlAnd

TLastCell = Range("A2").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    For i = 1 To Rcount
       With Sheets("PBA").AutoFilter.Range
           'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
           If Rcount >= 1 Then
                Range("U" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
                ActiveCell.value = ActiveCell.Offset(0, 1).value
                Range(ActiveCell.Address).Offset(0, -12).value = ActiveCell.Offset(0, -4).value
                ActiveSheet.Range("A2:X2").AutoFilter Field:=21, Criteria1:=""
            End If

        End With
    Next i
    ActiveSheet.Range("A2:X2").AutoFilter

'4) Hoover Drug Customer and non PBA customer Exlcusion

FilterValue = "Parata claims that Hoover Drug is not their customer, hence no System Cost; no rebate paid"
ActiveSheet.Range("A2:X2").AutoFilter Field:=21, Criteria1:=""
ActiveSheet.Range("A2:U2").AutoFilter Field:=22, Criteria1:="*" & FilterValue & "*", Operator:=xlAnd

TLastCell = Range("A2").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    For i = 1 To Rcount
       With Sheets("PBA").AutoFilter.Range
           'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
           If Rcount >= 1 Then
                Range("U" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
                ActiveCell.value = ActiveCell.Offset(0, 1).value
                Range(ActiveCell.Address).Offset(0, -12).value = 0
                ActiveSheet.Range("A2:X2").AutoFilter Field:=21, Criteria1:=""
            End If

        End With
    Next i
    ActiveSheet.Range("A2:X2").AutoFilter
    
' Non POB accounts
FilterValue = "This customer is no longer PBA and PPA(recoup) has been issued as per request through SFDC 12024780"

ActiveSheet.Range("A2:X2").AutoFilter Field:=22, Criteria1:="*" & FilterValue & "*", Operator:=xlAnd

TLastCell = Range("A2").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    For i = 1 To Rcount
       With Sheets("PBA").AutoFilter.Range
           'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
           If Rcount >= 1 Then
                Range("U" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
                ActiveCell.value = ActiveCell.Offset(0, 1).value
                Range(ActiveCell.Address).Offset(0, -12).value = 0
                ActiveSheet.Range("A2:X2").AutoFilter Field:=21, Criteria1:=""
            End If

        End With
    Next i
    ActiveSheet.Range("A2:X2").AutoFilter
    
'5) Payments
ActiveSheet.Range("A2:X2").AutoFilter Field:=14, Criteria1:="Y"
ActiveSheet.Range("A2:X2").AutoFilter Field:=21, Criteria1:=""

TLastCell = Range("A2").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    For i = 1 To Rcount
       With Sheets("PBA").AutoFilter.Range
           'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
           If Rcount >= 1 Then
                Range("U" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
                NP = ActiveCell.Offset(0, -2).value
                CarrCost = ActiveCell.Offset(0, 3).value
                SysCost = ActiveCell.Offset(0, -4).value
                AnnMonth = ActiveCell.Offset(0, 2).value
                AnMonth = Format(AnnMonth, "mmmm")
                Yr = Year(Date) + 1
                yrr = Replace(Yr, "20", "")
                Curr = Format(Date, "mmmm")
                
                If CarrCost >= 10000 And NP >= 10000 And SysCost > 0 Then
                    ActiveCell.value = "10K NTE met. Not to be Paid Until " & AnMonth & "'" & yrr
                    If AnMonth = Curr Then
                        Range(ActiveCell.Address).Offset(0, -12).value = 10000
                    Else
                        Range(ActiveCell.Address).Offset(0, -12).value = 0
                    End If
                    
                    ActiveSheet.Range("A2:X2").AutoFilter Field:=21, Criteria1:=""
                End If
                
                If CarrCost < 10000 And SysCost > 0 And NP > SysCost Then
                    ActiveCell.value = "Paid on system cost as low/no carryover cost"
                    Range(ActiveCell.Address).Offset(0, -12).value = ActiveCell.Offset(0, -4).value
                    ActiveSheet.Range("A5:X2").AutoFilter Field:=21, Criteria1:=""
                End If
                      
            End If

        End With
    Next i
    ActiveSheet.Range("A2:X2").AutoFilter
    
'6) 10K NTE already Met Overide
ActiveSheet.Range("A2:X2").AutoFilter
FilterValue = "10K NTE met."
ActiveSheet.Range("A2:X2").AutoFilter Field:=22, Criteria1:="*" & FilterValue & "*", Operator:=xlAnd
ActiveSheet.Range("A2:X2").AutoFilter Field:=21, Criteria1:="<>" & FilterValue & "*", Operator:=xlAnd

TLastCell = Range("A2").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    For i = 1 To Rcount
       With Sheets("PBA").AutoFilter.Range
           'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
            Range("U" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
            AnnMonth = ActiveCell.Offset(0, 2).value
            AnMonth = Format(AnnMonth, "mmmm")
            Yr = Year(Date) + 1
            yrr = Replace(Yr, "20", "")
            Curr = Format(Date, "mmmm")
           
           If Rcount >= 1 Then
                Range("U" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
                If AnMonth <> Curr Then
                    ActiveCell.value = ActiveCell.Offset(0, 1).value
                    Range(ActiveCell.Address).Offset(0, -12).value = 0
                    ActiveSheet.Range("A2:X2").AutoFilter Field:=21, Criteria1:=""
                End If
            End If

        End With
    Next i
    ActiveSheet.Range("A2:X2").AutoFilter
    
Range("A2").Activate

WorkingFile.Save
WorkingFile.Close

MsgBox "Completed", vbInformation, "Sucess"

End Sub

