Attribute VB_Name = "APSC"
'Global Variables
Public CostFile As String
Public BWFile As String
Public APSCFile As Workbook
Public RenamedFile As String
Sub APSCPaymentFile()
Dim SourceFile As String
Dim DestinationFile As String

'Copying Previous Month Payment File
CurentYear = DateAdd("M", -1, Date)
CurrentYear = Format(CurentYear, "YY")
CurrentYear2 = Format(CurentYear, "YYYY")

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
CostFileDestination = "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\Payment Files\APSC\"
SourceFileForIPC = "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Payment Files\" & CurrentYear2 & "\" & LastMonthFolder & LastMonthFolder2 & "\APSC\APSC Tech Payment Summary " & WorkFileName & " - Working File.xlsx"
CostFile.CopyFile SourceFileForIPC, CostFileDestination, True

'Setting up Sheets
Set APSCFile = Workbooks.Open(CostFileDestination & "APSC Tech Payment Summary " & WorkFileName & " - Working File.xlsx")
APSCFile.Activate
Sheets("Payment File").Activate
LMonth = LMonth
CurrentMnth = DateAdd("M", -1, Date)
lastmonth3 = Format(CurrentMnth, "mmmm")
LMonth2 = Format(CurrentMnth, "MM")
Range("B3").value = CurrentYear2 & LMonth2

'Clearing BW File Data from Working File
CurrYear = Year(Date)
Currmonth = Format(DateAdd("M", -2, Now), "MM") 'Change -2 to -1
lastmonth = CurrYear & Currmonth

Sheets("BW-Compliance Data").Activate
LastCell = Range("A2").End(xlDown).Row
Range("A2:DH" & LastCell).Clear

'Putting Data from BW to WF
Dim BWFile As Workbook
Set BWFile = Workbooks.Open("\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\BW Queries\APSC.xlsx") 'Change BW Path
BWFile.Activate
Sheets("Table").Activate
LastCell = Range("G16").End(xlDown).Row
Range("G16:DN" & LastCell).Copy
APSCFile.Activate
Sheets("BW-Compliance Data").Activate
ActiveSheet.AutoFilterMode = False
Range("A2").PasteSpecial xlPasteAll
Application.CutCopyMode = False
' Calling function to sort data by Totoal Purchases in Decensding order
Call Sorting
BWFile.Close

'Setting CarryOver Cost
APSCFile.Activate
Sheets("Carryover cost").Activate

LastColumn = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Address
Range(LastColumn).Select
NewLastCol = ActiveCell.Offset(0, -2).Address
Range(LastColumn & ":" & NewLastCol).Copy
Range(LastColumn).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.PasteSpecial xlPasteAll
lastRow = Range("A1").End(xlDown).Row

Sheets(1).Activate
LastCell = Range("J6").End(xlDown).Row
LastCell = LastCell - 1
Range("J6:J" & LastCell).Copy
Sheets("Carryover cost").Activate

Range(LastColumn).Select
NewLastCol = ActiveCell.Offset(1, 1).Address
Range(NewLastCol).PasteSpecial xlPasteValues

'Copying Formula of CarryOver Cost From Preivous Month val
Sheets("Carryover cost").Activate
NLastCell = Range("A1").End(xlDown).Row

Range(LastColumn).Select
NewLastCol = ActiveCell.Offset(1, 0).Address
Range(NewLastCol).Copy
Range(LastColumn).Select
NewLastCol2 = ActiveCell.Offset(1, 3).Address
Range(NewLastCol2).PasteSpecial xlPasteAll

Dim ComRow As String
ComRow = Range("A1").End(xlDown).Row
CarryLastColumn = Range(LastColumn).Offset(ComRow - 1, 3).Address

Range(NewLastCol2).Copy
Range(NewLastCol2 & ":" & CarryLastColumn).PasteSpecial xlPasteAll

LastTwoMonthName = DateAdd("M", -2, Date)
LTwoMonthName = Format(LastTwoMonthName, "mmmm")

LastOneMonthName = DateAdd("M", -1, Date)
LOMonthName = Format(LastOneMonthName, "mmmm")


Range(LastColumn).Offset(0, 1).value = LTwoMonthName & " Tech Rebate"
Range(LastColumn).Offset(0, 2).value = "COST USED " & LOMonthName
Range(LastColumn).Offset(0, 3).value = LOMonthName & " Carryover Cost"

'Setting Up Payment Sheet
Sheets(1).Activate
LastCell = Range("A6").End(xlDown).Row
Range("H6:H" & LastCell).value = ""
Range("M6:M" & LastCell).value = ""
Range("T6:T" & LastCell).value = ""
Range("K6:K" & LastCell).value = ""
Range("L6:L" & LastCell).value = ""
Range("R6:R" & LastCell).Copy
Range("U6").PasteSpecial xlPasteValues
Range("R6:R" & LastCell).value = ""

LastOneMonthName = DateAdd("M", -1, Date)
LOMonthName = Format(LastOneMonthName, "mm")
CurrentMonthName = DateAdd("M", -0, Date)
CurrMonthName = Format(CurrentMonthName, "mm")

Range("K6:K" & LastCell).value = CurrentYear2 & LOMonthName
Range("L6:L" & LastCell).value = Year(Date) & CurrMonthName

Sheets("BW-Compliance Data").Activate
LastRowin = Range("D2").End(xlDown).Row
Range("D2").Activate
With Range("D2:D" & LastRowin)
    .NumberFormat = "General"
    .value = .value
End With

Sheets(1).Activate
Range("M6").value = "=VLOOKUP(C6,'BW-Compliance Data'!D:BF,55,0)"
Range("M6").Copy
Range("M7:M" & LastCell).PasteSpecial xlPasteAll
Range("M6:M" & LastCell).Copy
Range("M6").PasteSpecial xlPasteValues
Range("M7:M" & LastCell).Replace What:="#N/A", Replacement:="0"

'Getting System Cost
CurrYear = Year(Date)
Currmonth = Format(DateAdd("M", -1, Now), "MM") 'Change -1 to 0
CurrentMonth = CurrentYear2 & Currmonth

Set CstFile = Workbooks.Open("\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\System Cost\CostFiles_Template\Cost File Template_ " & CurrentMonth & ".xlsx")

CstFile.Activate
Sheets("Sheet1").Activate
APSCFile.Activate
Sheets(1).Activate
Columns(18).EntireColumn.Insert
Columns(19).EntireColumn.Insert
Columns(20).EntireColumn.Insert
Range("R6").value = "=VLOOKUP(E6,'[Cost File Template_ " & CurrentMonth & ".xlsx]Sheet1'!$A:$B,2,0)"
Range("S6").value = "=VLOOKUP(C6,'[Cost File Template_ " & CurrentMonth & ".xlsx]Parata '!$B:$C,2,0)"
Range("T6").value = "=VLOOKUP(C6,'[Cost File Template_ " & CurrentMonth & ".xlsx]Prescribed Wellness '!$B:$C,2,0)"

LastCell = Range("A6").End(xlDown).Row
Range("R6:T6").Copy
Range("R7:T" & LastCell).PasteSpecial xlPasteAll
Range("R6:T" & LastCell).Copy
Range("R6").PasteSpecial xlPasteValues

CstFile.Close

'Replacing N/A Value to 0
Range("R6:T" & LastCell).Replace What:="#N/A", Replacement:="0", MatchCase:=True

Dim F As Integer
Dim Cellval As Integer
Cellval = 6
Dim NewCost
For F = 6 To LastCell
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

Range("Q6:Q" & LastCell).Copy
Sheets(3).Activate
Range(LastColumn).Offset(1, 2).PasteSpecial xlPasteValues

'Getting Carryover in Working sheet
Sheets(3).Activate
VCol = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
Sheets(1).Activate
Range("T6").value = "=VLOOKUP(C6,'Carryover cost'!A:XFD," & VCol & ",0)"
Range("T6").Copy
Range("T6:T" & LastCell).PasteSpecial xlPasteAll
Range("T6:T" & LastCell).Copy
Range("T6:T" & LastCell).PasteSpecial xlPasteValues
Application.CutCopyMode = False

Range("A6").Activate

'Payment Setup

'1) Moved to Liberty
FilterValue = "Confirmed that account moved to Liberty and earning rebates through the Liberty program. Hence no rebate paid"
ActiveSheet.Range("A5:U5").AutoFilter Field:=21, Criteria1:="*" & FilterValue & "*", Operator:=xlAnd
TLastCell = Range("A5").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    For i = 1 To Rcount
       With Sheets(1).AutoFilter.Range
           'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
           If Rcount >= 1 Then
                Range("R" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
                ActiveCell.value = FilterValue
                Range(ActiveCell.Address).Offset(0, -10).value = 0
                ActiveSheet.Range("A5:U5").AutoFilter Field:=18, Criteria1:=""
            End If

        End With
    Next i
    ActiveSheet.Range("A5:U5").AutoFilter
    

'2) No Np. No Rebate
ActiveSheet.Range("A5:U5").AutoFilter Field:=18, Criteria1:=""

ActiveSheet.Range("A5:U5").AutoFilter Field:=14, Criteria1:="$-"
TLastCell = Range("A5").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    For i = 1 To Rcount
       With Sheets(1).AutoFilter.Range
           'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
           If Rcount >= 1 Then
                Range("R" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
                ActiveCell.value = "No NP. No rebate Paid"
                Range(ActiveCell.Address).Offset(0, -10).value = 0
                ActiveSheet.Range("A5:U5").AutoFilter Field:=18, Criteria1:=""
            End If

        End With
    Next i
    ActiveSheet.Range("A5:U5").AutoFilter
    
'3) Negative NP

ActiveSheet.Range("A5:U5").AutoFilter Field:=18, Criteria1:=""

ActiveSheet.Range("A5:U5").AutoFilter Field:=14, Criteria1:="<0"
TLastCell = Range("A5").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    For i = 1 To Rcount
       With Sheets(1).AutoFilter.Range
           'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
           If Rcount >= 1 Then
                Range("R" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
                ActiveCell.value = "Negative NP. No Rebate Paid"
                Range(ActiveCell.Address).Offset(0, -10).value = 0
                ActiveSheet.Range("A5:U5").AutoFilter Field:=18, Criteria1:=""
            End If

        End With
    Next i
    ActiveSheet.Range("A5:U5").AutoFilter

'4) No System Cost No Rebate Paid
ActiveSheet.Range("A5:U5").AutoFilter Field:=18, Criteria1:=""

ActiveSheet.Range("A5:U5").AutoFilter Field:=17, Criteria1:="$0.00"
TLastCell = Range("A5").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    For i = 1 To Rcount
       With Sheets(1).AutoFilter.Range
           'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
           If Rcount >= 1 Then
                Range("R" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
                ActiveCell.value = "No system cost; hence no rebate paid"
                Range(ActiveCell.Address).Offset(0, -10).value = 0
                ActiveSheet.Range("A5:U5").AutoFilter Field:=18, Criteria1:=""
            End If

        End With
    Next i
    ActiveSheet.Range("A5:U5").AutoFilter

'5) Payments

ActiveSheet.Range("A5:U5").AutoFilter Field:=18, Criteria1:=""

'ActiveSheet.Range("A5:U5").AutoFilter Field:=17, Criteria1:="$0.00"
TLastCell = Range("A5").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    For i = 1 To Rcount
       With Sheets(1).AutoFilter.Range
           'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
           If Rcount >= 1 Then
                Range("R" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
                FixedCell = ActiveCell.Address
                SystemCost = Range(FixedCell).Offset(0, -1).value
                CarryOCost = Range(FixedCell).Offset(0, 2).value
                AnniversaryDate = Range(FixedCell).Offset(0, 1).value
                
                yy = DateAdd("Y", 1, AnniversaryDate)
                
                D2 = Format(AnniversaryDate, "mmmm")
                

                NP = Range(FixedCell).Offset(0, -4).value
                
                If CarryOCost > 15000 And NP > 15000 And SystemCost > 0 Then
                    Range(FixedCell).Offset(0, -10).value = 15000
                    Range(FixedCell).value = "15K NTE met. Not to be paid until " & D2 & "'" & yy
                End If
                If CarryOCost < 15000 And NP > SystemCost Then
                    Range(FixedCell).Offset(0, -10).value = SystemCost
                    Range(FixedCell).value = "Paid on System Cost as no/low Carry Over Cost. Watch for NTE"
                End If
                If CarryOCost <= NP And NP < SystemCost Then
                    Range(FixedCell).Offset(0, -10).value = NP
                    Range(FixedCell).value = "Paid on NP using Carryover Cost"
                End If
                If CarryOCost > NP And NP < 15000 And NP < SystemCost Then
                    Range(FixedCell).Offset(0, -10).value = NP
                    Range(FixedCell).value = "Paid on NP using Carryover Cost"
                End If
                If CarryOCost >= SystemCost And NP > SystemCost And SystemCost > 0 Then
                    Range(FixedCell).Offset(0, -10).value = SystemCost
                    Range(FixedCell).value = "Paid on System Cost as no/low Carry Over Cost. Watch for NTE"
                End If
                
                If CarryOCost < 15000 And NP < 15000 And SystemCost > 0 Then
                    Range(FixedCell).Offset(0, -10).value = SystemCost
                    Range(FixedCell).value = "Paid on System Cost as no/low Carry Over Cost. Watch for NTE"
                End If
                
                ActiveSheet.Range("A5:U5").AutoFilter Field:=18, Criteria1:=""
            End If

        End With
    Next i
    ActiveSheet.Range("A5:U5").AutoFilter

'6) NTE Already Met

FilterValue = "15K NTE met"
ActiveSheet.Range("A5:U5").AutoFilter Field:=21, Criteria1:="*" & FilterValue & "*", Operator:=xlAnd

'ActiveSheet.Range("A5:U5").AutoFilter Field:=17, Criteria1:="$0.00"
TLastCell = Range("A5").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    For i = 1 To Rcount
       With Sheets(1).AutoFilter.Range
           'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
           If Rcount >= 1 Then
                Range("R" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
                ActiveCell.value = ActiveCell.Offset(0, 3).value
                Range(ActiveCell.Address).Offset(0, -10).value = 0
                ActiveSheet.Range("A5:U5").AutoFilter Field:=18, Criteria1:=""
            End If

        End With
    Next i
    ActiveSheet.Range("A5:U5").AutoFilter

Call APCIRemove

APSCFile.Save
APSCFile.Close

Mth = DateAdd("M", -1, Date)
Mth = Format(Mth, "MM")
Mth2 = DateAdd("M", -2, Date)
Mth2 = Format(Mth2, "MM")

Fname = CurrentYear2 & Mth
Fname2 = "20" & CurrentYear3 & Mth2

Name "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\Payment Files\APSC\APSC Tech Payment Summary " & Fname2 & " - Working File.xlsx" As _
   "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\Payment Files\APSC\APSC Tech Payment Summary " & Fname & " - Working File.xlsx"

MsgBox "Completed", vbInformation, "Success"

End Sub
Function APCIRemove()

'initalizing arraylist to store chains from BW sheet
Dim Chain As ArrayList
Set Chain = New ArrayList

Sheets("BW-Compliance Data").Activate

'Reading all the cell of chain using Loop
For i = 2 To Range("H2").End(xlDown).Row
    ChainTxt = Range("H" & i).value
    'Checking if chaintext contains APCI
    If ChainTxt = "APCI" Then
        Chainval = Range("H" & i).Offset(0, 1).value
        'if the current chain from the cell value already exist below code will not let the duplicate value add
        ex = Chain.Contains(Chainval)
        If ex Then
            'nothing
        Else
            Chain.Add Chainval
        End If
    End If
Next i

'Setting up rebate amount 0 and adding comment "Moved to APCI" if their chain is APCI chain
Sheets("Payment File").Activate

Dim Checkchain
'Reading all the chain from Payment sheet using Loop
For n = 6 To Range("F6").End(xlDown).Row
    Checkchain = Range("F" & n).value
    'Checking if chain has #N/A then it will convert chainvalue to 0
    If IsError(Checkchain) Then
        Checkchain = 0
    End If
    
    'Checking if chain from the payment sheets matches to the arraylist from BW sheet for APCI only, it will put 0 rebate amount and add comment.
    For m = 0 To Chain.Count - 1
        If Checkchain = Chain(m) Then
            Range("H" & n).value = 0
            Range("R" & n).value = "Moved to APCI"
        End If
    Next m
Next n
End Function
Sub Sorting()

Sheets("BW-Compliance Data").Activate

ActiveSheet.Range("A1:DH1").AutoFilter
ActiveWorkbook.Worksheets("BW-Compliance Data").AutoFilter.Sort.SortFields. _
    Clear
ActiveWorkbook.Worksheets("BW-Compliance Data").AutoFilter.Sort.SortFields. _
    Add2 Key:=Range("BF1:BF176"), SortOn:=xlSortOnValues, Order:=xlDescending _
    , DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("BW-Compliance Data").AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
ActiveSheet.Range("A1:DH1").AutoFilter

End Sub

Sub AddingNewCust()

'Getting Payment file name to identify which customer needs to be added from enrollment tracker
CurrentFileName = ActiveWorkbook.Name

'Opening Enrollment tracker
If InStr(CurrentFileName, "APCI") Then
    Dim EnrollmentWB As Workbook
    Set EnrollmentWB = Workbooks.Open("G:\MODEL\ENROLLMENT TRACKER log.xlsx")
    
    'Trying to find the row number and address for checkpoint cell which has cell value - Check for Reliant TRAs
    J = Range("A:A").Rows.Count
    LastCell = Range("A" & J).End(xlUp).Row
    Dim foundCell As Range
    Set foundCell = ActiveSheet.UsedRange.Find(What:="Check for Reliant TRAs", _
                                     LookIn:=xlValues, _
                                     LookAt:=xlWhole, _
                                     SearchOrder:=xlByRows, _
                                     SearchDirection:=xlNext, _
                                     MatchCase:=False)

    'Declaring variables to cell address and row number found for Check for Reliant TRAs
    CheckPoint_Address = foundCell.Address
    CheckPoint_Row = foundCell.Row

End If


'Declaring variable with preivous month in mmmm'yy format
CurDate = DateAdd("M", -1, Date)
PrevDate = Format(CurDate, "mmmm'yy")

'Checking which cell in F column belongs to buying group in the filename
CheckPoint_Row_F = CheckPoint_Row + 1
For m = CheckPoint_Row_F To LastCell
    
    'Getting Payment months in which customer should be included for the payment
    If Trim(Range("F" & m).value) = "APCI" Then
        PaymentMonth = Range("F" & m).Offset(0, 5).value
        PaymentMonth = Replace(PaymentMonth, "Rebate From ", "")
        PaymentMonth = Trim(PaymentMonth)
        If PrevDate = PaymentMonth Then
            MsgBox "Should be Added"
        End If
    End If
Next m
End Sub
