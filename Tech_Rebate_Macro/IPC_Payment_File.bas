Attribute VB_Name = "IPC_Buying_Group"
Public NC, C As String
Public BPR, GPR, GCR
Public HM As String
Public Cat As String
Public PreComments As String
Public IPCworkingFile As Workbook
Public TrendFile As Workbook
Sub IPCPaymentFile()

'Copying IPC File

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
CostFileDestination = "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\Payment Files\IPC\"
SourceFileForIPC = "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Payment Files\" & CurrentYear & "\" & LastMonthFolder & LastMonthFolder2 & "\IPC\IPC Payment Summary " & WorkFileName & "_Working File.xlsx"
CostFile.CopyFile SourceFileForIPC, CostFileDestination, True

Dim WorkingFile As String
WorkingFile = "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\Payment Files\IPC\IPC Payment Summary " & WorkFileName & "_Working File.xlsx"
Set IPCworkingFile = Workbooks.Open(WorkingFile)

IPCworkingFile.Activate
Sheets(1).Copy Before:=Sheets(1)

Mth = DateAdd("M", -1, Date)
Mth = Format(Mth, "MM")

Sheets(1).Name = CurrentYear & Mth & " IPC Tech Rebates"
Application.DisplayAlerts = False
Sheets(3).Delete
Application.DisplayAlerts = True
Sheets(1).Activate
Range("A3").value = Year(Date) & Mth

'Clearing BW File Data from Working File
CurrYear = Year(Date)
Currmonth = Format(DateAdd("M", -2, Now), "MM") 'Change -2 to -1
lastmonth = CurrYear & Currmonth

Sheets("BW-Compliance Data").Activate
ActiveSheet.AutoFilterMode = False
LastCell = Range("A2").End(xlDown).Row
Range("A2:DH" & LastCell).Clear

'Putting Data from BW to WF
Dim BWFile As Workbook
Set BWFile = Workbooks.Open("\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\BW Queries\IPC.xlsx") 'Change BW Path
BWFile.Activate
Sheets("Table").Activate
LastCell = Range("G16").End(xlDown).Row
Range("G16:DN" & LastCell).Copy
IPCworkingFile.Activate
Sheets("BW-Compliance Data").Activate
Range("A2").PasteSpecial xlPasteAll
Application.CutCopyMode = False
Application.DisplayAlerts = False
' Calling function to sort data by Totoal Purchases in Decensding order
Call APSC.Sorting
BWFile.Close

'Setting CarryOver Cost
IPCworkingFile.Activate
Sheets("Carryover cost").Activate

LastColumn = ActiveSheet.Cells(2, ActiveSheet.Columns.Count).End(xlToLeft).Address
Range(LastColumn).Select
NewLastCol = ActiveCell.Offset(0, -2).Address
Range(LastColumn & ":" & NewLastCol).Copy
Range(LastColumn).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.PasteSpecial xlPasteAll
lastRow = Range("A2").End(xlDown).Row
    
Sheets(1).Activate
LastCell = Range("K6").End(xlDown).Row
LastCell = LastCell - 1
Range("K7:K" & LastCell).Copy
Sheets("Carryover cost").Activate

Range(LastColumn).Select
NewLastCol = ActiveCell.Offset(1, 1).Address
Range(NewLastCol).PasteSpecial xlPasteValues

'Copying Formula of CarryOver Cost From Preivous Month val
Sheets("Carryover cost").Activate
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

'Getting Last Month Rebate Amount
Sheets(1).Activate
lastRow = Range("A6").End(xlDown).Row
Range("K7:K" & lastRow).Copy

Sheets("Carryover cost").Activate
Range(LastColumn).Offset(1, 2).Select
Range(ActiveCell.Address).PasteSpecial xlPasteValues

'copyingFormat
Lcell = Range("A2").End(xlDown).Row
C1 = Range(LastColumn).Offset(0, -2).Address
C2 = Range(LastColumn).Offset(Lcell, -1).Address
Range(C1 & ":" & C2).Copy
JState = Range(LastColumn).Offset(0, 1).Address
Range(JState).PasteSpecial xlPasteFormats
ActiveSheet.UsedRange.EntireColumn.AutoFit
Range(LastColumn).Offset(0, 3).Select

'Setting Column Month
Last2month = DateAdd("M", -2, Date)
Last2month = Format(Last2month, "mmmm")

lastmonth = DateAdd("M", -1, Date)
lastmonth = Format(lastmonth, "mmmm")
Range(LastColumn).Offset(0, 1).value = Last2month & " Month Payment"
Range(LastColumn).Offset(0, 2).value = "Cost " & lastmonth
Range(LastColumn).Offset(0, 3).value = lastmonth & " CARRY OVER COST"

'Copying Preivous Month Details
Sheets(1).Activate
lastRow = Range("A6").End(xlDown).Row

'Copying Paid Month Date into Rebate Month
Range("P7:P" & lastRow).Copy
Range("P7").PasteSpecial xlPasteValuesAndNumberFormats

'Copying Last Two Months Rebate Amount
Range("L7:M" & lastRow).Copy
Range("M7").PasteSpecial xlPasteValuesAndNumberFormats

lastmonth = DateAdd("M", -1, Date)
lastmonth = Format(lastmonth, "mmmm")
Last2month = DateAdd("M", -2, Date)
Last2month = Format(Last2month, "mmmm")
Last3month = DateAdd("M", -3, Date)
Last3month = Format(Last3month, "mmmm")

Range("L6").value = "Last PMT Month-" & lastmonth
Range("M6").value = "Last PMT Month-" & Last2month
Range("N6").value = "Last PMT Month-" & Last3month

NowMonth = DateAdd("M", 0, Date)
NowMonth = Format(NowMonth, "MM")
NowYear = DateAdd("Y", 0, Date)
NowYear = Format(NowYear, "YYYY")
PaidMonth = NowYear & NowMonth

Now2Month = DateAdd("M", -1, Date)
Now2Month = Format(Now2Month, "MM")
Paid2Month = CurrentYear & Now2Month

Range("P7:P" & lastRow).value = PaidMonth
Range("O7:O" & lastRow).value = Paid2Month

''Copying Last Months Rebate Amount.
Range("K7:K" & lastRow).Copy
Range("L7").PasteSpecial xlPasteValuesAndNumberFormats
'Copying Last Months Coments
Range("Y7:Y" & lastRow).Copy
Range("AF7").PasteSpecial xlPasteValues

'Clearing Rows
Range("I7:I" & lastRow).Clear
Range("Q7:Q" & lastRow).Clear
Range("W7:W" & lastRow).Clear
Range("R7:R" & lastRow).Clear
Range("AE7:AE" & lastRow).Clear
Range("Z7:AC" & lastRow).Clear
Range("Y7:Y" & lastRow).Clear

'Getting Data from BW Sheet
Sheets("BW-Compliance Data").Activate
ActiveSheet.AutoFilterMode = False
LastRowin = Range("D2").End(xlDown).Row
Range("D2").Activate
With Range("D2:D" & LastRowin)
    .NumberFormat = "General"
    .value = .value
End With

Sheets(1).Activate

Range("R7").value = "=VLOOKUP(D7,'BW-Compliance Data'!D:BF,17,0)"
Range("R7").Copy
Range("R7:R" & lastRow).PasteSpecial xlPasteAll
Range("R7:R" & lastRow).Copy
Range("R7:R" & lastRow).PasteSpecial xlPasteValues

Range("Z7").value = "=CONCATENATE(VLOOKUP(D7,'BW-Compliance Data'!D:BC,48,0),$Z$1)"
Range("Z7").Copy
Range("Z7:Z" & lastRow).PasteSpecial xlPasteAll
Range("Z7:Z" & lastRow).Copy
Range("Z7:Z" & lastRow).PasteSpecial xlPasteValues

Range("AA7").value = "=CONCATENATE(VLOOKUP(D7,'BW-Compliance Data'!D:BC,52,0),$Z$1)"
Range("AA7").Copy
Range("AA7:AA" & lastRow).PasteSpecial xlPasteAll
Range("AA7:AA" & lastRow).Copy
Range("AA7:AA" & lastRow).PasteSpecial xlPasteValues

Range("AB7").value = "=CONCATENATE(VLOOKUP(D7,'BW-Compliance Data'!D:BC,49,0),$Z$1)"
Range("AB7").Copy
Range("AB7:AB" & lastRow).PasteSpecial xlPasteAll
Range("AB7:AB" & lastRow).Copy
Range("AB7:AB" & lastRow).PasteSpecial xlPasteValues


Range("AC7").value = "=VLOOKUP(D7,'BW-Compliance Data'!D:BC,15,0)"
Range("AC7").Copy
Range("AC7:AC" & lastRow).PasteSpecial xlPasteAll
Range("AC7:AC" & lastRow).Copy
Range("AC7:AC" & lastRow).PasteSpecial xlPasteValues


Range("W7").value = "=VLOOKUP(D7,'BW-Compliance Data'!D:BF,55,0)"
Range("W7").Copy
Range("W7:W" & lastRow).PasteSpecial xlPasteAll
Range("W7:W" & lastRow).Copy
Range("W7:W" & lastRow).PasteSpecial xlPasteValues

'Removing N/A
Range("R7:R" & lastRow).Replace What:="#N/A", Replacement:="#"
Range("W7:W" & lastRow).Replace What:="#N/A", Replacement:="0"
Range("Z7:AB" & lastRow).Replace What:="#N/A", Replacement:="0"
Range("AC7:AC" & lastRow).Replace What:="#N/A", Replacement:=""
Range("V7").Copy
Range("W7:W" & lastRow).PasteSpecial xlPasteFormats

For i = 26 To 28
    With ActiveSheet.Columns(i)
        .NumberFormat = "0"
        .value = .value
    End With
Next i
Range("Z7:AB" & lastRow).Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"


'Getting Cost from Cost File
CurrYear = Year(Date)
Currmonth = Format(DateAdd("M", -1, Now), "MM") 'Change -1 to 0
CurrentMonth = CurrentYear & Currmonth

Dim CstFile As Workbook
Set CstFile = Workbooks.Open("\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\System Cost\CostFiles_Template\Cost File Template_ " & CurrentMonth & ".xlsx")

CstFile.Activate
Sheets("Sheet1").Activate
IPCworkingFile.Activate
Sheets(1).Activate
Columns(23).EntireColumn.Insert
Columns(23).EntireColumn.Insert
Columns(23).EntireColumn.Insert
Range("W7").value = "=VLOOKUP(F7,'[Cost File Template_ " & CurrentMonth & ".xlsx]Sheet1'!$A:$B,2,0)"
Range("X7").value = "=VLOOKUP(D7,'[Cost File Template_ " & CurrentMonth & ".xlsx]Parata '!$B:$C,2,0)"
Range("Y7").value = "=VLOOKUP(D7,'[Cost File Template_ " & CurrentMonth & ".xlsx]Prescribed Wellness '!$B:$C,2,0)"

LastCell = Range("A6").End(xlDown).Row
Range("W7:Y7").Copy
Range("W7:Y" & lastRow).PasteSpecial xlPasteAll
Range("W7:Y" & lastRow).Copy
Range("W7").PasteSpecial xlPasteValues

CstFile.Close

'Replacing N/A Value to 0
Range("W7:Y" & LastCell).Replace What:="#N/A", Replacement:="0", MatchCase:=True

Dim F As Integer
Dim Cellval As Integer
Cellval = 7
Dim NewCost
For F = 1 To lastRow
    MPS = Range("W" & Cellval).value
    Parata = Range("X" & Cellval).value
    PW = Range("Y" & Cellval).value
    NewCost = MPS + Parata + PW
    Range("V" & Cellval).value = NewCost
    Cellval = Cellval + 1
Next F

'Deleting Extra Created Columns
For n = 1 To 3
    Columns(23).EntireColumn.Delete
Next n

Range("W7").Copy
Range("V7:V" & lastRow).PasteSpecial xlPasteFormats

'Putting Cost into Carrover sheet
Range("V7:V" & lastRow).Copy

Sheets("Carryover Cost").Activate

Range(LastColumn).Offset(1, 2).Select
Range(ActiveCell.Address).PasteSpecial xlPasteValues

'Getting Carryover in Working sheet
Sheets("Carryover cost").Activate
LastCol = ActiveSheet.Cells(2, ActiveSheet.Columns.Count).End(xlToLeft).Column
Sheets(1).Activate
Range("AE7").value = "=VLOOKUP(D7,'Carryover cost'!A:XFD," & LastCol & ",0)"
Range("AE7").Copy
Range("AE7:AE" & lastRow).PasteSpecial xlPasteAll
Range("AE7:AE" & lastRow).Copy
Range("AE7:AE" & lastRow).PasteSpecial xlPasteValues

Range("AE7:AE" & lastRow).Select
    Selection.Style = "Currency"

Range("A7").Activate


Call CheckCompliance

End Sub
Function CheckCompliance()
lastRow = Range("A6").End(xlDown).Row
For i = 7 To lastRow
    'Setting Cells values to variables
    Cat = Range("A" & i).value
    Cat = Trim(Cat)
    BPR = Range("Z" & i).value
    GCR = Range("AA" & i).value
    GPR = Range("AB" & i).value
    HM = Range("AC" & i).value
    NP = Range("X" & i).value
    
    'Setting COmpliant for "OLD" Catagory Customers
    If Cat = "Old" Then
        If BPR >= 0.9 And GCR >= 0.1 And GPR >= 0.9 And HM = "Y" Then 'Condition - 1 -> If BPR,GPR,GCR and HM are meeting Set customer to Compliant
           Range("Q" & i).value = "Y"
        Else
            If GCR >= 0.16 And HM = "Y" Then 'Condition -2 -> Overriding Above Condition. IF GCR and HM are meeting then set customer to Compliant
                Range("Q" & i).value = "Y"
            Else
                Range("Q" & i).value = "N" 'Condition 3- -> If Both Condition fails then set customer to non compliant
            End If
        End If
    'Setting Compliant for "NEW" Catagory Customers
    Else
        If BPR >= 0.9 And GCR >= 0.12 And GPR >= 0.9 Then 'Condition - 1 -> If BPR,GPR,and GCR are meeting Set customer to Compliant
           Range("Q" & i).value = "Y"
        Else
            If GCR >= 0.2 Then 'Condion 2 -> Overriding Above Condition. IF GCR and HM are meeting then set customer to Compliant
                Range("Q" & i).value = "Y"
            Else
                Range("Q" & i).value = "N" 'Condition 3 -> If Both Condition fails then set customer to non compliant
            End If
        End If
    End If
'Calling Comments Writing function for Non Compliant Customer within Loop
    Call CommentWriter(BPR, GPR, GCR, HM, Cat, "Y" & i)

'Setting Comments -  No Data in BW
    If NP = 0 And BPR = 0 And GPR = 0 And GCR = 0 And HM = "" Then
        Range("Y" & i).value = "No Data In BW."
    End If

Next i

'Filtering Non Compliant and Putting Zero for their Amount
Call FilterSetup(17, "N", "I", 0)

'Overwritting Comments based on last month Comments
Call FilterSetup(32, "No System Cost, no rebate paid", "Y", "No System Cost, no rebate paid")
Range("A6:AF6").AutoFilter Field:=32

Call FilterSetupContains(32, "10K NTE", "Y")

Call ChangeAnniversaryComment


End Function
Function CommentWriter(xBPR, xGPR, xGCR, xHM, xCat, CellV)

BPR = xBPR
GPR = xGPR
GCR = xGCR
HM = xHM
Cat = xCat
NC = "Non Compliant. Missing "
If Cat = "Old" Then
    If GCR >= 0.16 And HM = "Y" Then
        NC = NC
    Else
        If BPR < 0.9 Then
           NC = NC & "BPR"
        End If
        
        If GPR < 0.9 Then
            If InStr(NC, "BPR") > 0 Then
                NC = NC & " and GPR"
            Else
                NC = NC & "GPR"
            End If
        End If
        
        If GCR < 0.1 Then
            If InStr(NC, "BPR") > 0 And InStr(NC, "GPR") > 0 Then
                NC = Replace(NC, "and", ",")
                NC = NC & " and GCR"
            ElseIf InStr(NC, "BPR") > 0 Or InStr(NC, "GPR") > 0 Then
                NC = Replace(NC, "and", ",")
                NC = NC & " and GCR"
            Else
                NC = NC & "GCR"
            End If
        End If
        
        If HM = "N" Or HM = "" Then
            If InStr(NC, "BPR") > 0 Then
                NC = Replace(NC, "and", ",")
                NC = NC & " and HM"
            ElseIf InStr(NC, "BPR") > 0 And InStr(NC, "GPR") > 0 Then
                NC = Replace(NC, "and", ",")
                NC = NC & "and HM"
            ElseIf InStr(NC, "BPR") > 0 And InStr(NC, "GPR") > 0 And InStr(NC, "GCR") > 0 Then
                NC = Replace(NC, "and", ",")
                NC = NC & " and HM"
             ElseIf InStr(NC, "BPR") > 0 Or InStr(NC, "GPR") > 0 Or InStr(NC, "GCR") > 0 Then
                NC = Replace(NC, "and", ",")
                NC = NC & " and HM"
            Else
                NC = NC & "HM"
            End If
        End If
        
        If NC = "Non Compliant. Missing " Then
            NC = NC
        Else
            Range(CellV).value = NC
        End If
    End If
End If

If Cat = "New" Then
    If GCR >= 0.2 Then
        NC = NC
    Else
        If BPR < 0.9 Then
           NC = NC & "BPR"
        End If
        
        If GPR < 0.9 Then
            If InStr(NC, "BPR") > 0 Then
                NC = NC & " and GPR"
            Else
                NC = NC & "GPR"
            End If
        End If
        
        If GCR < 0.12 Then
            If InStr(NC, "BPR") > 0 And InStr(NC, "GPR") > 0 Then
                NC = Replace(NC, "and", ",")
                NC = NC & " and GCR"
            ElseIf InStr(NC, "BPR") > 0 Or InStr(NC, "GPR") > 0 Then
                NC = Replace(NC, "and", ",")
                NC = NC & " and GCR"
            Else
                NC = NC & "GCR"
            End If
        End If
        
        If NC = "Non Compliant. Missing " Then
            NC = NC
        Else
            Range(CellV).value = NC
        End If
    End If
End If

End Function

Public Function FilterSetup(ColumnToPutFilterOn As Integer, FilterValue, ColumnToSetValue, value)
ActiveSheet.Range("A6:AF6").AutoFilter Field:=ColumnToPutFilterOn, Criteria1:=FilterValue
TLastCell = Range("A6").End(xlDown).Row
   With Sheets(1).AutoFilter.Range
       'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
       Range(ColumnToSetValue & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
       CurrentCell = ActiveCell.Address
       'Getting Last row of the visible cells after filtering
       Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
       If Rcount > 1 Then
        'Putting Zeros in amount cells
           Range(CurrentCell & ":" & ColumnToSetValue & TLastCell).SpecialCells(xlCellTypeVisible).SpecialCells(xlCellTypeVisible).value = value
       End If
    End With
'ActiveSheet.AutoFilterMode = False
End Function

Function FilterSetupContains(ColumnToPutFilterOn As Integer, FilterValue, ColumnToSetValue)
Dim lastmonth As String
lastmonth = Format(DateAdd("m", -1, Date), "mmmm")
lastMonthYear = Right(Year(DateAdd("m", -1, Date)), 2)
NotSelect = lastmonth & "'" & lastMonthYear
ActiveSheet.Range("A6:AF6").AutoFilter Field:=ColumnToPutFilterOn, Criteria1:="*" & FilterValue & "*", Operator:=xlAnd, Criteria2:="<>*" & NotSelect
TLastCell = Range("A6").End(xlDown).Row
   With Sheets(1).AutoFilter.Range
       'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
       Range(ColumnToSetValue & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
       CurrentCell = ActiveCell.Address
       'Getting Last row of the visible cells after filtering
       Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
       Range("Y" & .Offset(1, 0).SpecialCells(xlCellTypeVisible).Row).Select
        Cfrom = ActiveCell.Row
        Addr = Range("AF" & Rows.Count).End(xlUp).Row
        
        Range("Y" & Cfrom).value = "=AF" & Cfrom
        Range("Y" & Cfrom).Copy
        Range("Y" & Cfrom & ":Y" & Addr).SpecialCells(xlCellTypeVisible).PasteSpecial xlPasteAll
        Application.CutCopyMode = False
        Range("A6:AF6").AutoFilter Field:=32
        Range("A6:AF6").AutoFilter Field:=17
        Range("Y7:Y" & Addr).Copy
        Range("Y7:Y" & Addr).PasteSpecial xlPasteValues
        Application.CutCopyMode = False
    End With
'ActiveSheet.AutoFilterMode = False
End Function
Public Function ChangeAnniversaryComment()
FilterValue = "Anniversary Month"
ActiveSheet.Range("A6:AF6").AutoFilter Field:=25, Criteria1:="*" & FilterValue & "*", Operator:=xlAnd
TLastCell = Range("A6").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    For i = 1 To Rcount
       With Sheets(1).AutoFilter.Range
           'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
           If Rcount >= 1 Then
            Range("Y" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
            End If
           CurrentCell = ActiveCell.Address
           CelCon = ActiveCell.value
           'Getting Last row of the visible cells after filtering
            CelCon = Replace(CelCon, FilterValue, "")
            Filter2 = " met using carryover cost"
            CelCon = Replace(CelCon, Filter2, "")
            CelCon = Replace(CelCon, ". 10K", "10K")
            Range(CurrentCell).value = CelCon
        End With
    Next i
    
Call NonPaymentCompliant
End Function
Sub NonPaymentCompliant()
ActiveSheet.AutoFilterMode = False
FilterValue = "10K NTE"
ActiveSheet.Range("A6:AF6").AutoFilter Field:=17, Criteria1:="Y"
Dim lastmonth As String
lastmonth = Format(DateAdd("m", -1, Date), "mmmm")
lastMonthYear = Right(Year(DateAdd("m", -1, Date)), 2)
NotSelect = lastmonth & "'" & lastMonthYear
ActiveSheet.Range("A6:AF6").AutoFilter Field:=32, Criteria1:="*" & FilterValue & "*", Operator:=xlAnd, Criteria2:="<>*" & NotSelect
TLastCell = Range("A6").End(xlDown).Row

For i = 1 To TLastCell
    With Sheets(1).AutoFilter.Range
        'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
        If TLastCell >= 1 Then
            Range("Y" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
        End If
        CurrentCell = ActiveCell.Address
        CelCon = 0
        'Getting Last row of the visible cells after filtering
        Range(CurrentCell).Offset(0, -16).value = CelCon
    End With
Next i

With Sheets(1).AutoFilter.Range
    'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
    Range("Y" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    CurrentCell = ActiveCell.Address
       
   'Getting Last row of the visible cells after filtering
   Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
   Range("Y" & .Offset(1, 0).SpecialCells(xlCellTypeVisible).Row).Select
    Cfrom = ActiveCell.Row
    Addr = Range("AF" & Rows.Count).End(xlUp).Row
    
    Range("Y" & Cfrom).value = "=AF" & Cfrom
    Range("Y" & Cfrom).Copy
    Range("Y" & Cfrom & ":Y" & Addr).SpecialCells(xlCellTypeVisible).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    Range("A6:AF6").AutoFilter Field:=32
    Range("A6:AF6").AutoFilter Field:=17
    Range("Y7:Y" & Addr).Copy
    Range("Y7:Y" & Addr).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
End With
'ActiveSheet.AutoFilterMode = False

FilterValue = "Anniversary Month"
ActiveSheet.Range("A6:AF6").AutoFilter Field:=25, Criteria1:="*" & FilterValue & "*", Operator:=xlAnd
TLastCell = Range("A6").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    With Sheets(1).AutoFilter.Range
        If Rcount >= 1 Then
            'Range("Y" & .Offset(0, 0).SpecialCells(xlCellTypeVisible)(TLastCell).Row).Select
            Range(ActiveCell.Address & ":Y" & TLastCell).SpecialCells(xlCellTypeVisible).Select
        End If
        Selection.Replace What:=FilterValue & ". ", Replacement:="", LookAt:= _
            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="TU to ", Replacement:="", LookAt:= _
            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:=" Using Carryover cost", Replacement:="", LookAt:= _
            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    End With
    
ActiveSheet.AutoFilterMode = False
Range("Y6").Select

'Setting No System Cost No Rebate Paid Customer
ActiveSheet.Range("A6:AF6").AutoFilter Field:=17, Criteria1:="Y"
ActiveSheet.Range("A6:AF6").AutoFilter Field:=22, Criteria1:="<0.1"
ActiveSheet.Range("A6:AF6").AutoFilter Field:=25, Criteria1:=""
Call FilterSetup(32, "No System Cost, no rebate paid", "Y", "No System Cost, no rebate paid")
Call FilterSetup(32, "No System Cost, no rebate paid", "I", 0)
ActiveSheet.AutoFilterMode = False

'Calling Paid on Following Trend
Call PaidOnFollowingTrend


End Sub
Function PaidOnFollowingTrend()

'Paid on Following Trend
Dim TrendCustomer As ArrayList
Set TrendCustomer = New ArrayList

TrendFileLoc = "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\\Macros\Payment Files\Tech Rebate Payments_Consolidated WC.xlsx"
Set TrendFile = Workbooks.Open(TrendFileLoc)
IPCworkingFile.Activate

TrendCustomer.Add "571372"
TrendCustomer.Add "824555"
TrendCustomer.Add "825163"

For Each Cust In TrendCustomer
    ActiveSheet.Range("A6:AF6").AutoFilter Field:=4, Criteria1:=Cust
    With Sheets(1).AutoFilter.Range
        Range("V" & .Offset(1, 0).SpecialCells(xlCellTypeVisible).Row).Select
        SystemCost = ActiveCell.value
        CurrentAddress = ActiveCell.Address
        Compliance = Range(CurrentAddress).Offset(, -5).value
        NP = Range(CurrentAddress).Offset(, 2).value
        CarryOverCost1 = Range(CurrentAddress).Offset(, 11).value
        If SystemCost > "0.1" And Compliance = "Y" Then
            Range("Y" & .Offset(1, 0).SpecialCells(xlCellTypeVisible).Row).value = "Paid following historical trend"
            
            GetCostCell = Range("I" & .Offset(1, 0).SpecialCells(xlCellTypeVisible).Row).Address
            TrendFile.Activate
            Sheets(1).Activate
            LastCellt = Range("C1").End(xlDown).Row
            ActiveSheet.Range("A1:M1").AutoFilter Field:=4, Criteria1:=Cust
            ActiveSheet.Range("A1:M1").AutoFilter Field:=5, Criteria1:=">0.01"
            Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
            Range("E1:E" & LastCellt).SpecialCells(xlCellTypeVisible).End(xlDown).Select
            TrendingSystemCost = ActiveCell.value
            
            IPCworkingFile.Activate
            If NP < TrendingSystemCost And CarryOverCost1 > NP Then
                Range(GetCostCell).value = NP
            Else
                If NP > TrendingSystemCost Then
                    Range(GetCostCell).value = TrendingSystemCost
                End If
            End If
        End If
    End With
    
Next Cust

Call Pay10K
End Function
Function Pay10K()
Dim lastmonth As String
IPCworkingFile.Activate
FilterValue = "10K NTE met"
lastmonth = Format(DateAdd("m", -1, Date), "mmmm")
lastMonthYear = Right(Year(DateAdd("m", 0, Date)), 2)
NotSelect = lastmonth & "'" & lastMonthYear
ActiveSheet.Range("A6:AF6").AutoFilter Field:=1, Criteria1:="*" & "Yes" & "*"
ActiveSheet.Range("A6:AF6").AutoFilter Field:=17, Criteria1:="Y"
ActiveSheet.Range("A6:AF6").AutoFilter Field:=32, Criteria1:="*" & FilterValue & "*", Operator:=xlAnd, Criteria2:="*" & NotSelect & "*"
TLastCell = Range("A6").End(xlDown).Row
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
'TrendFileLoc = "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\\Macros\Payment Files\Tech Rebate Payments_Consolidated WC.xlsx"
'Set TrendFile = Workbooks.Open(TrendFileLoc)
'TrendFile.Activate
'Sheets(1).Activate
ActiveSheet.Range("A1:M1").AutoFilter

IPCworkingFile.Activate
   
    For i = 1 To Rcount
        With Sheets(1).AutoFilter.Range
            'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
            NP = Range("X" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(i).Row).value
            CarryOverCost2 = Range("AE" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(i).Row).value
            LastComment = Range("AF" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(i).Row).value
            CustomerNum = Range("D" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(i).Row).value
            AnnMonth = Range("AD" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(i).Row).value
            '11/1/2023
            LastYr = Right(AnnMonth, 2)
            AnMonth = Left(AnnMonth, 7)
            AnneMonth = AnMonth & LastYr
            'TrendFile.Activate
            'ActiveSheet.Range("A1:M1").AutoFilter Field:=1, Criteria1:=CustomerNum
            'CRcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
            
            'For i = 1 To CRcount
                'RebMonth = Range("H" & i).Value
                'If RebMonth = AnneMonth Then
                    'GotCelladd = ActiveCell.Address
                    
                
            If NP > 10000 And CarryOverCost2 > 10000 Then
                 Range("I" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(i).Row).value = 10000
                 L2 = lastMonthYear + 1
                 LastComment = Replace(LastComment, lastMonthYear, L2)
                 LastComment = Replace(LastComment, "10K NTE met", "10K NTE met using carryover cost")
                 Range("Y" & .Offset(i, 0).SpecialCells(xlCellTypeVisible)(i).Row).value = "Anniversary Month. " & LastComment
                ActiveSheet.Range("A6:AF6").AutoFilter Field:=25, Criteria1:=""
            End If
        End With
    Next i
    ActiveSheet.Range("A6:AF6").AutoFilter Field:=25
       

'ActiveSheet.AutoFilterMode = False
Call ExcludedCustomers
End Function

Function ExcludedCustomers()
Dim ExComments As ArrayList
Set ExComments = New ArrayList

Sheets(1).AutoFilterMode = False

ExComments.Add "Pharmacy Rx is not an approved System for IPC"
ExComments.Add "Per Nicholas, account terminated from IPC"
ExComments.Add "Moved to Liberty"


For Each jef In ExComments
    TLastCell = Range("A6").End(xlDown).Row
    ActiveSheet.Range("A6:AF6").AutoFilter Field:=32, Criteria1:=jef
    With Sheets(1).AutoFilter.Range
        'Setting 0 amount is rebate amount
        Range("I" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
        CurrentCell = ActiveCell.Address
        Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
        If Rcount > 1 Then
            Range(CurrentCell & ":I" & TLastCell).SpecialCells(xlCellTypeVisible).SpecialCells(xlCellTypeVisible).value = 0
        End If
        
        'Setting Comments for Excluded Amounts
        Range("Y" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
        CurrentCell = ActiveCell.Address
        Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
        CellCom = Range(CurrentCell).Offset(0, 7).value
        If Rcount > 1 Then
            Range(CurrentCell & ":Y" & TLastCell).SpecialCells(xlCellTypeVisible).SpecialCells(xlCellTypeVisible).value = CellCom
        End If
        
    End With
    Sheets(1).AutoFilterMode = False
    ActiveSheet.Range("A6:AF6").AutoFilter Field:=17, Criteria1:="Y"
Next jef
Call AcutalPayments
End Function
Function AcutalPayments()

ActiveSheet.Range("A6:AF6").AutoFilter Field:=17, Criteria1:="Y"
ActiveSheet.Range("A6:AF6").AutoFilter Field:=25, Criteria1:=""
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
'For Old
TrendFileLoc = "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\\Macros\Payment Files\Tech Rebate Payments_Consolidated WC.xlsx"
Set TrendFile = Workbooks.Open(TrendFileLoc)
IPCworkingFile.Activate
For i = 1 To Rcount
    With Sheets(1).AutoFilter.Range
        'Selecting First Visible Cell After Putting Filter and getting its curent Cell Address
        Catagory = Range("A" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
        Catagory = Range("A" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).value
        Catagory = Trim(Catagory)
        If Catagory = "Old" Then
            CurrentAdd = ActiveCell.Address
            CarryOverCost3 = Range("AE" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).value
            NP = Range("X" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).value
            SystemCost = Range("V" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).value
            customerID = Range("D" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).value
            
            TrendFile.Activate
            Sheets(1).Activate
            LastCellt = Range("C1").End(xlDown).Row
            ActiveSheet.Range("A1:M1").AutoFilter Field:=4, Criteria1:=customerID
            ActiveSheet.Range("A1:M1").AutoFilter Field:=5, Criteria1:=">0.01"
            Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
            Range("E1:E" & LastCellt).SpecialCells(xlCellTypeVisible).End(xlDown).Select
            NTE = ActiveCell.value
            NTE2 = NTE * 2
            IPCworkingFile.Activate
            
            'Paid on NP
            If NP < NTE And CarryOverCost3 >= NP Then
                Range(CurrentAdd).Offset(0, 8).value = NP
                Range(CurrentAdd).Offset(0, 24).value = "Paid on NP"
            End If
            
            'Paid on NTE
            If SystemCost > NTE And NP > NTE Then
                Range(CurrentAdd).Offset(0, 8).value = NTE
                Range(CurrentAdd).Offset(0, 24).value = "Paid on NTE"
            End If
        
            'TU to NTE using Carry Over Cost
            If SystemCost < NTE And NP > NTE2 And CarryOverCost3 > NTE2 Then
                Range(CurrentAdd).Offset(0, 8).value = NTE
                Range(CurrentAdd).Offset(0, 24).value = "TU to NTE using Carry Over Cost"
            End If
            
            'Paid on system cost as no/low carry over cost
            If SystemCost < NTE And CarryOverCost3 < NTE And NP > SystemCost Then
                Range(CurrentAdd).Offset(0, 8).value = SystemCost
                Range(CurrentAdd).Offset(0, 24).value = "Paid on system cost as no/low carry over cost"
            End If
        Else
        'For New
            CurrentAdd = ActiveCell.Address
            CarryOverCost3 = Range("AE" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).value
            NP = Range("X" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).value
            SystemCost = Range("V" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).value
            NTE = 10000
            If VarType(CarryOverCost3) = vbInteger Then
                CarryOverCost4 = CarryOverCost3 * 1.1
            Else
                CarryOverCost3 = 0
            End If
            
            'CarryOverCost4 = CarryOverCost3 * 1.1
          
            'Negative NP
            If NP < 0 Then
                Range(CurrentAdd).Offset(0, 8).value = 0
                Range(CurrentAdd).Offset(0, 24).value = "Negative NP. No Rebate Paid"
            End If
            
            'Paid on NP
            If NP < NTE And CarryOverCost3 >= NP And NP > 0 Then
                Range(CurrentAdd).Offset(0, 8).value = NP
                Range(CurrentAdd).Offset(0, 24).value = "Paid on NP"
            End If
            
            'Paid on NTE
            If CarryOverCost4 > NTE And NP > NTE Then
                Range(CurrentAdd).Offset(0, 8).value = NTE
                Range(CurrentAdd).Offset(0, 24).value = "Paid on NTE"
            End If
        
            'TU to NTE using Carry Over Cost
            If SystemCost < NTE And NP > NTE And CarryOverCost3 > NTE Then
                Range(CurrentAdd).Offset(0, 8).value = NTE
                Range(CurrentAdd).Offset(0, 24).value = "TU to NTE using Carry Over Cost"
            End If
            
            'Paid on system cost as no/low carry over cost
            If SystemCost < NTE And CarryOverCost3 < NTE And NP > SystemCost Then
                Range(CurrentAdd).Offset(0, 8).value = SystemCost
                Range(CurrentAdd).Offset(0, 24).value = "Paid on system cost as no/low carry over cost"
            End If
        End If
    
        ActiveSheet.Range("A6:AF6").AutoFilter Field:=25, Criteria1:=""
        
    End With
Next i

Call Cleanup
End Function
Function Cleanup()
Application.DisplayAlerts = False
TrendFile.Close
Application.DisplayAlerts = True

ActiveSheet.AutoFilterMode = False

'Assigning Negative NP
ActiveSheet.Range("A6:AF6").AutoFilter Field:=24, Criteria1:="<0"
ActiveSheet.Range("A6:AF6").AutoFilter Field:=25, Criteria1:="<>*" & "Negative NP. No Rebate Paid"
Rcount = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
For i = 1 To Rcount
    With Sheets(1).AutoFilter.Range
        ActiveSheet.Range("A6:AF6").AutoFilter Field:=25, Criteria1:="<>*" & "Negative NP. No Rebate Paid"
        Neg = Range("Y" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
        ActiveCell.value = "Negative NP. No Rebate Paid"
        ActiveCell.Offset(0, -16).value = 0
    End With
Next i
ActiveSheet.AutoFilterMode = False

'Setting No System Cost should have zero in rebate payment
LastCell = Range("A6").End(xlDown).Row
ActiveSheet.Range("A6:AF6").AutoFilter Field:=25, Criteria1:="No System Cost, no rebate paid"
    
    With Sheets(1).AutoFilter.Range
        FirstCell = Range("I" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Address
        Range(FirstCell & ":I" & LastCell).SpecialCells(xlCellTypeVisible).SpecialCells(xlCellTypeBlanks).value = "0.00"
    End With

ActiveSheet.AutoFilterMode = False
Sheets(1).Activate
PMonth = DateAdd("M", -1, Date)
PrevMonth = Format(PMonth, "mm/01/yyyy")
Range("A3").value = PrevMonth
Call RenamePaymentFile
End Function
Function RenamePaymentFile()

IPCworkingFile.Save
IPCworkingFile.Close


Mths = DateAdd("M", -1, Date)
Mth = Format(Mths, "MM")
Mth2 = DateAdd("M", -2, Date)
Mth2 = Format(Mth2, "MM")
CurrentYear = Format(Mths, "YYYY")

L2year = DateAdd("M", -2, Date)
CurrentYear2 = Format(L2year, "YYYY")

Fname = CurrentYear & Mth
Fname2 = CurrentYear2 & Mth2

Name "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\Payment Files\IPC\IPC Payment Summary " & Fname2 & "_Working File.xlsx" As _
   "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\Payment Files\IPC\IPC Payment Summary " & Fname & "_Working File.xlsx"

MsgBox "Completed", vbInformation, "Success"

End Function

