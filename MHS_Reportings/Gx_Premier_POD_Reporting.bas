Attribute VB_Name = "Gx_Premier_POD_Reporting"
Sub ReportCreation()

Version_Control = "867010710"
Test_vs_Prd = "P"
Vendor_Name = "McKesson"
Premier_Agreement = "PPPW18MBF01 - Amendment 1"

'Copying Format File and Creating New Month File
Mnth = DateAdd("M", -1, Date)
Mth = Format(Mnth, "mm")
yr = Format(Mnth, "YYYY")

Dim FormatFile As Object
Set FormatFile = CreateObject("Scripting.Filesystemobject")
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False
SourceFile = "C:\Users\eo5v4x3\Desktop\MHS Reportings\Gx_Premimer_POD_Format_File.xls" 'Variable
FormatFileDestination = "C:\Users\eo5v4x3\Desktop\MHS Reportings\Reports\Gx\" & yr & Mth & " Premier OS Admin Fee - New PODs.xls"  'Variable
FormatFile.CopyFile SourceFile, FormatFileDestination, True

'Reporting File Preparation
Dim ReportingFile As Workbook
Set ReportingFile = Workbooks.Open(FormatFileDestination)
ReportingFile.Activate
Sheets("Premier Template").Activate
Lrow = Range("A1").End(xlDown).Row
Range("A2:AN" & Lrow).Value = ""

'Openning BW
Dim BWFile As Workbook
BWFilePath = "C:\Users\eo5v4x3\Desktop\MHS Reportings\BW Queries\Gx_Premier_POD_BW_Query_Long_Report.xlsx" 'Variable
Set BWFile = Workbooks.Open(BWFilePath)
BWFile.Activate
Sheets("Table").Activate
BWLastRow = Range("J15").End(xlDown).Row
BWLastRow = BWLastRow - 1

Dim BWColList As ArrayList
Dim RPColList As ArrayList
Set BWColList = New ArrayList
Set RPColList = New ArrayList

'Adding BW File Column Names to a List
BWColList.Add "J" 'Customer Number
BWColList.Add "BQ" 'Sales Amount
BWColList.Add "BS" 'Rebate Amount
BWColList.Add "AK" 'DEA Number
BWColList.Add "K" 'Facility Name

'Adding Reporting File Column Names to a List
RPColList.Add "AN" 'Customer Number
RPColList.Add "AF" 'Sales Amount
RPColList.Add "AL" 'Rebate Amount
RPColList.Add "J" 'DEA Number
RPColList.Add "O" 'Facility Name

'Copying Data
For i = 0 To BWColList.Count - 1
    BWFile.Activate
    Range(BWColList(i) & 16 & ":" & BWColList(i) & BWLastRow).Copy
    ReportingFile.Activate
    Range(RPColList(i) & 2).PasteSpecial xlPasteValues
Next i

BWFile.Close

'Getting Customer Addresses from External Rebate Report
Dim ExtRebFile As Workbook
ExtRbtFilePath = "C:\Users\eo5v4x3\Desktop\MHS Reportings\Required Files\External Rebate Reports\72268_Ext_Rbt.XLSX"
Set ExtRebFile = Workbooks.Open(ExtRbtFilePath)
ExtRebFile.Activate
Sheets(1).Activate
ExLastRow = Range("A1").End(xlDown).Row
With Range("A2:A" & ExLastRow)
    .NumberFormat = "General"
    .Value = .Value
End With

ReportingFile.Activate
RPLastRow = Range("AN1").End(xlDown).Row

'Changing Customer Numbers and Zip Codes from Text string to Number format
With Range("AN2:AN" & RPLastRow)
    .NumberFormat = "General"
    .Value = .Value
End With

With Range("T2:T" & RPLastRow)
    .NumberFormat = "General"
    .Value = .Value
End With

'Getting Address
Dim Addr_Range As Range
'Street Name
For Each Addr_Range In Range("P2:P" & RPLastRow)
    Addr_Range.Value = "=XLookup(" & Addr_Range.Offset(0, 24).Value & ",[72268_Ext_Rbt.XLSX]Sheet1!$A1:$A65536,[72268_Ext_Rbt.XLSX]Sheet1!$D1:$D65536)"
Next Addr_Range

'City
For Each Addr_Range In Range("R2:R" & RPLastRow)
    Addr_Range.Value = "=XLookup(" & Addr_Range.Offset(0, 22).Value & ",[72268_Ext_Rbt.XLSX]Sheet1!$A1:$A65536,[72268_Ext_Rbt.XLSX]Sheet1!$E1:$E65536)"
Next Addr_Range

'State Name
For Each Addr_Range In Range("S2:S" & RPLastRow)
    Addr_Range.Value = "=XLookup(" & Addr_Range.Offset(0, 21).Value & ", [72268_Ext_Rbt.XLSX]Sheet1!$A1:$A65536, [72268_Ext_Rbt.XLSX]Sheet1!$F1:$F65536)"
Next Addr_Range

'Zip Code
For Each Addr_Range In Range("T2:T" & RPLastRow)
    Addr_Range.Value = "=XLookup(" & Addr_Range.Offset(0, 20).Value & ", [72268_Ext_Rbt.XLSX]Sheet1!$A1:$A65536, [72268_Ext_Rbt.XLSX]Sheet1!$G1:$G65536)"
Next Addr_Range

Range("P2:T" & RPLastRow).Copy
Range("P2:T" & RPLastRow).PasteSpecial xlPasteValues
Application.CutCopyMode = False
Range("P2").Select

ExtRebFile.Close

'Setting up text in  Fixed Columns
ReportingFile.Activate
Range("A2:A" & RPLastRow).Value = Version_Control
Range("B2:B" & RPLastRow).Value = Test_vs_Prd
Range("C2:C" & RPLastRow).Value = Vendor_Name
Range("L2:L" & RPLastRow).Value = Premier_Agreement


Start_Date = DateAdd("M", -1, Date)
Start_Date01 = Format(Start_Date, "YYYYMM01")

Dim currentDate As Date
Dim lastDay As Date
' Get the current date
currentDate = Date

' Calculate the first day of the current month
Dim firstDayOfCurrentMonth As Date
firstDayOfCurrentMonth = DateSerial(Year(currentDate), Month(currentDate), 1)

' Calculate the last day of the second-to-last month
lastDay = firstDayOfCurrentMonth - 1
LDate = Format(lastDay, "yyyyMMdd")

'Setting Dates
Range("M2:M" & RPLastRow).Value = Start_Date01
Range("N2:N" & RPLastRow).Value = LDate


'Copying Format of first row on all Data
Range("A2:AN2").Copy
Range("A3:AN" & RPLastRow).PasteSpecial xlPasteFormats
Range("A1").Select

'Save and Exit
ReportingFile.Save
ReportingFile.Close

'Completion Msg
MsgBox "Completed", vbInformation, "Success"
End Sub

