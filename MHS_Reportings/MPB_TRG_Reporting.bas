Attribute VB_Name = "MPB_TRG_Reporting"
Sub ReportCreation()

'Copying Format File and Creating New Month File
Mnth = DateAdd("M", -1, Date)
Mth = Format(Mnth, "mmmm")
yr = Format(Mnth, "YYYY")

Dim FormatFile As Object
Set FormatFile = CreateObject("Scripting.Filesystemobject")
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False
SourceFile = "C:\Users\eo5v4x3\Desktop\MHS Reportings\MPB TRG Format_File.xlsx" 'Variable
FormatFileDestination = "C:\Users\eo5v4x3\Desktop\MHS Reportings\Reports\MPB\MPB TRG Admin Fee Report_" & Mth & "'" & yr & ".xlsx" 'Variable
FormatFile.CopyFile SourceFile, FormatFileDestination, True

'Reporting File Preparation
Dim ReportingFile As Workbook
Set ReportingFile = Workbooks.Open(FormatFileDestination)
ReportingFile.Activate
Sheets("Admin Fee").Activate
Lrow = Range("A1").End(xlDown).Row
Range("A2:O" & Lrow).Value = ""

'Openning BW
Dim BWFile As Workbook
BWFilePath = "C:\Users\eo5v4x3\Desktop\MHS Reportings\BW Queries\MPB_TRG MPB Long Report.xlsx" 'Variable
Set BWFile = Workbooks.Open(BWFilePath)
BWFile.Activate
Sheets("Table").Activate
BWLastRow = Range("J15").End(xlDown).Row

Dim BWColList As ArrayList
Dim RPColList As ArrayList
Set BWColList = New ArrayList
Set RPColList = New ArrayList

'Adding BW File Column Names to a List
BWColList.Add "J" 'Customer Number
BWColList.Add "J" 'Customer Number
BWColList.Add "BR" 'Sales Amount
BWColList.Add "BT" 'Rebate Amount
BWColList.Add "AL" 'DEA Number
BWColList.Add "K" 'Facility Name
BWColList.Add "M" 'National Group

'Adding Reporting File Column Names to a List
RPColList.Add "B" 'Customer Number
RPColList.Add "C" 'Customer Number
RPColList.Add "M" 'Sales Amount
RPColList.Add "O" 'Rebate Amount
RPColList.Add "J" 'DEA Number
RPColList.Add "D" 'Facility Name
RPColList.Add "A" 'National Group

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
ExtRbtFilePath = "C:\Users\eo5v4x3\Desktop\MHS Reportings\Required Files\External Rebate Reports\85876_Ext_Rbt.XLSX"
Set ExtRebFile = Workbooks.Open(ExtRbtFilePath)
ExtRebFile.Activate
Sheets(1).Activate
ExLastRow = Range("A1").End(xlDown).Row

With Range("A2:A" & ExLastRow)
    .NumberFormat = "General"
    .Value = .Value
End With

ReportingFile.Activate
RPLastRow = Range("A1").End(xlDown).Row

'Changing Customer Numbers and Zip Codes from Text string to Number format
With Range("A2:C" & RPLastRow)
    .NumberFormat = "General"
    .Value = .Value
End With

'Getting Address
Dim Addr_Range As Range
'Street Name
For Each Addr_Range In Range("E2:E" & RPLastRow)
    Addr_Range.Value = "=VLOOKUP(" & Addr_Range.Offset(0, -3).Value & ",[85876_Ext_Rbt.XLSX]Sheet1!$A:$G,4,0)"
Next Addr_Range

'City
For Each Addr_Range In Range("F2:F" & RPLastRow)
    Addr_Range.Value = "=VLOOKUP(" & Addr_Range.Offset(0, -4).Value & ",[85876_Ext_Rbt.XLSX]Sheet1!$A:$G,5,0)"
Next Addr_Range

'State Code
For Each Addr_Range In Range("G2:G" & RPLastRow)
    Addr_Range.Value = "=VLOOKUP(" & Addr_Range.Offset(0, -5).Value & ",[85876_Ext_Rbt.XLSX]Sheet1!$A:$G,6,0)"
Next Addr_Range

'Zip Code
For Each Addr_Range In Range("H2:H" & RPLastRow)
    Addr_Range.Value = "=VLOOKUP(" & Addr_Range.Offset(0, -6).Value & ",[85876_Ext_Rbt.XLSX]Sheet1!$A:$G,7,0)"
Next Addr_Range

Range("E2:H" & RPLastRow).Copy
Range("E2:H" & RPLastRow).PasteSpecial xlPasteValues
Application.CutCopyMode = False
Range("A2").Select

ExtRebFile.Close

'Setting up text in  Fixed Columns
ReportingFile.Activate
Range("I2:I" & RPLastRow).Value = "MCKES-0073766"
Range("N2:N" & RPLastRow).Value = "2.85%"

Start_Date = DateAdd("M", -1, Date)
Start_Date01 = Format(Start_Date, "YYYYMM")

'Setting Dates
Range("L2:L" & RPLastRow).Value = Start_Date01

'Copying Format of first row on all Data
Range("A2:V2").Copy
Range("A3:V" & RPLastRow).PasteSpecial xlPasteFormats
Range("A1").Select

'Save and Exit
ReportingFile.Save
ReportingFile.Close

'Completion Msg
MsgBox "Completed", vbInformation, "Success"
End Sub

