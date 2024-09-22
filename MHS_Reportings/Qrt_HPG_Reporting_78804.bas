Attribute VB_Name = "Qrt_HPG_Reporting_78804"
Sub ReportCreation()

'Copying Format File and Creating New Month File
Mnth = DateAdd("M", -1, Date)
Mth = Format(Mnth, "mm")
yr = Format(Mnth, "YYYY")

Dim FormatFile As Object
Set FormatFile = CreateObject("Scripting.Filesystemobject")
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False
SourceFile = "C:\Users\eo5v4x3\Desktop\MHS Reportings\HPG Contract #78804 Format_File.xlsx" 'Variable
FormatFileDestination = "C:\Users\eo5v4x3\Desktop\MHS Reportings\Reports\MPB\MPB HPG Q3_" & yr & Mth & ".xlsx" 'Variable
FormatFile.CopyFile SourceFile, FormatFileDestination, True

'Reporting File Preparation
Dim ReportingFile As Workbook
Set ReportingFile = Workbooks.Open(FormatFileDestination)
ReportingFile.Activate
Sheets("MPB HPG").Activate
Lrow = Range("A1").End(xlDown).Row
Range("A2:K" & Lrow).Value = ""

'Openning BW
Dim BWFile As Workbook
BWFilePath = "C:\Users\eo5v4x3\Desktop\MHS Reportings\BW Queries\Qrt_HPG MPB_Long Report.xlsx" 'Variable
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
BWColList.Add "BQ" 'Sales Amount
BWColList.Add "BS" 'Rebate Amount
BWColList.Add "AK" 'DEA Number
BWColList.Add "K" 'Facility Name

'Adding Reporting File Column Names to a List
RPColList.Add "B" 'Customer Number
RPColList.Add "I" 'Sales Amount
RPColList.Add "K" 'Rebate Amount
RPColList.Add "H" 'DEA Number
RPColList.Add "C" 'Facility Name

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
ExtRbtFilePath = "C:\Users\eo5v4x3\Desktop\MHS Reportings\Required Files\External Rebate Reports\92607_Ext_Rbt.XLSX"
Set ExtRebFile = Workbooks.Open(ExtRbtFilePath)
ExtRebFile.Activate
Sheets(1).Activate
ExLastRow = Range("A1").End(xlDown).Row

With Range("A2:A" & ExLastRow)
    .NumberFormat = "General"
    .Value = .Value
End With

ReportingFile.Activate
RPLastRow = Range("B1").End(xlDown).Row

'Changing Customer Numbers and Zip Codes from Text string to Number format
With Range("B2:B" & RPLastRow)
    .NumberFormat = "General"
    .Value = .Value
End With

'Getting Address
Dim Addr_Range As Range
'Street Name
For Each Addr_Range In Range("D2:D" & RPLastRow)
    Addr_Range.Value = "=VLOOKUP(" & Addr_Range.Offset(0, -2).Value & ",[92607_Ext_Rbt.XLSX]Sheet1!$A:$G,4,0)"
Next Addr_Range

'City
For Each Addr_Range In Range("E2:E" & RPLastRow)
    Addr_Range.Value = "=VLOOKUP(" & Addr_Range.Offset(0, -3).Value & ",[92607_Ext_Rbt.XLSX]Sheet1!$A:$G,5,0)"
Next Addr_Range

'State Code
For Each Addr_Range In Range("F2:F" & RPLastRow)
    Addr_Range.Value = "=VLOOKUP(" & Addr_Range.Offset(0, -4).Value & ",[92607_Ext_Rbt.XLSX]Sheet1!$A:$G,6,0)"
Next Addr_Range

'Zip Code
For Each Addr_Range In Range("G2:G" & RPLastRow)
    Addr_Range.Value = "=VLOOKUP(" & Addr_Range.Offset(0, -5).Value & ",[92607_Ext_Rbt.XLSX]Sheet1!$A:$G,7,0)"
Next Addr_Range

Range("D2:G" & RPLastRow).Copy
Range("D2:G" & RPLastRow).PasteSpecial xlPasteValues
Application.CutCopyMode = False
Range("A2").Select

ExtRebFile.Close

'Setting up text in  Fixed Columns
ReportingFile.Activate
Range("A2:A" & RPLastRow).Value = "92607"
Range("J2:J" & RPLastRow).Value = "1.00%"

'Copying Format of first row on all Data
Range("A2:K2").Copy
Range("A3:K" & RPLastRow).PasteSpecial xlPasteFormats
Range("A1").Select

'Save and Exit
ReportingFile.Save
ReportingFile.Close

'Completion Msg
MsgBox "Completed", vbInformation, "Success"
End Sub

