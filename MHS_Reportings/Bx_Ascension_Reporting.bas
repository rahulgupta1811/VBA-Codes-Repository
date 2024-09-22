Attribute VB_Name = "Bx_Ascension_Reporting"
Sub AscReporting()

'Copying Format File and Creating New Month File
Mnth = DateAdd("M", -1, Date)
Mth = Format(Mnth, "mm")
yr = Format(Mnth, "YY")

Dim FormatFile As Object
Set FormatFile = CreateObject("Scripting.Filesystemobject")
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False
SourceFile = "C:\Users\eo5v4x3\Desktop\MHS Reportings\Asc FormatFile.xlsx" 'Variable
FormatFileDestination = "C:\Users\eo5v4x3\Desktop\MHS Reportings\Reports\Bx\" & Mth & yr & " McKesson Ascension Admin Fee Report" & ".xlsx" 'Variable
FormatFile.CopyFile SourceFile, FormatFileDestination, True

'Reporting File Preparation
Dim ReportingFile As Workbook
Set ReportingFile = Workbooks.Open(FormatFileDestination)
ReportingFile.Activate
Sheets("Admin Fee").Activate
Lrow = Range("B1").End(xlDown).Row
Range("A2:W" & Lrow).Value = ""

Dim BWFile As Workbook
BWFilePath = "C:\Users\eo5v4x3\Desktop\MHS Reportings\BW Queries\Bx_Ascension_BW_Query.xls" 'Variable
Set BWFile = Workbooks.Open(BWFilePath)
BWFile.Activate
Sheets("Table").Activate
BWLastRow = Range("F15").End(xlDown).Row
BWLastRow = BWLastRow - 1

Dim BWColList As ArrayList
Dim RPColList As ArrayList
Set BWColList = New ArrayList
Set RPColList = New ArrayList

'Adding BW File Column Names to a List
BWColList.Add "H" 'Customer Number
BWColList.Add "S" 'Sales Amount
BWColList.Add "T" 'Rebate Percentage
BWColList.Add "V" 'Rebate Amount
BWColList.Add "N" 'DEA Number
BWColList.Add "I" 'Facility Name
BWColList.Add "J" 'Facility Address
BWColList.Add "K" 'Facility City
BWColList.Add "L" 'Facility State
BWColList.Add "M" 'Facility Zip Code
BWColList.Add "G" 'National Group Code

'Adding Reporting File Column Names to a List
RPColList.Add "B" 'Customer Number
RPColList.Add "U" 'Sales Amount
RPColList.Add "V" 'Rebate Percentage
RPColList.Add "W" 'Rebate Amount
RPColList.Add "R" 'DEA Number
RPColList.Add "D" 'Facility Name
RPColList.Add "E" 'Facility Address
RPColList.Add "F" 'Facility City
RPColList.Add "G" 'Facility State
RPColList.Add "H" 'Facility Zip Code
RPColList.Add "A" 'National Group Code


'Copying Data
For i = 0 To BWColList.Count - 1
    BWFile.Activate
    Range(BWColList(i) & 16 & ":" & BWColList(i) & BWLastRow).Copy
    ReportingFile.Activate
    Range(RPColList(i) & 2).PasteSpecial xlPasteValues
Next i

BWFile.Close

'Setting up text in  Fixed Columns
ReportingFile.Activate
RPLastRow = Range("B1").End(xlDown).Row
Range("N2:O" & RPLastRow).Value = "Mckesson"
Range("Q2:Q" & RPLastRow).Value = "MCKES-0001974"

Start_Date = DateAdd("M", -1, Date)
Start_Date01 = Format(Start_Date, "YYYYMM")

'Setting Dates
Range("T2:T" & RPLastRow).Value = Start_Date01
Range("V2:V" & RPLastRow).Value = "0.60%"
TotalValRow = RPLastRow + 1
Range("U" & TotalValRow).Value = "=SUM(U2:U" & RPLastRow & ")"
Range("W" & TotalValRow).Value = "=SUM(W2:W" & RPLastRow & ")"


'Setting Customer NUmber with starting with 20
For i = 2 To RPLastRow
    Range("C" & i).Value = 20000000 + Range("B" & i).Value
Next i

'Changing Customer Numbers and Zip Codes from Text string to Number format
With Range("A2:B" & RPLastRow)
    .NumberFormat = "General"
    .Value = .Value
End With

With Range("H2:H" & RPLastRow)
    .NumberFormat = "General"
    .Value = .Value
End With

'Copying Format of first row on all Data
Range("A2:W2").Copy
Range("A3:W" & RPLastRow).PasteSpecial xlPasteFormats
Range("A1").Select

'Save and Exit
ReportingFile.Save
ReportingFile.Close

'Completion Msg
MsgBox "Completed", vbInformation, "Success"
End Sub
