Attribute VB_Name = "BX_Premier_Acurity_Reporting"
Public Function ReportCreation(ReportingFilePath, BWFilePath, Version_Control, Test_vs_Prd, Vendor_Name)

'Copying Format File and Creating New Month File
Mnth = DateAdd("M", -1, Date)
Mth = Format(Mnth, "mm")
yr = Format(Mnth, "YY")

Dim FormatFile As Object
Set FormatFile = CreateObject("Scripting.Filesystemobject")
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False
SourceFile = "C:\Users\eo5v4x3\Desktop\MHS Reportings\Format File.xlsx" 'Variable
FormatFileDestination = "C:\Users\eo5v4x3\Desktop\MHS Reportings\Reports\Bx\" & Mth & yr & ReportingFilePath & ".xlsx" 'Variable
FormatFile.CopyFile SourceFile, FormatFileDestination, True

'Reporting File Preparation
Dim ReportingFile As Workbook
Set ReportingFile = Workbooks.Open(FormatFileDestination)
ReportingFile.Activate
Sheets("Acurity Template").Activate
Lrow = Range("A1").End(xlDown).Row
Range("A2:AN" & Lrow).Value = ""

'Openning BW
Dim BWFile As Workbook
'BWFilePath = "C:\Users\eo5v4x3\Desktop\MHS Reportings\BW Queries\" & BWFile 'Variable
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
BWColList.Add "G" 'Customer Number
If InStr(1, ReportingFilePath, "Premier") Then
    BWColList.Add "P" 'Sales Amount
Else
    BWColList.Add "O" 'Sales Amount
End If
BWColList.Add "R" 'Rebate Amount
BWColList.Add "M" 'DEA Number
BWColList.Add "H" 'Facility Name
BWColList.Add "I" 'Facility Address
BWColList.Add "J" 'Facility City
BWColList.Add "K" 'Facility State
BWColList.Add "L" 'Facility Zip Code

'Adding Reporting File Column Names to a List
RPColList.Add "AN" 'Customer Number
RPColList.Add "AF" 'Sales Amount
RPColList.Add "AL" 'Rebate Amount
RPColList.Add "J" 'DEA Number
RPColList.Add "O" 'Facility Name
RPColList.Add "P" 'Facility Address
RPColList.Add "R" 'Facility City
RPColList.Add "S" 'Facility State
RPColList.Add "T" 'Facility Zip Code

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
RPLastRow = Range("AN1").End(xlDown).Row
Range("A2:A" & RPLastRow).Value = Version_Control
Range("B2:B" & RPLastRow).Value = Test_vs_Prd
Range("C2:C" & RPLastRow).Value = Vendor_Name

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

'Changing Customer Numbers and Zip Codes from Text string to Number format
With Range("AN2:AN" & RPLastRow)
    .NumberFormat = "General"
    .Value = .Value
End With

With Range("T2:T" & RPLastRow)
    .NumberFormat = "General"
    .Value = .Value
End With

If InStr(1, ReportingFilePath, "Premier") Then
    Range("L2:L" & RPLastRow).Value = "PPPW18MBF01"
    ActiveSheet.Name = "Premier Template"
    Range("L1").Value = "Premier_Agreement_#"
Else
    Range("L2:L" & RPLastRow).Value = "ACU-PH-29"
End If


'Copying Format of first row on all Data
Range("A2:AN2").Copy
Range("A3:AN" & RPLastRow).PasteSpecial xlPasteFormats
Range("A1").Select

'Save and Exit
ReportingFile.Save
ReportingFile.Close

'Completion Msg
MsgBox "Completed", vbInformation, "Success"

End Function

