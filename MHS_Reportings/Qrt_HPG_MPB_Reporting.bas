Attribute VB_Name = "Qrt_HPG_MPB_Reporting"
Sub ReportCreation()

'Copying Format File and Creating New Month File
Mnth = DateAdd("M", -1, Date)
Mth = Format(Mnth, "mm")
yr = Format(Mnth, "YY")

Dim FormatFile As Object
Set FormatFile = CreateObject("Scripting.Filesystemobject")
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False
SourceFile = "C:\Users\eo5v4x3\Desktop\MHS Reportings\HPG Contract #78804 Format_File.xlsx" 'Variable
FormatFileDestination = "C:\Users\eo5v4x3\Desktop\MHS Reportings\Reports\Bx\" & Mth & yr & " HPG Admin Fee Report_HPG Contract #78804.xlsx" 'Variable
FormatFile.CopyFile SourceFile, FormatFileDestination, True

'Reporting File Preparation
Dim ReportingFile As Workbook
Set ReportingFile = Workbooks.Open(FormatFileDestination)
ReportingFile.Activate
Sheets("HPG Admin Fee #78804").Activate
Lrow = Range("B7").End(xlDown).Row
Range("A8:M" & Lrow).Value = ""

'Openning BW
Dim BWFile As Workbook
BWFilePath = "C:\Users\eo5v4x3\Desktop\MHS Reportings\BW Queries\BW-HPG 0.60% 78804 (Qtr).xlsx" 'Variable
Set BWFile = Workbooks.Open(BWFilePath)
BWFile.Activate
Sheets("Table").Activate
BWLastRow = Range("G15").End(xlDown).Row
BWLastRow = BWLastRow - 1

'Coying Data from BW
'Copying Rebate Agreement Number, Sold to Party, Name, StreetName, City, Region, Postal Code and DEA Number
Range("F16:M" & BWLastRow).Copy
ReportingFile.Activate
Range("A8").PasteSpecial xlPasteValues

'Copying Net Sales
BWFile.Activate
Range("R16:R" & BWLastRow).Copy
ReportingFile.Activate
Range("I8").PasteSpecial xlPasteValues

'Copying Add Sub/Net CNT and Rebateable Sales
BWFile.Activate
Range("T16:U" & BWLastRow).Copy
ReportingFile.Activate
Range("J8").PasteSpecial xlPasteValues

'Copying Rebate Amount
BWFile.Activate
Range("Y16:Y" & BWLastRow).Copy
ReportingFile.Activate
Range("M8").PasteSpecial xlPasteValues

BWFile.Close

ReportingFile.Activate
RPLastRow = Range("B7").End(xlDown).Row

'Changing Customer Numbers and Zip Codes from Text string to Number format
With Range("A8:B" & RPLastRow)
    .NumberFormat = "General"
    .Value = .Value
End With

With Range("G8:G" & RPLastRow)
    .NumberFormat = "General"
    .Value = .Value
End With

'Setting up text in  Fixed Columns
ReportingFile.Activate
yrr = Format(Mnth, "YYYY")
calendar_quarter = Choose(Month(Mnth), 1, 1, 1, 2, 2, 2, 3, 3, 3, 4, 4, 4)
Range("A4").Value = "Q" & calendar_quarter & yrr
Range("L8:L" & RPLastRow).Value = "0.60%"

'Copying Format of first row on all Data
Range("A8:M8").Copy
Range("A9:M" & RPLastRow).PasteSpecial xlPasteFormats
Range("A7").Select
Application.CutCopyMode = False

Range("M" & RPLastRow).Offset(2, 0).Value = "=Sum(M8:M" & RPLastRow & ")"
Range("K" & RPLastRow).Offset(2, 0).Value = "=Sum(K8:K" & RPLastRow & ")"
Range("J" & RPLastRow).Offset(2, 0).Value = "=Sum(J8:J" & RPLastRow & ")"
Range("I" & RPLastRow).Offset(2, 0).Value = "=Sum(I8:I" & RPLastRow & ")"

'Save and Exit
ReportingFile.Save
ReportingFile.Close

'Completion Msg
MsgBox "Completed", vbInformation, "Success"
End Sub
