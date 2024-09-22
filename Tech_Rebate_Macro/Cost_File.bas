Attribute VB_Name = "CostFile"
Public WorkingFile As Workbook
Public CostFile As Workbook
Public PaymentFile As Workbook
Public DestinationFolderName As String
Public CostFileDesti As String
Public SourceFile As String
Public ParentPath As String
Sub StartProc()
Dim DestLocation As String
Dim Subject As ArrayList
Dim lastmonth As String
Dim PrevMonth As String
Dim MSearch As String

ParentPath = ThisWorkbook.Path

'Deleting Previous Cost Request Files
On Error GoTo C
Kill ParentPath & "\System Cost\Downloaded_Cost_Files\*.*"
Kill ParentPath & "\System Cost\CostFiles_Template\*.*"
C:


'Global Declration
DestinationFolderName = ParentPath & "\System Cost\Downloaded_Cost_Files"

PrevMonth = MonthName(Month(DateAdd("M", -1, Date)))
RawDate = DateAdd("M", -1, Date)
Prevyear = Right(RawDate, 4)

'Setting up Last Date for subject Line
CurrYear = Year(Date)
Currmonth = Format(DateAdd("M", -1, Now), "MM")
lastmonth = CurrYear & Currmonth

PrevMonth = MonthName(Month(DateAdd("M", -1, Date)))
RawDate = DateAdd("M", -1, Date)
Prevyear = Right(RawDate, 4)
MSearch = PrevMonth & " " & Prevyear

Set Subject = New ArrayList

'SubjLineforLiberty = "RE: Liberty Cost Request - " & Lastmonth
SubjLineforPW = "RE: [EXTERNAL]: PRESCRIBED WELLNESS COST REQUEST - " & MSearch
SubjLineforPW2 = "RE: [EXTERNAL]: RE: PRESCRIBED WELLNESS COST REQUEST - " & MSearch
SubjLineforMPS = "Tech Rebates " & MSearch
SubjLineforParata = "RE: Parata Cost - " & lastmonth

'MsgBox SubjLineforLiberty
'MsgBox SubjLineforPW

'Subject.Add SubjLineforLiberty
Subject.Add SubjLineforPW
Subject.Add SubjLineforPW2
Subject.Add SubjLineforMPS
Subject.Add SubjLineforParata

Dim MailSubject As String
For Each Subs In Subject    ' Iterate through each element.
    MailSubject = Subs
    Call Download_Attachments(DestLocation, MailSubject)
Next


strDir = "\\ddcf2015\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\Macros\System Cost\Downloaded_Cost_Files"
Set FSO = CreateObject("Scripting.FileSystemObject")
Set objFiles = FSO.GetFolder(strDir).Files
lngFileCount = objFiles.Count

If lngFileCount <> 3 Then
    MsgBox "All Cost file was not downloaded. Please check and add missing cost file Manually to the Folder and then Press OK button", vbCritical, "Retry"
    Call TemplateCreation
Else
    MsgBox "Cost Files Downloaded", vbInformation, "Success"
    Call TemplateCreation
    MsgBox "Completed", vbInformation, "Success"
End If

End Sub

Public Function Download_Attachments(DestinationFolderName As String, subjectFilter As String)

On Error GoTo Err_Control
Dim OutlookOpened As Boolean
Dim outApp As Outlook.Application
Dim outNs As Outlook.Namespace
Dim outFolder As Outlook.MAPIFolder
Dim outAttachment As Outlook.Attachment
Dim outItem As Object
'Dim DestinationFolderName As String
Dim saveFolder As String
Dim outMailItem As Outlook.MailItem
Dim inputDate As String, sFolderName As String
Dim FSO As Object
Dim SourceFileName As String, DestinFileName As String
Set FSO = CreateObject("Scripting.FileSystemObject")
Set FSO = CreateObject("Scripting.Filesystemobject")

'sFolderName = Format(Now, "yyyyMMdd")
sMailName = Format(Now, "dd/MM/yyyy")

DestinationFolderName = ParentPath & "\System Cost\Downloaded_Cost_Files"
    
saveFolder = DestinationFolderName

'subjectFilter = "PRESCRIBED WELLNESS COST REQUEST -January 2023"    'REPLACE WORD SUBJECT TO FIND

OutlookOpened = False
On Error Resume Next
Set outApp = GetObject(, "Outlook.Application")
If Err.Number <> 0 Then
    Set outApp = New Outlook.Application
    OutlookOpened = True
End If
On Error GoTo Err_Control

If outApp Is Nothing Then
    MsgBox "Cannot start Outlook.", vbExclamation
    Exit Function
End If

Set outNs = outApp.GetNamespace("MAPI")
Set outFolder = outNs.GetDefaultFolder(olFolderInbox)

If Not outFolder Is Nothing Then
    For Each outItem In outFolder.Items
        If outItem.Class = Outlook.OlObjectClass.olMail Then
            Set outMailItem = outItem
                If InStr(1, outMailItem.Subject, subjectFilter) > 0 Then 'removed the quotes around subjectFilter
                    For Each outAttachment In outMailItem.Attachments
                        If InStr(outAttachment.DisplayName, ".xlsx") _
                            Or InStr(outAttachment.DisplayName, ".xls") _
                            Or InStr(outAttachment.DisplayName, ".xlsm") Then
                                If Dir(saveFolder, vbDirectory) = "" Then FSO.CreateFolder (saveFolder)
                                    outAttachment.SaveAsFile saveFolder & "\" & outAttachment.FileName
                                End If
                            
                        Set outAttachment = Nothing
                    Next
                End If
         End If
    Next
End If


    'SourceFileName = "C:\Users\eo5v4x3\Desktop\Tech Rebates\Payment Files"
    'DestinFileName = saveFolder

    'FSO.MoveFile SourceFileName, DestinFileName


If OutlookOpened Then outApp.Quit
Set outApp = Nothing

Err_Control:
If Err.Number <> 0 Then
    'MsgBox Err.Description
End If

End Function
Function TemplateCreation()
Dim COfile As Object
Dim CostFileArr As ArrayList
Set COfile = CreateObject("Scripting.Filesystemobject")
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False

CurYear = DateAdd("M", -1, Date)
CurrYear = Format(CurYear, "YYYY")
Currmonth = Format(DateAdd("M", -2, Now), "MM")
lastmonth = CurrYear & Currmonth

'Copying CostFile Template of Previous Month
CostFileDesti = ParentPath & "\System Cost\CostFiles_Template\Cost File Template_ " & lastmonth & ".xlsx"
SourceFile = "\\ddcf3007\d-fin1$\Promacct\Customer Rebates - ISMC\Tech Rebate\1. Clean Up\System Costs\Cost File\" & CurrYear & "\Cost File Template_ " & lastmonth & ".xlsx"
COfile.CopyFile SourceFile, CostFileDesti, True

'Rename File
Currmonth = Format(DateAdd("M", -1, Now), "MM")
CurrMth = CurrYear & Currmonth

Name ParentPath & "\System Cost\CostFiles_Template\Cost File Template_ " & lastmonth & ".xlsx" As _
    ParentPath & "\System Cost\CostFiles_Template\Cost File Template_ " & CurrMth & ".xlsx"

Set COfile = CreateObject("Scripting.Filesystemobject")
Set Downfolder = COfile.GetFolder(ParentPath & "\System Cost\Downloaded_Cost_Files")
For Each ObjFile In Downfolder.Files
    Call CostIntoTemplate(ObjFile.Path, ParentPath & "\System Cost\CostFiles_Template\Cost File Template_ " & CurrMth & ".xlsx", "")
Next ObjFile

End Function
Function CostIntoTemplate(costFileName As String, TemplateFile As String, ShName1 As String)

Dim TWf As Workbook
Dim CostFile As Workbook
Dim Regex As Object

Set TWf = Workbooks.Open(TemplateFile)
Set CostFile = Workbooks.Open(costFileName)

If InStr(costFileName, "Liberty") <> 0 Then
    
    TWf.Activate
    Sheets("Liberty").Activate

    lastRow = Range("A1").End(xlDown).Row
    Range("A2:C" & lastRow).Clear

    
    'Copying Account Number
    CostFile.Activate
    Sheets(1).Activate
    lastRow = Range("A3").End(xlDown).Row
    Range("A3:A" & lastRow).Copy
    TWf.Activate
    Sheets("Liberty").Activate
    Range("B2").PasteSpecial xlPasteValues
    
    'Copying Amounts
    CostFile.Activate
    LastColumn = ActiveSheet.Cells(3, ActiveSheet.Columns.Count).End(xlToLeft).Address
    Set Regex = CreateObject("VBScript.RegExp")
    Regex.Global = True
    Regex.Pattern = "[0-9]"
    OutStr = Regex.Replace(LastColumn, "")
    OutStr = Replace(OutStr, "$", "")
    Range(OutStr & "3:" & OutStr & lastRow).Copy
    TWf.Activate
    Sheets("Liberty").Activate
    Range("C2").PasteSpecial xlPasteValues
    
    'Setting Date in A column of Prev Month
    CurrYear = Year(Date)
    Currmonth = Format(DateAdd("M", -1, Now), "MM")
    lastmonth = CurrYear & Currmonth
    Range("A2:A" & lastRow).value = lastmonth
    
ElseIf InStr(costFileName, "PW") <> 0 Then
    Sheets(2).Activate
    ActiveSheet.PivotTables("PivotTable2").RefreshTable
    
    'Copying Account Number and Amount
    Sheets(2).Activate
    lastRow = Range("A4").End(xlDown).Row
    lastRow = lastRow - 1
    Range("A4:B" & lastRow).Copy
    TWf.Activate
    Sheets("Prescribed Wellness ").Activate
    Range("B2").PasteSpecial xlPasteValues
    
    'Setting Date in A column of Prev Month
    CurrYear = Year(Date)
    Currmonth = Format(DateAdd("M", -1, Now), "MM")
    lastmonth = CurrYear & Currmonth
    lastRow = Range("B2").End(xlDown).Row
    Range("A2:A" & lastRow).value = lastmonth

ElseIf InStr(costFileName, "Parata") <> 0 Then
    
    'Pivot Updated Source
    Dim pt As PivotTable
    Dim newSource As String
    Dim LastCell
    Sheets(1).Activate
    LastCell = Range("C2").End(xlDown).Row
      
    'Calling function to fix missing costs in main sheet
    Call CostsetupParataOnly
    Sheets(2).Activate
    Set pt = ActiveSheet.PivotTables("PivotTable1")
    
    newSource = "'Parata Cost'!$A$1:$AF$" & LastCell
    pt.ChangePivotCache ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=newSource)
    
    ActiveSheet.PivotTables("PivotTable1").RefreshTable
    
    'Copying Account Number and Amount
    lastRow = Range("A4").End(xlDown).Row
    lastRow = lastRow - 1
    Range("A4:B" & lastRow).Copy
    TWf.Activate
    Sheets("Parata ").Activate
    Range("B2").PasteSpecial xlPasteValues
    
    'Setting Date in A column of Prev Month
    CurrYear = Year(Date)
    Currmonth = Format(DateAdd("M", -1, Now), "MM")
    lastmonth = CurrYear & Currmonth
    lastRow = Range("B2").End(xlDown).Row
    Range("A2:A" & lastRow).value = lastmonth
    
ElseIf InStr(costFileName, "Tech Rebates") <> 0 Then

    'clearing Sheet
    TWf.Activate
    Sheets("MPS").Activate
    LastCell = Range("B2").End(xlDown).Row
    Range("A2:R" & LastCell).Clear
    
    'Copying Data From Cost File
    CostFile.Activate
    Sheets(1).Activate
    ActiveSheet.AutoFilterMode = False
    LastCell = Range("B2").End(xlDown).Row
    Range("A2:Q" & LastCell).Copy
    
    'Pasting Data in Consolidated File
    TWf.Activate
    Sheets("MPS").Activate
    Range("B2").PasteSpecial xlPasteValues
    
    'setting up File
    Dim i As Integer
    
    For i = 1 To LastCell
    Convalue = Range("P" & i).value
    If Convalue = "M2" Then
        Range("A" & i).value = "Enterprise Rx"
    ElseIf Convalue = "M0" Then
        Range("A" & i).value = "POS"
    Else
        Range("A" & i).value = "POS"
    End If
    Next i
    
    'MPS Pivot Update
    Dim pvt As PivotTable
    Dim nSource As String
    Sheets("MPS").Activate
    LastCell = Range("B2").End(xlDown).Row
    
    Sheets("Sheet1").Activate
    Set pvt = ActiveSheet.PivotTables("PivotTable1")
    nSource = "'MPS'!$A$1:$R$" & LastCell
    pvt.ChangePivotCache ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=nSource)
    
    ActiveSheet.PivotTables("PivotTable1").RefreshTable
    
    
End If
CostFile.Save
CostFile.Close
End Function

Sub CostsetupParataOnly()

ActualLastCell = Range("A1").End(xlDown).Row
CostLastCell = Range("AA1").End(xlDown).Row

DiffCell = ActualLastCell - CostLastCell
Range("AA" & ActualLastCell).Select

For i = 1 To DiffCell
    ActiveCell.value = "=Sum(" & ActiveCell.Offset(0, -1) + ActiveCell.Offset(0, -2) & ")"
    ActiveCell.Offset(-1, 0).Select
Next i
End Sub

