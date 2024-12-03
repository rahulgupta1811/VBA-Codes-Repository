Attribute VB_Name = "Data_Creation"
Sub Data_Generate()

Dim ExcelFilePath As String
Dim AccessFilePath As String
Dim AccessApp As Object
Dim ParenthPath As String
Dim ExcelWB As Workbook

User = Environ("USERNAME")
ParentPath = ThisWorkbook.Path
ParentPath = Replace(ParentPath, "https://mydrive.merck.com/personal/" & User & "_merck_com/Documents", "C:\Users\" & User & "\OneDrive - Merck Sharp & Dohme LLC")
ParentPath = Replace(ParentPath, "/", "\")

ExcelFilePath = "C:\Users\dasakas\Downloads\Spend_Issues_Export_10_17_2024_08100027.xlsx"
'ExcelFilePath = UserForm2.TextBox1.Value
AccessFilePath = ParentPath & "\Dashboard.accdb"
AccessTable = "Spend_Issue"

'Adding Fresh Excel Data to Access Database
Set AccessApp = CreateObject("Access.Application")
AccessApp.OpenCurrentDatabase AccessFilePath

AccessApp.DoCmd.TransferSpreadsheet _
        TransferType:=acImport, _
        SpreadsheetType:=acSpreadSheetTypeExcel12, _
        tableName:=AccessTable, _
        Filename:=ExcelFilePath, _
        HasFieldNames:=True, _
        Range:="GeneralSpend Upload Template!"
AccessApp.Quit

'Adding Usernames, Aging to the database
Call Queries(AccessFilePath)

'Fetching Data from database
Call FetchDataFromDB(AccessFilePath)

End Sub

Private Sub FetchDataFromDB(AccessDataBase)
Dim conn As Object
Dim strSQL As String
Dim rs As Object
Dim DBPath As String
Dim WhereClause As String

'Sheets("Sheet2").Activate
strSQL = "SELECT * FROM Spend_Issue"
DBPath = AccessDataBase

Set conn = CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath

Set rs = CreateObject("ADODB.Recordset")
rs.Open strSQL, conn

Sheets("Sheet2").UsedRange.Clear

For i = 1 To rs.Fields.Count
    Sheets("Sheet2").Cells(1, i).Value = rs.Fields(i - 1).Name
Next i

Sheets("Sheet2").Range("A2").CopyFromRecordset rs
Sheets("Sheet2").Columns.AutoFit

rs.Close
conn.Close

'Deleting Empty Columns
LastCol = Sheets("Sheet2").Range("A1:" & Range("A1").End(xlToRight).Address).Count
LastRow = Sheets("Sheet2").Range("A1").End(xlDown).Row

For Col = LastCol To 1 Step -1
    If Application.WorksheetFunction.CountA(Sheets("Sheet2").Range(Cells(2, Col), Cells(LastRow, Col))) = 0 Then
        Sheets("Sheet2").Columns(Col).Delete
    End If
Next Col

'Moving Columns
For i = 1 To 4
    LastCol = Sheets("Sheet2").Range("A1").End(xlToRight).Column
    Sheets("Sheet2").Columns(LastCol).Cut
    Sheets("Sheet2").Columns(1).Insert Shift:=xlToRight
Next i

Application.CutCopyMode = False

MsgBox "Dashboard Updated!", vbInformation, "Updated"

End Sub
Public Sub ShowData()

Dim HeaderList As ArrayList
Set HeaderList = New ArrayList
Dim ShowRng


LastCol = Sheets("Views").Range("B1:" & Range("B1").End(xlToRight).Address).Count
Lastcol_Name = Sheets("Views").Range("A1").End(xlToRight).Address
Lastcol_Name = Replace(Lastcol_Name, "1", "")
Row_No = Sheets("Views").Range("B1").End(xlDown).Row

'LastRow = Sheets("Views").Range("B1").End(xlDown).Row
ShowRng = Sheets("Views").Range("A1:" & Lastcol_Name & Row_No).Address
'SourceAdd = Sheets("Views").UsedRange.Address

UserForm1.ListBox1.ColumnCount = LastCol
UserForm1.ListBox1.RowSource = "Views!" & (ShowRng)

End Sub
Public Sub YearFilterDataAdder()

Dim DataArray As New ArrayList
Set DataArray = New ArrayList
Dim adFind As Range

'Adding Spend Year
Set adFind = Sheets("Sheet2").UsedRange.Find(What:="Spend Date", LookAt:=xlWhole)
ColAddress = Replace(adFind.Address, "$", "")
LastRow = Sheets("Sheet2").Range("E1").End(xlDown).Row
For i = 2 To LastRow
    If Len(Sheets("Sheet2").Range(Replace(ColAddress, "1", "") & i).Value) > 0 Then
        FormVal = Format(Sheets("Sheet2").Range(Replace(ColAddress, "1", "") & i).Value, "yyyy")
        If Not DataArray.Contains(FormVal) Then
             DataArray.Add FormVal
        End If
    End If
Next i
UserForm1.ComboBox1.AddItem "--Please Select--"
UserForm1.ComboBox1.AddItem "(All)"
UserForm1.ComboBox1.ListIndex = 0
For m = 0 To DataArray.Count - 1
    UserForm1.ComboBox1.AddItem DataArray(m)
Next m
Set DataArray = Nothing
End Sub

Public Sub CountryFilterDataAdder()

Dim DataArray As New ArrayList
Set DataArray = New ArrayList
Dim adFind As Range

'Adding Spend Country
Set adFind = Sheets("Sheet2").UsedRange.Find(What:="Expense Country", LookAt:=xlWhole)
ColAddress = Replace(adFind.Address, "$", "")
LastRow = Sheets("Sheet2").Range("E1").End(xlDown).Row
For i = 2 To LastRow
    If Len(Sheets("Sheet2").Range(Replace(ColAddress, "1", "") & i).Value) > 0 Then
        If Not DataArray.Contains(Sheets("Sheet2").Range(Replace(ColAddress, "1", "") & i).Value) Then
             DataArray.Add Sheets("Sheet2").Range(Replace(ColAddress, "1", "") & i).Value
        End If
    End If
Next i
UserForm1.ComboBox3.AddItem "--Please Select--"
UserForm1.ComboBox3.AddItem "(All)"
UserForm1.ComboBox3.ListIndex = 0
For m = 0 To DataArray.Count - 1
    UserForm1.ComboBox3.AddItem DataArray(m)
Next m
Set DataArray = Nothing
End Sub


Public Sub HomeSystemIdentiferDataAdder()

Dim DataArray As New ArrayList
Set DataArray = New ArrayList
Dim adFind As Range

'Adding Spend Home System Identifier
Set adFind = Sheets("Sheet2").UsedRange.Find(What:="Home System Identifier", LookAt:=xlWhole)
ColAddress = Replace(adFind.Address, "$", "")
LastRow = Sheets("Sheet2").Range("E1").End(xlDown).Row
For i = 2 To LastRow
    If Len(Sheets("Sheet2").Range(Replace(ColAddress, "1", "") & i).Value) > 0 Then
        If Not DataArray.Contains(Sheets("Sheet2").Range(Replace(ColAddress, "1", "") & i).Value) Then
             DataArray.Add Sheets("Sheet2").Range(Replace(ColAddress, "1", "") & i).Value
        End If
    End If
Next i

UserForm1.ComboBox2.AddItem "--Please Select--"
UserForm1.ComboBox2.AddItem "(All)"
UserForm1.ComboBox2.ListIndex = 0
For m = 0 To DataArray.Count - 1
    UserForm1.ComboBox2.AddItem DataArray(m)
Next m
Set DataArray = Nothing
End Sub

Public Sub IssueTypeDataAdder()

Dim DataArray() As Variant
Dim ErrorList As ArrayList
Set ErrorList = New ArrayList
Dim adFind As Range

Sheets("Sheet2").AutoFilterMode = False
'Adding Error Descripition
Set adFind = Sheets("Sheet2").UsedRange.Find(What:="Error Reason", LookAt:=xlWhole)
ColAddress = Replace(adFind.Address, "$", "")
LastRow = Sheets("Sheet2").Range("E1").End(xlDown).Row
DataArray = Sheets("Sheet2").Range("E2:E" & LastRow).Value
For i = 1 To UBound(DataArray, 1)
    'If Len(Sheets("Sheet2").Range(Replace(ColAddress, "1", "") & i).Value) > 0 Then
        If Not ErrorList.Contains(DataArray(i, 1)) Then
             ErrorList.Add DataArray(i, 1)
        End If
    'End If
Next i

UserForm1.ComboBox7.AddItem "--Please Select--"
UserForm1.ComboBox7.AddItem "(All)"
UserForm1.ComboBox7.ListIndex = 0
For m = 0 To ErrorList.Count - 1
    UserForm1.ComboBox7.AddItem ErrorList(m)
Next m
Set ErrorList = Nothing
End Sub
Public Sub AgingFilter()

Dim DataArray As New ArrayList
Set DataArray = New ArrayList
Dim adFind As Range

'Adding Error Descripition
Set adFind = Sheets("Sheet2").UsedRange.Find(What:="Aging_Days", LookAt:=xlWhole)
ColAddress = Replace(adFind.Address, "$", "")
LastRow = Sheets("Sheet2").Range("E1").End(xlDown).Row
For i = 2 To LastRow
    If Len(Sheets("Sheet2").Range(Replace(ColAddress, "1", "") & i).Value) > 0 Then
        If Not DataArray.Contains(Sheets("Sheet2").Range(Replace(ColAddress, "1", "") & i).Value) Then
             DataArray.Add Sheets("Sheet2").Range(Replace(ColAddress, "1", "") & i).Value
        End If
    End If
Next i
UserForm1.ComboBox4.AddItem "--Please Select--"
UserForm1.ComboBox4.AddItem "(All)"
UserForm1.ComboBox4.ListIndex = 0
For m = 0 To DataArray.Count - 1
    UserForm1.ComboBox4.AddItem DataArray(m)
Next m
Set DataArray = Nothing
End Sub
Public Sub AddUsersFilter()

Dim DataArray As New ArrayList
Set DataArray = New ArrayList
Dim adFind As Range

'Adding Error Descripition
Set adFind = Sheets("Sheet3").UsedRange.Find(What:="USER", LookAt:=xlWhole)
ColAddress = Replace(adFind.Address, "$", "")
LastRow = Sheets("Sheet3").Range("A1").End(xlDown).Row
For i = 2 To LastRow
    If Len(Sheets("Sheet3").Range(Replace(ColAddress, "1", "") & i).Value) > 0 Then
        If Not DataArray.Contains(Sheets("Sheet3").Range(Replace(ColAddress, "1", "") & i).Value) Then
             DataArray.Add Sheets("Sheet3").Range(Replace(ColAddress, "1", "") & i).Value
        End If
    End If
Next i
'MsgBox DataArray.Count
UserForm1.ComboBox5.AddItem "--Please Select--"
UserForm1.ComboBox5.AddItem "(All)"
UserForm1.ComboBox5.ListIndex = 0
For m = 0 To DataArray.Count - 1
    UserForm1.ComboBox5.AddItem DataArray(m)
Next m
Set DataArray = Nothing
End Sub
Public Sub StatusFilter()

Dim DataArray As New ArrayList
Set DataArray = New ArrayList
Dim adFind As Range

'Adding Error Descripition
Set adFind = Sheets("Sheet2").UsedRange.Find(What:="Completed_Status", LookAt:=xlWhole)
ColAddress = Replace(adFind.Address, "$", "")
LastRow = Sheets("Sheet2").Range("E1").End(xlDown).Row
For i = 2 To LastRow
    If Len(Sheets("Sheet2").Range(Replace(ColAddress, "1", "") & i).Value) > 0 Then
        If Not DataArray.Contains(Sheets("Sheet2").Range(Replace(ColAddress, "1", "") & i).Value) Then
             DataArray.Add Sheets("Sheet2").Range(Replace(ColAddress, "1", "") & i).Value
        End If
    End If
Next i
UserForm1.ComboBox8.AddItem "--Please Select--"
UserForm1.ComboBox8.AddItem "(All)"
UserForm1.ComboBox8.ListIndex = 0
For m = 0 To DataArray.Count - 1
    UserForm1.ComboBox8.AddItem DataArray(m)
Next m
Set DataArray = Nothing

End Sub
Public Sub PresentData()

F_Year = UserForm1.ComboBox1.Value
F_Country = UserForm1.ComboBox3.Value
F_HomeIdentifyer = UserForm1.ComboBox2.Value
F_Aging = UserForm1.ComboBox4.Value
ErrorDescrption = UserForm1.ComboBox7.Value
F_User = UserForm1.ComboBox5.Value
F_Status = UserForm1.ComboBox8.Value

FCell = Sheets("Sheet2").Range("A1").End(xlToRight).Address

'Selecting Year
If Not F_Year = "--Please Select--" Then
    If Not F_Year = "(All)" Then
        Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=33, Criteria1:="*" & F_Year & "*"
    Else
        Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=33
    End If
Else
    Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=33
End If

'Selecting Country
If Not F_Country = "--Please Select--" Then
    If Not F_Country = "(All)" Then
        Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=63, Criteria1:=F_Country
    Else
        Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=63
    End If
Else
    Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=63
End If

'Seleting Home System Identifier
If Not F_HomeIdentifyer = "--Please Select--" Then
    If Not F_HomeIdentifyer = "(All)" Then
        Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=26, Criteria1:=F_HomeIdentifyer
    Else
       Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=26
    End If
Else
    Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=26
End If

'Selecting Aging
If Not F_Aging = "--Please Select--" Then
    If Not F_Aging = "(All)" Then
        Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=4, Criteria1:=F_Aging
    Else
        Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=4
    End If
Else
    Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=4
End If

'Selecting Error Description
If Not ErrorDescrption = "--Please Select--" Then
    If Not ErrorDescrption = "(All)" Then
        Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=5, Criteria1:=ErrorDescrption
    Else
        Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=5
    End If
Else
    Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=5
End If

'Selecting user
If Not F_User = "--Please Select--" Then
    If Not F_User = "(All)" Then
        Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=1, Criteria1:=F_User
    Else
        Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=1
    End If
Else
    Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=1
End If

'Selecting Status
If Not F_Status = "--Please Select--" Then
    If Not F_Status = "(All)" Then
        Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=4, Criteria1:=F_Status
    Else
        Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=4
    End If
Else
    Sheets("Sheet2").Range("A1:" & Sheets("Sheet2").Range("A1").End(xlToRight).Address).AutoFilter Field:=4
End If

On Error Resume Next
Sheets("Views").Delete
Sheets.Add.Name = "Views"
Sheets("Views").Visible = xlSheetHidden
On Error GoTo 0

Sheets("Sheet2").UsedRange.Copy
Sheets("Views").Range("A1").PasteSpecial xlPasteAll

Call ShowData

End Sub

Private Function Queries(AccessDatabasePath)

Dim conn As Object
Dim strSQL As String
Dim rst As Object
Dim DBPath As String

'Sheets("Sheet2").Activate
DBPath = AccessDatabasePath

Set conn = CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath

On Error Resume Next
    conn.Execute "ALTER TABLE Spend_Issue ADD COLUMN [USER] TEXT (255)"
    conn.Execute "ALTER TABLE Spend_Issue ADD COLUMN [Aging_Days] TEXT (255)"
    conn.Execute "ALTER TABLE Spend_Issue ADD COLUMN [Pull_Date] TEXT (255)"
    conn.Execute "ALTER TABLE Spend_Issue ADD COLUMN [Completed_Status] TEXT (255)"
On Error GoTo 0

'Mapping User Names to dictionay from Sheet3
Dim CountryMap As Dictionary
Set CountryMap = New Dictionary

For i = 2 To Sheets("Sheet3").Range("A1").End(xlDown).Row
    CountryMap.Add Sheets("Sheet3").Range("B" & i).Value, Range("A" & i).Value
Next i


strSQL = "SELECT Country,[USER] FROM Spend_Issue"
Set rst = CreateObject("ADODB.Recordset")
rst.Open strSQL, conn, 1, 3

Set rs = CreateObject("ADODB.Recordset")
rs.Open strSQL, conn

'Adding Users to the country Names
Do While Not rst.EOF
    CurrentCountry = rst.Fields("Country").Value
    CurrentCountry = StrConv(CurrentCountry, vbProperCase)
    If CountryMap.ContainsKey(CurrentCountry) Then
        
        CurrentUserName = CountryMap(CurrentCountry)
        rst.Fields("USER").Value = CurrentUserName
        rst.Update
        
    End If
        rst.MoveNext
Loop

'Adding Current Date
CurDate = Format(DateAdd("M", 0, Date), "MM/dd/yyyy")
PdateQuery = "UPDATE Spend_Issue SET Pull_Date = Date() WHERE Pull_Date is Null"
conn.Execute PdateQuery


'Aging Days Add Query
AgingQuery = "UPDATE Spend_Issue SET Aging_Days = DateDiff('d',Pull_Date,Date()) WHERE Pull_Date IS NOT NULL and Completed_Status IS NULL"
conn.Execute AgingQuery
conn.Close

End Function

Public Sub DatabaseUpdate()

Dim conn As Object
Dim strSQL As String
Dim rst As Object
Dim DBPath As String

User = Environ("USERNAME")
ParentPath = ThisWorkbook.Path
ParentPath = Replace(ParentPath, "https://mydrive.merck.com/personal/" & User & "_merck_com/Documents", "C:\Users\" & User & "\OneDrive - Merck Sharp & Dohme LLC")
ParentPath = Replace(ParentPath, "/", "\")

DBPath = ParentPath & "\Dashboard.accdb"

Set conn = CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath

'Create Temp Table
CreatTable = "CREATE TABLE TempTable (" & _
            "[Home System Identifier] TEXT(50), " & _
            "[Completed_Status] TEXT(50));"
            
conn.Execute CreatTable

'Adding Data into Temp Table
strSQL = "INSERT INTO TempTable ([Home System Identifier], Completed_Status) " & _
            "SELECT [Home System Identifier],[Completed_Status] FROM [Excel 12.0 Xml;HDR=YES;IMEX=1;Database=" & "C:\Users\nanp\OneDrive - Merck Sharp & Dohme LLC\desktop\Spend_Issue\SpendIssues_Macro.xlsm" & "].[Views$];"

conn.Execute strSQL

'Performing Join to Update Status into Spend_issue Table
JoinSQL = "UPDATE Spend_Issue AS T " & _
            "INNER JOIN TempTable AS S " & _
            "ON T.[Home System Identifier] = S.[Home System Identifier] " & _
            "SET T.Completed_Status = S.[Completed_Status];"


conn.Execute JoinSQL

'Dropping TempTable
DropSQL = "DROP TABLE TempTable;"
conn.Execute DropSQL

conn.Close

MsgBox "Updated", vbInformation, "Success"

End Sub
