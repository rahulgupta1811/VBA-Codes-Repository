Attribute VB_Name = "Module1"
Sub CalculateAging()
    Dim conn As Object
    Dim rs As Object
    Dim strSQL As String
    Dim tableName As String
    Dim agingColumn As String
    Dim idColumn As String
    Dim dateColumn As String
    Dim calculatedAging As Long
    Dim id As Long

    ' Define table and column names
    tableName = "TestTable"   ' Replace with your table name
    agingColumn = "Aging"          ' Column to store aging results
    idColumn = "GSD_ID"                ' Column that holds the ID
    dateColumn = "OccurrenceDate"   ' Column that holds the date

    ' Create a new ADODB connection
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\rahulgupta312039\Pictures\TestDb.accdb;" ' Replace with your database path

    ' Step 1: Open the recordset to loop through the table
    strSQL = "SELECT " & idColumn & ", MIN(" & dateColumn & ") AS MinDate, MAX(" & dateColumn & ") AS MaxDate " & _
             "FROM " & tableName & " " & _
             "GROUP BY " & idColumn & " " & _
             "HAVING " & idColumn & " IS NOT NULL;"

    ' Debug: Print the SELECT query to check if it's valid
    Debug.Print "SELECT SQL: " & strSQL

    ' Execute the SELECT query to fetch the results
    On Error GoTo ErrHandler
    Set rs = conn.Execute(strSQL)

    ' Debug: Check if the recordset is returning results
    If rs.EOF Then
        Debug.Print "No records found in the recordset."
    End If

    ' Step 2: Loop through the recordset and calculate the aging
    Do While Not rs.EOF
        ' Calculate the aging (difference in days)
        calculatedAging = DateDiff("d", rs("MinDate"), rs("MaxDate"))

        ' Debug: Check the values being updated
        Debug.Print "Updating ID: " & rs(idColumn) & " with Aging: " & calculatedAging

        ' Update the Aging column for each ID
        strSQL = "UPDATE " & tableName & " SET " & agingColumn & " = " & calculatedAging & " " & _
                 "WHERE " & idColumn & " = '" & rs(idColumn) & "';" ' Ensure value is properly enclosed in single quotes

        ' Debug: Print the UPDATE query to check if it's valid
        Debug.Print "UPDATE SQL: " & strSQL

        conn.Execute strSQL
        
        rs.MoveNext
    Loop

    MsgBox "Aging calculated and updated successfully!"

    ' Clean up
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description
    Set rs = Nothing
    Set conn = Nothing
End Sub


