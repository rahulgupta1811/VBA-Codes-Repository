Attribute VB_Name = "Clone"
Sub Clone_PrevSheet()

'On Error GoTo Defaulterror

    Dim Exist As Boolean
    Dim Todaysdate As String
    Dim YesterdayDate As String
    Dim NowDay As String
    Dim Selectoption As String
    
    Selectoption = MsgBox("Do you want to set date manually ?", vbYesNoCancel, "Selection")
    If Selectoption = vbYes Then
        UserForm1.Show
    Else
    
    Todaysdate = Format(Date, "dd MMMM yyyy")
    NowDay = Format(Date, "dddd")
    'Checking if today is Monday
        If NowDay = "Wednesday" Then
            YesterdayDate = Format(Date - 3, "dd MMMM yyyy")
        Else
            YesterdayDate = Format(Date - 1, "dd MMMM yyyy")
        End If
    
    'Checking if todays sheet exist or not
        For i = 1 To Worksheets.Count
            If Worksheets(i).Name = Todaysdate Then
                Exist = True
            End If
        Next i
        
        If Not Exist Then
            'MsgBox YesterdayDate
            'MsgBox Todaysdate
            Sheets("Sample").Copy before:=Sheets("Sample")
            ActiveSheet.Name = Todaysdate
            
            Dim Sheetcount As Integer
            Sheetcount = Worksheets.Count
            Sheetcount = Sheetcount - 2
            Worksheets(Sheetcount).Tab.Color = False
            Sheets(Todaysdate).Tab.Color = vbRed
            Dim CompleteDate  As String
            CompleteDate = Format(Date, "dddd, MMMM dd, yyyy")
            ActiveSheet.[C2].Value = CompleteDate
            Dim msg As Integer
            msg = MsgBox("Completed 1 of 2 Macro", vbInformation, "Done")
           
        Else
            msg = MsgBox("(" & Todaysdate & ")" & " Sheet Already Exists", vbExclamation, "Stop")
        
        End If
    End If
'Done:
    'Exit Sub

'Defaulterror:
    'MsgBox "There is some data processing error", vbCritical, Error
    
Sheets("Sample").Visible = False
End Sub
