VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Select Date"
   ClientHeight    =   850
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   3530
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
If DTPicker1 = Null Then
    MsgBox "Please select any date", vbCritical, "Error"
Else
    Dim Sdate As String
    Sdate = DTPicker1.Value
    Sheetdate = Format(Sdate, "dd MMMM yyyy")
    CompleteDate = Format(Sdate, "dddd, MMMM dd, yyyy")
    ActiveSheet.[C2].Value = CompleteDate
    'MsgBox Sdate
    

    For i = 1 To Worksheets.Count
            If Worksheets(i).Name = Sheetdate Then
                Exist = True
            End If
        Next i
        
        If Not Exist Then
            'MsgBox YesterdayDate
            'MsgBox Todaysdate
            Sheets("Sample").Copy before:=Sheets("Sample")
            ActiveSheet.Name = Sheetdate
            
            Dim Sheetcount As Integer
            Sheetcount = Worksheets.Count
            Sheetcount = Sheetcount - 2
            Worksheets(Sheetcount).Tab.Color = False
            Sheets(Sheetdate).Tab.Color = vbRed
            
            Dim msg As Integer
            msg = MsgBox("Completed 1 of 2 Macro", vbInformation, "Done")
           
        Else
            msg = MsgBox("(" & Sheetdate & ")" & " Sheet Already Exists", vbExclamation, "Stop")
        
        End If
    
End If
'Worksheets("Sample").Visible = False

Me.Hide
End Sub
