VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Input Selector"
   ClientHeight    =   3540
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   5220
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

If Len(UserForm1.TextBox1.Value) = 0 Then
    MsgBox "Enter Input File Path", vbCritical, "Input File Missing"
ElseIf Len(UserForm1.TextBox2.Value) = 0 Then
    MsgBox "Enter Reference ID", vbCritical, "Reference ID Missing"
Else
    Call GL_Activity.Start_activity
End If

End Sub

Private Sub CommandButton2_Click()

UserForm1.Hide

End Sub

Private Sub CommandButton3_Click()
Dim fd As FileDialog
Dim filePath As String

Set fd = Application.FileDialog(msoFileDialogFilePicker)

With fd
    .Title = "Select a File"
    .AllowMultiSelect = False
    .Filters.Clear
    .Show
    
    If .SelectedItems.Count > 0 Then
        filePath = .SelectedItems(1)
        UserForm1.TextBox1.Value = filePath
    End If
End With
    Set fd = Nothing

End Sub
