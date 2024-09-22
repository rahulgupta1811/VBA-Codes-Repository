VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Select WF File"
   ClientHeight    =   2010
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4890
   OleObjectBlob   =   "File_Select_Userform1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

If TextBox1.Value = "" Then
    MsgBox "Please enter the WF File Path", vbCritical, "Error"
Else
    Path = TextBox1.Value
    Path = Replace(Path, Chr(34), "")
    TextBox1.Value = Path
    Me.Hide
    Call CPH_WF
End If
End Sub

Private Sub CommandButton2_Click()
Dim File As Office.FileDialog
Set File = Application.FileDialog(msoFileDialogFilePicker)
With File
    .AllowMultiSelect = False
    .Filters.Clear
    '.Filters.Add "CSV", "*CSV?"
    If .Show = True Then
        Filename = Dir(.SelectedItems(1))
        TextBox1.Value = Filename
    End If
End With
End Sub
