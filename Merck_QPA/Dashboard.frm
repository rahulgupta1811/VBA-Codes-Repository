VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Spend_Issue_Dashboard"
   ClientHeight    =   8300.001
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   15550
   OleObjectBlob   =   "Dashboard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Dim NB As Workbook


Set NB = Workbooks.Add

ThisWorkbook.Activate
Sheets("Views").UsedRange.Copy

NB.Activate
Sheets(1).Activate

Range("A1").PasteSpecial xlPasteAll
Columns.AutoFit

ParentPath = ThisWorkbook.Path
ParentPath = Replace(ParentPath, "https://mydrive.merck.com/personal/dasakas_merck_com/Documents", "C:\Users\dasakas\OneDrive - Merck Sharp & Dohme LLC")
ParentPath = Replace(ParentPath, "/", "\")

NewFilePath = ParentPath & "\" & "Report_" & Format(DateAdd("M", 0, Now()), "mmddyy_HH_SS") & ".xlsx"

NB.SaveAs Filename:=NewFilePath
NB.Close
MsgBox "Report Exported", vbInformation, "Success"

End Sub

Public Sub CommandButton2_Click()
Call Data_Creation.PresentData
End Sub

Private Sub CommandButton3_Click()
UserForm2.Show

End Sub

Private Sub CommandButton4_Click()
Data_Creation.DatabaseUpdate
End Sub
