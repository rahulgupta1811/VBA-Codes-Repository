VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Launcher"
   ClientHeight    =   3010
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()
    If Me.ComboBox1.value = "Payment File" Then
        Call UpdateFile
    
    ElseIf Me.ComboBox1.value = "Payment File" Then
        Call UpdateFile

    ElseIf Me.ComboBox1.value = "Cost File" Then
        Call CostFile.StartProc
        
    ElseIf Me.ComboBox1.value = "IPC Payment File" Then
        Call IPC_Buying_Group.IPCPaymentFile
        
    ElseIf Me.ComboBox1.value = "APSC" Then
        Call APSC.APSCPaymentFile
        
    ElseIf Me.ComboBox1.value = "PBA" Then
        Call IPC_PBA.PBAPaymentFile
        
    ElseIf Me.ComboBox1.value = "Reliant" Then
        Call Reliant.ReliantPaymentFile
        
        ElseIf Me.ComboBox1.value = "APCI Payment File" Then
        Call APCI.APCIPaymentFile
        
    Else
        MsgBox "Please select any activity or close this file", vbCritical, "Selection Error"
    End If
End Sub

Private Sub UserForm_Initialize()

With ComboBox1
        .AddItem "Payment File"
        .AddItem "Cost File"
        .AddItem "IPC Payment File"
        .AddItem "PBA"
        .AddItem "APSC"
        .AddItem "Reliant"
        .AddItem "APCI Payment File"
        .AddItem "Updated BW"
End With

End Sub
