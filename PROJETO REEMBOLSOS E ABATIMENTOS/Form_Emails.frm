VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Emails 
   Caption         =   "Electrolux Group"
   ClientHeight    =   4635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8295.001
   OleObjectBlob   =   "Form_Emails.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_Emails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_cancelar_SAP_Click()
    Form_Emails.Hide
    End
End Sub

Private Sub btn_OK_SAP_Click()
    Form_Emails.Hide
End Sub

Private Sub opt_ambos_Click()
    With Me
        .opt_ambos = True
        .opt_apenas_abatimentos = False
        .opt_apenas_reembolsos = False
    End With
End Sub
Private Sub opt_apenas_abatimentos_Click()
    With Me
        .opt_ambos = False
        .opt_apenas_abatimentos = True
        .opt_apenas_reembolsos = False
    End With
End Sub
Private Sub opt_apenas_reembolsos_Click()
    With Me
        .opt_ambos = False
        .opt_apenas_abatimentos = False
        .opt_apenas_reembolsos = True
    End With
End Sub

Private Sub UserForm_Initialize()
    With Me
        .opt_ambos = True
        .opt_apenas_abatimentos = False
        .opt_apenas_reembolsos = False
        .Height = 261
        .Width = 427
    End With
End Sub
