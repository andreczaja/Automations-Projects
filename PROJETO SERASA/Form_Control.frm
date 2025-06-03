VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Control 
   Caption         =   "Electrolux Group"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5325
   OleObjectBlob   =   "Form_Control.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_Control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_processar_Click()
    Me.Hide
End Sub

Private Sub opt_ambos_Click()
    With Me
        .opt_ambos = True
        .opt_apenas_exclusao = False
        .opt_apenas_inclusao = False
    End With
End Sub

Private Sub opt_apenas_exclusao_Click()
    With Me
        .opt_ambos = False
        .opt_apenas_exclusao = True
        .opt_apenas_inclusao = False
    End With
End Sub
Private Sub opt_apenas_inclusao_Click()
    With Me
        .opt_ambos = False
        .opt_apenas_exclusao = False
        .opt_apenas_inclusao = True
    End With
End Sub

Private Sub UserForm_Click()
    Form_Control.Hide
End Sub

Private Sub UserForm_Initialize()
    With Me
        .Height = 230
        .Width = 280
        .opt_ambos = True
        .opt_apenas_exclusao = False
        .opt_apenas_inclusao = False
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        End ' Encerra completamente a execução do código
    End If
End Sub

