VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Ellevo 
   Caption         =   "Electrolux Group"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6690
   OleObjectBlob   =   "Form_Ellevo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_Ellevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    Call declaracao_vars
    
    data_agrupado_pagamento = aba_reembolsos_aprovados.Range("BC1").Value
    
    With Me
      
        .Width = 347
        .Height = 393
        .txt_box_data_agrupado_pgto_ellevo.tabIndex = 0
        .txt_box_data_agrupado_pgto_ellevo.tabStop = True
        .txt_box_data_agrupado_pgto_ellevo.text = data_agrupado_pagamento
        .txtbox_login_ellevo.tabIndex = 1
        .txtbox_login_ellevo.tabStop = True
        .txtbox_login_ellevo.text = "contasareceber@electrolux.com"
        .txtbox_senha_ellevo.tabIndex = 2
        .txtbox_senha_ellevo.tabStop = True
        .txtbox_senha_ellevo.text = "Elux@2024"
        .btn_OK_ellevo.tabIndex = 3
        .btn_OK_ellevo.tabStop = True
        .btn_cancelar_ellevo.tabIndex = 4
        .btn_cancelar_ellevo.tabStop = True
      
    End With
    
End Sub
Private Sub txtbox_login_ellevo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyBack Then
        txtbox_login.text = ""
    End If

End Sub
Private Sub txtbox_senha_ellevo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyBack Then
        txtbox_senha.text = ""
    End If

End Sub

Private Sub txtbox_senha_Change()
txtbox_senha.PasswordChar = "*"
End Sub
Private Sub btn_OK_ellevo_Click()

    Dim texto_data As String
    
    
    
    texto_data = Replace(Me.txt_box_data_agrupado_pgto_ellevo.text, ".", "/")
    
    If texto_data = "" Or Me.txtbox_login_ellevo = "" Or Me.txtbox_senha_ellevo = "" Then
        MsgBox "Preencha todos os campos obrigatórios do formulário.", vbOKOnly
        Exit Sub
    End If
    
    If Len(texto_data) < 10 Then
        MsgBox "Digite uma data válida!", vbOKOnly
        Me.txt_box_data_agrupado_pgto_ellevo.text = ""
        Exit Sub
    End If
    
    Debug.Print CInt(Mid(texto_data, 4, 2))
    If (CInt(Left(texto_data, 2)) < 1 Or CInt(Left(texto_data, 2)) > 31) Or _
            (CInt(Mid(texto_data, 4, 2)) < 1 Or CInt(Mid(texto_data, 4, 2)) > 12) Then
        MsgBox "Digite uma data válida!", vbOKOnly
        Me.txt_box_data_agrupado_pgto_ellevo.text = ""
        Exit Sub
    End If
    If CDate(texto_data) < Date Then
        MsgBox "Digite uma data anterior a data de hoje!", vbOKOnly
        Me.txt_box_data_agrupado_pgto_ellevo.text = ""
        Exit Sub
    End If
    Me.Hide

End Sub

Private Sub txt_box_data_agrupado_pgto_ellevo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyBack Then
        txt_box_data_agrupado_pgto_ellevo.text = ""
    End If
End Sub



Private Sub txt_box_data_agrupado_pgto_ellevo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim char As String
    Dim text As String

    ' Permitir apenas números, barra e backspace
    char = VBA.Chr(KeyAscii)
    If Not (char Like "[0-9/]" Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If

    ' Limitar o comprimento a 10 caracteres (DD/MM/YYYY)
    If Len(txt_box_data_agrupado_pgto_ellevo.text) >= 10 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
    
    ' Adicionar barra automaticamente
    text = txt_box_data_agrupado_pgto_ellevo.text
    If (Len(text) = 2 Or Len(text) = 5) And KeyAscii <> vbKeyBack Then
        txt_box_data_agrupado_pgto_ellevo.text = text & "/"
        txt_box_data_agrupado_pgto_ellevo.SelStart = Len(txt_box_data_agrupado_pgto_ellevo.text) + 1
    End If
End Sub


Private Sub btn_cancelar_ellevo_Click()
    Me.Hide
    End
End Sub

Private Sub UserForm_Deactivate()
    End
End Sub
