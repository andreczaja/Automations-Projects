VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_SAP 
   Caption         =   "Electrolux Group"
   ClientHeight    =   9780.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6135
   OleObjectBlob   =   "Form_SAP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_SAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public approver As String
Private Sub btn_OK_SAP_Click()

    Dim texto_data As String
    
    If Not checkbox_verificar_abrir_chamado_reembolsos_aprovados And Not checkbox_enviar_aprov_reembolsos_antigos And Not checkbox_processamento_novos_chamados Then
        MsgBox "Selecione ao menos uma etapa da automação para continuar!", vbInformation
        Exit Sub
    End If
    
    texto_data = Replace(Me.txt_box_data_agrupado_pgto_SAP.text, ".", "/")
    
    If texto_data = "" Or _
        (Me.opt_evelin_approver = False And Me.opt_luana_approver = False And Me.opt_thiago_approver = False And Me.opt_outro_approver = True And Me.txt_box_outro_approver = "") Then
        MsgBox "Preencha todos os campos obrigatórios do formulário.", vbOKOnly
        Exit Sub
    End If
    
    If Len(texto_data) <> 10 Then
        MsgBox "Digite uma data válida!", vbOKOnly
        Me.txt_box_data_agrupado_pgto_SAP.text = ""
        Exit Sub
    End If
    
    Debug.Print CInt(Mid(texto_data, 4, 2))
    If (CInt(Left(texto_data, 2)) < 1 Or CInt(Left(texto_data, 2)) > 31) Or _
            (CInt(Mid(texto_data, 4, 2)) < 1 Or CInt(Mid(texto_data, 4, 2)) > 12) Then
        MsgBox "Digite uma data válida!", vbOKOnly
        Me.txt_box_data_agrupado_pgto_SAP.text = ""
        Exit Sub
    End If
    If CDate(texto_data) < Date Then
        MsgBox "Digite uma data anterior a data de hoje!", vbOKOnly
        Me.txt_box_data_agrupado_pgto_SAP.text = ""
        Exit Sub
    End If
    Me.Hide

End Sub
Private Sub opt_evelin_approver_Click()
    With Me
        .opt_luana_approver = False
        .opt_thiago_approver = False
        .opt_outro_approver = False
        .txt_box_outro_approver.Visible = False
        .txt_box_outro_approver.text = ""
        .lbl_outro_biz.Visible = False
    End With
    approver = "SIZANEVE"
End Sub


Private Sub opt_luana_approver_Click()
    With Me
        .opt_evelin_approver = False
        .opt_thiago_approver = False
        .opt_outro_approver = False
        .txt_box_outro_approver.Visible = False
        .txt_box_outro_approver.text = ""
        .lbl_outro_biz.Visible = False
    End With
    approver = "RUCHILUA"
End Sub

Private Sub opt_outro_approver_Click()
    With Me
        .opt_evelin_approver = False
        .opt_luana_approver = False
        .opt_thiago_approver = False
        .txt_box_outro_approver.Visible = True
        .lbl_outro_biz.Visible = True
    End With
    approver = Me.txt_box_outro_approver.text
End Sub

Private Sub opt_thiago_approver_Click()
    With Me
        .opt_evelin_approver = False
        .opt_luana_approver = False
        .opt_outro_approver = False
        .txt_box_outro_approver.Visible = False
        .txt_box_outro_approver.text = ""
        .lbl_outro_biz.Visible = False
        
    End With
    approver = "BRAMBTHI"
End Sub

Private Sub txt_box_data_agrupado_pgto_SAP_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyBack Then
        txt_box_data_agrupado_pgto_SAP.text = ""
    End If
End Sub



Private Sub txt_box_data_agrupado_pgto_SAP_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim char As String
    Dim text As String

    ' Permitir apenas números, barra e backspace
    char = VBA.Chr(KeyAscii)
    If Not (char Like "[0-9/]" Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If

    ' Limitar o comprimento a 10 caracteres (DD/MM/YYYY)
    If Len(txt_box_data_agrupado_pgto_SAP.text) >= 10 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
    
    ' Adicionar barra automaticamente
    text = txt_box_data_agrupado_pgto_SAP.text
    If (Len(text) = 2 Or Len(text) = 5) And KeyAscii <> vbKeyBack Then
        txt_box_data_agrupado_pgto_SAP.text = text & "/"
        txt_box_data_agrupado_pgto_SAP.SelStart = Len(txt_box_data_agrupado_pgto_SAP.text) + 1
    End If
    
    
End Sub


Private Sub btn_cancelar_SAP_Click()
    Me.Hide
    End
End Sub

Private Sub UserForm_Deactivate()
    End
End Sub

Private Sub UserForm_Initialize()
    
    With Me
      
      .Width = 347
      .Height = 520
      .txt_box_data_agrupado_pgto_SAP.tabIndex = 0
      .txt_box_data_agrupado_pgto_SAP.tabStop = True
      .txt_box_data_agrupado_pgto_SAP.tabStop = True
      .opt_evelin_approver.tabIndex = 1
      .opt_evelin_approver.tabStop = True
      .opt_thiago_approver.tabIndex = 2
      .opt_thiago_approver.tabStop = True
      .opt_luana_approver.tabIndex = 3
      .opt_luana_approver.tabStop = True
      .opt_outro_approver.tabIndex = 4
      .opt_outro_approver.tabStop = True
      .btn_OK_SAP.tabIndex = 5
      .btn_OK_SAP.tabStop = True
      .btn_cancelar_SAP.tabIndex = 6
      .btn_cancelar_SAP.tabStop = True
      .opt_luana_approver = False
      .opt_thiago_approver = False
      .opt_outro_approver = False
      .txt_box_outro_approver.Visible = False
      .lbl_outro_biz.Visible = False
      .opt_evelin_approver = True
      .opt_evelin_approver.SetFocus
      .checkbox_enviar_aprov_reembolsos_antigos = True
      .checkbox_processamento_novos_chamados = True
      .checkbox_verificar_abrir_chamado_reembolsos_aprovados = True
    End With
    
End Sub
