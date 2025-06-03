VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_sim_nao_dt_manual 
   Caption         =   "Electrolux Group"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "frm_sim_nao_dt_manual.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_sim_nao_dt_manual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_nao_Click()
    With Me
        .Height = 299
        .btn_sim.Enabled = False
        .btn_nao.Enabled = False
        .btn_ok.Visible = True
        .btn_voltar.Visible = True
        .lbl_data_inicial.Visible = True
        .txtbox_data_inicial.Visible = True
        .lbl_data_final.Visible = True
        .txtbox_data_final.Visible = True
    End With
        
    
End Sub

Private Sub btn_ok_Click()
    If frm_sim_nao_dt_manual.txtbox_data_inicial = "" Or frm_sim_nao_dt_manual.txtbox_data_final = "" Then
        MsgBox "Preencha as datas inicias e finais para continuar", vbOKOnly
        Call btn_nao_Click
    End If
    frm_sim_nao_dt_manual.Hide
    data_inicial = frm_sim_nao_dt_manual.txtbox_data_inicial
    data_final = frm_sim_nao_dt_manual.txtbox_data_final
End Sub

Private Sub btn_sim_Click()
    frm_sim_nao_dt_manual.Hide
End Sub

Private Sub btn_voltar_Click()
    With Me
        btn_sim.Enabled = True
        btn_nao.Enabled = True
    End With
    Call UserForm_Initialize
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    End
End Sub


Private Sub UserForm_Initialize()
    With Me
        .Height = 185
        .Width = 400
        .label_datas.Caption = "As datas inicias e finais calculadas pela automação são, respectivamente " & _
            data_inicial & " e " & data_final & ". Deseja seguir?"
        .btn_ok.Visible = False
        .btn_voltar.Visible = False
        .lbl_data_inicial.Visible = False
        .txtbox_data_inicial.Visible = False
        .lbl_data_final.Visible = False
        .txtbox_data_final.Visible = False
    End With
End Sub

Private Sub txtbox_data_inicial_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyBack Then
        txtbox_data_inicial.text = ""
    End If
End Sub


Private Sub txtbox_data_inicial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim char As String
    Dim text As String

    ' Permitir apenas números, barra e backspace
    char = VBA.Chr(KeyAscii)
    If Not (char Like "[0-9/]" Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If

    ' Limitar o comprimento a 10 caracteres (DD/MM/YYYY)
    If Len(txtbox_data_inicial.text) >= 10 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
    
    ' Adicionar barra automaticamente
    text = txtbox_data_inicial.text
    If (Len(text) = 2 Or Len(text) = 5) And KeyAscii <> vbKeyBack Then
        txtbox_data_inicial.text = text & "/"
        txtbox_data_inicial.SelStart = Len(txtbox_data_inicial.text) + 1
    End If
End Sub

Private Sub txtbox_data_inicial_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Validar se a data está no formato DD/MM/YYYY
    On Error Resume Next
    Dim dateValue As Date
    dateValue = CDate(txtbox_data_inicial.text)
    If Err.Number <> 0 Then
        MsgBox "Por favor, insira uma data válida no formato DD/MM/YYYY.", vbExclamation
        txtbox_data_inicial.SetFocus
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtbox_data_final_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyBack Then
        txtbox_data_final.text = ""
    End If
End Sub


Private Sub txtbox_data_final_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim char As String
    Dim text As String

    ' Permitir apenas números, barra e backspace
    char = VBA.Chr(KeyAscii)
    If Not (char Like "[0-9/]" Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If

    ' Limitar o comprimento a 10 caracteres (DD/MM/YYYY)
    If Len(txtbox_data_final.text) >= 10 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
    
    ' Adicionar barra automaticamente
    text = txtbox_data_final.text
    If (Len(text) = 2 Or Len(text) = 5) And KeyAscii <> vbKeyBack Then
        txtbox_data_final.text = text & "/"
        txtbox_data_final.SelStart = Len(txtbox_data_final.text) + 1
    End If
End Sub

Private Sub txtbox_data_final_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Validar se a data está no formato DD/MM/YYYY
    On Error Resume Next
    Dim dateValue As Date
    dateValue = CDate(txtbox_data_final.text)
    If Err.Number <> 0 Then
        MsgBox "Por favor, insira uma data válida no formato DD/MM/YYYY.", vbExclamation
        txtbox_data_final.SetFocus
        Cancel = True
    End If
    On Error GoTo 0
End Sub



