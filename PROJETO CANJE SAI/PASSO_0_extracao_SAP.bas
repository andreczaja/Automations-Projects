Attribute VB_Name = "PASSO_0_extracao_SAP"
'variavel que possibilita que o codigo encerre caso se clique no botao de Cancelar no meio
' da execucao do codigo

Public btnCancelClicked As Boolean
Public SapGui As Object, WSHShell As Object, session As Object, Applic As Object, connection As Object, SapGuiWB As Object
Public CaminhoPasta As String
Public linha, linha_fim As Integer
Public canje As Worksheet


Sub SAP_Login()

    frmDate.Show


    MsgBox "Por favor, escolha a pasta onde as extra��es do SAP e Acepta ser�o salvos." & _
            "OBS: Certique-se de alterar o caminho das conex�es do Power Query com os arquivos!", vbInformation, "Aviso"

    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' O usu�rio selecionou uma pasta
            CaminhoPasta = .SelectedItems(1) & "\"
        Else
            ' O usu�rio cancelou a sele��o da pasta
            MsgBox "Nenhuma pasta selecionada. O processo foi cancelado."
            Exit Sub
        End If
    End With
    ' Verifica se o SAP est� aberto
    On Error Resume Next
    Set SapGui = GetObject("SAPGUI")
    On Error GoTo 0
            
    If Not SapGui Is Nothing Then
        
        ' SAP ABERTO
        Set SapGui = GetObject("SAPGUI")
        Set Applic = SapGui.GetScriptingEngine
        
        On Error Resume Next
        Set connection = Applic.Connections(0)
        
        If connection Is Nothing Then
            MsgBox "Favor iniciar a automatiza��o apenas com o SAP aberto e logado!", vbOKOnly, "Electrolux Group"
            Exit Sub
        End If
        Set session = connection.Children(0)
        
    Else:
      
      ' SAP FECHADO
        MsgBox "Favor iniciar a automatiza��o apenas com o SAP aberto e logado!", vbOKOnly, "Electrolux Group"
        End

    End If


'clica no bot�o para entrar no SAP

    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[0]").Maximize
'voltar para a pagina maior
    On Error Resume Next
    session.FindById("wnd[0]/tbar[0]/okcd").text = "/nBPMDG/UTL_BROWSER"
    session.FindById("wnd[0]").SendVKey 0

    Set canje = ThisWorkbook.Sheets("0. CANJE")
    linha_fim = canje.Range("B99999").End(xlUp).Row
    canje.Range("B22" & ":B" & linha_fim).Copy

'troca de m�duo (transa��o me2l)
me2l

    canje.Range("A22" & ":A" & linha_fim).Copy

'troca de m�duo (transa��o fbl5n)
fbl5n
    
    ActiveWorkbook.RefreshAll

    Set Applic = Nothing
    Set SapGui = Nothing
    Set connection = Nothing
    
'troca de m�duo (extra��o acepta)

verificacao_notas_acepta_
    
    MsgBox "Relat�rios extra�dos do SAP e relat�rio do Acepta salvo na pasta informada (" & CaminhoPasta & ").", vbOKOnly
    
End Sub

Sub me2l()




        session.FindById("wnd[0]/tbar[0]/okcd").text = "ME2L"
        session.FindById("wnd[0]").SendVKey 0
        session.FindById("wnd[0]/usr/ctxtEL_LIFNR-LOW").text = Range("A2")
        session.FindById("wnd[0]/usr/ctxtLISTU").text = "alv"
        session.FindById("wnd[0]/usr/btn%_EL_LIFNR_%_APP_%-VALU_PUSH").Press
        'cola sele��o de partidas acredoras
        session.FindById("wnd[1]/tbar[0]/btn[16]").Press
        session.FindById("wnd[1]/tbar[0]/btn[24]").Press
        session.FindById("wnd[1]/tbar[0]/btn[8]").Press
        session.FindById("wnd[0]/usr/ctxtS_BEDAT-LOW").text = frmDate.lbl_data_inicio
        session.FindById("wnd[0]/usr/ctxtS_BEDAT-HIGH").text = frmDate.lbl_data_final
        session.FindById("wnd[0]/tbar[1]/btn[8]").Press
        session.FindById("wnd[0]/tbar[1]/btn[45]").Press
        session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
        session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
        session.FindById("wnd[1]/tbar[0]/btn[0]").Press
        session.FindById("wnd[1]/usr/ctxtDY_PATH").text = CaminhoPasta
        session.FindById("wnd[1]/usr/ctxtDY_FILENAME").text = "ME2L" & ".txt"
        session.FindById("wnd[1]/tbar[0]/btn[11]").Press


End Sub


Sub fbl5n()

        session.FindById("wnd[0]/tbar[0]/okcd").text = "/nBPMDG/UTL_BROWSER"
        session.FindById("wnd[0]").SendVKey 0
        session.FindById("wnd[0]/tbar[0]/okcd").text = "FBL5N"
        session.FindById("wnd[0]").SendVKey 0
        session.FindById("wnd[0]/usr/btn%_DD_KUNNR_%_APP_%-VALU_PUSH").Press
        session.FindById("wnd[1]/tbar[0]/btn[16]").Press
        session.FindById("wnd[1]/tbar[0]/btn[24]").Press
        session.FindById("wnd[1]/tbar[0]/btn[8]").Press
        session.FindById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = "tc04"
        session.FindById("wnd[0]/usr/ctxtPA_VARI").text = "/SUR CTA CTE"
        session.FindById("wnd[0]/tbar[1]/btn[8]").Press
        session.FindById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
        session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
        session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
        session.FindById("wnd[1]/tbar[0]/btn[0]").Press
        session.FindById("wnd[1]/usr/ctxtDY_PATH").text = CaminhoPasta
        session.FindById("wnd[1]/usr/ctxtDY_FILENAME").text = "FBL5N.txt"
        session.FindById("wnd[1]/tbar[0]/btn[11]").Press

End Sub
