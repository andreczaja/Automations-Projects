Attribute VB_Name = "PASSO_0_REEMBOLSOS_APROVADOS"
Public linhas_SBWP, reembolsos_aprovados As Integer
Sub buscar_novas_linhas_SBWP_rodada_anterior()

    Call declaracao_vars
    

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    reembolsos_com_dados_bancarios_processados = 0
    
    
    If aba_reembolsos_pendentes.Range("A2").Value = "" Then
        Folder = BuscarPasta("", True)
        FileCopy Folder & "\REEMBOLSOS APROVADOS - BASE VAZIA.txt", Folder & "\FBL5N - REEMBOLSOS APROVADOS.txt"
        tabela_reembolsos_aprovados.QueryTable.BackgroundQuery = False
        tabela_reembolsos_aprovados.QueryTable.Refresh
        MsgBox "Nenhuma linha na SBWP a ser buscada. A base de Reembolsos Aguardando Aprovação está vazia.", vbOKOnly
        Exit Sub
    End If
    
    On Error Resume Next
    Set SapGui = GetObject("SAPGUI")
    Set Applic = SapGui.GetScriptingEngine
    Set connection = Applic.Connections(0)
    Set session = connection.Children(0)
    Call VerificarFormatoPadraoSAP
    Form_SAP.Show
    data_agrupado_pagamento = Form_SAP.txt_box_data_agrupado_pgto_SAP
    aba_reembolsos_aprovados.Range("BC1").Value = data_agrupado_pagamento
    Folder = BuscarPasta("", True)
    On Error GoTo 0
    
    ' verificacao de linhas vazias na base reembolsos ag aprovacao, se for vazia, será buscada a pasta Arquivos SAP Macro Reembolsos e Adiantamentos
    ' e o arquivo com a base vazia substituirá o antigo de reembolsos aprovados, para que o sistema entenda que não existem reembolsos a serem processados na SBWP
    ' caso contrário irá procurar esses reembolsos na SBWP para fazer o processo de clicar em Documento > Completo em cada um deles (Processo que envia para aprovação)
    If Not VerificarLinhasSBWP(session) Then
        Call extrair_base_reembolsos_aprovados
        Exit Sub
    Else
        Call alterar_campos_linhas_rodada_anterior_enviadas_aprovacao
        Call extrair_base_reembolsos_aprovados
    End If
    
    Call LimparFiltros(tabela_aba_reembolsos_pendentes)
    
    MsgBox "Foram enviados para a aprovação " & reembolsos_com_dados_bancarios_processados & vbNewLine & _
        "Além disso, foram verificados " & reembolsos_aprovados & _
            " reembolsos aprovados, os quais devem ser aberto chamados na ellevo e enviado o e-mail de notificação ao cliente (FASE 2)", vbOKOnly
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

Sub alterar_campos_linhas_rodada_anterior_enviadas_aprovacao()

    linha_fim_aba_reembolsos_pendentes = aba_reembolsos_pendentes.Range("A1048576").End(xlUp).Row
    tabela_aba_reembolsos_pendentes.Range.AutoFilter Field:=30, Criteria1:=VBA.Format(Date, "dd/mm/yyyy")
    tabela_aba_reembolsos_pendentes.Range.AutoFilter Field:=31, Criteria1:="Aguardando Aprovação"

    session.findById("wnd[0]/tbar[0]/okcd").text = "/N FBL5N"
    session.findById("wnd[0]").sendVKey 0
    tipo_data_sap = VerificarFormatoDatas(session.findById("wnd[0]/usr/ctxtPA_STIDA").text)

    aba_reembolsos_pendentes.Range("C2:C" & linha_fim_aba_reembolsos_pendentes).Copy
    session.findById("wnd[0]/usr/btn%_DD_KUNNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/chkX_SHBV").Selected = True
    session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = "BR10"
    session.findById("wnd[0]/tbar[1]/btn[16]").press
    
    aba_reembolsos_pendentes.Range("G2:G" & linha_fim_aba_reembolsos_pendentes).Copy
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN012_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/txt%%DYN011-LOW").text = "AG. APROV REEMB"

    session.findById("wnd[0]/usr/ctxtPA_VARI").text = "/ABATREEMB"
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' COLOCAR AQUI NA ASIGNACION A DATA DO AGRUPADO DE PAGAMENTO
    session.findById("wnd[0]").sendVKey 5
    session.findById("wnd[0]/tbar[1]/btn[45]").press
    If session.findById("wnd[0]/sbar").text = "Marcar pelo menos uma partida" Then
        Exit Sub
    End If
    session.findById("wnd[1]/usr/txt*BSEG-ZUONR").text = "REEMB AUT " & VBA.Format(data_agrupado_pagamento, "dd.mm.yy")
    session.findById("wnd[0]").sendVKey 0
    On Error Resume Next
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    On Error GoTo 0
    
End Sub
Sub extrair_base_reembolsos_aprovados()

    ' depois de realizar todo o processo na SBWP irá verificar se existem linhas as quais estavam pendente de aprovação e que foram aprovadas
    ' ENTÃO BUSCARÁ NA FBL5N AS PARTIDAS ABERTAS (COM ASIGNACION DIFERENTE DE "PROCESSADO AUTOMACAO") e USAR CHAVE DE REF 2 IGUAL A AUTOMACAO
    'ESSAS PARTIDAS ABERTAS QUER DIZER QUE JÁ FORAM APROVADOS PELO THIAGO OU LUANA
    ' NO EXCEL ENTÃO VEREMOS A CORRESPONDÊNCIA ENTRE A ABA AG APROVAÇÃO E APROVADOS PARA ENTENDER AQUI SE PODEMOS OU NÃO SEGUIR COM A CRIAÇÃO DO CHAMADO ELLEVO DA DETERMINADA LINHA
    'NAS PARTIDAS ENCONTRADAS DO EXCEL, DEVEMOS ANTES DE EXPORTAR, TROCAR A ASIGNACION DE TODAS QUE NÃO FORAM PROCESSADAS PARA "PROCESSADO AUTOMACAO" NO CAMPO ASIGNACION
    'PARA QUE NO MOMENTO DA FILTRAGEM DE PROXIMAS LEVAS, NÃO PUXE ESSAS LINHAS

    Call declaracao_vars
    Folder = BuscarPasta("", True)
    session.findById("wnd[0]/tbar[0]/okcd").text = "/N FBL5N"
    session.findById("wnd[0]").sendVKey 0
    
    
    tipo_data_sap = VerificarFormatoDatas(session.findById("wnd[0]/usr/ctxtPA_STIDA").text)
    session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = "BR10"
    ' ESCOLHENDO VARIANTE "REEMBO. AUTOMA" E DEMAIS FILTROS PARA GERAÇÃO DA BASE
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").Select
    session.findById("wnd[1]/usr/txtV-LOW").text = "REEMBO. AUTOMA"
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    session.findById("wnd[0]/usr/ctxtPA_STIDA").text = VBA.Format(Date, tipo_data_sap)
    
    session.findById("wnd[0]/tbar[1]/btn[16]").press
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN011_%_APP_%-VALU_PUSH").press
    
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,0]").text = "AUTOMACAO DEV"
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    
    linha_fim_aba_reembolsos_pendentes = aba_reembolsos_pendentes.Range("A1048576").End(xlUp).Row
    aba_reembolsos_pendentes.Range("C2:C" & linha_fim_aba_reembolsos_pendentes).Copy
    
    session.findById("wnd[0]/usr/btn%_DD_KUNNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    If Left(session.findById("wnd[0]/sbar").text, 12) <> "São exibidas" And linhas_SBWP = 0 Then
        MsgBox "Não existem reembolsos aprovados nem enviados para aprovação. Além disso, nenhuma das linhas pendentes de aprovação foi aprovada. Por favor, siga para a FASE 4", vbOKOnly
        'FileCopy Folder & "\REEMBOLSOS APROVADOS - BASE VAZIA.txt", Folder & "\FBL5N - REEMBOLSOS APROVADOS.txt"
        'tabela_reembolsos_aprovados.QueryTable.Refresh False
        Exit Sub
    ElseIf Left(session.findById("wnd[0]/sbar").text, 12) <> "São exibidas" And linhas_SBWP > 0 Then
        MsgBox "Nenhuma das linhas pendentes de aprovação foi aprovada. Por favor, siga para a FASE 4", vbOKOnly
        'FileCopy Folder & "\REEMBOLSOS APROVADOS - BASE VAZIA.txt", Folder & "\FBL5N - REEMBOLSOS APROVADOS.txt"
        'tabela_reembolsos_aprovados.QueryTable.Refresh False
        Exit Sub
    End If
    
    session.findById("wnd[0]").sendVKey 5
    session.findById("wnd[0]/tbar[1]/btn[45]").press
    session.findById("wnd[1]/usr/txt*BSEG-ZUONR").text = "PROCESSADO AUTOMAC"
    session.findById("wnd[0]").sendVKey 0
    On Error Resume Next
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    On Error GoTo 0
    
    
    Call SalvarArquivo(session, "FBL5N - REEMBOLSOS APROVADOS.txt")
    Call LimparFiltros(tabela_reembolsos_aprovados)
    Call LimparFiltros(tabela_aba_reembolsos_pendentes)
    
    
    linha_fim = aba_reembolsos_aprovados.Range("A1048576").End(xlUp).Row
    
    reembolsos_aprovados = linha_fim - 1
    
    Call AtualizarBase(aba_reembolsos_aprovados, tabela_reembolsos_aprovados, linha_fim)
    
    MsgBox "Linhas de reembolsos atualizadas.", vbOKOnly

End Sub
