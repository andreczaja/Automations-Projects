Attribute VB_Name = "PASSO0_extracao_SAP"
'variavel que possibilita que o codigo encerre caso se clique no botao de Cancelar no meio
' da execucao do codigo
Public aba_export_sap, aba_nao_cobraveis, aba_relatorio_portal_devolucoes, aba_datas_sap, aba_calendarizacao As Worksheet
Public SapGui As Object, WSHShell As Object, session As Object, Applic As Object, connection As Object, SapGuiWB As Object
Private ultima_linha_preenchida_payers_exclusos As String
Public tabela_aba_export_sap, tabela_aba_nao_cobraveis, tabela_aba_relatorio_portal_devolucoes, tabela_aba_datas_sap, tabela_aba_calendarizacao As ListObject
Public linha, linha_fim, i As Integer
Public Folder, Pasta_Diaria, tipo_data_sap As String
Public data_final, data_inicial As Date



Sub SAP_Login()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Set aba_export_sap = ThisWorkbook.Sheets("Export SAP")
    Set tabela_aba_export_sap = aba_export_sap.ListObjects("Export_FBL5N___Cobráveis")
    Set aba_nao_cobraveis = ThisWorkbook.Sheets("Payers Não Cobraveis")
    Set tabela_aba_nao_cobraveis = aba_nao_cobraveis.ListObjects("Plan_Distr_Não_Cobrar")
    Set aba_datas_sap = ThisWorkbook.Sheets("Data Inicial X Final")
    Set tabela_aba_datas_sap = aba_datas_sap.ListObjects("Data_Inicial_e_Final_Extração_SAP")
    Set aba_calendarizacao = ThisWorkbook.Sheets("Calendarização")
    Set tabela_aba_calendarizacao = aba_calendarizacao.ListObjects("Calendarização")
    data_final = VBA.Date
    
    linha_fim = aba_datas_sap.Range("A1048576").End(xlUp).Row
    
    For linha = 2 To linha_fim
        If data_final = aba_datas_sap.Range("A" & linha).Value Then
            data_inicial = aba_datas_sap.Range("B" & linha - 1).Value
            Exit For
        End If
    Next linha
    
    frm_sim_nao_dt_manual.Show
    
    ' chamada da função buscar pasta para encontrar onde o arquivo do sap FBL5N deve ser descarregado - parâmetro Citrix como True já que é uma etapa de SAP Scripting BR
    Folder = BuscarPasta("", True)

    Set SapGui = GetObject("SAPGUI")
    Set Applic = SapGui.GetScriptingEngine
    Set connection = Applic.Connections(0)
    Set session = connection.Children(0)
    
    Call VerificarFormatoPadraoSAP
    
    
    tabela_aba_datas_sap.QueryTable.BackgroundQuery = False
    tabela_aba_datas_sap.QueryTable.Refresh False
    tabela_aba_nao_cobraveis.QueryTable.BackgroundQuery = False
    tabela_aba_nao_cobraveis.QueryTable.Refresh False
    tabela_aba_calendarizacao.QueryTable.BackgroundQuery = False
    tabela_aba_calendarizacao.QueryTable.Refresh False
    
    On Error Resume Next
    aba_nao_cobraveis.ShowAllData
    On Error GoTo 0

'voltar para a pagina maior
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n FBL5N"
    session.findById("wnd[0]").sendVKey 0
    
    '''''''''' SELECIONANDO VARIANTE MACRO COB
    
    tipo_data_sap = VerificarFormatoDatas(session.findById("wnd[0]/usr/ctxtPA_STIDA").text)

    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").Select
    session.findById("wnd[1]/usr/txtV-LOW").text = "MACRO COB"
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ctxtPA_STIDA").text = VBA.Format(aba_calendarizacao.Range("B2").Value, tipo_data_sap)

    ''''''''' EXCLUINDO PAYERS NÃO COBRAVEIS DA BASE
    linha_fim = aba_nao_cobraveis.Range("A1048576").End(xlUp).Row
    
    aba_nao_cobraveis.Range("A2:A" & linha_fim).Copy

    session.findById("wnd[0]/usr/btn%_DD_KUNNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
    ultima_linha_preenchida_payers_exclusos = Replace(Replace(VBA.Right(session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").text, 5), "(", ""), ")", "")
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.Position = CInt(ultima_linha_preenchida_payers_exclusos) + 2
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    session.findById("wnd[0]/usr/ctxtSO_FAEDT-LOW").text = VBA.Format(data_inicial, tipo_data_sap) ' definindo data inicial do vencimento liquido pag inicial FBL5N
    session.findById("wnd[0]/usr/ctxtSO_FAEDT-HIGH").text = VBA.Format(data_final, tipo_data_sap) ' definindo data final do vencimento liquido pag inicial FBL5N
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = Folder
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "FBL5N.txt"
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    
    tabela_aba_export_sap.QueryTable.BackgroundQuery = False
    tabela_aba_export_sap.QueryTable.Refresh False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Base de faturas cobráveis atualizada", vbOKOnly

End Sub




Sub atualizar()
 
    Sheets("Export SAP").ListObjects("Export_FBL5N___Cobráveis").QueryTable.BackgroundQuery = False
    Sheets("Export SAP").ListObjects("Export_FBL5N___Cobráveis").QueryTable.Refresh False
    Sheets("Relatório Portal de Devoluções").ListObjects("Tabela_Relatório_Portal_de_Devoluções").QueryTable.BackgroundQuery = False
    Sheets("Relatório Portal de Devoluções").ListObjects("Tabela_Relatório_Portal_de_Devoluções").QueryTable.Refresh False
    Sheets("Payers Não Cobraveis").ListObjects("Plan_Distr_Não_Cobrar").QueryTable.BackgroundQuery = False
    Sheets("Payers Não Cobraveis").ListObjects("Plan_Distr_Não_Cobrar").QueryTable.Refresh False
    
End Sub

