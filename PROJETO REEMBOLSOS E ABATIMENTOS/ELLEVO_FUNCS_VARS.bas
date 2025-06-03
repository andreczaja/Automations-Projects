Attribute VB_Name = "ELLEVO_FUNCS_VARS"
Public elemento_login, elemento_senha, elemento_entrar, elemento_barra_busca_chamado, elemento_chamado_encontrado_page_inicial, elemento_chamado_selecionado, _
      elemento_chamado_encontrado_page_chamado, elemento_botao_assumir_chamado, elemento_novo_tramite, elemento_status_chamado, elemento_observacoes, _
      elemento_bt_enviar_trâmite_por_email, elemento_texto_chamado, elemento_popup_status_chamado, elemento_caixa_texto_tramite_chamado_cliente, elemento_popup_confirmar, elemento_servico, elemento_natureza, elemento_fornecedores_clientes, _
      elemento_pagamento_normal, elemento_classificacao, elemento_reembolso_clientes, elemento_empresa, elemento_codigo_cliente, elemento_nome_fornecedor_cliente, elemento_forma_de_pagamento, elemento_novo_tramite_chamado_aberto, _
      elemento_nota_fiscal, elemento_documento_sap, elemento_valor_total, elemento_data_pagamento, elemento_anexar_arquivos, elemento_salvar, elemento_barra_busca_servicos, elemento_solic_pag_nacional, elemento_popup_chamado_aberto, _
      elemento_natureza_fechar, elemento_natureza_expand, elemento_responsavel, elemento_responsavel_busca, elemento_responsavel_selecao, elemento_salvar_chamado, elemento_responsavel2, elemento_natureza_fechar2, elemento_login_incorreto, elemento_severidade, _
      elemento_lista_suspensa, elemento_pgto_agrupado As String
Public driver As New EdgeDriver
Public by As New by
Public contador As Integer

Sub Variaveis_Ellevo()



'Atribuição de variantes publicas, ordenado de A-Z
elemento_anexar_arquivos = "/html/body/app-root/app-attendant-main/div/div/div/div/app-opening/app-ticket-opening/div/app-ticket-opening-form/div[1]/app-portlet/form/div/div[3]/div/div/button[1]"
elemento_barra_busca_chamado = "/html/body/app-root/app-attendant-main/div/div/app-header/div[1]/app-global-search/app-search-input/input"
elemento_barra_busca_servicos = "/html/body/div/div/div/div/div/div[2]/div/div/div/div/app-search-input/input"
elemento_botao_assumir_chamado = "/html/body/app-root/app-attendant-main/div/div/div/div/app-toolbar/div/div[1]/app-assume-toolbar-button/app-toolbar-button"
elemento_bt_enviar_trâmite_por_email = "/html/body/div/div[2]/div/mat-dialog-container/div/div[2]/div[3]/div/app-switch/nz-switch/button" 'FAZER AUTOMAÇÃO VERIFICAR SE JÁ ESTÁ HABILITADA A OPÇÃO
elemento_caixa_texto_tramite_chamado_cliente = "/html/body/div[1]/div[2]/div/mat-dialog-container/div/div[2]/div[1]/app-text-editor/p-editor/div/div[2]/div[1]"
elemento_chamado_encontrado_page_inicial = "/html/body/div/div[4]/div/div/div/div[2]/div/div/div/ul"
elemento_chamado_encontrado_page_chamado = "/html/body/div/div/div/div/div/div[2]/div/div/div/ul"
elemento_classificacao = "/html/body/app-root/app-attendant-main/div/div/div/div/app-opening/app-ticket-opening/div/app-ticket-opening-form/div[1]/app-portlet/form/div/div[2]/app-form-renderer/div/app-section/nz-collapse/div/nz-collapse-panel/div[2]/div/div/app-contextual-faq-container[3]/div/div/app-dropdown/span/span[1]/span"
elemento_chamado_selecionado = "/html/body/app-root/app-attendant-main/div/div/div/div/app-view/div/div/div[1]/app-main-portlet/app-portlet/form/div/div[1]/div[1]/div/h3/span"
elemento_codigo_cliente = "/html/body/app-root/app-attendant-main/div/div/div/div/app-opening/app-ticket-opening/div/app-ticket-opening-form/div[1]/app-portlet/form/div/div[2]/app-form-renderer/div/app-section/nz-collapse/div/nz-collapse-panel/div[2]/div/div/app-contextual-faq-container[5]/div/div/input"
elemento_data_pagamento = "/html/body/app-root/app-attendant-main/div/div/div/div/app-opening/app-ticket-opening/div/app-ticket-opening-form/div[1]/app-portlet/form/div/div[2]/app-form-renderer/div/app-section/nz-collapse/div/nz-collapse-panel/div[2]/div/div/app-contextual-faq-container[10]/div/div/app-calendar/div/input"
elemento_documento_sap = "/html/body/app-root/app-attendant-main/div/div/div/div/app-opening/app-ticket-opening/div/app-ticket-opening-form/div[1]/app-portlet/form/div/div[2]/app-form-renderer/div/app-section/nz-collapse/div/nz-collapse-panel/div[2]/div/div/app-contextual-faq-container[8]/div/div/input"
elemento_entrar = "/html/body/app-root/app-login/div/div[2]/div/app-portlet[1]/div/div/div/form/app-generic-login-button/button"
elemento_empresa = "/html/body/app-root/app-attendant-main/div/div/div/div/app-opening/app-ticket-opening/div/app-ticket-opening-form/div[1]/app-portlet/form/div/div[2]/app-form-renderer/div/app-section/nz-collapse/div/nz-collapse-panel/div[2]/div/div/app-contextual-faq-container[4]/div/div/app-dropdown/span/span[1]/span"
elemento_fornecedores_clientes = "/html/body/app-root/app-attendant-main/div/div/div/div/app-opening/app-ticket-opening/div/app-ticket-opening-form/div[1]/app-portlet/form/div/div[2]/app-form-renderer/div/app-section/nz-collapse/div/nz-collapse-panel/div[2]/div/div/app-contextual-faq-container[2]/div/div/app-radio[1]/label"
elemento_forma_de_pagamento = "/html/body/app-root/app-attendant-main/div/div/div/div/app-opening/app-ticket-opening/div/app-ticket-opening-form/div[1]/app-portlet/form/div/div[2]/app-form-renderer/div/app-section/nz-collapse/div/nz-collapse-panel/div[2]/div/div/app-contextual-faq-container[11]/div/div/app-dropdown/span/span[1]/span"
elemento_login = "/html/body/app-root/app-login/div/div[2]/div/app-portlet[1]/div/div/div/form/div[2]/input"
elemento_login_incorreto = "/html/body/app-root/app-login/div/div[2]/div/app-portlet[1]/div/div/div/div[1]"
elemento_natureza = "/html/body/app-root/app-attendant-main/div/div/div/div/app-opening/app-ticket-opening/div/app-ticket-opening-form/div[1]/app-portlet/form/div/div[2]/div[4]/div[2]/app-dropdown/span/span[1]/span"
elemento_natureza_fechar = "/html/body/app-root/app-attendant-main/div/div/div/div/app-view/div/div/div[2]/div[1]/div/app-details-portlet/app-portlet/form/div/div[2]/div[3]/app-dropdown/span/span[1]/span/span[1]/span[1]"
elemento_natureza_fechar2 = "/html/body/app-root/app-attendant-main/div/div/div/div/app-view/div/div/div[2]/div[1]/div/app-involved-portlet/app-portlet/form/div/div[2]/div[3]/app-dropdown/span/span[1]/span/span[1]/span[1]"
elemento_natureza_expand = "/html/body/app-root/app-attendant-main/div/div/div/div/app-view/div/div/div[2]/div[1]/div/app-details-portlet/app-portlet/form/div/div[2]/div[3]/app-dropdown/span/span[1]/span/span[2]"
elemento_nota_fiscal = "/html/body/app-root/app-attendant-main/div/div/div/div/app-opening/app-ticket-opening/div/app-ticket-opening-form/div[1]/app-portlet/form/div/div[2]/app-form-renderer/div/app-section/nz-collapse/div/nz-collapse-panel/div[2]/div/div/app-contextual-faq-container[7]/div/div/input"
elemento_nome_fornecedor_cliente = "/html/body/app-root/app-attendant-main/div/div/div/div/app-opening/app-ticket-opening/div/app-ticket-opening-form/div[1]/app-portlet/form/div/div[2]/app-form-renderer/div/app-section/nz-collapse/div/nz-collapse-panel/div[2]/div/div/app-contextual-faq-container[6]/div/div/input"
elemento_novo_tramite = "/html/body/app-root/app-attendant-main/div/div/div/div/app-view/div/div/div[1]/app-ticket-tabs/div/app-portlet/div/div[2]/div/div/button"
elemento_novo_tramite_chamado_aberto = "/html/body/app-root/app-attendant-main/div/div/div/div/app-view/div/div/div[1]/app-ticket-tabs/div/app-portlet/div/div[2]/div/div[1]/button"
elemento_observacoes = "/html/body/app-root/app-attendant-main/div/div/div/div/app-opening/app-ticket-opening/div/app-ticket-opening-form/div[1]/app-portlet/form/div/div[2]/app-form-renderer/div/app-section/nz-collapse/div/nz-collapse-panel/div[2]/div/div/app-contextual-faq-container[3]/div/div/textarea"
elemento_popup_status_chamado = "/html/body/div/div[2]/div/mat-dialog-container/div/div[2]/form/div/div[1]/app-dropdown/span/span[1]/span"
elemento_responsavel = "/html/body/app-root/app-attendant-main/div/div/div/div/app-view/div/div/div[2]/div[1]/div/app-involved-portlet/app-portlet/form/div/div[2]/div[3]/app-dropdown/span/span[1]/span/span[1]/span[1]"
elemento_responsavel2 = "/html/body/app-root/app-attendant-main/div/div/div/div/app-view/div/div/div[2]/div[1]/div/app-involved-portlet/app-portlet/form/div/div[2]/div[3]/app-dropdown/span/span[1]/span/span[1]"
elemento_responsavel_busca = "/html/body/app-root/app-attendant-main/div/div/div/div/app-view/div/div/div[2]/div[1]/div/app-involved-portlet/app-portlet/form/div/div[2]/div[3]/app-dropdown/span/span[1]/span/span[1]"
elemento_responsavel_selecao = "/html/body/span/span/span[1]/input"
elemento_senha = "/html/body/app-root/app-login/div/div[2]/div/app-portlet[1]/div/div/div/form/div[3]/input"
elemento_pagamento_normal = "/html/body/span/span/span[2]/ul/li[1]/span"
elemento_popup_chamado_aberto = "/html/body/div/div[2]/div/mat-dialog-container"
elemento_popup_confirmar = "/html/body/div/div[2]/div/mat-dialog-container/div/div[3]/div/button[2]"
elemento_servico = "/html/body/app-root/app-attendant-main/div/div/div/div/app-opening/app-ticket-opening/div/app-ticket-opening-form/div[1]/app-portlet/form/div/div[2]/div[3]/div/app-service-dropdown-tree/app-dropdown-tree/div"
elemento_status_chamado = "/html/body/app-root/app-attendant-main/div/div/div/div/app-view/div/div/div[2]/div[1]/div/app-details-portlet/app-portlet/form/div/div[2]/div[1]/app-dropdown/span/span[1]/span"
elemento_solic_pag_nacional = "/html/body/div/div/div/div/div/div[2]/div/div/div/app-tree/div/div/nz-tree/ul/nz-tree-node/li/ul/nz-tree-node/li/ul/nz-tree-node/li/div/span[1]"
elemento_salvar = "/html/body/app-root/app-attendant-main/div/div/div/div/app-opening/app-ticket-opening/div/app-ticket-opening-form/div[1]/app-portlet/form/div/div[3]/div/div/button[3]"
elemento_salvar_chamado = "/html/body/app-root/app-attendant-main/div/div/div/div/app-toolbar/div/div[1]/app-save-toolbar-button/app-toolbar-button"
elemento_texto_chamado = "/html/body/div[2]/div[2]/div/mat-dialog-container/div/div[2]/div[1]/app-text-editor/p-editor/div/div[2]/div[1]"
elemento_valor_total = "/html/body/app-root/app-attendant-main/div/div/div/div/app-opening/app-ticket-opening/div/app-ticket-opening-form/div[1]/app-portlet/form/div/div[2]/app-form-renderer/div/app-section/nz-collapse/div/nz-collapse-panel/div[2]/div/div/app-contextual-faq-container[9]/div/div/input"
elemento_lista_suspensa = "/html/body/span/span/span[1]/input"
elemento_severidade = "/html/body/app-root/app-attendant-main/div/div/div/div/app-opening/app-ticket-opening/div/app-ticket-opening-form/div[1]/app-portlet/form/div/div[2]/div[4]/div[3]"
elemento_pgto_agrupado = "/html/body/span/span/span[2]/ul/li[2]"

Form_Ellevo.Show
'Acesso ao portal Ellevo
driver.Get "https://electrolux.ellevo.com/"
driver.Window.Maximize

For contador = 1 To 50
    If contador < 60 Then
        If driver.IsElementPresent(by.xpath(elemento_login)) Then
            If driver.FindElementByXPath(elemento_login).IsDisplayed And driver.FindElementByXPath(elemento_login).IsEnabled Then
                Exit For
            Else
                Application.Wait (Now + TimeValue("00:00:01"))
                contador = contador + 1
            End If
        Else
            Application.Wait (Now + TimeValue("00:00:01"))
            contador = contador + 1
        End If
    Else
        MsgBox "Verifique sua conexão, não foi possível carregar a página da Ellevo", vbOKOnly
        End
    End If
Next contador

driver.FindElementByXPath(elemento_login).SendKeys Form_Ellevo.txtbox_login_ellevo
driver.FindElementByXPath(elemento_senha).SendKeys Form_Ellevo.txtbox_senha_ellevo
driver.FindElementByXPath(elemento_entrar).Click

Application.Wait (Now + TimeValue("00:00:02"))

If driver.IsElementPresent(by.xpath(elemento_login_incorreto)) Then
    If driver.FindElementByXPath(elemento_login_incorreto).text = "Usuário e/ou Senha Inválidos." Then
        MsgBox "Usuário e/ou Senha Inválidos, por favor, verifique.", vbOKOnly
        End
    End If
End If


End Sub

Public Function VerificarElemento(driver As EdgeDriver, ByRef tipo_verificacao, ByVal xpath As String)

Dim contador_de_segundos As Integer
Dim elemento_a_verificar As WebElement
Dim by As New by

    contador_de_segundos = 1
    
    ' botão de novo trâmite
    
    If tipo_verificacao = "ENABLED" Then
        On Error Resume Next
        Do Until driver.FindElementByXPath(xpath).IsEnabled
            If contador_de_segundos < 20 Then
                Application.Wait (Now + TimeValue("00:00:01"))
                contador_de_segundos = contador_de_segundos + 1
            Else
                VerificarElemento = False
                Exit Function
            End If
            
        Loop
        On Error GoTo 0
    ElseIf tipo_verificacao = "DISPLAYED" Then
        On Error Resume Next
        Do Until driver.FindElementByXPath(xpath).IsDisplayed
            If contador_de_segundos < 20 Then
                Application.Wait (Now + TimeValue("00:00:01"))
                contador_de_segundos = contador_de_segundos + 1
            Else
                VerificarElemento = False
                Exit Function
            End If
        Loop
        On Error GoTo 0
    ElseIf tipo_verificacao = "PRESENT" Then
        On Error Resume Next
        Do Until driver.IsElementPresent(by.xpath(xpath))
            If contador_de_segundos < 20 Then
                Application.Wait (Now + TimeValue("00:00:01"))
                contador_de_segundos = contador_de_segundos + 1
            Else
                VerificarElemento = False
                Exit Function
            End If
        Loop
        On Error GoTo 0
    ElseIf tipo_verificacao = "ENABLED_DISPLAYED_PRESENT" Then
        On Error Resume Next
        Do Until driver.IsElementPresent(by.xpath(xpath))
            If contador_de_segundos < 20 Then
                Application.Wait (Now + TimeValue("00:00:01"))
                contador_de_segundos = contador_de_segundos + 1
            Else
                VerificarElemento = False
                Exit Function
            End If
        Loop
        Do Until driver.FindElementByXPath(xpath).IsEnabled And driver.FindElementByXPath(xpath).IsDisplayed
            If contador_de_segundos < 20 Then
                Application.Wait (Now + TimeValue("00:00:01"))
                contador_de_segundos = contador_de_segundos + 1
            Else
                VerificarElemento = False
                Exit Function
            End If
        Loop
        On Error GoTo 0
    ElseIf tipo_verificacao = "ENABLED_DISPLAYED" Then
        On Error Resume Next
        Do Until driver.FindElementByXPath(xpath).IsEnabled And driver.FindElementByXPath(xpath).IsDisplayed
            If contador_de_segundos < 20 Then
                Application.Wait (Now + TimeValue("00:00:01"))
                contador_de_segundos = contador_de_segundos + 1
            Else
                VerificarElemento = False
                Exit Function
            End If
        Loop
        On Error GoTo 0
    ElseIf tipo_verificacao = "ENABLED_PRESENT" Then
        On Error Resume Next
        Do Until driver.FindElementByXPath(xpath).IsEnabled And driver.IsElementPresent(by.xpath(xpath))
            If contador_de_segundos < 20 Then
                Application.Wait (Now + TimeValue("00:00:01"))
                contador_de_segundos = contador_de_segundos + 1
            Else
                VerificarElemento = False
                Exit Function
            End If
        Loop
        On Error GoTo 0
    ElseIf tipo_verificacao = "DISPLAYED_PRESENT" Then
        On Error Resume Next
        Do Until driver.FindElementByXPath(xpath).IsDisplayed And driver.IsElementPresent(by.xpath(xpath))
            If contador_de_segundos < 20 Then
                Application.Wait (Now + TimeValue("00:00:01"))
                contador_de_segundos = contador_de_segundos + 1
            Else
                VerificarElemento = False
                Exit Function
            End If
        Loop
        On Error GoTo 0
    End If

    VerificarElemento = True


End Function
Public Function VerificarItemListaSuspensa(ByVal driver As Object, xpath_parte1 As String, xpath_parte2 As String, texto_procurado As String, acao As String, counter_inicial As Integer, counter_final As Integer) As Boolean
Dim x As Integer
    
    VerificarItemListaSuspensa = False
    For x = counter_inicial To counter_final
        Debug.Print driver.FindElementByXPath(xpath_parte1 & x & xpath_parte2).text
        If driver.FindElementByXPath(xpath_parte1 & x & xpath_parte2).text = texto_procurado Then
            If acao = "CLICK" Then
                driver.FindElementByXPath(xpath_parte1 & x & xpath_parte2).Click
            ElseIf acao = "CLICKDOUBLE" Then
                driver.FindElementByXPath(xpath_parte1 & x & xpath_parte2).ClickDouble
            End If
            VerificarItemListaSuspensa = True
            Exit For
        End If
    Next x
End Function
