Attribute VB_Name = "PASSO_1_ELLEVO"
Option Explicit
Private elemento_campo_observacoes, elementoinput As WebElement
Private valor_reembolsos As Single
Private resultado As String
Public chamado_criado As String
'A BASE DE DADOS PARA ENVIAR NA ELLEVO É AQUILO QUE ESTÁ NA BASE REEMBOLSOS APROVADOS
Sub ELLEVO()

Dim by As New by
Dim chamado, texto, cod_cliente, nota_fiscal, documento_sap, nome_cliente, chamado_criado As String
Dim range_linhas As Range
Dim contador As Integer

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    ' declarando todas as vars já que é o começo da FASE 3 do processo
    Call declaracao_vars
    
    ' atualizando apenas por precaucao a aba de reembolsos aprovados pois é a aba que devem ser copiadas as linhas para criacao do chamado ellevo
    Call LimparFiltros(tabela_reembolsos_aprovados)
    tabela_reembolsos_aprovados.QueryTable.BackgroundQuery = False
    tabela_reembolsos_aprovados.QueryTable.Refresh
    
    Call verificar_linhas_reembolsos_aprovados
    
    Call Variaveis_Ellevo
    data_agrupado_pagamento = Form_Ellevo.txt_box_data_agrupado_pgto_ellevo
    aba_reembolsos_aprovados.Range("BC1").Value = data_agrupado_pagamento
    
    ' a etapa não irá rodar caso a aba esteja com uma base vazia
    If aba_reembolsos_aprovados.Range("A2").Value = "" Then
        MsgBox "Nenhum chamado de reembolso a ser criado, a aba de Reembolsos Aprovados está vazia.", vbOKOnly
        End
    End If
    
    ' apagando o conteudo da coluna AB que é a coluna de verificacao de linhas já incluídas na base histórica
    ' isso para evitar que linhas que já foram processadas na Ellevo como reembolso sejam processadas novamente
    linha_fim = aba_reembolsos_aprovados.Range("A1048576").End(xlUp).Row
    
    ' verificando se as linhas da aba de reembolsos aprovados são novas, se sim, na coluna AB terão valor vazio,
    ' caso contrário serão preenchidas com "Sim" para Processadas Anteriormente
    Call LimparFiltros(tabela_aba_base_historica)
    linha_fim_base_historica = aba_base_historica.Range("A1048576").End(xlUp).Row
    
    For linha = 2 To linha_fim
        If aba_reembolsos_aprovados.Range("AC" & linha).Value <> "Sim" Then
            valor_reembolsos = valor_reembolsos + Abs(aba_reembolsos_aprovados.Range("P" & linha).Value)
        End If
    Next linha
    
    
    ' setando variaveis fixas caso seja apenas uma linha de processamento para a etapa de preenchimento
    ' dos dados no formulário de abertura de chamado da ellevo
    If linha_fim = 2 And aba_reembolsos_aprovados.Range("B2").Value <> "Sim" Then
        cod_cliente = aba_reembolsos_aprovados.Range("B2").Value
        nome_cliente = aba_reembolsos_aprovados.Range("C2").Value
        nota_fiscal = aba_reembolsos_aprovados.Range("I2").Value
    End If
    


    ' começo da execução dentro da ellevo
    ' espera a página carregar para começar as ações
    Call VerificarElemento(driver, "ENABLED_DISPLAYED_PRESENT", elemento_barra_busca_chamado)
    
    driver.Get "https://electrolux.ellevo.com/attendant/ticket/opening"
    
    driver.ExecuteScript "document.body.style.zoom='50%';"

    Application.Wait (Now + TimeValue("00:00:02"))

    ' começa a preencher e clicar nos diferentes elementos da Ellevo para abertura do tipo de chamado correto
    
    driver.FindElementByXPath(elemento_servico).Click
    driver.FindElementByXPath(elemento_barra_busca_servicos).SendKeys "Solicitação de pagamento nacional"
    contador = 0
    Do Until driver.FindElementByXPath(elemento_solic_pag_nacional).text = "Solicitação de pagamento nacional"
        If contador > 20 Then
            MsgBox "Não foi possível carregar a página da Ellevo. Por favor, verifique sua conexão.", vbOKOnly
            End
        End If
        Application.Wait (Now + TimeValue("00:00:01"))
        contador = contador + 1
    Loop
        
    driver.FindElementByXPath(elemento_solic_pag_nacional).Click
selecionar_pgto_normal_novamente:
    driver.FindElementByXPath(elemento_natureza).Click

    Application.Wait (Now + TimeValue("00:00:02"))
    Call VerificarElemento(driver, "ENABLED_DISPLAYED", elemento_pagamento_normal)
    driver.FindElementByXPath(elemento_pagamento_normal).Click
    
    contador = 1
    For i = 1 To 10
        If Not driver.IsElementPresent(by.xpath(elemento_fornecedores_clientes)) Then
            driver.FindElementByXPath(elemento_natureza).Click
            Application.Wait (Now + TimeValue("00:00:01"))
            Call VerificarElemento(driver, "ENABLED_DISPLAYED", elemento_pagamento_normal)
            driver.FindElementByXPath(elemento_pagamento_normal).Click
        ElseIf contador = 10 Then
            Exit For
        End If
    Next i
    
    driver.FindElementByXPath(elemento_fornecedores_clientes).Click
    driver.FindElementByXPath(elemento_classificacao).Click
clicar_pagamento_agrupado:
    driver.FindElementByXPath(elemento_severidade).Click
    Application.Wait (Now + TimeValue("00:00:02"))
    If Not driver.FindElementByXPath(elemento_pgto_agrupado).IsDisplayed Then
        GoTo clicar_pagamento_agrupado
    End If
    driver.FindElementByXPath(elemento_pgto_agrupado).Click
    If Not driver.IsElementPresent(by.xpath(elemento_fornecedores_clientes)) Then
        GoTo selecionar_pgto_normal_novamente
    End If
    driver.FindElementByXPath(elemento_fornecedores_clientes).Click
    driver.FindElementByXPath(elemento_classificacao).Click
    driver.FindElementByXPath(elemento_lista_suspensa).Click
    driver.FindElementByXPath(elemento_lista_suspensa).SendKeys "Reembolso clientes (OTC)"
        
clicar_reembolso_clientes_novamente:
    If Not VerificarItemListaSuspensa("/html/body/span/span/span[2]/ul/li[", "]", "Reembolso clientes (OTC)", "CLICKDOUBLE", 1, 30) Then
        GoTo clicar_reembolso_clientes_novamente
    End If
    driver.FindElementByXPath(elemento_empresa).Click
    
        
    contador = 1
    Do Until VerificarItemListaSuspensa("/html/body/span/span/span[2]/ul/li[", "]", "BR10", "CLICKDOUBLE", 1, 5) Or contador = 10
        Application.Wait (Now + TimeValue("00:00:01"))
        contador = contador + 1
    Loop
        
    If linha_fim > 2 Then
        driver.FindElementByXPath(elemento_codigo_cliente).Click
        driver.FindElementByXPath(elemento_codigo_cliente).SendKeys "VARIOS"
        driver.FindElementByXPath(elemento_nome_fornecedor_cliente).Click
        driver.FindElementByXPath(elemento_nome_fornecedor_cliente).SendKeys "VARIOS"
        driver.FindElementByXPath(elemento_nota_fiscal).Click
        driver.FindElementByXPath(elemento_nota_fiscal).SendKeys "VARIOS"
        driver.FindElementByXPath(elemento_documento_sap).Click
        driver.FindElementByXPath(elemento_documento_sap).SendKeys "VARIOS"
    Else
        driver.FindElementByXPath(elemento_codigo_cliente).Click
        driver.FindElementByXPath(elemento_codigo_cliente).SendKeys cod_cliente
        driver.FindElementByXPath(elemento_nome_fornecedor_cliente).Click
        driver.FindElementByXPath(elemento_nome_fornecedor_cliente).SendKeys nome_cliente
        driver.FindElementByXPath(elemento_nota_fiscal).Click
        driver.FindElementByXPath(elemento_nota_fiscal).SendKeys nota_fiscal
        driver.FindElementByXPath(elemento_documento_sap).Click
        driver.FindElementByXPath(elemento_documento_sap).SendKeys nota_fiscal
    End If
    
    driver.FindElementByXPath(elemento_valor_total).Click
    driver.FindElementByXPath(elemento_valor_total).SendKeys valor_reembolsos
    driver.FindElementByXPath(elemento_data_pagamento).Click
    driver.FindElementByXPath(elemento_data_pagamento).SendKeys data_agrupado_pagamento
    driver.FindElementByXPath(elemento_forma_de_pagamento).Click
        
    contador = 1
    Do Until VerificarItemListaSuspensa("/html/body/span/span/span[2]/ul/li[", "]", "TED", "CLICKDOUBLE", 1, 5) Or contador = 10
        Application.Wait (Now + TimeValue("00:00:01"))
        contador = contador + 1
    Loop
    Call LimparFiltros(tabela_reembolsos_aprovados)
    On Error Resume Next
    tabela_reembolsos_aprovados.Range.AutoFilter Field:=29, Criteria1:=""
    Set range_linhas = aba_reembolsos_aprovados.Range("A1:AA" & linha_fim).SpecialCells(xlCellTypeVisible)
    resultado = RangeParaString(range_linhas)
    On Error GoTo 0

' se a len(resultado) for menor que 1802 (limite maximo de caracteres no campo observacoes) a automacao preenche no proprio campo de observacos
' as informações que devem ser copiadas e coladas no excel, caso contrário, abre o chamado criado e cria um tramite com as informações
        If Len(resultado) <= 1802 Then
verificar_pop_up_novamente_1:
            For i = 1 To 20
                If driver.IsElementPresent(by.xpath("/html/body/app-root/app-attendant-main/div/div/div/div/app-opening/app-ticket-opening/div/app-ticket-opening-form/div[1]/app-portlet/form/div/div[2]/app-form-renderer/div/app-section/nz-collapse/div/nz-collapse-panel/div[2]/div/div/app-contextual-faq-container[" & i & "]/div/div/textarea")) Then
                    Set elementoinput = driver.FindElementByXPath("/html/body/app-root/app-attendant-main/div/div/div/div/app-opening/app-ticket-opening/div/app-ticket-opening-form/div[1]/app-portlet/form/div/div[2]/app-form-renderer/div/app-section/nz-collapse/div/nz-collapse-panel/div[2]/div/div/app-contextual-faq-container[" & i & "]/div/div/textarea")
                    elemento_campo_observacoes = "/html/body/app-root/app-attendant-main/div/div/div/div/app-opening/app-ticket-opening/div/app-ticket-opening-form/div[1]/app-portlet/form/div/div[2]/app-form-renderer/div/app-section/nz-collapse/div/nz-collapse-panel/div[2]/div/div/app-contextual-faq-container[" & i & "]/div/div/textarea"
                    If elementoinput.IsDisplayed Then
                        Call VerificarElemento(driver, "ENABLED_DISPLAYED", elemento_campo_observacoes)
                        Application.Wait (Now + TimeValue("00:00:03"))
                        elementoinput.Click
                        elementoinput.SendKeys "Segue abaixo listagem das linhas a serem coladas no Excel (utilizar texto para colunas com o delimitador '|')." & _
                            vbNewLine & vbNewLine & resultado
                        Exit For
                    End If
                End If
            Next i
            driver.FindElementByXPath(elemento_salvar).Click
            
            contador = 1
            Do Until driver.IsElementPresent(by.xpath(elemento_popup_chamado_aberto)) Or contador = 10
                Application.Wait (Now + TimeValue("00:00:01"))
                contador = contador + 1
            Loop
            contador = 1
            Do Until VerificarElemento(driver, "ENABLED_DISPLAYED", elemento_popup_chamado_aberto) Or contador = 10
                Application.Wait (Now + TimeValue("00:00:01"))
                contador = contador + 1
            Loop
            If driver.FindElementByXPath(elemento_popup_chamado_aberto).IsDisplayed And driver.FindElementByXPath(elemento_popup_chamado_aberto).IsEnabled Then
                For i = 1 To 10
                    If driver.IsElementPresent(by.xpath("/html/body/div[" & i & "]/div[2]/div/mat-dialog-container/div/div[2]/div/div[1]/h5")) Then
                        ' verificando o texto do elemento que abriga o numero do chamado criado
                        If driver.FindElementByXPath("/html/body/div[" & i & "]/div[2]/div/mat-dialog-container/div/div[2]/div/div[1]/h5").IsDisplayed Then
                            chamado_criado = driver.FindElementByXPath("/html/body/div[" & i & "]/div[2]/div/mat-dialog-container/div/div[2]/div/div[1]/h5").text
                            aba_reembolsos_aprovados.Range("BB1").Value = Replace(chamado_criado, "Chamado", "ELLEVO")
                            Exit For
                        End If
                    End If
                Next i
            Else
                GoTo verificar_pop_up_novamente_1
            End If
            
        ElseIf Len(resultado) > 1802 Then
verificar_pop_up_novamente_2:
            driver.FindElementByXPath(elemento_salvar).Click
            
            Do Until VerificarElemento(driver, "PRESENT", elemento_popup_chamado_aberto) Or contador = 10
                Application.Wait (Now + TimeValue("00:00:01"))
                contador = contador + 1
            Loop
            contador = 1
            Do Until VerificarElemento(driver, "ENABLED_DISPLAYED", elemento_popup_chamado_aberto) Or contador = 10
                Application.Wait (Now + TimeValue("00:00:01"))
                contador = contador + 1
            Loop
            If driver.FindElementByXPath(elemento_popup_chamado_aberto).IsDisplayed And driver.FindElementByXPath(elemento_popup_chamado_aberto).IsEnabled Then
                ' botão de visualizar chamado
                For i2 = 1 To 10
                    If driver.IsElementPresent(by.xpath("/html/body/div[" & i2 & "]/div[2]/div/mat-dialog-container/div/app-crud-footer/div/div/div[2]/button")) Then
                        contador = 1
                        Do Until VerificarItemListaSuspensa("/html/body/div[", "]/div[2]/div/mat-dialog-container/div/app-crud-footer/div/div/div[2]/button", "VISUALIZAR CHAMADO", "CLICK", 1, 10) Or contador = 10
                           contador = contador + 1
                        Loop
                        Call VerificarElemento(driver, "ENABLED_DISPLAYED", "/html/body/div[" & i2 & "]/div[2]/div/mat-dialog-container/div/app-crud-footer/div/div/div[2]/button")
                        Call VerificarElemento(driver, "ENABLED_DISPLAYED", "/html/body/div[" & i2 & "]/div[2]/div/mat-dialog-container/div/app-crud-footer/div/div/div[2]/button")
                        Call VerificarElemento(driver, "ENABLED_DISPLAYED", "/html/body/div[" & i2 & "]/div[2]/div/mat-dialog-container/div/app-crud-footer/div/div/div[2]/button")
                        Call VerificarElemento(driver, "ENABLED_DISPLAYED", "/html/body/div[" & i2 & "]/div[2]/div/mat-dialog-container/div/app-crud-footer/div/div/div[2]/button")
                        ' verificando o texto do elemento que abriga o numero do chamado criado
                        chamado_criado = driver.FindElementByXPath("/html/body/div[" & i & "]/div[2]/div/mat-dialog-container/div/div[2]/div/div[1]/h5").text
                        driver.FindElementByXPath("/html/body/div[" & i2 & "]/div[2]/div/mat-dialog-container/div/app-crud-footer/div/div/div[2]/button").Click
                        Application.Wait (Now + TimeValue("00:00:05"))
                        aba_reembolsos_aprovados.Range("BB1").Value = Replace(chamado_criado, "Chamado", "ELLEVO")
                        Exit For
                    End If
                Next i2
            Else
                GoTo verificar_pop_up_novamente_2
            End If
        
            ' botão de novo trâmite
            Call VerificarElemento(driver, "ENABLED_DISPLAYED", elemento_novo_tramite_chamado_aberto)
            driver.FindElementByXPath(elemento_novo_tramite_chamado_aberto).Click
            ' botão de inserir texto
            For i2 = 1 To 10
                If driver.IsElementPresent(by.xpath("/html/body/div[" & i2 & "]/div[2]/div/mat-dialog-container/div/div[2]/div[1]/app-text-editor/p-editor/div/div[2]/div[1]")) Then
                    Set elementoinput = driver.FindElementByXPath("/html/body/div[" & i2 & "]/div[2]/div/mat-dialog-container/div/div[2]/div[1]/app-text-editor/p-editor/div/div[2]/div[1]")
                    Do Until elementoinput.IsDisplayed
                        Call VerificarElemento(driver, "ENABLED", elemento_novo_tramite_chamado_aberto)
                        driver.FindElementByXPath(elemento_novo_tramite_chamado_aberto).Click
                    Loop
                    driver.FindElementByXPath("/html/body/div[" & i2 & "]/div[2]/div/mat-dialog-container/div/div[2]/div[1]/app-text-editor/p-editor/div/div[2]/div[1]").Click
                    driver.FindElementByXPath("/html/body/div[" & i2 & "]/div[2]/div/mat-dialog-container/div/div[2]/div[1]/app-text-editor/p-editor/div/div[2]/div[1]").Clear
                    driver.FindElementByXPath("/html/body/div[" & i2 & "]/div[2]/div/mat-dialog-container/div/div[2]/div[1]/app-text-editor/p-editor/div/div[2]/div[1]").SendKeys "Segue abaixo listagem das linhas a serem coladas no Excel (utilizar texto para colunas com o delimitador '|')." & vbNewLine & vbNewLine & resultado
                    ' botão confirmar tramite
                    driver.FindElementByXPath("/html/body/div[" & i2 & "]/div[2]/div/mat-dialog-container/div/div[3]/div/button[2]").Click
                    Exit For
                End If
            Next i2
        End If

    driver.Quit
    
    ' chama a sub explicitamente para enviar os e-mails aos clientes referente a esses reembolsos recém solicitados na ellevo
    Call alterar_status_reembolsos_chamado_ellevo_criado
    Call preencher_reembolsos_base_historica
    

    Call emails_etapa_1
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

' sub especifica para apagar da aba ag aprovacao aquilo que já foi aprovado conforme constar na aba aprovados
' e passar base a base historica com o status de processamento igual a "Reembolso"
Sub verificar_linhas_reembolsos_aprovados()
    
    linha_fim = aba_reembolsos_aprovados.Range("A1048576").End(xlUp).Row
    linha_fim_base_historica = aba_base_historica.Range("A1048576").End(xlUp).Row
    For linha = 2 To linha_fim
        referencia = aba_reembolsos_aprovados.Range("F" & linha).Value
        num_doc = aba_reembolsos_aprovados.Range("G" & linha).Value
        item = aba_reembolsos_aprovados.Range("H" & linha).Value
        For i = 2 To linha_fim_base_historica
            If referencia = aba_base_historica.Range("F" & i).Value And _
                num_doc = aba_base_historica.Range("G" & i).Value And _
                    item = aba_base_historica.Range("H" & i).Value Then
  
                    aba_reembolsos_aprovados.Range("AC" & linha).Value = "Sim"
                    Exit For
                    
            End If
        Next i
    Next linha

End Sub
Sub preencher_reembolsos_base_historica()

    Call declaracao_vars

    Dim linha_fim_base_historica2 As Integer
    
    linha_fim_aba_reembolsos_pendentes = aba_reembolsos_pendentes.Range("A1048576").End(xlUp).Row
    
    tabela_aba_reembolsos_pendentes.Range.AutoFilter Field:=31, Criteria1:="Ellevo Criado"
    
    aba_reembolsos_pendentes.Range("A2:AB" & linha_fim_aba_reembolsos_pendentes).SpecialCells(xlCellTypeVisible).Copy
                    
    ' ADICIONANDO LINHA DE ABATIMENTO PARCIAL NA BASE HISTORICA
    linha_fim_base_historica = aba_base_historica.Range("A1048576").End(xlUp).Row
    
    aba_reembolsos_pendentes.Range("A2" & ":AB" & linha_fim_aba_reembolsos_pendentes).Copy
    If aba_base_historica.Range("A" & linha_fim_base_historica).Value <> "" Then
        linha_fim_base_historica = linha_fim_base_historica + 1
    End If
    aba_base_historica.Range("A" & linha_fim_base_historica).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    linha_fim_base_historica2 = aba_base_historica.Range("A1048576").End(xlUp).Row
    
    aba_base_historica.Range("AC" & linha_fim_base_historica & ":AC" & linha_fim_base_historica2).Value = "Reembolso"
    aba_base_historica.Range("AD" & linha_fim_base_historica & ":AD" & linha_fim_base_historica2).Value = Date

End Sub

Sub alterar_status_reembolsos_chamado_ellevo_criado()

    Dim status_processado_anteriormente As String
    array_docs_F65 = Array()
    linha_fim = aba_reembolsos_aprovados.Range("A1048576").End(xlUp).Row
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For linha = 2 To linha_fim
        doc_f65 = aba_reembolsos_aprovados.Range("G" & linha).Value
        status_processado_anteriormente = aba_reembolsos_aprovados.Range("AC" & linha).Value
        If status_processado_anteriormente <> "Sim" Then
            If Not UBound(VBA.Filter(array_docs_F65, doc_f65)) >= 0 Then
                Call Add_ao_Array(array_docs_F65, doc_f65)
            End If
        End If
    Next linha
    
    Call LimparFiltros(tabela_aba_reembolsos_pendentes)
    linha_fim_aba_reembolsos_pendentes = aba_reembolsos_pendentes.Range("A1048576").End(xlUp).Row
    For linha = 2 To linha_fim_aba_reembolsos_pendentes
        doc_f65 = aba_reembolsos_pendentes.Range("AC" & linha).Value
        If UBound(VBA.Filter(array_docs_F65, doc_f65)) >= 0 Then
            aba_reembolsos_pendentes.Range("AE" & linha).Value = "Ellevo Criado"
        End If
    Next linha


End Sub

Private Function RangeParaString(range_linhas As Range) As String
    Dim celula As Range
    Dim resultado As String
    ' Inicializa a string
    resultado = ""
    
    ' Loop pelas células no intervalo
    For Each celula In range_linhas
        ' Adiciona o valor da célula à string com separador de vírgula ou outro caractere desejado
        If celula.Column = 28 Then
            resultado = resultado & celula.Value & "|" & vbNewLine
        Else
            resultado = resultado & celula.Value & "|"
        End If
    Next celula
        
    ' Remove a última vírgula e o espaço extra
    If Len(resultado) > 0 Then
        resultado = Left(resultado, Len(resultado) - 3)
        
    End If
    
    RangeParaString = resultado
    
End Function

Private Function VerificarItemListaSuspensa(xpath_parte1, xpath_parte2, texto_procurado As String, acao As String, counter_inicial As Integer, counter_final As Integer) As Boolean
Dim x As Integer
Dim elemento As WebElement
    
    VerificarItemListaSuspensa = False
    For x = counter_inicial To counter_final
        On Error Resume Next
        Set elemento = driver.FindElementByXPath(xpath_parte1 & x & xpath_parte2)
        If UCase(elemento.text) = UCase(texto_procurado) Then
            If acao = "CLICK" Then
                elemento.Click
            ElseIf acao = "CLICKDOUBLE" Then
                elemento.ClickDouble
            End If
            VerificarItemListaSuspensa = True
            Exit For
        End If
        On Error GoTo 0
    Next x
End Function

