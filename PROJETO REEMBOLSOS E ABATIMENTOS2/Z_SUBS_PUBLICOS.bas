Attribute VB_Name = "Z_SUBS_PUBLICOS"
' Esta fun��o preenche a coluna BB da planilha de consolidado com os valores de um array
' Par�metros:
'   - array_(): Array contendo os valores a serem listados
' Retorna: Nada (Sub impl�cita)
Public Sub ListarInfosColunaBB(array_() As Variant)
   Dim i6, i7 As Integer
   
   ' Tenta mostrar todos os dados em qualquer filtro de tabela que possa estar ativo
    On Error Resume Next
    tabela_aba_consolidado.AutoFilter.ShowAllData
    On Error GoTo 0
   
    ' Limpa o conte�do existente na coluna BB da planilha
    aba_consolidado.Range("BB2:BB1048576").ClearContents
    ' Come�a a preencher a partir da linha 2
    i6 = 2
    
    ' Percorre o array de valores e preenche cada c�lula na coluna BB
    For i7 = LBound(array_) To UBound(array_)
        aba_consolidado.Range("BB" & i6).Value = array_(i7)
        i6 = i6 + 1
    Next i7
   
End Sub
' essa fun��o alimenta a base historica com OCs j� processadas/verificadas e seu respectivo status para que caso
' sejam consultadas novamente, o sistema bata com a base historica e n�o processe novamente nem entre no circuito de verificacao
Public Sub AlimentarAbaHistorica(ByVal status_OC As String)

    Dim linha_fim_aba_historica As Long
    
    linha_fim_aba_historica = aba_historica.Range("A1048576").End(xlUp).Offset(1, 0).Row
    
    If linha_fim_aba_historica > 1000000 Then
        aba_historica.Rows("2:5000").Delete Shift:=xlUp
        linha_fim_aba_historica = aba_historica.Range("A1048576").End(xlUp).Offset(1, 0).Row
    End If

    
    aba_historica.Range("A" & linha_fim_aba_historica).Value = numero_OC
    aba_historica.Range("B" & linha_fim_aba_historica).Value = chamado
    aba_historica.Range("C" & linha_fim_aba_historica).Value = status_OC
    If status_OC = "ABATIMENTO" Then
        aba_historica.Range("D" & linha_fim_aba_historica).Value = Date
    ElseIf status_OC = "REEMBOLSO" Then
        aba_historica.Range("D" & linha_fim_aba_historica).Value = Form_SAP.txt_box_data_agrupado_pgto_SAP
    End If
    
    

End Sub
' preenche no array todos os clientes que tem dados bancarios
Public Sub PreencherArrayPayersCOMDadosBancarios()

    Dim linha_fim_payers_com_dados_bancarios As Long
    Dim payer As String
    
    linha_fim_payers_com_dados_bancarios = aba_dados_bancarios.Range("A1048576").End(xlUp).Row
    
    For linha = 2 To linha_fim_payers_com_dados_bancarios
        payer = aba_dados_bancarios.Range("A" & linha).Value
        If Not UBound(VBA.Filter(array_payers_com_dados_bancarios, payer)) >= 0 Then
            array_payers_com_dados_bancarios = Add_ao_Array(array_payers_com_dados_bancarios, payer)
        End If
    Next linha
    
End Sub
' essa fun��o � responsavel por adicionar e criar as respectivas chaves e valores no dicionario de relatorio de processamento
' para posteriormente, no fim da execu��o, registrar tudo o que foi processado para visualiza�ao do usuario
Public Sub AlimentarDicionario_Relatorio_Processamento(ByVal chave As String, ByVal valor As String)

    If Dicionario_Relatorio_Processamento.exists(chave) Then
        If Not InStr(1, Dicionario_Relatorio_Processamento(chave), valor, vbTextCompare) Then
            Dicionario_Relatorio_Processamento(chave) = _
                Dicionario_Relatorio_Processamento(chave) & "-" & valor
        End If
    Else
        Dicionario_Relatorio_Processamento(chave) = valor
    End If

End Sub
' Esta fun��o coleta informa��es de payers/CNPJs a partir de uma tela do SAP
' Par�metros:
'   - session_number: Objeto que representa a sess�o SAP
'   - indice_inicio: �ndice inicial vertical (Y) para come�ar a busca
'   - indice_fim: �ndice final vertical (Y) para terminar a busca
'   - id_parte_inicial: String base para o ID do elemento SAP
'   - indice_coluna: �ndice da coluna (X) onde est�o os dados de payer/CNPJ
' Retorna: Implicitamente preenche o array_payers_encontrados (vari�vel global)
Public Sub PreencherPayersCNPJ(ByVal session_number As Object, indice_inicio As Integer, ByVal indice_fim As Integer, id_parte_inicial As String, ByVal indice_coluna As Integer)
   Dim elemento_sap As Object
   Dim payer, texto_tela_cnpjs As String
   Dim qtde_payers As Integer
   Dim cnpjs_visiveis As Integer
   
   ' Inicializa o array que armazenar� payers encontrados
   array_payers_encontrados = Array()
   
   ' Cria um objeto para express�o regular
   Set regex = CreateObject("VBScript.RegExp")
   ' Define o padr�o para extrair o n�mero de cliente
   regex.Pattern = "N� cliente (\d+) Entrada"
   regex.IgnoreCase = True
   regex.Global = True
   
   ' Obt�m o texto da tela de CNPJs
   texto_tela_cnpjs = session.findById("wnd[1]").text
   ' Extrai o n�mero de clientes usando express�o regular
   Set ocorrencia = regex.Execute(texto_tela_cnpjs)
   texto_tela_cnpjs = ocorrencia(0).SubMatches(0)
   qtde_payers = CInt(texto_tela_cnpjs)
   
   ' Loop at� encontrar todos os payers esperados
   Do Until UBound(array_payers_encontrados) + 1 = qtde_payers
       ' Primeiro loop: determina quantos CNPJs est�o vis�veis na tela atual
       For i2 = indice_inicio To indice_fim
           On Error Resume Next
           Set elemento_sap = session_number.findById(id_parte_inicial & indice_coluna & "," & i2 & "]")
           elemento_sap.SetFocus
           If Err.number <> 0 Then
               ' Se ocorrer erro, encontrou o limite de CNPJs vis�veis
               cnpjs_visiveis = i2 - 1
               Exit For
           End If
           On Error GoTo 0
       Next i2
       
       ' Segundo loop: coleta os payers vis�veis
       For i2 = indice_inicio To cnpjs_visiveis
           On Error Resume Next
           Set elemento_sap = session_number.findById(id_parte_inicial & indice_coluna & "," & i2 & "]")
           On Error GoTo 0
           payer = elemento_sap.text
           ' Adiciona o payer ao array se ainda n�o estiver l�
           If UBound(VBA.Filter(array_payers_encontrados, payer)) < 0 Then
               array_payers_encontrados = Add_ao_Array(array_payers_encontrados, payer)
           Else
               ' Se encontrou um payer repetido, sai do loop (provavelmente est� na mesma p�gina)
               Exit For
           End If
       Next i2
       
       ' Avan�a para a pr�xima p�gina (tecla F22)
       session_number.findById("wnd[0]").sendVKey 82
       i2 = 3
   Loop
End Sub
Public Sub PreencherArrayDocsSap(array_() As Variant, ByVal elemento_tabela As Object, nome_coluna As String)

    ' Tratamento de erro: ignora erros para evitar interrup��es se a tabela estiver vazia ou n�o estiver totalmente carregada.
    On Error Resume Next
    ' Seleciona a primeira linha da tabela.
    elemento_tabela.selectedRows = 1
    elemento_tabela.currentCellRow = 1
    ' Verifica se ocorreu um erro ao selecionar a linha (por exemplo, tabela vazia).
    If Err.number <> 0 Then
        ' Se houver um erro, desfaz a sele��o da linha.
        elemento_tabela.selectedRows = 0
        elemento_tabela.currentCellRow = 0
        ' Obt�m o valor da c�lula na primeira linha e coluna especificada.
        num_doc_sap = elemento_tabela.GetCellValue(elemento_tabela.currentCellRow, nome_coluna)
        ' Adiciona o valor ao array.
        array_ = Add_ao_Array(array_, num_doc_sap)
        ' Sai da fun��o, pois n�o h� mais linhas para processar.
        Exit Sub
    End If
    ' Restaura o tratamento de erro padr�o.
    On Error GoTo 0

    ' Tratamento de erro: ignora erros durante o loop para continuar processando outras linhas.
    On Error Resume Next
    i2 = 0
    ' Loop atrav�s das linhas da tabela (at� um m�ximo de 150).
    For i2 = 0 To 150
        ' Seleciona a linha atual.
        elemento_tabela.selectedRows = i2
        elemento_tabela.currentCellRow = i2
        ' Verifica se ocorreu um erro ao selecionar a linha.
        If Err.number <> 0 Then
            ' Se houver um erro, sai do loop (fim das linhas da tabela).
            Exit For
        End If

        ' Obt�m o valor da c�lula na linha atual e coluna especificada.
        num_doc_sap = elemento_tabela.GetCellValue(elemento_tabela.currentCellRow, nome_coluna)
        ' Adiciona o valor ao array.
        array_ = Add_ao_Array(array_, num_doc_sap)
    Next i2
    ' Restaura o tratamento de erro padr�o.
    On Error GoTo 0

End Sub

Public Sub SetEixoXColunas()

    '***********************************************************
    ' Esta sub-rotina identifica e atribui as posi��es das colunas na interface SAP
    ' Cada vari�vel x_* armazena a posi��o de uma coluna espec�fica na tela do SAP
    ' A fun��o VerificarColuna procura por um texto espec�fico na interface para identificar colunas
    '***********************************************************
    
    ' Par�metros da fun��o VerificarColuna:
    ' 1�: Linha na interface (2 significa a linha de cabe�alho)
    ' 2�: Coluna inicial a partir da qual come�ar a busca
    ' 3�: Limite m�ximo de colunas a serem verificadas
    ' 4�: Caminho base do elemento na interface SAP
    ' 5�: Texto a ser procurado no cabe�alho da coluna

    x_doc_compensacao = VerificarColuna(2, 1, 10000, "wnd[0]/usr/lbl[", "DocCompens")
    x_data_compensacao = VerificarColuna(2, x_doc_compensacao, 100, "wnd[0]/usr/lbl[", "Compensa�.")
    x_texto = VerificarColuna(2, x_data_compensacao, 100, "wnd[0]/usr/lbl[", "Texto")
    x_tipo_doc = VerificarColuna(2, x_texto, 10000, "wnd[0]/usr/lbl[", "Tip")
    x_cliente = VerificarColuna(2, x_tipo_doc, 10000, "wnd[0]/usr/lbl[", "Cliente")
    x_nome = VerificarColuna(2, x_cliente, 10000, "wnd[0]/usr/lbl[", "Nome 1")
    x_referencia = VerificarColuna(2, x_nome, 10000, "wnd[0]/usr/lbl[", "Refer�ncia")
    x_montante = VerificarColuna(2, x_referencia, 10000, "wnd[0]/usr/lbl[", "Mont.moeda doc.")
    x_venc_liquido = VerificarColuna(2, x_montante, 10000, "wnd[0]/usr/lbl[", "VencL�quid")
    x_bloq_adv = VerificarColuna(2, x_venc_liquido, 10000, "wnd[0]/usr/lbl[", "Bloq.")
    x_chave_ref_1 = VerificarColuna(2, x_bloq_adv, 10000, "wnd[0]/usr/lbl[", "Chv.ref.1")
    x_chave_ref_2 = VerificarColuna(2, x_chave_ref_1, 10000, "wnd[0]/usr/lbl[", "Chv.ref.2")
    x_chave_ref_3 = VerificarColuna(2, x_chave_ref_2, 10000, "wnd[0]/usr/lbl[", "Chave refer�ncia 3")
    x_num_doc = VerificarColuna(2, x_chave_ref_3, 10000, "wnd[0]/usr/lbl[", "N� doc.")
    x_item = VerificarColuna(2, x_num_doc, 10000, "wnd[0]/usr/lbl[", "Itm")
    x_atribuicao = VerificarColuna(2, x_item, 10000, "wnd[0]/usr/lbl[", "Atribui��o")
    
End Sub
Public Sub PreencherArrayLinhasCondicaoAtual(ByVal session_number As Object, ByVal iterator1 As Integer, ByVal iterator2 As Integer, acao As String)
    '***********************************************************
    ' Esta sub-rotina preenche arrays com dados da interface SAP com base na condi��o atual
    ' Par�metros:
    '   session_number: Objeto da sess�o SAP
    '   iterator1: Contador para controle de navega��o pelas p�ginas
    '   iterator2: Contador para controle de navega��o pelas linhas
    '   acao: Define o tipo de opera��o ("ABATIMENTO", "VERIFICACAO" ou "CHAMADO_CTA_A_PAGAR")
    '***********************************************************
    
    ' Declara��o de vari�veis
    Dim primeiro_num_doc_item, texto_sbar As String
    Dim quantidade_linhas, linha_index_array_atual As Integer
    

    ' Cria objeto de express�o regular para extrair quantidade de partidas da barra de status
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "S�o exibidas (.*?) partidas"
    regex.IgnoreCase = True
    regex.Global = True

    ' Inicializa contador de p�gina
    qtde_page_downs = 0
    texto_sbar = session_number.findById("wnd[0]/sbar").text
    ' Extrai a quantidade de partidas usando express�o regular
    Set ocorrencia = regex.Execute(texto_sbar)
    
    
    
    ' Se houver um resultado na express�o regular
    If ocorrencia.Count > 0 Then
        texto_sbar = ocorrencia(0).SubMatches(0)
        quantidade_linhas = CInt(texto_sbar)  ' Converte para inteiro
        linhas_visiveis = VerificarQuantidadeLinhasVisiveis(session_number, 4, 100, "wnd[0]/usr/lbl[")
        
        ' Calcula quantas p�ginas ser�o necess�rias para ver todas as linhas
        If quantidade_linhas > linhas_visiveis - 3 Then
            qtde_page_downs = Application.WorksheetFunction.Floor(quantidade_linhas / (linhas_visiveis - 3), 1)
        End If
    Else
        ' Se n�o conseguiu extrair com regex, tenta verificar diretamente
        linhas_visiveis = VerificarQuantidadeLinhasVisiveis(session_number, 4, 100, "wnd[0]/usr/lbl[")
        quantidade_linhas = VerificarQuantidadeLinhasTotais(session_number, 4, 100, "wnd[0]/usr/lbl[")
        session_number.findById("wnd[0]").sendVKey 80  ' Tecla F8 (refresh)
        
        ' Calcula p�ginas necess�rias
        If quantidade_linhas = linhas_visiveis Then
            qtde_page_downs = 0
        ElseIf quantidade_linhas > linhas_visiveis Then
            qtde_page_downs = Application.WorksheetFunction.Floor(quantidade_linhas / linhas_visiveis, 1)
        End If
    End If
    

    ' Loop principal que navega por todas as p�ginas necess�rias
    iterator1 = 0
    Do Until iterator1 > qtde_page_downs
        linhas_visiveis = VerificarQuantidadeLinhasVisiveis(session_number, 4, 100, "wnd[0]/usr/lbl[")
        
        ' Loop que percorre cada linha vis�vel na p�gina atual
        For iterator2 = 4 To linhas_visiveis
            ' Inicializa array para armazenar dados da linha atual
            array_linha_atual = Array()
            
            ' Define o tamanho do array conforme a a��o
            If acao = "ABATIMENTO" Then
                ReDim array_linha_atual(18)  ' Array maior para opera��o de abatimento
            Else
                ReDim array_linha_atual(16)  ' Tamanho padr�o para outras opera��es
            End If
        
            ' Captura os dados de cada coluna da linha atual
            array_linha_atual(0) = session_number.findById("wnd[0]/usr/lbl[" & x_doc_compensacao & "," & iterator2 & "]").text
            array_linha_atual(1) = session_number.findById("wnd[0]/usr/lbl[" & x_data_compensacao & "," & iterator2 & "]").text
            array_linha_atual(2) = session_number.findById("wnd[0]/usr/lbl[" & x_texto & "," & iterator2 & "]").text
            array_linha_atual(3) = session_number.findById("wnd[0]/usr/lbl[" & x_tipo_doc & "," & iterator2 & "]").text
            array_linha_atual(4) = session_number.findById("wnd[0]/usr/lbl[" & x_cliente & "," & iterator2 & "]").text
            array_linha_atual(5) = session_number.findById("wnd[0]/usr/lbl[" & x_nome & "," & iterator2 & "]").text
            array_linha_atual(6) = session_number.findById("wnd[0]/usr/lbl[" & x_referencia & "," & iterator2 & "]").text
            ' Converte o valor monet�rio para formato num�rico (troca v�rgula por ponto e remove ponto de milhar)
            array_linha_atual(7) = CSng(VBA.Trim(Replace(Replace(session_number.findById("wnd[0]/usr/lbl[" & x_montante & "," & iterator2 & "]").text, ".", ""), ",", ".")))
            array_linha_atual(8) = session_number.findById("wnd[0]/usr/lbl[" & x_venc_liquido & "," & iterator2 & "]").text
            array_linha_atual(9) = session_number.findById("wnd[0]/usr/lbl[" & x_bloq_adv & "," & iterator2 & "]").text
            array_linha_atual(10) = session_number.findById("wnd[0]/usr/lbl[" & x_chave_ref_1 & "," & iterator2 & "]").text
            array_linha_atual(11) = session_number.findById("wnd[0]/usr/lbl[" & x_chave_ref_2 & "," & iterator2 & "]").text
            array_linha_atual(12) = session_number.findById("wnd[0]/usr/lbl[" & x_chave_ref_3 & "," & iterator2 & "]").text
            array_linha_atual(13) = session_number.findById("wnd[0]/usr/lbl[" & x_num_doc & "," & iterator2 & "]").text
            array_linha_atual(14) = session_number.findById("wnd[0]/usr/lbl[" & x_item & "," & iterator2 & "]").text
            array_linha_atual(15) = session_number.findById("wnd[0]/usr/lbl[" & x_atribuicao & "," & iterator2 & "]").text
            Dim chave_3, array_chave_atual() As Variant
            Dim iterador_dict_ZFI105 As Integer
            If Not dictOCNumeroDOC_ZFI105 Is Nothing Then
                For Each chave_3 In dictOCNumeroDOC_ZFI105.keys
                    If chave_3 = array_linha_atual(13) Then
                        numero_OC = dictOCNumeroDOC_ZFI105(chave_3)
                        array_linha_atual(16) = numero_OC
                        Exit For
                    End If
                Next chave_3
            Else
                array_linha_atual(16) = numero_OC
            End If
            
            
            
            ' Processa os dados conforme a a��o especificada
            If acao = "ABATIMENTO" Then
                ' Adiciona informa��es espec�ficas para abatimento com base no tipo de documento
                If array_linha_atual(3) = "R1" Then
                    array_linha_atual(17) = "Cr�dito utilizado"
                    array_linha_atual(18) = "-"
                ElseIf array_linha_atual(3) = "AB" Then
                    array_linha_atual(17) = "Valor Residual a Pagar ->"
                    array_linha_atual(18) = valor_residual_AB
                ElseIf array_linha_atual(3) = "RV" Then
                    If array_linha_atual(15) = "ABATIDO TOTAL" Then
                        array_linha_atual(17) = "T�tulo Abatido Integralmente"
                        array_linha_atual(18) = "-"
                    ElseIf array_linha_atual(15) = "ABATIDO PARCIAL" Then
                        array_linha_atual(17) = "T�tulo Abatido Parcialmente"
                        array_linha_atual(18) = valor_residual_AB
                    End If
                    
                End If
                
                ' Adiciona a linha atual ao array de detalhes de abatimento
                linha_index_array_atual = UBound(array_linhas_detalhe_abatimento) + 1
                ReDim Preserve array_linhas_detalhe_abatimento(LBound(array_linhas_detalhe_abatimento) To linha_index_array_atual)
                array_linhas_detalhe_abatimento(linha_index_array_atual) = array_linha_atual

            ElseIf acao = "VERIFICACAO" Then
                ' Processa dados para verifica��o, classificando-os conforme o tipo
                If array_linha_atual(0) <> "" And session_number.info.Transaction = "FBL5N" Then
                    ' Se tem documento de compensa��o e est� na transa��o FBL5N, adiciona ao array de documentos compensados
                    If UBound(VBA.Filter(array_doc_compensacoes, array_linha_atual(0))) < 0 Then
                        array_doc_compensacoes = Add_ao_Array(array_doc_compensacoes, array_linha_atual(0))
                    End If
                    linhas_compensadas = True
                    
                ElseIf array_linha_atual(0) = "" And session_number.info.Transaction = "FBL5N" Then
                    Dim x2 As Integer
                    Dim linha_ja_processada As Boolean
                    linha_ja_processada = False
                    For x2 = LBound(array_atribuicoes_proibidas) To UBound(array_atribuicoes_proibidas)
                        If UCase(array_linha_atual(15)) Like UCase(array_atribuicoes_proibidas(x2)) Then
                            linha_ja_processada = True
                            Exit For
                        End If
                    Next x2
                    
                    ' Se n�o tem documento de compensa��o e est� na transa��o FBL5N, adiciona ao array de linhas abertas
                    If Not linha_ja_processada Then
                        linha_index_array_atual = UBound(array_geral_linhas_abertas_FBL5N) + 1
                        ReDim Preserve array_geral_linhas_abertas_FBL5N(LBound(array_geral_linhas_abertas_FBL5N) To linha_index_array_atual)
                        array_geral_linhas_abertas_FBL5N(linha_index_array_atual) = array_linha_atual
                        linhas_abertas = True
                    End If
                    
                ElseIf array_linha_atual(0) <> "" And session_number.info.Transaction = "FB03" Then
                    If array_linha_atual(16) = "" And acao = "VERIFICACAO" Then
                        array_linha_atual(16) = "Outros Cr�ditos/D�bitos que comp�em a compensa��o de baixa"
                    End If
                    ' Se tem documento de compensa��o e est� na transa��o FB03, adiciona ao array de linhas compensadas FB03
                    linha_index_array_atual = UBound(array_linhas_compensadas_FB03) + 1
                    ReDim Preserve array_linhas_compensadas_FB03(LBound(array_linhas_compensadas_FB03) To linha_index_array_atual)
                    array_linhas_compensadas_FB03(linha_index_array_atual) = array_linha_atual
                End If
            ElseIf acao = "CHAMADO_CTA_A_PAGAR" Then

                linha_index_array_atual = UBound(array_linhas_chamado_contas_pagar) + 1
                ReDim Preserve array_linhas_chamado_contas_pagar(LBound(array_linhas_chamado_contas_pagar) To linha_index_array_atual)
                array_linhas_chamado_contas_pagar(linha_index_array_atual) = array_linha_atual
            End If
            
        Next iterator2
        
        ' Vai para a pr�xima p�gina
        iterator1 = iterator1 + 1
        session_number.findById("wnd[0]").sendVKey 82  ' Page Down (tecla F20)
    Loop
    
    ' Retorna para a primeira p�gina ap�s processar todas
    session_number.findById("wnd[0]").sendVKey 80  ' F8 (refresh/primeira p�gina)
    
End Sub
Public Sub IncluirAtualizarChamadoPendente()

    ' Se a quantidade de NFD_OC_chamado for "01", sai da subrotina.
    If qtde_NFD_OC_chamado = "01" Then
        Exit Sub
    End If

    ' Define o c�digo de refer�ncia do gerador.
    generatorReferenceCode = "CONTASARECEBERELECT-638618399315"
    ' Obt�m o token da API da planilha "API KEY".
    API_token = ThisWorkbook.Sheets("API KEY").Range("A1").Value

    ' Cria um objeto XMLHTTP para fazer requisi��es HTTP.
    Set http = CreateObject("MSXML2.XMLHTTP")
    ' Define o link da API para obter a lista de chamados.
    link = "https://electrolux.ellevo.com/api/v1/ticket/ticket-list/" & chamado
    ' Envia uma requisi��o GET para a API e obt�m a resposta.
    response = EnviarRequisicao("GET", http, link, "")
    ' Converte a resposta JSON em um objeto JSON.
    Set json = JsonConverter.ParseJson(response)
    ' Obt�m os procedimentos do chamado a partir do objeto JSON.
    Set proceedings = json("proceedings")
    ' Obt�m o n�mero de procedimentos do chamado.
    numero_tramites_chamado = proceedings.Count
    ' Verifica se o dicion�rio dictChamadosPendentes est� vazio.
    If IsEmpty(dictChamadosPendentes) Then
        ' Se estiver vazio, cria um novo dicion�rio.
        Set dictChamadosPendentes = CreateObject("Scripting.Dictionary")
        ' Adiciona o chamado e o n�mero de tr�mites ao dicion�rio.
        dictChamadosPendentes.Add chamado, numero_tramites_chamado
    Else
        ' Se n�o estiver vazio, atualiza o n�mero de tr�mites do chamado no dicion�rio.
        dictChamadosPendentes(chamado) = numero_tramites_chamado
    End If


    ' Verifica se o chamado j� existe na coluna A da aba "aba_historico_chamados_pendentes".
    If Application.WorksheetFunction.CountIf(aba_historico_chamados_pendentes.Columns("A:A"), chamado) > 0 Then
        ' Obt�m o n�mero da �ltima linha preenchida na coluna A.
        linha_fim_aba_historico_chamados_pendentes = aba_historico_chamados_pendentes.Range("A1048576").End(xlUp).Row
        ' Loop atrav�s das linhas da aba.
        For linha = 2 To linha_fim_aba_historico_chamados_pendentes
            ' Se o valor da coluna A corresponder ao chamado.
            If CStr(aba_historico_chamados_pendentes.Range("A" & linha).Value) = chamado Then
                ' Atualiza o n�mero de tr�mites na coluna B.
                aba_historico_chamados_pendentes.Range("B" & linha).Value = numero_tramites_chamado
                ' Sai do loop.
                Exit For
            End If
        Next linha
    Else
        ' Se o chamado n�o existir na aba, obt�m o n�mero da pr�xima linha vazia.
        linha_fim_aba_historico_chamados_pendentes = aba_historico_chamados_pendentes.Range("A1048576").End(xlUp).Offset(1, 0).Row
        ' Insere o chamado na coluna A.
        aba_historico_chamados_pendentes.Range("A" & linha_fim_aba_historico_chamados_pendentes).Value = chamado
        ' Insere o n�mero de tr�mites na coluna B.
        aba_historico_chamados_pendentes.Range("B" & linha_fim_aba_historico_chamados_pendentes).Value = numero_tramites_chamado
    End If

End Sub
Public Sub EtapaFB03(ByVal session_number As Object, ByVal doc_compensacao As String, tipo_acao_array As String)

    Dim primeiro_num_doc_item As String

    ' Entra na transa��o FB03 (Exibir Documento Cont�bil) na sess�o especificada
    session_number.findById("wnd[0]/tbar[0]/okcd").text = "/N FB03"
    ' Simula a tecla Enter
    session_number.findById("wnd[0]").sendVKey 0
    ' Preenche o campo de n�mero do documento com o valor passado
    session_number.findById("wnd[0]/usr/txtRF05L-BELNR").text = doc_compensacao
    ' Preenche o campo de c�digo da empresa com "BR10"
    session_number.findById("wnd[0]/usr/ctxtRF05L-BUKRS").text = "BR10"
    ' Inicializa um contador para diminuir o ano na busca
    i3 = 0
diminuir_ano:
    ' Preenche o campo de ano com o ano atual menos o contador
    session_number.findById("wnd[0]/usr/txtRF05L-GJAHR").text = VBA.Year(Date) - i3
    ' Simula a tecla Enter
    session_number.findById("wnd[0]").sendVKey 0
    ' Se a barra de status n�o estiver vazia (indicando erro ao encontrar o documento no ano atual)
    If session_number.findById("wnd[0]/sbar").text <> "" Then
        ' Incrementa o contador do ano
        i3 = i3 + 1
        ' Volta para a linha 'diminuir_ano' para tentar o ano anterior
        GoTo diminuir_ano
    End If
    ' Seleciona a op��o de menu "Ambiente -> Detalhes do cabe�alho"
    session_number.findById("wnd[0]/mbar/menu[5]/menu[3]").Select
    
    If InStr(1, session_number.findById("wnd[0]/sbar").text, "N�o foi encontrado nenhum item compensado p/o documento") Then
        Exit Sub
    End If
    ' Define o foco no label da coluna de texto na segunda linha
    session_number.findById("wnd[0]/usr/lbl[" & x_texto & ",2]").SetFocus
    ' Simula a tecla Shift+F2 (provavelmente para acessar detalhes)
    session_number.findById("wnd[0]").sendVKey 2
    ' Clica no bot�o de ajuda de pesquisa para o campo de texto
    session_number.findById("wnd[0]/tbar[1]/btn[38]").press
    ' Clica no bot�o de m�ltipla sele��o para o campo de texto
    session_number.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
    ' Seleciona a aba "Valores N�o Individuais"
    session_number.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").Select
    ' Preenche os campos de sele��o com "*Documento cont�bil*" e "*Lan�to.de pagamento*" para buscar esses textos
    session_number.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "*Documento cont�bil*"
    session_number.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").text = "*Lan�to.de pagamento*"
    ' Clica no bot�o de executar na janela de m�ltipla sele��o
    session_number.findById("wnd[2]/tbar[0]/btn[8]").press
    ' Clica no bot�o de copiar na janela de pesquisa
    session_number.findById("wnd[1]/tbar[0]/btn[0]").press
    ' Define o foco no label da coluna de tipo de documento na segunda linha
    session_number.findById("wnd[0]/usr/lbl[" & x_tipo_doc & ",2]").SetFocus
    ' Simula a tecla Shift+F2 (provavelmente para acessar detalhes)
    session_number.findById("wnd[0]").sendVKey 2
    ' Clica no bot�o para exibir mais detalhes do documento
    session_number.findById("wnd[0]/tbar[1]/btn[41]").press

    ' Chama a sub-rotina para preencher um array com informa��es da linha atual, passando a sess�o, �ndices e o tipo de a��o
    Call PreencherArrayLinhasCondicaoAtual(session_number, i4, i5, tipo_acao_array)

End Sub


Public Sub AlterarAtribuicao(ByVal session_number As Object, texto As String)


    ' Clica no bot�o para alterar o texto de atribui��o
    session_number.findById("wnd[0]/tbar[1]/btn[45]").press
    ' Preenche o campo de texto de atribui��o com o valor passado
    session_number.findById("wnd[1]/usr/txt*BSEG-ZUONR").text = texto
    ' Simula a tecla Enter
    session_number.findById("wnd[0]").sendVKey 0
    ' Trata poss�vel erro se a janela n�o existir
    On Error Resume Next
    ' Clica no bot�o de copiar (se a janela de atribui��o ainda estiver aberta)
    session_number.findById("wnd[1]/tbar[0]/btn[0]").press
    ' Desativa o tratamento de erros
    On Error GoTo 0

End Sub
