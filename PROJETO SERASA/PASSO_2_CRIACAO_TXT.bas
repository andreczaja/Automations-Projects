Attribute VB_Name = "PASSO_2_CRIACAO_TXT"
Option Explicit

Public num_titulos_inclusao, num_titulos_exclusao, remessa_inicial, remessa_final As Integer
Public linha, i, i_fim As Integer
Public aba_correspondente As Worksheet
Private tabela_aba_correspondente As ListObject
Public array_nfs_buscadas_port_dev_com_ocorrencia(), array_nfs_buscadas_port_dev_sem_ocorrencia() As Variant


Sub criacao_txt(ByVal tipo_processo As String)


Dim driver As New EdgeDriver
Dim fso As Object
Dim pasta_destino_txt_criado As Object
Dim check_base_preenchida, check_sem_cnpj, check_bloqueio_adv, check_portal_devolucoes, check_plan_distribuicao, check_linha_duplicada, browser_open As Boolean
Dim bloqueio_adv, payer, referencia As String


''''''''''''''' STRINGS '''''''''''''''''

Dim linha_1 As String
    ' variaveis proprias da linha 1
    Dim cnpj_informante, data_movimento, num_ddd_informante, num_fone_informante, num_ramal_informante, nome_contato_informante, id_arquivo, num_remessa, cod_envio_arquivo, diferencial_remessa, _
        logon_contab_cartas, cod_erros_1 As String

Dim linha_2 As String
    ' variaveis proprias da linha 2
    Dim cod_operacao, filial_digito_cnpj, data_ocorrencia, data_termino_contrato, cod_nat_operacao, cod_praca_embratel, tipo_pessoa_principal, tipo_doc_principal, doc_principal, _
        motivo_baixa, tipo_segundo_doc_principal, segundo_doc_principal, uf_doc_principal, tipo_pessoa_coobrigado, tipo_doc_coobrigado, doc_coobrigado, tipo_segundo_doc_coobrigado, doc_segundo_coobrigado, _
        uf_doc_coobrigado, nome_devedor, data_nascimento_devedor, nome_pai, nome_mae, endereco_completo, bairro, municipio, sigla_uf, cod_postal, valor_divida, numero_contrato, numero_serasa, _
        complemento_endereco_devedor, data_compromisso_devedor, valor_total_compromisso, indicativo_envio_reg_dados, indicativo_tipo_comunicado_devedor, cod_erros_2 As String
        
Dim linha_3 As String
    ' variaveis proprias da linha 3
    Dim email_devedor, data_optin_email_devedor, data_optin_fone_devedor As String
    
Dim linha_4 As String

' variaveis em comum linhas
Dim cod_registro, num_ddd_devedor, num_fone_devedor, num_linha, seq_reg_arquivo_string As String

Dim seq_reg_arquivo_integer As Integer
Dim string_total As String

''''''''''''''' STRINGS '''''''''''''''''

    
    If tipo_processo = "I" Then
        Set aba_correspondente = ThisWorkbook.Sheets("FBL5H - Base Geral")
        Set tabela_aba_correspondente = aba_correspondente.ListObjects("Tabela_FBL5H_Base_Geral")
    ElseIf tipo_processo = "E" Then
        Set aba_correspondente = ThisWorkbook.Sheets("FBL5H - Base Compensados SERASA")
        Set tabela_aba_correspondente = aba_correspondente.ListObjects("Tabela_FBL5H_Base_Compensados_SERASA")
    End If
    
    Set aba_base_historica = ThisWorkbook.Sheets("Base Histórica")
    Set tabela_aba_base_historica = aba_base_historica.ListObjects("Tabela_Base_Histórica")
  
    Set aba_numero_remessa = ThisWorkbook.Sheets("Nº Remessa")
    
    i_fim = aba_correspondente.Range("A1048576").End(xlUp).Row
    
    If tipo_processo = "E" And aba_correspondente.Range("A2").Value = "" Then
        Exit Sub
    End If
    
    aba_correspondente.Range("AD2:AE1048576").ClearContents
    
    array_nfs_buscadas_port_dev_com_ocorrencia = Array()
    array_nfs_buscadas_port_dev_sem_ocorrencia = Array()
    browser_open = False
    
    seq_reg_arquivo_integer = 1
    string_total = ""
    
    '''''''''''''''''''''''''' ETAPA COMPOSIÇÃO LINHA 1 '''''''''''''''''''''''''''''''''
    cod_registro = "0"
    cnpj_informante = "076487032"
    data_movimento = CStr(VBA.Format(Date, "YYYYMMDD"))
    num_ddd_informante = "0041"
    num_fone_informante = "99999999"
    num_ramal_informante = "0000"
    nome_contato_informante = "ELECTROLUX DO BRASIL S/A                                              "
    id_arquivo = "SERASA-CONVEM04"
    num_remessa = Preencher(aba_numero_remessa.Range("A1").Value, 6, "ZEROS", "ESQUERDA")
    cod_envio_arquivo = "E"
    diferencial_remessa = "0000"
    ' 3 CARACTERES EM BRANCO
    logon_contab_cartas = "        "
    ' 392 CARACTERES EM BRANCO
    cod_erros_1 = "                                                            "
    seq_reg_arquivo_string = Preencher(seq_reg_arquivo_integer, 7, "ZEROS", "ESQUERDA")
    seq_reg_arquivo_integer = seq_reg_arquivo_integer + 1
    
    linha_1 = cod_registro & cnpj_informante & data_movimento & num_ddd_informante & num_fone_informante & num_ramal_informante & nome_contato_informante & id_arquivo & num_remessa & cod_envio_arquivo & diferencial_remessa & "   " & logon_contab_cartas & String(392, " ") & cod_erros_1 & seq_reg_arquivo_string


    For linha = 2 To i_fim
        
        payer = aba_correspondente.Range("B" & linha).Value
        referencia = aba_correspondente.Range("E" & linha).Value
        bloqueio_adv = aba_correspondente.Range("H" & linha).Value
        '''' CHECKS IMPEDITIVOS DE INCLUSÃO/EXCLUSÃO '''''
        If aba_correspondente.Range("A" & linha).Value = "" Then
            aba_correspondente.Range("AD" & linha).Value = "Linha Vazia"
            GoTo proxima_linha
        ElseIf aba_correspondente.Range("V" & linha).Value = "" Then
            aba_correspondente.Range("AD" & linha).Value = "Payer sem CNPJ preenchido - não é possível enviar ou tirar dívida ao SERASA"
            GoTo proxima_linha
        End If

        ' se for linha nova, não executa, pq não foi incluída em nenhum momento
        If Not VerificarLinhaDuplicada(aba_base_historica, aba_correspondente, linha, tipo_processo) Then
            GoTo proxima_linha
        End If

        If aba_correspondente.Range("AC" & linha).Value = "" Then
            If VerificarBloqueioAdvertencia(bloqueio_adv, tipo_processo) Then
                check_bloqueio_adv = True
            Else
                check_bloqueio_adv = False
            End If

            If VerificarPlanilhaDistribuicao(payer, aba_correspondente, tipo_processo) Then
                check_plan_distribuicao = True
            Else
                check_plan_distribuicao = False
            End If

            If VerificarPortalDevolucoes(driver, aba_correspondente, payer, referencia, linha, tipo_processo, browser_open) Then
                check_portal_devolucoes = True
            Else
                If VerificarPortalDevolucoes(driver, aba_correspondente, payer, referencia, linha, tipo_processo, browser_open) Then
                    check_portal_devolucoes = True
                Else
                    check_portal_devolucoes = False
                End If
            End If


            If tipo_processo = "I" Then
                If Not check_bloqueio_adv Or Not check_plan_distribuicao Or Not check_portal_devolucoes Then
                    If aba_correspondente.Range("AD" & linha).Value = "" Then
                        aba_correspondente.Range("AD" & linha).Value = "Favor processar linha novamente"
                    End If
                    GoTo proxima_linha
                End If
            ElseIf tipo_processo = "E" Then
                If Not check_bloqueio_adv And Not check_plan_distribuicao And Not check_portal_devolucoes Then
                    If aba_correspondente.Range("AD" & linha).Value = "" Then
                        aba_correspondente.Range("AD" & linha).Value = "Favor processar linha novamente"
                    End If
                    GoTo proxima_linha
                End If
            End If
        End If
        
        
        '''''''''''''''''''''''''' ETAPA COMPOSIÇÃO LINHA 2 '''''''''''''''''''''''''''''''''
        cod_registro = "1"
        cod_operacao = tipo_processo
        filial_digito_cnpj = "000125"
        
        data_ocorrencia = CStr(VBA.Format(aba_correspondente.Range("K" & linha).Value, "YYYYMMDD"))
    
        
        data_termino_contrato = data_ocorrencia
        cod_nat_operacao = " DP"
        cod_praca_embratel = "    "
        
        ' VERIFICANDO SE O VALOR DA COLUNA AG É CNPJ OU CPF
        If Len(aba_correspondente.Range("V" & linha).Value) = 14 Then
            tipo_pessoa_principal = "J"
            tipo_doc_principal = "1"
        ElseIf Len(aba_correspondente.Range("V" & linha).Value) = 11 Then
            tipo_pessoa_principal = "F"
            tipo_doc_principal = "2"
        Else
            tipo_pessoa_principal = "J"
            tipo_doc_principal = "1"
        End If
        doc_principal = Preencher(aba_correspondente.Range("V" & linha).Value, 15, "ZEROS", "ESQUERDA")
        If cod_operacao = "I" Then
            motivo_baixa = "  "
        ElseIf cod_operacao = "E" Then
            motivo_baixa = "01"
        End If
        
        tipo_segundo_doc_principal = " "
        segundo_doc_principal = "               "
        uf_doc_principal = "  "
        tipo_pessoa_coobrigado = " "
        tipo_doc_coobrigado = " "
        doc_coobrigado = "               "
        ' 2 espaços vazios
        tipo_segundo_doc_coobrigado = " "
        doc_segundo_coobrigado = "               "
        uf_doc_coobrigado = "  "
        nome_devedor = Preencher(aba_correspondente.Range("C" & linha).Value, 70, "ESPAÇOS", "DIREITA")
        data_nascimento_devedor = "00000000"
        nome_pai = "                                                                      "
        nome_mae = "                                                                      "
        endereco_completo = Preencher(aba_correspondente.Range("O" & linha).Value, 45, "ESPAÇOS", "DIREITA")
        If Len(aba_correspondente.Range("U" & linha).Value) > 20 Then
            bairro = aba_correspondente.Range("U" & linha).Value
            bairro = Reduzir(bairro, 20)
        Else
            bairro = Preencher(aba_correspondente.Range("U" & linha).Value, 20, "ESPAÇOS", "DIREITA")
        End If
        If Len(aba_correspondente.Range("Q" & linha).Value) > 25 Then
            municipio = aba_correspondente.Range("Q" & linha).Value
            municipio = Reduzir(bairro, 25)
        Else
            municipio = Preencher(aba_correspondente.Range("Q" & linha).Value, 25, "ESPAÇOS", "DIREITA")
        End If
        sigla_uf = Preencher(aba_correspondente.Range("S" & linha).Value, 2, "ESPAÇOS", "ESQUERDA")
        cod_postal = Preencher(aba_correspondente.Range("R" & linha).Value, 8, "ZEROS", "ESQUERDA")
        valor_divida = Preencher(Replace(Replace(aba_correspondente.Range("L" & linha).Value, ",", ""), ".", ""), 15, "ZEROS", "ESQUERDA")
        numero_contrato = Replace((aba_correspondente.Range("E" & linha).Value & CStr(aba_correspondente.Range("F" & linha).Value)), "-", "")
        
        If Len(numero_contrato) > 16 Then
            numero_contrato = Reduzir(numero_contrato, 16)
        ElseIf Len(numero_contrato) < 16 Then
            numero_contrato = Preencher(numero_contrato, 16, "ZEROS", "ESQUERDA")
        End If
        
        numero_serasa = "         "
        complemento_endereco_devedor = "                         "
        num_ddd_devedor = "00" & Left(aba_correspondente.Range("P" & linha).Value, 2)
        num_ddd_devedor = Preencher(num_ddd_devedor, 4, "ZEROS", "ESQUERDA")
        ' FAZER VERIFICAÇÃO DE COMPRIMENTO
        num_fone_devedor = aba_correspondente.Range("P" & linha).Value
        
        If Len(num_fone_devedor) > 9 Then
            num_fone_devedor = Reduzir(num_fone_devedor, 9)
        ElseIf Len(num_fone_devedor) < 9 Then
            num_fone_devedor = Preencher(num_fone_devedor, 9, "ZEROS", "ESQUERDA")
        End If
    
        data_compromisso_devedor = "        "
        valor_total_compromisso = "               "
        indicativo_envio_reg_dados = "S"
        ' 5 espaços vazios
        indicativo_tipo_comunicado_devedor = " "
        ' 2 espaços vazios
        cod_erros_2 = "                                                            "
        seq_reg_arquivo_string = Preencher(seq_reg_arquivo_integer, 7, "ZEROS", "ESQUERDA")
        seq_reg_arquivo_integer = seq_reg_arquivo_integer + 1
        
        linha_2 = cod_registro & cod_operacao & filial_digito_cnpj & data_ocorrencia & data_termino_contrato & cod_nat_operacao & cod_praca_embratel & tipo_pessoa_principal & tipo_doc_principal & doc_principal & _
            motivo_baixa & tipo_segundo_doc_principal & segundo_doc_principal & uf_doc_principal & tipo_pessoa_coobrigado & tipo_doc_coobrigado & doc_coobrigado & "  " & tipo_segundo_doc_coobrigado & doc_segundo_coobrigado & _
            uf_doc_coobrigado & nome_devedor & data_nascimento_devedor & nome_pai & nome_mae & endereco_completo & bairro & municipio & sigla_uf & cod_postal & valor_divida & numero_contrato & numero_serasa & complemento_endereco_devedor & num_ddd_devedor & _
            num_fone_devedor & data_compromisso_devedor & valor_total_compromisso & indicativo_envio_reg_dados & "     " & indicativo_tipo_comunicado_devedor & "  " & cod_erros_2 & seq_reg_arquivo_string
        Debug.Print Len(linha_2)
        
        '''''''''''''''''''''''''' ETAPA COMPOSIÇÃO LINHA 3 '''''''''''''''''''''''''''''''''
        cod_registro = "5"
        If aba_correspondente.Range("W" & linha).Value <> "" Then
            email_devedor = Preencher(aba_correspondente.Range("W" & linha).Value, 100, "ESPAÇOS", "DIREITA")
        ElseIf aba_correspondente.Range("X" & linha).Value <> "" Then
            email_devedor = Preencher(aba_correspondente.Range("X" & linha).Value, 100, "ESPAÇOS", "DIREITA")
        ElseIf aba_correspondente.Range("Y" & linha).Value <> "" Then
            email_devedor = Preencher(aba_correspondente.Range("Y" & linha).Value, 100, "ESPAÇOS", "DIREITA")
        ElseIf aba_correspondente.Range("Z" & linha).Value <> "" Then
            email_devedor = Preencher(aba_correspondente.Range("Z" & linha).Value, 100, "ESPAÇOS", "DIREITA")
        Else
            email_devedor = String(100, " ")
        End If
        data_optin_email_devedor = "        "
        num_ddd_devedor = num_ddd_devedor
        ' FAZER VERIFICAÇÃO DE COMPRIMENTO
        num_fone_devedor = num_fone_devedor
        data_optin_fone_devedor = "        "
        ' 463 espaços vazios
        seq_reg_arquivo_string = Preencher(seq_reg_arquivo_integer, 7, "ZEROS", "ESQUERDA")
        seq_reg_arquivo_integer = seq_reg_arquivo_integer + 1
        linha_3 = cod_registro & email_devedor & data_optin_email_devedor & num_ddd_devedor & num_fone_devedor & data_optin_fone_devedor & String(463, " ") & seq_reg_arquivo_string
        
        If string_total = "" Then
            string_total = linha_2 & vbNewLine & linha_3
        Else
            string_total = string_total & vbNewLine & linha_2 & vbNewLine & linha_3
        End If
        
        If tipo_processo = "I" Then
            num_titulos_inclusao = num_titulos_inclusao + 1
        ElseIf tipo_processo = "E" Then
            num_titulos_exclusao = num_titulos_exclusao + 1
        End If
        num_remessa = aba_numero_remessa.Range("A1").Value
        aba_correspondente.Range("AE" & linha).Value = num_remessa
        Call ExcluirIncluirLinhaBaseHistorica(tipo_processo, linha, aba_correspondente, aba_base_historica)
        
proxima_linha:
    Next linha
    
    If string_total = "" Then
        If tipo_processo = "I" Then
            MsgBox "Nenhum remessa de inclusão gerada"
        ElseIf tipo_processo = "E" Then
            MsgBox "Nenhum remessa de exclusão gerada"
        End If
        Exit Sub
    End If
    
    '''''''''''''''''''''''''' ETAPA COMPOSIÇÃO LINHA 4 '''''''''''''''''''''''''''''''''
    cod_registro = "9"
    ' 2 espaços vazios
    seq_reg_arquivo_string = Preencher(seq_reg_arquivo_integer, 7, "ZEROS", "ESQUERDA")
    linha_4 = cod_registro & String(592, " ") & seq_reg_arquivo_string
    
    
    aba_numero_remessa.Range("A1").Value = aba_numero_remessa.Range("A1").Value + 1
    string_total = linha_1 & vbNewLine & string_total & vbNewLine & linha_4
    
    '''''''''''''''''''''''''' ETAPA CRIAÇÃO TXT '''''''''''''''''''''''''''
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set pasta_destino_txt_criado = fso.getfolder(Replace(caminho_pasta, "Arquivo TXT SERASA SAP", "Remessas Geradas"))
    If pasta_destino_txt_criado Is Nothing Then
         MsgBox "Por favor, escolha a pasta do seu computador para onde será descarregado o TXT do Serasa.", vbOKCancel, "Aviso"

        With Application.FileDialog(msoFileDialogFolderPicker)
            If .Show = -1 Then ' O usuário selecionou uma pasta
                caminho_pasta = .SelectedItems(1) & "\"
                Set pasta_destino_txt_criado = fso.getfolder(caminho_pasta)
            Else
                ' O usuário cancelou a seleção da pasta
                MsgBox "Nenhuma pasta selecionada. O processo foi cancelado."
                Exit Sub
            End If
        End With
    End If
    On Error GoTo 0

    Dim fileNum As Integer
    Dim Emplacement, nome_arquivo As String
    fileNum = FreeFile

    If tipo_processo = "I" Then
        nome_arquivo = "Remessas Serasa Inclusao " & VBA.Format(VBA.Date, "dd.mm.yyyy") & ".txt"
    ElseIf tipo_processo = "E" Then
        nome_arquivo = "Remessas Serasa Exclusao " & VBA.Format(VBA.Date, "dd.mm.yyyy") & ".txt"
    End If
    If Dir(pasta_destino_txt_criado & "\" & nome_arquivo) <> "" Then
        Kill pasta_destino_txt_criado & "\" & nome_arquivo
    End If

    Emplacement = pasta_destino_txt_criado & "\" & nome_arquivo


    Open Emplacement For Output As fileNum
    Print #fileNum, string_total
    Close fileNum


End Sub


Private Function Preencher(ByVal variavel As String, num_caracteres As Integer, tipo As String, sentido As String) As String
    If Len(variavel) < num_caracteres And tipo = "ZEROS" Then
        Do Until Len(variavel) = num_caracteres
            If sentido = "ESQUERDA" Then
                variavel = "0" & variavel
            ElseIf sentido = "DIREITA" Then
                variavel = variavel & "0"
            End If
        Loop
    ElseIf Len(variavel) < num_caracteres And tipo = "ESPAÇOS" Then
        Do Until Len(variavel) = num_caracteres
            If sentido = "ESQUERDA" Then
                variavel = " " & variavel
            ElseIf sentido = "DIREITA" Then
                variavel = variavel & " "
            End If
        Loop
    End If
    Preencher = variavel
End Function

Private Function Reduzir(ByVal variavel As String, num_caracteres As Integer) As String

Dim i As Integer
    i = 1
    Do Until Len(variavel) = num_caracteres
        variavel = Mid(variavel, 1, Len(variavel) - i)
    Loop
    Reduzir = variavel
End Function


Private Function VerificarPortalDevolucoes(ByRef driver As Object, ByVal aba_correspondente As Worksheet, ByVal payer As String, ByVal referencia As String, ByVal linha_aba_correspondente As Integer, ByVal tipo_processo As String, ByRef browser_open As Boolean) As Boolean
''''''''''''' ETAPA SELENIUM ''''''''''''''
    Dim data_inicial, elemento_home_barra_lateral, elemento_pesquisar, elemento_login, elemento_data_inicial, elemento_data_final, _
         elemento_botao_filtro, elemento_codigo_cliente_sap, elemento_nenhuma_ocorrencia_encontrada, elemento_buscar, elemento_nf_electrolux, elemento_ocorrencia_encontrada, elemento_tabela_ocorrencias_baixadas As String
    Dim by As New by
    Dim contador As Integer
    
    elemento_login = "/html/body/app-root/div[1]/div/div/div/app-login/div/div/div/div/div/div/div[2]/div/div/div[1]/button"
    elemento_home_barra_lateral = "/html/body/app-root/div/div/div[2]/div[1]/app-nav-bar/nav/div/div[1]/button[2]"
    elemento_pesquisar = "/html/body/app-root/div/div/div[1]/app-side-bar-desktop/nav/app-side-bar-menu/ul/div[5]/li/a"
    elemento_data_inicial = "/html/body/app-root/div/div/div[2]/div[2]/app-search-occurrence/app-listing-occurences/div[3]/div[1]/div/div/div/div[1]/div[3]/div/div/input[1]"
    elemento_data_final = "/html/body/app-root/div/div/div[2]/div[2]/app-search-occurrence/app-listing-occurences/div[3]/div[1]/div/div/div/div[1]/div[3]/div/div/input[2]"
    elemento_nf_electrolux = "/html/body/app-root/div/div/div[2]/div[2]/app-search-occurrence/app-listing-occurences/div[3]/div[1]/div/div/div/div[2]/div[1]/div/input"
    elemento_botao_filtro = "/html/body/app-root/div/div/div[2]/div[2]/app-search-occurrence/app-listing-occurences/div[3]/div[1]/div/div/div/div[7]/div/button[1]"
    elemento_codigo_cliente_sap = "/html/body/app-root/div/div/div[2]/div[2]/app-search-occurrence/app-listing-occurences/div[3]/div[1]/div/div/div/div[6]/div[2]/div/input"
    elemento_buscar = "/html/body/app-root/div/div/div[2]/div[2]/app-search-occurrence/app-listing-occurences/div[3]/div[1]/div/div/div/div[7]/div/button[4]"
    elemento_nenhuma_ocorrencia_encontrada = "/html/body/app-root/div/div/div[2]/div[2]/app-search-occurrence/app-listing-occurences/div[3]/app-alert"
    elemento_ocorrencia_encontrada = "/html/body/app-root/div/div/div[2]/div[2]/app-search-occurrence/app-listing-occurences/div[3]/div[2]/div/div/div/div[2]/div/div/table/thead/tr/th[1]"
    elemento_tabela_ocorrencias_baixadas = "/html/body/app-root/div/div/div[2]/div[2]/app-search-occurrence/app-listing-occurences/div[3]/div[2]/div/div/div[2]"
    
    referencia = TratativasReferencia(referencia)
    
    If UBound(VBA.Filter(array_nfs_buscadas_port_dev_com_ocorrencia, referencia)) = 0 And tipo_processo = "I" Then
        VerificarPortalDevolucoes = False
        Exit Function
    ElseIf UBound(VBA.Filter(array_nfs_buscadas_port_dev_com_ocorrencia, referencia)) = 0 And tipo_processo = "E" Then
        VerificarPortalDevolucoes = True
        Exit Function
    ElseIf UBound(VBA.Filter(array_nfs_buscadas_port_dev_sem_ocorrencia, referencia)) = 0 And tipo_processo = "I" Then
        VerificarPortalDevolucoes = True
        Exit Function
    ElseIf UBound(VBA.Filter(array_nfs_buscadas_port_dev_sem_ocorrencia, referencia)) = 0 And tipo_processo = "E" Then
        VerificarPortalDevolucoes = False
        Exit Function
    End If
    
    If linha_aba_correspondente = 2 Or Not browser_open Then
        browser_open = True
        driver.Get "https://portaldevolucoes.electrolux.com.br/login"
        driver.Window.Maximize
        
        
        driver.FindElementByXPath(elemento_login).Click
        Do Until driver.IsElementPresent(by.XPath(elemento_home_barra_lateral))
            Application.Wait (Now + TimeValue("00:00:01"))
        Loop
        driver.ExecuteScript "document.body.style.zoom='80%'"
        driver.Get "https://portaldevolucoes.electrolux.com.br/search_occurrence/default"
        Call PreencherDatasPortDev(driver, elemento_data_inicial, elemento_data_final)
    End If
    
refresh_page:
    
    If condicao_ocorrencia_encontrada Then
        condicao_ocorrencia_encontrada = False
        driver.Get "https://portaldevolucoes.electrolux.com.br/search_occurrence/default"
        Application.Wait (Now + TimeValue("00:00:01"))
        Call PreencherDatasPortDev(driver, elemento_data_inicial, elemento_data_final)
    End If
    contador = 1
    Do Until driver.IsElementPresent(by.XPath(elemento_data_inicial))
        If contador < 7 Then
            Application.Wait (Now + TimeValue("00:00:01"))
            contador = contador + 1
        Else
            driver.Get "https://portaldevolucoes.electrolux.com.br/search_occurrence/default"
            Application.Wait (Now + TimeValue("00:00:01"))
            GoTo refresh_page
        End If
    Loop

    driver.ExecuteScript "document.body.style.zoom='80%'"
    
    driver.FindElementByXPath(elemento_nf_electrolux).Click
    driver.FindElementByXPath(elemento_nf_electrolux).Clear
    driver.FindElementByXPath(elemento_nf_electrolux).SendKeys referencia
    If Not driver.FindElementByXPath(elemento_codigo_cliente_sap).IsDisplayed Or Not driver.FindElementByXPath(elemento_codigo_cliente_sap).IsEnabled Then
clicar_botao_filtro_novamente:
        On Error GoTo clicar_botao_filtro_novamente
        Application.Wait (Now + TimeValue("00:00:01"))
        driver.FindElementByXPath(elemento_botao_filtro).Click
        Application.Wait (Now + TimeValue("00:00:01"))
        On Error GoTo 0
    End If
    
    

    driver.FindElementByXPath(elemento_codigo_cliente_sap).Click
    driver.FindElementByXPath(elemento_codigo_cliente_sap).Clear
    driver.FindElementByXPath(elemento_codigo_cliente_sap).SendKeys payer
    
verificar_novamente_ocorrencia:
    contador = 1
    Do Until driver.IsElementPresent(by.XPath(elemento_buscar))
        If contador > 5 Then
            Exit Do
        Else
            contador = contador + 1
        End If
        Application.Wait (Now + TimeValue("00:00:01"))
    Loop

    If driver.IsElementPresent(by.XPath(elemento_buscar)) Then
        driver.FindElementByXPath(elemento_buscar).Click
        Application.Wait (Now + TimeValue("00:00:01"))
    End If
    
    If driver.IsElementPresent(by.XPath(elemento_ocorrencia_encontrada)) Then
        condicao_ocorrencia_encontrada = True
        If tipo_processo = "I" Then
            aba_correspondente.Range("AD" & linha_aba_correspondente).Value = "Cliente com Ocorrência em aberto referente a essa NF"
            VerificarPortalDevolucoes = False
        ElseIf tipo_processo = "E" Then
            VerificarPortalDevolucoes = True
        End If
        
        If UBound(VBA.Filter(array_nfs_buscadas_port_dev_com_ocorrencia, referencia)) < 0 Then
            ReDim Preserve array_nfs_buscadas_port_dev_com_ocorrencia(LBound(array_nfs_buscadas_port_dev_com_ocorrencia) To UBound(array_nfs_buscadas_port_dev_com_ocorrencia) + 1)
            array_nfs_buscadas_port_dev_com_ocorrencia(UBound(array_nfs_buscadas_port_dev_com_ocorrencia)) = referencia
        End If
        Exit Function
        
    End If
    If driver.IsElementPresent(by.XPath(elemento_data_inicial)) And driver.FindElementByXPath(elemento_nenhuma_ocorrencia_encontrada).Text = "" Then
        GoTo verificar_novamente_ocorrencia
    End If
    If driver.FindElementByXPath(elemento_nenhuma_ocorrencia_encontrada).Text = "Não houve resultado para os filtros selecionados!" Then
        condicao_ocorrencia_encontrada = False
        If tipo_processo = "E" Then
            VerificarPortalDevolucoes = False
        ElseIf tipo_processo = "I" Then
            VerificarPortalDevolucoes = True
        End If
    End If
    If UBound(VBA.Filter(array_nfs_buscadas_port_dev_sem_ocorrencia, referencia)) < 0 Then
        ReDim Preserve array_nfs_buscadas_port_dev_sem_ocorrencia(LBound(array_nfs_buscadas_port_dev_sem_ocorrencia) To UBound(array_nfs_buscadas_port_dev_sem_ocorrencia) + 1)
        array_nfs_buscadas_port_dev_sem_ocorrencia(UBound(array_nfs_buscadas_port_dev_sem_ocorrencia)) = referencia
    End If
  
    
End Function
Private Function PreencherDatasPortDev(ByRef driver As Object, ByVal elemento_data_inicial As String, ByVal elemento_data_final As String)
    driver.FindElementByXPath(elemento_data_inicial).Click
    driver.FindElementByXPath(elemento_data_inicial).Clear
    driver.FindElementByXPath(elemento_data_inicial).SendKeys Format(Date - 500, "dd/mm/yyyy")
    driver.FindElementByXPath(elemento_data_final).Click
    driver.FindElementByXPath(elemento_data_final).Clear
    driver.FindElementByXPath(elemento_data_final).SendKeys Format(Date, "dd/mm/yyyy")
End Function


