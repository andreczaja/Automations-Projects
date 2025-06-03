Attribute VB_Name = "Y_CRIACAO_ARQUIVOS"
Public Sub AlimentarArquivoCliente()

    Dim aba_1 As Worksheet, aba_2 As Worksheet, aba_3 As Worksheet, aba_4 As Worksheet
    Dim tbl As ListObject, tbl_2 As ListObject, tbl_3 As ListObject
    Dim indexlinha As Integer, indexcoluna As Integer, quantidade_abas_arquivo_cliente As Integer
    Dim ws As Variant

    If qtde_NFD_OC_chamado = "01" Then
        ' Define o workbook ativo como o arquivo do cliente
        Set arquivo_cliente = Workbooks.Add
        ' Define o nome da primeira planilha
        arquivo_cliente.Sheets(1).Name = "Cr�d disp a abater.reembolsar"
        arquivo_cliente.Sheets.Add After:=arquivo_cliente.Sheets(1)
        Set aba_1 = arquivo_cliente.Sheets(1)
        ' Define o nome da primeira planilha
        arquivo_cliente.Sheets(2).Name = "Cr�ditos Ja Utilizados"
        Set aba_2 = arquivo_cliente.Sheets(2)
        
    ElseIf qtde_NFD_OC_chamado = "Acima de 01" Then
        Set arquivo_cliente = arquivo_anexo_chamado_atual
        arquivo_cliente.Sheets.Add After:=arquivo_cliente.Sheets(1)
        ' Define o nome da primeira planilha
        arquivo_cliente.Sheets(2).Name = "Cr�d disp a abater.reembolsar"
        ' Define a vari�vel para a primeira planilha
        Set aba_1 = arquivo_cliente.Sheets(2)
        ' Adiciona uma nova planilha ap�s a primeira
        arquivo_cliente.Sheets.Add After:=aba_1
        ' Define o nome da segunda planilha
        arquivo_cliente.Sheets(3).Name = "Cr�ditos Ja Utilizados"
        ' Define a vari�vel para a segunda planilha
        Set aba_2 = arquivo_cliente.Sheets(3)
        
    End If
    
    
    
    ' Define o formato de n�mero para as colunas H nas duas primeiras planilhas
    aba_1.Columns("H:H").NumberFormat = "#,###,###.##"
    aba_2.Columns("H:H").NumberFormat = "#,###,###.##"
    ' Se a condi��o do payer for "abatidos"
    If condicao_payer = "abatidos" Then
        ' Adiciona uma nova planilha ap�s a segunda
        arquivo_cliente.Sheets.Add After:=aba_2
        ' Define o nome da terceira planilha
        arquivo_cliente.Sheets(aba_2.Index + 1).Name = "Detalhe Abatimento"
        ' Define a vari�vel para a terceira planilha
        Set aba_3 = arquivo_cliente.Sheets("Detalhe Abatimento")
        ' Define o formato de n�mero para as colunas H e S na terceira planilha
        aba_3.Columns("H:H").NumberFormat = "#,###,###.##"
        aba_3.Columns("S:S").NumberFormat = "#,###,###.##"
    End If
    ' Se houver dados no array de linhas abertas da FBL5N
    If UBound(array_geral_linhas_abertas_FBL5N) > 0 Then
        ' Inicializa os �ndices de linha e coluna
        indexlinha = 1
        indexcoluna = 1
        ' Loop atrav�s das linhas do array
        For i = LBound(array_geral_linhas_abertas_FBL5N) To UBound(array_geral_linhas_abertas_FBL5N)
            ' Reinicializa o �ndice da coluna para cada linha
            indexcoluna = 1
            ' Loop atrav�s dos itens de cada linha do array
            For i2 = LBound(array_geral_linhas_abertas_FBL5N(i)) To UBound(array_geral_linhas_abertas_FBL5N(i))
                ' Preenche a c�lula com o valor do array
                aba_1.Cells(indexlinha, indexcoluna).Value = array_geral_linhas_abertas_FBL5N(i)(i2)
                ' Incrementa o �ndice da coluna
                indexcoluna = indexcoluna + 1
            Next i2
            ' Incrementa o �ndice da linha
            indexlinha = indexlinha + 1
        Next i
        
        ' Define o range dos dados preenchidos na primeira planilha
        Set rng = aba_1.Range("A1:Q" & indexlinha - 1)
        ' Cria uma tabela a partir do range
        Set tbl = aba_1.ListObjects.Add(xlSrcRange, rng, , xlYes)
        ' Ajusta automaticamente a largura das colunas
        aba_1.Columns("A:Q").AutoFit
    Else
        ' Se n�o houver linhas abertas, informa na planilha
        aba_1.Range("A1").Value = "Nenhuma linha a ser abatida de um t�tulo ou reembolsada/devolvida ao cliente."
    End If
    ' Se houver dados no array de linhas compensadas da FB03
    If UBound(array_linhas_compensadas_FB03) > 0 Then
        ' Inicializa os �ndices de linha e coluna
        indexlinha = 1
        indexcoluna = 1
        Dim ultimo_doc_compensacao As String
        ultimo_doc_compensacao = array_linhas_compensadas_FB03(0)(0)
        ' Loop atrav�s das linhas do array
        For i = LBound(array_linhas_compensadas_FB03) To UBound(array_linhas_compensadas_FB03)
            If ultimo_doc_compensacao <> array_linhas_compensadas_FB03(i)(0) Then
                If ultimo_doc_compensacao <> "DocCompens" Then
                    indexlinha = indexlinha + 1
                End If
                ultimo_doc_compensacao = array_linhas_compensadas_FB03(i)(0)
            End If
            
            ' Loop atrav�s dos itens de cada linha do array
            For i2 = LBound(array_linhas_compensadas_FB03(i)) To UBound(array_linhas_compensadas_FB03(i))
                ' Preenche a c�lula com o valor do array
                aba_2.Cells(indexlinha, indexcoluna).Value = array_linhas_compensadas_FB03(i)(i2)
                ' Incrementa o �ndice da coluna
                indexcoluna = indexcoluna + 1
            Next i2
            ' Incrementa o �ndice da linha
            indexlinha = indexlinha + 1
            ' Reinicializa o �ndice da coluna
            indexcoluna = 1
        Next i
        
        ' Se o documento de compensa��o for diferente de "DocCompens" e n�o estiver vazio
        If doc_compensacao <> "DocCompens" And doc_compensacao <> "" Then
            ' Define o range dos dados preenchidos na segunda planilha
            Set rng = aba_2.Range("A1:Q" & indexlinha - 1)
            ' Cria uma tabela a partir do range
            Set tbl_2 = aba_2.ListObjects.Add(xlSrcRange, rng, , xlYes)
            ' Ajusta automaticamente a largura das colunas
            aba_2.Columns("A:Q").AutoFit
        Else
            ' Se n�o houver documentos de compensa��o v�lidos, exclui a planilha
            aba_2.Delete
        End If
    Else
        ' Se n�o houver cr�ditos utilizados, informa na planilha
        aba_2.Range("A1").Value = "Nenhum cr�dito utilizado anteriormente referente a(s) OC(s) informadas."
    End If
    
    ' Se a condi��o do payer for "abatidos"
    If condicao_payer = "abatidos" Then
        
        ' Inicializa os �ndices de linha e coluna
        indexlinha = 1
        indexcoluna = 1
        ' Loop atrav�s das linhas do array de detalhes de abatimento
        For i = LBound(array_linhas_detalhe_abatimento) To UBound(array_linhas_detalhe_abatimento)
            ' Loop atrav�s dos itens de cada linha do array
            For i2 = LBound(array_linhas_detalhe_abatimento(i)) To UBound(array_linhas_detalhe_abatimento(i))
                ' Preenche a c�lula com o valor do array
                aba_3.Cells(indexlinha, indexcoluna).Value = array_linhas_detalhe_abatimento(i)(i2)
                ' Incrementa o �ndice da coluna
                indexcoluna = indexcoluna + 1
            Next i2
            ' Incrementa o �ndice da linha
            indexlinha = indexlinha + 1
            ' Reinicializa o �ndice da coluna
            indexcoluna = 1
        Next i
        ' Define o range dos dados preenchidos na terceira planilha
        Set rng = aba_3.Range("A1:S" & indexlinha - 1)
        ' Cria uma tabela a partir do range
        Set tbl_3 = aba_3.ListObjects.Add(xlSrcRange, rng, , xlYes)
        ' Ajusta automaticamente a largura das colunas
        aba_3.Columns("A:S").AutoFit
        ' Obt�m a �ltima linha preenchida na terceira planilha
        linha_fim_aba_abatimento_arquivo_cliente = aba_3.Range("A1048576").End(xlUp).Row
        ' Define a fonte da �ltima linha como negrito
        aba_3.Rows(linha_fim_aba_abatimento_arquivo_cliente).Font.Bold = True
        ' Define a fonte da coluna S como negrito
        aba_3.Columns("S:S").Font.Bold = True
    ' Sen�o, se a condi��o do payer for "reembolsados"
    ElseIf condicao_payer = "reembolsados" Then
        arquivo_cliente.Sheets("Cr�d disp a abater.reembolsar").Name = "Detalhe Reembolso"
    End If
    
    
    
End Sub

Public Sub CriarChamadoReembolsosAprovados()
    '***********************************************************
    ' Esta sub-rotina cria chamados no sistema Ellevo para reembolsos que foram aprovados
    ' Verifica os registros pendentes e j� aprovados, os processa no SAP e cria chamados
    ' correspondentes no sistema de gest�o Ellevo para dar sequ�ncia ao pagamento
    '***********************************************************
    
    ' Declara��o de vari�veis
    Dim tbl As ListObject          ' Objeto para armazenar tabela de dados
    Dim aba1 As Worksheet          ' Objeto que referencia a primeira aba da planilha
    Dim indexlinha, indexcoluna As Integer  ' Vari�veis para controle de posi��o na planilha
    Dim rng As Range               ' Objeto para definir um intervalo de c�lulas
    Dim pasta_diaria As String     ' Caminho da pasta onde ser�o salvos arquivos di�rios
    Dim arquivo_contas_a_pagar As Workbook  ' Refer�ncia ao arquivo Excel de contas a pagar

    ' Descri��o do processo completo:
    ' depois de realizar todo o processo na SBWP ir� verificar se existem linhas as quais estavam pendente de aprova��o e que foram aprovadas
    ' ENT�O BUSCAR� NA FBL5N AS PARTIDAS ABERTAS (COM ASIGNACION DIFERENTE DE "PROCESSADO AUTOMACAO") e USAR CHAVE DE REF 2 IGUAL A AUTOMACAO
    'ESSAS PARTIDAS ABERTAS QUER DIZER QUE J� FORAM APROVADOS PELO THIAGO OU LUANA
    ' NO EXCEL ENT�O VEREMOS A CORRESPOND�NCIA ENTRE A ABA AG APROVA��O E APROVADOS PARA ENTENDER AQUI SE PODEMOS OU N�O SEGUIR COM A CRIA��O DO CHAMADO ELLEVO DA DETERMINADA LINHA
    'NAS PARTIDAS ENCONTRADAS DO EXCEL, DEVEMOS ANTES DE EXPORTAR, TROCAR A ASIGNACION DE TODAS QUE N�O FORAM PROCESSADAS PARA "PROCESSADO AUTOMACAO" NO CAMPO ASIGNACION
    'PARA QUE NO MOMENTO DA FILTRAGEM DE PROXIMAS LEVAS, N�O PUXE ESSAS LINHAS
    
    ' Inicializa o array que ser� usado para armazenar os dados do cabe�alho
    array_linhas_chamado_contas_pagar = Array(array_cabecalho_arquivo_cliente)
    
    ' Encontra a �ltima linha com dados na aba de reembolsos pendentes
    linha_fim_aba_reembolsos_pendentes = aba_reembolsos_pendentes.Range("A1048576").End(xlUp).Row
    
    ' Verifica se h� dados para processar
    If linha_fim_aba_reembolsos_pendentes = 1 Then
        ' Se s� tem o cabe�alho (linha 1), n�o h� reembolsos pendentes
        Dicionario_Relatorio_Processamento("Chamado Ellevo de reembolso enviado para pagamento") = "Nenhum chamado de reembolso foi enviado ao contas a pagar"
        Exit Sub
    End If

    ' Acessa a transa��o FBL5N no SAP
    session.findById("wnd[0]/tbar[0]/okcd").text = "/N FBL5N"
    session.findById("wnd[0]").sendVKey 0
    
    ' Seleciona op��o de layout
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").Select
    
    ' Preenche campos para buscar o layout espec�fico
    session.findById("wnd[1]/usr/txtV-LOW").text = "id328"
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    ' Define a data limite (dois dias ap�s a data atual)
    session.findById("wnd[0]/usr/ctxtPA_STIDA").text = VBA.Format(Date + 2, tipo_data_sap)
    
    ' Bot�o para inserir valores do payers
    session.findById("wnd[0]/usr/btn%_DD_KUNNR_%_APP_%-VALU_PUSH").press
    
    ' Copia os payers da planilha e cola no SAP
    aba_reembolsos_pendentes.Range("C2:C" & linha_fim_aba_reembolsos_pendentes).Copy
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    ' Pressiona bot�o de crit�rios de sele��o
    session.findById("wnd[0]/tbar[1]/btn[16]").press
    
    ' Define o crit�rio de chave de refer�ncia como "AUTOMACAO DEV"
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/txt%%DYN002-LOW").text = "AUTOMACAO DEV"
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' Verifica se h� resultados na pesquisa
    If Left(session.findById("wnd[0]/sbar").text, 12) <> "S�o exibidas" Then
        Exit Sub
    End If
    
    ' Configura colunas e preenche o array com os dados da consulta SAP
    Call SetEixoXColunas
    Call PreencherArrayLinhasCondicaoAtual(session, i4, i5, "CHAMADO_CTA_A_PAGAR")

    payer_associado_OC = "VARIOS"
    ' Inicializa a vari�vel para soma de cr�ditos/devolu��es
    soma_cred_dev = 0
    
    ' Calcula a soma total dos valores na coluna 7 (montante)
    For i = LBound(array_linhas_chamado_contas_pagar) + 1 To UBound(array_linhas_chamado_contas_pagar)
        soma_cred_dev = soma_cred_dev + array_linhas_chamado_contas_pagar(i)(7)
    Next i
    
    ' Cria um novo arquivo Excel para armazenar os dados
    Set arquivo_contas_a_pagar = Workbooks.Add
    Set aba1 = arquivo_contas_a_pagar.Sheets(1)
    aba1.Name = "Reembolsos Pendentes"
    
    ' Inicializa array para armazenar os documentos F65
    array_docs_F65 = Array()
    
    ' Preenche o arquivo Excel com os dados e tamb�m coleta os n�meros dos documentos F65
    indexlinha = 1
    indexcoluna = 1
    For i = LBound(array_linhas_chamado_contas_pagar) To UBound(array_linhas_chamado_contas_pagar)
        indexcoluna = 1
        If i > 0 Then
            ' Adiciona o n�mero do documento (coluna 13) ao array de documentos F65
            array_docs_F65 = Add_ao_Array(array_docs_F65, array_linhas_chamado_contas_pagar(i)(13))
        End If
        
        ' Preenche cada coluna da linha atual com os dados do array
        For i2 = LBound(array_linhas_chamado_contas_pagar(i)) To UBound(array_linhas_chamado_contas_pagar(i))
            aba1.Cells(indexlinha, indexcoluna).Value = array_linhas_chamado_contas_pagar(i)(i2)
            indexcoluna = indexcoluna + 1
        Next i2
        indexlinha = indexlinha + 1
    Next i
    
    ' Formata os dados como uma tabela
    Set rng = aba1.Range("A1:P" & indexlinha - 1)
    Set tbl = aba1.ListObjects.Add(xlSrcRange, rng, , xlYes)
    
    ' Ajusta a largura das colunas e formata os valores monet�rios
    aba1.Columns("A:P").AutoFit
    aba1.Columns("H:H").NumberFormat = "#,###,###.##"
    
    ' Define o caminho para salvar o arquivo di�rio
    pasta_diaria = pasta_arquivos_clientes & "\" & VBA.Format(VBA.Date, "dd.mm.yyyy")
    caminho_arquivo = pasta_diaria & "\Reembolsos Pendentes.xlsx"
    
    ' Cria a pasta di�ria se n�o existir
    If Dir(pasta_diaria, vbDirectory) = "" Then
        MkDir (pasta_diaria)
    End If
    
    ' Remove o arquivo se j� existir
    If Dir(caminho_arquivo, vbDirectory) <> "" Then
        Kill caminho_arquivo
    End If
    
    ' Salva e fecha o arquivo
    arquivo_contas_a_pagar.SaveAs caminho_arquivo
    arquivo_contas_a_pagar.Close
    Set arquivo_contas_a_pagar = Nothing
    
    ' Cria o chamado no sistema Ellevo
    Call AbrirChamadoContasAPagar
    
    ' Atualiza a atribui��o no SAP com o n�mero do chamado Ellevo
    session.findById("wnd[0]").sendVKey 5
    Call AlterarAtribuicao(session, "ELLEVO#" & chamado_ellevo_aberto_contas_pagar)
    
    ' Registra o n�mero do chamado no dicion�rio de relat�rio de processamento
    Call AlimentarDicionario_Relatorio_Processamento("Chamado Ellevo de reembolso enviado para pagamento: ", chamado_ellevo_aberto_contas_pagar)
    
    ' Obt�m a data agrupada de pagamento
    data_agrupado_pagamento = Form_SAP.txt_box_data_agrupado_pgto_SAP
    
    ' Processa cada documento F65 encontrado
    For i = LBound(array_docs_F65) To UBound(array_docs_F65)
        doc_f65 = array_docs_F65(i)
        
        ' Busca informa��es relacionadas ao documento na planilha de reembolsos pendentes
        chamado = Application.WorksheetFunction.VLookup(CLng(doc_f65), aba_reembolsos_pendentes.Columns("A:B"), 2, False)
        data_solicitacao_reembolso = CDate(Application.WorksheetFunction.VLookup(CLng(doc_f65), aba_reembolsos_pendentes.Columns("A:D"), 4, False))
        soma_cred_dev = Round(CCur(Application.WorksheetFunction.VLookup(CLng(doc_f65), aba_reembolsos_pendentes.Columns("A:F"), 6, False)), 2)
        qtde_NFD_OC_chamado = Application.WorksheetFunction.VLookup(CLng(doc_f65), aba_reembolsos_pendentes.Columns("A:G"), 7, False)
        
        ' Cria a notifica��o de que o reembolso foi aprovado
        Call CriarTramiteNotificacaoReembolsoAprovado
        
        ' Remove as linhas processadas da planilha de reembolsos pendentes
        linha_fim_aba_reembolsos_pendentes = aba_reembolsos_pendentes.Range("A1048576").End(xlUp).Row
        For i2 = linha_fim_aba_reembolsos_pendentes To 2 Step -1
            If CStr(aba_reembolsos_pendentes.Range("A" & i2).Value) = doc_f65 Then
                aba_reembolsos_pendentes.Rows(i2).Delete
            End If
        Next i2
    Next i
    
End Sub

Public Sub CriarArquivoAnexoReembolso()

    Dim arquivo_anexo_reembolso As Workbook
    Dim aba As Worksheet
    Dim indexlinha, indexcoluna As Integer
    Dim rng As Range
    Dim tbl As ListObject

    ' Obt�m a �ltima linha preenchida na coluna A da planilha "aba_reembolsos_pendentes"
    linha_fim_aba_reembolsos_pendentes = aba_reembolsos_pendentes.Range("A1048576").End(xlUp).Row

    ' Cria um novo workbook para o arquivo anexo de reembolso
    Set arquivo_anexo_reembolso = Workbooks.Add
    ' Define a primeira planilha do novo workbook
    Set aba = arquivo_anexo_reembolso.Sheets(1)

    ' Inicializa o �ndice da linha
    indexlinha = 1
    'Loop atrav�s do array de linhas abertas da FBL5N para preencher os dados
    indexcoluna = 1
    For i = LBound(array_geral_linhas_abertas_FBL5N) To UBound(array_geral_linhas_abertas_FBL5N)
        ' Reinicializa o �ndice da coluna para cada linha
        indexcoluna = 1
        ' Loop atrav�s dos itens de cada linha do array
        For i2 = LBound(array_geral_linhas_abertas_FBL5N(i)) To UBound(array_geral_linhas_abertas_FBL5N(i))
            ' Preenche a c�lula com o valor do array
            aba.Cells(indexlinha, indexcoluna).Value = array_geral_linhas_abertas_FBL5N(i)(i2)
            ' Incrementa o �ndice da coluna
            indexcoluna = indexcoluna + 1
        Next i2
        ' Incrementa o �ndice da linha para a pr�xima linha
        indexlinha = indexlinha + 1
    Next i

    ' Define o range dos dados preenchidos
    Set rng = aba.Range("A1:Q" & indexlinha - 1)
    ' Cria uma tabela a partir do range
    Set tbl = aba.ListObjects.Add(xlSrcRange, rng, , xlYes)
    ' Ajusta automaticamente a largura das colunas
    aba.Columns("A:Q").AutoFit
    ' Define o formato de n�mero para a coluna H
    aba.Columns("H:H").NumberFormat = "#,###,###.##"
    ' Define o nome da planilha
    aba.Name = "Detalhe Reembolso"


    ' Define o caminho do arquivo a ser salvo
    caminho_arquivo = pasta_anexos_detalhe_reembolso & "\" & doc_f65 & ".xlsx"
    ' Cria a pasta se n�o existir
    If Dir(pasta_anexos_detalhe_reembolso, vbDirectory) = "" Then
        MkDir (pasta_anexos_detalhe_reembolso)
    End If
    ' Exclui o arquivo se j� existir
    If Dir(caminho_arquivo, vbDirectory) <> "" Then
        Kill caminho_arquivo
    End If
    ' Salva o arquivo
    arquivo_anexo_reembolso.SaveAs caminho_arquivo
    ' Fecha o arquivo
    arquivo_anexo_reembolso.Close
    ' Libera a vari�vel do objeto
    Set arquivo_anexo_reembolso = Nothing

End Sub

