Attribute VB_Name = "Z_FUNCOES_PUBLICAS"

' essa função simplesmente preenche no dict de infos do chamado o tipo da info como chave e a designação como valor
Public Function PreencherDicionario() As Object
    Dim DictInfosUnicas As Object
    Dim tipo, designacao As String

    
    Set DictInfosUnicas = CreateObject("Scripting.Dictionary")

    ' Preenche o dicionário com chamados únicos
    For linha = 2 To linha_fim
        If chamado = VBA.Trim(aba_consolidado.Range("A" & linha).Value) Then
            tipo = aba_consolidado.Range("C" & linha).Value
            designacao = aba_consolidado.Range("D" & linha).Value
            DictInfosUnicas.Add tipo, designacao
        End If
    Next linha

    Set PreencherDicionario = DictInfosUnicas
End Function


' Esta função verifica e processa solicitações de reembolso na transação SBWP do SAP
' Parâmetros:
'   - session_number: Objeto que representa a sessão SAP
'   - tipo: String que define o tipo de processamento ("GERAL" ou "UNITARIA")
' Retorna: Boolean indicando se existem linhas para processar na SBWP
Public Function VerificarLinhasSBWP(ByVal session_number As Object, tipo As String) As Boolean

    Dim data_solicitacao_reembolso As String
    
    ' Inicializa o retorno da função como True (existem linhas para processar)
    VerificarLinhasSBWP = True
    ' Inicializa o array que armazenará os documentos F65
    array_docs_F65 = Array()
    
    If tipo = "GERAL" Then
        ' Processa todas as solicitações de reembolso pendentes
        
        ' Determina a última linha da planilha com dados
        linha_fim_aba_reembolsos_pendentes = aba_reembolsos_pendentes.Range("A1048576").End(xlUp).Row
        
        ' Percorre todas as linhas da aba de reembolsos pendentes
        For linha = 2 To linha_fim_aba_reembolsos_pendentes
            doc_f65 = aba_reembolsos_pendentes.Range("A" & linha).Value
            data_criacao = aba_reembolsos_pendentes.Range("D" & linha).Value
            status_SBWP = aba_reembolsos_pendentes.Range("E" & linha).Value
            
            ' Verifica se o status está como "Não Solicitada Aprovação"
            If status_SBWP = "Não Solicitada Aprovação" Then
                ' Adiciona o documento ao array se ainda não existir
                If UBound(VBA.Filter(array_docs_F65, doc_f65)) < 0 Then
                    Call Add_ao_Array(array_docs_F65, doc_f65)
                End If
            End If
        Next linha
        
        ' Se não houver documentos para processar, sai da função
        If UBound(array_docs_F65) = -1 Then
            VerificarLinhasSBWP = False
            Call AlimentarDicionario_Relatorio_Processamento("Linhas Enviadas para Aprovação via Transação SBWP: ", "Nenhuma reembolso a ser enviado para aprovação")
            Exit Function
        End If
    ElseIf tipo = "UNITARIA" Then
        ' Processa apenas um documento específico
        ' Aguarda 30 segundos antes de continuar
        Application.Wait (Now + TimeValue("00:00:30"))
        ' Pega o último documento da lista
        array_docs_F65 = Array(aba_reembolsos_pendentes.Range("A1048576").End(xlUp).Value)
    End If
    
    ' Navega para a transação SBWP no SAP
    session_number.findById("wnd[0]/tbar[0]/okcd").text = "/N SBWP"
    session_number.findById("wnd[0]").sendVKey 0
    
    ' Expande os nós da árvore na interface do SAP
    session_number.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode "          2"
    session_number.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode "          5"
    session_number.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[0]/shell").topNode = "          1"
    session_number.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "          5"
    
    ' Define o elemento da tabela na interface do SAP
    Set elemento_tabela_SBWP = session_number.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell")
    
    ' Manipula as colunas da tabela
    elemento_tabela_SBWP.selectColumn "WI_CD"
    elemento_tabela_SBWP.pressColumnHeader "WI_CD"
    elemento_tabela_SBWP.selectColumn "WI_CT"
    elemento_tabela_SBWP.pressColumnHeader "WI_CT"
    elemento_tabela_SBWP.selectColumn "WI_CT"
    elemento_tabela_SBWP.pressColumnHeader "WI_CT"
    
    contador = 1
    
' Ponto de entrada para recomeçar a busca após processar um documento
recomecar_busca:
    ' Conta o número de linhas disponíveis na tabela
    For i = 0 To 1000000
        On Error Resume Next
        elemento_tabela_SBWP.setCurrentCell i, "WI_TEXT"
        elemento_tabela_SBWP.selectedRows = i
        If Err.number <> 0 Then
            linhas_SBWP = i
            Exit For
        End If
        On Error GoTo 0
    Next i
    
    ' Se não houver linhas, sai da função
    If linhas_SBWP = 0 Then
        If Not Dicionario_Relatorio_Processamento.exists("Linhas Enviadas para Aprovação via Transação SBWP: ") Then
            Call AlimentarDicionario_Relatorio_Processamento("Linhas Enviadas para Aprovação via Transação SBWP: ", "Nenhuma reembolso a ser enviado para aprovação")
            VerificarLinhasSBWP = False
        End If
            Exit Function
    End If

    ' Começa a processar os documentos encontrados
    elemento_tabela_SBWP.setCurrentCell 0, "WI_TEXT"
    elemento_tabela_SBWP.selectedRows = 0
    
    ' Percorre cada linha da tabela SBWP
    For i = 0 To linhas_SBWP - 1
        ' Extrai o número do documento F65 do texto
        doc_f65 = VBA.Right(elemento_tabela_SBWP.GetCellValue(i, "WI_TEXT"), 10)
        
        ' Verifica se o documento está no array de documentos a processar
        If UBound(VBA.Filter(array_docs_F65, CDbl(doc_f65))) >= 0 Then
            On Error Resume Next
            ' Verifica se o status do documento na planilha é "Não Solicitada Aprovação"
            If Application.VLookup(CLng(doc_f65), aba_reembolsos_pendentes.Columns("C:E"), 3, False) = "Não Solicitada Aprovação" Then
                ' Obtém a data de solicitação do reembolso
                data_solicitacao_reembolso = VBA.Format(Application.WorksheetFunction.VLookup(CLng(doc_f65), aba_reembolsos_pendentes.Columns("A:D"), 4, False), "dd.mm.yyyy")
            End If
            On Error GoTo 0
            
            ' Seleciona a linha e abre o documento
            elemento_tabela_SBWP.currentCellRow = i
            elemento_tabela_SBWP.selectedRows = i
            elemento_tabela_SBWP.doubleClickCurrentCell
            
            ' Adiciona um hífen ao texto do documento
            session_number.findById("wnd[0]/usr/txtBKPF-BKTXT").text = session_number.findById("wnd[0]/usr/txtBKPF-BKTXT").text & "-"
            
            ' Adiciona um anexo ao documento
            session_number.findById("wnd[0]/titl/shellcont/shell").pressContextButton "%GOS_TOOLBOX"
            session_number.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem "%GOS_PCATTA_CREA"
            session_number.findById("wnd[1]/usr/ctxtDY_PATH").text = Replace(pasta_anexos_detalhe_reembolso, VBA.Format(Date, "dd.mm.yyyy"), data_solicitacao_reembolso) & "\"
            session_number.findById("wnd[1]/usr/ctxtDY_FILENAME").text = doc_f65 & ".xlsx"
            session_number.findById("wnd[1]/tbar[0]/btn[0]").press
            
            ' Envia o documento para aprovação
            session_number.findById("wnd[0]/mbar/menu[0]/menu[6]").Select
            session_number.findById("wnd[1]/usr/ctxtG_INPUT").text = Form_SAP.approver
            session_number.findById("wnd[1]/usr/btnG_OK").press
            
            ' Atualiza o status do documento na planilha
            For i2 = 2 To linha_fim_aba_reembolsos_pendentes
                If aba_reembolsos_pendentes.Range("A" & i2).Value = CLng(doc_f65) Then
                    aba_reembolsos_pendentes.Range("E" & i2).Value = "Aguardando Aprovação"
                    Exit For
                End If
            Next i2
            
            ' Registra o processamento do documento
            Call AlimentarDicionario_Relatorio_Processamento("Linhas Enviadas para Aprovação via Transação SBWP: ", doc_f65)
            
            ' Comportamento baseado no tipo de processamento
            If tipo = "GERAL" Then
                ' Atualiza a visualização e recomeça a busca
                elemento_tabela_SBWP.pressToolbarButton "EREF"
                On Error Resume Next
                elemento_tabela_SBWP.currentCellRow = 0
                On Error GoTo 0
                GoTo recomecar_busca
            ElseIf tipo = "UNITARIA" Then
                ' Sai do loop se for processamento unitário
                Exit For
            End If
        End If
    Next i

End Function


' Esta função busca uma coluna específica em uma tabela SAP com base no texto do cabeçalho
' Parâmetros:
'   - indice_fixo_cabecalho: Posição vertical (Y) onde se encontram os cabeçalhos
'   - indice_inicio: Índice inicial horizontal (X) para começar a busca
'   - indice_fim: Índice final horizontal (X) para terminar a busca
'   - id_parte_inicial: String base para o ID do elemento SAP
'   - nome_coluna_buscada: Texto do cabeçalho que está sendo procurado
' Retorna: Índice da coluna encontrada ou 0 se não encontrar
Public Function VerificarColuna(indice_fixo_cabecalho As Integer, ByVal indice_inicio As Integer, indice_fim As Integer, id_parte_inicial As String, nome_coluna_buscada As String) As Integer
   Dim elemento_sap As Object
   Dim nome_coluna_atual As String
   Dim tentativas As Integer
   
   ' Loop com até 3 tentativas para encontrar a coluna
   Do
       ' Percorre horizontalmente as colunas possíveis
       For i = indice_inicio To indice_fim
           On Error Resume Next
           ' Tenta localizar o elemento de cabeçalho na interface SAP
           Set elemento_sap = session.findById(id_parte_inicial & i & "," & indice_fixo_cabecalho & "]")
           On Error GoTo 0
   
           ' Se o elemento foi encontrado, verifica o texto
           If Not elemento_sap Is Nothing Then
               nome_coluna_atual = Trim(elemento_sap.text)
               ' Se o texto corresponde ao buscado, retorna o índice
               If nome_coluna_atual = nome_coluna_buscada Then
                   VerificarColuna = i
                   Exit Function
               End If
           Else
               ' Se não encontrou o elemento, sai do loop
               Exit For
           End If
       Next i
   
       ' Pressiona tecla F22 (código 82) para navegar para a próxima página
       session.findById("wnd[0]").sendVKey 82
       ' Aguarda 1 segundo
       Application.Wait Now + TimeValue("0:00:01")
       ' Incrementa o contador de tentativas
       tentativas = tentativas + 1
       ' Sai se já fez 3 tentativas
       If tentativas > 3 Then Exit Do
   
   Loop
   ' Retorna 0 implicitamente se não encontrar a coluna
End Function



' Esta função verifica quantas linhas estão visíveis em uma tabela SAP
' Parâmetros:
'   - session_number: Objeto que representa a sessão SAP
'   - indice_inicio: Índice inicial vertical (Y) para começar a verificação
'   - indice_fim: Índice final vertical (Y) para terminar a verificação
'   - id_parte_inicial: String base para o ID do elemento SAP
' Retorna: Número de linhas visíveis na tabela
Public Function VerificarQuantidadeLinhasVisiveis(ByVal session_number As Object, indice_inicio As Integer, ByVal indice_fim As Integer, id_parte_inicial As String) As Integer
   
   'Dim elemento_sap As Object
   Dim i6 As Integer
   
   ' Inicializa o retorno com 0
   VerificarQuantidadeLinhasVisiveis = 0
   
   ' Percorre as linhas de indice_inicio até indice_fim
   For i6 = indice_inicio To indice_fim
       On Error Resume Next
       ' Tenta localizar o elemento na interface SAP usando variável global x_num_doc
       'Set elemento_sap = session_number.findById(id_parte_inicial & x_num_doc & "," & i6 & "]")
       session_number.findById(id_parte_inicial & x_num_doc & "," & i6 & "]").SetFocus
       If Err.number <> 0 Then
           ' Se encontrar erro, significa que chegou ao fim das linhas visíveis
           VerificarQuantidadeLinhasVisiveis = i6 - 1
           Exit Function
       End If
       On Error GoTo 0
   Next i6
   
   ' Se chegar até aqui sem erro, todas as linhas de indice_inicio até indice_fim estão visíveis
End Function
' Esta função calcula o número total de linhas em uma tabela SAP, incluindo as não visíveis
' que requerem rolagem usando a tecla F22 (Page Down)
' Parâmetros:
'   - session_number: Objeto que representa a sessão SAP
'   - indice_inicio: Índice inicial vertical (Y) para começar a contagem
'   - indice_fim: Índice final vertical (Y) máximo para terminar a contagem
'   - id_parte_inicial: String base para o ID do elemento SAP
' Retorna: Número total de linhas na tabela
Public Function VerificarQuantidadeLinhasTotais(ByVal session_number As Object, indice_inicio As Integer, ByVal indice_fim As Integer, id_parte_inicial As String) As Integer
   
   Dim elemento_sap As Object
   Dim primeiro_num_doc_item As String
   Dim i6 As Integer
   
   ' Inicializa o contador com 3 (offset inicial)
   VerificarQuantidadeLinhasTotais = 3
   
   ' Armazena o texto do primeiro documento + item como referência
   ' Usa as variáveis globais x_num_doc e x_item para identificar as colunas
   primeiro_num_doc_item = session_number.findById(id_parte_inicial & x_num_doc & "," & indice_inicio & "]").text & session_number.findById(id_parte_inicial & x_item & "," & indice_inicio & "]").text
   
   ' Percorre as linhas de indice_inicio até indice_fim
    For i6 = indice_inicio To indice_fim
    
        On Error Resume Next
        ' Tenta localizar o elemento na interface SAP
        session_number.findById(id_parte_inicial & x_num_doc & "," & i6 & "]").SetFocus
        
        If Err.number <> 0 Then
            Err.Clear
            session_number.findById("wnd[0]").sendVKey 82
            session_number.findById(id_parte_inicial & x_num_doc & "," & indice_inicio & "]").SetFocus
            If Err.number <> 0 Then
                Exit Function
            End If
            ' Verifica se voltou ao início (comparando com o primeiro documento+item)
            If primeiro_num_doc_item <> session_number.findById(id_parte_inicial & x_num_doc & "," & indice_inicio & "]").text & session_number.findById(id_parte_inicial & x_item & "," & indice_inicio & "]").text Then
                ' Se estiver em um novo conjunto de linhas, reinicia a contagem nesta página
                i6 = indice_inicio - 1
                primeiro_num_doc_item = session_number.findById(id_parte_inicial & x_num_doc & "," & indice_inicio & "]").text & session_number.findById(id_parte_inicial & x_item & "," & indice_inicio & "]").text
            Else
                ' Se voltou ao início, termina o loop
                Exit For
            End If
        Else
            ' Incrementa o contador de linhas totais
            VerificarQuantidadeLinhasTotais = VerificarQuantidadeLinhasTotais + 1
        End If
        On Error GoTo 0
       
       
   Next i6
End Function



' Esta função verifica e ajusta o formato padrão de data e decimal no SAP
' Parâmetros: Nenhum
' Retorna: String contendo o tipo de formato de data (dd/mm/yyyy, mm/dd/yyyy, yyyy/mm/dd)
Public Function VerificarFormatoPadraoSAP() As String
   Dim tipo As String
   
   ' Abre a transação SU3 (Dados do usuário)
   session.findById("wnd[0]").SetFocus
   session.findById("wnd[0]/tbar[0]/okcd").text = "/N SU3"
   session.findById("wnd[0]").sendVKey 0
   
   ' Seleciona a aba de padrões
   session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA").Select
   
   ' Obtém o formato de data e substitui "A" por "Y" (formato de ano)
   tipo = Trim(LCase(Replace(session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-DATFM").text, "A", "Y")))
   
   ' Verifica se o formato decimal precisa ser ajustado
   If session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-DCPFM").Key <> "" Then
       ' Define o formato decimal para padrão (vazio)
       session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-DCPFM").Key = ""
       
       ' Salva as alterações
       session.findById("wnd[0]/tbar[0]/btn[11]").press
       
       ' Fecha todas as sessões SAP abertas e reconecta
       i = connection.Children.Count - 1
       Do Until connection.Children(CInt(i)) Is Nothing
           Set session = connection.Children(CInt(i))
           session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
           session.findById("wnd[0]").sendVKey 0
           session.findById("wnd[0]").Close
           
           ' Verifica se há diálogo de confirmação e confirma
           On Error Resume Next
           session.findById("wnd[1]").SetFocus
           If Err.number = 0 Then
               session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
               
               ' Reabre uma nova conexão após fechar tudo
               Set SapGui = GetObject("SAPGUI")
               Set Applic = SapGui.GetScriptingEngine
               Set connection = Applic.OpenConnection("002. P1L - SAP ECC Latin America (Single Sign On)", True)
               Set session = connection.Children(0)
               Exit Do
           End If
           On Error GoTo 0
           i = i - 1
       Loop
   End If
   
   ' Retorna o tipo de formato de data
   VerificarFormatoPadraoSAP = tipo
   
End Function
Public Function InteracaoTelasSAP(ByRef session_number As Object, number As Integer, transacao As String) As Object

    ' Imprime o número atual de filhos (sessões) na conexão para depuração.
    Debug.Print connection.Children.Count

    ' Verifica se o número de sessões existentes é menor que o número desejado.
    If connection.Children.Count < number Then

        ' Se não houver sessões suficientes, cria uma nova.
        session.createsession

        ' Espera até que o número desejado de sessões seja estabelecido.
        Do Until connection.Children.Count = number
            ' Pausa a execução por 5 segundos.
            Application.Wait (Now + TimeValue("00:00:05"))
        Loop

        ' Imprime o número atualizado de filhos (sessões) após a criação.
        Debug.Print connection.Children.Count

        ' Loop através de cada uma das sessões filhas.
        For i = 0 To connection.Children.Count - 1
            ' Verifica se a sessão atual é o Gerenciador de Sessões.
            If connection.Children(CInt(i)).info.Transaction = "SESSION_MANAGER" Then
                ' Se for o Gerenciador de Sessões, atribui ao objeto session_number.
                Set session_number = connection.Children(CInt(i))
            End If
        Next i

        ' Insere um código de transação na sessão selecionada. "/N " geralmente indica uma nova transação.
        session_number.findById("wnd[0]/tbar[0]/okcd").text = "/N " & transacao
        ' Simula pressionar a tecla Enter (VKey 0) para executar a transação.
        session_number.findById("wnd[0]").sendVKey 0

    Else ' Se o número desejado de sessões já existir

        ' Imprime o número atual de sessões menos 1.
        Debug.Print connection.Children.Count - 1
        ' Loop através de cada uma das sessões existentes.
        For i = 0 To connection.Children.Count - 1
            ' Imprime o código de transação da sessão atual.
            Debug.Print connection.Children(CInt(i)).info.Transaction
            ' Verifica se a sessão atual é a transação SU3 (Manutenção de Usuário).
            If connection.Children(CInt(i)).info.Transaction = "SU3" Then
                ' Se for SU3, atribui ao objeto session.
                Set session = connection.Children(CInt(i))
                ' Tratamento de erro: Se a próxima sessão não existir, continua.
                On Error Resume Next
                ' Tenta definir o session_number para uma sessão deslocada por 'number - 1'.
                Set session_number = connection.Children(CInt(i + number - 1))
                 ' Se o session_number não foi atribuído corretamente (por exemplo, índice fora dos limites).
                If session_number Is Nothing Then
                    ' Tenta definir o session_number para uma sessão deslocada por 'number - 2'.
                    Set session_number = connection.Children(CInt(i + number - 2))
                End If
                ' Restaura o tratamento de erro padrão.
                On Error GoTo 0
                ' Insere o código da transação.
                session_number.findById("wnd[0]/tbar[0]/okcd").text = "/N " & transacao
                ' Simula pressionar Enter.
                session_number.findById("wnd[0]").sendVKey 0
                ' Sai do loop depois de encontrar e usar a sessão SU3.
                Exit For
            End If
        Next i
    End If

    ' Define o valor de retorno da função para o objeto de sessão selecionado.
    Set InteracaoTelasSAP = session_number

End Function

' Função para adicionar um item de string a um array dinâmico.
Public Function Add_ao_Array(ByRef array_() As Variant, ByVal item As String) As Variant
    ' Redimensiona o array, preservando os elementos existentes e adicionando mais um elemento.
    ReDim Preserve array_(LBound(array_) To UBound(array_) + 1)
    ' Adiciona o novo item ao último elemento do array.
    array_(UBound(array_)) = item
    ' Retorna o array modificado.
    Add_ao_Array = array_
End Function

Public Function TratativasOC(ByVal numero_OC As String) As String

    Dim resultado As String
    Dim char_init As String
    Dim j As Integer
    Dim startNum As Integer
    Dim endNum As Integer
    
    ' Inicializa variáveis
    resultado = ""
    startNum = 0
    endNum = 0

    ' Verifica se a string de referência não está vazia
    If Len(numero_OC) = 0 Then
        TratativasOC = resultado
        Exit Function
    End If

    ' Localiza o índice inicial da sequência numérica
    For j = 1 To Len(numero_OC)
        char_init = Mid(numero_OC, j, 1)
        
        If IsNumeric(char_init) And Val(char_init) <> 0 Then
            startNum = j ' Marca o índice inicial dos números
            Exit For
        End If
    Next j

    ' Se nenhum número foi encontrado, retorna string vazia
    If startNum = 0 Then
        TratativasOC = resultado
        Exit Function
    End If

    ' Localiza o índice final da sequência numérica
    For j = startNum To Len(numero_OC)
        char_init = Mid(numero_OC, j, 1)
        
        If Not IsNumeric(char_init) Then
            endNum = j - 1 ' Marca o índice final dos números
            Exit For
        End If
    Next j

    ' Caso o final não tenha sido definido (números até o final da string)
    If endNum = 0 Then endNum = Len(numero_OC)

    ' Extrai a sequência numérica com base nos índices encontrados
    resultado = Mid(numero_OC, startNum, endNum - startNum + 1)

    ' Retorna o resultado
    TratativasOC = resultado

End Function

Public Function VerificarContaBloqueada(transacao As String) As Boolean

    ' Inicializa a variável de controle de conta bloqueada como Falso
    conta_bloqueada = False
    ' Define o valor de retorno da função como Falso por padrão
    VerificarContaBloqueada = False

    ' Se a transação a ser verificada for "F-32" (Compensar Contas de Cliente)
    If transacao = "F-32" Then

        ' Entra na transação F-32 na terceira sessão
        session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N F-32"
        ' Simula a tecla Enter
        session_3.findById("wnd[0]").sendVKey 0
        ' Seleciona o radio button para contas de cliente
        session_3.findById("wnd[0]/usr/sub:SAPMF05A:0131/radRF05A-XPOS1[2,0]").Select
        ' Preenche o campo de cliente com o payer associado à OC
        session_3.findById("wnd[0]/usr/ctxtRF05A-AGKON").text = payer_associado_OC
        ' Preenche o campo de data de lançamento com a data atual no formato SAP
        session_3.findById("wnd[0]/usr/ctxtBKPF-BUDAT").text = Format(Date, tipo_data_sap)
        ' Preenche o campo de mês do documento com o mês atual
        session_3.findById("wnd[0]/usr/txtBKPF-MONAT").text = Month(Date)
        ' Preenche o campo de código da empresa com "BR10"
        session_3.findById("wnd[0]/usr/ctxtBKPF-BUKRS").text = "BR10"
        ' Preenche o campo de moeda com "BRL"
        session_3.findById("wnd[0]/usr/ctxtBKPF-WAERS").text = "BRL"
        ' Clica no botão para exibir as partidas em aberto
        session_3.findById("wnd[0]/tbar[1]/btn[16]").press

        ' Se a barra de status não estiver vazia (indicando alguma mensagem, possivelmente erro de conta bloqueada)
        If session_3.findById("wnd[0]/sbar").text <> "" Then
            ' Chama a sub-rotina para registrar no relatório que o payer tem conta bloqueada para processamento na F-32
            Call AlimentarDicionario_Relatorio_Processamento("Payers com contas bloqueada para processamento na F-32: ", payer_associado_OC)
            ' Define a flag de conta bloqueada como Verdadeira
            conta_bloqueada = True
            ' Define o valor de retorno da função como Verdadeiro (conta bloqueada)
            VerificarContaBloqueada = True
            ' Simula a tecla F5 (processar) na segunda sessão
            session_2.findById("wnd[0]").sendVKey 5
            
            Call AlterarAtribuicao(session_2, "CTA BLOQUEADA")

            ' Simula a tecla F8 (executar) na segunda sessão
            session_2.findById("wnd[0]").sendVKey 80
        End If
    ' Senão, se a transação a ser verificada for "F-65" (Estorno de Pagamento)
    ElseIf transacao = "F-65" Then
        ' Entra na transação F-65 na terceira sessão
        session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N F-65"
        ' Simula a tecla Enter
        session_3.findById("wnd[0]").sendVKey 0
        ' Preenche os campos necessários para a transação de estorno
        session_3.findById("wnd[0]/usr/ctxtBKPF-BLDAT").text = Format(Date, tipo_data_sap)
        session_3.findById("wnd[0]/usr/ctxtBKPF-BUDAT").text = Format(Date, tipo_data_sap)
        session_3.findById("wnd[0]/usr/ctxtBKPF-BLART").text = "ZD"
        session_3.findById("wnd[0]/usr/ctxtBKPF-WAERS").text = "BRL"
        session_3.findById("wnd[0]/usr/txtBKPF-XBLNR").text = "REEMB AUTOMACAO"
        session_3.findById("wnd[0]/usr/ctxtBKPF-BUKRS").text = "BR10"
        session_3.findById("wnd[0]/usr/txtBKPF-MONAT").text = Month(Date)
        session_3.findById("wnd[0]/usr/txtBKPF-BKTXT").text = Replace(UCase(VBA.Environ("USERPROFILE")), "C:\USERS\", "")
        session_3.findById("wnd[0]/usr/ctxtRF05V-NEWBS").text = "02"
        session_3.findById("wnd[0]/usr/ctxtRF05V-NEWKO").text = payer_associado_OC
        ' Simula a tecla Enter
        session_3.findById("wnd[0]").sendVKey 0

        ' Se os últimos 29 caracteres da barra de status indicarem que a conta está bloqueada para contabilização
        If VBA.Right(session_3.findById("wnd[0]/sbar").text, 29) = "bloqueada para contabilização" Then
            ' Chama a sub-rotina para registrar no relatório que o payer tem conta bloqueada para processamento na F-65
            Call AlimentarDicionario_Relatorio_Processamento("Payers com contas bloqueada para processamento na F-65: ", payer_associado_OC)
            Call AlterarAtribuicao(session_2, "CTA BLOQUEADA")
            ' Define a flag de conta bloqueada como Verdadeira
            conta_bloqueada = True
            ' Define o valor de retorno da função como Verdadeiro (conta bloqueada)
            VerificarContaBloqueada = True
        End If
    End If

End Function

Public Function BuscarPasta(ByRef caminho_pasta As String, Citrix As Boolean) As String
    Dim fso As Object
    Dim pastaGeral As Object
    Dim pastaAutomatizacoesBIRPA As Object
    Dim pastaAutomacoesEllevo As Object
    Dim pastaProcessoDevolucao As Object
    Dim pastaArquivosClientes As Object
    
    ' Cria o objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Define a pasta inicial
    If Citrix Then
        On Error Resume Next
        Set pastaGeral = fso.getfolder(Replace(VBA.Environ("USERPROFILE"), "C:\", "\\Client\C$\") & "\OneDrive - Electrolux\")
        On Error GoTo 0
        ' verificação pasta geral com onedrive - electrolux
        If Not VerificarPasta(pastaGeral) Then
            Set pastaGeral = fso.getfolder(Replace(VBA.Environ("USERPROFILE"), "C:\", "\\Client\C$\") & "\OneDrive\")
        End If
    ElseIf Not Citrix Then
        On Error Resume Next
        Set pastaGeral = fso.getfolder(VBA.Environ("USERPROFILE") & "\OneDrive - Electrolux\")
        On Error GoTo 0
        ' verificação pasta geral com onedrive - electrolux
        If Not VerificarPasta(pastaGeral) Then
            Set pastaGeral = fso.getfolder(VBA.Environ("USERPROFILE") & "\OneDrive\")
        End If
    End If
        
    ' verificação pasta geral só com onedrive
    If Not VerificarPasta(pastaGeral) Then GoTo pedir_caminho_manualmente
    
    On Error Resume Next
    ' verificação pasta Excelencia
    Set pastaAutomatizacoesBIRPA = fso.getfolder(pastaGeral.Path & "\AUTOMATIZAÇÕES, BIs & RPAs\")
    If Not VerificarPasta(pastaAutomatizacoesBIRPA) Then
        Set pastaAutomacoesEllevo = fso.getfolder(pastaGeral.Path & "\Automações Ellevo\")
    Else
        Set pastaAutomacoesEllevo = fso.getfolder(pastaAutomatizacoesBIRPA.Path & "\Automações Ellevo\")
    End If
    If Not VerificarPasta(pastaAutomacoesEllevo) Then
        Set pastaProcessoDevolucao = fso.getfolder(pastaGeral.Path & "\Processo Devolução\")
    Else
        Set pastaProcessoDevolucao = fso.getfolder(pastaAutomacoesEllevo.Path & "\Processo Devolução\")
    End If
    If Not VerificarPasta(pastaProcessoDevolucao) Then
        Set pastaArquivosClientes = fso.getfolder(pastaGeral.Path & "\Arquivos Clientes\")
    Else
        Set pastaArquivosClientes = fso.getfolder(pastaProcessoDevolucao.Path & "\Arquivos Clientes\")
    End If
    On Error GoTo 0
    If Not VerificarPasta(pastaArquivosClientes) Then
        GoTo pedir_caminho_manualmente
    Else
        caminho_pasta = pastaArquivosClientes.Path
        GoTo fim
    End If
    
pedir_caminho_manualmente:
    MsgBox "Favor escolher o caminho no seu Computador a ser descarregado o arquivo baixado do site Transbank. Se não possuir a pasta, crie o atalho no Sharepoint e execute novamente a automação." & _
        "(Escolha a pasta no seu computador equivalente à pasta do Sharepoint:" & _
            "Documentos > AUTOMATIZAÇÕES, BIs & RPAs > Automações Ellevo > Processo Devolução > Arquivos Clientes)"
     
    'Seleção da pasta para salvar o arquivo
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' O usuário selecionou uma pasta
            caminho_pasta = .SelectedItems(1) & "\"
        Else
             'O usuário cancelou a seleção da pasta
            MsgBox "Nenhuma pasta selecionada. O processo foi cancelado."
            End
        End If
    End With
    
fim:

    BuscarPasta = caminho_pasta
    ' Libera os objetos
    Set fso = Nothing
    Set pastaGeral = Nothing
    Set pastaAutomatizacoesBIRPA = Nothing
    Set pastaAutomacoesEllevo = Nothing
    Set pastaProcessoDevolucao = Nothing
    Set pastaArquivosClientes = Nothing
    
    
End Function
' funcao complementar da funcao BuscarPasta - apenas verifica se a pasta em questão existe
Public Function VerificarPasta(ByVal pasta As Object) As Boolean
    VerificarPasta = True
    If pasta Is Nothing Then
        VerificarPasta = False
    End If
End Function
