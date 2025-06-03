Attribute VB_Name = "Z_FUNCOES_PUBLICAS"

' essa fun��o simplesmente preenche no dict de infos do chamado o tipo da info como chave e a designa��o como valor
Public Function PreencherDicionario() As Object
    Dim DictInfosUnicas As Object
    Dim tipo, designacao As String

    
    Set DictInfosUnicas = CreateObject("Scripting.Dictionary")

    ' Preenche o dicion�rio com chamados �nicos
    For linha = 2 To linha_fim
        If chamado = VBA.Trim(aba_consolidado.Range("A" & linha).Value) Then
            tipo = aba_consolidado.Range("C" & linha).Value
            designacao = aba_consolidado.Range("D" & linha).Value
            DictInfosUnicas.Add tipo, designacao
        End If
    Next linha

    Set PreencherDicionario = DictInfosUnicas
End Function


' Esta fun��o verifica e processa solicita��es de reembolso na transa��o SBWP do SAP
' Par�metros:
'   - session_number: Objeto que representa a sess�o SAP
'   - tipo: String que define o tipo de processamento ("GERAL" ou "UNITARIA")
' Retorna: Boolean indicando se existem linhas para processar na SBWP
Public Function VerificarLinhasSBWP(ByVal session_number As Object, tipo As String) As Boolean

    Dim data_solicitacao_reembolso As String
    
    ' Inicializa o retorno da fun��o como True (existem linhas para processar)
    VerificarLinhasSBWP = True
    ' Inicializa o array que armazenar� os documentos F65
    array_docs_F65 = Array()
    
    If tipo = "GERAL" Then
        ' Processa todas as solicita��es de reembolso pendentes
        
        ' Determina a �ltima linha da planilha com dados
        linha_fim_aba_reembolsos_pendentes = aba_reembolsos_pendentes.Range("A1048576").End(xlUp).Row
        
        ' Percorre todas as linhas da aba de reembolsos pendentes
        For linha = 2 To linha_fim_aba_reembolsos_pendentes
            doc_f65 = aba_reembolsos_pendentes.Range("A" & linha).Value
            data_criacao = aba_reembolsos_pendentes.Range("D" & linha).Value
            status_SBWP = aba_reembolsos_pendentes.Range("E" & linha).Value
            
            ' Verifica se o status est� como "N�o Solicitada Aprova��o"
            If status_SBWP = "N�o Solicitada Aprova��o" Then
                ' Adiciona o documento ao array se ainda n�o existir
                If UBound(VBA.Filter(array_docs_F65, doc_f65)) < 0 Then
                    Call Add_ao_Array(array_docs_F65, doc_f65)
                End If
            End If
        Next linha
        
        ' Se n�o houver documentos para processar, sai da fun��o
        If UBound(array_docs_F65) = -1 Then
            VerificarLinhasSBWP = False
            Call AlimentarDicionario_Relatorio_Processamento("Linhas Enviadas para Aprova��o via Transa��o SBWP: ", "Nenhuma reembolso a ser enviado para aprova��o")
            Exit Function
        End If
    ElseIf tipo = "UNITARIA" Then
        ' Processa apenas um documento espec�fico
        ' Aguarda 30 segundos antes de continuar
        Application.Wait (Now + TimeValue("00:00:30"))
        ' Pega o �ltimo documento da lista
        array_docs_F65 = Array(aba_reembolsos_pendentes.Range("A1048576").End(xlUp).Value)
    End If
    
    ' Navega para a transa��o SBWP no SAP
    session_number.findById("wnd[0]/tbar[0]/okcd").text = "/N SBWP"
    session_number.findById("wnd[0]").sendVKey 0
    
    ' Expande os n�s da �rvore na interface do SAP
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
    
' Ponto de entrada para recome�ar a busca ap�s processar um documento
recomecar_busca:
    ' Conta o n�mero de linhas dispon�veis na tabela
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
    
    ' Se n�o houver linhas, sai da fun��o
    If linhas_SBWP = 0 Then
        If Not Dicionario_Relatorio_Processamento.exists("Linhas Enviadas para Aprova��o via Transa��o SBWP: ") Then
            Call AlimentarDicionario_Relatorio_Processamento("Linhas Enviadas para Aprova��o via Transa��o SBWP: ", "Nenhuma reembolso a ser enviado para aprova��o")
            VerificarLinhasSBWP = False
        End If
            Exit Function
    End If

    ' Come�a a processar os documentos encontrados
    elemento_tabela_SBWP.setCurrentCell 0, "WI_TEXT"
    elemento_tabela_SBWP.selectedRows = 0
    
    ' Percorre cada linha da tabela SBWP
    For i = 0 To linhas_SBWP - 1
        ' Extrai o n�mero do documento F65 do texto
        doc_f65 = VBA.Right(elemento_tabela_SBWP.GetCellValue(i, "WI_TEXT"), 10)
        
        ' Verifica se o documento est� no array de documentos a processar
        If UBound(VBA.Filter(array_docs_F65, CDbl(doc_f65))) >= 0 Then
            On Error Resume Next
            ' Verifica se o status do documento na planilha � "N�o Solicitada Aprova��o"
            If Application.VLookup(CLng(doc_f65), aba_reembolsos_pendentes.Columns("C:E"), 3, False) = "N�o Solicitada Aprova��o" Then
                ' Obt�m a data de solicita��o do reembolso
                data_solicitacao_reembolso = VBA.Format(Application.WorksheetFunction.VLookup(CLng(doc_f65), aba_reembolsos_pendentes.Columns("A:D"), 4, False), "dd.mm.yyyy")
            End If
            On Error GoTo 0
            
            ' Seleciona a linha e abre o documento
            elemento_tabela_SBWP.currentCellRow = i
            elemento_tabela_SBWP.selectedRows = i
            elemento_tabela_SBWP.doubleClickCurrentCell
            
            ' Adiciona um h�fen ao texto do documento
            session_number.findById("wnd[0]/usr/txtBKPF-BKTXT").text = session_number.findById("wnd[0]/usr/txtBKPF-BKTXT").text & "-"
            
            ' Adiciona um anexo ao documento
            session_number.findById("wnd[0]/titl/shellcont/shell").pressContextButton "%GOS_TOOLBOX"
            session_number.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem "%GOS_PCATTA_CREA"
            session_number.findById("wnd[1]/usr/ctxtDY_PATH").text = Replace(pasta_anexos_detalhe_reembolso, VBA.Format(Date, "dd.mm.yyyy"), data_solicitacao_reembolso) & "\"
            session_number.findById("wnd[1]/usr/ctxtDY_FILENAME").text = doc_f65 & ".xlsx"
            session_number.findById("wnd[1]/tbar[0]/btn[0]").press
            
            ' Envia o documento para aprova��o
            session_number.findById("wnd[0]/mbar/menu[0]/menu[6]").Select
            session_number.findById("wnd[1]/usr/ctxtG_INPUT").text = Form_SAP.approver
            session_number.findById("wnd[1]/usr/btnG_OK").press
            
            ' Atualiza o status do documento na planilha
            For i2 = 2 To linha_fim_aba_reembolsos_pendentes
                If aba_reembolsos_pendentes.Range("A" & i2).Value = CLng(doc_f65) Then
                    aba_reembolsos_pendentes.Range("E" & i2).Value = "Aguardando Aprova��o"
                    Exit For
                End If
            Next i2
            
            ' Registra o processamento do documento
            Call AlimentarDicionario_Relatorio_Processamento("Linhas Enviadas para Aprova��o via Transa��o SBWP: ", doc_f65)
            
            ' Comportamento baseado no tipo de processamento
            If tipo = "GERAL" Then
                ' Atualiza a visualiza��o e recome�a a busca
                elemento_tabela_SBWP.pressToolbarButton "EREF"
                On Error Resume Next
                elemento_tabela_SBWP.currentCellRow = 0
                On Error GoTo 0
                GoTo recomecar_busca
            ElseIf tipo = "UNITARIA" Then
                ' Sai do loop se for processamento unit�rio
                Exit For
            End If
        End If
    Next i

End Function


' Esta fun��o busca uma coluna espec�fica em uma tabela SAP com base no texto do cabe�alho
' Par�metros:
'   - indice_fixo_cabecalho: Posi��o vertical (Y) onde se encontram os cabe�alhos
'   - indice_inicio: �ndice inicial horizontal (X) para come�ar a busca
'   - indice_fim: �ndice final horizontal (X) para terminar a busca
'   - id_parte_inicial: String base para o ID do elemento SAP
'   - nome_coluna_buscada: Texto do cabe�alho que est� sendo procurado
' Retorna: �ndice da coluna encontrada ou 0 se n�o encontrar
Public Function VerificarColuna(indice_fixo_cabecalho As Integer, ByVal indice_inicio As Integer, indice_fim As Integer, id_parte_inicial As String, nome_coluna_buscada As String) As Integer
   Dim elemento_sap As Object
   Dim nome_coluna_atual As String
   Dim tentativas As Integer
   
   ' Loop com at� 3 tentativas para encontrar a coluna
   Do
       ' Percorre horizontalmente as colunas poss�veis
       For i = indice_inicio To indice_fim
           On Error Resume Next
           ' Tenta localizar o elemento de cabe�alho na interface SAP
           Set elemento_sap = session.findById(id_parte_inicial & i & "," & indice_fixo_cabecalho & "]")
           On Error GoTo 0
   
           ' Se o elemento foi encontrado, verifica o texto
           If Not elemento_sap Is Nothing Then
               nome_coluna_atual = Trim(elemento_sap.text)
               ' Se o texto corresponde ao buscado, retorna o �ndice
               If nome_coluna_atual = nome_coluna_buscada Then
                   VerificarColuna = i
                   Exit Function
               End If
           Else
               ' Se n�o encontrou o elemento, sai do loop
               Exit For
           End If
       Next i
   
       ' Pressiona tecla F22 (c�digo 82) para navegar para a pr�xima p�gina
       session.findById("wnd[0]").sendVKey 82
       ' Aguarda 1 segundo
       Application.Wait Now + TimeValue("0:00:01")
       ' Incrementa o contador de tentativas
       tentativas = tentativas + 1
       ' Sai se j� fez 3 tentativas
       If tentativas > 3 Then Exit Do
   
   Loop
   ' Retorna 0 implicitamente se n�o encontrar a coluna
End Function



' Esta fun��o verifica quantas linhas est�o vis�veis em uma tabela SAP
' Par�metros:
'   - session_number: Objeto que representa a sess�o SAP
'   - indice_inicio: �ndice inicial vertical (Y) para come�ar a verifica��o
'   - indice_fim: �ndice final vertical (Y) para terminar a verifica��o
'   - id_parte_inicial: String base para o ID do elemento SAP
' Retorna: N�mero de linhas vis�veis na tabela
Public Function VerificarQuantidadeLinhasVisiveis(ByVal session_number As Object, indice_inicio As Integer, ByVal indice_fim As Integer, id_parte_inicial As String) As Integer
   
   'Dim elemento_sap As Object
   Dim i6 As Integer
   
   ' Inicializa o retorno com 0
   VerificarQuantidadeLinhasVisiveis = 0
   
   ' Percorre as linhas de indice_inicio at� indice_fim
   For i6 = indice_inicio To indice_fim
       On Error Resume Next
       ' Tenta localizar o elemento na interface SAP usando vari�vel global x_num_doc
       'Set elemento_sap = session_number.findById(id_parte_inicial & x_num_doc & "," & i6 & "]")
       session_number.findById(id_parte_inicial & x_num_doc & "," & i6 & "]").SetFocus
       If Err.number <> 0 Then
           ' Se encontrar erro, significa que chegou ao fim das linhas vis�veis
           VerificarQuantidadeLinhasVisiveis = i6 - 1
           Exit Function
       End If
       On Error GoTo 0
   Next i6
   
   ' Se chegar at� aqui sem erro, todas as linhas de indice_inicio at� indice_fim est�o vis�veis
End Function
' Esta fun��o calcula o n�mero total de linhas em uma tabela SAP, incluindo as n�o vis�veis
' que requerem rolagem usando a tecla F22 (Page Down)
' Par�metros:
'   - session_number: Objeto que representa a sess�o SAP
'   - indice_inicio: �ndice inicial vertical (Y) para come�ar a contagem
'   - indice_fim: �ndice final vertical (Y) m�ximo para terminar a contagem
'   - id_parte_inicial: String base para o ID do elemento SAP
' Retorna: N�mero total de linhas na tabela
Public Function VerificarQuantidadeLinhasTotais(ByVal session_number As Object, indice_inicio As Integer, ByVal indice_fim As Integer, id_parte_inicial As String) As Integer
   
   Dim elemento_sap As Object
   Dim primeiro_num_doc_item As String
   Dim i6 As Integer
   
   ' Inicializa o contador com 3 (offset inicial)
   VerificarQuantidadeLinhasTotais = 3
   
   ' Armazena o texto do primeiro documento + item como refer�ncia
   ' Usa as vari�veis globais x_num_doc e x_item para identificar as colunas
   primeiro_num_doc_item = session_number.findById(id_parte_inicial & x_num_doc & "," & indice_inicio & "]").text & session_number.findById(id_parte_inicial & x_item & "," & indice_inicio & "]").text
   
   ' Percorre as linhas de indice_inicio at� indice_fim
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
            ' Verifica se voltou ao in�cio (comparando com o primeiro documento+item)
            If primeiro_num_doc_item <> session_number.findById(id_parte_inicial & x_num_doc & "," & indice_inicio & "]").text & session_number.findById(id_parte_inicial & x_item & "," & indice_inicio & "]").text Then
                ' Se estiver em um novo conjunto de linhas, reinicia a contagem nesta p�gina
                i6 = indice_inicio - 1
                primeiro_num_doc_item = session_number.findById(id_parte_inicial & x_num_doc & "," & indice_inicio & "]").text & session_number.findById(id_parte_inicial & x_item & "," & indice_inicio & "]").text
            Else
                ' Se voltou ao in�cio, termina o loop
                Exit For
            End If
        Else
            ' Incrementa o contador de linhas totais
            VerificarQuantidadeLinhasTotais = VerificarQuantidadeLinhasTotais + 1
        End If
        On Error GoTo 0
       
       
   Next i6
End Function



' Esta fun��o verifica e ajusta o formato padr�o de data e decimal no SAP
' Par�metros: Nenhum
' Retorna: String contendo o tipo de formato de data (dd/mm/yyyy, mm/dd/yyyy, yyyy/mm/dd)
Public Function VerificarFormatoPadraoSAP() As String
   Dim tipo As String
   
   ' Abre a transa��o SU3 (Dados do usu�rio)
   session.findById("wnd[0]").SetFocus
   session.findById("wnd[0]/tbar[0]/okcd").text = "/N SU3"
   session.findById("wnd[0]").sendVKey 0
   
   ' Seleciona a aba de padr�es
   session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA").Select
   
   ' Obt�m o formato de data e substitui "A" por "Y" (formato de ano)
   tipo = Trim(LCase(Replace(session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-DATFM").text, "A", "Y")))
   
   ' Verifica se o formato decimal precisa ser ajustado
   If session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-DCPFM").Key <> "" Then
       ' Define o formato decimal para padr�o (vazio)
       session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-DCPFM").Key = ""
       
       ' Salva as altera��es
       session.findById("wnd[0]/tbar[0]/btn[11]").press
       
       ' Fecha todas as sess�es SAP abertas e reconecta
       i = connection.Children.Count - 1
       Do Until connection.Children(CInt(i)) Is Nothing
           Set session = connection.Children(CInt(i))
           session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
           session.findById("wnd[0]").sendVKey 0
           session.findById("wnd[0]").Close
           
           ' Verifica se h� di�logo de confirma��o e confirma
           On Error Resume Next
           session.findById("wnd[1]").SetFocus
           If Err.number = 0 Then
               session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
               
               ' Reabre uma nova conex�o ap�s fechar tudo
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

    ' Imprime o n�mero atual de filhos (sess�es) na conex�o para depura��o.
    Debug.Print connection.Children.Count

    ' Verifica se o n�mero de sess�es existentes � menor que o n�mero desejado.
    If connection.Children.Count < number Then

        ' Se n�o houver sess�es suficientes, cria uma nova.
        session.createsession

        ' Espera at� que o n�mero desejado de sess�es seja estabelecido.
        Do Until connection.Children.Count = number
            ' Pausa a execu��o por 5 segundos.
            Application.Wait (Now + TimeValue("00:00:05"))
        Loop

        ' Imprime o n�mero atualizado de filhos (sess�es) ap�s a cria��o.
        Debug.Print connection.Children.Count

        ' Loop atrav�s de cada uma das sess�es filhas.
        For i = 0 To connection.Children.Count - 1
            ' Verifica se a sess�o atual � o Gerenciador de Sess�es.
            If connection.Children(CInt(i)).info.Transaction = "SESSION_MANAGER" Then
                ' Se for o Gerenciador de Sess�es, atribui ao objeto session_number.
                Set session_number = connection.Children(CInt(i))
            End If
        Next i

        ' Insere um c�digo de transa��o na sess�o selecionada. "/N " geralmente indica uma nova transa��o.
        session_number.findById("wnd[0]/tbar[0]/okcd").text = "/N " & transacao
        ' Simula pressionar a tecla Enter (VKey 0) para executar a transa��o.
        session_number.findById("wnd[0]").sendVKey 0

    Else ' Se o n�mero desejado de sess�es j� existir

        ' Imprime o n�mero atual de sess�es menos 1.
        Debug.Print connection.Children.Count - 1
        ' Loop atrav�s de cada uma das sess�es existentes.
        For i = 0 To connection.Children.Count - 1
            ' Imprime o c�digo de transa��o da sess�o atual.
            Debug.Print connection.Children(CInt(i)).info.Transaction
            ' Verifica se a sess�o atual � a transa��o SU3 (Manuten��o de Usu�rio).
            If connection.Children(CInt(i)).info.Transaction = "SU3" Then
                ' Se for SU3, atribui ao objeto session.
                Set session = connection.Children(CInt(i))
                ' Tratamento de erro: Se a pr�xima sess�o n�o existir, continua.
                On Error Resume Next
                ' Tenta definir o session_number para uma sess�o deslocada por 'number - 1'.
                Set session_number = connection.Children(CInt(i + number - 1))
                 ' Se o session_number n�o foi atribu�do corretamente (por exemplo, �ndice fora dos limites).
                If session_number Is Nothing Then
                    ' Tenta definir o session_number para uma sess�o deslocada por 'number - 2'.
                    Set session_number = connection.Children(CInt(i + number - 2))
                End If
                ' Restaura o tratamento de erro padr�o.
                On Error GoTo 0
                ' Insere o c�digo da transa��o.
                session_number.findById("wnd[0]/tbar[0]/okcd").text = "/N " & transacao
                ' Simula pressionar Enter.
                session_number.findById("wnd[0]").sendVKey 0
                ' Sai do loop depois de encontrar e usar a sess�o SU3.
                Exit For
            End If
        Next i
    End If

    ' Define o valor de retorno da fun��o para o objeto de sess�o selecionado.
    Set InteracaoTelasSAP = session_number

End Function

' Fun��o para adicionar um item de string a um array din�mico.
Public Function Add_ao_Array(ByRef array_() As Variant, ByVal item As String) As Variant
    ' Redimensiona o array, preservando os elementos existentes e adicionando mais um elemento.
    ReDim Preserve array_(LBound(array_) To UBound(array_) + 1)
    ' Adiciona o novo item ao �ltimo elemento do array.
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
    
    ' Inicializa vari�veis
    resultado = ""
    startNum = 0
    endNum = 0

    ' Verifica se a string de refer�ncia n�o est� vazia
    If Len(numero_OC) = 0 Then
        TratativasOC = resultado
        Exit Function
    End If

    ' Localiza o �ndice inicial da sequ�ncia num�rica
    For j = 1 To Len(numero_OC)
        char_init = Mid(numero_OC, j, 1)
        
        If IsNumeric(char_init) And Val(char_init) <> 0 Then
            startNum = j ' Marca o �ndice inicial dos n�meros
            Exit For
        End If
    Next j

    ' Se nenhum n�mero foi encontrado, retorna string vazia
    If startNum = 0 Then
        TratativasOC = resultado
        Exit Function
    End If

    ' Localiza o �ndice final da sequ�ncia num�rica
    For j = startNum To Len(numero_OC)
        char_init = Mid(numero_OC, j, 1)
        
        If Not IsNumeric(char_init) Then
            endNum = j - 1 ' Marca o �ndice final dos n�meros
            Exit For
        End If
    Next j

    ' Caso o final n�o tenha sido definido (n�meros at� o final da string)
    If endNum = 0 Then endNum = Len(numero_OC)

    ' Extrai a sequ�ncia num�rica com base nos �ndices encontrados
    resultado = Mid(numero_OC, startNum, endNum - startNum + 1)

    ' Retorna o resultado
    TratativasOC = resultado

End Function

Public Function VerificarContaBloqueada(transacao As String) As Boolean

    ' Inicializa a vari�vel de controle de conta bloqueada como Falso
    conta_bloqueada = False
    ' Define o valor de retorno da fun��o como Falso por padr�o
    VerificarContaBloqueada = False

    ' Se a transa��o a ser verificada for "F-32" (Compensar Contas de Cliente)
    If transacao = "F-32" Then

        ' Entra na transa��o F-32 na terceira sess�o
        session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N F-32"
        ' Simula a tecla Enter
        session_3.findById("wnd[0]").sendVKey 0
        ' Seleciona o radio button para contas de cliente
        session_3.findById("wnd[0]/usr/sub:SAPMF05A:0131/radRF05A-XPOS1[2,0]").Select
        ' Preenche o campo de cliente com o payer associado � OC
        session_3.findById("wnd[0]/usr/ctxtRF05A-AGKON").text = payer_associado_OC
        ' Preenche o campo de data de lan�amento com a data atual no formato SAP
        session_3.findById("wnd[0]/usr/ctxtBKPF-BUDAT").text = Format(Date, tipo_data_sap)
        ' Preenche o campo de m�s do documento com o m�s atual
        session_3.findById("wnd[0]/usr/txtBKPF-MONAT").text = Month(Date)
        ' Preenche o campo de c�digo da empresa com "BR10"
        session_3.findById("wnd[0]/usr/ctxtBKPF-BUKRS").text = "BR10"
        ' Preenche o campo de moeda com "BRL"
        session_3.findById("wnd[0]/usr/ctxtBKPF-WAERS").text = "BRL"
        ' Clica no bot�o para exibir as partidas em aberto
        session_3.findById("wnd[0]/tbar[1]/btn[16]").press

        ' Se a barra de status n�o estiver vazia (indicando alguma mensagem, possivelmente erro de conta bloqueada)
        If session_3.findById("wnd[0]/sbar").text <> "" Then
            ' Chama a sub-rotina para registrar no relat�rio que o payer tem conta bloqueada para processamento na F-32
            Call AlimentarDicionario_Relatorio_Processamento("Payers com contas bloqueada para processamento na F-32: ", payer_associado_OC)
            ' Define a flag de conta bloqueada como Verdadeira
            conta_bloqueada = True
            ' Define o valor de retorno da fun��o como Verdadeiro (conta bloqueada)
            VerificarContaBloqueada = True
            ' Simula a tecla F5 (processar) na segunda sess�o
            session_2.findById("wnd[0]").sendVKey 5
            
            Call AlterarAtribuicao(session_2, "CTA BLOQUEADA")

            ' Simula a tecla F8 (executar) na segunda sess�o
            session_2.findById("wnd[0]").sendVKey 80
        End If
    ' Sen�o, se a transa��o a ser verificada for "F-65" (Estorno de Pagamento)
    ElseIf transacao = "F-65" Then
        ' Entra na transa��o F-65 na terceira sess�o
        session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N F-65"
        ' Simula a tecla Enter
        session_3.findById("wnd[0]").sendVKey 0
        ' Preenche os campos necess�rios para a transa��o de estorno
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

        ' Se os �ltimos 29 caracteres da barra de status indicarem que a conta est� bloqueada para contabiliza��o
        If VBA.Right(session_3.findById("wnd[0]/sbar").text, 29) = "bloqueada para contabiliza��o" Then
            ' Chama a sub-rotina para registrar no relat�rio que o payer tem conta bloqueada para processamento na F-65
            Call AlimentarDicionario_Relatorio_Processamento("Payers com contas bloqueada para processamento na F-65: ", payer_associado_OC)
            Call AlterarAtribuicao(session_2, "CTA BLOQUEADA")
            ' Define a flag de conta bloqueada como Verdadeira
            conta_bloqueada = True
            ' Define o valor de retorno da fun��o como Verdadeiro (conta bloqueada)
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
        ' verifica��o pasta geral com onedrive - electrolux
        If Not VerificarPasta(pastaGeral) Then
            Set pastaGeral = fso.getfolder(Replace(VBA.Environ("USERPROFILE"), "C:\", "\\Client\C$\") & "\OneDrive\")
        End If
    ElseIf Not Citrix Then
        On Error Resume Next
        Set pastaGeral = fso.getfolder(VBA.Environ("USERPROFILE") & "\OneDrive - Electrolux\")
        On Error GoTo 0
        ' verifica��o pasta geral com onedrive - electrolux
        If Not VerificarPasta(pastaGeral) Then
            Set pastaGeral = fso.getfolder(VBA.Environ("USERPROFILE") & "\OneDrive\")
        End If
    End If
        
    ' verifica��o pasta geral s� com onedrive
    If Not VerificarPasta(pastaGeral) Then GoTo pedir_caminho_manualmente
    
    On Error Resume Next
    ' verifica��o pasta Excelencia
    Set pastaAutomatizacoesBIRPA = fso.getfolder(pastaGeral.Path & "\AUTOMATIZA��ES, BIs & RPAs\")
    If Not VerificarPasta(pastaAutomatizacoesBIRPA) Then
        Set pastaAutomacoesEllevo = fso.getfolder(pastaGeral.Path & "\Automa��es Ellevo\")
    Else
        Set pastaAutomacoesEllevo = fso.getfolder(pastaAutomatizacoesBIRPA.Path & "\Automa��es Ellevo\")
    End If
    If Not VerificarPasta(pastaAutomacoesEllevo) Then
        Set pastaProcessoDevolucao = fso.getfolder(pastaGeral.Path & "\Processo Devolu��o\")
    Else
        Set pastaProcessoDevolucao = fso.getfolder(pastaAutomacoesEllevo.Path & "\Processo Devolu��o\")
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
    MsgBox "Favor escolher o caminho no seu Computador a ser descarregado o arquivo baixado do site Transbank. Se n�o possuir a pasta, crie o atalho no Sharepoint e execute novamente a automa��o." & _
        "(Escolha a pasta no seu computador equivalente � pasta do Sharepoint:" & _
            "Documentos > AUTOMATIZA��ES, BIs & RPAs > Automa��es Ellevo > Processo Devolu��o > Arquivos Clientes)"
     
    'Sele��o da pasta para salvar o arquivo
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' O usu�rio selecionou uma pasta
            caminho_pasta = .SelectedItems(1) & "\"
        Else
             'O usu�rio cancelou a sele��o da pasta
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
' funcao complementar da funcao BuscarPasta - apenas verifica se a pasta em quest�o existe
Public Function VerificarPasta(ByVal pasta As Object) As Boolean
    VerificarPasta = True
    If pasta Is Nothing Then
        VerificarPasta = False
    End If
End Function
