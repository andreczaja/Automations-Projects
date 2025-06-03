Attribute VB_Name = "PASSO_2_PROCESS_REEMBOLSO_SAP"
Public doc_f65, status_SBWP As String
Public array_docs_F65() As Variant
Public linha_fim_aba_reembolsos_pendentes As Long
Public data_agrupado_pagamento As String

Sub Processamento_Reembolso_SAP()

    ' Obt�m a data do agrupado de pagamento da caixa de texto no formul�rio "Form_SAP"
    data_agrupado_pagamento = Form_SAP.txt_box_data_agrupado_pgto_SAP
    ' Se a data n�o foi encontrada ou est� vazia
    If data_agrupado_pagamento = ".." Or data_agrupado_pagamento = "" Then
        ' Exibe uma caixa de entrada para o usu�rio digitar a data no formato "DD/MM/AAAA"
        data_agrupado_pagamento = InputBox("A data do agrupado de pagamento n�o foi encontrada. Por favor digite-a abaixo no formato 'DD/MM/AAAA'")
        ' Grava a data digitada na c�lula BC1 da planilha "aba_reembolsos_aprovados"
        aba_reembolsos_aprovados.Range("BC1").Value = data_agrupado_pagamento
    End If

    ' Verifica se a conta n�o est� bloqueada para a transa��o "F-65" (Estorno de Pagamento)
    If Not VerificarContaBloqueada("F-65") Then

        ' Obt�m a pr�xima linha vazia na coluna A da planilha "aba_reembolsos_pendentes"
        linha_fim_aba_reembolsos_pendentes = aba_reembolsos_pendentes.Range("A1048576").End(xlUp).Offset(1, 0).Row

        ' Preenche o campo de valor na terceira sess�o do SAP com o valor absoluto da soma dos cr�ditos de devolu��o
        session_3.findById("wnd[0]/usr/txtBSEG-WRBTR").text = Abs(soma_cred_dev)
        If InStr(1, session_3.findById("wnd[0]/usr/txtBSEG-WRBTR").text, ".") Then
            session_3.findById("wnd[0]/usr/txtBSEG-WRBTR").text = Replace(session_3.findById("wnd[0]/usr/txtBSEG-WRBTR").text, ".", ",")
        End If
        ' Preenche o campo de atribui��o com "REEMB AUT" seguido da data do agrupado de pagamento formatada
        session_3.findById("wnd[0]/usr/txtBSEG-ZUONR").text = "REEMB AUTOMACAO"
        ' Preenche o campo de texto do item com a descri��o do processo
        session_3.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = "Processo autom�tico de reembolso de devolu��o"
        ' Simula o clique no bot�o "Pr�ximo Item" (ou similar)
        session_3.findById("wnd[0]/tbar[1]/btn[7]").press
        ' Preenche o campo de data base para condi��es de pagamento com a data atual no formato SAP
        session_3.findById("wnd[0]/usr/ctxtBSEG-FDTAG").text = VBA.Format(VBA.Date, tipo_data_sap)
        ' Preenche o campo de chave de lan�amento especial com "1D" (adiantamento ao cliente)
        session_3.findById("wnd[0]/usr/ctxtRF05V-NEWBS").text = "1D"
        ' Preenche o campo de conta de contrapartida com o payer associado � OC
        session_3.findById("wnd[0]/usr/ctxtRF05V-NEWKO").text = payer_associado_OC
        ' Simula a tecla Enter
        session_3.findById("wnd[0]").sendVKey 0
        ' Preenche novamente o campo de valor
        session_3.findById("wnd[0]/usr/txtBSEG-WRBTR").text = Abs(soma_cred_dev)
        If InStr(1, session_3.findById("wnd[0]/usr/txtBSEG-WRBTR").text, ".") Then
            session_3.findById("wnd[0]/usr/txtBSEG-WRBTR").text = Replace(session_3.findById("wnd[0]/usr/txtBSEG-WRBTR").text, ".", ",")
        End If
        
        ' Preenche o campo de atribui��o com "AUTOMACAO DEV"
        session_3.findById("wnd[0]/usr/txtBSEG-ZUONR").text = "AUTOMACAO DEV"
        session_3.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = "Processo autom�tico de reembolso de devolu��o"
        ' Preenche o campo de m�todo de pagamento com "T" (transfer�ncia banc�ria)
        session_3.findById("wnd[0]/usr/ctxtBSEG-ZLSCH").text = "T"
        ' Simula o clique no bot�o "Pr�ximo Item"
        session_3.findById("wnd[0]/tbar[1]/btn[7]").press
        ' Preenche o campo de chave de refer�ncia 2 com "AUTOMACAO"
        session_3.findById("wnd[0]/usr/txtBSEG-XREF2").text = "AUTOMACAO"
        ' Preenche o campo de data base para condi��es de pagamento novamente
        session_3.findById("wnd[0]/usr/ctxtBSEG-FDTAG").text = VBA.Format(VBA.Date, tipo_data_sap)
        ' Simula o clique no bot�o "Pr�ximo Item"
        session_3.findById("wnd[0]/tbar[1]/btn[7]").press
        ' Define o foco no campo de valor do desconto
        session_3.findById("wnd[0]/usr/txtBSEG-WSKTO").SetFocus
        ' Seleciona a op��o de menu "Documento -> Simular"
        session_3.findById("wnd[0]/mbar/menu[0]/menu[4]").Select

        ' Extrai o n�mero do documento F-65 da barra de status
        doc_f65 = Mid(session_3.findById("wnd[0]/sbar").text, 11, 10)
        ' Preenche as colunas da planilha "aba_reembolsos_pendentes" com as informa��es do reembolso
        aba_reembolsos_pendentes.Range("A" & linha_fim_aba_reembolsos_pendentes).Value = doc_f65
        aba_reembolsos_pendentes.Range("B" & linha_fim_aba_reembolsos_pendentes).Value = chamado
        aba_reembolsos_pendentes.Range("C" & linha_fim_aba_reembolsos_pendentes).Value = payer_associado_OC
        aba_reembolsos_pendentes.Range("D" & linha_fim_aba_reembolsos_pendentes).Value = Date
        aba_reembolsos_pendentes.Range("E" & linha_fim_aba_reembolsos_pendentes).Value = "N�o Solicitada Aprova��o"
        aba_reembolsos_pendentes.Range("F" & linha_fim_aba_reembolsos_pendentes).Value = Abs(soma_cred_dev)
        aba_reembolsos_pendentes.Range("G" & linha_fim_aba_reembolsos_pendentes).Value = qtde_NFD_OC_chamado
        aba_reembolsos_pendentes.Range("H" & linha_fim_aba_reembolsos_pendentes).Value = Replace(UCase(VBA.Environ("USERPROFILE")), "C:\USERS\", "")
        ' Chama a sub-rotina para criar o arquivo anexo de reembolso
        Call CriarArquivoAnexoReembolso
        ' Chama a sub-rotina para verificar linhas na SBWP (SAP Business Workplace) para aprova��o unit�ria
        Call VerificarLinhasSBWP(session_3, "UNITARIA")
        ' Simula F5 e altera a atribui��o para "AG PROCESS SBWP" (Aguardando Processamento SBWP)
        session.findById("wnd[0]").sendVKey 5
        Call AlterarAtribuicao(session, "AG PROCESS SBWP")
    Else
        ' Se a conta estiver bloqueada, simula F5 e altera a atribui��o para "CTA BLOQUEADA"
        session.findById("wnd[0]").sendVKey 5
        Call AlterarAtribuicao(session, "CTA BLOQUEADA")
    End If

    
    

End Sub
