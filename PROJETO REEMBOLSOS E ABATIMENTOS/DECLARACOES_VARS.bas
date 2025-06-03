Attribute VB_Name = "DECLARACOES_VARS"
Public aba_plan_distribuicao, aba_fbl5n_credito_devolucao, aba_fbl5n_AR, aba_dados_bancarios, aba_titulos_a_abater, _
    aba_reembolsos_pendentes, aba_reembolsos_aprovados, aba_base_historica, aba_modelos_de_emails, aba_emails_analistas, aba_home As Worksheet
    
Public tabela_aba_plan_distribuicao, tabela_aba_fbl5n_credito_devolucao, tabela_aba_fbl5n_AR, tabela_titulos_a_abater, _
    tabela_aba_reembolsos_pendentes, tabela_reembolsos_aprovados, tabela_aba_base_historica, tabela_aba_emails_analistas As ListObject

Public Sub declaracao_vars()

    ' declaracao de todas as abas e respectivas tabelas da planilha necess�rias para a automa��o

    Set aba_plan_distribuicao = ThisWorkbook.Sheets("Plan Distribui��o")
    Set tabela_aba_plan_distribuicao = aba_plan_distribuicao.ListObjects("Plan_Distribui��o")
    
    Set aba_fbl5n_credito_devolucao = ThisWorkbook.Sheets("FBL5N Cr�dito Devolu��o")
    Set tabela_aba_fbl5n_credito_devolucao = aba_fbl5n_credito_devolucao.ListObjects("FBL5N_Cr�ditos_Devolu��o")
    
    Set aba_fbl5n_AR = ThisWorkbook.Sheets("FBL5N AR")
    Set tabela_aba_fbl5n_AR = aba_fbl5n_AR.ListObjects("FBL5N_AR")
    
    Set aba_dados_bancarios = ThisWorkbook.Sheets("Check Dados Bancarios")
 
    Set aba_titulos_a_abater = ThisWorkbook.Sheets("T�tulos a Abater")
    Set tabela_titulos_a_abater = aba_titulos_a_abater.ListObjects("Tabela_Titulos_a_Abater")
    
    Set aba_reembolsos_pendentes = ThisWorkbook.Sheets("Reembolsos Pendentes")
    Set tabela_aba_reembolsos_pendentes = aba_reembolsos_pendentes.ListObjects("Tabela_Reembolsos_Pendentes")
    
    Set aba_reembolsos_aprovados = ThisWorkbook.Sheets("Reembolsos Aprovados")
    Set tabela_reembolsos_aprovados = aba_reembolsos_aprovados.ListObjects("Tabela_Reembolsos_Aprovados")
    
    Set aba_base_historica = ThisWorkbook.Sheets("Base Hist�rica")
    Set tabela_aba_base_historica = aba_base_historica.ListObjects("Tabela_Base_Hist�rica")
    
    Set aba_modelos_de_emails = ThisWorkbook.Sheets("Modelos de Emails")
    
    Set aba_home = ThisWorkbook.Sheets("Home")
    
End Sub
