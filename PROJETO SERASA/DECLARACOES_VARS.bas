Attribute VB_Name = "DECLARACOES_VARS"
Public Sub declaracao_variaveis()


    Set aba_fbl5h_base_geral = ThisWorkbook.Sheets("FBL5H - Base Geral")
    Set tabela_aba_fbl5h_base_geral = aba_fbl5h_base_geral.ListObjects("Tabela_FBL5H_Base_Geral")
    
    Set aba_fbl5h_base_compensados_serasa = ThisWorkbook.Sheets("FBL5H - Base Compensados SERASA")
    Set tabela_aba_fbl5h_base_compensados_serasa = aba_fbl5h_base_compensados_serasa.ListObjects("Tabela_FBL5H_Base_Compensados_SERASA")
    
    Set aba_base_historica = ThisWorkbook.Sheets("Base Histórica")
    Set tabela_aba_base_historica = aba_base_historica.ListObjects("Tabela_Base_Histórica")
    
    Set aba_numero_remessa = ThisWorkbook.Sheets("Nº Remessa")
    
End Sub
