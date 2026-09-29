Attribute VB_Name = "Imposto_NFe"

Public Enum PIS_CST
    Outras_Operacoes_de_Saida = 49
    operacao_com_Direito_a_credito__Vinculada_Exclusivamente_a_Receita_Tributada_no_Mercado_Interno = 50
    operacao_com_Direito_a_credito__Vinculada_Exclusivamente_a_Receita_Nao_Tributada_no_Mercado_Interno = 51
    operacao_com_Direito_a_credito_–_Vinculada_Exclusivamente_a_Receita_de_Exportacao = 52
    operacao_com_Direito_a_credito__Vinculada_a_Receitas_Tributadas_e_NaoTributadas_no_Mercado_Interno = 53
    operacao_com_Direito_a_credito__Vinculada_a_Receitas_Tributadas_no_Mercado_Interno_e_de_Exportacao = 54
    operacao_com_Direito_a_credito__Vinculada_a_Receitas_NaoTributadas_no_Mercado_Interno_e_de_Exportacao = 55
    operacao_com_Direito_a_credito__Vinculada_a_Receitas_Tributadas_e_NaoTributadas_no_Mercado_InternoE_de_Exportacao = 56
    credito_Presumido__operacao_de_Aquisicao_Vinculada_Exclusivamente_a_Receita_Tributada_no_Mercado_Interno = 60
    credito_Presumido__operacao_de_Aquisicao_Vinculada_Exclusivamente_a_Receita_NaoTributada_no_Mercado_Interno = 61
    credito_Presumido__operacao_de_Aquisicao_Vinculada_Exclusivamente_a_Receita_de_Exportacao = 62
End Enum
Public Enum TipoICMS
    
    ICMS00 = 0
    ICMS10 = 10
    ICMS20 = 20
    ICMS30 = 30
    ICMS40 = 40
    ICMSSN500 = 500
    ICMS51 = 51
    ICMS60 = 60
    ICMS70 = 70
    ICMS90 = 90
    ICMSPART = -1
    ICMSSN101 = 101
    ICMSSN102 = 102
    ICMSSN201 = 201
    ICMSSN202 = 202
    ICMSSN900 = 900
    ICMSST = -2
    
End Enum
Public Type TImposto
    vTotTrib As String
    TipoICMS As TipoICMS
    TipoPis As String
    TipoCofins As String
    
    i00 As ICMS_NFe.ICMS00
    i10 As ICMS_NFe.ICMS10
    i20 As ICMS_NFe.ICMS20
    i30 As ICMS_NFe.ICMS30
    i40 As ICMS_NFe.ICMS40
    i500 As ICMS_NFe.ICMS500
    i51 As ICMS_NFe.ICMS51
    i60 As ICMS_NFe.ICMS60
    i70 As ICMS_NFe.ICMS70
    i90 As ICMS_NFe.ICMS90
    iPart As ICMS_NFe.ICMSPART
    iSN101 As ICMS_NFe.ICMSSN101
    iSN102 As ICMS_NFe.ICMSSN102
    iSN201 As ICMS_NFe.ICMSSN201
    iSN202 As ICMS_NFe.ICMSSN202
    iSN900 As ICMS_NFe.ICMSSN900
    iST As ICMS_NFe.ICMSST
    
End Type
