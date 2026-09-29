Attribute VB_Name = "visao"
Public strVisao As String

Global strVisaoCaixa_usu As String
Global strVisaoEstoque_usu As String
Global strVisaoFinanceiro_usu As String
Global strVisaoAtacado_usu As String
Global strVisaoGeral_usu As String

Public Enum tipo_visao
    tCAIXA = 0
    
    tEstoque = 1
    tFinanceiro = 2
    tAtacado = 3
    tGeral = 4
    
End Enum
    
Sub MudaVisao(cod As tipo_visao)
    If cod = tCAIXA Then
        mdiFrm.muda_visao tCAIXA
    End If
End Sub
