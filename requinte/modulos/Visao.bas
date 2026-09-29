Attribute VB_Name = "Visao"
Public strVisao As String
Public Enum tipo_visao
    tCAIXA = 0
    tConsignado = 1
    tCrediario = 2
    tEstoque = 3
    TVendas = 4
    tPacote = 5
    tFinanceiro = 6
    tGeral = 7
    
End Enum
    
Sub MudaVisao(cod As tipo_visao)
    If cod = tCAIXA Then
        mdiFrm.muda_visao tCAIXA
    End If
End Sub
