Attribute VB_Name = "Padrao"
Type Padrao1
    codigo As String
    nome As String
End Type


Function ListaPadrao1(descricao As String, tabela As String, campo1 As String, campo2 As String, Optional campo3 As String = "") As Padrao1
    On Error Resume Next
    frmEdiPadrao1.tabela = tabela
    frmEdiPadrao1.campo1 = campo1
    frmEdiPadrao1.campo2 = campo2
    frmEdiPadrao1.campo3 = campo3
    frmEdiPadrao1.txtBuscar.text = descricao
    frmEdiPadrao1.caption = "Lista de " & tabela
    frmEdiPadrao1.tag = "x"
    frmEdiPadrao1.Show 1
    
    If frmEdiPadrao1.tag <> "" Then
        With frmEdiPadrao1.ado
            ListaPadrao1.codigo = .Fields(0).Value
            ListaPadrao1.nome = .Fields(1).Value
            
        End With
    End If
    Unload frmEdiPadrao1
End Function

Function SelectUnidade(descricao As String) As Padrao1
    SelectUnidade = ListaPadrao1(descricao, "sped_unidades", "unidade", "descricao")
End Function
Sub EditaUnidade()
    frmEdiPadrao1.tabela = "sped_unidades"
    frmEdiPadrao1.campo1 = "unidade"
    frmEdiPadrao1.campo2 = "descricao"
    frmEdiPadrao1.txtBuscar.text = descricao
    frmEdiPadrao1.caption = "Lista de " & tabela
    frmEdiPadrao1.Show 1
    Unload frmEdiPadrao1
    
End Sub

