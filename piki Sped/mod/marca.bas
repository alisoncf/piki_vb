Attribute VB_Name = "marca"
Type Marca
    codigo As String
    nome As String
End Type
Function buscaMarca(expr As String) As Marca.Marca
    frmEdiMar.tag = "x"
    frmEdiMar.txtNome.Text = expr
    frmEdiMar.Show 1
    On Error Resume Next
    If frmEdiMar.tag <> "" Then
        buscaMarca.codigo = frmEdiMar.ado!cod_mar
        buscaMarca.nome = frmEdiMar.ado!nome_mar
    End If
    Unload frmEdiMar

End Function
Sub CarregaMarca(obj As TextBox)
    Dim m As Marca.Marca
    m = buscaMarca(obj.Text)
    If m.codigo <> "" Then
        obj.Text = m.nome
        obj.tag = m.codigo
    End If
End Sub
