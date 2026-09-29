Attribute VB_Name = "categoria"
Enum tipo_categoria
    cat = 0
    subcat = 1
    Marca = 3
End Enum
Type categoria
    codigo As String
    nome As String
    descmax As Currency
    qtdmin As Integer
End Type
Type SubCategoria
    codigo As String
    nome As String
    descmax As Currency
End Type

Function BuscaCategoria(expr As String) As categoria.categoria
    Dim fr As New frmEdiCat
    fr.tag = "x"
    fr.txtNome.Text = expr
    fr.Show 1
    On Error Resume Next
    If fr.tag <> "" Then
        BuscaCategoria.codigo = fr.ado!cod_cat
        BuscaCategoria.nome = fr.ado!nome_cat
        BuscaCategoria.descmax = fr.ado!descmax_cat
        BuscaCategoria.qtdmin = fr.ado!qtddmin_cat
    End If
    Unload fr

End Function
Function BuscaSubCategoria(expr As String) As categoria.SubCategoria
    frmEdiSubCat.tag = "x"
    frmEdiSubCat.txtNome.Text = expr
    frmEdiSubCat.Show 1
    On Error Resume Next
    If frmEdiSubCat.tag <> "" Then
        BuscaSubCategoria.codigo = frmEdiSubCat.ado!cod_sub
        BuscaSubCategoria.nome = frmEdiSubCat.ado!nome_sub
    End If
    Unload frmEdiCat

End Function


Sub CarregaCategoria(obj As TextBox)
    Dim m As categoria.categoria
    m = BuscaCategoria(obj.Text)
    If m.codigo <> "" Then
        obj.Text = m.nome
        obj.tag = m.codigo
    End If
End Sub
Function GravaCategoria(sCod As String, nome_cat As String, Optional tipo_categoria = cat) As Boolean
    Dim strSQL As String
    Dim tabela As String
    If tipo_categoria = Marca Then
        tabela = TAB_MARCAS
    ElseIf tipo_categoria = subcat Then
        tabela = TAB_SUBCAT
    Else
        tabela = TAB_CATEGORIAS
    End If
    
    
    
    
    
    On Error GoTo trata_Erro
    If sCod = "" Then
        sCod = F_BuscaUmCampo(" max(cod_cat)", tabela, "")
    End If
    If sCod = "" Then
        sCod = "001"
    Else
        
        sCod = Format(Int(sCod) + 1, "00#")
    End If
    strSQL = "Insert into " & tabela & "  (cod_cat,nome_cat)values("
    strSQL = strSQL & "'" & sCod & "', "
    strSQL = strSQL & "'" & nome_cat & "')  "
    conn.Execute strSQL
    
    
    GravaCategoria = True
    
    Exit Function
trata_Erro:
    GravaCategoria = False
    
    
End Function
Function AlteraCategoria(codigo As String, nome As String, Optional tipo As tipo_categoria = cat) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    Dim tabela As String
    If tipo = Marca Then
        tabela = TAB_MARCAS
    ElseIf tipo = subcat Then
        tabela = TAB_SUBCAT
    Else
        tabela = TAB_CATEGORIAS
    End If
    
    strSQL = "update " & tabela & "  set "
    
    strSQL = strSQL & " nome_cat='" & nome & "'  "
    strSQL = strSQL & " where cod_cat='" & codigo & "' "
    conn.Execute strSQL
    
    AlteraCategoria = True
trata_Erro:
    AlteraCategoria = False
End Function

