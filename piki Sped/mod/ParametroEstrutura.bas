Attribute VB_Name = "ParametroEstrutura"
Public ListParametros() As parametro_estrutura

Public Type parametro_estrutura
    id As Integer
    regra As String
    valor As String
    

End Type


Function BuscaParametroEstrutura(IdOuRegra As String) As parametro_estrutura
    Dim strSQL As String
    Dim ado As New ADODB.Recordset
    Dim p As parametro_estrutura
    On Error GoTo trata_Erro
    strSQL = "select * from parametro_estrutura where id ='" & IdOuRegra & "' "
    strSQL = strSQL & " or regra='" & IdOuRegra & "' "
    abreObjetoAdo ado, strSQL
    With ado
        If .RecordCount > 0 Then
            BuscaParametroEstrutura.id = !id
            BuscaParametroEstrutura.regra = !regra
            BuscaParametroEstrutura.valor = !valor
        End If
        .Close
    End With
    
trata_Erro:


End Function


Public Sub CarregarParametros()

    Dim strSQL As String
    Dim ado As New ADODB.Recordset
    Dim p As parametro_estrutura
    On Error GoTo trata_Erro
    strSQL = "select * from parametro_estrutura "
    
    abreObjetoAdo ado, strSQL
    ReDim ListParametros(ado.RecordCount) As parametro_estrutura
    
    With ado
        Do While Not .EOF
            p.id = !id
            p.regra = !regra
            p.valor = !valor
            
            If LCase(!regra) = "usaconfecção" Then
                bUsaConfeccao = !valor = "S"
            End If
            If LCase(!regra) = "usaveiculo" Then
                bUsaVeiculo = !valor = "S"
            End If
            If LCase(!regra) = "usaentrega" Then
                bUsaEntrega = !valor = "S"
            End If
            If LCase(!regra) = "controlaestoque" Then
                bControlaEstoque = !valor = "S"
            End If
            If LCase(!regra) = "usacombustiveis" Then
                bUsaCombustiveis = !valor = "S"
            End If
            
            ListParametros(ado.AbsolutePosition - 1) = p
            .MoveNext
        Loop
        
    End With
Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub


Function BuscaParametroLista(IdOuRegra As String) As String
    Dim i As Integer
    Dim f As Integer
    f = UBound(ListParametros)
    
    For i = 0 To f
        If LCase(ListParametros(i).id) = LCase(IdOuRegra) Or LCase(ListParametros(i).regra) = LCase(IdOuRegra) Then
            BuscaParametroLista = ListParametros(i).valor
            Exit Function
        End If
    Next
End Function
