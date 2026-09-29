Attribute VB_Name = "XMLUtil"
Function RetornaValorTagXML(conteudo As String, tag As String, Optional tagFin As String = "") As String
' vazio não encontrado
' tag deve vir <>
    Dim i As Long
    Dim f As Long
    If tagFin = "" Then
        tagFin = "</" & Mid(tag, 2)
    End If
    
    i = InStr(1, conteudo, tag, vbBinaryCompare)
    If i > 0 Then
        f = InStr(i, conteudo, tagFin)
    End If
    If i > 0 And f > 0 Then
        RetornaValorTagXML = Mid(conteudo, i + Len(tag), f - i - Len(tag))
    End If


End Function

Function RetornaValorTagArquivo(conteudo As String, tag As String, tamanho As Integer) As String
' vazio não encontrado
' tag deve vir <>
    Dim i As Long
    Dim f As Long
    
    
    i = InStr(1, conteudo, tag, vbBinaryCompare)
    
    If i > 0 Then
        RetornaValorTagArquivo = Mid(conteudo, i + Len(tag), tamanho)
    End If


End Function

Function RetornaValorExpressaoArquivo(conteudo As String, str_inicio As String, str_fim As String) As String
' vazio não encontrado
' tag deve vir <>
    Dim i As Long
    Dim f As Long
    Dim j As Long
    
    i = InStr(1, conteudo, str_inicio, vbBinaryCompare)
    For j = i + Len(str_inicio) To Len(conteudo)
        If InStr(1, str_fim, Mid(conteudo, j, 1)) > 0 Then
           f = j
           Exit For
        End If
    Next
    
    
    
    If i > 0 Then
        RetornaValorExpressaoArquivo = Trim(Mid(conteudo, i + Len(str_inicio), f - (i + Len(str_inicio))))
    End If


End Function


Function TagXML(tag As String, Optional Parametros As String = "", Optional valor As String = "", Optional fecha As Boolean = False) As String
    Dim xml As String
    tag = XMLUtil.AjustaTagXML(tag)
    If Parametros <> "" Then
        Parametros = " " & Parametros
    End If
    xml = "<" & tag & Parametros & ">" & AjustaTextoXML(valor)
    If fecha Then
        xml = xml + "</" & tag & ">" & CrLf
    ElseIf valor = "" Then
        xml = xml + CrLf
    End If
    TagXML = xml

End Function
Function FechaTagXML(tag As String) As String
    tag = AjustaTagXML(tag)
    FechaTagXML = "</" & tag & ">" & CrLf
End Function
Function RetornoConteudoXML(caminho As String)
    RetornoConteudoXML = LeArquivoExternoUTF8(caminho)
End Function


Function XMLHead() As String
    XMLHead = "<?xml version=" & aspas & "1.0" & aspas & espaco & "encoding=" & aspas & "UTF-8" & aspas & "?>"

End Function

Public Function hex2ascii(ByVal hextext As String) As String
    
For y = 1 To Len(hextext)
    Num = Mid(hextext, y, 2)
    Value = Value & Chr(val("&h" & Num))
    y = y + 1
Next y

hex2ascii = Value
End Function

Function PegaCodigo(Texto As String, parada As String) As String
    On Error GoTo trata_erro
    PegaCodigo = Trim(Left(Texto, InStr(1, Texto, parada) - 1))
Exit Function
trata_erro:
    PegaCodigo = ""
End Function
Function AjustaTextoXML(Texto As String) As String
    Texto = Replace(Texto, "&", "&amp;")
    Texto = Replace(Texto, "<", "&lt;")
    Texto = Replace(Texto, ">", "&gt;")
    AjustaTextoXML = Texto
End Function
Function AjustaTagXML(Texto As String) As String
    Texto = Replace(Texto, " ", "_")
    Texto = Replace(Texto, "*", "_ast_")
    AjustaTagXML = Texto
End Function

Sub ExportaGridToXML(dbgrd As Object, dtVinculado As Object, Parametros As String, Optional caminhoArquivo As String = "")
    
    Dim c As Collection
    Set c = New Collection
    Dim conteudo As String
    Dim xml As String
    If caminhoArquivo = "" Then
        caminhoArquivo = BuscaNomeArquivoExterno("*.xml")
    
    End If
    
    On Error GoTo trata_erro
    dbgrd.Visible = False
    
    
    Dim sPath As String
    Dim strSQL As String
    Dim iQtdColsSelec As Integer
    Dim iCount As Integer
    Dim jCount As Integer
    Dim i As Integer
    Dim par As String
'
    P_Status "Gerando arquivo externo"
    conteudo = ""
    
    
    
    xml = CabXML
    xml = xml + XMLUtil.TagXML("registros", Parametros)
    With dbgrd
        
           
            

        
        
    End With
    'dtVinculado.Refresh
    With dtVinculado
        
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        Do While Not .EOF
            
            par = "seq=" & aspas & .AbsolutePosition & aspas
            
            conteudo = conteudo & XMLUtil.TagXML("registro", par)
            For iCount = 0 To dbgrd.Columns.Count - 1
                        
                If dbgrd.Columns(iCount).Visible = True Then
                    conteudo = conteudo & XMLUtil.TagXML(dbgrd.Columns(iCount).caption, "", dbgrd.Columns(iCount).Text, True)
                
                End If
                
            Next
            conteudo = conteudo & FechaTagXML("registro")
            P_Status .AbsolutePosition & "/" & .RecordCount
            .MoveNext
        Loop
    End With
    
    xml = xml + conteudo
    xml = xml + XMLUtil.FechaTagXML("registros")
    GravaArquivoExternoUTF8 caminhoArquivo, xml
    
    P_Status ""
    If MsgBox("Deseja abrir o arquivo?", vbYesNo + vbQuestion) = vbYes Then
        
        shell "explorer " & caminhoArquivo
    End If
    
    
    
    
    dbgrd.Visible = True
    
    
    
    
    Exit Sub
trata_erro:
    P_MostraErro
    
    On Error Resume Next
    dbgrd.Visible = True
    
    
End Sub

Function AjustaCaminho(caminho As String, Optional reverse As Boolean = False) As String
    Dim c As New Collection
    Dim i As Integer
    Dim s As String
    Set c = F_SeparaTexto(caminho, "\", True)
    For i = 1 To c.Count
        
            s = s + c(i)
        If i < c.Count Or Right(c(i), 1) = "\" Then
            s = s + "\"
        End If
    Next
    
    
    AjustaCaminho = s
End Function

Function UnAjustaCaminho(caminho As String) As String
    Dim c As New Collection
    Dim i As Integer
    Dim s As String
    
    
    Set c = F_SeparaTexto(caminho, "\", False)
    For i = 1 To c.Count
        
            s = s + c(i)
        If i < c.Count Then
            s = s + "\"
        End If
    Next
    
    If Left(caminho, 2) = "\\" Then
        s = "\\" & s
    End If
    UnAjustaCaminho = s
End Function



Function EncriptarToken(token As String, chave As String) As String
    Dim i As Long
    Dim resultado As String
    Dim tChar As Integer
    Dim kChar As Integer
    
    resultado = ""
    For i = 1 To Len(token)
        tChar = Asc(Mid(token, i, 1))
        kChar = Asc(Mid(chave, ((i - 1) Mod Len(chave)) + 1, 1))
        resultado = resultado & Right("0" & Hex(tChar Xor kChar), 2)
    Next i
    EncriptarToken = resultado
End Function
Function DesencriptarToken(cripto As String, chave As String) As String
    Dim i As Long
    Dim resultado As String
    Dim hexByte As String
    Dim tChar As Integer
    Dim kChar As Integer
    
    resultado = ""
    For i = 1 To Len(cripto) Step 2
        hexByte = Mid(cripto, i, 2)
        tChar = CInt("&H" & hexByte)
        kChar = Asc(Mid(chave, (((i \ 2)) Mod Len(chave)) + 1, 1))
        resultado = resultado & Chr(tChar Xor kChar)
    Next i
    DesencriptarToken = resultado
End Function

