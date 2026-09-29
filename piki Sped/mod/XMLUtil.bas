Attribute VB_Name = "XMLUtil"
Function RetornaValorTagXML(conteudo As String, tag As String, Optional tagFin As String = "") As String
' vazio não encontrado
' tag deve vir <>
    Dim i As Integer
    Dim f As Integer
    If tagFin = "" Then
        tagFin = "</" & Mid(tag, 2)
    End If
    
    i = InStr(1, conteudo, tag)
    If i > 0 Then
        f = InStr(i, conteudo, tagFin)
    End If
    If i > 0 And f > 0 Then
        RetornaValorTagXML = Mid(conteudo, i + Len(tag), f - i - Len(tag))
    End If


End Function
