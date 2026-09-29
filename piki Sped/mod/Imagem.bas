Attribute VB_Name = "Imagem"
Public caminhoImagem As String

Dim xPoint As Single
Dim yPoint As Single
Dim xSolta As Single
Dim ySolta As Single
Dim prop As Single

Dim bActive As Boolean
Dim listaImagens As New Collection
Dim Indice As Integer
Dim Quantidade As Integer


Type imagem
    id As Integer
    caminho As String
    cod_pro As String
    cod_usu As String
End Type

 Sub buscaCaminhoImagem()
    caminhoImagem = F_BuscaUmCampo("caminho_imagem", "tab_parametros", "")
End Sub
Sub AdicionarImagem(cod_pro As String, caminho As String)
    Dim strSQL As String
    Dim copia As String
    On Error GoTo trata_erro
    
    copia = CriarCopiaImagem(caminho, cod_pro)
    
    strSQL = "insert into imagem_produto(cod_Pro,caminho,cod_usu)values("
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & "'" & AjustaNomeArquivo(caminho) & "',"
    strSQL = strSQL & "'" & strCod_Usu & "')"
    
    conn.Execute strSQL
    
    Exit Sub
trata_erro:
    MsgBox Err.Description
End Sub

Sub AlterarImagem(id As String, cod_pro As String, caminho As String)
    Dim strSQL As String
    On Error GoTo trata_erro
    strSQL = "update imagem set  "
    strSQL = strSQL & "cod_pro='" & cod_pro & "',"
    
    strSQL = strSQL & "caminho='" & caminho & "' "
    strSQL = strSQL & " where id = " & id
    
    conn.Execute strSQL
    Exit Sub
trata_erro:
    MsgBox Err.Description
    
End Sub
Sub ExcluirImagem(id As String)
    Dim strSQL As String
    On Error GoTo trata_erro
    strSQL = "delete from imagem_produto where id=" & id
    conn.Execute strSQL
    Exit Sub
trata_erro:
    MsgBox Err.Description
End Sub
 Function CriarCopiaImagem(img As String, cod_pro As String) As String

    On Error GoTo trata_erro
    Dim s As String
    s = cod_pro & "_" & F_SoNumero(Now) & ".jpg"

    CriarCopiaImagem = s
    
    
    
    'If Dir(CaminhoImagens, vbDirectory) = "" Then
    '    MkDir CaminhoImagens
    'End If
    
    FileCopy img, CaminhoImagens & "\" & s
    CriarCopiaImagem = CaminhoImagens & "\" & s
    Exit Function
    
trata_erro:
    MsgBox "erro ao criar cópia da imagem em " & s & "." & Err.Description
    
End Function

Sub AbreImagens(cod_pro As String)
    frmImagens.tag = cod_pro
    frmImagens.txtProduto.Text = cod_pro
    frmImagens.txtProduto.tag = cod_pro
    
    frmImagens.Show 1
    
    
    
End Sub
Sub mostraMiniatura(cod_pro As String, Optional Form As Object, Optional x As Single = -1, Optional y As Single = -1, Optional wid As Single, Optional hei As Single)
    frmImagensMiniatura.ScaleMode = 1
    frmImagensMiniatura.tag = cod_pro
    frmImagensMiniatura.Show 0, Form
    If wid > 0 Then
        frmImagensMiniatura.width = wid
    End If
    If hei > 0 Then
        frmImagensMiniatura.height = hei
    End If
    If x > -1 Then
        frmImagensMiniatura.Left = x
    End If
    If y > -1 Then
        frmImagensMiniatura.Top = x
    End If
    
    
End Sub
Sub escondeMiniatura()
    Unload frmImagensMiniatura
End Sub








Sub AjusteImagem(pcImg As PictureBox, pcQuadro As PictureBox)
    If pcImg.image = 0 Then
    
    Else
        resizeImage pcImg, pcQuadro
    End If
End Sub
Sub resizeImage(img As Object, obj As PictureBox)
    Dim i As Integer
    Dim h As Single
    Dim w As Single
    Dim p As Single
    Dim horizontal As Single
    Dim proporcao As Single
    pcImg.Left = 0
    pcImg.Top = 0
    h = img.height
    w = img.width
    p = h / 10
    
    
    If w > obj.ScaleWidth Then
    
        For i = w To obj.width Step -p
            horizontal = i
            
            
        Next
        proporcao = i / w
     ElseIf h > obj.ScaleHeight Then
        For i = h To obj.height Step -p
            horizontal = i
            
            
        Next
        proporcao = i / w
    End If
    
    'MostraImagemMiniatura listaImagens(1), img.width * proporcao, img.height * proporcao
    prop = proporcao
End Sub




Sub CarregaImagem(id As String, pcImg As PictureBox)
    Set listaImagens = F_SeparaTexto(id, ",", False)
    Quantidade = listaImagens.Count
    
    If id <> "" Then
        'MostraImagemMiniatura listaImagens(1), pcImg
    End If
End Sub








Public Sub abreImagem()
    On Error Resume Next
    Shell "explorer " & txtCaminhoImg.Text
End Sub



Private Sub pcimg_MouseUp(pcImg As PictureBox, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        xSolta = (x - xPoint)
        ySolta = (y - yPoint)
        
        pcImg.Left = pcImg.Left + xSolta
        pcImg.Top = pcImg.Top + ySolta
        
    End If
End Sub

Public Function ImagemMiniatura(img As String, pcImg As PictureBox, Optional w As Single, Optional h As Single) As PictureBox
    Dim s As String
    On Error Resume Next
    pcImg.ScaleMode = 1
    
    pcImg.picture = Nothing
    pcImg.AutoSize = True

    
    If w > 0 And h > 0 Then
         pcImg.PaintPicture LoadPicture(img), 0, 0, w, h
    Else
        prop = 1
        pcImg.Left = 0
        pcImg.Top = 0
        pcImg.picture = LoadPicture(img)
    End If
  
    pcImg.Refresh
    
    txtCaminhoImg.Text = img
    
   ' Set pcMiniatura.Picture = LoadPicture(s)
   ImagemMiniatura = pcImg
    
End Function





Sub carregaImagemFit(pic As PictureBox, container As PictureBox, arquivo As String)
    container.ScaleMode = 1
    pic.ScaleMode = 1
    pic.AutoRedraw = True
    pic.AutoSize = True
    pic.picture = LoadPicture(arquivo)
    wid = pic.image.width
    hei = pic.image.height
    w = container.width
    h = container.height
    Do Until hei < h And wid < w
        hei = hei * 0.995
        wid = wid * 0.995
    Loop
    container.picture = Nothing
    container.PaintPicture LoadPicture(arquivo), 0, 0, pic.ScaleX(wid, container.ScaleMode, vbTwips), pic.ScaleY(hei, container.ScaleMode, vbTwips)
    container.AutoSize = True
    container.Refresh
End Sub

