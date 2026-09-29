Attribute VB_Name = "fotos"
Option Explicit
Dim Arquivo As Integer
Dim iCont As Integer
Dim TamArquivo As Long
Dim Fragmento As Long
Dim Chunk() As Byte
Dim Chunks As Integer
Const TamChunk As Integer = 16384


' 32-bit EZTWAIN functions for Visual Basic 5.0 para o scaner
Declare Function TWAIN_AcquireToClipboard Lib "EZTW32.DLL" (ByVal hwndApp&, ByVal wPixTypes&) As Long
Declare Function TWAIN_SelectImageSource Lib "EZTW32.DLL" (ByVal hwndApp&) As Long


'rsBLOB = Campo BLOB do Recordset
Public Function CarregaBLOB(rsBLOB As ADODB.Field, nFile As String) As String
    On Error GoTo trata_erro
    CarregaBLOB = ""
    Arquivo = 1
    If Not IsNull(rsBLOB) Then
        Open nFile For Binary Access Write As Arquivo
        ReDim BufferF(rsBLOB.ActualSize)
        Chunk() = rsBLOB.GetChunk(rsBLOB.ActualSize)
        Put Arquivo, , Chunk()
        Close Arquivo
        CarregaBLOB = nFile
    End If
    
trata_erro:
    Close Arquivo
End Function
'sArquivo = Nome do Arquivo com a imagem
'rsBLOB = Campo do Recordset
Public Function GravaBLOB(sArquivo As String, rsBLOB As ADODB.Field)
    If sArquivo > "" Then
         Arquivo = 1
         Open sArquivo For Binary Access Read As Arquivo
         TamArquivo = LOF(Arquivo)    ' Tamanho do Arquivo
         If TamArquivo > 0 Then
            ReDim Chunk(TamArquivo)
            Get Arquivo, , Chunk()
            rsBLOB.AppendChunk Chunk()
        End If
        Close Arquivo
    End If
End Function
Public Sub GravarFigura()
Dim cn As New ADODB.Connection 'Var de Conexao com o banco de dados
Dim rs As ADODB.Recordset 'Var com os resultados de uma consulta
        cn.Open 'COLOQUE AQUI OS PARÂMETROS DE ABERTURA DO BANCO
        Set rs = cn.Execute("SELECT imagem FROM BANCO where codigo = 1")
        GravaBLOB "C:\Temp\Figura.jpg", rs("imagem") 'Coloca a nova figura no campo do recordset
        rs.UpdateBatch 'realiza o update no Banco de Dados
        cn.Close 'Fecha Conexao
        Set cn = Nothing
        Set rs = Nothing
End Sub
Public Sub LerFigura(sNomeFigura As String)
Dim cn As New ADODB.Connection 'Var de Conexao com o banco de dados
Dim rs As ADODB.Recordset 'Var com os resultados de uma consulta
        cn.Open 'COLOQUE AQUI OS PARÂMETROS DE ABERTURA DO BANCO
        Set rs = cn.Execute("SELECT imagem FROM BANCO")
        'Image.Picture = LoadPicture(CarregaBLOB(rs("imagem"), "file"))
        cn.Close 'Fecha Conexao
        Set cn = Nothing
        Set rs = Nothing
End Sub


