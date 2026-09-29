Attribute VB_Name = "fotos"
Option Explicit
Dim Arquivo As Integer
Dim iCont As Integer
Dim TamArquivo As Long
Dim Fragmento As Long
Dim Chunk() As Byte
Dim Chunks As Integer
Const TamChunk As Integer = 16384
Public sGravaFoto_Temp As String
Public sFoto_Temp As String
' 32-bit EZTWAIN functions for Visual Basic 5.0 para o scaner
Declare Function TWAIN_AcquireToClipboard Lib "EZTW32.DLL" (ByVal hwndApp&, ByVal wPixTypes&) As Long
Declare Function TWAIN_SelectImageSource Lib "EZTW32.DLL" (ByVal hwndApp&) As Long


'rsBLOB = Campo BLOB do Recordset
Public Function CarregaBLOB(rsBLOB As adodb.Field, nFile As String) As String
    On Error GoTo trata_Erro
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
    
trata_Erro:
    Close Arquivo
End Function
'sArquivo = Nome do Arquivo com a imagem
'rsBLOB = Campo do Recordset
Public Function GravaBLOB(sArquivo As String, rsBLOB As adodb.Field)
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
Dim cn As New adodb.Connection 'Var de Conexao com o banco de dados
Dim rs As adodb.Recordset 'Var com os resultados de uma consulta
        cn.Open 'COLOQUE AQUI OS PARÂMETROS DE ABERTURA DO BANCO
        Set rs = cn.Execute("SELECT imagem FROM BANCO where codigo = 1")
        GravaBLOB "C:\Temp\Figura.jpg", rs("imagem") 'Coloca a nova figura no campo do recordset
        rs.UpdateBatch 'realiza o update no Banco de Dados
        cn.Close 'Fecha Conexao
        Set cn = Nothing
        Set rs = Nothing
End Sub
Public Sub LerFigura(sNomeFigura As String)
Dim cn As New adodb.Connection 'Var de Conexao com o banco de dados
Dim rs As adodb.Recordset 'Var com os resultados de uma consulta
        cn.Open 'COLOQUE AQUI OS PARÂMETROS DE ABERTURA DO BANCO
        Set rs = cn.Execute("SELECT imagem FROM BANCO")
        'Image.Picture = LoadPicture(CarregaBLOB(rs("imagem"), "file"))
        cn.Close 'Fecha Conexao
        Set cn = Nothing
        Set rs = Nothing
End Sub


Function F_RetornaSeqDoArquivo(sArq As String)
    Dim sAux As String
    Dim sBuf As String
    Dim i As Integer
    Dim iCount As Integer
    For iCount = Len(sArq) To 1 Step -1
        
        sBuf = Mid(sArq, iCount, 1)
        If sBuf = "\" Then
            Exit For
        End If
        sAux = sBuf & sAux
    Next
    F_RetornaSeqDoArquivo = F_SoNumero(sAux)
End Function

Function F_GravaFoto(sTab As String, sCampo As String, cod As String, sCaminho As String, Optional tipo As String = "") As Boolean
    Dim strSQL As String
    Dim sCod As String
    Dim sSeq As String
    On Error GoTo trata_Erro
    sSeq = MaxKeySQL(sTab, "seq_fot", "", "")
    strSQL = "insert into " & sTab & "(seq_Fot,data_Fot," & sCampo & " ,cod_usu,foto_Fot)values("
    strSQL = strSQL & "'" & sSeq & "', "
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'" & cod & "',"
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "load_file('" & F_AjustaTexto(sCaminho) & "'))"
    Conn.Execute strSQL
    F_GravaFoto = True
    Exit Function
trata_Erro:
F_GravaFoto = False
    MsgBox Err.Description, vbCritical
End Function

