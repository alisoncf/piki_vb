Attribute VB_Name = "QRCodeUtil"
Option Explicit
Private Declare Sub FastQRCode Lib "c:\tmp\QRCodeLib.dll" (ByVal Texto As String, ByVal FileName As String)
Private Declare Function QRCodeLibVer Lib "c:\tmp\QRCodeLib.dll" () As String




Public Function ImprimeQRCodeArquivo(conteudo As String, caminhoBMP As String) As String
'    On Error GoTo trata_Erro

    FastQRCode conteudo, caminhoBMP
    ImprimeQRCodeArquivo = caminhoBMP
trata_Erro:
    Exit Function

End Function

Public Function ImprimeQRCodePictureBox(pb As PictureBox, conteudo As String, Optional caminhoBMP As String = "") As Boolean
'    On Error GoTo trata_Erro
    
    If caminhoBMP = "" Then
        caminhoBMP = App.Path & "\QRcode.bmp"
    End If

    FastQRCode conteudo, caminhoBMP
    pb.PaintPicture LoadPicture(caminhoBMP), 0, 0
    
    
    ImprimeQRCodePictureBox = True
trata_Erro:
    Exit Function

End Function

