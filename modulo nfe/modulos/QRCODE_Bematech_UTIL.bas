Attribute VB_Name = "QRCODE_Bematech_UTIL"
Public Declare Function ImprimeCodigoQRCODE Lib "mp2032.dll" _
(ByVal errorCorrectionLevel As Integer, ByVal moduleSize As Integer, ByVal codeType As Integer, _
    ByVal QRCodeVersion As Integer, ByVal encodingModes As Integer, ByVal codeQr As String) As Integer


Sub testeQR()
    errorCorrectionLevel = 1
moduleSize = 10
codeType = 0
QRCodeVersion = 10
encodingModes = 1
codeQr = "123ABC"
iRetorno = ImprimeCodigoQRCODE(errorCorrectionLevel, moduleSize, codeType, QRCodeVersion, encodingModes, codeQr)


End Sub



