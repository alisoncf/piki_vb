VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Exemplo QR CODE"
   ClientHeight    =   4515
   ClientLeft      =   2970
   ClientTop       =   2685
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   7410
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4080
      TabIndex        =   1
      Text            =   "COM5"
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ImprimeCodigoQRCode"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin MSCommLib.MSComm MSC_COM 
      Left            =   2880
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   512
   End
   Begin VB.Label Label1 
      Caption         =   "Porta impressora"
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function ConfiguraModeloImpressora Lib "MP2032.DLL" (ByVal ModeloImpressora As Integer) As Integer
Private Declare Function IniciaPorta Lib "MP2032.DLL" (ByVal Porta As String) As Integer
Private Declare Function ComandoTX Lib "MP2032.DLL" (ByVal BufTrans As String, ByVal Flag As Integer) As Integer
Private Declare Function ImprimeCodigoQRCODE Lib "MP2032.DLL" (ByVal errorCorrectionLevel As Integer, ByVal moduleSize As Integer, ByVal codeType As Integer, ByVal QRCodeVersion As Integer, ByVal encodingModes As Integer, ByVal codeQr As String) As Integer
Private Declare Function FechaPorta Lib "MP2032.DLL" () As Integer
            
            Dim sPorta As String
            Dim iRetorno As Integer
            Dim sComando As String
            Dim sModuloWidth As String
            Dim sErrorCorrection As String
            Dim sStoreSymbolDate As String
            Dim sBarCodeData As String
            Dim sPrintQrCodeSymbol As String
            Dim sCenter As String
            
      
Private Sub Command1_Click()
            Dim iRetorno As Integer

iRetorno = ConfiguraModeloImpressora(8)
    sPorta = Text1.Text
        iRetorno = IniciaPorta(sPorta)
        Dim errorCorrectionLevel As Long
        Dim moduleSize As Long
        
        Dim codeType As Long
        Dim QRCodeVersion As Long
        Dim encodingModes As Long
        Dim codeQr As String
    errorCorrectionLevel = 1
    
    moduleSize = 7
    codeType = 0
    QRCodeVersion = 10
    encodingModes = 1
    codeQr = "123ABC"

    iRetorno = ImprimeCodigoQRCODE(errorCorrectionLevel, moduleSize, codeType, QRCodeVersion, encodingModes, codeQr)

iRetorno = FechaPorta()

End Sub

