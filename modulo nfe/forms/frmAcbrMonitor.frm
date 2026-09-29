VERSION 5.00
Begin VB.Form frmAcbrMonitor 
   Caption         =   "Monitor ACBR"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Setar CSC (Código de Segurança do Contribuinte)"
      Height          =   1095
      Left            =   360
      TabIndex        =   11
      Top             =   1320
      Width           =   8535
      Begin VB.CommandButton Command2 
         Caption         =   "Set CSC"
         Height          =   495
         Left            =   6480
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtCSC 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox txtIDCsC 
         Height          =   285
         Left            =   3240
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "CSC"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ID."
         Height          =   195
         Left            =   3240
         TabIndex        =   15
         Top             =   240
         Width           =   210
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resposta"
      Height          =   1935
      Left            =   360
      TabIndex        =   8
      Top             =   4320
      Width           =   8535
      Begin VB.CommandButton cmdResp 
         Caption         =   "Buscar Resposta"
         Height          =   375
         Left            =   6480
         TabIndex        =   10
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtStatusServ 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   240
         Width           =   8175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Outros Comandos"
      Height          =   1455
      Left            =   360
      TabIndex        =   6
      Top             =   2880
      Width           =   8535
      Begin VB.CommandButton Command1 
         Caption         =   "Testar Status do Serviço"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Setar Certificado SSL"
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   8535
      Begin VB.TextBox txtSenhaPFx 
         Height          =   285
         Left            =   3240
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtCaminhoPfx 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   3015
      End
      Begin VB.CommandButton cmdSetCertificado 
         Caption         =   "setCertificado"
         Height          =   495
         Left            =   6480
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Senha:"
         Height          =   195
         Left            =   3240
         TabIndex        =   4
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Caminho do Arquivo .pfx"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1725
      End
   End
End
Attribute VB_Name = "frmAcbrMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim empresa As sped_empresa.sped_empresa
Dim bActive As Boolean
Private Sub cmdResp_Click()
    AguardarArquivo
End Sub

Private Sub cmdSetCertificado_Click()
    
    
    
setarCertificado strCod_Emp
    
    
End Sub

Private Sub Command1_Click()
    If GravaArquivoExterno(CaminhoArquivoEntradaMonitor, "NFe.StatusServico") Then
        AguardarArquivo
    End If
End Sub

Sub AguardarArquivo()
    
    
    Sleep 1000
    If Dir(CaminhoArquivoSaidaMonitor, vbArchive) <> "" Then
        txtStatusServ.text = LeArquivoExternoUTF8(CaminhoArquivoSaidaMonitor)
        Kill CaminhoArquivoSaidaMonitor
    End If



End Sub

Private Sub Command2_Click()
    setarCertificado strCod_Emp
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    empresaSelecionada = BuscaSpedEmpresa(strCod_Emp)

    txtCaminhoPfx.text = empresaSelecionada.caminho_certificado
    txtSenhaPFx.text = empresaSelecionada.senha_certificado
    
    txtCSC.text = empresaSelecionada.token
    txtIDCsC.text = empresaSelecionada.idCSC
    bActive = True
    
    
End Sub

