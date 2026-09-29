VERSION 5.00
Begin VB.Form frmPrincipalECF 
   Caption         =   "Topa Tudo da Construção - Comércio de Pisos R & R LTDA. - CONTROLE ECF"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11370
   Icon            =   "frmPrincipalECF.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmPrincipalECF.frx":0ECA
   ScaleHeight     =   8595
   ScaleWidth      =   11370
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame frControle 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   6375
      Begin VB.CommandButton Command6 
         Caption         =   "Cadastro C&lientes"
         Height          =   1095
         Left            =   480
         Picture         =   "frmPrincipalECF.frx":60AC
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5160
         Width           =   5295
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Sair"
         Height          =   1095
         Left            =   480
         Picture         =   "frmPrincipalECF.frx":61F6
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6360
         Width           =   5295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Emissão de Cupom Fiscal"
         Height          =   1095
         Left            =   480
         Picture         =   "frmPrincipalECF.frx":6340
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   5295
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Relatório de Saídas"
         Height          =   1095
         Left            =   480
         Picture         =   "frmPrincipalECF.frx":68CA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3960
         Width           =   5295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Controle ECF"
         Height          =   1095
         Left            =   480
         Picture         =   "frmPrincipalECF.frx":6D0C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1560
         Width           =   5295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Edição Produtos/Pesquisa de Preços"
         Height          =   1095
         Left            =   480
         Picture         =   "frmPrincipalECF.frx":714E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2760
         Width           =   5295
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X-pert Softwares"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   9600
      TabIndex        =   7
      Top             =   7680
      Width           =   1845
   End
End
Attribute VB_Name = "frmPrincipalECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    frmEdiPro.Show 1
End Sub

Private Sub Command2_Click()
    If F_NumeroCupom = "" Then
        MsgBox "Não há Impressora Fiscal Conectada!", vbExclamation
        Exit Sub
    End If
    frmImpressoraECF.Show 1
    
End Sub

Private Sub Command3_Click()
    frmRelSai.Show 1
End Sub

Private Sub Command4_Click()
    If F_NumeroCupom = "" Then
        MsgBox "Não há impressora fiscal conectada!", vbExclamation
        Exit Sub
    End If
    frmSaidaCupom.Show 1
End Sub

Private Sub Command5_Click()
    If MsgBox("Deseja Realmente Sair?", vbYesNo + vbQuestion) = vbYes Then
        P_Fim
    End If
End Sub

Private Sub Command6_Click()
    frmEdiClientes.Show 1
End Sub

Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub

