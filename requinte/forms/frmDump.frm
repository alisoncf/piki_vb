VERSION 5.00
Begin VB.Form frmDump 
   Caption         =   "DUMP - Transferência de Dados"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      ScaleHeight     =   945
      ScaleWidth      =   5505
      TabIndex        =   5
      Top             =   2760
      Width           =   5535
      Begin VB.CommandButton cmdArquivo 
         Caption         =   "..."
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtNomeArq 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Text            =   "c:\"
         Top             =   360
         Width           =   3135
      End
      Begin VB.CommandButton cmdGerar 
         Caption         =   "Gerar Arquivo"
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salvar para o Arquivo:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tabelas"
      Height          =   1575
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   5415
      Begin VB.CheckBox chkClientes 
         Caption         =   "Clientes"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   3375
      End
      Begin VB.CheckBox chkProd 
         Caption         =   "Produtos (+Fornecedores+Seção+sub)"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.TextBox txtServer 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   5415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID Servidor"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   795
   End
End
Attribute VB_Name = "frmDump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGerar_Click()
    Shell txtServer.Text & "\mysql\bin\mysqldump.exe requinte tab_subsec > " & App.path & "\migras.txt", vbNormalFocus
End Sub

Private Sub Form_Load()
    txtNomeArq.Text = App.path & "\"
End Sub
