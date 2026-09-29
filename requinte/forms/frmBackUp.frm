VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBackUp 
   Caption         =   "Efetuar Backup"
   ClientHeight    =   4125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6975
   Icon            =   "frmBackUp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtZip 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Text            =   "ZIp"
      Top             =   3720
      Width           =   6495
   End
   Begin VB.TextBox txtCaminhoBanco 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Text            =   "\\10.1.1.10\c\mysql\data\requinte\*.*"
      Top             =   1200
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copiar"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   5880
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtCaminho 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5775
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      Height          =   285
      Left            =   6000
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   285
      Left            =   6000
      TabIndex        =   6
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caminho do Banco"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caminho da Cópia"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1290
   End
End
Attribute VB_Name = "frmBackUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPath_Click()
    Dim s As String
    dlg.filename = F_SoNumero(Now) & ".rar"
    dlg.Filter = "*.rar"
    dlg.ShowSave
    s = dlg.filename
    txtCaminho.Text = s
    
End Sub

Sub DoBackup()
    Dim Drive As String
    
    
    
End Sub

Private Sub Command1_Click()
    If strFechacai_usu <> "" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
    
    
    Compacta txtCaminho.Text, txtCaminhoBanco
    
End Sub

Private Sub Form_Load()
    DoEvents
    InicializaZip Me, txtZip
End Sub
