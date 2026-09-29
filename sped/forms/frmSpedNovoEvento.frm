VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSpedNovoEvento 
   Caption         =   "Novo Evento SPED"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2175
      Left            =   360
      TabIndex        =   8
      Top             =   2040
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3836
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   735
      Left            =   6240
      TabIndex        =   7
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Height          =   735
      Left            =   4800
      TabIndex        =   6
      Top             =   5040
      Width           =   1335
   End
   Begin VB.ComboBox cmbPeriodicidade 
      Height          =   315
      ItemData        =   "frmSpedNovoEvento.frx":0000
      Left            =   360
      List            =   "frmSpedNovoEvento.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox txtDescricao 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   5775
   End
   Begin VB.TextBox txtId 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Periodicidade:"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição:"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   210
   End
End
Attribute VB_Name = "frmSpedNovoEvento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()

End Sub

Private Sub txtDescricao_Change()

End Sub

Private Sub txtDescricao_GotFocus()
    P_SelecTexto txtDescricao
End Sub

Private Sub txtDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtDescricao_LostFocus()
    P_AjustaTexto txtDescricao
End Sub
