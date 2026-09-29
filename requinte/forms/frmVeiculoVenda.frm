VERSION 5.00
Begin VB.Form frmVeiculoVenda 
   Caption         =   "Veículo"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8655
   Icon            =   "frmVeiculoVenda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAbertura 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5880
      TabIndex        =   9
      Text            =   "dd/mm/yyyy HH:MM:SS"
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox txtDiagnostico 
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   7935
   End
   Begin VB.TextBox txtDescricao 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   7935
   End
   Begin VB.PictureBox pcBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   8655
      TabIndex        =   8
      Top             =   2805
      Width           =   8655
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   615
         Left            =   6600
         Picture         =   "frmVeiculoVenda.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Confirmar"
         Height          =   615
         Left            =   4800
         Picture         =   "frmVeiculoVenda.frx":1035
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   14
         Top             =   480
         Width           =   1410
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   525
      End
      Begin VB.Label lblVenda 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venda: "
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.TextBox txtPlaca 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Abertura:"
      Height          =   195
      Left            =   5880
      TabIndex        =   10
      Top             =   360
      Width           =   645
   End
   Begin VB.Label lblDiag 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Diagnóstico:"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição:"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Placa:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   450
   End
End
Attribute VB_Name = "frmVeiculoVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blnEdita As Boolean


Private Sub cmdOk_Click()
    If Me.tag = "" Then
        Exit Sub
    End If
    If blnEdita Then
        If Vendas.AlteraVeiculoVenda(Me.tag, txtPlaca.Text, txtDescricao.Text, txtAbertura.Text, _
            txtDiagnostico.Text, 0) Then
        End If
    Else
        If Vendas.GravaVeiculoVenda(Me.tag, txtPlaca.Text, txtDescricao.Text, txtAbertura.Text, _
            txtDiagnostico.Text, 0) Then
        End If
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub txtModelo_Change()
    txtModelo.tag = ""
End Sub

Private Sub txtModelo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
    End If
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

Private Sub txtDiagnostico_LostFocus()
    P_AjustaTexto txtDiagnostico
End Sub

Private Sub txtPlaca_GotFocus()
    P_SelecTexto txtPlaca
End Sub

Private Sub txtPlaca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtPlaca_LostFocus()
    AjustaPlaca txtPlaca
End Sub
