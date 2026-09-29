VERSION 5.00
Begin VB.Form frmDetalheProdutoUniformes 
   Caption         =   "Detalhe do produto"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13080
   Icon            =   "frmDetalheProdutoUniformes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   13080
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EEDFD5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   13080
      TabIndex        =   11
      Top             =   2595
      Width           =   13080
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   735
         Left            =   10800
         Picture         =   "frmDetalheProdutoUniformes.frx":1084A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdEfetivar 
         Caption         =   "Confirmar"
         Height          =   735
         Left            =   8760
         Picture         =   "frmDetalheProdutoUniformes.frx":10DD8
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.PictureBox pcConfeccao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2295
      ScaleWidth      =   12855
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   12855
      Begin VB.TextBox txtDetalhamento 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   1245
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   12375
      End
      Begin VB.TextBox txtCor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtSilk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   2640
         TabIndex        =   4
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtBordado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   5160
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bordado"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5160
         TabIndex        =   5
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Silk"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2640
         TabIndex        =   3
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label37 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cor"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label38 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Detalhamento:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmDetalheProdutoUniformes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdEfetivar_Click()
    Me.tag = "x"
    Me.Hide
End Sub



Private Sub cmdSair_Click()
    Me.tag = ""
    Me.Hide
End Sub

Private Sub Form_Load()
    pcConfeccao.Visible = True
End Sub

Private Sub txtCor_GotFocus()
    P_SelecTexto txtCor
End Sub

Private Sub txtCor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub


Private Sub txtCor_LostFocus()
    P_AjustaTexto txtCor, nada_muda
End Sub


Private Sub txtDetalhamento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = vbCtrlMask Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDetalhamento_LostFocus()
    P_AjustaTexto txtDetalhamento, nada_muda
End Sub

Private Sub txtSilk_GotFocus()
    P_SelecTexto txtSilk
End Sub

Private Sub txtSilk_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub


Private Sub txtSilk_LostFocus()
    P_AjustaTexto txtSilk, nada_muda
End Sub

Private Sub txtBordado_GotFocus()
    P_SelecTexto txtBordado
End Sub

Private Sub txtBordado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub


Private Sub txtBordado_LostFocus()
    P_AjustaTexto txtBordado, nada_muda
End Sub

