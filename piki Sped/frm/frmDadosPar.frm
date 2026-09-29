VERSION 5.00
Begin VB.Form frmDadosPar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dados da Parcela"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   Icon            =   "frmDadosPar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   6090
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.TextBox txtbxVencimento 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnConfirmar 
         Caption         =   "Con&firmar"
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   390
         Width           =   1455
      End
      Begin VB.CommandButton cmbxSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   375
         Left            =   4320
         TabIndex        =   6
         Top             =   390
         Width           =   1455
      End
      Begin VB.TextBox txtbxValor_Che 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   14
         TabIndex        =   4
         Text            =   "0,00"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimento:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Va&lor:"
         Height          =   195
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmDadosPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxSair_Click()
    Me.Hide
End Sub

Private Sub cmndbtnConfirmar_Click()
    If txtbxValor_Che.Text = "0,00" Then
        MsgBox "Informe um valor válido!", vbExclamation
        Exit Sub
    End If
    If F_Vazio(txtbxVencimento.Text) Then
        MsgBox "Informe um vencimento válido!", vbExclamation
        Exit Sub
    End If
    Me.tag = "X"
    Me.Hide
End Sub

Private Sub Form_Load()
    'MudaAparencia Me
End Sub

Private Sub txtbxValor_Che_GotFocus()
    ganhaFoco txtbxValor_Che
End Sub

Private Sub txtbxValor_Che_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxValor_Che_LostFocus()
    PerdeFoco txtbxValor_Che
End Sub

Private Sub txtbxVencimento_GotFocus()
    ganhaFoco txtbxVencimento
End Sub

Private Sub txtbxVencimento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxVencimento_LostFocus()
    P_AjustaData txtbxVencimento
End Sub
