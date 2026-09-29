VERSION 5.00
Begin VB.Form frmDadosChPag 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dados do Cheque a Pagar"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4740
   Icon            =   "frmDadosChPag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C6DADD&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Tag             =   "!"
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtbxNum_Che 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   12
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtbxVencimento 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnConfirmar 
         Caption         =   "Con&firmar"
         Height          =   615
         Left            =   360
         Picture         =   "frmDadosChPag.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2520
         Width           =   1575
      End
      Begin VB.ComboBox cmbxContas 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox txtbxValor_Che 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   14
         TabIndex        =   8
         Text            =   "0,00"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cheque Nº:"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "V&encimento:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   885
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Conta:"
         Height          =   195
         Left            =   540
         TabIndex        =   1
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor:"
         Height          =   195
         Left            =   600
         TabIndex        =   7
         Top             =   1560
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmDadosChPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnConfirmar_Click()
    Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    P_PreencheContas cmbxContas, 1, False
End Sub

Private Sub txtbxNum_Che_GotFocus()
    ganhaFoco txtbxNum_Che
End Sub

Private Sub txtbxNum_Che_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxNum_Che_LostFocus()
    AjustaTexto txtbxNum_Che
End Sub
Private Sub txtbxValor_Che_GotFocus()
    ganhaFoco txtbxNum_Che
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
    ganhaFoco txtbxNum_Che
End Sub

Private Sub txtbxVencimento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxVencimento_LostFocus()
    P_AjustaData txtbxVencimento
End Sub


