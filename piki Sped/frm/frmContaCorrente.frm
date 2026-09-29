VERSION 5.00
Begin VB.Form frmContaCorrente 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Escolha a Contas Corrente"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3285
   Icon            =   "frmContaCorrente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtbxData 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ListBox lstConta 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1605
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Escolha a conta corrente e tecle enter"
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmndbtnCOnfirma 
      Caption         =   "F7 - &Confirmar "
      Height          =   615
      Left            =   120
      Picture         =   "frmContaCorrente.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "F9 - Cancela&r"
      Height          =   615
      Left            =   1680
      Picture         =   "frmContaCorrente.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Data:"
      Height          =   195
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   390
   End
End
Attribute VB_Name = "frmContaCorrente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnConfirma_Click()
    Me.tag = Left(lstConta.Text, 3) & txtbxData.Text
    
    Me.Hide
End Sub

Private Sub cmndbtnSair_Click()
    Me.tag = ""
    Me.Hide
End Sub

Private Sub Form_Activate()
    lstConta.ListIndex = 0
    'MudaAparencia Me
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        cmndbtnConfirma_Click
    ElseIf KeyCode = vbKeyF9 Then
        cmndbtnSair_Click
    End If
End Sub

Private Sub Form_Load()
    txtbxData.Text = Date
    P_PreencheContas lstConta
End Sub

Private Sub lstConta_DblClick()
    cmndbtnConfirma_Click
End Sub

Private Sub lstConta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxData_GotFocus()
    ganhaFoco txtbxData
End Sub

Private Sub txtbxData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxData_LostFocus()
    P_AjustaData txtbxData
End Sub
