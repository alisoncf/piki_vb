VERSION 5.00
Begin VB.Form frmLeValor 
   Appearance      =   0  'Flat
   BackColor       =   &H00EF9C7C&
   Caption         =   "Informe um Valor"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   Icon            =   "frmLeValor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnCancela 
      Caption         =   "Cancela F9"
      Height          =   735
      Left            =   5280
      Picture         =   "frmLeValor.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox txtbxValor 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   7335
   End
   Begin VB.CommandButton cmndbtnOk 
      Cancel          =   -1  'True
      Caption         =   "Confirmar (Enter)"
      Height          =   735
      Left            =   120
      Picture         =   "frmLeValor.frx":2751
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmndbtnRetornar 
      Caption         =   "Retornar (Esc)"
      Height          =   735
      Left            =   2640
      Picture         =   "frmLeValor.frx":2BF9
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Escolha a Opção Desejada"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmLeValor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sOp As String

Private Sub Command1_Click()
    Me.Tag = ""
    Me.Hide
End Sub

Sub separaOpcao()
    Dim c As New Collection
    
End Sub

Private Sub cmndbtnCancela_Click()
    Me.Tag = "-1"
    Me.Hide
End Sub

Private Sub cmndbtnOk_Click()
    Command1_Click
End Sub

Private Sub cmndbtnRetornar_Click()
    Me.Tag = "1"
    Me.Hide
End Sub

Private Sub Form_Activate()
    Dim c As New Collection
    Dim iCount As Integer
    
    
    'Set c = F_SeparaTexto(sOp, ";", False)
    
    'For iCount = 1 To c.count
    '    lst.AddItem c(iCount)
    'Next
    'lst.ListIndex = 0
    'lst.SetFocus
    
End Sub

Private Sub lst_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Command1_Click
    End If
End Sub

Private Sub txtbxValor_GotFocus()
    P_SelecTexto txtbxValor
End Sub

Private Sub txtbxValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnOk_Click
    End If
End Sub
