VERSION 5.00
Begin VB.Form frmEscolhe 
   Appearance      =   0  'Flat
   BackColor       =   &H00EF9C7C&
   Caption         =   "Menu TEF"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   Icon            =   "frmEscolhe.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnCancela 
      Caption         =   "Cancela F9"
      Height          =   735
      Left            =   5280
      Picture         =   "frmEscolhe.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton cmndbtnOk 
      Cancel          =   -1  'True
      Caption         =   "Confirmar (Enter)"
      Height          =   735
      Left            =   120
      Picture         =   "frmEscolhe.frx":2751
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton cmndbtnRetornar 
      Caption         =   "Retornar (Esc)"
      Height          =   735
      Left            =   2520
      Picture         =   "frmEscolhe.frx":2BF9
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   2175
   End
   Begin VB.ListBox lst 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   7335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmEscolhe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sOp As String

Private Sub Command1_Click()
    Me.tag = ""
    Me.Hide
End Sub

Sub separaOpcao()
    Dim c As New Collection
    
    
    
End Sub

Private Sub cmndbtnCancela_Click()
    Me.tag = "-1"
    Me.Hide
End Sub

Private Sub cmndbtnOk_Click()
    Command1_Click
End Sub

Private Sub cmndbtnRetornar_Click()
    Me.tag = "1"
    Me.Hide
End Sub

Private Sub Form_Activate()
    Dim c As New Collection
    Dim iCount As Integer
    On Error Resume Next
    
    Set c = F_SeparaTexto(sOp, ";", False)
    
    For iCount = 1 To c.count
        lst.AddItem c(iCount)
    Next
    lst.ListIndex = 0
    lst.SetFocus
    
End Sub

Private Sub lst_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Command1_Click
    End If
End Sub
