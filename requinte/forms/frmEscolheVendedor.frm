VERSION 5.00
Begin VB.Form frmEscolheVendedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendedores:"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   Icon            =   "frmEscolheVendedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4020
      ItemData        =   "frmEscolheVendedor.frx":030A
      Left            =   0
      List            =   "frmEscolheVendedor.frx":030C
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lblNome_ven 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   4320
      Width           =   45
   End
End
Attribute VB_Name = "frmEscolheVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Me.Tag = lst.Tag
    Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Me.Tag = ""
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    P_PreencheVendedor lst, 1
    P_LocalizaOperacao F_BuscaUmCampo("nome_fun", "tab_funcionarios", _
        "cod_fun='" & Me.Tag & "'"), lst
    Me.Tag = ""
End Sub

Private Sub lst_Click()
    lst.Tag = F_BuscaUmCampo("cod_fun", "tab_funcionarios", "guerra_fun='" & _
        lst.Text & "'")
    lblNome_ven.caption = F_BuscaUmCampo("nome_fun", "tab_funcionarios", "guerra_fun='" & _
        lst.Text & "'")
End Sub

Private Sub lst_DblClick()
    Command1_Click
End Sub

Private Sub lst_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        lst_DblClick
    End If
End Sub


