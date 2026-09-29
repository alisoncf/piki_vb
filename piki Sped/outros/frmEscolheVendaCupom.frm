VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEscolheVendaCupom 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Escolha a Venda"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2778
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Escolha a Venda e Tecle <Enter>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2460
   End
End
Attribute VB_Name = "frmEscolheVendaCupom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    With msflxgrd
        .FixedRows = 1
        .FixedCols = 0
        .Cols = 4
        .rows = 1
        .col = 0: .Text = "Caixa"
        .col = 1: .Text = "Valor"
        .col = 2: .Text = "Data/Hora"
        .col = 3: .Text = "Cliente"
        .ColWidth(0) = 400
        .ColWidth(2) = 2200
        .ColWidth(3) = 3000
    End With
End Sub

Private Sub msflxgrd_DblClick()
    msflxgrd.col = 0
    Me.Tag = msflxgrd.Text
    Me.Hide
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        msflxgrd_DblClick
    End If
End Sub
