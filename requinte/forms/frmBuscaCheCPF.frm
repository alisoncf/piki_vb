VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBuscaCheCPF 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Busca de Contas do CPF:"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3625
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmBuscaCheCPF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
With msflxgrd
        .rows = 1
        .FixedCols = 0
        .Cols = 5
        .col = 0: .Text = "Cód."
        .col = 1: .Text = "Banco"
        .col = 2: .Text = "Agência"
        .col = 3: .Text = "Conta C/"
        .col = 4: .Text = "Titular"
        
        .ColWidth(1) = 2000
        .ColWidth(4) = 2500
    End With
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.Tag = "X"
    Me.Hide
End Sub
