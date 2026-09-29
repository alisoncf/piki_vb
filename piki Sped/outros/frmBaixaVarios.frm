VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBaixaVarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixa Vinculada"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   Icon            =   "frmBaixaVarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   2880
      TabIndex        =   2
      Top             =   2280
      Width           =   3735
      Begin VB.CommandButton Command3 
         Caption         =   "Fechar"
         Height          =   495
         Left            =   2520
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Baixar"
         Height          =   495
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Boletos"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxGrd 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3625
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   2655
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         Height          =   195
         Left            =   1080
         TabIndex        =   6
         Top             =   480
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parcelas:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   660
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         Height          =   195
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor total:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmBaixaVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    With msflxGrd
        .FixedCols = 0
        .rows = 1
        .Cols = 6
        .Col = 0: .Text = "Seq_Doc"
        .Col = 1: .Text = "Seq_Itd"
        .Col = 2: .Text = "Número"
        .Col = 3: .Text = "Parcela"
        .Col = 4: .Text = "Vence em:"
        .Col = 5: .Text = "Valor"
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 1000
        .ColWidth(3) = 800
        
    End With
End Sub
