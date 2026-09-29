VERSION 5.00
Begin VB.Form frmOpcoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Vendas"
      Height          =   2655
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton Command2 
         Caption         =   "&Edição de Vendas"
         Height          =   615
         Left            =   600
         TabIndex        =   2
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Nova Venda"
         Height          =   615
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmOpcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
