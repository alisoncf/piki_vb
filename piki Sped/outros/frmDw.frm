VERSION 5.00
Begin VB.Form frmDw 
   Caption         =   "Armazém de Dados"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11490
   Icon            =   "frmDw.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   11490
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Perspectiva"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.OptionButton Option2 
         Caption         =   "Movimento de Produtos"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Movimento Mensal"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmDw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
