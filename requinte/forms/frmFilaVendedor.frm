VERSION 5.00
Begin VB.Form frmFilaVendedor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fila de Vendedores"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9105
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lst 
      Height          =   2595
      ItemData        =   "frmFilaVendedor.frx":0000
      Left            =   120
      List            =   "frmFilaVendedor.frx":0002
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedores Cadastrados:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1830
   End
End
Attribute VB_Name = "frmFilaVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
