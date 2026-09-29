VERSION 5.00
Begin VB.Form frmTefAvulso 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Recebimento TEF "
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnFinanceira 
      Caption         =   "..."
      Height          =   285
      Left            =   4800
      TabIndex        =   4
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox txtbxFinanceira 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox txtbxValor 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Text            =   "0,00"
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Financeira:"
      Height          =   195
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   405
   End
End
Attribute VB_Name = "frmTefAvulso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
