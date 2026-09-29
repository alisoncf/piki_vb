VERSION 5.00
Begin VB.Form frmMovFiscalProduto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Movimento Fiscal"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Pesquisar"
      Height          =   495
      Left            =   8880
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtAte 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7080
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtDe 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5880
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   495
      Left            =   10080
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtProduto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label lblTotalInvent 
      AutoSize        =   -1  'True
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5040
      TabIndex        =   15
      Top             =   1200
      Width           =   390
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Inventário:"
      Height          =   195
      Left            =   5040
      TabIndex        =   14
      Top             =   840
      Width           =   750
   End
   Begin VB.Label lblTotalEntradas 
      AutoSize        =   -1  'True
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   13
      Top             =   1200
      Width           =   390
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Entrada:"
      Height          =   195
      Left            =   3360
      TabIndex        =   12
      Top             =   840
      Width           =   600
   End
   Begin VB.Label lblTotalNfe 
      AutoSize        =   -1  'True
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1800
      TabIndex        =   11
      Top             =   1200
      Width           =   390
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "NFE:"
      Height          =   195
      Left            =   1800
      TabIndex        =   10
      Top             =   840
      Width           =   360
   End
   Begin VB.Label lblTotalF 
      AutoSize        =   -1  'True
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ECF:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Até:"
      Height          =   195
      Left            =   7080
      TabIndex        =   5
      Top             =   0
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "De:"
      Height          =   195
      Left            =   5880
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   600
   End
End
Attribute VB_Name = "frmMovFiscalProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSair_Click()
    Unload Me
End Sub
