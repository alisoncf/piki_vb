VERSION 5.00
Begin VB.Form frmConfiguraSitef 
   Caption         =   "Configuração de Terminal"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtbxParamAdic 
      Height          =   285
      Left            =   1560
      TabIndex        =   10
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox txtbxIdReservado 
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox txtbxIdTerminal 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtbxIdLoja 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox txtbxEndSitef 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton cmndbtnOk 
      Caption         =   "Ok"
      Height          =   615
      Left            =   2400
      TabIndex        =   0
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Param. Adic."
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reservado:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID Terminal"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID Loja"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço Sitef:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmConfiguraSitef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnOk_Click()
    Dim i As Long
    Dim result As Integer
    
    'i = IniciaFuncaoSiTefInterativo(
    
    i = ConfiguraIntSiTefInterativo(txtbxEndSitef.Text, txtbxIdLoja.Text, txtbxIdTerminal.Text, result)
End Sub
