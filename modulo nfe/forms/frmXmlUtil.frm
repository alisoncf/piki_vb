VERSION 5.00
Begin VB.Form frmXmlUtil 
   Caption         =   "Utilitário XML/JSON"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14475
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   14475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProdEstoque 
      Caption         =   "Produtos (sem estoque)"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdTableToJson 
      Caption         =   "JSON"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtTabela 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tabela:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "frmXmlUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub TabJson()
    
    RecordSetToJson "select * from " & txtTabela.Text, txtTabela.Text, "", 0
    
    
    
    
    
End Sub

Private Sub cmdProdEstoque_Click()
    Dim strSQL As String

    
End Sub

Private Sub cmdTableToJson_Click()
    TabJson
End Sub
