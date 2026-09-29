VERSION 5.00
Begin VB.Form frmExport 
   Caption         =   "Exportar Dados "
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDocs 
      Caption         =   "Docs"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdProdutos 
      Caption         =   "Produtos"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDocs_Click()
    
    XMLBD.GeraXmlTabela "tab_doc"
End Sub

Private Sub cmdProdutos_Click()
    XMLBD.GeraXmlTabela "tab_produtos"
    XMLBD.GeraXmlTabela "tab_produtos"
End Sub
