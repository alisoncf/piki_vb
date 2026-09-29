VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmParametrosRegras 
   Caption         =   "Regras"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9810
   Icon            =   "frmParametrosRegras.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Adicionar"
      Height          =   495
      Left            =   6960
      TabIndex        =   5
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtValor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   6000
      Width           =   5535
   End
   Begin VB.TextBox txtChave 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grade 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   9763
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Left            =   1320
      TabIndex        =   3
      Top             =   5760
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Chave"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   5760
      Width           =   465
   End
End
Attribute VB_Name = "frmParametrosRegras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from parametro_estrutura order by id"
    
    abreObjetoAdo rPesq, strSQL
    
    AjustaGridData grade, "4", 1, 0, "N. ,id,regra,valor", "@item,id,regra,valor ", "1000,1000,3000,2000", rPesq, "", True, 0, False
    
    
    
    
End Sub
