VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDadosFornada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dados da Distribuição de Fornadas do Produto"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frControle 
      Caption         =   "Distribuição das Fornadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7575
      Begin VB.Data dtProd 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSDBGrid.DBGrid dbgrdProd 
         Bindings        =   "frmDadosFornada.frx":0000
         Height          =   2295
         Left            =   120
         OleObjectBlob   =   "frmDadosFornada.frx":0015
         TabIndex        =   2
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.CommandButton cmndbtnConfirma 
      Cancel          =   -1  'True
      Caption         =   "F&echar"
      Default         =   -1  'True
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   1815
   End
End
Attribute VB_Name = "frmDadosFornada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnConfirma_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
    dtProd.Connect = banco.Connect
    busca
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select (tf.qtd_pro-tf.reservado_pro) as total,tf.* "
    strSQL = strSQL & " from tab_fornadas as tf where cod_pro='" & Me.Tag & "'"
    'strSQL = strSQL & " and cod_loj='" & strCod_Loj & "' "
    dtProd.RecordSource = strSQL
    dtProd.Refresh
    
End Sub
