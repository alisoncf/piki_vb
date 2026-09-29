VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFila 
   Caption         =   "Fila de Transferência"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timer_doc 
      Interval        =   5000
      Left            =   0
      Top             =   0
   End
   Begin VB.CheckBox chkDoc 
      Caption         =   "Documentos"
      Height          =   735
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
   Begin VB.CheckBox chkProdutos 
      Caption         =   "Produto"
      Height          =   735
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid gridCli 
      Height          =   5175
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "codigo"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
      EndProperty
   End
   Begin VB.Timer timer_busca 
      Interval        =   5000
      Left            =   7920
      Top             =   960
   End
   Begin VB.CheckBox chkCliente 
      Caption         =   "Cliente"
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Timer timer_cliente 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7800
      Top             =   1800
   End
   Begin MSDataGridLib.DataGrid gridProd 
      Height          =   5175
      Left            =   2880
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "codigo"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid gridDoc 
      Height          =   5175
      Left            =   5280
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "codigo"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFila"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoCli As New ADODB.Recordset
Dim adoprod As New ADODB.Recordset
Dim adodoc As New ADODB.Recordset
Private Sub chkCliente_Click()
    
    timer_cliente.Enabled = chkCliente.Value = 1
End Sub

Private Sub chkDoc_Click()
    Dim ret As String
    ret = executa_fila("2")
End Sub

Private Sub chkProdutos_Click()
    Dim ret As String
    ret = executa_fila("3")
End Sub

Private Sub Form_Activate()
    busca_cli
    busca_prod
    busca_doc
End Sub

Private Sub timer_busca_Timer()
    busca_cli
    busca_prod
    busca_doc
End Sub

Private Sub timer_cliente_Timer()
    Dim ret As String
    ret = executa_fila("1")
    'lblRetornoCliente.caption = ret
End Sub
Sub busca_cli()
    Dim strSQL As String
    
    strSQL = " select * from fila_transferencia where tipo = 1 order by id "
    
    abreObjetoAdo adoCli, strSQL, gridCli
    chkCliente.caption = "Clientes (" & adoCli.RecordCount & ") "
    
    
    
End Sub
Sub busca_doc()
    Dim strSQL As String
    
    strSQL = " select * from fila_transferencia where tipo = 2 order by id "
    
    abreObjetoAdo adodoc, strSQL, gridDoc
    chkDoc.caption = "Documentos (" & adodoc.RecordCount & ") "
    
    
    
End Sub

Sub busca_prod()
    Dim strSQL As String
    
    strSQL = " select * from fila_transferencia where tipo = 3 order by id "
    
    abreObjetoAdo adoprod, strSQL, gridProd
    chkProdutos.caption = "Produtos (" & adoprod.RecordCount & ") "
    
    
    
End Sub
