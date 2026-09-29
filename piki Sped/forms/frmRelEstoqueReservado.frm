VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelEstoqueReservado 
   Caption         =   "Estoque Reservado"
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9465
   Icon            =   "frmRelEstoqueReservado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReservado 
      Caption         =   "Reservado"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Fe&char (Esc)"
      Height          =   615
      Left            =   7680
      Picture         =   "frmRelEstoqueReservado.frx":1EFA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnHtml 
      Caption         =   "&Imprimir Lista"
      Height          =   615
      Left            =   6120
      Picture         =   "frmRelEstoqueReservado.frx":2065
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   6800
      _Version        =   393216
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
      Caption         =   "Vendas em aberto encontradas"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "data_ven"
         Caption         =   "Data"
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
      BeginProperty Column01 
         DataField       =   "nomecli_ven"
         Caption         =   "Cliente"
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
      BeginProperty Column02 
         DataField       =   "qtd_ite"
         Caption         =   "Qtd"
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
      BeginProperty Column03 
         DataField       =   "vendedor"
         Caption         =   "Vendededor"
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
            ColumnWidth     =   2039,811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4605,166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   675,213
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtProd 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6135
   End
   Begin VB.CommandButton cmdProd 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6240
      Picture         =   "frmRelEstoqueReservado.frx":2472
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Produto"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "frmRelEstoqueReservado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim bActive As Boolean

Private Sub cmdProd_Click()
    buscaProd
End Sub

Private Sub cmdReservado_Click()
    P_AjustaEstoqueReservado txtProd.tag
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If bActive = True Then Exit Sub
    If Me.tag <> "" Then
        busca
    End If
    bActive = False
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
End Sub

Private Sub txtProd_Change()
    txtProd.tag = ""
End Sub

Private Sub txtProd_GotFocus()
    P_SelecTexto txtProd
End Sub

Private Sub txtProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaProd
    End If
End Sub
Sub buscaProd()
    Dim p As Produto
    
    p = F_BuscaProdSaida(txtProd.Text)
    If p.cod_pro <> "" Then
        txtProd.Text = p.nome_pro
        txtProd.tag = p.cod_pro
        busca
    End If
    
    
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select tv.*,tf.guerra_fun as vendedor,tp.er_pro,tp.reservado_pro,ti.qtd_ite from "
    strSQL = strSQL & TAB_VENDAS & " as tv inner join " & TAB_ITVENDAS & " as ti on tv.num_ven=ti.num_ven "
    strSQL = strSQL & " inner join " & TAB_PRODUTOS & " as tp on ti.cod_pro=tp.cod_pro "
    strSQL = strSQL & " left join tab_funcionarios as tf on tv.abrevenda_usu=tf.cod_fun "
    strSQL = strSQL & " where tp.cod_pro='" & txtProd.tag & "'"
    strSQL = strSQL & " and situacao_ven<4"
    abreObjetoAdo ado, strSQL, DataGrid1
    
    
End Sub
