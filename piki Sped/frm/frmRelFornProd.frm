VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelFornProd 
   Caption         =   "Relatório de Faturamento  Marca/Fornecedor"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   12210
   Icon            =   "frmRelFornProd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   12210
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   12210
      TabIndex        =   5
      Top             =   0
      Width           =   12210
      Begin VB.TextBox txtDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtForn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   9
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txtCat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8280
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   495
         Left            =   10800
         Picture         =   "frmRelFornProd.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdProd 
         Height          =   285
         Left            =   5400
         Picture         =   "frmRelFornProd.frx":04F5
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdCat 
         Height          =   285
         Left            =   7920
         Picture         =   "frmRelFornProd.frx":08A0
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdMarca 
         Height          =   285
         Left            =   10440
         Picture         =   "frmRelFornProd.frx":0C4B
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Até"
         Height          =   195
         Left            =   1320
         TabIndex        =   18
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedor"
         Height          =   195
         Left            =   2520
         TabIndex        =   17
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria:"
         Height          =   195
         Left            =   5760
         TabIndex        =   16
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marcar:"
         Height          =   195
         Left            =   8280
         TabIndex        =   15
         Top             =   120
         Width           =   540
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   12180
      TabIndex        =   1
      Top             =   7185
      Width           =   12210
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   10560
         Picture         =   "frmRelFornProd.frx":0FF6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   720
         Left            =   120
         Picture         =   "frmRelFornProd.frx":1544
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   1680
         Picture         =   "frmRelFornProd.frx":1973
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1455
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmRelFornProd.frx":1ABD
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   10821
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRelFornProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdCat_Click()
    Dim c As categoria.categoria
    
    If c.codigo <> "" Then
        txtCat.Text = c.nome
        txtCat.tag = c.codigo
        busca
    End If
    
End Sub

Private Sub cmdMarca_Click()
    buscaMarca
End Sub
Sub buscaMarca()
    Dim m As Marca.Marca
    If m.codigo <> "" Then
        txtMarca.Text = m.nome
        txtMarca.tag = m.codigo
        busca
    End If
    
    
    
End Sub

Private Sub cmdPesquisar_Click()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    
    strSQL = "select p.cod_pro,p.codfab_pro,p.nome_pro, m.nome_mar, c.nome_cat, sum(i.qcom) as qtd, sum(vprod) as total "
    strSQL = strSQL & " from nfe as n inner join nfe_item as i on n.id=i.id_nfe inner join " & TAB_PRODUTOS & " as p "
    strSQL = strSQL & " on p.cod_pro=i.cprod "
    strSQL = strSQL & " inner join " & TAB_MARCAS & " as m on m.cod_mar=p.cod_mar "
    strSQL = strSQL & " inner join " & TAB_CATEGORIAS & " as c on c.cod_CAT=p.cod_CAT"
    strSQL = strSQL & " where 1>0"
    If txtDe.Text <> "" Then
        strSQL = strSQL & " and left(dhemi,10)>=" & F_GravaData(txtDe.Text)
    End If
    If txtAte.Text <> "" Then
        strSQL = strSQL & " and left(dhemi,10)<=" & F_GravaData(txtAte.Text)
    End If
    
    abreObjetoAdo ado, strSQL, dtgrd
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    txtDe.Text = F_PrimDiaMes
    txtAte.Text = Date
    
    
    
End Sub

Private Sub txtAte_GotFocus()
    P_SelecTexto txtAte
End Sub

Private Sub txtAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAte_LostFocus()
    P_AjustaData txtAte
End Sub

Private Sub txtDe_GotFocus()
    P_SelecTexto txtDe
End Sub

Private Sub txtDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDe_LostFocus()
    P_AjustaData txtDe
End Sub
