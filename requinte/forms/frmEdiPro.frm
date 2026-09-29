VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiPro 
   Caption         =   "Edição de Produtos"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   660
   ClientWidth     =   18825
   Icon            =   "frmEdiPro.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   18825
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pcBotoes 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   0
      ScaleHeight     =   885
      ScaleWidth      =   18825
      TabIndex        =   16
      Top             =   7665
      Width           =   18825
      Begin VB.CommandButton cmdEtiquetaPromocional 
         Caption         =   "Etiqueta Prom."
         Height          =   615
         Left            =   7320
         Picture         =   "frmEdiPro.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdNCM 
         Caption         =   "NCM"
         Height          =   615
         Left            =   13080
         Picture         =   "frmEdiPro.frx":04FA
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdCest 
         Caption         =   "CEST"
         Height          =   615
         Left            =   14640
         Picture         =   "frmEdiPro.frx":0644
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "Exportar"
         Height          =   615
         Left            =   11640
         Picture         =   "frmEdiPro.frx":078E
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Lista de Preços"
         Height          =   615
         Left            =   4440
         Picture         =   "frmEdiPro.frx":0B18
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnHist 
         Caption         =   "&Histórico"
         Height          =   615
         Left            =   10200
         Picture         =   "frmEdiPro.frx":0F25
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnDelPro 
         Caption         =   "E&xcluir Produto"
         Height          =   615
         Left            =   8760
         Picture         =   "frmEdiPro.frx":1927
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnCodBarra 
         Caption         =   "Emitir C.B&arra"
         Height          =   615
         Left            =   5880
         Picture         =   "frmEdiPro.frx":1CB1
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnLista 
         Caption         =   "Impr &HTML"
         Height          =   615
         Left            =   3000
         Picture         =   "frmEdiPro.frx":203B
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "Fe&char"
         Height          =   615
         Left            =   16200
         Picture         =   "frmEdiPro.frx":2448
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   1560
         Picture         =   "frmEdiPro.frx":25B3
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiPro.frx":2FB5
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox pcBusca 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   18795
      TabIndex        =   15
      Top             =   6690
      Width           =   18825
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   12120
         ScaleHeight     =   975
         ScaleWidth      =   2775
         TabIndex        =   31
         Top             =   0
         Width           =   2775
         Begin VB.CheckBox chkSemNcm 
            Caption         =   "Sem NCM"
            Height          =   255
            Left            =   1440
            TabIndex        =   40
            Tag             =   "!"
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox chkEst 
            Caption         =   "Vis. Est"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Tag             =   "!"
            Top             =   240
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chkNaoQueima 
            Caption         =   "Não queima"
            Height          =   255
            Left            =   1440
            TabIndex        =   38
            Tag             =   "!"
            Top             =   0
            Width           =   1335
         End
         Begin VB.CheckBox chkQueima 
            Caption         =   "Queima"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Tag             =   "!"
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox chkTodas 
            Caption         =   "Listar &Todas"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Tag             =   "!"
            Top             =   0
            Value           =   1  'Checked
            Width           =   1335
         End
      End
      Begin VB.PictureBox pcEstCons 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   14880
         ScaleHeight     =   975
         ScaleWidth      =   1575
         TabIndex        =   27
         Top             =   0
         Width           =   1575
         Begin VB.CommandButton cmndbtnEstCons 
            Height          =   255
            Left            =   1200
            Picture         =   "frmEdiPro.frx":333F
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   120
            Width           =   255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Est. Cons:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   30
            Top             =   120
            Width           =   930
         End
         Begin VB.Label lblEstCons 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   600
            TabIndex        =   29
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.TextBox txtbxSubSec 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         MaxLength       =   40
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtbxSecao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   40
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtbxFab 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         MaxLength       =   40
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox cmbxCriterio 
         Height          =   315
         ItemData        =   "frmEdiPro.frx":36C9
         Left            =   10440
         List            =   "frmEdiPro.frx":36E2
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   345
         Width           =   1575
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiPro.frx":3746
         Left            =   8880
         List            =   "frmEdiPro.frx":3756
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSubSec 
         Height          =   285
         Left            =   6360
         Picture         =   "frmEdiPro.frx":3783
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnSecao 
         Height          =   285
         Left            =   4080
         Picture         =   "frmEdiPro.frx":3B0D
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnFabr 
         Height          =   285
         Left            =   8520
         Picture         =   "frmEdiPro.frx":3E97
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   14760
         TabIndex        =   26
         Tag             =   "!"
         Top             =   0
         Width           =   195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricante/Marca:"
         Height          =   195
         Left            =   6720
         TabIndex        =   8
         Top             =   120
         Width           =   1320
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Critério:"
         Height          =   195
         Left            =   10440
         TabIndex        =   13
         Top             =   120
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&xpressão:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   8880
         TabIndex        =   11
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblSecCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Seção:"
         Height          =   195
         Left            =   2160
         TabIndex        =   2
         Top             =   120
         Width           =   510
      End
      Begin VB.Label lblSubCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Subseção:"
         Height          =   195
         Left            =   4440
         TabIndex        =   5
         Top             =   120
         Width           =   765
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiPro.frx":4221
      Height          =   5295
      Left            =   120
      TabIndex        =   25
      Top             =   240
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   9340
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
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
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "sinal"
         Caption         =   "Marcacao"
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
         DataField       =   "cod_pro"
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
      BeginProperty Column02 
         DataField       =   "nome_pro"
         Caption         =   "Nome"
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
         DataField       =   "pv_pro"
         Caption         =   "Preço"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "er_pro"
         Caption         =   "Estoque"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "r_pro"
         Caption         =   "Reserv."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "estoque"
         Caption         =   "Previsto"
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
      BeginProperty Column07 
         DataField       =   "nome_sec"
         Caption         =   "Seção"
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
      BeginProperty Column08 
         DataField       =   "nome_sub"
         Caption         =   "Subseção"
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
      BeginProperty Column09 
         DataField       =   "pq_pro"
         Caption         =   "queima"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "ncm_pro"
         Caption         =   "NCM"
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
      BeginProperty Column11 
         DataField       =   "icms_pro"
         Caption         =   "ICMS"
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
            ColumnWidth     =   255,118
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3435,024
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1755,213
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1124,787
         EndProperty
         BeginProperty Column11 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnOpt 
      Caption         =   "Opções"
      Begin VB.Menu mnExpor 
         Caption         =   "Exportar para XML"
         Begin VB.Menu mnExportLista 
            Caption         =   "Exportar produtos listados "
         End
         Begin VB.Menu mnExporProdSel 
            Caption         =   "Exportar apenas produto selecionado"
         End
         Begin VB.Menu mnExporProd 
            Caption         =   "Exportar todos os produtos"
         End
      End
      Begin VB.Menu mnInternet 
         Caption         =   "Internet"
         Begin VB.Menu mnConsultar 
            Caption         =   "Consultar"
         End
         Begin VB.Menu mnMerge 
            Caption         =   "Enviar para Internet"
         End
         Begin VB.Menu mnAtualizar 
            Caption         =   "Atualizar da Internet"
         End
         Begin VB.Menu mnInserirFila 
            Caption         =   "Inserir todos na fila de transferência"
         End
      End
   End
End
Attribute VB_Name = "frmEdiPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim bActive As Boolean

Sub BuscaUltEntrada()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select num_not,entrada_not,tf.nome_for,ti.* "
    strSQL = strSQL & " from tab_cmp as tc, tab_itcmp as ti, "
    strSQL = strSQL & " tab_fornecedores as tf "
    strSQL = strSQL & " where tf.cod_for=tc.cod_for "
    strSQL = strSQL & " and tc.seq_not=ti.seq_not "
    strSQL = strSQL & " and ti.cod_pro='" & ado!cod_pro & "' "
    strSQL = strSQL & " order by entrada_not desc limit 1 "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then
            lblUltComp.caption = ""
            lblNomeFor.caption = ""
            lblNumNot.caption = ""
            lblQtd.caption = ""
            lblUnit.caption = ""
            lblEmb.caption = ""
        Else
            lblUltComp.caption = !entrada_not
            lblNomeFor.caption = !nome_for
            lblNumNot.caption = !Num_not
            lblQtd.caption = !qtd_ite
            lblUnit.caption = FormatVal(!unit_ite)
            lblEmb.caption = !qtd_emb
        End If
        .Close
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Sub P_TileMsg(msg As String)
    lblMsg.caption = msg
End Sub

Private Sub ado_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    buscaEstCons
End Sub

Private Sub chkEst_Click()
    If chkEst.Value = 1 Then
        dtgrd.Columns(3).Visible = True
    Else
        dtgrd.Columns(3).Visible = False
    End If
End Sub

Private Sub chkNaoQueima_Click()
    busca
End Sub

Private Sub chkQueima_Click()
    busca
End Sub

Private Sub chkSemNcm_Click()
    busca
End Sub

Private Sub chkTodas_Click()
    adc_Reposition
End Sub

Private Sub cmbxBusca_Click()
    'If Me.Visible = True Then busca
End Sub

Private Sub cmbxDiv_Click()
    lblDiv.caption = F_BuscaUmCampo("Cod_div", "tab_divisoes", "nome_div='" & cmbxDiv.text & "'")
    busca
End Sub

Private Sub cmbxDiv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxDiv.ListIndex = -1
    End If
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxCat.ListIndex = -1
    End If
End Sub


Private Sub cmbxLoja_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxCriterio_Click()
    busca
End Sub

Private Sub cmbxPri_Click()
    lblPri.caption = F_BuscaUmCampo("Cod_pri", "tab_principios", "nome_pri='" & cmbxPri.text & "'")
    busca
End Sub

Private Sub cmbxEsp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxEsp.ListIndex = -1
    End If
End Sub

Private Sub cmbxSec_Click()
    lblSec.caption = F_BuscaUmCampo("Cod_sec", "tab_secoes", "nome_sec='" & cmbxSec.text & "'")
    
    busca
    
End Sub

Private Sub cmbxGenero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxGenero.ListIndex = -1
    End If
End Sub

Private Sub cmbxMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxMarca.ListIndex = -1
    End If
End Sub

Private Sub cmbxOrd_Click()
    'If Me.Visible = True Then busca
End Sub

Private Sub cmbxOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxSub_Click()
    busca
End Sub

Private Sub cmdCest_Click()
    Dim sNCM As String
    On Error GoTo trata_Erro
    Dim strSQL As String
    If strAltprod_usu <> "S" Then
        MsgBox "Você não tem autorização para essa tarefa.", vbExclamation
        Exit Sub
    End If
    
    
    sNCM = InputBox("Informe o número CEST para todos os produtos listados sem pontos.(apenas dígitos)", "NCM")
    If sNCM = "" Then
        MsgBox "Não foi informado um CEST válido!", vbExclamation
        Exit Sub
    End If
    
    With ado
        If MsgBox("Confirma alteração do código CEST: " & sNCM & " para os " & .RecordCount & " produtos selecionados?", vbYesNo + vbQuestion) = vbYes Then
            .MoveFirst
        
            Do While Not .EOF
                strSQL = "update " & TAB_PRODUTOS & " set cest_pro='" & sNCM & "' "
                strSQL = strSQL & " where cod_pro='" & !cod_pro & "'"
                
                
                Conn.Execute strSQL
                
                .MoveNext
            Loop
            .Requery
            
        End If
        
        
    End With
    
    Exit Sub
trata_Erro:
    P_MostraErro


End Sub

Private Sub cmdEtiquetaPromocional_Click()
    On Error GoTo trata_Erro
    ImprEtiquetaPromocional ado!cod_pro
    Exit Sub
trata_Erro:
    P_MostraErro
    
End Sub

Private Sub cmdExport_Click()
    ExportaGridToCSV dtgrd, ado
End Sub

Private Sub cmdNCM_Click()
    If strAltprod_usu <> "S" Then
        MsgBox "Você não tem autorização para essa tarefa.", vbExclamation
        Exit Sub
    End If
    
        Dim sNCM As String
    On Error GoTo trata_Erro
    Dim strSQL As String
    sNCM = InputBox("Informe o número NCM para todos os produtos listados sem pontos.(apenas dígitos)", "NCM")
    If sNCM = "" Then
        MsgBox "Não foi informado um NCM válido!", vbExclamation
        Exit Sub
    End If
    
    With ado
        P_Status "Aguarde"
        If MsgBox("Confirma alteração do código NCM: " & sNCM & " para os " & .RecordCount & " produtos selecionados?", vbYesNo + vbQuestion) = vbYes Then
            .MoveFirst
            dtgrd.Visible = False
        bActive = False
            Do While Not .EOF
                strSQL = "update " & TAB_PRODUTOS & " set ncm_pro='" & sNCM & "' "
                strSQL = strSQL & " where cod_pro='" & !cod_pro & "'"
                If !ncm_pro <> "''" Then
                
                    Conn.Execute strSQL
                End If
                .MoveNext
            Loop
            .Requery
            
        End If
        
        
    End With
    P_Status ""
    dtgrd.Visible = True
    bActive = True
    Exit Sub
    
trata_Erro:
    P_Status ""
    P_MostraErro
    dtgrd.Visible = True
    bActive = True
    
End Sub

Private Sub cmndbtnCodBarra_Click()
    On Error Resume Next
    With ado
        frmGeraEtiquetas.Show 0
        frmGeraEtiquetas.txtbxExpr.text = !nome_pro
        frmGeraEtiquetas.txtbxExpr.tag = !cod_pro
        frmGeraEtiquetas.Show 0
        frmGeraEtiquetas.SetFocus
    End With
End Sub

Private Sub cmndbtnDelPro_Click()
    Dim strSQL As String
    If strCadprod_usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        Exit Sub
    End If
    
    'if msgbox("Você deseja excluir
    
    With ado
        If MsgBox("Confirma exclusão do Produto?", vbYesNo + vbCritical) = vbNo Then
            Exit Sub
        End If
        If .RecordCount = 0 Then
            MsgBox "Não há o que ser excluído. Selecione um produto!", vbExclamation
            Exit Sub
        End If
        sAux = F_BuscaUmCampo("cod_pro", "tab_itcmp", "cod_pro = '" & !cod_pro & "'")
        If sAux <> "" Then
            MsgBox "Existem compras com esse produto!" & Chr(13) & "Não será possível excluir o produto selecionado.", vbCritical
            Exit Sub
        End If
        sAux = F_BuscaUmCampo("cod_pro", "tab_itcompras", "cod_pro = '" & !cod_pro & "'")
        If sAux <> "" Then
            MsgBox "Existem compras com esse produto!" & Chr(13) & "Não será possível excluir o produto selecionado.", vbCritical
            Exit Sub
        End If
        sAux = F_BuscaUmCampo("cod_pro", "tab_itvendas", "cod_pro = '" & !cod_pro & "'")
        If sAux <> "" Then
            MsgBox "Existem vendas com esse produto!" & Chr(13) & "Não será possível excluir o produto selecionado.", vbCritical
            Exit Sub
        End If
        sAux = F_BuscaUmCampo("cod_pro", "tab_itsai", "cod_pro = '" & !cod_pro & "'")
        If sAux <> "" Then
            MsgBox "Existem vendas com esse produto!" & Chr(13) & "Não será possível excluir o produto selecionado.", vbCritical
            Exit Sub
        End If
        sAux = F_BuscaUmCampo("er_pro", "tab_estoque", "cod_pro = '" & !cod_pro & "'")
        If sAux > 0 And sAux <> "" Then
            MsgBox "Existe estoque com esse produto!" & Chr(13) & "Não será possível excluir o produto selecionado.", vbCritical
            Exit Sub
        End If
        sAux = F_BuscaUmCampo("ef_pro", "tab_estoque", "cod_pro = '" & !cod_pro & "'")
        If sAux <> "0" And sAux <> "" Then
            MsgBox "Existe estoque com esse produto!" & Chr(13) & "Não será possível excluir o produto selecionado.", vbCritical
            Exit Sub
        End If
        strSQL = "delete from tab_produtos where cod_pro = '" & !cod_pro & "'"
        Conn.Execute strSQL
        'strSQL = "delete from tab_estoque where cod_pro = '" & !cod_pro & "'"
        'Conn.Execute strSQL
        'strSQL = "delete from tab_codbarras where cod_pro = '" & !cod_pro & "'"
        'conn.execute strSQL
    End With
    adc.Refresh
    MsgBox "Produto Excluído!", vbExclamation
End Sub

Private Sub cmndbtnEdit_Click()
    On Error Resume Next
    
    Dim fr As New frmCadPro
    
    'If strCadprod_usu <> "S" Then
    '    MsgBox "Você não tem permissão para esta operação!", vbExclamation
    '    Exit Sub
    'End If
    If dtgrd.row = -1 Then
    Exit Sub
    End If
    
    With ado
        If .RecordCount = 0 Then
            MsgBox "Não há o que ser Editado!", vbExclamation
            Exit Sub
        End If
        fr.tag = !cod_pro
        
        fr.Show 0, mdiFrm
        
    End With
End Sub


Private Sub cmndbtnEstCons_Click()
'    On Error GoTo trata_Erro
    P_ProdutosPorRepresentante ado!cod_pro, ""
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnFabr_Click()
    buscaFor
End Sub

Private Sub cmndbtnHist_Click()
    On Error Resume Next
    P_HistProd ado!cod_pro
End Sub

Private Sub cmndbtnImpr_Click()
    imprimeListagemPrecos ado
End Sub

Private Sub cmndbtnLista_Click()
bActive = False
P_DtgrdToHtml "LISTA DE PREÇOS EM " & Date, dtgrd, ado
Exit Sub

If dtgrd.row = -1 Then
        MsgBox "Não há itens na sua pesquisa!", vbExclamation
        bActive = True
        Exit Sub
    End If
    With ado
        
        If .RecordCount = 0 Then
            MsgBox "Não há itens na sua pesquisa!", vbExclamation
            bActive = True
            Exit Sub
        End If
        P_ImprimeListagem ado, frmVisualiza.Pic
        frmVisualiza.Show 1
        If frmVisualiza.tag <> "" Then
            P_ImprimeListagem ado, Printer
            Printer.EndDoc
        End If
        Unload frmVisualiza
        .MoveFirst
    End With
    P_Status ""
    bActive = True
    
End Sub

Private Sub cmndbtnNew_Click()
    If strCadprod_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
        Exit Sub
    End If
    Dim fr As New frmCadPro
    fr.Show 0
        
End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    Unload Me
End Sub


Private Sub cmndbtnSecao_Click()
    Dim s As secao.secao
    s = F_BuscaSec(txtbxSecao.text)
    If s.cod_sec <> "" Then
        txtbxSecao.text = s.nome_sec
        txtbxSecao.tag = s.cod_sec
        txtbxSubSec.text = ""
        txtbxSubSec.tag = ""
        busca
        txtbxSubSec.SetFocus
    End If
    
    
    
End Sub

Private Sub cmndbtnSubSec_Click()
    Dim s As subsec
    s = F_BuscaSubSec(txtbxSecao.tag, txtbxSubSec.text)
    If s.cod_sub <> "" Then
        txtbxSubSec.text = s.nome_sub
        txtbxSubSec.tag = s.cod_sub
        busca
        txtbxFab.SetFocus
    End If
End Sub

Private Sub dtgrd_DblClick()
    
    cmndbtnEdit_Click
End Sub

Private Sub dtgrd_GotFocus()
    
    dtgrd.MarqueeStyle = 3
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnEdit_Click
    ElseIf KeyCode = vbKeyC Then
        dtgrd.col = 1
        Clipboard.SetText dtgrd.text
    ElseIf KeyCode = vbKeyB Or KeyCode = vbKeyN Or KeyCode = vbKeyM Or _
                KeyCode = vbKeyG Or KeyCode = vbKeyF Then
        
            
            On Error Resume Next
            frmBusca.Show 1
            If frmBusca.tag <> "" Then
                If KeyCode = vbKeyB Then
                    cmbxBusca.ListIndex = 1
                End If
                If KeyCode = vbKeyN Then
                    cmbxBusca.ListIndex = 0
                End If
                If KeyCode = vbKeyF Then
                    cmbxBusca.ListIndex = 2
                End If
                If KeyCode = vbKeyM Then
                    cmbxBusca.ListIndex = 3
                End If
                If KeyCode = vbKeyG Then
                    cmbxBusca.ListIndex = 4
                End If
                txtbxExpr.text = frmBusca.txtbxExpr.text
            End If
            Unload frmBusca
        
    End If
    If KeyCode = vbKeyI And Shift = 1 Then
        If strViscusto_usu <> "S" Then
            MsgBox "Você não tem permissão para isto!", vbCritical
            Exit Sub
        End If
        P_InfPro ado!cod_pro
    End If
    If KeyCode = vbKeyC And Shift = vbCtrlMask Then
        Clipboard.SetText dtgrd.text
    End If
End Sub


Private Sub dtgrd_LostFocus()
   dtgrd.MarqueeStyle = 0
End Sub

Private Sub adc_Reposition()
    On Error Resume Next
    If bActive = False Then Exit Sub
    With ado
        If .RecordCount > 0 Then
            buscaEstoque
            lblEst.caption = FormatVal(!er_pro)
            lblPreco.caption = FormatVal(!pv_pro)
            'lblEstRes.caption = FormatVal(!r_pro)
            lblMarca.caption = !nome_pri
            lblCategoria.caption = !nome_sec
            lblSub.caption = !nome_sub
            txtbxST.text = !s_pro
            lblNomeLab.caption = F_BuscaUmCampo("nome_for", "tab_fornecedores", "cod_for='" & !cod_fab & "'")
                'P_PreencheBarCode !cod_pro, lstBarCode
            'BuscaUltEntrada
            buscaEstCons
        Else
            
        End If
    End With
End Sub
Sub buscaEstCons()
        
        
    On Error GoTo trata_Erro
    
    If bActive = False Then Exit Sub
    
    lblEstCons.caption = F_BuscaUmCampo("sum(reman_con)", "tab_consignacoes", "cod_pro='" & ado!cod_pro & "' ")
    
    Exit Sub
trata_Erro:
    lblEstCons.caption = "Err"
    
End Sub

Private Sub dtgrd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnOpt
    End If
End Sub

Private Sub Form_Activate()
    
    If bActive Then busca
    
    
        
    txtbxExpr.SetFocus
    
    If strViscusto_usu = "S" Then
        chkEst.Enabled = True
        chkEst_Click
    Else
        chkEst.Enabled = False
    End If
    txtbxExpr.SetFocus
    
    
    bActive = True
    
End Sub

Private Sub Form_Load()
        
    
    dtgrd.Columns(3).Visible = False
    P_CentraForm Me
    'adc.ConnectionString = Conn.ConnectionString
   ' dtEst.Connect = conn.connectionstring
    
    Set ado = New ADODB.Recordset
    

    cmbxOrd.ListIndex = 0
    
    
    
    resizeForm frmEdiPro
    
    
    
    
    'Busca
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnExpor
    End If
End Sub

Private Sub Form_Resize()
   posicionaControles
   
End Sub
Sub posicionaControles()
    On Error Resume Next
    
   ' dtgrd.width = Me.width - dtgrd.Left - 300
   ' dtgrd.height = pcBusca.Top - dtgrd.Top - 100
    PosicionaGrade dtgrd, Me, pcBusca
    posicionaADireita cmndbtnSair, pcBusca
    
    
End Sub
Sub buscaEstoque()
    Dim strSQL As String
    On Error Resume Next
    strSQL = "select * from tab_estoque where cod_pro='" & ado!cod_pro & "' "
    
    If chkTodas.Value = 0 Then
        strSQL = strSQL & " and (er_pro > 0 or ef_pro>0) "
    End If
    
    
    
    strSQL = strSQL & " order by venc_pro"
    dtEst.RecordSource = strSQL
    dtEst.Refresh
End Sub

Private Sub mnAtualizar_Click()
    On Error GoTo trata_Erro:
    With ado
        dtgrd.Visible = False
        Do While Not .EOF
            ApiAtualizarProdutos ado!cod_pro
            
            .MoveNext
        Loop
    End With
    dtgrd.Visible = True
trata_Erro:
dtgrd.Visible = True
End Sub

Private Sub mnConsultar_Click()
    On Error GoTo trata_Erro
    'APIBus ado!cod_pro
trata_Erro:
End Sub

Private Sub mnExporProd_Click()
    If strVisaoEstoque_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
    XMLBD.GeraXmlTabela "tab_produtos", "select * from ", ""
End Sub

Private Sub mnExporProdSel_Click()
    On erro GoTo trata_Erro
    
    If strVisaoEstoque_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
    
    XMLBD.GeraXmlTabela "tab_produtos", "select * from ", "where cod_pro='" & ado!cod_pro & "'"
    Exit Sub
trata_Erro:
End Sub

Private Sub mnExportLista_Click()
    Dim Parametros As String
    Parametros = "cnpj=" & aspas & F_SoNumero(strCNPJ_Emp) & aspas & " produtos=" & aspas & ado.RecordCount & aspas
    XMLUtil.ExportaGridToXML dtgrd, ado, Parametros
    
End Sub

Private Sub mnInserirFila_Click()
     On Error GoTo trata_Erro
    dtgrd.Visible = False
    With ado
        .MoveFirst
        Do While Not .EOF
            
            insere_fila !cod_pro, 3
            .MoveNext
        Loop
        .MoveFirst
    End With
    P_Status
    dtgrd.Visible = True
trata_Erro:
End Sub

Private Sub mnMerge_Click()
    On Error GoTo trata_Erro
    dtgrd.Visible = True
    With ado
        Do While Not .EOF
            
            APIMergeProduto ado!cod_pro
            P_Status ado!cod_pro
            .MoveNext
        Loop
        
    End With
    P_Status
    dtgrd.Visible = True
trata_Erro:
End Sub

Private Sub pcBotoes_GotFocus()
    Sendkeys Chr(9)
End Sub

Private Sub pcBusca_GotFocus()
    Sendkeys Chr(9)
End Sub

Sub busca(Optional kw As String = "")
    Dim strSQL As String
    Dim sOrder As String
    Dim sCampo As String
    Dim k1 As String
    Dim k2 As String
    Dim i As Integer
    Dim c As New Collection
    Dim c2 As New Collection
    On Error GoTo trata_Erro
    
    
    
    'txtbxExpr.Text = Trim(txtbxExpr.Text)
    'adc.ConnectionString = Conn.ConnectionString
    
    
    
    
    
    If InStr(1, txtbxExpr.text, "{", vbTextCompare) = 1 _
        And InStr(1, txtbxExpr.text, "}", vbTextCompare) = 1 _
        And InStr(1, txtbxExpr.text, ";", vbTextCompare) > 0 Then
    
        Set c = F_SeparaTexto(txtbxExpr.text, ";", False, "", "")
        If c.Count > 0 Then
'
        End If
        
    End If
    
    
    
    If cmbxOrd.ListIndex = 1 Then
        sOrder = "tp.cod_pro,tp.nome_pro"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrder = "tp.ref_pro,tp.nome_pro,tp.cod_pro"
    ElseIf cmbxOrd.ListIndex = 3 Then
        sOrder = "tp.endereco_pro,tp.nome_pro,tp.cod_pro"
    Else
        sOrder = "tp.nome_pro,tp.cod_pro "
    End If
    strSQL = "select distinct ts.nome_sub, tse.nome_sec,"
    strSQL = strSQL & " tp.*,er_pro+r_pro as estoque,((tp.pv_pro/tp.cr_pro-1)*100)-" & F_GravaDec(F_TaxasEntradas) & "  as liq, "
    strSQL = strSQL & " ((tp.pv_pro/tp.cr_pro-1)*100) as lb, "
    strSQL = strSQL & " (((tp.pv_pro/tp.cr_pro-1)*100)-" & F_GravaDec(F_TaxasEntradas) & ") * pv_pro/100  as lr, "
    strSQL = strSQL & " ((tp.pv_pro/tp.cr_pro-1)*100) * pv_pro/100  as bruto "
    strSQL = strSQL & " from tab_produtos as tp left join tab_subsec as ts on tp.cod_sub=ts.cod_sub "
    strSQL = strSQL & " left join tab_secoes as tse on ts.cod_sec=tse.cod_sec "
    strSQL = strSQL & " left join  tab_codbarras as tb  on tb.cod_pro=tp.cod_pro "
    
    strSQL = strSQL & " where (nome_pro like '%"
    strSQL = strSQL & txtbxExpr.text & "%'  "
    If F_SoNumero(txtbxExpr.text) = txtbxExpr.text And txtbxExpr.text <> "" Then
        strSQL = strSQL & " or ref_pro =" & txtbxExpr.text
        strSQL = strSQL & " or ncm_pro =" & txtbxExpr.text
    End If
    strSQL = strSQL & " or codbar_pro like '" & txtbxExpr.text & "%' "
    strSQL = strSQL & " or ean_pro like '" & txtbxExpr.text & "%' "
    strSQL = strSQL & " or codbarra_pro like '" & txtbxExpr.text & "%' "
    
    strSQL = strSQL & " or endereco_pro like '%" & txtbxExpr.text & "%' "
    strSQL = strSQL & " or tp.cod_pro like '" & txtbxExpr.text & "%'"
    strSQL = strSQL & " or tp.ncm_pro like '" & txtbxExpr.text & "%'"
    strSQL = strSQL & " or tb.cod_bar='" & txtbxExpr.text & "') "
    
     
    If chkSemNcm.Value Then
        strSQL = strSQL & " and (tp.ncm_pro='' or length(tp.ncm_pro) <> 8) "
    End If
    If txtbxFab.tag <> "" Then
        strSQL = strSQL & " and tp.cod_fab='" & txtbxFab.tag & "' "
    End If
    If chkQueima.Value = 1 Then
        strSQL = strSQL & " and pq_pro>0 "
    Else
        If chkNaoQueima.Value = 1 Then
            strSQL = strSQL & " and pq_pro=0 "
        End If
    End If
    
     
    If chkTodas.Value = 0 Then
        strSQL = strSQL & " and er_pro >0"
    End If
    
    If txtbxSecao.tag <> "" Then
        strSQL = strSQL & " and ts.cod_sec='" & txtbxSecao.tag & "'  "
    End If
    If txtbxSubSec.tag <> "" Then
        strSQL = strSQL & " and ts.cod_sub='" & txtbxSubSec.tag & "'  "
    End If
    
    If cmbxCriterio.ListIndex = 1 Then
        strSQL = strSQL & " and ef_pro<er_pro"
    ElseIf cmbxCriterio.ListIndex = 2 Then
        strSQL = strSQL & " and ef_pro>er_pro"
    ElseIf cmbxCriterio.ListIndex = 3 Then
        strSQL = strSQL & " and ef_pro<0"
    ElseIf cmbxCriterio.ListIndex = 4 Then
        strSQL = strSQL & " and er_pro<estmin_pro"
    ElseIf cmbxCriterio.ListIndex = 5 Then
        strSQL = strSQL & " and ef_pro>(er_pro*2)"
    ElseIf cmbxCriterio.ListIndex = 6 Then
        strSQL = strSQL & " and er_pro>(ef_pro*2)"
    End If
    
    
    strSQL = strSQL & " order by " & sOrder

    abreObjetoAdo ado, strSQL, dtgrd
    dtgrd.caption = "Registros: " & ado.RecordCount

    Set dtgrd.DataSource = ado
'    adc.RecordSource = strSQL
'    adc.Refresh
    'lblList.caption = "[Listadas..: " & ado.RecordCount & _
        "   Cadastradas: " & F_BuscaUmCampo("count(*)", "tab_produtos", "") & "]"
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub txtbxExpr_Change()
    busca
End Sub

Private Sub txtbxExpr_GotFocus()
    P_TileMsg "Informe algum dado do produto e tecle enter..."
    P_SelecTexto txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        busca
       If ado.RecordCount > 0 Then
            dtgrd.SetFocus
        End If
    ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dtgrd.SetFocus
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    P_AjustaTexto txtbxExpr
    P_TileMsg ""
    
End Sub

Private Sub txtbxFab_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaFor
    End If
End Sub
Sub buscaFor()
    Dim f As fornecedor.fornecedor
    f = fornecedor.F_SelectForn(txtbxFab.text)
    If f.cod_For <> "" Then
        txtbxFab.text = f.nome_for
        txtbxFab.tag = f.cod_For
        busca
    End If
    
End Sub

Private Sub txtbxSecao_Change()
    txtbxSecao.tag = ""
End Sub

Private Sub txtbxSecao_GotFocus()
    P_SelecTexto txtbxSecao
End Sub

Private Sub txtbxSecao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnSecao_Click
    End If
End Sub

Private Sub txtbxSubSec_Change()
    txtbxSubSec.tag = ""
End Sub

Private Sub txtbxSubsec_GotFocus()
    P_SelecTexto txtbxSubSec
End Sub

Private Sub txtbxSubsec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnSubSec_Click
    End If
End Sub
