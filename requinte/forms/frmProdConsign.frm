VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmProdConsig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecione Um Produto"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12645
   ControlBox      =   0   'False
   ForeColor       =   &H80000004&
   Icon            =   "frmProdConsign.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   12645
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00F79971&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   12615
      TabIndex        =   12
      Top             =   0
      Width           =   12645
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         Picture         =   "frmProdConsign.frx":014A
         ScaleHeight     =   195
         ScaleWidth      =   180
         TabIndex        =   14
         Top             =   68
         Width           =   180
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione um Produto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   480
         TabIndex        =   13
         Top             =   0
         Width           =   3075
      End
   End
   Begin VB.PictureBox pcDados 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   12615
      TabIndex        =   5
      Top             =   7635
      Width           =   12645
      Begin VB.Label lblNomeRep 
         BackColor       =   &H00869865&
         Caption         =   "nome"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   16
         Top             =   720
         Width           =   15720
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Repres:"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   15
         Top             =   720
         Width           =   915
      End
      Begin VB.Label lblNomeFab 
         BackColor       =   &H007DBFDB&
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   11
         Top             =   480
         Width           =   15720
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H007DBFDB&
         Caption         =   "Fabricante:"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   0
         TabIndex        =   10
         Top             =   480
         Width           =   915
      End
      Begin VB.Label lblSec 
         BackColor       =   &H007DBFDB&
         Caption         =   "nome"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   9
         Top             =   0
         Width           =   15720
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H007DBFDB&
         Caption         =   "Seção"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   915
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "SubSeção:"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lblSub 
         BackColor       =   &H00C0FFFF&
         Caption         =   "nome"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   15720
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EF9C7C&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   12615
      TabIndex        =   0
      Top             =   8610
      Width           =   12645
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Selecionar"
         Height          =   375
         Left            =   11280
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbxOpcao 
         Height          =   315
         ItemData        =   "frmProdConsign.frx":0360
         Left            =   120
         List            =   "frmProdConsign.frx":036A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtbxBuscar 
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   7215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opções"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expressão:"
         Height          =   195
         Left            =   1920
         TabIndex        =   1
         Top             =   0
         Width           =   780
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmProdConsign.frx":0388
      Height          =   6495
      Left            =   0
      TabIndex        =   17
      Top             =   600
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   11456
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
      ColumnCount     =   14
      BeginProperty Column00 
         DataField       =   "ref_pro"
         Caption         =   "Referência"
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
         Caption         =   "Descrição"
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
         DataField       =   "reman_con"
         Caption         =   "Reman."
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
      BeginProperty Column04 
         DataField       =   "qtd_con"
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
      BeginProperty Column05 
         DataField       =   "reservado_con"
         Caption         =   "Reservado"
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
      BeginProperty Column06 
         DataField       =   "restante"
         Caption         =   "Restante"
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
         DataField       =   "unit_con"
         Caption         =   "Unit."
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
         DataField       =   "data_con"
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
      BeginProperty Column09 
         DataField       =   "endereco_pro"
         Caption         =   "Endereço"
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
      BeginProperty Column10 
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
      BeginProperty Column11 
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
      BeginProperty Column12 
         DataField       =   "seq_con"
         Caption         =   "Seq."
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
      BeginProperty Column13 
         DataField       =   "pq_pro"
         Caption         =   "Queima"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   2324,977
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1110,047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4995,213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1230,236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2160
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1484,787
         EndProperty
         BeginProperty Column13 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProdConsig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents adoprod As ADODB.Recordset
Attribute adoprod.VB_VarHelpID = -1
Dim bActive As Boolean

Private Sub chkTodas_Click()
    busca
End Sub

Private Sub cmdSelect_Click()
     With adoprod
        If .RecordCount = 0 Then
            Exit Sub
        End If
        Me.tag = !cod_pro
        Me.Hide
        
    End With
End Sub

Private Sub dtgrd_DblClick()
    On Error Resume Next
    With adoprod
        If .RecordCount = 0 Then
            Exit Sub
        End If
        cmdSelect_Click
        
    End With
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    End If
    If KeyCode = vbKeyN Or KeyCode = vbKeyP Or KeyCode = vbKeyM Then
        frmBusca.tag = "N"
        frmBusca.Show 1
        txtbxBuscar.Text = frmBusca.tag
        Unload frmBusca
        busca
    End If
End Sub

Private Sub adc_Reposition()
    On Error Resume Next
    With adoprod
        lblNomeFab.caption = F_BuscaUmCampo("nome_for", "tab_fornecedores", "cod_for='" & !cod_fab & "'")
        lblSub.caption = !nome_sub
        lblSec.caption = !Nome_Sec
        
        'P_BuscaPrecos !pv_pro, !nome_pro, msflxgrdPrecos
    End With
    
    
    
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If bActive = True Then Exit Sub
    
    Form_Resize
    txtbxBuscar.Text = Me.tag
    busca
    dtgrd.SetFocus
    dtgrd.MarqueeStyle = 3
    bActive = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    Set adoprod = New ADODB.Recordset
    bActive = False
End Sub
Sub busca1(Optional limit As String = "")
    Dim strSQL As String
    Dim sOrder As String
    Dim sCampo As String
    Dim iCount As Integer
    On Error GoTo trata_erro
    txtbxBuscar.Text = F_AjustaTexto(txtbxBuscar.Text)
    txtbxBuscar.SelLength = 0
    txtbxBuscar.SelStart = Len(txtbxBuscar.Text)
    
        strSQL = "select  er_pro-r_pro as er, ts.nome_sec, tu.nome_sub,tu.comiss_sub,tu.descmax_sub, "
        
        strSQL = strSQL & "tp.*,tc.*,sum(reman_con) as reman,sum(qtd_con) as qtd "
        strSQL = strSQL & " from tab_produtos as tp, "
        strSQL = strSQL & " tab_subsec as tu,tab_secoes as ts,tab_consignacoes as tc  "
        strSQL = strSQL & " where tp.cod_sub=tu.cod_sub "
        strSQL = strSQL & " and tu.cod_sec=ts.cod_sec "
        strSQL = strSQL & " and tc.cod_pro=tp.cod_pro "
        strSQL = strSQL & " and tc.cod_rep='" & lblNomeRep.tag & "' "
        strSQL = strSQL & " and reman_con>0"
        
            strSQL = strSQL & " and ( codbar_PRO like '"
            strSQL = strSQL & txtbxBuscar.Text & "' "
            
                strSQL = strSQL & " or tp.COD_PRO like '"
                strSQL = strSQL & txtbxBuscar.Text & "%' "
            
        If F_SoNumero(txtbxBuscar.Text) = txtbxBuscar.Text And txtbxBuscar.Text <> "" Then
            strSQL = strSQL & " or ref_pro =" & txtbxBuscar.Text
        End If
        
        
            strSQL = strSQL & " or NOME_PRO like '%"
            strSQL = strSQL & txtbxBuscar.Text & "%') "
         
        
        
    
        
        strSQL = strSQL & " group by tp.cod_pro "
        strSQL = strSQL & " order by tp.nome_pro"
        
        FechaObjeto adoprod
        adoprod.Open strSQL, Conn_Client
        Set dtgrd.DataSource = adoprod
        If adoprod.RecordCount = 1 Then
            dtgrd_DblClick
            
        End If
    
    Exit Sub
trata_erro:
    P_MostraErro
End Sub
Sub busca(Optional limit As String = "")
    Dim strSQL As String
    Dim sOrder As String
    Dim sCampo As String
    Dim iCount As Integer
    On Error GoTo trata_erro
    txtbxBuscar.Text = F_AjustaTexto(txtbxBuscar.Text)
    txtbxBuscar.SelLength = 0
    txtbxBuscar.SelStart = Len(txtbxBuscar.Text)
    
        strSQL = "select  reman_con-reservado_con as restante,"
        strSQL = strSQL & " er_pro-r_pro as er, ts.nome_sec, tu.nome_sub,tu.comiss_sub,tu.descmax_sub, "
        
        strSQL = strSQL & "tp.*,tc.* "
        strSQL = strSQL & " from tab_produtos as tp, "
        strSQL = strSQL & " tab_subsec as tu,tab_secoes as ts,tab_consignacoes as tc  "
        strSQL = strSQL & " where tp.cod_sub=tu.cod_sub "
        strSQL = strSQL & " and tu.cod_sec=ts.cod_sec "
        strSQL = strSQL & " and tc.cod_pro=tp.cod_pro "
        strSQL = strSQL & " and tc.cod_rep='" & lblNomeRep.tag & "' "
        strSQL = strSQL & " and reman_con>0"
        
            strSQL = strSQL & " and ( codbar_PRO like '"
            strSQL = strSQL & txtbxBuscar.Text & "' "
            
                strSQL = strSQL & " or tp.COD_PRO like '"
                strSQL = strSQL & txtbxBuscar.Text & "%' "
            
        If F_SoNumero(txtbxBuscar.Text) = txtbxBuscar.Text And txtbxBuscar.Text <> "" Then
            strSQL = strSQL & " or ref_pro =" & txtbxBuscar.Text
        End If
        
        
            strSQL = strSQL & " or NOME_PRO like '%"
            strSQL = strSQL & txtbxBuscar.Text & "%') "
         
        
        
    
        
        'strSQL = strSQL & " group by tp.cod_pro "
        strSQL = strSQL & " order by tp.nome_pro,data_con"
        
        abreObjetoAdo adoprod, strSQL
        
        Set dtgrd.DataSource = adoprod
        If adoprod.RecordCount = 1 Then
            dtgrd_DblClick
            
        End If
    
    Exit Sub
trata_erro:
    P_MostraErro
End Sub


Private Sub Form_Resize()
    dtgrd.height = pCdados.Top - dtgrd.Top - 100
    
    dtgrd.width = Me.width - dtgrd.Left
End Sub

Private Sub txtbxBuscar_Change()
    busca
End Sub

Private Sub txtbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dtgrd.SetFocus
    End If
    If KeyCode = vbKeyReturn Then
        busca
    End If
    
End Sub
