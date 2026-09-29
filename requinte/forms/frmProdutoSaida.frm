VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmProdutoSaida 
   BorderStyle     =   0  'None
   Caption         =   "Selecione Um Produto"
   ClientHeight    =   9345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12645
   ControlBox      =   0   'False
   ForeColor       =   &H80000004&
   Icon            =   "frmProdutoSaida.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   12645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmProdutoSaida.frx":014A
      Height          =   6975
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   12303
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
      ColumnCount     =   10
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "pv_pro"
         Caption         =   "Preço"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "pq_pro"
         Caption         =   "Queima"
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
         DataField       =   "er_pro"
         Caption         =   "Estoque"
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
      BeginProperty Column08 
         DataField       =   "nome_for"
         Caption         =   "Fabricante"
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
         DataField       =   ""
         Caption         =   "codfab_pro"
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
            ColumnWidth     =   1110,047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4995,213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1230,236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1305,071
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   12615
      TabIndex        =   4
      Top             =   0
      Width           =   12645
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   0
         Picture         =   "frmProdutoSaida.frx":015C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   6
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione um Produto"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   600
         TabIndex        =   5
         Top             =   120
         Width           =   2775
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
      TabIndex        =   3
      Top             =   7635
      Width           =   12645
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
      Begin VB.TextBox txtbxBuscar 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pesquisar:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmProdutoSaida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bActive As Boolean
Public WithEvents adoR As ADODB.Recordset
Attribute adoR.VB_VarHelpID = -1
Private Sub chkTodas_Click()
    busca
End Sub

Private Sub adoR_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    'lblNomeFab.caption = F_BuscaUmCampo("nome_for", "tab_fornecedores", "cod_for='" & adoR!cod_fab & "'")
    'lblSec.caption = adoR!Nome_Sec
    'lblSub.caption = adoR!nome_sub
   '
    'buscaPrecos
    
End Sub
Sub buscaPrecos()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
  '  msflxgrdPrecos.rows = 1
    strSQL = "select * from tab_tipoven "
    abreObjetoAdo rPesq, strSQL
    'msflxgrdPrecos.Cols = 3
'   msflxgrdPrecos.ColAlignment(0) = 3
 '  msflxgrdPrecos.ColAlignment(1) = 3
    With rPesq
        Do While Not .EOF
        
            strSQL = !descricao_tip & Chr(9) & FormatVal(adoR!pv_pro * !fator_tip / !parcelas_tip)
         '   msflxgrdPrecos.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
  '  ZebraGrid msflxgrdPrecos, 1, &HB4DFE0
End Sub
Private Sub dtgrd_DblClick()
    On Error Resume Next
    With adoR
        If .RecordCount = 0 Then
            Exit Sub
        End If
        
        Me.tag = adoR!cod_pro
        Me.Hide
    End With
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    End If
    If KeyCode = vbKeyN Or KeyCode = vbKeyP Or KeyCode = vbKeyM Then
        frmBusca.tag = "N"
        frmBusca.Show 1
        txtbxBuscar.text = frmBusca.tag
        Unload frmBusca
        busca
    End If
End Sub


Private Sub Form_Activate()
    On Error Resume Next
    Form_Resize
    If bActive = True Then Exit Sub
    
    
    txtbxBuscar.text = Me.tag
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
    Set adoR = New ADODB.Recordset
    bActive = False
End Sub
Sub busca(Optional limit As String = "")
    Dim strSQL As String
    Dim sOrder As String
    Dim sCampo As String
    Dim iCount As Integer
    On Error GoTo trata_Erro
    txtbxBuscar.text = F_AjustaTexto(txtbxBuscar.text)
    txtbxBuscar.SelLength = 0
    txtbxBuscar.SelStart = Len(txtbxBuscar.text)
    
        strSQL = "select distinct er_pro as er, ts.nome_sec, tu.nome_sub,tu.comiss_sub,tu.descmax_sub, "
        
        strSQL = strSQL & "tp.* "
        strSQL = strSQL & " from tab_produtos as tp left join "
        strSQL = strSQL & " tab_subsec as tu on tp.cod_sub=tu.cod_sub "
        strSQL = strSQL & " left join tab_secoes as ts  on ts.cod_sec=tu.cod_sec "
        strSQL = strSQL & " left join tab_codbarras as tb on tb.cod_pro=tp.cod_pro "
        
        strSQL = strSQL & " where 1 > 0"
        
        
        
            strSQL = strSQL & " and ( codbar_PRO like '"
            strSQL = strSQL & txtbxBuscar.text & "' "
            strSQL = strSQL & " or ean_pro like '"
            strSQL = strSQL & txtbxBuscar.text & "' "
            strSQL = strSQL & " or codfab_pro like '"
            strSQL = strSQL & txtbxBuscar.text & "%' "
            
            strSQL = strSQL & " or codbarra_pro like '"
            strSQL = strSQL & txtbxBuscar.text & "' "
                strSQL = strSQL & " or tp.COD_PRO like '"
                strSQL = strSQL & txtbxBuscar.text & "%' "
            
        If F_SoNumero(txtbxBuscar.text) = txtbxBuscar.text And txtbxBuscar.text <> "" Then
            strSQL = strSQL & " or ref_pro =" & txtbxBuscar.text
        End If
        If F_SoNumero(txtbxBuscar.text) = txtbxBuscar.text And txtbxBuscar.text <> "" Then
            strSQL = strSQL & " or tp.cod_pro =" & txtbxBuscar.text
        End If
            strSQL = strSQL & " or NOME_PRO like '%"
            strSQL = strSQL & txtbxBuscar.text & "%') "
        
        
    
        
        strSQL = strSQL & " group by tp.cod_pro "
        strSQL = strSQL & " order by tp.nome_pro"
        
        abreObjetoAdo adoR, strSQL, dtgrd
        
        If adoR.RecordCount = 1 Then
            dtgrd_DblClick
            
        End If
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub Form_Resize()
    'dtgrd.height = pcDados.top - dtgrd.top - 100
    'msflxgrdPrecos.height = dtgrd.height
    'msflxgrdPrecos.Left = Me.width - msflxgrdPrecos.width - 100
    'dtgrd.width = msflxgrdPrecos.Left - dtgrd.Left
    PosicionaGrade dtgrd, Me, pCdados
End Sub

Private Sub Label16_Click()
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
