VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiCaiOld 
   Caption         =   "Edição de Fluxo de Caixa"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12900
   Icon            =   "frmEdiCaiOld.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   12900
   Begin VB.PictureBox pcLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   0
      ScaleHeight     =   7545
      ScaleWidth      =   2265
      TabIndex        =   8
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton cmndbtnVendasNaoCaixa 
         Caption         =   "Vendas Não Caixa"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   6720
         Width           =   2055
      End
      Begin VB.ComboBox cmbxTurno 
         Height          =   315
         ItemData        =   "frmEdiCaiOld.frx":014A
         Left            =   120
         List            =   "frmEdiCaiOld.frx":0157
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   3840
         Width           =   1935
      End
      Begin VB.ComboBox cmbxCai 
         Height          =   315
         ItemData        =   "frmEdiCaiOld.frx":0175
         Left            =   120
         List            =   "frmEdiCaiOld.frx":0182
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   3240
         Width           =   1935
      End
      Begin VB.OptionButton optnDesc 
         Caption         =   "Desc"
         Height          =   255
         Left            =   1320
         TabIndex        =   17
         Top             =   1560
         Width           =   735
      End
      Begin VB.OptionButton optnAsc 
         Caption         =   "Asc"
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   1320
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   15
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   13
         Top             =   2640
         Width           =   1695
      End
      Begin VB.ComboBox cmbxBusca 
         Height          =   315
         ItemData        =   "frmEdiCaiOld.frx":01A0
         Left            =   120
         List            =   "frmEdiCaiOld.frx":01AD
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiCaiOld.frx":01D4
         Left            =   120
         List            =   "frmEdiCaiOld.frx":01E4
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ComboBox cmbxTIpo 
         Height          =   315
         ItemData        =   "frmEdiCaiOld.frx":020C
         Left            =   120
         List            =   "frmEdiCaiOld.frx":0219
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox chkVales 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Vales"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   4200
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Height          =   285
         Left            =   1800
         Picture         =   "frmEdiCaiOld.frx":0237
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   5640
         Width           =   615
      End
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1425
         TabIndex        =   35
         Top             =   5640
         Width           =   390
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Crédito:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label lblCred_din 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1500
         TabIndex        =   33
         Top             =   4560
         Width           =   315
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Débito:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label lblDeb_Din 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1500
         TabIndex        =   31
         Top             =   4920
         Width           =   315
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   5280
         Width           =   615
      End
      Begin VB.Label lblTotal_Din 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1500
         TabIndex        =   29
         Top             =   5280
         Width           =   315
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno:"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   3600
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   3000
         Width           =   435
      End
      Begin VB.Label lbldataFin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   1200
         TabIndex        =   24
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label lblDataIni 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Buscar Por:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lblExpr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&xpressão:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   0
         Width           =   360
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   12870
      TabIndex        =   0
      Top             =   7575
      Width           =   12900
      Begin VB.CommandButton cmndbtnVale 
         Caption         =   "Impr. Vale"
         Height          =   615
         Left            =   8520
         Picture         =   "frmEdiCaiOld.frx":05C1
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   6840
         Picture         =   "frmEdiCaiOld.frx":09CE
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   10200
         Picture         =   "frmEdiCaiOld.frx":0DDB
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnFechaCAi 
         Caption         =   "Fec&hamento"
         Height          =   615
         Left            =   5160
         Picture         =   "frmEdiCaiOld.frx":0F46
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3480
         Picture         =   "frmEdiCaiOld.frx":109F
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   1800
         Picture         =   "frmEdiCaiOld.frx":11E9
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiCaiOld.frx":1333
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiCaiOld.frx":16BD
      Height          =   7095
      Left            =   2400
      TabIndex        =   38
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   12515
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "seq_cai"
         Caption         =   "Id.Cai"
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
         DataField       =   "data_cai"
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
      BeginProperty Column02 
         DataField       =   "valor_cai"
         Caption         =   "Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "historico_cai"
         Caption         =   "Histórico"
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
         DataField       =   "tipo_cai"
         Caption         =   "Tipo"
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
         DataField       =   "num_ven"
         Caption         =   "Venda"
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
         DataField       =   "esp_cai"
         Caption         =   "Esp"
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
         DataField       =   "desconto_cai"
         Caption         =   "Desconto"
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
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   4199,811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1184,882
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiCaiOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Private Sub Text1_Change()

End Sub

Private Sub ado_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    If ado!num_ven = "" Then
        cmndbtnEdit.Enabled = True
        cmndbtnExcl.Enabled = True
    Else
        cmndbtnEdit.Enabled = False
        cmndbtnExcl.Enabled = False
    End If
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    
    On Error Resume Next
    
    On Error GoTo trata_Erro
    
    
    If txtbxdataFin.Visible = True Then
        If F_DataMaior(txtbxDataIni.Text, txtbxdataFin.Text) Then
            MsgBox "A data final deve ser maior que a inicial!", vbExclamation
            txtbxdataFin.SetFocus
            Exit Sub
        End If
    End If
    strSQL = "select distinct tc.*,tb.nome_cai from " & TAB_CAIXA & " as tc "
    'strSQL = strSQL & " left join tab_baixafin as tbf on tbf.seq_fin=tc.seq_cai "
    'strSQL = strSQL & " inner join tab_baixaitd2 as tbi on tbi.seq_cab=tbf.seq_cab "
    strSQL = strSQL & ",  tab_caixas    as tb "
    
    
    
    If chkVales.Value = 1 Then
        strSQL = strSQL & " ,tab_valescai as tv "
        strSQL = strSQL & " where tv.seq_cai=tc.seq_cai and "
    Else
        strSQL = strSQL & " where "
    End If
    
    
    strSQL = strSQL & "tc.cod_cai=tb.cod_cai  and "
    If cmbxBusca.ListIndex = 0 Then
        strSQL = strSQL & " historico_cai like '" & txtbxExpr.Text & "%'"
    ElseIf cmbxBusca.ListIndex = 1 Then
        strSQL = strSQL & " num_ven like '" & txtbxExpr.Text & "%'"
    ElseIf cmbxBusca.ListIndex = 2 Then
        strSQL = strSQL & " valor_cai = " & F_GravaDec(txtbxExpr.Text)
    End If
    strSQL = strSQL & " and data_cai >= " & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and data_cai <= " & F_GravaData(txtbxdataFin.Text)
    If cmbxTIpo.ListIndex <> 0 Then
        strSQL = strSQL & " and tipo_cai='" & Left(cmbxTIpo.Text, 1) & "' "
    End If
    If cmbxCai.Text <> "" Then
        strSQL = strSQL & " and tc.cod_cai='" & Left(cmbxCai.Text, 2) & "' "
    End If
    If cmbxTurno.Text <> "" Then
        strSQL = strSQL & " and tc.cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
    End If
    If pcLeft.BackColor <> vbYellow Then
        strSQL = strSQL & " and visivel_cai='S'"
    
    
    End If
    
    If cmbxOrd.ListIndex = 0 Then
        strSQL = strSQL & " order by data_cai,seq_cai"
    ElseIf cmbxOrd.ListIndex = 1 Then
        strSQL = strSQL & " order by seq_cai"
    ElseIf cmbxOrd.ListIndex = 2 Then
        strSQL = strSQL & " order by historico_cai, seq_cai"
    ElseIf cmbxOrd.ListIndex = 3 Then
        strSQL = strSQL & " order by valor_cai,seq_cai"
    End If
        
    
    
    If optnAsc.Value Then
        strSQL = strSQL & " asc "
    Else
        strSQL = strSQL & " desc "
    End If
    'ado.RecordSource = strSQL
    'ado.Refresh
    
    abreObjetoAdo ado, strSQL, dtgrd
    
    
    
    P_Resumo
trata_Erro:
Exit Sub
P_MostraErro
End Sub
Sub buscaF8()
    Dim strSQL As String
    If txtbxdataFin.Visible = True Then
        If F_DataMaior(txtbxDataIni.Text, txtbxdataFin.Text) Then
            MsgBox "A data final deve ser maior que a inicial!", vbExclamation
            txtbxdataFin.SetFocus
            Exit Sub
        End If
    End If
    strSQL = "select distinct tc.*,tb.nome_cai from tab_caixa as tc, tab_caixas as tb "
    If chkVales.Value = 1 Then
        strSQL = strSQL & " ,tab_valescai as tv "
        strSQL = strSQL & " where tv.seq_cai=tc.seq_cai and "
    Else
        strSQL = strSQL & " where "
    End If
    
    
    strSQL = strSQL & "tc.cod_cai=tb.cod_cai  and "
    If cmbxBusca.ListIndex = 0 Then
        strSQL = strSQL & " historico_cai like '" & txtbxExpr.Text & "%'"
    ElseIf cmbxBusca.ListIndex = 1 Then
        strSQL = strSQL & " num_ven like '" & txtbxExpr.Text & "%'"
    ElseIf cmbxBusca.ListIndex = 2 Then
        strSQL = strSQL & " valor_cai = " & F_GravaDec(txtbxExpr.Text)
    End If
    strSQL = strSQL & " and data_cai >= " & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and data_cai <= " & F_GravaData(txtbxdataFin.Text)
    If cmbxTIpo.ListIndex <> 0 Then
        strSQL = strSQL & " and tipo_cai='" & Left(cmbxTIpo.Text, 1) & "' "
    End If
    If cmbxCai.Text <> "" Then
        strSQL = strSQL & " and tc.cod_cai='" & Left(cmbxCai.Text, 2) & "' "
    End If
    If cmbxTurno.Text <> "" Then
        strSQL = strSQL & " and tc.cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
    End If
    If cmbxOrd.ListIndex = 0 Then
        strSQL = strSQL & " order by data_cai,seq_cai"
    ElseIf cmbxOrd.ListIndex = 1 Then
        strSQL = strSQL & " order by seq_cai"
    ElseIf cmbxOrd.ListIndex = 2 Then
        strSQL = strSQL & " order by historico_cai, seq_cai"
    ElseIf cmbxOrd.ListIndex = 3 Then
        strSQL = strSQL & " order by valor_cai,seq_cai"
    End If
        
    
    
    If optnAsc.Value Then
        strSQL = strSQL & " asc "
    Else
        strSQL = strSQL & " desc "
    End If
    FechaObjeto ado
    ado.Open strSQL
    
    P_Resumo
End Sub

Sub P_Resumo()
Dim cCheDeb As Currency
    Dim cCheCred As Currency
    Dim cDinDeb As Currency
    Dim cDinCred As Currency
    
    With ado
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            
            If !esp_cai = "DIN" Then
                If !tipo_cai = "C" Then
                    cDinCred = cDinCred + !valor_Cai
                Else
                    cDinDeb = cDinDeb + !valor_Cai
                End If
            Else
                If !tipo_cai = "C" Then
                    cCheCred = cCheCred + !valor_Cai
                Else
                    cCheDeb = cCheDeb + !valor_Cai
                End If
            End If
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
    End With
    'lblCred_Che.caption = FormatVal(cCheCred)
    lblCred_din.caption = FormatVal(cDinCred)
    lblDeb_Din.caption = FormatVal(cDinDeb)
 '   lblDeb_Che.caption = FormatVal(cCheDeb)
  '  lblTotal_Che.caption = FormatVal(cCheCred - cCheDeb)
    lblTotal_Din.caption = FormatVal(cDinCred - cDinDeb)
'    lblSaldo.caption = FormatVal(CCur(lblTotal_Che.caption) + CCur(lblTotal_Din.caption))
End Sub
Private Sub cmndbtnEdit_Click()
    Dim fr As New frmLanCaixa
    With ado
        fr.tag = !seq_cai
        fr.txtbxData_Cai.Text = !data_cai
        fr.txtbxHist.Text = !historico_cai
        fr.txtbxValor_Cai.Text = FormatVal(!valor_Cai)
        If !tipo_cai = "C" Then
            fr.optnCerd.Value = True
        Else
            fr.optnDeb.Value = True
        End If
        fr.optnCerd.Enabled = False
        fr.optnDeb.Enabled = False
        fr.Show 0
    End With
End Sub

Private Sub cmndbtnExcl_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    If ado.RecordCount = 0 Then
        cmndbtnExcl.Enabled = False
        Exit Sub
    End If
    If ado!esp_cai <> "DIN" And ado!esp_cai <> "MOV" Then
        MsgBox "Este lançamento não pode ser cancelado!", vbExclamation
        Exit Sub
    End If
    
    If strExclcai_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma exclusão do Lançamento? ", vbYesNo + vbCritical) = vbYes Then
        strSQL = "Delete from tab_caixa where seq_cai ='" & ado!seq_cai & "'"
        conn.Execute strSQL
        ado.Requery
    End If
    Exit Sub
trata_Erro:
    
End Sub

Private Sub cmndbtnFechaCAi_Click()
    frmFechaCaixa.Show 1
    'frmMovDiarioCx.Show 1
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmLanCaixa
    fr.Show 0
 End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub


Private Sub cmndbtnVale_Click()
  '  On Error Resume Next
    With ado
        P_ImprimeValeCai !seq_cai
    End With
End Sub

Private Sub cmndbtnVendasNaoCaixa_Click()
'    P_VerificaVendaNaoCaixa txtbxDataIni.Text, Left(cmbxCai.Text, 2)
End Sub

Private Sub dbgrd_GotFocus()
    
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub dtServ_Reposition()
    'On Error Resume Next
    If ado.RecordCount = 0 Or dbgrd.row = -1 Then
        cmndbtnEdit.Enabled = False
        cmndbtnExcl.Enabled = False
        Exit Sub
    
    End If
    
    If ado!num_ven _
     <> "" Then
        cmndbtnEdit.Enabled = False
        cmndbtnExcl.Enabled = False
    Else
        cmndbtnEdit.Enabled = True
        cmndbtnExcl.Enabled = True
    End If
End Sub

Private Sub Form_Activate()
    'MudaAparencia Me
    If F_IsCaixa Then
        P_LocalizaOperacao strCod_Cai, cmbxCai
        P_LocalizaOperacao strCod_Tur, cmbxTurno
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF9 Then
        If pcLeft.BackColor = vbYellow Then
            pcLeft.BackColor = vbWhite
            cmndbtnBuscar_Click
        Else
            pcLeft.BackColor = vbYellow
            cmndbtnBuscar_Click
        End If
    ElseIf KeyCode = vbKeyF11 Then
        VerificaVendaDuplicada Date
    End If
    
End Sub
Sub VerificaVendaDuplicada(data As String)
    Dim strSQL As String
    strSQL = "select tv.num_ven, tc.seq_cai,tc.valor_cai  from tab_vendas as tv, tab_caixa as tc "
    strSQL = strSQL & " where tv.num_ven=tc.num_ven "
    
End Sub
Private Sub Form_Load()
   ' Me.Width = 11730
   ' Me.Height = 6705
    'Dim db As Connection
  'Set db = New Connection
'  db.CursorLocation = adUseClient
'  db.Open "PROVIDER=MSDASQL;dsn=requinte;uid=root;pwd=master;database=requinte;"
'
'  Set ado = New Recordset
'  ado.Open "select SEQ_CAI,DATA_CAI,VALOR_CAI,entrada_CAI,troco_CAI,HISTORICO_CAI,TIPO_CAI,esp_cai,NUM_VEN,cod_con,cod_cai,ip_cai,cod_usu,saldo_Cai,cod_tur,cod_pla,visivel_cai from tab_caixa Order by DATA_CAI", db, adOpenStatic, adLockOptimistic
'
'  Set grdDataGrid.DataSource = ado


  mbDataChanged = False
    
    Set ado = New ADODB.Recordset
    P_PreencheCaixas cmbxCai
    P_PreencheTurnos cmbxTurno
    cmbxBusca.ListIndex = 0
    cmbxOrd.ListIndex = 0
    cmbxTIpo.ListIndex = 0
    txtbxdataFin.Text = Date
    txtbxDataIni.Text = Date
    'ado.ConnectionString = Conn.ConnectionString
    cmndbtnBuscar_Click
    resizeForm Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    dtgrd.Height = pcBot.Top - dtgrd.Top - 100
    dtgrd.width = Me.width - dtgrd.Left - 200
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxdataFin
    If F_DataMaior(txtbxDataIni.Text, txtbxdataFin.Text) Then
        MsgBox "A data final deve ser maior que a inicial!", vbExclamation
        txtbxdataFin.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxExpr_GotFocus()
    P_SelecTexto txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    If cmbxBusca.ListIndex = 1 Then
        txtbxExpr.Text = Format(F_SoNumero(txtbxExpr.Text), "000000#")
    ElseIf cmbxBusca.ListIndex = 2 Then
        P_ValorLostFocus txtbxExpr
    Else
        P_AjustaTexto txtbxExpr
    End If
End Sub
