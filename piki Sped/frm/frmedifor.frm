VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiFor 
   Caption         =   "Edição de Fornecedores"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10890
   Icon            =   "frmedifor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   10890
   Begin VB.PictureBox pcBusca 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   10890
      TabIndex        =   0
      Top             =   5910
      Width           =   10890
      Begin VB.CheckBox chkInv 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Listar Invisíveis"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8400
         TabIndex        =   17
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmedifor.frx":014A
         Left            =   120
         List            =   "frmedifor.frx":015D
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   40
         TabIndex        =   4
         Top             =   240
         Width           =   4095
      End
      Begin VB.CommandButton cmndbtnLocaliza 
         Caption         =   "Pesquisar"
         Height          =   525
         Left            =   5760
         Picture         =   "frmedifor.frx":0189
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&xpressão:"
         Height          =   195
         Left            =   1560
         TabIndex        =   3
         Top             =   0
         Width           =   780
      End
      Begin VB.Label lblList 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listados: 0"
         Height          =   195
         Left            =   8400
         TabIndex        =   13
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total de Fornecedores: 0"
         Height          =   195
         Left            =   8400
         TabIndex        =   12
         Top             =   0
         Width           =   1785
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   10860
      TabIndex        =   6
      Top             =   6645
      Width           =   10890
      Begin VB.CommandButton cmdInvisivel 
         Caption         =   "Invisível"
         Height          =   615
         Left            =   7200
         Picture         =   "frmedifor.frx":0513
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSelect 
         Caption         =   "&Selecionar"
         Height          =   615
         Left            =   2880
         Picture         =   "frmedifor.frx":089D
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cheques"
         Height          =   615
         Left            =   5760
         Picture         =   "frmedifor.frx":0E27
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnDocs 
         Caption         =   "&Documentos"
         Height          =   615
         Left            =   4320
         Picture         =   "frmedifor.frx":13B1
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   8880
         Picture         =   "frmedifor.frx":15C7
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   1440
         Picture         =   "frmedifor.frx":1732
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   0
         Picture         =   "frmedifor.frx":187C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmedifor.frx":1C06
      Height          =   4695
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   8281
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
      Caption         =   "Clientes"
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "cod_for"
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
         DataField       =   "cnpj_for"
         Caption         =   "CNPJ"
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
         DataField       =   "nome_for"
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
         DataField       =   "fantasia_for"
         Caption         =   "Fantasia"
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
         DataField       =   "endereco_for"
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
      BeginProperty Column05 
         DataField       =   "bairro_for"
         Caption         =   "Bairro"
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
         DataField       =   "cidade_for"
         Caption         =   "Cidade"
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
         DataField       =   "uf_for"
         Caption         =   "uf"
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
         DataField       =   "telefone_for"
         Caption         =   "Telefone"
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
         DataField       =   "fax_for"
         Caption         =   "Fax"
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
         DataField       =   "email_for"
         Caption         =   "email"
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
            ColumnWidth     =   929,764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1470,047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3000,189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1934,929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2789,858
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2039,811
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   480,189
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   330
      Left            =   120
      Top             =   5760
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmEdiFor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkInv_Click()
    cmndbtnLocaliza_Click
End Sub

Private Sub cmdInvisivel_Click()
    Dim strSQL As String
On Error GoTo trata_Erro
    With adc.Recordset
        If adc.Recordset.RecordCount <= 0 Then
            MsgBox "Não há o que ser modificado!", vbCritical
            Exit Sub
        End If
        If !visivel_for = "S" Then
            If MsgBox("Deseja tornar o fornecedor selecionado invisivel?", vbYesNo + vbQuestion) = vbYes Then
                strSQL = "update " & TAB_FORNECEDORES & " set visivel_for='N' where cod_for='" & !cod_for & "'"
                conn.Execute strSQL
                cmndbtnLocaliza_Click
            End If
        Else
            If MsgBox("Deseja tornar o fornecedor selecionado visível?", vbYesNo + vbQuestion) = vbYes Then
                strSQL = "update " & TAB_FORNECEDORES & " set visivel_for='S' where cod_for='" & !cod_for & "'"
                conn.Execute strSQL
                cmndbtnLocaliza_Click
            End If
        End If
    End With
    Exit Sub
trata_Erro:
    MostraErro
    
End Sub

Private Sub cmndbtnEdit_Click()
    If adc.Recordset.RecordCount <= 0 Then
        MsgBox "Não há o que ser editado!", vbCritical
        Exit Sub
    End If
    frmIncFor.tag = adc.Recordset!cod_for
    frmIncFor.Show 1
End Sub

Private Sub cmndbtnLocaliza_Click()
    Dim strSQL As String
    Dim sOrderBy As String
    
    strSQL = "select * from " & TAB_FORNECEDORES & " where "
    strSQL = strSQL & "(Nome_for like '%" & txtbxExpr.Text & "%' "
    strSQL = strSQL & "or cod_for like '" & txtbxExpr.Text & "%' "
    strSQL = strSQL & "or cnpj_for like '" & txtbxExpr.Text & "%' "
    strSQL = strSQL & "or ie_for like '" & txtbxExpr.Text & "%') "
        
    If chkInv.Value = 0 Then
        strSQL = strSQL & " and visivel_for<>'N'"
    Else
        strSQL = strSQL & " and visivel_for='N'"
    End If
    If cmbxOrd.ListIndex = 0 Then
        sOrderBy = "cod_for"
    ElseIf cmbxOrd.ListIndex = 1 Then
        sOrderBy = "Nome_for,cnpj_for"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrderBy = "CNPJ_for"
    ElseIf cmbxOrd.ListIndex = 3 Then
        sOrderBy = "Cidade_for,uf_for,NOme_for"
    ElseIf cmbxOrd.ListIndex = 4 Then
        sOrderBy = "UF_for,Cidade_for"
    End If
    strSQL = strSQL & " order by " & sOrderBy
    adc.RecordSource = strSQL
    adc.Refresh
    
    lblTotal.caption = "Fornecedores Cadastrados...: " & F_BuscaUmCampo("count(cod_for)", TAB_FORNECEDORES, "")
    lblList.caption = "Listados...: " & adc.Recordset.RecordCount
End Sub

Private Sub cmndbtnNew_Click()
    
    frmIncFor.Show 1
    'If frmIncFor.Visible = False Then
        'frmIncFor.Visible = True
    'End If
End Sub

Private Sub Command4_Click()

End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
    
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub cmndbtnSelect_Click()
    Me.tag = "x"
    Me.Hide
End Sub

Private Sub dtgrd_DblClick()
    If Me.tag = "" Then
        cmndbtnEdit_Click
    Else
        cmndbtnSelect_Click
    End If
End Sub

Private Sub dtgrd_GotFocus()
    dtgrd.MarqueeStyle = 3
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    End If
End Sub

Private Sub dtgrd_LostFocus()
    dtgrd.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    
    
    cmbxOrd.ListIndex = 0
    cmndbtnLocaliza_Click
    
    If Me.tag = "" Then
        cmndbtnSelect.Enabled = False
    Else
        On Error Resume Next
        cmndbtnSelect.Enabled = True
        dtgrd.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    resizeForm Me
    
    adc.ConnectionString = strConnect
    
    
       
    
End Sub

Private Sub Form_Resize()
    dtgrd.width = Me.width - 300
    dtgrd.Height = pcBusca.Top - 300
    cmndbtnSair.Left = Me.width - cmndbtnSair.width - 200
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnLocaliza_Click
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    AjustaTexto txtbxExpr
    'If cmbxBusca.Text = "CNPJ" Then
       
    'End If
End Sub

