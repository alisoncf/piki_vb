VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiVales 
   Caption         =   "Vales"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13665
   Icon            =   "frmEdiVales.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   13665
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab sstb 
      Height          =   6495
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Detalhado"
      TabPicture(0)   =   "frmEdiVales.frx":27A2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dtg"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Resumido"
      TabPicture(1)   =   "frmEdiVales.frx":27BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "gridResumido"
      Tab(1).ControlCount=   1
      Begin MSDataGridLib.DataGrid dtg 
         Bindings        =   "frmEdiVales.frx":27DA
         Height          =   5895
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   10398
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "Id"
            Caption         =   "Id."
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
            DataField       =   "nome_fun"
            Caption         =   "Funcionário"
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
            DataField       =   "valor"
            Caption         =   "Valor"
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
            DataField       =   "parcela"
            Caption         =   "Parcela"
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
            DataField       =   "historico"
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
         BeginProperty Column05 
            DataField       =   "referencia"
            Caption         =   "referência"
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
            DataField       =   "data"
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
         BeginProperty Column07 
            DataField       =   "cod_usu"
            Caption         =   "Feito por"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """R$ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   ""
            Caption         =   "Reservado para Anotações"
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
               ColumnWidth     =   959,811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3374,929
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1184,882
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   3195,213
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1319,811
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1305,071
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridResumido 
         Bindings        =   "frmEdiVales.frx":27EC
         Height          =   5895
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   10398
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "cod_fun"
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
            DataField       =   "nome_fun"
            Caption         =   "Funcionário"
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
            DataField       =   "valor"
            Caption         =   "Valor"
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
            DataField       =   "anot"
            Caption         =   "anot"
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
               ColumnWidth     =   959,811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   5309,858
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1184,882
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   4860,284
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   13635
      TabIndex        =   17
      Top             =   6780
      Width           =   13665
      Begin VB.TextBox txtMes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         MaxLength       =   40
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiVales.frx":27FE
         Left            =   6720
         List            =   "frmEdiVales.frx":280B
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtFun 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         MaxLength       =   40
         TabIndex        =   7
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton cmdBuscarFun 
         Height          =   285
         Left            =   6240
         Picture         =   "frmEdiVales.frx":282F
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mês:"
         Height          =   195
         Left            =   2280
         TabIndex        =   4
         Top             =   0
         Width           =   345
      End
      Begin VB.Label label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   8520
         TabIndex        =   19
         Top             =   120
         Width           =   360
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   9120
         TabIndex        =   18
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   2
         Top             =   0
         Width           =   285
      End
      Begin VB.Label lblDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   6720
         TabIndex        =   9
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionário"
         Height          =   195
         Left            =   3360
         TabIndex        =   6
         Top             =   0
         Width           =   825
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   13635
      TabIndex        =   15
      Top             =   7395
      Width           =   13665
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   705
         Left            =   6360
         Picture         =   "frmEdiVales.frx":2928
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "Pesquisar"
         Height          =   705
         Left            =   120
         Picture         =   "frmEdiVales.frx":2D57
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   705
         Left            =   4800
         Picture         =   "frmEdiVales.frx":2E50
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   705
         Left            =   1680
         Picture         =   "frmEdiVales.frx":2F9A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   705
         Left            =   12120
         Picture         =   "frmEdiVales.frx":3324
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdLista 
         Caption         =   "Lista"
         Height          =   705
         Left            =   3240
         Picture         =   "frmEdiVales.frx":3872
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmEdiVales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim WithEvents adoResumido As ADODB.Recordset
Attribute adoResumido.VB_VarHelpID = -1

Private Sub cmbxOrd_Click()
    busca
End Sub

Private Sub cmdBuscarFun_Click()
    Dim f As funcionario.funcionario
    f = funcionario.F_SelectFuncionario(txtFun.text)
    If f.cod_fun <> "" Then
        txtFun.text = f.nome_fun
        txtFun.tag = f.cod_fun
        cmdPesquisar_Click
    End If

End Sub

Private Sub cmdExportar_Click()
    If sstb.tag = 0 Then
        ExportaGridToCSV dtg, ado
    Else
        ExportaGridToCSV gridResumido, adoResumido
    End If
End Sub

Private Sub cmdLista_Click()
    If sstb.Tab = 0 Then
    
        P_DtgrdToHtml "Vales (" & txtbxDe.text & "-" & txtbxAte.text & ")", dtg, ado, "Funcionário: " & txtFun.text, "2"
    Else
        P_DtgrdToHtml "Vales (" & txtbxDe.text & "-" & txtbxAte.text & ")", gridResumido, adoResumido, "", "2"
    End If
End Sub

Private Sub cmdPesquisar_Click()
    busca
    buscaResumido
End Sub

Private Sub cmndbtnNew_Click()
    If strLancaixa_usu = "S" Then
        frmLancaVale.Show 1
    Else
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    Set adoResumido = New ADODB.Recordset
    
    txtbxDe.text = F_PrimDiaMes(Date)
    txtbxAte.text = F_UltDiaMes(Date)
    sstb.Tab = 0
    
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select v.*,nivel_usu,tf.nome_fun  from vales as v left join tab_funcionarios as tf "
    strSQL = strSQL & " on tf.cod_fun=v.cod_fun "
    strSQL = strSQL & " left join tab_usuarios as tu on tu.cod_fun=tf.cod_fun "
    strSQL = strSQL & " where 1>0 "
    
    If txtFun.tag <> "" Then
        strSQL = strSQL & " and tf.cod_fun='" & txtFun.tag & "' "
    End If
    
    If txtbxDe.text <> "" Then
        strSQL = strSQL & " and referencia>=" & F_GravaData(txtbxDe.text)
    End If
    If txtbxAte.text <> "" Then
        strSQL = strSQL & " and referencia<=" & F_GravaData(txtbxAte.text)
    End If
    strSQL = strSQL & " and (nivel_usu<=" & strNivel_Usu & " or nivel_usu is null)"
    If cmbxOrd.ListIndex = 0 Then
        strSQL = strSQL & " order by nome_fun"
    ElseIf cmbxOrd.ListIndex = 1 Then
        strSQL = strSQL & " order by nome_fun,referencia,valor desc"
    ElseIf cmbxOrd.ListIndex = 2 Then
        strSQL = strSQL & " order by nome_fun,valor desc, referencia"
    Else
    
    End If
    
    
    
    abreObjetoAdo ado, strSQL, dtg
    
    soma
    
End Sub
Sub buscaResumido()
    Dim strSQL As String
    strSQL = "select v.cod_fun,sum(v.valor) as valor,tf.nome_fun,"
    strSQL = strSQL & "'" & String(40, "_") & CrLf & CrLf & CrLf & "' as anot   from vales as v left join tab_funcionarios as tf "
    strSQL = strSQL & " on tf.cod_fun=v.cod_fun "
    strSQL = strSQL & " left join tab_usuarios as tu on tu.cod_fun=tf.cod_fun "
    strSQL = strSQL & " where 1>0 "
    
    If txtFun.tag <> "" Then
        strSQL = strSQL & " and tf.cod_fun='" & txtFun.tag & "' "
    End If
    
    If txtbxDe.text <> "" Then
        strSQL = strSQL & " and referencia>=" & F_GravaData(txtbxDe.text)
    End If
    If txtbxAte.text <> "" Then
        strSQL = strSQL & " and referencia<=" & F_GravaData(txtbxAte.text)
    End If
     strSQL = strSQL & " and (nivel_usu<=" & strNivel_Usu & " or nivel_usu is null)"
    strSQL = strSQL & " group by v.cod_fun "
    If cmbxOrd.ListIndex = 0 Then
        strSQL = strSQL & " order by nome_fun"
    ElseIf cmbxOrd.ListIndex = 1 Then
        strSQL = strSQL & " order by nome_fun,referencia,valor desc"
    ElseIf cmbxOrd.ListIndex = 2 Then
        strSQL = strSQL & " order by nome_fun,valor desc, referencia"
    Else
    
    End If
    
    
    
    abreObjetoAdo adoResumido, strSQL, gridResumido
    
    soma
    
End Sub

Sub soma()
    On Error GoTo trata_Erro
    Dim total As Currency
    dtg.Visible = False
    
    With ado
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            total = total + !valor
            
            
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
    End With


    lblTotal.caption = FormatVal(total)
dtg.Visible = True
trata_Erro:
dtg.Visible = True
    Exit Sub
    lblTotal = "erro"
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    'PosicionaGrade dtg, Me, pcBot
   ' posicionaADireita cmndbtnSair, pcBot
    
End Sub

Private Sub txtbxAte_GotFocus()
    P_SelecTexto txtbxAte
End Sub

Private Sub txtbxAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxAte_LostFocus()
    P_AjustaData txtbxAte
    'txtbxAte.Text = F_UltDiaMes(txtbxAte.Text)
End Sub
Private Sub txtbxDe_GotFocus()
    P_SelecTexto txtbxDe
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDe_LostFocus()
    P_AjustaData txtbxDe
'    txtbxDe.Text = F_PrimDiaMes(txtbxDe.Text)
End Sub

Private Sub txtFun_GotFocus()
    P_SelecTexto txtFun
End Sub

Private Sub txtFun_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdBuscarFun_Click
    End If
    
End Sub

Private Sub txtMes_GotFocus()
    P_SelecTexto txtMes
End Sub

Private Sub txtMes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtMes_LostFocus()
    P_AjustaData txtMes, "mm/yyyy"
    If txtMes.text <> "" Then
        txtbxDe.text = "01/" & txtMes.text
        txtbxAte.text = F_UltDiaMes(txtbxDe.text)
    End If
    
End Sub
