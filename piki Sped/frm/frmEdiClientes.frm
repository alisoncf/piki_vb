VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiClientes 
   Caption         =   "Clientes"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14430
   Icon            =   "frmEdiClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7905
   ScaleWidth      =   14430
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiClientes.frx":014A
      Height          =   5535
      Left            =   120
      TabIndex        =   24
      Top             =   840
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   9763
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
      ColumnCount     =   14
      BeginProperty Column00 
         DataField       =   "cod_cli"
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
         DataField       =   "cnpjcpf"
         Caption         =   "CNPJ/CPF"
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
         DataField       =   "ie_cli"
         Caption         =   "I.E"
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
         DataField       =   "nome_cli"
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
      BeginProperty Column04 
         DataField       =   "fantasia_cli"
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
      BeginProperty Column05 
         DataField       =   "logradouro_cli"
         Caption         =   "Logradouro"
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
         DataField       =   "complemento_cli"
         Caption         =   "Complemento"
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
         DataField       =   "numero_cli"
         Caption         =   "Número"
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
         DataField       =   "bairro_cli"
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
      BeginProperty Column09 
         DataField       =   "nome_cid"
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
      BeginProperty Column10 
         DataField       =   "uf_cid"
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
      BeginProperty Column11 
         DataField       =   "telefone_cli"
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
      BeginProperty Column12 
         DataField       =   "cel_cli"
         Caption         =   "Celular"
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
         DataField       =   "email_cli"
         Caption         =   "Email"
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
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   4410,142
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2865,26
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1665,071
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   2819,906
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   14400
      TabIndex        =   13
      Top             =   0
      Width           =   14430
      Begin VB.TextBox txtbxUf 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtbxEnd 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4800
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiClientes.frx":015C
         Left            =   120
         List            =   "frmEdiClientes.frx":016F
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnLocaliza 
         Height          =   315
         Left            =   7800
         Picture         =   "frmEdiClientes.frx":019F
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblList 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   9960
         TabIndex        =   20
         Top             =   360
         Width           =   120
      End
      Begin VB.Label lblTotal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   9960
         TabIndex        =   19
         Top             =   120
         Width           =   120
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Listados:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8880
         TabIndex        =   17
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastrados:"
         Height          =   195
         Left            =   8880
         TabIndex        =   16
         Top             =   120
         Width           =   930
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Totais:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8160
         TabIndex        =   15
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF:"
         Height          =   195
         Left            =   4200
         TabIndex        =   6
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço/Cidade:"
         Height          =   195
         Left            =   1560
         TabIndex        =   14
         Top             =   0
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome/CPF/RG/Telefone"
         Height          =   195
         Left            =   4800
         TabIndex        =   0
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   615
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
      ScaleWidth      =   14400
      TabIndex        =   8
      Top             =   6930
      Width           =   14430
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   615
         Left            =   8640
         Picture         =   "frmEdiClientes.frx":0529
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnHist 
         Caption         =   "&Histórico"
         Height          =   615
         Left            =   5760
         Picture         =   "frmEdiClientes.frx":0A1A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnCheques 
         Caption         =   "Cheques"
         Height          =   615
         Left            =   10200
         Picture         =   "frmEdiClientes.frx":0B04
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnDoc 
         Caption         =   "Documentos"
         Height          =   615
         Left            =   7200
         Picture         =   "frmEdiClientes.frx":17CE
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   1335
      End
      Begin VB.CheckBox chkNegativo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clientes Negativos"
         Height          =   555
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   11640
         Picture         =   "frmEdiClientes.frx":1D58
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   2880
         Picture         =   "frmEdiClientes.frx":1EC3
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   1440
         Picture         =   "frmEdiClientes.frx":28C5
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "&Imprimir"
         Height          =   615
         Left            =   4320
         Picture         =   "frmEdiClientes.frx":2C4F
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmEdiClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub DBGrid1_Click()

End Sub

Private Sub chkNegativo_Click()
    cmndbtnLocaliza_Click
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    
End Sub

Private Sub cmbxOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmdExportar_Click()
    ExportaGridToCSV dtgrd, ado
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmCadCli
    If ado.RecordCount <= 0 Then
        MsgBox "Não há o que ser editado!", vbCritical
        Exit Sub
    End If
    strHora = Time
    fr.tag = ado!cod_cli
    
    fr.Show 1
End Sub

Private Sub cmndbtnHist_Click()
    On Error Resume Next
    P_histCli ado!cod_cli
End Sub

Private Sub cmndbtnImprimir_Click()
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub imprime(obj As Object)
    On Error GoTo trata_Erro
    If chkNegativo.Value Then
        'P_CabRelTimbr "LISTAGEM DE CLIENTES NEGATIVADOS", obj
    Else
      '  P_CabRelTimbr "LISTAGEM DE CLIENTES", obj
    End If
    
    P_ImprX obj, "CPF/CNPJ:", 5
    P_ImprX obj, "Nome:", 45
    P_ImprX obj, "Cidade:", 155, False
    P_line1 obj
    With ado
        .MoveFirst
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                
                P_RodapeVis obj, "continua..."
                P_AddPicVis obj
                If chkNegativo.Value Then
                    P_CabRelTimbr "LISTAGEM DE CLIENTES NEGATIVADOS", obj
                Else
                    P_CabRelTimbr "LISTAGEM DE CLIENTES", obj
                End If
            End If
            P_ImprX obj, !cpf_cli, 5
            P_ImprX obj, !nome_Cli, 45
            P_ImprX obj, !nome_cid, 155, False
            
            .MoveNext
        Loop
        P_RodapeVis obj, strNome_Emp
        P_AddPicVis obj
    End With
    Exit Sub
trata_Erro:
    MostraErro ""
    
End Sub
Private Sub cmndbtnLocaliza_Click()
    Dim strSQL As String
    Dim sOrderBy As String
    On Error GoTo trata_Erro:
    strSQL = "select tc.cod_cli,if(tipo_cli='J',tc.cnpj_cli,tc.cpf_cli) as cnpjcpf,tc.cnpj_cli,tc.cpf_cli,ie_cli,"
    strSQL = strSQL & " email_cli,fantasia_cli,logradouro_cli,nome_cli,endereco_cli,complemento_cli,numero_Cli,telefone_cli,"
    strSQL = strSQL & "bairro_cli,cel_cli,fax_cli,"
    strSQL = strSQL & " ti.nome_cid,ti.uf_cid from " & TAB_CLIENTES & " as tc left join " & TAB_CIDADES & " as ti "
    strSQL = strSQL & " on tc.cod_cid=ti.cod_cid "
    
    strSQL = strSQL & " where (Nome_cli like '%" & txtbxExpr.text & "%' "

    strSQL = strSQL & "or cod_cli like '" & txtbxExpr.text & "%' "
    strSQL = strSQL & "or fantasia_cli like '" & txtbxExpr.text & "%' "
    If F_SoNumero(txtbxExpr.text) <> "" Then
        strSQL = strSQL & "or cpf_cli like '" & txtbxExpr.text & "%' "
        strSQL = strSQL & "or cnpj_cli like '" & txtbxExpr.text & "%' "
        strSQL = strSQL & "or cpf_cli like '" & F_FormataCPF(txtbxExpr.text) & "%' "
        strSQL = strSQL & "or cnpj_cli like '" & F_FormataCNPJ(txtbxExpr.text) & "%' "
    End If
    strSQL = strSQL & "or telefone_cli like '%" & txtbxExpr.text & "%' "
    strSQL = strSQL & "or cel_cli like '%" & txtbxExpr.text & "%' "
    strSQL = strSQL & "or rg_cli like '" & txtbxExpr.text & "%') "
    
    strSQL = strSQL & " and cod_cli <> 0 " 'venda a consumidor
    If txtbxEnd.text <> "" Then
        strSQL = strSQL & " and (endereco_cli like '%" & txtbxEnd.text & "%' "
        strSQL = strSQL & " or nome_cid like '%" & txtbxEnd.text & "%' "
        strSQL = strSQL & " or bairro_cli like '%" & txtbxEnd.text & "%') "
    End If
    If Trim(txtbxUf.text) <> "" Then
        strSQL = strSQL & " and uf_cid='" & txtbxUf.text & "'"
    End If
    If chkNegativo.Value = 1 Then
        strSQL = strSQL & " and status_cli=2"
    End If
    If cmbxOrd.ListIndex = 0 Then
        sOrderBy = "cod_cli"
    ElseIf cmbxOrd.ListIndex = 1 Then
        sOrderBy = "Nome_Cli,cpf_Cli"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrderBy = "CPF_CLI"
    ElseIf cmbxOrd.ListIndex = 3 Then
        sOrderBy = "nome_cid,uf_cid,NOme_Cli"
    ElseIf cmbxOrd.ListIndex = 4 Then
        sOrderBy = "uf_cid,nome_cid"
    End If
    strSQL = strSQL & " order by " & sOrderBy
    abreObjetoAdo ado, strSQL, dtgrd
    
    
    
    lblTotal.caption = F_BuscaUmCampo("count(cod_Cli)", TAB_CLIENTES, "")
    lblList.caption = ado.RecordCount
    Exit Sub
trata_Erro:
    MostraErro
    
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadCli
    fr.Show 0
'    If frmIncClientes.Visible = False Then
'        frmIncClientes.Visible = True
'    End If
End Sub

Private Sub Command4_Click()

End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    ado.Close
    Unload Me
    
End Sub

Private Sub dtgrd_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub dtgrd_GotFocus()
    dtgrd.MarqueeStyle = 3
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    ElseIf KeyCode = vbKeyC And Shift = vbCtrlMask Then
        'P_HistCheCPF  ado!cpf_cli
    End If
    
End Sub

Private Sub dtgrd_LostFocus()
    dtgrd.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    'MudaAparencia Me
    txtbxExpr.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    'Me.Width = 10530
    'Me.Height = 6810
    
    Set ado = New ADODB.Recordset
    
    cmbxOrd.ListIndex = 0
    
    'txtbxExpr.SetFocus
    resizeForm Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    dtgrd.height = pcBot.Top - dtgrd.Top - 100
    dtgrd.width = Me.width - dtgrd.Left - 200
    
    cmndbtnSair.Left = Me.width - cmndbtnSair.width - 400
    
    
End Sub

Private Sub txtbxEnd_Change()
    cmndbtnLocaliza_Click
End Sub

Private Sub txtbxEnd_GotFocus()
    ganhaFoco txtbxEnd
End Sub

Private Sub txtbxEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    
End Sub

Private Sub txtbxEnd_LostFocus()
    AjustaTexto txtbxEnd
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnLocaliza_Click
    End If
    If KeyCode = vbKeyDown Then
        dtgrd.SetFocus
    End If
    
End Sub

Private Sub txtbxExpr_LostFocus()
    AjustaTexto txtbxExpr
    
End Sub

Private Sub txtbxUf_GotFocus()
    ganhaFoco txtbxUf
End Sub

Private Sub txtbxUf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxUf_LostFocus()
    AjustaTexto txtbxUf
    cmndbtnLocaliza_Click
End Sub
