VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiNFSaida 
   Caption         =   "Edição de Notas Fiscais de Saída"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   ScaleHeight     =   9810
   ScaleWidth      =   14085
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9810
      Left            =   0
      ScaleHeight     =   9780
      ScaleWidth      =   2265
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   15
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   14
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   13
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmEdiNFSaida.frx":0000
         Left            =   120
         List            =   "frmEdiNFSaida.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiNFSaida.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   8160
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "&Imprimir"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiNFSaida.frx":0182
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4560
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "Nova Nota Fiscal"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiNFSaida.frx":058F
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnCancelar 
         Caption         =   "&Cancelar"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiNFSaida.frx":0919
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6000
         Width           =   2055
      End
      Begin VB.TextBox txtbxCliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "Buscar"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiNFSaida.frx":0A63
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnImprGrade 
         Caption         =   "&Imprimir Grade"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiNFSaida.frx":0DED
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5280
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Excluir"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiNFSaida.frx":11FA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   6720
         Width           =   2055
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "Exportar NFe"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiNFSaida.frx":1344
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   7440
         Width           =   2055
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Editar"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiNFSaida.frx":16CE
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3840
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnTerc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         Picture         =   "frmEdiNFSaida.frx":20D0
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   18
         Top             =   1200
         Width           =   285
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   510
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiNFSaida.frx":245A
      Height          =   9015
      Left            =   2400
      TabIndex        =   21
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   15901
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
      Caption         =   "Clientes"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "nota_sai"
         Caption         =   "Numero"
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
         DataField       =   "nome_cli"
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
         DataField       =   "cpf_cli"
         Caption         =   "CPF/CNPJ"
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
         DataField       =   "valor_sai"
         Caption         =   "Valor"
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
         DataField       =   "emissao_sai"
         Caption         =   "Emissão"
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
         DataField       =   "nfe_sai"
         Caption         =   "NFE"
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
         DataField       =   "situacao_sai"
         Caption         =   "Situação"
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
            ColumnWidth     =   3075,024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1874,835
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1665,071
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiNFSaida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmbxBuscar_Click()
    
    End Sub

Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxTipo_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub


Private Sub cmdExport_Click()
    On Error GoTo trata_erro
    If MsgBox("Gerar arquivo da NF " & adc.Recordset!seq_sai & "?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    'ExportaNfe adc.Recordset!seq_sai
    Exit Sub
trata_erro:
    'MostraErro
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim sOrd As String
    If cmbxOrder.ListIndex = 0 Then
        sOrd = "nota_sai desc "
    ElseIf cmbxOrder.ListIndex = 1 Then
        sOrd = "Nome_cli, nota_sai desc"
    ElseIf cmbxOrder.ListIndex = 2 Then
        sOrd = "Emissao_sai desc, nota_sai desc "
    Else
        sOrd = "emissao_sai desc, nota_sai desc "
    End If
    strSQL = "select ts.*,tc.nome_cli,tc.uf_cli,tc.cpf_cli from tab_clientes as tc, tab_saidasnf as ts "
    strSQL = strSQL & " where ts.cod_cli=tc.cod_cli "
    If cmndbtnTerc.Visible = True Then
        strSQL = strSQL & " and tc.nome_cli like '" & txtbxExpr.Text & "%' "
    Else
        strSQL = strSQL & " and ts.nota_sai like '%" & txtbxExpr.Text & "%' "
    End If
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and ts.emissao_sai < " & F_PesqData(txtbxdataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and ts.emissao_sai > " & F_PesqData(txtbxDataIni.Text)
    End If
    strSQL = strSQL & " and ts.serie_sai not in('CUPM','CANC')"
    strSQL = strSQL & " order by " & sOrd
    abreObjetoAdo ado, strSQL, dtgrd
End Sub

Private Sub cmndbtnCancelar_Click()
    Dim strSQL As String
    With adc.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If !situacao_sai = "3" Then
            MsgBox "A nota fiscal já foi cancelada!", vbExclamation
            Exit Sub
        End If
       
        If MsgBox("Deseja Realmente Excluir a Nota Fiscal " & !nota_sai & _
            " ?", vbYesNo + vbCritical) = vbYes Then
            
            
            P_cancelaNotaFiscalSaida !seq_sai, ""
            
            
            adc.Refresh
        End If
    End With
End Sub

Private Sub cmdEdit_Click()
    
    
    If adc.Recordset.RecordCount <= 0 Or dtgrd.row = -1 Then
        MsgBox "Não há notas fiscais nesta pesquisa!", vbExclamation
        Exit Sub
    End If
    
    With adc.Recordset
        
        
        frmGeraNota.tag = !seq_sai
        frmGeraNota.setSeq_Sai !seq_sai
        frmGeraNota.Show 1
        
    End With
End Sub


Private Sub cmndbtnImprGrade_Click()
    'P_DbgrdToHtml dtgrd.caption & "De: " & txtbxDataIni.Text & " até " & txtbxdataFin.Text, _
        dtgrd, adc
End Sub

Private Sub cmndbtnImprimir_Click()
    On Error GoTo trata_erro
    With adc.Recordset
        'imprimenota !seq_sai
    
    End With
    Exit Sub
trata_erro:
    'MostraErro
End Sub

Private Sub cmndbtnNew_Click()
    frmGeraNota.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    adc.Recordset.Close
    Unload Me
End Sub

Private Sub cmndbtnTerc_Click()
    'Dim c As tCliente
    'c = Clientes.F_BuscaCliente("C", txtbxCliente.Text)
    If c.cod_cli <> "" Then
        txtbxCliente.Text = c.nome_Cli
        txtbxCliente.tag = c.nome_Cli
        cmndbtnBuscar_Click
        txtbxExpr.SetFocus
    End If
    
End Sub

Private Sub dtgrd_DblClick()
    cmdEdit_Click
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

Private Sub adc_Reposition()
    With adc.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        
        
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        
        
        
    End If
End Sub

Private Sub Form_Load()
   ' Me.Width = 10560
   ' Me.Height = 6585
    resizeForm Me
    Set ado = New ADODB.Recordset
    
    cmbxOrder.ListIndex = 0
    
    cmbxBuscar_Click
    cmndbtnBuscar_Click
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    dtgrd.Width = Me.Width - dtgrd.Left - 200
    dtgrd.height = Me.height - dtgrd.Top - 400
    
End Sub

Private Sub txtbxCliente_Change()
    txtbxCliente.tag = ""
End Sub

Private Sub txtbxdataFin_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxdataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxdataFin_LostFocus()
    P_AjustaData txtbxdataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        SendKeys Chr(9)
        cmndbtnBuscar_Click
    End If
    
End Sub




