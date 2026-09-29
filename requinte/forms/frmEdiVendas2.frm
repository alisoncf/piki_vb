VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiVendas2 
   Caption         =   "Edição de Vendas"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14460
   Icon            =   "frmEdiVendas2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   14460
   Begin MSAdodcLib.Adodc adc 
      Height          =   330
      Left            =   6360
      Top             =   5400
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.PictureBox pcVenda 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   14430
      TabIndex        =   25
      Top             =   6120
      Width           =   14460
      Begin VB.Label lblCrediario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venda Disponível"
         Height          =   195
         Left            =   5400
         TabIndex        =   27
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label lblVendaAberta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venda Disponível"
         Height          =   195
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   1275
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   14430
      TabIndex        =   14
      Top             =   7110
      Width           =   14460
      Begin VB.CommandButton cmndbtnDestrava 
         Caption         =   "&Destravar Venda"
         Height          =   585
         Left            =   11640
         Picture         =   "frmEdiVendas2.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnReabre 
         Caption         =   "&Reabrir Venda"
         Height          =   585
         Left            =   10200
         Picture         =   "frmEdiVendas2.frx":1054
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnReem 
         Caption         =   "&Reemitir CUPM"
         Height          =   585
         Left            =   7320
         Picture         =   "frmEdiVendas2.frx":119E
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnCancel 
         Caption         =   "&Cancelar"
         Height          =   585
         Left            =   8760
         Picture         =   "frmEdiVendas2.frx":12E8
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnFinalizarVenda 
         Caption         =   "&Efetivar Venda"
         Height          =   585
         Left            =   4440
         Picture         =   "frmEdiVendas2.frx":1432
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Nova"
         Height          =   585
         Left            =   120
         Picture         =   "frmEdiVendas2.frx":17BC
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   585
         Left            =   3000
         Picture         =   "frmEdiVendas2.frx":1B46
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   585
         Left            =   13080
         Picture         =   "frmEdiVendas2.frx":1C90
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   585
         Left            =   1560
         Picture         =   "frmEdiVendas2.frx":1DDA
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnFinalizar 
         Caption         =   "&Finalizar"
         Height          =   585
         Left            =   5880
         Picture         =   "frmEdiVendas2.frx":2164
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Timer tmr 
      Interval        =   1000
      Left            =   2640
      Top             =   4200
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   14430
      TabIndex        =   13
      Top             =   6495
      Width           =   14460
      Begin VB.TextBox txtProd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7800
         MaxLength       =   40
         TabIndex        =   31
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtbxVendedor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         MaxLength       =   40
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbxSituacao_ven 
         Height          =   315
         ItemData        =   "frmEdiVendas2.frx":24EE
         Left            =   120
         List            =   "frmEdiVendas2.frx":2501
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiVendas2.frx":254E
         Left            =   10080
         List            =   "frmEdiVendas2.frx":255E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Height          =   285
         Left            =   7440
         Picture         =   "frmEdiVendas2.frx":2589
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtbxDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3960
         MaxLength       =   40
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         MaxLength       =   40
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdProd 
         Height          =   285
         Left            =   9600
         Picture         =   "frmEdiVendas2.frx":2682
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   7800
         TabIndex        =   33
         Top             =   0
         Width           =   600
      End
      Begin VB.Label lblVenFec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Formas Pag:"
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
         Left            =   12600
         TabIndex        =   29
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pagamento:"
         Height          =   195
         Left            =   12600
         TabIndex        =   28
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   6120
         TabIndex        =   10
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Situação da Venda:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   10080
         TabIndex        =   0
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número/Cliente"
         Height          =   195
         Left            =   1800
         TabIndex        =   4
         Top             =   0
         Width           =   1110
      End
      Begin VB.Label lblDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   3960
         TabIndex        =   6
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   5040
         TabIndex        =   8
         Top             =   0
         Width           =   285
      End
   End
   Begin MSDataGridLib.DataGrid dtg 
      Bindings        =   "frmEdiVendas2.frx":277B
      Height          =   5895
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   14175
      _ExtentX        =   25003
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
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "num_ven"
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
         DataField       =   "data_ven"
         Caption         =   "data"
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
         DataField       =   "dataefet_ven"
         Caption         =   "Dt. Efet."
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
         DataField       =   "valor_ven"
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
      BeginProperty Column05 
         DataField       =   "desconto_ven"
         Caption         =   "Desconto"
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
         DataField       =   "total"
         Caption         =   "Vlr Total"
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
         DataField       =   "situacao_ven"
         Caption         =   "S"
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
         DataField       =   "guerra_fun"
         Caption         =   "Vendedor"
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
         DataField       =   "nome"
         Caption         =   "Fechado por:"
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
         DataField       =   "nome_for"
         Caption         =   "Forma Pag."
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
         DataField       =   "autorizadeconto_ven"
         Caption         =   "Autoriza Desconto."
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
         DataField       =   "cancelavenda_usu"
         Caption         =   "Cancelamento"
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
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3374,929
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
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
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiVendas2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Dim sTab As String
Dim sCod As String
Dim sNome As String
Dim WithEvents AdoVendas As ADODB.Recordset
Attribute AdoVendas.VB_VarHelpID = -1


'Private Sub adc_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    On Error Resume Next
'
'    Dim strSQL As String
'    If dtVendas.Recordset.RecordCount = 0 Then
'        cmndbtnEdit.Enabled = False
'        cmndbtnImpr.Enabled = False
'        Exit Sub
'    Else
'        cmndbtnImpr.Enabled = True
'        cmndbtnEdit.Enabled = True
'    End If
'    If dtVendas.Recordset!situacao_ven > 1 Then
'        cmndbtnEdit.Enabled = False
'
'    Else
'        cmndbtnEdit.Enabled = True
'
'    End If
'
'    strSQL = F_BuscaUmCampo("ip_ven", "tab_vendaaberta", "num_ven='" & dtVendas.Recordset!num_ven & "'")
'    If strSQL = "" Then
'        lblVendaAberta.caption = "..."
'    Else
'        lblVendaAberta.caption = "Esta venda está sendo utilizada pelo Terminal " & strSQL
'    End If
'    buscaCrediario
'    buscaVenFec
'End Sub

Private Sub cmbxBusca_Click()
    P_Busca
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxOrd_Click()
    P_Busca
End Sub

Private Sub cmbxOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxSituacao_ven_Click()
   P_Busca
End Sub

Private Sub cmdProd_Click()
    buscaProd
End Sub
Sub buscaProd()
    Dim p As Produto
    
    p = Produtos.F_BuscaProdSaida(txtProd.text)
    If p.cod_pro <> "" Then
        txtProd.text = p.nome_pro
        txtProd.tag = p.cod_pro
        P_Busca
    End If
    
End Sub
Private Sub cmndbtnBuscar_Click()
    P_Status "Aguarde..."
    P_Busca
    P_Status ""
End Sub

Private Sub cmndbtnCancel_Click()
    If strExclVenda_usu = "S" Then
        If MsgBox("Deseja realmente cancelar todos os lancamentos dessa venda?", vbYesNo + vbQuestion) = vbYes Then
            If dtVendas.Recordset!situacao_ven < 3 Then
                MsgBox "A venda não pode ser cancelada por aqui!", vbCritical
                Exit Sub
            End If
            If Not F_CancelaVendaAposCupom(dtVendas.Recordset!num_ven) Then
                dtVendas.Refresh
            End If
        End If
    Else
        MsgBox "Você não tem permissão para isso!", vbCritical
    End If
End Sub

Private Sub cmndbtnDestrava_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    If strExclVenda_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
        Exit Sub
    End If
    
    If Not F_senha(True) Then
        Exit Sub
    End If
    strSQL = F_BuscaUmCampo("ip_ven", "tab_vendaaberta", "num_ven='" & dtVendas.Recordset!num_ven & "'")
    If strSQL = "" Then
        MsgBox "A venda não está marcada como sendo utilizada em nenhum computador!", vbExclamation
        Unload Me
    End If
    strSQL = "A Venda " & dtVendas.Recordset!num_ven & " está marcada como sendo ulitazada pelo computador " & strSQL & Chr(13)
    strSQL = strSQL & "Tenha certeza que há um problema de visualização da venda, "
    strSQL = strSQL & "e que esta não se encontra aberta em nenhum computador."
    strSQL = strSQL & " Caso contrário, poderá haver problemas de valores e anomalias nesta Venda. " & Chr(13)
    strSQL = strSQL & "Tem certeza que deseja forçar a abertura desta venda?"
    If MsgBox(strSQL, vbCritical + vbYesNo) = vbYes Then
        P_DeletaVendaAberta dtVendas.Recordset!num_ven
    End If
    
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnEdit_Click()
    If Me.tag = "U" Then
        editaVendaFunc
    Else
        editaVenda
    End If
End Sub
Sub editaVenda()
    Dim fr As New frmVenda
    P_Status "Montando venda..."
    
    
    With adc.Recordset
        If !situacao_ven <> 1 Then
            MsgBox "Esta venda não pode ser editada. Ela não está na situação 'em aberto'", vbExclamation
            Exit Sub
        End If
        fr.tag = !num_ven
        'fr.lblAbert.caption = !abertura_ven
        'frmVenda.lblNome_Cli.caption = !nome_cli
        'frmVenda.lblNome_Cli.Tag = !cod_cli
        'frmVenda.lblHora.caption = !data_ven
        'frmVenda.lblCpf.caption = F_BuscaUmCampo("CPF_Cli", "tab_clientes", "cod_cli='" & !cod_cli & "'")
        'frmVenda.lblCPFCap.caption = "CPF/CNPJ:"
        'frmVenda.lblNome_Ven.caption = F_BuscaUmCampo("Guerra_Fun", "tab_Funcionarios", _
            "cod_fun='" & !cod_usu & "'")
        'frmVenda.txtbxEndereco_ven.Text = Format(!endereco_ven)
        fr.Atacado = True
        fr.lblValor.caption = FormatVal(!valor_ven)
        fr.caption = "Venda " & !num_ven & "-" & !nome_Cli
        
        'frmVenda.txtbxObs.Text = Format(!obs_ven)
        'fr.Frame1.caption = Format(!nomecli_ven) & " Atendente[" & strNome_Usu & "]"
        fr.Show 0, mdiFrm
        'fr.lblAtacado.tag = "X" 'indica que a venda foi editada
    End With
    P_Status ""
    dtVendas.Refresh
End Sub
Sub editaVendaFunc()
    Dim fr As New frmVendaFunc
    P_Status "Montando venda..."
    With dtVendas.Recordset
        fr.tag = !num_ven
        'fr.lblAbert.caption = !abertura_ven
        'frmVenda.lblNome_Cli.caption = !nome_cli
        'frmVenda.lblNome_Cli.Tag = !cod_cli
        'frmVenda.lblHora.caption = !data_ven
        'frmVenda.lblCpf.caption = F_BuscaUmCampo("CPF_Cli", "tab_clientes", "cod_cli='" & !cod_cli & "'")
        'frmVenda.lblCPFCap.caption = "CPF/CNPJ:"
        'frmVenda.lblNome_Ven.caption = F_BuscaUmCampo("Guerra_Fun", "tab_Funcionarios", _
            "cod_fun='" & !cod_usu & "'")
        'frmVenda.txtbxEndereco_ven.Text = Format(!endereco_ven)
        fr.Atacado = True
        fr.lblValor.caption = FormatVal(!valor_ven)
        fr.caption = "Venda " & !num_ven & "-" & !nome_Cli
        
        'frmVenda.txtbxObs.Text = Format(!obs_ven)
        'fr.Frame1.caption = Format(!nomecli_ven) & " Atendente[" & strNome_Usu & "]"
        fr.Show 0, mdiFrm
        'fr.lblAtacado.tag = "X" 'indica que a venda foi editada
    End With
    P_Status ""
    dtVendas.Refresh
End Sub
Private Sub cmndbtnFinalizar_Click()
    Dim s As String
    Dim sStatus As String
    On Error GoTo trata_Erro
    With dtVendas.Recordset
        If !situacao_ven <> "3" Then
            MsgBox "A venda não está apta a ser finalizada!", vbExclamation
            Exit Sub
        End If
        sStatus = F_BuscaUmCampo("status_Cai", "tab_caixas", "cod_cai='" & strCod_Cai & "'")
        If sStatus = "F" Then
            MsgBox "Não é possível efetuar o recebimento pois este caixa está fechado!", vbExclamation
            Exit Sub
        End If
        frmLanParcelasVen.tag = !num_ven
        frmLanParcelasVen.Show 1
        
        
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnFinalizarVenda_Click()
    Dim s As String
    Dim sStatus As String
    
    On Error GoTo trata_Erro
    With adc.Recordset
        
        If strEfetivarVenda_Usu = "N" Then
            MsgBox "Você não tem permissão para efetuar venda " & strNome_Usu, vbExclamation
            Exit Sub
        End If
        
        s = F_BuscaUmCampo("status_cre", "tab_crediarioven", "num_ven='" & !num_ven & "'")
        If s = "1" Then
            MsgBox "Esta venda está sendo apreciada pelo crediário!", vbExclamation
            Exit Sub
        End If
        If s = "3" Then
            MsgBox "Esta venda foi  reprovada pelo crediário!", vbExclamation
            Exit Sub
        End If
        If strEfetivarVenda_Usu <> "S" Then
            MsgBox "Este usuário não tem permissão para efetivar venda!", vbExclamation
            Exit Sub
        End If
        
        If !situacao_ven <> "2" Then
            MsgBox "A venda não está apta a ser finalizada!", vbExclamation
            Exit Sub
        End If
            
        If Not F_IsCaixa Then
            MsgBox "Este computador não é um caixa, portanto,não pode efetivar venda!", vbExclamation
            Exit Sub
        End If
            
        sStatus = F_BuscaUmCampo("status_Cai", "tab_caixas", "cod_cai='" & strCod_Cai & "'")
        If sStatus = "F" Then
            MsgBox "Não é possível efetuar o recebimento pois este caixa está fechado!", vbExclamation
            Exit Sub
        End If
        s = F_Troco(!num_ven)
        
        
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnImpr_Click()
    ImprVenda dtVendas.Recordset!num_ven
End Sub

Private Sub cmndbtnNew_Click()
    On Error Resume Next
   Dim strSQL As String
    
Dim s As String
    If Me.tag = "U" Then
        s = F_GeraNovaVendaFuncionario()
    
    
    
    
        If s <> "" Then
            frmVendaFunc.tag = s
            frmVendaFunc.Show 0
        End If
    
    Else
        If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    strSQL = F_GeraNovaVenda("00000")
    
    
    
    frmVenda.caption = "Venda a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
    frmVenda.tag = strSQL
    On Error Resume Next
    frmVenda.Show 0, mdiFrm
    
    
    End If
    
End Sub



Private Sub cmndbtnReabre_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    If dtVendas.Recordset!situacao_ven <> "2" Then
        MsgBox "A Venda não pode ser reaberta!", vbExclamation
        Exit Sub
    End If
    If strExclVenda_usu <> "S" Then
        MsgBox "Você não tem permissão para Reabrir Vendas/Excluir Vendas!", vbExclamation
        Exit Sub
    
    
    End If
    
    If MsgBox("Confirma Reabertura da Venda?", vbYesNo + vbQuestion) = vbYes Then
        If Vendas.F_AlteraStatusVenda(dtVendas.Recordset!num_ven, "1") Then
            MsgBox "Venda reaberta com sucesso!", vbInformation
            dtVendas.Refresh
        End If
        
    End If
    Exit Sub
    
trata_Erro:
P_MostraErro
End Sub

Private Sub cmndbtnReem_Click()
    frmReemissao.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    dtVendas.Recordset.Close
    Unload Me
End Sub


Private Sub cmndbtnVendasEfet_Click()
    frmEfetivaVen.Show 0
End Sub


Private Sub DBGrid1_DblClick()
    On Error GoTo trata_Erro
    If cmndbtnEdit.Enabled = True Then
        cmndbtnEdit_Click
    End If
    If dtVendas.Recordset!situacao_ven = "2" Then
        cmndbtnFinalizarVenda_Click
    ElseIf dtVendas.Recordset!situacao_ven = "3" Then
        cmndbtnFinalizar_Click
    
    End If
    Exit Sub
trata_Erro:
    
End Sub

Private Sub DBGrid1_GotFocus()
    DBGrid1.MarqueeStyle = 3
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        DBGrid1_DblClick
    End If
End Sub

Private Sub DBGrid1_LostFocus()
    DBGrid1.MarqueeStyle = 0
End Sub

Private Sub dtVendas_Reposition()
    On Error Resume Next
    Dim strSQL As String
    If dtVendas.Recordset.RecordCount = 0 Then
        cmndbtnEdit.Enabled = False
        cmndbtnImpr.Enabled = False
        Exit Sub
    Else
        cmndbtnImpr.Enabled = True
        cmndbtnEdit.Enabled = True
    End If
    If dtVendas.Recordset!situacao_ven > 1 Then
        cmndbtnEdit.Enabled = False
        
    Else
        cmndbtnEdit.Enabled = True
        
    End If
    
    strSQL = F_BuscaUmCampo("ip_ven", "tab_vendaaberta", "num_ven='" & dtVendas.Recordset!num_ven & "'")
    If strSQL = "" Then
        lblVendaAberta.caption = "..."
    Else
        lblVendaAberta.caption = "Esta venda está sendo utilizada pelo Terminal " & strSQL
    End If
    buscaCrediario
    buscaVenFec
    
End Sub
Sub buscaVenFec()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    On Error GoTo trata_Erro:
    lblVenFec.caption = ""
    strSQL = "select * from tab_venfec where num_ven='" & dtVendas.Recordset!num_ven & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        strSQL = ""
        Do While Not .EOF
            strSQL = strSQL & !esp_cai & ": " & FormatVal(!valor_ven) & " "
            .MoveNext
        Loop
        .Close
    End With
    lblVenFec.caption = strSQL
    Exit Sub
trata_Erro:
End Sub
Sub buscaCrediario()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error Resume Next
    strSQL = "select * from tab_Crediarioven where num_ven='" & dtVendas.Recordset!num_ven & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            lblCrediario.caption = "Crediário. Status=" & !status_cre & " " & !avaliacao_cre & "  Data: " & !data_cre & " Fechamento: " & !fechamento_Cre & " Usuário: " & F_BuscaUmCampo("guerra_fun", "tab_funcionarios", "cod_fun='" & !cod_usu & "'")
        Else
        
            lblCrediario.caption = ""
        End If
        .Close
    End With
End Sub

Private Sub Form_Activate()
    buscaTabela
    
    
    
    
    
    P_Busca
End Sub
Sub buscaTabela()
    If Me.tag = "U" Then
        sTab = "tab_funcionarios"
        sCod = "cod_fun"
        sNome = "nome_fun"
        Me.caption = "Edição de Vendas de Funcionários"
    Else
        sTab = "tab_clientes"
        sCod = "cod_cli"
        sNome = "nome_cli"
    End If
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        If cmbxSituacao_ven.ListCount <= 2 Then
            cmbxSituacao_ven.AddItem "3- Canceladas"
            
            cmbxSituacao_ven.AddItem "4- Efetivadas"
        End If
    End If
    If KeyCode = vbKeyF5 Then
        P_Busca
    End If
    If KeyCode = vbKeyF9 Then
        If pcVenda.BackColor <> vbYellow Then
            pcVenda.BackColor = vbYellow
        Else
            pcVenda.BackColor = vbWhite
        End If
    End If
End Sub

Private Sub Form_Load()
    resizeForm Me
    txtbxDe.text = Date
    txtbxAte.text = Date
'    adc.ConnectionString = Conn.ConnectionString
    
    
    Set AdoVendas = New ADODB.Recordset
    
    
    cmbxSituacao_ven.ListIndex = 0
    If strVisualizarvenda_Usu <> "S" Then
        cmbxSituacao_ven.Enabled = False
    End If
    'cmbxBusca.ListIndex = 0
    cmbxOrd.ListIndex = 1
    
    
    bActive = True
End Sub
Sub P_Busca()
    Dim strSQL As String
    Dim sBusca As String
    Dim sOrd As String
    buscaTabela
    
    If cmbxOrd.ListIndex = 0 Then
        sOrd = "nomecli_ven"
    ElseIf cmbxOrd.ListIndex = 1 Then
        sOrd = "tv.num_ven"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrd = "data_ven"
    Else
        sOrd = "dataefet_ven"
    End If
    
    
    strSQL = "select distinct tfo.nome_for,tc." & sNome & " as nome_cli, "
    strSQL = strSQL & " tv.num_ven,tv.nomecli_ven,tv.valor_ven,"
    strSQL = strSQL & "tv.desconto_ven,tv.situacao_ven,tv.data_ven,tv.dataefet_ven,tv.*,"
    strSQL = strSQL & " tv.abrevenda_usu,tv.valor_ven+tv.desconto_ven as total, "
    strSQL = strSQL & " tf.guerra_fun "
    strSQL = strSQL & " from " & sTab & " as tc, "
    strSQL = strSQL & " tab_vendas as tv, tab_funcionarios as tf,tab_formaspag as tfo  "
    If Left(cmbxSituacao_ven.text, 1) = "4" And pcVenda.BackColor <> vbYellow Then
        strSQL = strSQL & ", tab_caixa as tx "
    End If
    
    If txtProd.tag <> "" Then
        strSQL = strSQL & " , tab_itvendas as ti "
                
    End If
    
    
    strSQL = strSQL & " where tf.cod_fun=tv.abrevenda_usu "
    
    If txtProd.tag <> "" Then
        strSQL = strSQL & " and tv.num_ven=ti.num_ven "
        
        strSQL = strSQL & " and ti.cod_pro='" & txtProd.tag & "'"
    End If
    If txtbxExpr.text <> "" Then
        strSQL = strSQL & " and (nomecli_ven like '" & txtbxExpr.text & "%' "
        strSQL = strSQL & " or num_ven like '" & txtbxExpr.text & "%' "
    
    End If
    If Left(cmbxSituacao_ven.text, 1) = "4" And pcVenda.BackColor <> vbYellow Then
        strSQL = strSQL & " and tx.num_ven=tv.num_ven "
        
        If pcVenda.BackColor <> vbYellow Then
            strSQL = strSQL & " and tx.visivel_cai='S' "
        End If
        
    End If
    
    strSQL = strSQL & " and tf.cod_fun=tv.abrevenda_usu "
    strSQL = strSQL & " and tfo.cod_for=tv.cod_for "
    strSQL = strSQL & " and tv.situacao_ven='" & Left(cmbxSituacao_ven.text, 1) & "' "
    strSQL = strSQL & " and tc." & sCod & "=tv.cod_cli "
    If txtbxVendedor.tag <> "" Then
        strSQL = strSQL & " and abrevenda_usu='" & txtbxVendedor.tag & "' "
    End If
    
    'If UCase(App.Title) = "PIKIBOX_CX" Then
    '    strSQL = strSQL & " and tv.cod_cai='" & strCod_Cai & "' "
    'End If
    
    If txtbxDe.text <> "" Then
        If Left(cmbxSituacao_ven.text, 1) = "4" And pcVenda.BackColor = vbYellow Then
            strSQL = strSQL & " and dataefet_ven >=" & F_GravaData(txtbxDe.text)
        Else
            strSQL = strSQL & " and data_ven >=" & F_GravaDataTempo(txtbxDe.text)
        End If
    End If
    If txtbxAte.text <> "" Then
        If Left(cmbxSituacao_ven.text, 1) = "4" And pcVenda.BackColor = vbYellow Then
            strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtbxAte.text)
        Else
            strSQL = strSQL & " and data_ven <=" & F_GravaDataTempo(txtbxAte.text & " 23:59:59")
        End If
    End If
    
    strSQL = strSQL & " order by " & sOrd & " desc "
    
    FechaObjeto AdoVendas
    AdoVendas.Open strSQL, Conn_Client
    
    Set dtg.DataSource = AdoVendas
    'adc.RecordSource = strSQL
    'adc.Refresh
    P_Status ""
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    dtg.width = Me.width - 200
    dtg.height = pcBot.top - dtg.top - 400
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    dtVendas.Recordset.Close
End Sub

Private Sub tmr_Timer()
    If cmbxSituacao_ven.ListIndex = 0 Then
        RefreshDataPosition adc
    End If
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
End Sub

Private Sub txtbxExpr_DblClick()
    P_Busca
End Sub

Private Sub txtbxExpr_GotFocus()
    P_SelecTexto txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    
        P_Busca
        
    End If
    If KeyCode = vbKeyDown Then
        
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    P_AjustaTexto txtbxExpr
End Sub

Private Sub txtbxVendedor_Change()
    txtbxVendedor.tag = ""
End Sub

Private Sub txtbxVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaFun
    End If
End Sub
Sub buscaFun()
    Dim f As funcionario.funcionario
    f = F_SelectVendedor(txtbxVendedor.text)
    txtbxVendedor.text = f.guerra_fun
    txtbxVendedor.tag = f.cod_fun
    
End Sub

Private Sub txtProd_Change()
    txtProd.tag = ""
End Sub
