VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiVendasAntigas 
   Caption         =   "Venda ainda Não Migradas"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11535
   Icon            =   "frmEdiVendasAntigas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid dtgrd 
      Height          =   4335
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "Vendas Localizadas"
      ColumnCount     =   5
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
         DataField       =   "data_ven"
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
         DataField       =   "valor_ven"
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
      BeginProperty Column03 
         DataField       =   "cod_usu"
         Caption         =   "Usuário"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2520
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3809,764
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   11505
      TabIndex        =   0
      Top             =   5025
      Width           =   11535
      Begin VB.CommandButton cmdDanfe 
         Caption         =   "Gerar DANFE"
         Height          =   615
         Left            =   6600
         TabIndex        =   9
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtNumero 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   615
         Left            =   10080
         TabIndex        =   11
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdCFCe 
         Caption         =   "Gerar NFC-e"
         Height          =   615
         Left            =   5280
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdiNFE 
         Caption         =   "Abrir NFEs"
         Height          =   615
         Left            =   8040
         TabIndex        =   10
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "Pesquisar"
         Height          =   615
         Left            =   3960
         TabIndex        =   7
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtCaixa 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         Height          =   195
         Left            =   2040
         TabIndex        =   5
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         Height          =   195
         Left            =   1080
         TabIndex        =   3
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmEdiVendasAntigas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdCFCe_Click()
    Dim id As String
    Dim cpf As String
    cpf = InputBox("Informe o CPF", "CPF")
    
    cpf = F_SoNumero(cpf)
    
    ImportarVendaMigracaoNfe ado!num_ven, cpf
    
    
    
    id = NFe_DAO.GravaNFCeVenda(ado!num_ven, strTipo_IcmsNFe, "Venda de Produtos", NFCe_65, True)
    If id <> "" Then
        If MsgBox("Geração concluída. Deseja transmitir NFC-e?", vbYesNo + vbQuestion) = vbYes Then
            TransmitirNFcE id, ""
        End If
        
    End If
End Sub

Private Sub cmdDanfe_Click()
    
    
        Dim id As String
    Dim cpf As String
    cpf = InputBox("Informe o CPF", "CPF")
    
    cpf = F_SoNumero(cpf)
    
    ImportarVendaMigracaoNfe ado!num_ven, cpf

    
    id = NFe_DAO.GravaNFCeVenda(ado!num_ven, strTipo_IcmsNFe, "Venda de Produtos", NFe_55, True)
    If id <> "" Then
        If MsgBox("Geração concluída. Deseja transmitir NFC-e?", vbYesNo + vbQuestion) = vbYes Then
            TransmitirNFcE id, ""
        End If
        
    End If
End Sub

Private Sub cmdEdiNFE_Click()
    frmEdiNFe.Show 1
End Sub

Private Sub cmdPesquisar_Click()
    busca
End Sub
Sub busca()

    Dim strSQL As String
    Dim c As ADODB.Connection
        
    FechaObjeto c
    FechaObjeto ado
        
        
    Set c = New ADODB.Connection
    
    
    c.Open "muitomais"
    c.CursorLocation = adUseClient
    strSQL = "select tv.*,tc.nome_cli from tab_vendas as tv left join tab_clientes as tc on tv.cod_cli=tc.cod_cli "
    strSQL = strSQL & " where dataefet_ven=" & F_GravaData(txtData.Text)
    If txtCaixa.Text <> "" Then
        strSQL = strSQL & " and cod_cai='" & txtCaixa.Text & "'"
    End If
    If txtNumero.Text <> "" Then
        strSQL = strSQL & " and num_ven='" & txtNumero.Text & "' "
    End If
    strSQL = strSQL & " order by num_ven desc "
    
    ado.Open strSQL, c
    Set dtgrd.DataSource = ado
    
         
    
    
    
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    busca
End Sub

Private Sub Form_Load()
    txtCaixa.Text = strCod_Cai
    txtData.Text = Date
    Set ado = New ADODB.Recordset
End Sub
