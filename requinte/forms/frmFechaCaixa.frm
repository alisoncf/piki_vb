VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFechaCaixa 
   Caption         =   "Fechamento de Caixa"
   ClientHeight    =   8190
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   13275
   Icon            =   "frmFechaCaixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   13275
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcDir 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6360
      Left            =   13140
      ScaleHeight     =   6360
      ScaleWidth      =   135
      TabIndex        =   26
      Top             =   855
      Width           =   135
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   13245
      TabIndex        =   15
      Top             =   7215
      Width           =   13275
      Begin VB.PictureBox pcTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         ScaleHeight     =   705
         ScaleWidth      =   3465
         TabIndex        =   19
         Top             =   120
         Width           =   3495
         Begin VB.Label lblTotalEfetivo 
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
            Left            =   1200
            TabIndex        =   25
            Top             =   0
            Width           =   390
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Efetivo:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   0
            Width           =   540
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Simbólico:"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   720
         End
         Begin VB.Label lblTotalSimb 
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
            Left            =   1200
            TabIndex        =   22
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Geral:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   420
         End
         Begin VB.Label lblTotalGeral 
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
            Left            =   1200
            TabIndex        =   20
            Top             =   480
            Width           =   510
         End
      End
      Begin VB.CommandButton cmndbtnFecharCaixa 
         Caption         =   "&Fechar Caixa"
         Height          =   735
         Left            =   3720
         Picture         =   "frmFechaCaixa.frx":1084A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   5520
         Picture         =   "frmFechaCaixa.frx":10994
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSair 
         Caption         =   "&Sair"
         Height          =   735
         Left            =   11400
         Picture         =   "frmFechaCaixa.frx":10BF0
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.PictureBox pcTopo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   13245
      TabIndex        =   4
      Top             =   0
      Width           =   13275
      Begin VB.TextBox txtbxData 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cmbxCaixa 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox cmbxTurno 
         Height          =   315
         ItemData        =   "frmFechaCaixa.frx":1113E
         Left            =   3600
         List            =   "frmFechaCaixa.frx":11140
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cmbxUsu 
         Height          =   315
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "Buscar"
         Height          =   495
         Left            =   8640
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkBaixa 
         Caption         =   "Baixa"
         Height          =   255
         Left            =   7320
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa:"
         Height          =   195
         Left            =   1320
         TabIndex        =   13
         Top             =   120
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno:"
         Height          =   195
         Left            =   3600
         TabIndex        =   12
         Top             =   120
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário:"
         Height          =   195
         Left            =   4680
         TabIndex        =   11
         Top             =   120
         Width           =   585
      End
   End
   Begin TabDlg.SSTab sstb 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Resumo"
      TabPicture(0)   =   "frmFechaCaixa.frx":11142
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "gridResumo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalhado"
      TabPicture(1)   =   "frmFechaCaixa.frx":1115E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "gridDet"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Baixa de Carnês"
      TabPicture(2)   =   "frmFechaCaixa.frx":1117A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "gridBaixa"
      Tab(2).ControlCount=   1
      Begin MSDataGridLib.DataGrid gridBaixa 
         Height          =   5415
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   9551
         _Version        =   393216
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
         Caption         =   "Cobrança Resumida"
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "cod"
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
            DataField       =   "cobranca"
            Caption         =   "Cobrança"
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
            DataField       =   "total"
            Caption         =   "Total"
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
               ColumnWidth     =   1124,787
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4559,811
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1289,764
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridResumo 
         Height          =   5535
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   9763
         _Version        =   393216
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
         Caption         =   "Movimento Resumido"
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "sigla_cai"
            Caption         =   "Sigla"
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
            DataField       =   "descricao_cai"
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
            DataField       =   "entrada_cai"
            Caption         =   "Entrada"
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
            DataField       =   "saida_cai"
            Caption         =   "Saída"
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
            DataField       =   "total"
            Caption         =   "Total"
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
               ColumnWidth     =   824,882
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4844,977
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1019,906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1094,74
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridDet 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   9128
         _Version        =   393216
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
         Caption         =   "Movimento Detalhado"
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "seq_cai"
            Caption         =   "SEQ"
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
            DataField       =   "esp_cai"
            Caption         =   "Espécie"
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
         BeginProperty Column03 
            DataField       =   "valor_cai"
            Caption         =   "Valor"
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
            DataField       =   "entrada_cai"
            Caption         =   "Entrada"
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
            DataField       =   "saida_cai"
            Caption         =   "Saida"
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
         BeginProperty Column07 
            DataField       =   "guerra_fun"
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
         BeginProperty Column08 
            DataField       =   "hora_cai"
            Caption         =   "Hora"
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
               ColumnWidth     =   854,929
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   780,095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   4529,764
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   569,764
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1679,811
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmFechaCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Public WithEvents adoBx As ADODB.Recordset
Attribute adoBx.VB_VarHelpID = -1
Public WithEvents adoResumo As ADODB.Recordset
Attribute adoResumo.VB_VarHelpID = -1


Private Sub chkBaixa_Click()
    If chkBaixa.Value = 0 Then
        sstb.TabVisible(2) = False
    Else
        sstb.TabVisible(2) = True
    End If
End Sub

Private Sub cmbxCaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxCaixa.ListIndex = -1
    End If
End Sub

Private Sub cmbxUsu_Click()
    cmbxUsu.tag = F_BuscaUmCampo("cod_fun", "tab_funcionarios", "guerra_fun='" & cmbxUsu.Text & "'")
End Sub

Private Sub cmbxUsu_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxUsu.ListIndex = -1
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    P_Status "Aguarde..."
    busca
    P_Status ""
End Sub

Private Sub cmndbtnImprimir_Click()
    If sstb.Tab = 0 Then
        P_DtgrdToHtml "Resumo de Movimento de Caixa", gridResumo, adoResumo, _
        "Data: " & txtbxData.Text & " Caixa: " & cmbxCaixa.Text & " Usu:" & cmbxUsu.Text, 4
    
    ElseIf sstb.Tab = 1 Then
        P_DtgrdToHtml "Movimento de Caixa Detalhado", gridDet, ado, _
        "Data: " & txtbxData.Text & " Caixa: " & cmbxCaixa.Text & " Usu:" & cmbxUsu.Text, "4,5"
    Else
        P_DtgrdToHtml "Resumo das Baixas", gridBaixa, adoBx, _
        "Data: " & txtbxData.Text & " Caixa: " & cmbxCaixa.Text & " Usu:" & cmbxUsu.Text, 2
    End If
        
        
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF9 Then
        If pcBot.BackColor = vbYellow Then
            pcBot.BackColor = vbWhite
        Else
            pcBot.BackColor = vbYellow
        End If
    Else
    
    End If
End Sub

Private Sub Form_Load()
    P_PreencheFunc cmbxUsu, 1
    P_PreencheCaixas cmbxCaixa
    P_PreencheTurnos cmbxTurno
    
    txtbxData.Text = Date
    
    Set ado = New ADODB.Recordset
    Set adoBx = New ADODB.Recordset
    Set adoResumo = New ADODB.Recordset
    
    
    
    chkBaixa_Click
    P_LocalizaOperacao strCod_Cai, cmbxCaixa
    
    sstb.Tab = 0
End Sub
Sub ajustaGrid()
    
End Sub
Private Sub cmndbtnFecharCaixa_Click()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim sCod_Con  As String
    Dim id_fin As String
'    If pcECF.BackColor = vbGreen Then
'        MsgBox "Impossível fechar esse caixa. Observe os valores e as cores!", vbExclamation
'        Exit Sub
'    End If
    If sstb.Tab = 0 Then
'        If VerificaVendaAbertaNoCaixa(Left(cmbxCaixa.Text, 2)) Then
'            MsgBox "Não é possível fechar o caixa pois existem vendas em aberto! Procure o responsável!", vbCritical
'            Exit Sub
'
'        End If
    End If
    'On Error GoTo trata_erro
    strSQL = F_BuscaUmCampo("count(*)", "tab_vendas", "situacao_ven=2 and cod_cai='" & strCod_Cai & "'")
    If Int(strSQL) > 0 Then
        MsgBox "Não é possível fechar o caixa pois existem vendas a efetivar!", vbExclamation
        Exit Sub
    End If
'    If F_BuscaUmCampo("data_cai", "tab_fechacaixa", "data_cai=" & _
'        F_PesqData(txtbxData.Text) & " and cod_cai='" & Left(cmbxCaixa.Text, 2) & "'") <> "" Then
'        MsgBox "Este caixa já foi fechado neste dia!", vbExclamation
'        Exit Sub
'    End If

    'If CCur(lblTotalDin.caption) - CCur(txtbxFundoCaixa.Text) < 0 Then
    '    MsgBox "O fundo de caixa está maior que o valor disponível!", vbExclamation
    '    Exit Sub
    'End If

'    If VendasFechadas = True Then
'        MsgBox "Existem vendas condicionais a serem efetivadas!", vbExclamation
'        On Error Resume Next
'        frmEfetivaVen.Show 1
'        Exit Sub
'    End If
    If MsgBox("Deseja realmente fechar o Turno do dia " & txtbxData.Text _
        & " Com " & " de Fundo de caixa?", vbYesNo + vbCritical) = vbYes Then
'        If CCur(txtbxFundoCaixa.Text) = 0 Then
'
'            If MsgBox("Este fechamento não contém fundo de caixa!" & Chr(13) & _
'                "Deseja realmente continuar?", vbYesNo + vbQuestion) = vbYesNo Then
'                Exit Sub
'            End If
'        End If
        If F_BuscaUmCampo("status_tur", "tab_turncai", "cod_Cai='" & Left(cmbxCaixa.Text, 2) & "' and cod_tur='" & Left(cmbxTurno.Text, 1) & "' and data_tur=" & F_GravaData(txtbxData.Text)) = "F" Then
            MsgBox "Este turno já se encontra Fechado!", vbCritical
            Exit Sub
         End If
'        strSQL = "insert into tab_fechacaixa(data_Cai,valor_FEC,fundo_cai,cod_usu,COD_CAI)values("
'        strSQL = strSQL & F_GravaData(txtbxData.Text) & ","
'        strSQL = strSQL & F_GravaDec(lblTotalAB.caption) & ","
'        strSQL = strSQL & F_GravaDec(txtbxFundoCaixa.Text) & ","
'        strSQL = strSQL & "'" & strCod_Usu & "','" & Left(cmbxCaixa.Text, 2) & "') "
'        Banco.Execute strSQL
        
        
        strSQL = "update tab_turncai set logout_usu='" & strCod_Usu & "',"
        strSQL = strSQL & "Fechamento_tur=" & F_GravaHora(Now) & ","
        strSQL = strSQL & "valorfechamento_tur=" & F_GravaDec(lblTotalGeral.caption) & ","
        strSQL = strSQL & "status_tur='F' "
        strSQL = strSQL & " where cod_Cai='" & Left(cmbxCaixa.Text, 2) & "' "
        strSQL = strSQL & " and cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
        Conn.Execute strSQL
        
        'strSQL = F_GravaCai("D", txtbxData.Text, CCur(lblTotalDin.caption) - CCur(txtbxFundoCaixa.Text), _
            "TRANSFERÊNCIA DE CAIXA", "", "DIN", "01", 0, 0, Left(cmbxCaixa.Text, 2), strCod_Tur)
        
        id_fin = F_GravaFinanceiro("C", txtbxData.Text, CCur(lblTotalGeral.caption), _
             "TRANSFERÊNCIA DE CAIXA", constContaCaixa, "S", "000")
         
        
            
            
        strSQL = "Update tab_Caixas set status_Cai='F' where cod_Cai='" & Left(cmbxCaixa.Text, 2) & "'"
        Conn.Execute strSQL
        MsgBox "Caixa " & cmbxCaixa.Text & "  Fechado com sucesso!", vbInformation
        
        
    End If
        
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Sub soma()
    Dim cTotal As Currency
    Dim cEfetivo As Currency
    Dim cSimb As Currency
    With adoResumo
        If Not .EOF Then
            .MoveFirst
        End If
        Do While Not .EOF
            
            If !sigla_cai = "DUP" Or !sigla_cai = "TRO" Or !sigla_cai = "DEV" Or !sigla_cai = "TRA" Then
                cSimb = cSimb + !total
            Else
                cEfetivo = cEfetivo + !total
            End If
            cTotal = cTotal + !total
            .MoveNext
        Loop
        If Not .EOF Then
            .MoveFirst
        End If
    End With
    lblTotalEfetivo.caption = FormatVal(cEfetivo)
    lblTotalGeral.caption = FormatVal(cTotal)
    lblTotalSimb.caption = FormatVal(cSimb)
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cCred As Currency
    Dim sSigla As String
    Dim cDeb As Currency
    Dim sCab As String
    strSQL = "drop table if exists tab_fccaitmp "
    Conn.Execute strSQL
    strSQL = "create table tab_fccaitmp( data_cai date,"
    
    strSQL = strSQL & " sigla_cai char(3),"
    strSQL = strSQL & " descricao_cai char(40),"
    strSQL = strSQL & " entrada_cai double(14,2) ,"
    strSQL = strSQL & " saida_cai double(14,2) ) "
    
    
    
    Conn.Execute strSQL
    
    strSQL = " select tc.data_cai,if(tc.tipo_cai='C','+','-') as sinal,tc.tipo_cai,sum(tc.valor_cai) as total,te.sigla_cai,te.descricao_cai from tab_caixa as tc,"
    strSQL = strSQL & " tab_espcai as te "
    strSQL = strSQL & " where data_cai=" & F_GravaData(txtbxData.Text)
    strSQL = strSQL & " and te.sigla_cai=tc.esp_cai "

    
    
    If cmbxCaixa.Text <> "" Then
        strSQL = strSQL & " and cod_Cai='" & Left(cmbxCaixa.Text, 2) & "' "
    End If
    If cmbxTurno.Text <> "" Then
        strSQL = strSQL & " and cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
    End If
    If cmbxUsu.tag <> "" Then
        strSQL = strSQL & " and cod_usu='" & cmbxUsu.tag & "' "
    End If
    strSQL = strSQL & " group by te.sigla_cai,tc.tipo_cai "
    strSQL = strSQL & " order by te.sigla_cai,tc.tipo_cai "
    
    abreObjetoAdo rPesq, strSQL
    sSigla = ""
    sCab = ""
    cCred = 0
    cDeb = 0
    With rPesq
        Do While Not .EOF
            
            If sSigla <> !sigla_cai Then
                
                If sCab <> "" Then
                    'Banco.Execute strSQL
                End If
                sSigla = !sigla_cai
                If sCab = "" Then
                    sCab = "insert into tab_fccaitmp(data_cai,sigla_cai,descricao_cai,saida_cai,entrada_cai)values("
                    sCab = sCab & F_GravaData(!data_cai) & ","
                    sCab = sCab & "'" & sSigla & "',"
                    sCab = sCab & "'" & !descricao_cai & "',"
                Else
                    sCab = sCab & F_GravaDec(cDeb) & ","
                    sCab = sCab & F_GravaDec(cCred) & ")"
                    Conn.Execute sCab
                    sCab = ""
                    cDeb = 0
                    cCred = 0
                    sCab = "insert into tab_fccaitmp(data_cai,sigla_cai,descricao_cai,saida_cai,entrada_cai)values("
                    sCab = sCab & F_GravaData(!data_cai) & ","
                    sCab = sCab & "'" & sSigla & "',"
                    sCab = sCab & "'" & !descricao_cai & "',"
                End If
                
                
                
                    
            End If
            If !sinal = "+" Then
                    
                cCred = !total
            Else
                cDeb = !total
            End If
            
            .MoveNext
        Loop
        
        
        If sCab <> "" Then
            sCab = sCab & F_GravaDec(cDeb) & ","
            sCab = sCab & F_GravaDec(cCred) & ")"
            Conn.Execute sCab
        End If
    End With
    
    strSQL = "select tf.*,entrada_cai-saida_cai as total from tab_fccaitmp as tf order by sigla_cai "
    abreObjetoAdo adoResumo, strSQL, gridResumo
    strSQL = "drop table tab_fccaitmp "
    Conn.Execute strSQL
    
    
    strSQL = " select tf.guerra_fun,tc.*,if(tc.tipo_cai='C','+','-') as sinal,te.sigla_cai,te.descricao_cai,hora_cai from tab_caixa as tc inner join "
    strSQL = strSQL & " tab_espcai as te on te.sigla_cai=tc.esp_cai "
    strSQL = strSQL & " inner join tab_funcionarios as tf on tf.cod_fun=tc.cod_usu "
    strSQL = strSQL & " where data_cai=" & F_GravaData(txtbxData.Text)
    
    If cmbxCaixa.Text <> "" Then
        strSQL = strSQL & " and cod_Cai='" & Left(cmbxCaixa.Text, 2) & "' "
    End If
    If cmbxTurno.Text <> "" Then
        strSQL = strSQL & " and cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
    End If
    If cmbxUsu.tag <> "" Then
        strSQL = strSQL & " and cod_usu='" & cmbxUsu.tag & "' "
    End If

    strSQL = strSQL & " order by tc.esp_cai,tc.tipo_cai,tc.seq_cai  "
    abreObjetoAdo ado, strSQL, gridDet
    
    If chkBaixa.Value = 1 Then
    
        strSQL = F_SelectBaixaCarnesPorCaixa(cmbxCaixa.Text, txtbxData.Text, txtbxData.Text, cmbxUsu.tag)
    
        abreObjetoAdo adoBx, strSQL, gridBaixa
        
    End If
    soma
    
    
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    PosicionaGradeVal sstb, pcDir.Left, pcBot.top
     PosicionaGradeVal gridResumo, sstb.width, sstb.height
'     gridBaixa.width = gridResumo.width
'     gridBaixa.height = gridResumo.height
'     gridBaixa.Left = gridResumo.Left
'     gridBaixa.top = gridResumo.top
End Sub

Private Sub txtbxData_GotFocus()
    P_SelecTexto txtbxData
End Sub

Private Sub txtbxData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxData_LostFocus()
    P_AjustaData txtbxData
End Sub
