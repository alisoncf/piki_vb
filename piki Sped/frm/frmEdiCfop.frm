VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiCfop 
   Caption         =   "Edição de CFOP"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   Icon            =   "frmEdiCfop.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   6900
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc adc 
      Height          =   330
      Left            =   240
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.TextBox txtbxExpr 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Top             =   4440
      Width           =   3135
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   6870
      TabIndex        =   0
      Top             =   4845
      Width           =   6900
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   5520
         Picture         =   "frmEdiCfop.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Height          =   615
         Left            =   4080
         Picture         =   "frmEdiCfop.frx":04F5
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnSelect 
         Caption         =   "&Selecionar"
         Height          =   615
         Left            =   2760
         Picture         =   "frmEdiCfop.frx":087F
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnEditar 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   1440
         Picture         =   "frmEdiCfop.frx":09DB
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnNovo 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiCfop.frx":13DD
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo:"
      Height          =   615
      Left            =   4080
      TabIndex        =   8
      Top             =   4200
      Width           =   2775
      Begin VB.OptionButton optnSaida 
         Caption         =   "Saída"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optnEntrada 
         Caption         =   "Entrada"
         Height          =   195
         Left            =   1080
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid dbgrd 
      Bindings        =   "frmEdiCfop.frx":1541
      Height          =   3855
      Left            =   0
      TabIndex        =   11
      Top             =   240
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6800
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblBuscar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buscar"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   4440
      Width           =   855
   End
End
Attribute VB_Name = "frmEdiCfop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEditar_Click()
    On Error GoTo trata_erro
    With adc.Recordset
        frmCadCFOP.tag = !cfop_cfo
        frmCadCFOP.txtbxCFOP.Text = !cfop_cfo
        frmCadCFOP.txtbxNome.Text = !nome_cfo
        If !tipo_cfo = "S" Then
            frmCadCFOP.optnSaida.Value = True
        Else
            frmCadCFOP.optnEntrada.Value = True
        End If
        frmCadCFOP.Show 1
        busca
    End With
    Exit Sub
trata_erro:
    MostraErro
End Sub

Private Sub cmndbtnExcl_Click()
    On Error GoTo trata_erro
    If MsgBox("Deseja realmente excluir?", vbYesNo + vbCritical) = vbNo Then
        Exit Sub
    End If
    P_DeleteCFOP adc.Recordset!cfop_cfo
    Exit Sub
trata_erro:
    
    
End Sub

Private Sub cmndbtnNovo_Click()
    frmCadCFOP.Show 1
    busca
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Me.Hide
    Else
        Unload Me
    End If
End Sub

Private Sub cmndbtnSelect_Click()
    Me.tag = "x"
    Me.Hide
End Sub

Sub busca()
    Dim strSQL As String
    On Error GoTo trata_erro:
    strSQL = F_SelectCFOP(txtbxExpr.Text, IIf(optnSaida.Value, "S", "E"))
    
    adc.RecordSource = strSQL
    adc.Refresh
    Exit Sub
trata_erro:
    MostraErro

End Sub

Private Sub dbgrd_DblClick()
    If Me.tag = "" Then
        cmndbtnEditar_Click
    Else
        cmndbtnSelect_Click
    End If
End Sub

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dbgrd_DblClick
    End If
End Sub

Private Sub Form_Activate()
    busca
End Sub

Private Sub Form_Load()
    adc.ConnectionString = strConnect
    
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
    
End Sub
