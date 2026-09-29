VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmProdProm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relação dos Produtos em Promoção"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7905
   Icon            =   "frmProdProm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmndbtnExcl 
      Caption         =   "&Excluir"
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmndbtnAdd 
      Caption         =   "&Adicionar"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmndbtnImpr 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4200
      Picture         =   "frmProdProm.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   6000
      Picture         =   "frmProdProm.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox txtbxFinPro 
      Height          =   285
      Left            =   4200
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txtbxProd 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   1440
      Top             =   2880
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Caption         =   "..."
      Height          =   285
      Left            =   3600
      TabIndex        =   6
      Top             =   5520
      Width           =   375
   End
   Begin MSDataGridLib.DataGrid dbgrd 
      Bindings        =   "frmProdProm.frx":06D6
      Height          =   3855
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6800
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
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4394,835
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   330
      Left            =   240
      Top             =   3720
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Até:"
      Height          =   195
      Left            =   4200
      TabIndex        =   3
      Top             =   5280
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Produto:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Width           =   600
   End
End
Attribute VB_Name = "frmProdProm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnIr_Click()

End Sub

Private Sub cmndbtnAdd_Click()
    If txtbxProd.tag = "" Then
        MsgBox "Escolha o produto primeiramente!", vbExclamation
        txtbxProd.SetFocus
        Exit Sub
    End If
    If CDate(txtbxFinPro.Text) < Date Then
        MsgBox "A Data Final deve ser posterior ou igual a hoje!", vbExclamation
        txtbxFinPro.SetFocus
        Exit Sub
    End If
    strSQL = "delete from tab_prodpromo where cod_pro='" & txtbxProd.tag & "'"
    conn.Execute strSQL
    strSQL = "insert into tab_prodpromo(cod_pro,datafin_pro)values("
    strSQL = strSQL & "'" & txtbxProd.tag & "',"
    strSQL = strSQL & F_GravaData(txtbxFinPro.Text) & ") "
    conn.Execute strSQL
    Timer1_Timer
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim p As Produto
    p = F_BuscaProdSaida(txtbxProd.Text)
    If p.cod_pro <> "" Then
        txtbxProd.Text = p.nome_pro
        txtbxProd.tag = p.cod_pro
    End If
    
    
    
End Sub

Private Sub cmndbtnExcl_Click()
    Dim strSQL As String
    On Error GoTo trata_erro
    strSQL = "delete from tab_prodpromo where cod_pro='" & dtPro.Recordset!cod_pro & "'"
    conn.Execute strSQL
    Timer1_Timer
    Exit Sub
trata_erro:
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtbxFinPro.Text = Date
    dtPro.ConnectionString = strConnect
    P_CentraForm Me
    Timer1_Timer
End Sub

Private Sub Timer1_Timer()
    Dim strSQL As String
    Dim i As Integer
    On Error Resume Next
    strSQL = "select tp.nome_pro,tp.cod_Pro,tp.codfab_pro,tp.pv_pro,tpp.datafin_pro "
    strSQL = strSQL & " from " & TAB_PRODUTOS & " as tp,tab_prodpromo as tpp"
    strSQL = strSQL & " where tp.cod_pro=tpp.cod_pro "
    strSQL = strSQL & " and tpp.datafin_pro >=" & F_GravaData(Date)
    
        strSQL = strSQL & " order by tp.nome_pro "
    
    
    i = dtPro.Recordset.AbsolutePosition
    dtPro.RecordSource = strSQL
    
    dtPro.Refresh
    If i > -1 Then
        dtPro.Recordset.Move i
    End If
End Sub

Private Sub txtbxFinPro_GotFocus()
    ganhaFoco txtbxFinPro
End Sub

Private Sub txtbxFinPro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxFinPro_LostFocus()
    P_AjustaData txtbxFinPro
End Sub

Private Sub txtbxProd_GotFocus()
    ganhaFoco txtbxProd
End Sub

Private Sub txtbxProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxProd_LostFocus()
    AjustaTexto txtbxProd
End Sub
