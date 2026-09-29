VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmHistCli 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Histórico de Produtos Adquiridos pelo Cliente"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "Im&primir"
      Height          =   735
      Left            =   5520
      Picture         =   "frmHistCli.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   7080
      Picture         =   "frmHistCli.frx":040D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordem:"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   5640
      Width           =   2775
      Begin VB.OptionButton Option1 
         Caption         =   "&Produto"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optnData 
         Caption         =   "&Data"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   375
      Left            =   1440
      Top             =   5760
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
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
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmHistCli.frx":0557
      Height          =   5535
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9763
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
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
End
Attribute VB_Name = "frmHistCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    Dim iPag As Integer
    P_CabRelTimbr "Histórico de Cliente", obj
    P_ImprX obj, "Cliente: " & Me.caption, 5, False
    P_Line obj, 0, obj.CurrentY, obj.ScaleWidth, obj.CurrentY
    
    P_ImprX obj, "Data", 2
    P_ImprX obj, "Venda", 20
    P_ImprX obj, "Ref.", 35
    P_ImprX obj, "Produto", 60
    P_ImprX obj, "Qtd.", 150, True, True
    P_ImprX obj, "Unitário", 170, True, True
    P_ImprX obj, "Usuário", 175, False
    P_line1 obj
    With dtVendas.Recordset
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If obj.CurrentY > obj.ScaleHeight - 30 Then
                 
                P_RodapeVis obj, "", (iPag + 1)
                iPag = iPag + 1
                
                P_AddPicVis obj
                
                P_ImprX obj, "Data", 2
                P_ImprX obj, "Venda", 20
                P_ImprX obj, "Ref.", 35
                P_ImprX obj, "Produto", 60
                P_ImprX obj, "Qtd.", 150, True, True
                P_ImprX obj, "Unitário", 170, True, True
                P_ImprX obj, "Usuário", 175, False
                P_line1 obj
                
            End If
            P_ImprX obj, !dataefet_ven, 2
            P_ImprX obj, !seqefet_ven, 20
            P_ImprX obj, !codfab_pro, 35
            P_ImprX obj, StrConv(!nome_pro, vbProperCase), 60
            P_ImprX obj, !qtd_ite, 150, True, True
            P_ImprX obj, !unit_ite, 170, True, True
            P_ImprX obj, !guerra_fun, 175, False
            
            .MoveNext
        Loop
    End With
    
    
    P_RodapeVis obj
    P_AddPicVis obj
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Busca
End Sub

Private Sub Form_Load()
    dtVendas.ConnectionString = strConnect
    P_CentraForm Me
End Sub
Sub Busca()
    Dim strSQL As String
    strSQL = "select tf.guerra_fun,tc.cod_cli,tc.nome_cli,data_ven,dataefet_ven,"
    strSQL = strSQL & " tp.*, ti.unit_ite,ti.qtd_ite "
    strSQL = strSQL & " from " & TAB_VENDAS & " as tv, " & TAB_ITVENDAS & " as ti,"
    strSQL = strSQL & " " & TAB_PRODUTOS & " as tp," & TAB_CLIENTES & " as tc,tab_funcionarios as tf  "
    strSQL = strSQL & " where ti.num_ven=tv.num_ven "
    strSQL = strSQL & " and tc.cod_Cli=tv.cod_cli "
    strSQL = strSQL & " and tc.cod_cli='" & Me.tag & "' "
    strSQL = strSQL & " and ti.cod_pro=tp.cod_pro "
    strSQL = strSQL & " and tf.cod_fun=tv.fechavenda_usu "
    If optnData.Value Then
        strSQL = strSQL & " order by dataefet_ven,nome_pro "
    Else
        strSQL = strSQL & " order by nome_pro,Ref_pro,dataefet_ven "
    End If
    dtVendas.RecordSource = strSQL
    dtVendas.Refresh
    
    With dtVendas.Recordset
        dbgrd.caption = .RecordCount & " Registros encontrados"
        If .RecordCount > 0 Then
            Me.caption = !cod_cli & "-" & !nome_Cli
        End If
    End With
    End Sub

Private Sub Option1_Click()
    Busca
End Sub

Private Sub optnData_Click()
    Busca
End Sub
