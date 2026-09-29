VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelVendas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Vendas"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10170
   Icon            =   "frmRelVendas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   10170
   Begin VB.Data dtSai 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   7920
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelVendas.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelVendas.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelVendas.frx":1016
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox cmbxOrderBY 
         Height          =   315
         ItemData        =   "frmRelVendas.frx":1320
         Left            =   120
         List            =   "frmRelVendas.frx":132D
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1065
         Width           =   1935
      End
      Begin VB.TextBox txtbxDAtaFIn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   840
      End
   End
   Begin MSDBGrid.DBGrid dbgrdSai 
      Bindings        =   "frmRelVendas.frx":1350
      Height          =   855
      Left            =   7920
      OleObjectBlob   =   "frmRelVendas.frx":1364
      TabIndex        =   12
      Top             =   4200
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   375
      Left            =   1200
      Top             =   4920
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
      Bindings        =   "frmRelVendas.frx":1D70
      Height          =   5415
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9551
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
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "Total Alcançado:"
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
      Left            =   7920
      TabIndex        =   11
      Top             =   5400
      Width           =   1470
   End
   Begin VB.Label lblTotalGerado 
      AutoSize        =   -1  'True
      Caption         =   "Total Gerado:"
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
      Left            =   7920
      TabIndex        =   10
      Top             =   5160
      Width           =   1185
   End
End
Attribute VB_Name = "frmRelVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Private Sub cmbxOrderBY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    busca
    soma
End Sub

Private Sub cmndbtnImprimir_Click()
    If dt.Recordset.RecordCount <= 0 Then
        MsgBox "Não há o quer ser impresso!", vbExclamation
        Exit Sub
    End If
    'If MsgBox("Confirma Impressão do Relatório?", vbYesNo + vbQuestion) = vbYes Then
        imprime frmVisualiza.Pic
        frmVisualiza.Show 1
        
        If frmVisualiza.tag <> "" Then
        
            imprime Printer
        End If
        Unload frmVisualiza
    'End If
End Sub
Sub imprime(obj As Object)
    Dim iItens As Integer
    Dim iPag As Integer
    Dim strSQL As String
    bActive = True
    P_CabRelTimbr "Relatório de Saídas no Período " & de & txtbxDataIni.Text & _
        " Até " & txtbxdataFin.Text, obj
    P_ImprX obj, "Código", 0
    P_ImprX obj, "Descrição", 20
    P_ImprX obj, "UN", 120
    P_ImprX obj, "QTD", 140, True, True
    P_ImprX obj, "Ger.", 160, True, True
    P_ImprX obj, "Alc.", 180, False, True
    
    
    P_ImprLinha obj
    
    With dt.Recordset
        
        .MoveFirst
        strSQL = CStr(.RecordCount / 50)
        If InStr(1, strSQL, ",") <> 0 Then
            iPag = Int(strSQL) + 1
        End If
        
        Do While Not .EOF
            iItens = iItens + 1
            If iItens = 50 Then
                If TypeName(obj) = "Printer" Then
                    P_Rodape "Continua...", True, "Pág:" & Printer.Page & " de " & iPag
                End If
                iItens = 0
                P_ImprX obj, "Código", 0
                P_ImprX obj, "Descrição", 20
                P_ImprX obj, "UN", 120
                P_ImprX obj, "QTD", 140, True, True
                P_ImprX obj, "Ger.", 160, True, True
                P_ImprX obj, "Alc.", 180, False, True
    
                P_ImprLinha obj
            End If
            P_ImprX obj, !cod_pro, 0
            P_ImprX obj, !nome_pro, 20
            P_ImprX obj, !unidade_pro, 120
            P_ImprX obj, FormatVal(!soma), 140, True, True
            P_ImprX obj, FormatVal(!gerado), 160, True, True
            With dtSai.Recordset
                If .RecordCount > 0 Then
                    P_ImprX obj, FormatVal(!soma), 180, False, True
                Else
                    P_ImprX obj, "0,00", 180, False, True
                End If
            End With
           
            .MoveNext
        Loop
        P_ImprLinha obj
        P_ImprX obj, lblTotalGerado.caption, 80, True
        P_ImprX obj, lblTotal.caption, 130, False
        If TypeName(obj) = "Printer" Then
            P_Rodape strNome_Emp, True, "Pág: " & Printer.Page & " de " & Printer.Page
        End If
    End With
    
    If TypeName(obj) = "Printer" Then
        Printer.EndDoc
    End If
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub dt_Reposition()
    If bActive = True Then Exit Sub
    buscaSaidas
End Sub

Private Sub Form_Load()
    dt.ConnectionString = strConnect
    adc.ConnectionString = strConnect
    cmbxOrderBY.ListIndex = 0
    P_CentraForm Me
End Sub
Sub soma()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim cTotalGer As Currency
    bActive = True
    With dt.Recordset
    
        Do While Not .EOF
           ' cTotal = cTotal + !Total
            cTotalGer = cTotalGer + !gerado
            .MoveNext
        
        Loop
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
    
    lblTotal.caption = "Total: " & FormatVal(F_BuscaUmCampo( _
        "sum(valor_cup)", "cupons", "data_cup>=" & _
            F_GravaDataTempo(txtbxDataIni.Text) & _
            "  and data_cup<=" & F_GravaDataTempo(txtbxdataFin.Text & " 23:59:59")))
    lblTotalGerado.caption = "Total Gerado: " & FormatVal(cTotalGer)
    bActive = False
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select tp.nome_pro,tp.unidade_pro,ti.*,sum(ti.qtd_ite) as soma,"
    strSQL = strSQL & "sum(ti.final_ite) as gerado "
    
    strSQL = strSQL & " from " & TAB_VENDAS & " as tv, " & TAB_ITVENDAS & " as ti,"
    strSQL = strSQL & "" & TAB_PRODUTOS & " as tp where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and ti.num_ven=tv.num_ven "
    
    strSQL = strSQL & " and left(tv.dataefet_ven,10) >=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and left(tv.dataefet_ven,10) <=" & F_GravaData(txtbxdataFin.Text)
    strSQL = strSQL & " and tv.situacao_ven>3 "
    strSQL = strSQL & " group by ti.cod_pro "
    
    strSQL = strSQL & " order by "
    If cmbxOrderBY.ListIndex = 0 Then
        strSQL = strSQL & "tp.nome_pro"
    ElseIf cmbxOrderBY.ListIndex = 1 Then
        strSQL = strSQL & "tp.cod_Pro "
    Else
        strSQL = strSQL & "gerado "
    End If
    dt.RecordSource = strSQL
    dt.Refresh
End Sub
Sub buscaSaidas()
    On Error Resume Next
    Dim strSQL As String
    strSQL = "select tp.nome_pro,tp.unidade_pro,ti.*,sum(ti.qtd_its) as soma,"
    strSQL = strSQL & "sum(ti.qtd_its*ti.unit_its) as gerado "
    
    strSQL = strSQL & " from cupons as ts, tab_itcup as ti,"
    strSQL = strSQL & "" & TAB_PRODUTOS & " as tp where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and ti.seq_cup=ts.seq_cup "
    strSQL = strSQL & " and ts.data_cup >=" & F_GravaDataTempo(txtbxDataIni.Text)
    strSQL = strSQL & " and ts.data_cup <=" & F_GravaDataTempo(txtbxdataFin.Text & " 23:59:59")
    strSQL = strSQL & " and tp.cod_pro='" & dt.Recordset!cod_pro & "'"
    strSQL = strSQL & " group by ti.cod_pro "
    
    strSQL = strSQL & " order by "
    If cmbxOrderBY.ListIndex = 0 Then
        strSQL = strSQL & "tp.nome_pro"
    ElseIf cmbxOrderBY.ListIndex = 1 Then
        strSQL = strSQL & "tp.cod_Pro "
    Else
        strSQL = strSQL & "gerado "
    End If
    dtSai.RecordSource = strSQL
    dtSai.Refresh
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    ganhaFoco txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxdataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    ganhaFoco txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub
