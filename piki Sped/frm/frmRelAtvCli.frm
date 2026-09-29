VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelAtvCli 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Atividade de Clientes por Vendas"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   Icon            =   "frmRelAtvCli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10005
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   9705
      TabIndex        =   12
      Top             =   6120
      Width           =   9735
      Begin VB.Label lblQtd 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   9120
         TabIndex        =   24
         Top             =   120
         Width           =   120
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade:"
         Height          =   195
         Left            =   8160
         TabIndex        =   23
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblDesc 
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
         Left            =   7200
         TabIndex        =   22
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto:"
         Height          =   195
         Left            =   6360
         TabIndex        =   21
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblFinan 
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
         Left            =   5400
         TabIndex        =   20
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Financeira:"
         Height          =   195
         Left            =   4560
         TabIndex        =   19
         Top             =   120
         Width           =   780
      End
      Begin VB.Label lblAp 
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
         Left            =   3720
         TabIndex        =   18
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Prazo:"
         Height          =   195
         Left            =   3000
         TabIndex        =   17
         Top             =   120
         Width           =   600
      End
      Begin VB.Label lblAv 
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
         Left            =   2160
         TabIndex        =   16
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "À Vista:"
         Height          =   195
         Left            =   1560
         TabIndex        =   15
         Top             =   120
         Width           =   540
      End
      Begin VB.Label lblFinal 
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
         Left            =   600
         TabIndex        =   14
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Final:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   2400
      TabIndex        =   5
      Top             =   0
      Width           =   1695
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmRelAtvCli.frx":0442
         Left            =   120
         List            =   "frmRelAtvCli.frx":045B
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar por:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   885
      End
   End
   Begin VB.CommandButton cmndbtnGraf 
      Caption         =   "&Gráfico"
      Enabled         =   0   'False
      Height          =   735
      Left            =   5640
      Picture         =   "frmRelAtvCli.frx":04BB
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   735
      Left            =   8520
      Picture         =   "frmRelAtvCli.frx":08FD
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "Im&primir"
      Height          =   735
      Left            =   7080
      Picture         =   "frmRelAtvCli.frx":11C7
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Caption         =   "&Buscar"
      Height          =   735
      Left            =   4200
      Picture         =   "frmRelAtvCli.frx":14D1
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   120
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   255
      End
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   375
      Left            =   1200
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
      Bindings        =   "frmRelAtvCli.frx":17DB
      Height          =   4935
      Left            =   120
      TabIndex        =   25
      Top             =   960
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   8705
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
Attribute VB_Name = "frmRelAtvCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxOrder_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub


Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    strSQL = "select count(tf.num_ven) as qtd,tc.nome_cli,tc.cod_cli,"
    strSQL = strSQL & "sum(tf.valorfinal_ven) as final,sum(tf.avista_ven) as avista,   "
    strSQL = strSQL & " sum(tf.aprazo_ven) as aprazo, sum(tf.finan_ven) as finan,sum(tf.desconto_ven) as desconto "
    strSQL = strSQL & " from " & TAB_VENDAS & " as tv,  " & TAB_VENFEC & "  as tf, " & TAB_CLIENTES & " as tc "
    strSQL = strSQL & " where tc.cod_cli=tv.cod_cli and tv.num_ven=tf.num_ven "
    If txtbxDAtaFIn.Text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtbxDAtaFIn.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
        
    End If
    strSQL = strSQL & " and situacao_ven='4' "
    strSQL = strSQL & " group by tc.cod_cli "
    cmndbtnGraf.Enabled = True
    If cmbxOrder.ListIndex = 0 Then
        strSQL = strSQL & " order by nome_cli "
        cmndbtnGraf.Enabled = False
    ElseIf cmbxOrder.ListIndex = 1 Then
        strSQL = strSQL & " order by final desc "
    ElseIf cmbxOrder.ListIndex = 2 Then
        strSQL = strSQL & " order by avista desc "
    ElseIf cmbxOrder.ListIndex = 3 Then
        strSQL = strSQL & " order by aprazo desc "
    ElseIf cmbxOrder.ListIndex = 4 Then
        strSQL = strSQL & " order by finan desc "
    ElseIf cmbxOrder.ListIndex = 5 Then
        strSQL = strSQL & " order by desconto desc "
    ElseIf cmbxOrder.ListIndex = 6 Then
        strSQL = strSQL & " order by qtd desc "
    End If
    dt.RecordSource = strSQL
    dt.Refresh
    soma
End Sub
Sub soma()
    Dim cFinal As Currency
    Dim cAv As Currency
    Dim cAp As Currency
    Dim cFinan As Currency
    Dim cDesc As Currency
    Dim cQtd As Currency
    On Error GoTo trata_Erro
    With dt.Recordset
        Do While Not .EOF
            cFinal = cFinal + !Final
            cAv = cAv + !avista
            cAp = cAp + !aprazo
            cFinan = cFinan + !finan
            cDesc = cDesc + !Desconto
            cQtd = cQtd + !qtd
            
            .MoveNext
        Loop
        lblAp.caption = FormatVal(cAp)
        lblAv.caption = FormatVal(cAv)
        lblDesc.caption = FormatVal(cDesc)
        lblFinal.caption = FormatVal(cFinal)
        lblFinan.caption = FormatVal(cFinan)
        lblQtd.caption = cQtd
        If .RecordCount > 0 Then .MoveFirst
    End With
    Exit Sub
trata_Erro:
MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnGraf_Click()
    If cmbxOrder.ListIndex = 1 Then
        grafico dt.Recordset!Final
    
    ElseIf cmbxOrder.ListIndex = 2 Then
        grafico dt.Recordset!avista
    ElseIf cmbxOrder.ListIndex = 3 Then
        grafico dt.Recordset!aprazo
    ElseIf cmbxOrder.ListIndex = 4 Then
        grafico dt.Recordset!finan
    ElseIf cmbxOrder.ListIndex = 5 Then
        grafico dt.Recordset!Desconto
    ElseIf cmbxOrder.ListIndex = 6 Then
        grafico dt.Recordset!qtd
    End If
End Sub
Sub grafico(f As Field)
    Dim b(11) As valChart
    Dim iCount As Integer
    Dim cRest As Currency
    
        With dt.Recordset
            .MoveFirst
            For iCount = 0 To 9
                
                b(iCount).caption = !nome_Cli
                b(iCount).Value = f
                 
                .MoveNext
            Next
            Do While Not .EOF
                cRest = cRest + f
                
                .MoveNext
            Loop
            b(10).caption = "RESTANTE"
            b(10).Value = cRest
        End With
        
    
    P_Graph b, 11, 1, True
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
    P_CabRelTimbr Me.caption & " de " & txtbxDataIni.Text & " até  " & txtbxDAtaFIn.Text, obj
    P_ImprX obj, "Cód", 0
    P_ImprX obj, "Cliente", 15
    P_ImprX obj, "Final", 115, True, True
    P_ImprX obj, "À Vista", 135, True, True
    P_ImprX obj, "A Prazo", 155, True, True
    P_ImprX obj, "Financeira", 175, True, True
    P_ImprX obj, "Desconto", 195, True, True
    P_ImprX obj, "Qtd", 196, False
    P_ImprLinha obj
    With dt.Recordset
        .MoveFirst
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_RodapeVis obj, "", "Continua..."
                P_ImprX obj, "Cód", 0
                P_ImprX obj, "Cliente", 15
                P_ImprX obj, "Final", 115, True, True
                P_ImprX obj, "À Vista", 135, True, True
                P_ImprX obj, "A Prazo", 155, True, True
                P_ImprX obj, "Financeira", 175, True, True
                P_ImprX obj, "Desconto", 195, True, True
                P_ImprX obj, "Qtd", 196, False
                P_ImprLinha obj
            End If
            P_ImprX obj, !cod_cli, 0
            P_ImprX obj, !nome_Cli, 15
            P_ImprX obj, FormatVal(!Final), 115, True, True
            P_ImprX obj, FormatVal(!avista), 135, True, True
            P_ImprX obj, FormatVal(!aprazo), 155, True, True
            P_ImprX obj, FormatVal(!finan), 175, True, True
            P_ImprX obj, FormatVal(!Desconto), 195, True, True
            P_ImprX obj, !qtd, 198, False
            
            .MoveNext
        Loop
        P_ImprLinha obj
        P_ImprX obj, "Totais...", 0
        P_ImprX obj, lblFinal.caption, 115, True, True
        P_ImprX obj, lblAv.caption, 135, True, True
        P_ImprX obj, lblAp.caption, 155, True, True
        P_ImprX obj, lblFinan.caption, 175, True, True
        P_ImprX obj, lblDesc.caption, 195, True, True
        P_ImprX obj, lblQtd.caption, 198, False
        
        P_ImprLinha obj, "_"
        P_RodapeVis obj
    End With
    
    
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    'MudaAparencia Me
End Sub

Private Sub Form_Load()
    dt.ConnectionString = strConnect
    P_CentraForm Me
    cmbxOrder.ListIndex = 0
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    ganhaFoco txtbxDAtaFIn
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxDAtaFIn
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
