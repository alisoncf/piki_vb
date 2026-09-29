VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelAtvCliProd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Atividade de Clientes por Produto"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   Icon            =   "frmRelAtvCliProd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9855
   Begin VB.Frame Frame3 
      Caption         =   "Opções de Busca:"
      Height          =   735
      Left            =   4440
      TabIndex        =   23
      Top             =   840
      Width           =   5295
      Begin VB.OptionButton optnAna 
         Caption         =   "&Analítico"
         Height          =   255
         Left            =   1680
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optnSint 
         Caption         =   "&Sintético"
         Height          =   255
         Left            =   3120
         TabIndex        =   24
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   4
         Top             =   360
         Width           =   975
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1080
         TabIndex        =   3
         Top             =   120
         Width           =   285
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   2280
      TabIndex        =   5
      Top             =   0
      Width           =   1935
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmRelAtvCliProd.frx":014A
         Left            =   120
         List            =   "frmRelAtvCliProd.frx":015A
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar por:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   885
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   9705
      TabIndex        =   16
      Top             =   6120
      Width           =   9735
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblQtd 
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
         TabIndex        =   21
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   1560
         TabIndex        =   20
         Top             =   120
         Width           =   360
      End
      Begin VB.Label lblTotal 
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
         TabIndex        =   19
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Final"
         Height          =   195
         Left            =   3000
         TabIndex        =   18
         Top             =   120
         Width           =   330
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
         Left            =   3720
         TabIndex        =   17
         Top             =   120
         Width           =   390
      End
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   735
      Left            =   8400
      Picture         =   "frmRelAtvCliProd.frx":0190
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   100
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "Im&primir"
      Height          =   735
      Left            =   7080
      Picture         =   "frmRelAtvCliProd.frx":0A5A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   100
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnGraf 
      Caption         =   "&Gráfico"
      Enabled         =   0   'False
      Height          =   735
      Left            =   5760
      Picture         =   "frmRelAtvCliProd.frx":0D64
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   100
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Caption         =   "&Buscar"
      Height          =   735
      Left            =   4440
      Picture         =   "frmRelAtvCliProd.frx":11A6
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   100
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   4215
      Begin VB.TextBox txtbxNomeCli 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   3735
      End
      Begin VB.CommandButton cmndbtnCliente 
         Height          =   285
         Left            =   3840
         Picture         =   "frmRelAtvCliProd.frx":14B0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   525
      End
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   375
      Left            =   1080
      Top             =   6600
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
      Bindings        =   "frmRelAtvCliProd.frx":15FA
      Height          =   4215
      Left            =   0
      TabIndex        =   26
      Top             =   1800
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7435
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
Attribute VB_Name = "frmRelAtvCliProd"
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
    If optnAna.Value Then
        ana
    Else
        sint
    End If
    soma

End Sub
Sub ana()
    Dim strSQL As String
    
    strSQL = "select ti.qtd_ite,tp.nome_pro,tp.cod_pro,tq.pv_pro,"
    strSQL = strSQL & "ti.final_ite,(ti.qtd_ite*ti.unit_ite) as total,tv.*   "
    
    strSQL = strSQL & " from " & TAB_VENDAS & " as tv,  " & TAB_VENFEC & "  as tf, " & TAB_CLIENTES & " as tc,"
    strSQL = strSQL & " " & TAB_ITVENDAS & " as ti, " & TAB_PRODUTOS & " as tp, tab_composicao as tq   "
    strSQL = strSQL & " where tc.cod_cli=tv.cod_cli and tv.num_ven=tf.num_ven "
    strSQL = strSQL & " and tq.cod_pro=tp.cod_pro"
    strSQL = strSQL & " and ti.num_ven=tv.num_ven "
    strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtbxdataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
        
    End If
    strSQL = strSQL & " and situacao_ven='4' "
    strSQL = strSQL & " and tc.cod_cli='" & txtbxNomeCli.tag & "' "
    
    cmndbtnGraf.Enabled = True
    If cmbxOrder.ListIndex = 0 Then
        strSQL = strSQL & " order by nome_pro "
        cmndbtnGraf.Enabled = False
    ElseIf cmbxOrder.ListIndex = 1 Then
        strSQL = strSQL & " order by total desc "
    ElseIf cmbxOrder.ListIndex = 2 Then
        strSQL = strSQL & " order by final_ite desc "
    ElseIf cmbxOrder.ListIndex = 3 Then
        strSQL = strSQL & " order by qtd_ite desc "
    End If
    dt.RecordSource = strSQL
    dt.Refresh
End Sub
Sub sint()
    Dim strSQL As String
    
    strSQL = "select sum(ti.qtd_ite) as qtd,tp.*,"
    strSQL = strSQL & "sum(ti.final_ite) as final,sum(ti.qtd_ite*ti.unit_ite) as total   "
    
    strSQL = strSQL & " from " & TAB_VENDAS & " as tv,  " & TAB_VENFEC & "  as tf, " & TAB_CLIENTES & " as tc,"
    strSQL = strSQL & " " & TAB_ITVENDAS & " as ti, " & TAB_PRODUTOS & " as tp "
    strSQL = strSQL & " where tc.cod_cli=tv.cod_cli and tv.num_ven=tf.num_ven "
    
    strSQL = strSQL & " and ti.num_ven=tv.num_ven "
    strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtbxdataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
        
    End If
    strSQL = strSQL & " and situacao_ven='4' "
    strSQL = strSQL & " and tc.cod_cli='" & txtbxNomeCli.tag & "' "
        
    strSQL = strSQL & " group by tp.cod_pro  "
    
    
    cmndbtnGraf.Enabled = True
    If cmbxOrder.ListIndex = 0 Then
        strSQL = strSQL & " order by nome_pro "
        cmndbtnGraf.Enabled = False
    ElseIf cmbxOrder.ListIndex = 1 Then
        strSQL = strSQL & " order by total desc "
    ElseIf cmbxOrder.ListIndex = 2 Then
        strSQL = strSQL & " order by final desc "
    ElseIf cmbxOrder.ListIndex = 3 Then
        strSQL = strSQL & " order by qtd desc "
    End If
    dt.RecordSource = strSQL
    dt.Refresh
End Sub
Sub soma()
    Dim cFinal As Currency
    Dim cTotal As Currency
    Dim cQtd As Currency
    On Error GoTo trata_erro
    With dt.Recordset
        Do While Not .EOF
            If optnSint.Value Then
                cFinal = cFinal + !Final
                cTotal = cTotal + !total
                cQtd = cQtd + !qtd
            Else
                cFinal = cFinal + !final_ite
                cTotal = cTotal + !total
                cQtd = cQtd + !qtd_ite
            End If
            .MoveNext
        Loop
        
        lblTotal.caption = FormatVal(cTotal)
        lblQtd.caption = cQtd
        lblFinal.caption = FormatVal(cFinal)
        
        If .RecordCount > 0 Then .MoveFirst
    End With
    Exit Sub
trata_erro:
MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnCliente_Click()
    
        frmSelectCli.txtbxNOme.tag = "C"
        frmSelectCli.txtbxNOme.Text = txtbxNomeCli.Text
    
    On Error Resume Next
    
    frmSelectCli.Show 1
    
'    If F_BuscaUmCampo("status_cli", "tab_clientes", "cod_cli='" & _
'        frmSelectCli.dtBusca.Recordset!cod & "'") <> "1" Then
'        MsgBox "Não é possível abrir venda para este cliente pois a situação esté negativa!", vbCritical
'        Exit Sub
'        txtbxNomeCli.Tag = ""
'        txtbxNomeCli.Text = ""
'    End If
    
    txtbxNomeCli.Text = frmSelectCli.adoR!nome
    txtbxNomeCli.tag = frmSelectCli.adoR!cod
    
    
    
    
    Unload frmSelectCli
End Sub

Private Sub cmndbtnGraf_Click()
    On Error GoTo trata_erro
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
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub grafico(f As Field)
    Dim b As BlChart
    Dim iCount As Integer
    Dim cRest As Currency
    Dim iRows As Integer
    
        With dt.Recordset
            .MoveFirst
            iRows = .RecordCount
            If iRows > 9 Then
                iRows = 9
                b.count = 10 + 1
            Else
                b.count = iRows
            End If
            
            For iCount = 0 To iRows - 1
                
                b.val(iCount).caption = !nome_pro
                b.val(iCount).Value = f
                 
                .MoveNext
            Next
            Do While Not .EOF
                cRest = cRest + f
                
                .MoveNext
            Loop
            If cRest > 0 Then
                b.val(iRows + 1).caption = "RESTANTE"
                b.val(iRows + 1).Value = cRest
            End If
        End With
        
    
    'P_Graph b, 1, True
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
    P_CabRelTimbr Me.caption & " de " & txtbxDataIni.Text & " até  " & txtbxdataFin.Text, obj
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
            P_ImprX obj, txtbxNomeCli.tag, 0
            P_ImprX obj, txtbxNomeCli.Text, 15
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

Private Sub optnAna_Click()
    With dbgrd
        .Columns(5).Visible = True
        .Columns(6).Visible = True
        .Columns(0).DataField = "Cod_pro"
        .Columns(1).DataField = "nome_pro"
        .Columns(2).DataField = "qtd_ite"
        .Columns(3).DataField = "total"
        .Columns(4).DataField = "final_ite"
        
        .Columns(5).DataField = "dataefet_ven"
        .Columns(6).DataField = "num_ven"
        .Columns(5).caption = "Data"
        .Columns(6).caption = "Venda"
    End With
End Sub

Private Sub optnSint_Click()
    With dbgrd
        .Columns(5).Visible = False
        .Columns(6).Visible = False
        .Columns(0).DataField = "Cod_pro"
        .Columns(1).DataField = "nome_pro"
        .Columns(2).DataField = "qtd"
        .Columns(3).DataField = "total"
        .Columns(4).DataField = "final"
    End With
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

