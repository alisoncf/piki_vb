VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRecebimentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relação de Recebimentos"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   Icon            =   "frmRecebimentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   10215
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FEF4E4&
      ForeColor       =   &H00FEF4E4&
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2265
      ScaleWidth      =   2145
      TabIndex        =   22
      Top             =   4200
      Width           =   2175
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   2175
         TabIndex        =   29
         Top             =   1440
         Width           =   2175
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
            ForeColor       =   &H00FEF8ED&
            Height          =   195
            Left            =   855
            TabIndex        =   31
            Tag             =   "!"
            Top             =   480
            Width           =   390
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Recebido:"
            ForeColor       =   &H00FEF8ED&
            Height          =   195
            Left            =   480
            TabIndex        =   30
            Tag             =   "!"
            Top             =   120
            Width           =   1140
         End
      End
      Begin VB.Label lblJuros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Tag             =   "!"
         Top             =   1080
         Width           =   315
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   1440
         TabIndex        =   27
         Tag             =   "!"
         Top             =   1080
         Width           =   315
      End
      Begin VB.Label lblValDoc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   915
         TabIndex        =   26
         Tag             =   "!"
         Top             =   450
         Width           =   315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total das Parcelas:"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   390
         TabIndex        =   25
         Tag             =   "!"
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Juros:"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Tag             =   "!"
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descontos:"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   1185
         TabIndex        =   23
         Tag             =   "!"
         Top             =   840
         Width           =   810
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Buscar Por:"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   650
      Width           =   2175
      Begin VB.ComboBox cmbxBuscar 
         Height          =   315
         ItemData        =   "frmRecebimentos.frx":2CFA
         Left            =   120
         List            =   "frmRecebimentos.frx":2D07
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "E&xpressão:"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   1280
      Width           =   2175
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "&Ordenar Por:"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   1940
      Width           =   2175
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmRecebimentos.frx":2D27
         Left            =   120
         List            =   "frmRecebimentos.frx":2D3A
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   2175
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   2
         Tag             =   "!"
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Tag             =   "!"
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   120
      Picture         =   "frmRecebimentos.frx":2D75
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "Im&primir"
      Height          =   615
      Left            =   1200
      Picture         =   "frmRecebimentos.frx":363F
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Caption         =   "&Buscar"
      Height          =   615
      Left            =   120
      Picture         =   "frmRecebimentos.frx":3949
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   6615
      Left            =   2400
      TabIndex        =   14
      Top             =   0
      Width           =   7815
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         ForeColor       =   &H80000008&
         Height          =   350
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   7545
         TabIndex        =   19
         Top             =   4935
         Width           =   7575
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000001&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   7035
            Picture         =   "frmRecebimentos.frx":3C53
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   20
            Top             =   0
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Formas de Baixa:"
            ForeColor       =   &H00FEF4E4&
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Tag             =   "!"
            Top             =   75
            Width           =   1215
         End
      End
      Begin MSFlexGridLib.MSFlexGrid msflgrd 
         Height          =   1215
         Left            =   120
         TabIndex        =   13
         Top             =   5280
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   2143
         _Version        =   393216
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         ForeColor       =   &H80000008&
         Height          =   350
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   7545
         TabIndex        =   16
         Top             =   250
         Width           =   7575
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000001&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   7035
            Picture         =   "frmRecebimentos.frx":594D
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   17
            Top             =   0
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Documentos Baixados no período:"
            ForeColor       =   &H00FEF4E4&
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Tag             =   "!"
            Top             =   75
            Width           =   2460
         End
      End
      Begin MSAdodcLib.Adodc adc 
         Height          =   375
         Left            =   1320
         Top             =   6480
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
         Bindings        =   "frmRecebimentos.frx":9DB7
         Height          =   4095
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   7223
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
End
Attribute VB_Name = "frmRecebimentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean

Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim sOrderBy As String
    bActive = False
    If cmbxOrder.ListIndex = 0 Then
        sOrderBy = "ti.venc_itd"
    ElseIf cmbxOrder.ListIndex = 1 Then
        sOrderBy = "tc.nome_cli"
    ElseIf cmbxOrder.ListIndex = 2 Then
        sOrderBy = "td.num_doc"
    ElseIf cmbxOrder.ListIndex = 3 Then
        sOrderBy = "td.valor_doc"
    ElseIf cmbxOrder.ListIndex = 4 Then
        sOrderBy = "ti.baixa_itd"
    End If
    strSQL = "select ti.*,tc.nome_cli,tc.cod_cli,td.num_doc as doc,td.valor_doc, "
    strSQL = strSQL & "(ti.Valor_Itd + ti.juros_itd) - ti.reman_itd - ti.desconto_itd as baixado "
    strSQL = strSQL & "from  " & TAB_ITDOC & " 1 as ti," & TAB_CLIENTES & " as tc, " & TAB_DOC & " 1 as td "
    strSQL = strSQL & "where tc.cod_cli = td.cod_ter and td.seq_doc = ti.seq_doc "
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & "and ti.baixa_itd >=" & F_GravaData(txtbxDataIni.Text) & " "
    End If
    If txtbxDAtaFIn.Text <> "" Then
        strSQL = strSQL & "and ti.baixa_itd <=" & F_GravaData(txtbxDAtaFIn.Text) & " "
    End If
    If cmbxBuscar.ListIndex = 0 Then
        strSQL = strSQL & "and td.num_doc like '%" & txtbxExpr.Text & "%' "
    ElseIf cmbxBuscar.ListIndex = 1 Then
        strSQL = strSQL & "and tc.nome_cli like '%" & txtbxExpr.Text & "%' "
    ElseIf cmbxBuscar.ListIndex = 2 Then
        strSQL = strSQL & "and td.valor_doc like '" & txtbxExpr.Text & "%' "
    End If
    strSQL = strSQL & "and ti.reman_itd <> ti.valor_itd and ti.baixa_itd <> '' "
    strSQL = strSQL & "order by " & sOrderBy
    dtNot.RecordSource = strSQL
    dtNot.Refresh
    soma
    msflgrd.rows = 1
    bActive = True
    If dtNot.Recordset.RecordCount > 0 Then dtNot.Recordset.MoveFirst
End Sub

Private Sub cmndbtnImprimir_Click()
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
    End If
    Unload frmVisualiza
End Sub

Sub imprime(obj As Object)
    Dim iItens As Integer
    Dim iPag As Integer
    Dim strSQL As String
    Dim sTipo As String
    Dim sCodTer As String
    On Error Resume Next
    If cmbxBuscar.ListIndex = 1 Then
        sTipo = "Relação de Documentos Baixados por Cliente "
    Else
        sTipo = "Relação de Documentos Baixados"
    End If
    If txtbxDataIni.Text <> "" Or txtbxDAtaFIn.Text <> "" Then
        P_CabRelTimbr sTipo & " de " & txtbxDataIni.Text & _
            " Até " & txtbxDAtaFIn.Text, obj
    Else
        P_CabRelTimbr sTipo, obj
    End If
    
    obj.FontName = "arial"
    obj.fontsize = 12
    P_ImprLinha obj
    obj.fontsize = 9
    P_ImprX obj, "Doc.", 0
    P_ImprX obj, "Terceiro", 15
    P_ImprX obj, "Baixado", 100, True, True
    P_ImprX obj, "Acrésc.", 115, True, True
    P_ImprX obj, "Desc.", 130, True, True
    P_ImprX obj, "Reman.", 135
    P_ImprX obj, "Valor Doc.", 165, True, True
    P_ImprX obj, "Data", 173, False
    P_ImprLinha obj
  
    With dtNot.Recordset
        .MoveFirst
        strSQL = CStr(.RecordCount / 50)
        If InStr(1, strSQL, ",") <> 0 Then
            iPag = Int(strSQL) + 1
        End If
            iItens = iItens + 1
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                If TypeName(obj) = "Printer" Then
                    P_Rodape "Continua...", True, "Pág:" & IIf(TypeName(obj) = "Printer", obj.Page, "%")
                Else
                    P_RodapeVis obj, "Continua...", "Pág:" & IIf(TypeName(obj) = "Printer", "N", "%")
                End If
                iItens = 0
                obj.FontName = "arial"
                obj.fontsize = 9
                P_ImprX obj, "Doc.", 0
                P_ImprX obj, "Terceiro", 15
                P_ImprX obj, "Baixado", 100, True, True
                P_ImprX obj, "Acrésc.", 115, True, True
                P_ImprX obj, "Desc.", 130
                P_ImprX obj, "Reman.", 135
                P_ImprX obj, "Valor Doc.", 165, True, True
                P_ImprX obj, "Data", 173, False
                P_ImprLinha obj
            End If
            obj.FontName = "arial"
            obj.fontsize = 9
        Do While Not .EOF
            P_ImprX obj, !doc, 0
            P_ImprX obj, !nome_Cli, 15
            P_ImprX obj, FormatVal(!valor_itd), 100, True, True
            P_ImprX obj, FormatVal(!juros_itd), 115, True, True
            P_ImprX obj, FormatVal(!desconto_itd), 130, True, True
            P_ImprX obj, FormatVal(!reman_itd), 145, True, True
            P_ImprX obj, FormatVal(!valor_doc), 165, True, True
            P_ImprX obj, !baixa_itd, 190, False, True
            .MoveNext
        Loop
        P_ImprLinha obj
        obj.FontBold = True
        P_ImprX obj, "Totais............: ", 15
        P_ImprX obj, FormatVal(lblTOtal.caption), 100, True, True
        P_ImprX obj, FormatVal(lblJuros.caption), 115, True, True
        P_ImprX obj, FormatVal(lblDesc.caption), 130, False, True
        
        If TypeName(obj) = "Printer" Then
            P_RodapeVis obj, strNome_Emp, "Pág: " & obj.Page & " de " & obj.Page
        End If
    End With
    If TypeName(obj) = "Printer" Then
        obj.EndDoc
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub dtNot_Reposition()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    'On Error Resume Next
    With dtNot.Recordset
        If .EOF Then Exit Sub
        'Busca baixa em dinheiro
        msflgrd.rows = 1
        strSQL = "select tf.*,tc.nome_con from  " & TAB_FINANCEIRO & "  as tf, tab_contas as tc where seq_lan = '" & _
            !seq_itd & "' and tf.cod_con = tc.cod_con "
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and tf.data_fin >=" & F_GravaData(txtbxDataIni.Text)
        End If
        If txtbxDAtaFIn.Text <> "" Then
            strSQL = strSQL & " and tf.data_fin <=" & F_GravaData(txtbxDAtaFIn.Text)
        End If
        rPesq.Open strSQL, conn_client
                With rPesq
                    Do While Not .EOF
                        strSQL = "Dinheiro" & Chr(9) & rPesq!seq_fin & Chr(9) & FormatVal(rPesq!valor_fin) & _
                            Chr(9) & rPesq!historico_fin & Chr(9) & rPesq!nome_con & Chr(9) & rPesq!data_fin
                        msflgrd.AddItem strSQL
                        .MoveNext
                    Loop
                    .Close
                End With
        'Busca Baixa em cheques
        strSQL = "select td.seq_doc, td.data_doc,tc.*,tcon.nome_con from  " & TAB_DOC & " 1 as td, tab_cheques1 as tc,tab_contas as tcon "
        strSQL = strSQL & "where num_doc ='" & !doc & "' and td.data_doc =" & F_GravaData(!baixa_itd) & " and "
        strSQL = strSQL & "td.seq_doc = tc.seq_doc and tc.cod_con = tcon.cod_con "
        rPesq.Open strSQL, conn_client
            With rPesq
                Do While Not .EOF
                    strSQL = "Cheque" & Chr(9) & rPesq!num_che & Chr(9) & FormatVal(rPesq!Valor_che) & _
                        Chr(9) & rPesq!nome_che & Chr(9) & rPesq!nome_con & Chr(9) & rPesq!Venc_che
                    msflgrd.AddItem strSQL
                    .MoveNext
                Loop
                .Close
            End With
        'Busca Baixa em Cartão
        strSQL = "select tf.nome_fin,tv.* from tab_financeiras as tf,tab_venfin as tv where num_ven like 'B%' and data_ven=" & _
            F_GravaData(!baixa_itd) & " and valor_ven=" & F_GravaDec(!baixado) & " and tv.cod_fin = tf.cod_fin "
        rPesq.Open strSQL, conn_client
                With rPesq
                    Do While Not .EOF
                        strSQL = "Financeira" & Chr(9) & rPesq!num_ven & Chr(9) & FormatVal(rPesq!valor_ven) & _
                            Chr(9) & rPesq!nome_fin & Chr(9) & "" & Chr(9) & rPesq!data_ven
                        msflgrd.AddItem strSQL
                        .MoveNext
                    Loop
                    .Close
                End With
    End With
 
End Sub

Private Sub Form_Activate()
'    'MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        dbgrd.SetFocus
    End If
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    Me.Height = 7035
    Me.Width = 10335
    Me.Top = 140
    Me.Left = 100
    adc.ConnectionString = strConnect
    cmbxOrder.ListIndex = 0
    cmbxBuscar.ListIndex = 0
    With msflgrd
        .FixedCols = 0
        .Cols = 6
        .rows = 1
        .col = 0
        .ColWidth(0) = 1200
        .Text = "Baixa com:"
        .col = 1
        .ColWidth(1) = 900
        .Text = "Ident."
        .col = 2
        .ColWidth(2) = 800
        .Text = "Valor"
        .col = 3
        .ColWidth(3) = 3100
        .Text = "Descrição"
        .col = 4
        .ColWidth(4) = 1700
        .Text = "Conta"
        .col = 5
        .ColWidth(5) = 1000
        .Text = "Data"
    End With
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

Private Sub txtbxExpr_GotFocus()
    ganhaFoco txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    AjustaTexto txtbxExpr
End Sub

Sub soma()
    Dim cTotal As Currency, cJuros As Currency, cDesc As Currency, cValDoc As Currency
    If dtNot.Recordset.RecordCount = 0 Then Exit Sub
    With dtNot.Recordset
        Do While Not .EOF
            cTotal = cTotal + !baixado
            cJuros = cJuros + !juros_itd
            cDesc = cDesc + !desconto_itd
            cValDoc = cValDoc + !valor_itd
            .MoveNext
        Loop
    End With
    lblTOtal.caption = FormatVal(cTotal)
    lblJuros.caption = FormatVal(cJuros)
    lblDesc.caption = FormatVal(cDesc)
    lblValDoc.caption = FormatVal(cValDoc)
End Sub

