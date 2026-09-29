VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRelProdLucr 
   Caption         =   "Relatório de Produtos por Lucratividade"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   Icon            =   "frmRelProdLucr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   10410
   Begin VB.ComboBox cmbxOrderBY 
      Height          =   315
      ItemData        =   "frmRelProdLucr.frx":0442
      Left            =   5640
      List            =   "frmRelProdLucr.frx":0455
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "Im&primir"
      Height          =   615
      Left            =   9000
      Picture         =   "frmRelProdLucr.frx":0494
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   9000
      Picture         =   "frmRelProdLucr.frx":05DE
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Caption         =   "&Buscar"
      Height          =   495
      Left            =   5640
      Picture         =   "frmRelProdLucr.frx":0728
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   960
      Width           =   1695
   End
   Begin TabDlg.SSTab sstb 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2778
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Busca"
      TabPicture(0)   =   "frmRelProdLucr.frx":0872
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Composição"
      TabPicture(1)   =   "frmRelProdLucr.frx":088E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmbxMarca"
      Tab(1).Control(1)=   "cmbxCat"
      Tab(1).Control(2)=   "txtbxDesc"
      Tab(1).Control(3)=   "Label1"
      Tab(1).Control(4)=   "Label2"
      Tab(1).Control(5)=   "Label4"
      Tab(1).ControlCount=   6
      Begin VB.Frame Frame2 
         Caption         =   "Período Útil:"
         Height          =   975
         Left            =   2880
         TabIndex        =   5
         Top             =   360
         Width           =   2295
         Begin VB.TextBox txtbxdataFin 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   9
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtbxDataIni 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   7
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "&Até:"
            Height          =   195
            Left            =   1200
            TabIndex        =   8
            Top             =   240
            Width           =   285
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "&De:"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.ComboBox cmbxMarca 
         Height          =   315
         ItemData        =   "frmRelProdLucr.frx":08AA
         Left            =   -73080
         List            =   "frmRelProdLucr.frx":08B7
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   480
         Width           =   3015
      End
      Begin VB.ComboBox cmbxCat 
         Height          =   315
         ItemData        =   "frmRelProdLucr.frx":08DA
         Left            =   -73080
         List            =   "frmRelProdLucr.frx":08E7
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtbxDesc 
         Height          =   285
         Left            =   -73080
         TabIndex        =   20
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Frame Frame1 
         Caption         =   "Buscar:"
         Height          =   975
         Left            =   30
         TabIndex        =   1
         Top             =   360
         Width           =   2775
         Begin VB.OptionButton optnAbaixoCrit 
            Caption         =   "C/ Margem Inferior ao Critério"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   2415
         End
         Begin VB.OptionButton optnTodos 
            Caption         =   "&Todos os Produtos"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optnInfDesp 
            Caption         =   "C/ Margem Inferior as Depesas"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   2535
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Marca:"
         Height          =   195
         Left            =   -73680
         TabIndex        =   15
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Categ:"
         Height          =   195
         Left            =   -73680
         TabIndex        =   17
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Descrição:"
         Height          =   195
         Left            =   -74040
         TabIndex        =   19
         Top             =   1200
         Width           =   765
      End
   End
   Begin VB.PictureBox pc 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   10305
      TabIndex        =   22
      Top             =   6720
      Width           =   10335
      Begin VB.Label lblList 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listados:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9600
         TabIndex        =   24
         Top             =   0
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "***LS=lucro Sujerido; LE=Lucro Empregado;PS=Preço Sujerido;PE=Preço Empregado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   6285
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   5055
      Left            =   0
      TabIndex        =   21
      Top             =   1680
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8916
      _Version        =   393216
      GridColor       =   16777215
      GridColorFixed  =   16777215
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Ordenar:"
      Height          =   195
      Left            =   5640
      TabIndex        =   10
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frmRelProdLucr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ajustaGrid()
    With Msflxgrd
        .Cols = 19
        .FixedCols = 0
        .rows = 1
        .col = 0: .Text = "Código"
        .col = 1: .Text = "Descrição"
        .col = 2: .Text = "Un."
        .col = 3: .Text = "Custo"
        .col = 4: .Text = "L.S(%)"
        .col = 5: .Text = "L.E(%)"
        .col = 6: .Text = "Dif(%)"
        .col = 7: .Text = "L.L(%)"
        .col = 8: .Text = "P.S(R$)"
        .col = 9: .Text = "P.E(R$)"
        .col = 10: .Text = "Dif.(R$)"
        .col = 11: .Text = "Q.G."
        .col = 12: .Text = "V.G."
        .col = 13: .Text = "G.B."
        .col = 14: .Text = "G.L."
        .col = 15: .Text = "Est"
        .col = 16: .Text = "Disp"
        .col = 17: .Text = "Categoria"
        .col = 18: .Text = "Marca"
        
        .ColWidth(0) = 1100: .ColAlignment(0) = 0
        .ColWidth(1) = 3500: .ColAlignment(1) = 0
        .ColWidth(2) = 600: .ColAlignment(2) = 0
        .ColWidth(3) = 800:
        .ColWidth(4) = 800:
        .ColWidth(5) = 800:
        .ColWidth(6) = 800:
        
        .ColWidth(7) = 800:
        .ColWidth(8) = 800:
        .ColWidth(9) = 800:
        .ColWidth(10) = 800:
        .ColWidth(11) = 800:
        .ColWidth(12) = 800:
        .ColWidth(17) = 2000: .ColAlignment(13) = 0
        .ColWidth(18) = 2000: .ColAlignment(14) = 0
    End With
End Sub

Private Sub cmbxCaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxCaixa.ListIndex = -1
    End If
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    If KeyCode = vbKeyF5 Then
        P_PreencheCateg cmbxCat
    End If
End Sub

Private Sub cmbxMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    If KeyCode = vbKeyF5 Then
        P_PreencheMarca cmbxMarca
    End If
End Sub

Private Sub cmbxOrderBY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxTurno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxTurno.ListIndex = -1
    End If
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim sTime As String
    sTime = Now
    Msflxgrd.Visible = False
    P_Status "Montando Busca..."
    busca
    P_Status ""
    Msflxgrd.Visible = True
    MsgBox "Pesquisa concluída em " & DateDiff("s", sTime, Now) & " segundos!", vbInformation
End Sub
Sub busca1()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rVendas As New ADODB.Recordset
    Dim cTotal As Currency
    Dim sMsg As String
    Msflxgrd.rows = 1
    
        strSQL = "select tp.*,(tp.cr_pro+(tp.cr_pro*tp.lucro_pro/100)) as ps,"
        strSQL = strSQL & "IF (tp.lucro_pro>0 ,(tp.lucro_pro*tp.pv_pro/(tp.cr_pro+(tp.cr_pro*tp.lucro_pro/100))),(tp.cr_pro/tp.pv_pro*100)) as le,"
        strSQL = strSQL & " tm.nome_mar,tc.nome_cat "
        
        strSQL = strSQL & " from " & TAB_PRODUTOS & " as tp,"
        strSQL = strSQL & " tab_Categorias as tc," & TAB_MARCAS & " as tm, "
        strSQL = strSQL & " " & TAB_VENDAS & " as tv, " & TAB_ITVENDAS & " as ti "
        strSQL = strSQL & " where tp.cod_Cat=tc.cod_Cat"
        strSQL = strSQL & " and tp.cod_mar=tm.cod_mar "
        strSQL = strSQL & " and tv.num_ven=ti.num_ven "
        strSQL = strSQL & " and ti.cod_pro=tp.cod_pro "
        strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
        strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtbxdataFin.Text)
        If cmbxCat.Text <> "" Then
            strSQL = strSQL & " and  tc.nome_cat='" & cmbxCat.Text & "' "
        End If
        If cmbxMarca.Text <> "" Then
            strSQL = strSQL & " and tm.nome_mar='" & cmbxMarca.Text & "' "
        End If
        
        If txtbxDesc.Text <> "" Then
            strSQL = strSQL & " and tp.nome_pro like '" & txtbxDesc.Text & "%'"
        End If
        If optnAbaixoCrit.Value Then
            strSQL = strSQL & " having tp.lucro_pro > le "
        End If
        If optnInfDesp.Value Then
            'strSQL = strSQL & " having (tp.lucro_pro -  " & F_GravaDec(F_TaxasEntradas) & ")<0 "
            strSQL = strSQL & " having (le -  " & F_GravaDec(F_TaxasEntradas) & ")<0 "
        End If
        strSQL = strSQL & " group by ti.cod_pro "
        If cmbxOrderBY.ListIndex = 1 Then
            strSQL = strSQL & " order by ti.cod_pro"
        ElseIf cmbxOrderBY.ListIndex = 2 Then
            strSQL = strSQL & " order by nome_pro"
        ElseIf cmbxOrderBY.ListIndex = 3 Then
        '    strSQL = strSQL & " order by total desc"
        ElseIf cmbxOrderBY.ListIndex = 4 Then
        '    strSQL = strSQL & " order by total desc"
        End If
    
        rPesq.Open strSQL, conn_client
        With rPesq
            Do While Not .EOF
                sMsg = FormatCodigo(!cod_pro) & Chr(9)
                sMsg = sMsg & !nome_pro & Chr(9) & !unidade_pro & Chr(9)
                sMsg = sMsg & FormatVal(!cr_pro) & Chr(9)
                sMsg = sMsg & FormatVal(!lucro_pro) & Chr(9)
                sMsg = sMsg & FormatVal(!le) & Chr(9)
                sMsg = sMsg & FormatVal(!le - !lucro_pro) & Chr(9)
                sMsg = sMsg & FormatVal(!le - F_TaxasEntradas) & Chr(9)
                sMsg = sMsg & FormatVal(!ps) & Chr(9)
                sMsg = sMsg & FormatVal(!pv_pro) & Chr(9)
                sMsg = sMsg & FormatVal(!pv_pro - !ps) & Chr(9)
                
                strSQL = "select sum(ti.qtd_ite) as qtd, sum(ti.qtd_ite*ti.unit_ite) as total, "
                strSQL = strSQL & " sum(ti.qtd_ite*(ti.unit_ite-tp.cr_pro)) as gb, "
                strSQL = strSQL & " sum(ti.qtd_ite*(ti.unit_ite-tp.cr_pro))-sum(ti.qtd_ite*(ti.unit_ite-tp.cr_pro) * " & F_GravaDec(F_TaxasEntradas / 100) & ") as gl "
                strSQL = strSQL & " from " & TAB_PRODUTOS & " as tp, " & TAB_VENDAS & " as tv, " & TAB_ITVENDAS & " as ti "
                strSQL = strSQL & " where tv.num_ven=ti.num_ven and tp.cod_pro=ti.cod_pro "
                strSQL = strSQL & " and tp.cod_pro='" & !cod_pro & "' "
                strSQL = strSQL & " and tv.situacao_ven=4 "
                rVendas.Open strSQL
                With rVendas
                    If .RecordCount > 0 Then
                        sMsg = sMsg & FormatVal(!qtd) & Chr(9)
                        sMsg = sMsg & FormatVal(!total) & Chr(9)
                        sMsg = sMsg & FormatVal(!gb) & Chr(9)
                        sMsg = sMsg & FormatVal(!gl) & Chr(9)
                    Else
                        sMsg = sMsg & "0" & Chr(9) & "0" & Chr(9) & "0" & Chr(9)
                    End If
                End With
                sMsg = sMsg & FormatVal(!er_pro) & Chr(9)
                sMsg = sMsg & FormatVal(!ef_pro) & Chr(9)
                sMsg = sMsg & (!nome_cat) & Chr(9)
                sMsg = sMsg & (!nome_mar)
                Msflxgrd.AddItem sMsg
                P_Status "Aguarde enquanto monto a tabela " & Round((.AbsolutePosition + 1) / .RecordCount * 100) & "%"
                .MoveNext
            Loop
            .Close
        End With
        lblList.caption = "Listados: " & Msflxgrd.rows - 1
    
    
    
    ZebraGrid Msflxgrd, 2, &HD3E7EB

End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rVendas As New ADODB.Recordset
    Dim cTotal As Currency
    Dim sMsg As String
    Msflxgrd.rows = 1
    
        strSQL = "select tp.*,(tp.cr_pro+(tp.cr_pro*tp.lucro_pro/100)) as ps,"
        strSQL = strSQL & "IF (tp.lucro_pro>0 ,(tp.lucro_pro*tp.pv_pro/(tp.cr_pro+(tp.cr_pro*tp.lucro_pro/100))),(tp.cr_pro/tp.pv_pro*100)) as le,"
        strSQL = strSQL & " tm.nome_mar,tc.nome_cat, "
        strSQL = strSQL & " sum(ti.qtd_ite) as qtd, sum(ti.qtd_ite*ti.unit_ite) as total, "
        strSQL = strSQL & " sum(ti.qtd_ite*(ti.unit_ite-tp.cr_pro)) as gb, "
        strSQL = strSQL & " sum(ti.qtd_ite*(ti.unit_ite-tp.cr_pro))-sum(ti.qtd_ite*(ti.unit_ite-tp.cr_pro) * " & F_GravaDec(F_TaxasEntradas / 100) & ") as gl "
        strSQL = strSQL & " from " & TAB_PRODUTOS & " as tp,"
        strSQL = strSQL & " tab_Categorias as tc, " & TAB_MARCAS & " as tm, "
        strSQL = strSQL & " " & TAB_VENDAS & " as tv, " & TAB_ITVENDAS & " as ti "
        strSQL = strSQL & " where tp.cod_Cat=tc.cod_Cat"
        strSQL = strSQL & " and tp.cod_mar=tm.cod_mar "
        strSQL = strSQL & " and tv.num_ven=ti.num_ven "
        strSQL = strSQL & " and ti.cod_pro=tp.cod_pro "
        strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
        strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtbxdataFin.Text)
        If cmbxCat.Text <> "" Then
            strSQL = strSQL & " and  tc.nome_cat='" & cmbxCat.Text & "' "
        End If
        If cmbxMarca.Text <> "" Then
            strSQL = strSQL & " and tm.nome_mar='" & cmbxMarca.Text & "' "
        End If
        
        If txtbxDesc.Text <> "" Then
            strSQL = strSQL & " and tp.nome_pro like '" & txtbxDesc.Text & "%'"
        End If
        strSQL = strSQL & " group by ti.cod_pro "
        If optnAbaixoCrit.Value Then
            strSQL = strSQL & " having tp.lucro_pro > le "
        End If
        If optnInfDesp.Value Then
            'strSQL = strSQL & " having (tp.lucro_pro -  " & F_GravaDec(F_TaxasEntradas) & ")<0 "
            strSQL = strSQL & " having (le -  " & F_GravaDec(F_TaxasEntradas) & ")<0 "
        End If
        
        If cmbxOrderBY.ListIndex = 1 Then
            strSQL = strSQL & " order by ti.cod_pro"
        ElseIf cmbxOrderBY.ListIndex = 2 Then
            strSQL = strSQL & " order by nome_pro"
        ElseIf cmbxOrderBY.ListIndex = 3 Then
            strSQL = strSQL & " order by total desc"
        ElseIf cmbxOrderBY.ListIndex = 4 Then
            strSQL = strSQL & " order by gb desc"
        End If
    
        rPesq.Open strSQL, conn_client
        With rPesq
            Do While Not .EOF
                sMsg = FormatCodigo(!cod_pro) & Chr(9)
                sMsg = sMsg & !nome_pro & Chr(9) & !unidade_pro & Chr(9)
                sMsg = sMsg & FormatVal(!cr_pro) & Chr(9)
                sMsg = sMsg & FormatVal(!lucro_pro) & Chr(9)
                sMsg = sMsg & FormatVal(!le) & Chr(9)
                sMsg = sMsg & FormatVal(!le - !lucro_pro) & Chr(9)
                sMsg = sMsg & FormatVal(!le - F_TaxasEntradas) & Chr(9)
                sMsg = sMsg & FormatVal(!ps) & Chr(9)
                sMsg = sMsg & FormatVal(!pv_pro) & Chr(9)
                sMsg = sMsg & FormatVal(!pv_pro - !ps) & Chr(9)
                sMsg = sMsg & FormatVal(!qtd) & Chr(9)
                sMsg = sMsg & FormatVal(!total) & Chr(9)
                sMsg = sMsg & FormatVal(!gb) & Chr(9)
                sMsg = sMsg & FormatVal(!gl) & Chr(9)
                sMsg = sMsg & FormatVal(!er_pro) & Chr(9)
                sMsg = sMsg & FormatVal(!ef_pro) & Chr(9)
                sMsg = sMsg & (!nome_cat) & Chr(9)
                sMsg = sMsg & (!nome_mar)
                Msflxgrd.AddItem sMsg
                P_Status "Aguarde enquanto monto a tabela " & Round((.AbsolutePosition + 1) / .RecordCount * 100) & "%"
                .MoveNext
            Loop
            .Close
        End With
        lblList.caption = "Listados: " & Msflxgrd.rows - 1
    
    
    
    ZebraGrid Msflxgrd, 2, &HD3E7EB

End Sub

Private Sub cmndbtnImprimir_Click()
    Msflxgrd.Visible = False
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
    Msflxgrd.Visible = True
End Sub
Sub imprime(obj As Object)
    Dim iCount As Integer
    Dim iCol As Integer
    Dim sWith As Single
    Printer.Orientation = vbPRORLandscape
    P_CabRelTimbr "Relatório de Saídas Produtos por Lucratividade", obj
    obj.fontsize = 9
    
    P_ImprX obj, "Marca: " & cmbxMarca.Text, 0
    P_ImprX obj, "Cat: " & cmbxCat.Text, 60, True
        
    P_ImprX obj, "Expr: " & txtbxDesc.Text, 200, False, True
    
    obj.fontsize = 9
    P_ImprLinha obj
    
        With Msflxgrd
            For iCol = 0 To .Cols - 3
                .Row = 0
                .col = iCol
                
                P_ImprX obj, .Text, ScaleX(sWith, 1, 6), , , 1
                sWith = sWith + .CellWidth
            Next
            P_ImprX obj, "", 0, False
            P_ImprLinha obj
            sWith = 0
            If obj.CurrentY >= obj.ScaleWidth - 30 Then
                P_FimRel obj, "Continua..."
                For iCol = 0 To .Cols - 3
                    .Row = 0
                    .col = iCol
                
                    P_ImprX obj, .Text, ScaleX(sWith, 1, 6), , , 1
                    sWith = sWith + .CellWidth
                    Next
                    P_ImprX obj, "", 0, False
            End If
            sWith = 0
            For iCount = 1 To .rows - 1
                .Row = iCount
                For iCol = 0 To .Cols - 3
                    .col = iCol
                
                    P_ImprX obj, .Text, ScaleX(sWith, 1, 6), , , 1
                sWith = sWith + .CellWidth
            Next
            sWith = 0
            P_ImprX obj, "", 0, False
            Next
            P_ImprLinha obj
            
            P_FimRel obj, strNome_Emp
        End With
    
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    ajustaGrid
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF9 Then
        pc.BackColor = vbYellow
    End If
    If KeyCode = vbKeyF8 Then
        pc.BackColor = vbGreen
    End If
End Sub

Private Sub Form_Load()
    Me.Width = 10530
    Me.Height = 7485
    Me.Top = 10
    Me.Left = 10
    P_PreencheMarca cmbxMarca
    P_PreencheCateg cmbxCat
    cmbxOrderBY.ListIndex = 0
    'pc.BackColor = vbGreen
End Sub

Private Sub optnAgrupada_Click()
    ajustaGrid
    
End Sub

Private Sub optnAgrupada_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar.SetFocus
    End If
End Sub

Private Sub optnSeq_Click()
    ajustaGrid
    
End Sub

Private Sub optnSeq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar.SetFocus
    End If
End Sub



Private Sub optnAbaixoCrit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub optnInfDesp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub optnTodos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
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

Private Sub txtbxDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        AjustaTexto txtbxDesc
        busca
    End If
End Sub

