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
      TabPicture(0)   =   "frmRelProdLucr.frx":0AB2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Composição"
      TabPicture(1)   =   "frmRelProdLucr.frx":0ACE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command2"
      Tab(1).Control(1)=   "Text1"
      Tab(1).Control(2)=   "cmndbtnMatriz"
      Tab(1).Control(3)=   "txtbxPri"
      Tab(1).Control(4)=   "Label1"
      Tab(1).Control(5)=   "Label2"
      Tab(1).Control(6)=   "Label4"
      Tab(1).ControlCount=   7
      Begin VB.CommandButton Command2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -70080
         Picture         =   "frmRelProdLucr.frx":0AEA
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74040
         TabIndex        =   24
         Top             =   840
         Width           =   3975
      End
      Begin VB.CommandButton cmndbtnMatriz 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -70080
         Picture         =   "frmRelProdLucr.frx":0BE3
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox txtbxPri 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74040
         TabIndex        =   22
         Top             =   480
         Width           =   3975
      End
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
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Princípio:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Subseção:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   16
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Descrição:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   17
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
      TabIndex        =   19
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   0
         Width           =   6285
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   5055
      Left            =   0
      TabIndex        =   18
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
Sub ajustagrid()
    With msflxgrd
        .Cols = 19
        .FixedCols = 0
        .rows = 1
        .Col = 0: .Text = "Código"
        .Col = 1: .Text = "Descrição"
        .Col = 2: .Text = "Un."
        .Col = 3: .Text = "Custo"
        .Col = 4: .Text = "L.S(%)"
        .Col = 5: .Text = "L.E(%)"
        .Col = 6: .Text = "Dif(%)"
        .Col = 7: .Text = "L.L(%)"
        .Col = 8: .Text = "P.S(R$)"
        .Col = 9: .Text = "P.E(R$)"
        .Col = 10: .Text = "Dif.(R$)"
        .Col = 11: .Text = "Q.G."
        .Col = 12: .Text = "V.G."
        .Col = 13: .Text = "G.B."
        .Col = 14: .Text = "G.L."
        .Col = 15: .Text = "Est"
        .Col = 16: .Text = "Disp"
        .Col = 17: .Text = "Categoria"
        .Col = 18: .Text = "Marca"
        
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
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    If KeyCode = vbKeyF5 Then
        'P_PreencheCateg cmbxCat
    End If
End Sub

Private Sub cmbxMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    If KeyCode = vbKeyF5 Then
        
    End If
End Sub

Private Sub cmbxOrderBY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxTurno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxTurno.ListIndex = -1
    End If
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim sTime As String
    sTime = Now
    msflxgrd.Visible = False
    P_Status "Montando Busca..."
    busca
    P_Status ""
    msflxgrd.Visible = True
    MsgBox "Pesquisa concluída em " & DateDiff("s", sTime, Now) & " segundos!", vbInformation
End Sub
Sub busca1()
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    Dim rVendas As adodb.Recordset
    Dim cTotal As Currency
    Dim sMsg As String
    msflxgrd.rows = 1
    
        strSQL = "select tp.*,(tp.cr_pro+(tp.cr_pro*tp.lucro_pro/100)) as ps,"
        strSQL = strSQL & "IF (tp.lucro_pro>0 ,(tp.lucro_pro*tp.pv_pro/(tp.cr_pro+(tp.cr_pro*tp.lucro_pro/100))),(tp.cr_pro/tp.pv_pro*100)) as le,"
        strSQL = strSQL & " tpr.nome_pri,tu.nome_sub,ts.nome_sec  "
        
        strSQL = strSQL & " from tab_produtos as tp,"
        strSQL = strSQL & " tab_principios as tpr,tab_subsec as tu, tab_secoes as ts"
        strSQL = strSQL & " tab_vendas as tv, tab_itvendas as ti "
        strSQL = strSQL & " where tp.cod_sub=tu.cod_sub"
        strSQL = strSQL & " and tu.cod_sec=ts.cod_sec "
        strSQL = strSQL & " and tpr.cod_pri=tp.cod_pri "
        strSQL = strSQL & " and tv.num_ven=ti.num_ven "
        strSQL = strSQL & " and ti.cod_pro=tp.cod_pro "
        strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
        strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtbxdataFin.Text)
        
        'If cmbxCat.Text <> "" Then
        '    strSQL = strSQL & " and  tc.nome_cat='" & cmbxCat.Text & "' "
        'End If
        
        'If cmbxMarca.Text <> "" Then
        '    strSQL = strSQL & " and tm.nome_mar='" & cmbxMarca.Text & "' "
        'End If
        
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
    
        abreObjetoAdo rPesq, strSQL
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
                strSQL = strSQL & " from tab_produtos as tp, tab_vendas as tv, tab_itvendas as ti "
                strSQL = strSQL & " where tv.num_ven=ti.num_ven and tp.cod_pro=ti.cod_pro "
                strSQL = strSQL & " and tp.cod_pro='" & !cod_pro & "' "
                strSQL = strSQL & " and tv.situacao_ven=4 "
                abreObjetoAdo rVendas, strSQL
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
                msflxgrd.AddItem sMsg
                P_Status "Aguarde enquanto monto a tabela " & Round((.AbsolutePosition + 1) / .RecordCount * 100) & "%"
                .MoveNext
            Loop
            .Close
        End With
        lblList.caption = "Listados: " & msflxgrd.rows - 1
    
    
    
    ZebraGrid msflxgrd, 2, &HD3E7EB

End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    Dim rVendas As New adodb.Recordset
    Dim cTotal As Currency
    Dim sMsg As String
    msflxgrd.rows = 1
    
        strSQL = "select tp.*,(tp.cr_pro+(tp.cr_pro*tp.lucro_pro/100)) as ps,sum(er_pro) as estoque,sum(ef_pro) as ef,"
        strSQL = strSQL & "IF (tp.lucro_pro>0 ,(tp.lucro_pro*tp.pv_pro/(tp.cr_pro+(tp.cr_pro*tp.lucro_pro/100))),(tp.cr_pro/tp.pv_pro*100)) as le,"
        strSQL = strSQL & " nome_sub,nome_sec,nome_pri, "
        strSQL = strSQL & " sum(ti.qtd_ite) as qtd, sum(ti.qtd_ite*ti.unit_ite) as total, "
        strSQL = strSQL & " sum(ti.qtd_ite*(ti.unit_ite-tp.cr_pro)) as gb, "
        strSQL = strSQL & " sum(ti.qtd_ite*(ti.unit_ite-tp.cr_pro))-sum(ti.qtd_ite*(ti.unit_ite-tp.cr_pro) * " & F_GravaDec(F_TaxasEntradas / 100) & ") as gl "
        strSQL = strSQL & " from tab_produtos as tp,"
        strSQL = strSQL & " tab_principios as tpr,tab_subsec as tu, tab_secoes as ts, "
        strSQL = strSQL & " tab_vendas as tv, tab_itvendas as ti,tab_estoque as te  "
        strSQL = strSQL & " where tv.num_ven=ti.num_ven "
        strSQL = strSQL & " and tp.cod_pri=tpr.cod_pri "
        strSQL = strSQL & " and ts.cod_sec=tu.cod_sec "
        strSQL = strSQL & " and tp.cod_sub=tu.cod_sub "
        strSQL = strSQL & " and ti.cod_pro=tp.cod_pro "
        strSQL = strSQL & " and tp.cod_pro=te.cod_pro "
        strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
        strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtbxdataFin.Text)
        'If cmbxCat.Text <> "" Then
        '    strSQL = strSQL & " and  tc.nome_cat='" & cmbxCat.Text & "' "
        'End If
        'If cmbxMarca.Text <> "" Then
        '    strSQL = strSQL & " and tm.nome_mar='" & cmbxMarca.Text & "' "
        'End If
        
        'If txtbxDesc.Text <> "" Then
        '    strSQL = strSQL & " and tp.nome_pro like '" & txtbxDesc.Text & "%'"
        'End If
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
    
        abreObjetoAdo rPesq, strSQL
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
                sMsg = sMsg & FormatVal(!Estoque) & Chr(9)
                sMsg = sMsg & FormatVal(!ef) & Chr(9)
                sMsg = sMsg & (!nome_pri) & Chr(9)
                sMsg = sMsg & (!nome_sub)
                msflxgrd.AddItem sMsg
                P_Status "Aguarde enquanto monto a tabela " & Round((.AbsolutePosition + 1) / .RecordCount * 100) & "%"
                .MoveNext
            Loop
            .Close
        End With
        lblList.caption = "Listados: " & msflxgrd.rows - 1
    
    
    
    ZebraGrid msflxgrd, 2, &HD3E7EB

End Sub

Private Sub cmndbtnImprimir_Click()
    msflxgrd.Visible = False
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
    msflxgrd.Visible = True
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
    
        With msflxgrd
            For iCol = 0 To .Cols - 3
                .row = 0
                .Col = iCol
                
                P_ImprX obj, .Text, ScaleX(sWith, 1, 6), , , 1
                sWith = sWith + .CellWidth
            Next
            P_ImprX obj, "", 0, False
            P_ImprLinha obj
            sWith = 0
            If obj.CurrentY >= obj.ScaleWidth - 30 Then
                P_FimRel obj, "Continua..."
                For iCol = 0 To .Cols - 3
                    .row = 0
                    .Col = iCol
                
                    P_ImprX obj, .Text, ScaleX(sWith, 1, 6), , , 1
                    sWith = sWith + .CellWidth
                    Next
                    P_ImprX obj, "", 0, False
            End If
            sWith = 0
            For iCount = 1 To .rows - 1
                .row = iCount
                For iCol = 0 To .Cols - 3
                    .Col = iCol
                
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
    ajustagrid
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
    Me.width = 10530
    Me.height = 7485
    Me.Top = 10
    Me.Left = 10
    'P_PreencheMarca cmbxMarca
    'P_PreencheCateg cmbxCat
    cmbxOrderBY.ListIndex = 0
    'pc.BackColor = vbGreen
End Sub

Private Sub optnAgrupada_Click()
    ajustagrid
    
End Sub

Private Sub optnAgrupada_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar.SetFocus
    End If
End Sub

Private Sub optnSeq_Click()
    ajustagrid
    
End Sub

Private Sub optnSeq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar.SetFocus
    End If
End Sub



Private Sub optnAbaixoCrit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub optnInfDesp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub optnTodos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxdataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        busca
    End If
End Sub

