VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
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
   Begin VB.Data dt 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmRelVendas.frx":1350
      Height          =   5535
      Left            =   0
      OleObjectBlob   =   "frmRelVendas.frx":1361
      TabIndex        =   10
      Top             =   120
      Width           =   7815
   End
   Begin MSDBGrid.DBGrid dbgrdSai 
      Bindings        =   "frmRelVendas.frx":2262
      Height          =   855
      Left            =   7920
      OleObjectBlob   =   "frmRelVendas.frx":2276
      TabIndex        =   13
      Top             =   4200
      Width           =   2175
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
      TabIndex        =   12
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
      TabIndex        =   11
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
        SendKeys Chr(9)
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
        " Até " & txtbxDAtaFIn.Text, obj
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
    dt.Connect = Banco.Connect
    dtSai.Connect = Banco.Connect
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
        "sum(valor_sai-desconto_sai)", "tab_saidas", "emissao_sai>=" & _
            F_GravaData(txtbxDataIni.Text) & _
            "  and emissao_sai<=" & F_GravaData(txtbxDAtaFIn.Text)))
    lblTotalGerado.caption = "Total Gerado: " & FormatVal(cTotalGer)
    bActive = False
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select tp.nome_pro,tp.unidade_pro,ti.*,sum(ti.qtd_ite) as soma,"
    strSQL = strSQL & "sum(ti.final_ite) as gerado "
    
    strSQL = strSQL & " from tab_vendas as tv, tab_itvendas as ti,"
    strSQL = strSQL & "tab_produtos as tp where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and ti.num_ven=tv.num_ven "
    
    strSQL = strSQL & " and left(tv.dataefet_ven,10) >=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and left(tv.dataefet_ven,10) <=" & F_GravaData(txtbxDAtaFIn.Text)
    strSQL = strSQL & " and tv.situacao_ven=4 "
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
    
    strSQL = strSQL & " from tab_saidas as ts, tab_itsai as ti,"
    strSQL = strSQL & "tab_produtos as tp where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and ti.nota_sai=ts.nota_sai "
    strSQL = strSQL & " and ts.emissao_sai >=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and ts.emissao_sai <=" & F_GravaData(txtbxDAtaFIn.Text)
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
    P_SelecTexto txtbxDAtaFIn
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxDAtaFIn
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub
