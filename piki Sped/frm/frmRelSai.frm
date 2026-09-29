VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelSai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Saídas"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   10785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      Begin VB.ComboBox cmbxOrderBY 
         Height          =   315
         ItemData        =   "frmRelSai.frx":0000
         Left            =   3120
         List            =   "frmRelSai.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   465
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   495
         Left            =   8760
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   495
         Left            =   6960
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   495
         Left            =   5160
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtbxDAtaFIn 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtbxDataIni 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   3120
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   240
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
      Height          =   735
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   4080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmRelSai.frx":0030
      Height          =   5535
      Left            =   0
      OleObjectBlob   =   "frmRelSai.frx":0041
      TabIndex        =   10
      Top             =   1200
      Width           =   10575
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
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
      Left            =   8760
      TabIndex        =   12
      Top             =   6720
      Width           =   510
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
      Left            =   6000
      TabIndex        =   11
      Top             =   6720
      Width           =   1185
   End
End
Attribute VB_Name = "frmRelSai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
    If Dt.Recordset.RecordCount <= 0 Then
        MsgBox "Não há o quer ser impresso!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma Impressão do Relatório?", vbYesNo + vbQuestion) = vbYes Then
        imprime Printer
    End If
End Sub
Sub imprime(obj As Object)
    Dim iItens As Integer
    Dim iPag As Integer
    Dim strSQL As String
    P_CabRelTimbr "Relatório de Saídas no Período " & de & txtbxDataIni.Text & _
        " Até " & txtbxdataFin.Text, obj
    P_ImprX obj, "Código", 0
    P_ImprX obj, "Descrição", 20
    P_ImprX obj, "UN", 120
    P_ImprX obj, "QTD", 140, True, True
    P_ImprX obj, "Vlr Ger.", 160, True, True
    P_ImprX obj, "Total", 180, False, True
    
    P_ImprLinha obj
    
    With Dt.Recordset
        
        .MoveFirst
        strSQL = CStr(.RecordCount / 50)
        If InStr(1, strSQL, ",") <> 0 Then
            iPag = Int(strSQL) + 1
        End If
        
        Do While Not .EOF
            iItens = iItens + 1
            If iItens = 50 Then
                P_Rodape "Continua...", True, "Pág:" & Printer.Page & " de " & iPag
                iItens = 0
                P_ImprX obj, "Código", 0
                P_ImprX obj, "Descrição", 20
                P_ImprX obj, "UN", 120
                P_ImprX obj, "QTD", 140, True, True
                P_ImprX obj, "Vlr Ger.", 160, True, True
                P_ImprX obj, "Total", 180, False, True
    
                P_ImprLinha obj
            End If
            P_ImprX obj, !cod_pro, 0
            P_ImprX obj, !nome_pro, 20
            P_ImprX obj, !unidade_Pro, 120
            P_ImprX obj, FormatVal(!soma), 140, True, True
            P_ImprX obj, FormatVal(!gerado), 160, True, True
            P_ImprX obj, FormatVal(!Total), 180, False, True
           
            .MoveNext
        Loop
        P_ImprLinha obj
        P_ImprX obj, lblTotalGerado.caption, 80, True
        P_ImprX obj, lblTotal.caption, 130, False
        P_Rodape strNome_Emp, True, "Pág: " & Printer.Page & " de " & Printer.Page
    End With
    
    
    Printer.EndDoc
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dt.Connect = banco.Connect
    cmbxOrderBy.ListIndex = 0
End Sub
Sub soma()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim cTotalGer As Currency
    With Dt.Recordset
        Do While Not .EOF
            cTotal = cTotal + !Total
            cTotalGer = cTotalGer + !gerado
            .MoveNext
        Loop
        
    End With
    lblTotal.caption = "Total: " & FormatVal(cTotal)
    lblTotalGerado.caption = "Total Gerado: " & FormatVal(cTotalGer)
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select tp.nome_pro,tp.unidade_pro,ti.*,sum(ti.qtd_its) as soma,"
    strSQL = strSQL & "sum(ti.qtd_its*ti.unit_its) as gerado,"
    strSQL = strSQL & "sum(ti.qtd_its*tp.preco_pro) as total "
    strSQL = strSQL & " from tab_saidas as ts, tab_itsai as ti,"
    strSQL = strSQL & "tab_produtos as tp where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and ti.nota_sai=ts.nota_sai and ts.serie_sai=ti.serie_sai "
    strSQL = strSQL & " and ts.emissao_sai>=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and ts.emissao_Sai <=" & F_GravaData(txtbxdataFin.Text)
    strSQL = strSQL & " group by ti.cod_pro "
    
    strSQL = strSQL & " order by "
    If cmbxOrderBy.ListIndex = 0 Then
        strSQL = strSQL & "tp.nome_pro"
    ElseIf cmbxOrderBy.ListIndex = 1 Then
        strSQL = strSQL & "tp.cod_Pro "
    Else
        strSQL = strSQL & "gerado "
    End If
    Dt.RecordSource = strSQL
    Dt.Refresh
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDAtaFIn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
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
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub
