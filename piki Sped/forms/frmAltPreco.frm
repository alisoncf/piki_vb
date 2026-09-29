VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAltPreco 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listagem de Alteração de Preços"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   14565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcBusca 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   14535
      TabIndex        =   10
      Top             =   7575
      Width           =   14565
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   595
         Left            =   9960
         Picture         =   "frmAltPreco.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdPesquisar 
         Appearance      =   0  'Flat
         Caption         =   "Pesquisar"
         Height          =   595
         Left            =   8520
         Picture         =   "frmAltPreco.frx":042F
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir"
         Height          =   595
         Left            =   11400
         Picture         =   "frmAltPreco.frx":07DA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   595
         Left            =   12960
         Picture         =   "frmAltPreco.frx":0924
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtbxDe 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cmbxMarca 
         Height          =   315
         ItemData        =   "frmAltPreco.frx":0E72
         Left            =   3360
         List            =   "frmAltPreco.frx":0E74
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cmbxCat 
         Height          =   315
         ItemData        =   "frmAltPreco.frx":0E76
         Left            =   1320
         List            =   "frmAltPreco.frx":0E78
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtbxExpr 
         Height          =   285
         Left            =   5520
         MaxLength       =   40
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apartir de:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Marca:"
         Height          =   255
         Left            =   3360
         TabIndex        =   4
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Categoria"
         Height          =   195
         Left            =   1320
         TabIndex        =   2
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&xpressão:"
         Height          =   195
         Left            =   5520
         TabIndex        =   6
         Top             =   120
         Width           =   780
      End
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   330
      Left            =   1320
      Top             =   4920
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmAltPreco.frx":0E7A
      Height          =   7215
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   12726
      _Version        =   393216
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
Attribute VB_Name = "frmAltPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmbxCat_Click()
    cmbxCat.tag = F_BuscaUmCampo("cod_cat", "tab_categorias", "nome_cat='" & cmbxCat.Text & "'")
    busca
End Sub

Private Sub cmbxCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    If KeyCode = vbKeyDelete Then
        cmbxCat.ListIndex = -1
    End If
End Sub

Private Sub cmbxMarca_Click()
    cmbxMarca.tag = F_BuscaUmCampo("cod_mar", "tab_marcas", "nome_mar='" & cmbxMarca.Text & "'")
    busca
End Sub

Private Sub cmbxMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    If KeyCode = vbKeyDelete Then
        cmbxMarca.ListIndex = -1
    End If
End Sub

Private Sub cmdExportar_Click()
    ExportaGridToCSV dtgrd, adc
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
    P_CabRelTimbr "Listagem de Alteração de Preços ", obj
    
    obj.Print ""
    Dim yIni As Single, yFin As Single
    yIni = obj.CurrentY
    P_ImprX obj, "Código", 5
    P_ImprX obj, "Produto", 25
    P_ImprX obj, "Pç. Antigo", 115, True, True
    P_ImprX obj, "Novo Preço", 140, True, True
    P_ImprX obj, "Preço Atual", 160, True, True
    P_ImprX obj, "Alterado Em", 184, False, True
    P_Line obj, 4, obj.CurrentY, 200, obj.CurrentY
    With adc.Recordset
        .MoveFirst
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 20 Then
                
                If TypeName(obj) = "Printer" Then
                    Printer.NewPage
                    yIni = obj.CurrentY
                End If
            End If
            
            P_ImprX obj, !cod_pro, 5
            P_ImprX obj, !nome_pro, 25
            P_ImprX obj, FormatVal(!antigo_pro), 115, True, True
            P_ImprX obj, FormatVal(!novo_pro), 140, True, True
            P_ImprX obj, FormatVal(!pv_pro), 160, True, True
            P_ImprX obj, !data_alt, 164, False
            P_Line obj, 4, obj.CurrentY, 200, obj.CurrentY, vbDot
            
            
            
            .MoveNext
        Loop
        yFin = obj.CurrentY
        P_Line obj, 4, obj.CurrentY, 200, obj.CurrentY
        P_Line obj, 4, yIni, 4, yFin
        P_Line obj, 200, yIni, 200, yFin
        
    End With
    
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    busca
    Form_Resize
End Sub

Private Sub Form_Load()
    
    txtbxDe.Text = DateAdd("D", -1, Date)
    adc.ConnectionString = strConnect
    resizeForm Me
    
    P_PreencheMarca cmbxMarca
    P_PreencheCateg cmbxCat
    
End Sub

Private Sub Form_Resize()
    PosicionaGrade dtgrd, Me, pcBusca
    posicionaADireita cmndbtnSair, pcBusca
    
End Sub

Private Sub txtbxDe_GotFocus()
    ganhaFoco txtbxDe
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDe_LostFocus()
    P_AjustaData txtbxDe
    busca
End Sub
Sub busca()
    Dim strSQL As String
    AjustaTexto txtbxExpr
    
    strSQL = "select  tp.nome_pro,tp.pv_pro,ta.* from " & TAB_PRODUTOS & " as tp, tab_altpreco as ta, "
    strSQL = strSQL & " tab_marcas as tm, tab_categorias as tc "
    strSQL = strSQL & " where tp.cod_cat=tc.cod_cat "
    strSQL = strSQL & " and tm.cod_mar=tp.cod_mar "
    strSQL = strSQL & " and tp.cod_pro=ta.cod_pro "
    strSQL = strSQL & " and left(ta.data_alt,10) >=" & F_GravaData(txtbxDe.Text)
    If cmbxMarca.Text <> "" Then
        strSQL = strSQL & " and tm.nome_mar='" & cmbxMarca.Text & "' "
    End If
    
    If cmbxCat.Text <> "" Then
        strSQL = strSQL & " and tc.nome_Cat='" & cmbxCat.Text & "' "
    End If
    If txtbxExpr.Text <> "" Then
        strSQL = strSQL & " and nome_pro like '" & txtbxExpr.Text & "%' "
    End If
    
    strSQL = strSQL & " order by data_alt "
    
    
    
    adc.RecordSource = strSQL
    adc.Refresh
End Sub

Private Sub txtbxExpr_Change()
    busca
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub
