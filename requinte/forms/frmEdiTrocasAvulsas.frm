VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiTrocasAvulsas 
   Caption         =   "Trocas Avulsas Efetuadas"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14370
   Icon            =   "frmEdiTrocasAvulsas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   14370
   Begin VB.PictureBox pcTop 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1305
      ScaleWidth      =   14340
      TabIndex        =   15
      Top             =   7110
      Width           =   14370
      Begin VB.TextBox txtVendedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Imprimir Troca"
         Height          =   615
         Left            =   9720
         Picture         =   "frmEdiTrocasAvulsas.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "Imprimir Lista"
         Height          =   615
         Left            =   11280
         Picture         =   "frmEdiTrocasAvulsas.frx":06D4
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   12840
         Picture         =   "frmEdiTrocasAvulsas.frx":0AE1
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox txtProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   4575
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         TabIndex        =   5
         Top             =   375
         Width           =   975
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         TabIndex        =   7
         Top             =   375
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   8160
         Picture         =   "frmEdiTrocasAvulsas.frx":0C4C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnLocaliza 
         Height          =   315
         Left            =   4680
         Picture         =   "frmEdiTrocasAvulsas.frx":0FD6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdVendedor 
         Height          =   315
         Left            =   2760
         Picture         =   "frmEdiTrocasAvulsas.frx":1360
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Produto"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5040
         TabIndex        =   4
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6120
         TabIndex        =   6
         Top             =   120
         Width           =   285
      End
   End
   Begin MSDataGridLib.DataGrid dtg 
      Bindings        =   "frmEdiTrocasAvulsas.frx":16EA
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   12303
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "Id"
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
         DataField       =   "cliente"
         Caption         =   "Cliente"
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
      BeginProperty Column02 
         DataField       =   "telefone"
         Caption         =   "telefone"
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
      BeginProperty Column03 
         DataField       =   "vendedor"
         Caption         =   "Vendedor"
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
      BeginProperty Column04 
         DataField       =   "valor_in"
         Caption         =   "Entrou"
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
      BeginProperty Column05 
         DataField       =   "valor_out"
         Caption         =   "Saiu"
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
      BeginProperty Column06 
         DataField       =   "desconto"
         Caption         =   "Desconto"
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
      BeginProperty Column07 
         DataField       =   "obs"
         Caption         =   "Obs"
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
            ColumnWidth     =   510,236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2729,764
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2670,236
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   330
      Left            =   120
      Top             =   7680
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
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
      Appearance      =   0
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
      Caption         =   ""
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
End
Attribute VB_Name = "frmEdiTrocasAvulsas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmbxBusca_Click()
    
End Sub

Private Sub cmdVendedor_Click()
   Dim f As funcionario.funcionario
    f = funcionario.F_SelectVendedor(txtVendedor.Text)
    If f.cod_fun <> "" Then
        txtVendedor.Text = f.guerra_fun
        txtVendedor.tag = f.cod_fun
        cmbxBusca_Click
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select distinct tt.*  from tab_produtos as tp, tab_trocaavulsa as tt, "
    strSQL = strSQL & " tab_trocaavulsait as ti where ti.id_tro=tt.id "
    strSQL = strSQL & " and ti.codigo=tp.cod_pro "
    If txtProduto.tag <> "" Then
        strSQL = strSQL & " and tp.cod_pro='" & txtProduto.tag & "' "
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and data>=" & F_GravaData(txtbxDataIni.Text)
    End If
    
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and data<=" & F_GravaData(txtbxdataFin.Text)
    End If
    
    If txtVendedor.tag <> "" Then
        strSQL = strSQL & " and tt.vendedor='" & txtVendedor.tag & "'"
    End If
        
    abreObjetoAdo ado, strSQL, dtg
    
End Sub
Sub soma()
    Dim cTotalTro As Currency
    Dim cTotalVen As Currency
    With ado
        Do While Not .EOF
            cTotalTro = cTotalTro + !valor_in
            cTotalVen = cTotalVen + !valor_out
            .MoveNext
        Loop
        adc.Refresh
    End With
    
End Sub

Private Sub cmndbtnImpr_Click()
    On Error Resume Next
    P_DbgrdToHtml "Trocas Efetuadas", dtg, adc, "De: " & txtbxDataIni.Text & " Até " & txtbxdataFin.Text
    
End Sub

Private Sub cmndbtnLocaliza_Click()
    buscaProduto
End Sub
Sub buscaProduto()
    Dim p As Produto
    p = F_BuscaProdSaida(txtProduto.Text)
    If p.cod_pro <> "" Then
        txtProduto.Text = p.nome_pro
        txtProduto.tag = p.cod_pro
        busca
        
        
        
        
    End If
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    On Error GoTo trata_Erro
    imprimeTrocaAvulsa ado!id, frmVisualiza.Pic
    frmVisualiza.Show 1
    
    If frmVisualiza.tag <> "" Then
        imprimeTrocaAvulsa ado!id, Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
    
    Exit Sub
    
    
trata_Erro:
    P_MostraErro
    
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    resizeForm Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    PosicionaGrade dtg, Me, pcTop
    posicionaADireita cmndbtnSair, pcTop
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

Sub buscaCli()
    frmSelectCli.txtbxNome.tag = "C"
    frmSelectCli.txtbxNome.Text = txtbxExpr.Text
    
    On Error Resume Next
    
    frmSelectCli.Show 1
    
    
    
    txtbxExpr.Text = frmSelectCli.adoR!nome
    txtbxExpr.tag = frmSelectCli.adoR!cod
    
    
    
    
    Unload frmSelectCli
End Sub
Sub buscaProd()
    Dim p As Produto
    p = F_BuscaProdSaida(txtbxExpr.Text)
    txtbxExpr.Text = p.nome_pro
    txtbxExpr.tag = p.cod_pro
End Sub
Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    
        If cmbxBusca.ListIndex = 2 Then
             buscaCli
        ElseIf cmbxBusca.ListIndex = 3 Then
            buscaProd
        End If
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtVendedor_Change()
    txtVendedor.tag = ""
End Sub
