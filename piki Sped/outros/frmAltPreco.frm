VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmAltPreco 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listagem de Alteração de Preços"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   12360
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
      ScaleWidth      =   12330
      TabIndex        =   10
      Top             =   8115
      Width           =   12360
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   8520
         Picture         =   "frmAltPreco.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   10320
         Picture         =   "frmAltPreco.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1695
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
         ItemData        =   "frmAltPreco.frx":0698
         Left            =   3360
         List            =   "frmAltPreco.frx":069A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cmbxCat 
         Height          =   315
         ItemData        =   "frmAltPreco.frx":069C
         Left            =   1320
         List            =   "frmAltPreco.frx":069E
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
   Begin VB.Data dtPro 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmAltPreco.frx":06A0
      Height          =   5415
      Left            =   120
      OleObjectBlob   =   "frmAltPreco.frx":06B4
      TabIndex        =   11
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "frmAltPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxCat_Click()
    cmbxCat.Tag = F_BuscaUmCampo("cod_cat", "tab_categorias", "nome_cat='" & cmbxCat.Text & "'")
    busca
End Sub

Private Sub cmbxCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxCat.ListIndex = -1
    End If
End Sub

Private Sub cmbxMarca_Click()
    cmbxMarca.Tag = F_BuscaUmCampo("cod_mar", "tab_marcas", "nome_mar='" & cmbxMarca.Text & "'")
    busca
End Sub

Private Sub cmbxMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxMarca.ListIndex = -1
    End If
End Sub

Private Sub cmndbtnImprimir_Click()
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.Tag <> "" Then
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
    With dtPro.Recordset
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
            P_ImprX obj, FormatVal(!pV_pro), 160, True, True
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
    
End Sub

Private Sub Form_Load()
    
    txtbxDe.Text = DateAdd("D", -1, Date)
    dtPro.Connect = Banco.Connect
    resizeForm Me
    
    P_PreencheMarca cmbxMarca
    P_PreencheCateg cmbxCat
    
End Sub

Private Sub Form_Resize()
    dbgrd.Height = pcBusca.Top - 200
    dbgrd.Width = Me.Width - 300
End Sub

Private Sub txtbxDe_GotFocus()
    P_SelecTexto txtbxDe
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDe_LostFocus()
    P_AjustaData txtbxDe
    busca
End Sub
Sub busca()
    Dim strSQL As String
    P_AjustaTexto txtbxExpr
    
    strSQL = "select  tp.nome_pro,tp.pv_pro,ta.* from tab_produtos as tp, tab_altpreco as ta, "
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
    
    
    
    dtPro.RecordSource = strSQL
    dtPro.Refresh
End Sub

Private Sub txtbxExpr_Change()
    busca
End Sub
