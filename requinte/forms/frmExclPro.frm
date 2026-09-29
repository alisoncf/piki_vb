VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmExclPro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exclusão de Produtos"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   Icon            =   "frmExclPro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9615
   Begin VB.Frame Frame3 
      Caption         =   "Opções de Busca"
      Height          =   855
      Left            =   4320
      TabIndex        =   22
      Top             =   0
      Width           =   5295
      Begin VB.TextBox txtbxExpr 
         Height          =   285
         Left            =   3000
         MaxLength       =   40
         TabIndex        =   25
         Top             =   480
         Width           =   2175
      End
      Begin VB.ComboBox cmbxBusca 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmExclPro.frx":128A
         Left            =   120
         List            =   "frmExclPro.frx":129D
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmExclPro.frx":12D1
         Left            =   1440
         List            =   "frmExclPro.frx":12DE
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Buscar Por:"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&xpressão:"
         Height          =   195
         Left            =   3000
         TabIndex        =   27
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   1440
         TabIndex        =   26
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Composição:"
      Height          =   855
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   4215
      Begin VB.ComboBox cmbxCat 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox cmbxMarca 
         Height          =   315
         Left            =   2160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Categoria"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Marca:"
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblCat 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   1800
         TabIndex        =   19
         Tag             =   "!"
         Top             =   240
         Width           =   270
      End
      Begin VB.Label lblMar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   3840
         TabIndex        =   18
         Tag             =   "!"
         Top             =   240
         Width           =   270
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Produtos:"
      Height          =   4935
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   9615
      Begin VB.CommandButton cmndbtnDesabilita 
         Caption         =   "Desa&bilita"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   6360
         Picture         =   "frmExclPro.frx":1300
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnClear 
         Caption         =   "&Limpa Grid"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   4800
         Picture         =   "frmExclPro.frx":144A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpa o Grid Abaixo para que se possa selecionar outros dois produtos."
         Top             =   2760
         Width           =   1335
      End
      Begin VB.PictureBox PctSelecionaProd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   2100
         ScaleHeight     =   1185
         ScaleWidth      =   5865
         TabIndex        =   11
         Top             =   3540
         Width           =   5895
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5600
            Picture         =   "frmExclPro.frx":1594
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   13
            Top             =   12
            Width           =   255
         End
         Begin VB.CommandButton cmndbtnConf 
            Height          =   440
            Left            =   5200
            Picture         =   "frmExclPro.frx":16DE
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Confirma exclusão do produto selecionado."
            Top             =   430
            Width           =   485
         End
         Begin VB.OptionButton OptnCodPro2 
            Caption         =   "OptnCodPro2"
            Height          =   495
            Left            =   120
            TabIndex        =   5
            Top             =   680
            Width           =   4935
         End
         Begin VB.OptionButton OptnCodPro1 
            Caption         =   "OptnCodPro1"
            Height          =   495
            Left            =   120
            TabIndex        =   4
            Top             =   180
            Width           =   4935
         End
         Begin VB.Shape Shape1 
            Height          =   1180
            Left            =   6
            Top             =   0
            Width           =   5870
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Selecione a mercadoria a ser excluída:"
            Height          =   195
            Left            =   1560
            TabIndex        =   12
            Top             =   10
            Width           =   2760
         End
      End
      Begin VB.CommandButton cmndbtnAgrupa 
         Caption         =   "&Agrupar Produtos"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   3240
         Picture         =   "frmExclPro.frx":1FA8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Mescla os dois produtos abaixo, excluindo o selecionado."
         Top             =   2760
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrd 
         Height          =   1335
         Left            =   120
         TabIndex        =   10
         Top             =   3480
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2355
         _Version        =   393216
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.CommandButton cmndbtnDupl 
         Caption         =   "&Duplicados"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   1680
         Picture         =   "frmExclPro.frx":20F2
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Insere o produto selecionado acima no Grid abaixo."
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnExcluir 
         Caption         =   "&Excluir Produto"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   120
         Picture         =   "frmExclPro.frx":223C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exclui o produto do Banco de Dados."
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Data dtPro 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   4080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmExclPro.frx":2386
         Height          =   2415
         Left            =   120
         OleObjectBlob   =   "frmExclPro.frx":239A
         TabIndex        =   0
         Top             =   240
         Width           =   9375
      End
      Begin VB.Shape Shape2 
         Height          =   2450
         Left            =   110
         Top             =   230
         Width           =   9400
      End
      Begin VB.Label lblCodPro 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   2400
         Width           =   480
      End
      Begin VB.Label lblList 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Listadas:"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   7320
         TabIndex        =   8
         Tag             =   "!"
         Top             =   0
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmExclPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sCodPro1 As String
Dim sCodPro2 As String
Dim Flag As Integer
Dim iAux As Integer

Private Sub cmbxBusca_Click()
    If txtbxExpr.Text <> "" Then
        busca
    End If
End Sub
Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxOrd_Click()
    If txtbxExpr.Text <> "" Then
        busca
    End If
End Sub

Private Sub cmbxOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnAgrupa_Click()
    Dim strSQL As String
    With msflxgrd
        If .Rows = 1 Then
            MsgBox "Não existem produtos no Grid!", vbExclamation + vbCritical
            Exit Sub
        ElseIf .Rows = 2 Then
            MsgBox "Só existe um produto no Grid, você deve selecionar dois produtos!", vbExclamation + vbCritical
            Exit Sub
        End If
    
    MsgBox "Ao agrupar dois produtos, todas as referências nas tabelas de compras e vendas serão atualizadas. O estoque será o resultado da soma dos estoques das duas mercadorias.", vbInformation
    
    OptnCodPro1.caption = sCodPro1 & " - " & F_BuscaUmCampo("nome_pro", "tab_produtos", "cod_pro ='" & sCodPro1 & "'")
    OptnCodPro2.caption = sCodPro2 & " - " & F_BuscaUmCampo("nome_pro", "tab_produtos", "cod_pro ='" & sCodPro2 & "'")
    
    PctSelecionaProd.Visible = True
    
    End With
End Sub

Private Sub cmndbtnClear_Click()
    msflxgrd.Rows = 1
    iAux = 0
End Sub

Private Sub CmndbtnConf_Click()
    Dim strSQL As String
    Dim Ep1, Ep2, Ep3, Ep4 As Integer
    Dim Er As Integer
    On Error Resume Next
    
    If OptnCodPro1 = False And OptnCodPro2 = False Then
        MsgBox "Não foi selecionado o produto a excluir!", vbCritical
        Exit Sub
    End If
    
    Ep1 = F_BuscaUmCampo("e1_pro", "tab_composicao", "cod_pro ='" & sCodPro1 & "'")
    Ep2 = F_BuscaUmCampo("e1_pro", "tab_composicao", "cod_pro ='" & sCodPro2 & "'")
    Ep3 = F_BuscaUmCampo("e2_pro", "tab_composicao", "cod_pro ='" & sCodPro1 & "'")
    Ep4 = F_BuscaUmCampo("e2_pro", "tab_composicao", "cod_pro ='" & sCodPro2 & "'")
        
    Er = (F_BuscaUmCampo("r_pro", "tab_composicao", "cod_pro ='" & sCodPro1 & "'")) + _
    (F_BuscaUmCampo("r_pro", "tab_composicao", "cod_pro ='" & sCodPro2 & "'"))

    If strExclProd_usu = "S" Then
        If OptnCodPro1 Then
            strSQL = "update tab_composicao set "
            strSQL = strSQL & "e1_pro=" & F_GravaDec(Ep1 + Ep2) & ", "
            strSQL = strSQL & "e2_pro=" & F_GravaDec(Ep3 + Ep4) & ", "
            strSQL = strSQL & "r_pro=" & F_GravaDec(Er)
            strSQL = strSQL & " where cod_pro='" & sCodPro2 & "' "
            Banco.Execute strSQL
            
            strSQL = ""
            strSQL = strSQL & "update tab_itcompras set "
            strSQL = strSQL & "cod_pro ='" & sCodPro2 & "'"
            strSQL = strSQL & " where cod_pro='" & sCodPro1 & "' "
            Banco.Execute strSQL
            strSQL = ""
            
            strSQL = strSQL & "update tab_itcmp set "
            strSQL = strSQL & "cod_pro ='" & sCodPro2 & "'"
            strSQL = strSQL & " where cod_pro='" & sCodPro1 & "' "
            Banco.Execute strSQL
            strSQL = ""
            
            strSQL = strSQL & "update tab_itvendas set "
            strSQL = strSQL & "cod_pro ='" & sCodPro2 & "'"
            strSQL = strSQL & " where cod_pro='" & sCodPro1 & "' "
            Banco.Execute strSQL
            strSQL = ""
            
            strSQL = strSQL & "update tab_itsai set "
            strSQL = strSQL & "cod_pro ='" & sCodPro2 & "'"
            strSQL = strSQL & " where cod_pro='" & sCodPro1 & "' "
            Banco.Execute strSQL
            strSQL = ""
            
            strSQL = strSQL & "update tab_itperdas set "
            strSQL = strSQL & "cod_pro ='" & sCodPro2 & "'"
            strSQL = strSQL & " where cod_pro='" & sCodPro1 & "' "
            Banco.Execute strSQL
            strSQL = ""
            
            strSQL = strSQL & "update tab_itcont set "
            strSQL = strSQL & "cod_pro ='" & sCodPro2 & "'"
            strSQL = strSQL & " where cod_pro='" & sCodPro1 & "' "
            Banco.Execute strSQL
            strSQL = ""
            
            strSQL = strSQL & "update tab_itdev set "
            strSQL = strSQL & "cod_pro ='" & sCodPro2 & "'"
            strSQL = strSQL & " where cod_pro='" & sCodPro1 & "' "
            Banco.Execute strSQL
            strSQL = ""

            strSQL = "delete from tab_produtos where cod_pro = '" & sCodPro1 & "'"
            Banco.Execute strSQL
            strSQL = ""
            
            strSQL = "delete from tab_composicao where cod_pro = '" & sCodPro1 & "'"
            Banco.Execute strSQL
            
            MsgBox "Produto Excluído!", vbExclamation
            msflxgrd.Rows = 1
            iAux = 0
            dtPro.Refresh
            
        ElseIf OptnCodPro2 Then
            strSQL = "update tab_composicao set "
            strSQL = strSQL & "e1_pro=" & F_GravaDec(Ep1 + Ep2) & ", "
            strSQL = strSQL & "e2_pro=" & F_GravaDec(Ep3 + Ep4) & ", "
            strSQL = strSQL & "r_pro=" & F_GravaDec(Er)
            strSQL = strSQL & " where cod_pro='" & sCodPro2 & "' "
            Banco.Execute strSQL
            
            strSQL = strSQL & "update tab_itcompras set "
            strSQL = strSQL & "cod_pro ='" & sCodPro1 & "'"
            strSQL = strSQL & " where cod_pro='" & sCodPro2 & "' "
            Banco.Execute strSQL
            
            strSQL = strSQL & "update tab_itcmp set "
            strSQL = strSQL & "cod_pro ='" & sCodPro1 & "'"
            strSQL = strSQL & " where cod_pro='" & sCodPro2 & "' "
            Banco.Execute strSQL
            
            strSQL = strSQL & "update tab_itvendas set "
            strSQL = strSQL & "cod_pro ='" & sCodPro1 & "'"
            strSQL = strSQL & " where cod_pro='" & sCodPro2 & "' "
            Banco.Execute strSQL
            
            strSQL = strSQL & "update tab_itsai set "
            strSQL = strSQL & "cod_pro ='" & sCodPro1 & "'"
            strSQL = strSQL & " where cod_pro='" & sCodPro2 & "' "
            Banco.Execute strSQL
            
            strSQL = strSQL & "update tab_itperdas set "
            strSQL = strSQL & "cod_pro ='" & sCodPro1 & "'"
            strSQL = strSQL & " where cod_pro='" & sCodPro2 & "' "
            Banco.Execute strSQL
            
            strSQL = strSQL & "update tab_itcont set "
            strSQL = strSQL & "cod_pro ='" & sCodPro2 & "'"
            strSQL = strSQL & " where cod_pro='" & sCodPro2 & "' "
            Banco.Execute strSQL
            strSQL = ""
            
            strSQL = strSQL & "update tab_itdev set "
            strSQL = strSQL & "cod_pro ='" & sCodPro2 & "'"
            strSQL = strSQL & " where cod_pro='" & sCodPro2 & "' "
            Banco.Execute strSQL
            strSQL = ""
            
            
            strSQL = "delete from tab_produtos where cod_pro = '" & sCodPro2 & "'"
            Banco.Execute strSQL
            
            strSQL = "delete from tab_composicao where cod_pro = '" & sCodPro2 & "'"
            Banco.Execute strSQL
            
            MsgBox "Produto Excluído!", vbExclamation
            msflxgrd.Rows = 1
            iAux = 0
            dtPro.Refresh
        End If
    Else
        MsgBox "Você não tem permissão para excluir produtos!", vbCritical
        Exit Sub
    End If
    PctSelecionaProd.Visible = False
End Sub

Private Sub cmndbtnDesabilita_Click()
    MsgBox "Em desenvolvimento!", vbInformation
End Sub

Private Sub cmndbtnDupl_Click()
    If dbgrd.Row = -1 Then
        MsgBox "Selecione um produto!", vbExclamation
        msflxgrd.Rows = 1
        txtbxExpr.SetFocus
        Exit Sub
    End If
    If msflxgrd.Rows >= 3 Then
        MsgBox "Não é possível adicionar mais que dois produtos ao Grid!", vbExclamation
        Exit Sub
    End If
    msflxgrd.Rows = msflxgrd.Rows + 1
    iAux = iAux + 1
    With dtPro.Recordset
        If iAux = 2 Then
            If sCodPro1 = !cod_pro Then
                MsgBox "Esse produto é igual ao anterior, selecione outro!", vbExclamation
                msflxgrd.Rows = 2
                iAux = 1
                Exit Sub
            End If
        End If
        msflxgrd.Row = iAux
        msflxgrd.Col = 0
        msflxgrd.CellAlignment = flexAlignCenterCenter
        msflxgrd.Text = iAux
        msflxgrd.Col = 1
        msflxgrd.Text = !cod_pro
        msflxgrd.Col = 2
        msflxgrd.Text = !codbarra_pro
        msflxgrd.Col = 3
        msflxgrd.Text = !nome_pro
        msflxgrd.Col = 4
        msflxgrd.Text = !nome_mar
        msflxgrd.Col = 5
        msflxgrd.Text = !nome_cat
        msflxgrd.Col = 6
        msflxgrd.Text = !p2_pro
        msflxgrd.Col = 7
        msflxgrd.Text = !unidade_pro
        msflxgrd.Col = 8
        msflxgrd.Text = !codfab_pro
        msflxgrd.Col = 9
        msflxgrd.Text = !Endereco_pro
        msflxgrd.Col = 10
        msflxgrd.Text = !s_pro
        msflxgrd.Col = 11
        msflxgrd.Text = F_BuscaUmCampo("(e2_pro - r_pro)", "tab_composicao", "cod_pro ='" & !cod_pro & "'")
            
        If msflxgrd.Row = 1 Then
            sCodPro1 = !cod_pro
        ElseIf msflxgrd.Row = 2 Then
            sCodPro2 = !cod_pro
        End If
    End With
    dbgrd.SetFocus
End Sub

Private Sub cmndbtnExcluir_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim E1 As String
    Dim E2 As String
    
    If dbgrd.Row = -1 Then
        MsgBox "Selecione um produto!", vbExclamation
        txtbxExpr.SetFocus
        Exit Sub
    End If
        
    strSQL = "select * from tab_itcompras where cod_pro ='" & lblCodPro.caption & "'"
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    If rPesq.RecordCount >= 1 Then
        MsgBox "Existem notas de entrada com esse produto!", vbExclamation
        Exit Sub
    End If
    
    strSQL = "select * from tab_itcmp where cod_pro ='" & lblCodPro.caption & "'"
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    If rPesq.RecordCount >= 1 Then
        MsgBox "Existem notas de entrada com esse produto!", vbExclamation
        Exit Sub
    End If
    
    strSQL = "select * from tab_itsai where cod_pro ='" & lblCodPro.caption & "'"
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    If rPesq.RecordCount >= 1 Then
        MsgBox "Existem vendas com esse produto!", vbExclamation
        Exit Sub
    End If
    
    strSQL = "select * from tab_itvendas where cod_pro ='" & lblCodPro.caption & "'"
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    If rPesq.RecordCount >= 1 Then
        MsgBox "Existem vendas com esse produto!", vbExclamation
        Exit Sub
    End If
    
    strSQL = "select * from tab_itperdas where cod_pro ='" & lblCodPro.caption & "'"
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    If rPesq.RecordCount >= 1 Then
        MsgBox "Existem entradas com esse produto!", vbExclamation
        Exit Sub
    End If
    
    strSQL = "select * from tab_itcont where cod_pro ='" & lblCodPro.caption & "'"
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    If rPesq.RecordCount >= 1 Then
        MsgBox "Existem contagens com esse produto!", vbExclamation
        Exit Sub
    End If
    
    strSQL = "select * from tab_itdev where cod_pro ='" & lblCodPro.caption & "'"
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    If rPesq.RecordCount >= 1 Then
        MsgBox "Existem devoluções com esse produto!", vbExclamation
        Exit Sub
    End If

    E1 = F_BuscaUmCampo("e1_pro", "tab_composicao", "cod_pro ='" & lblCodPro.caption & "'")
    E2 = F_BuscaUmCampo("e2_pro", "tab_composicao", "cod_pro ='" & lblCodPro.caption & "'")
    If E1 Or E2 <> 0 Then
        MsgBox "Existe estoque para esse produto!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Confirma exclusão do produto: " & _
        Chr(13) & lblCodPro.caption & " - " & _
        F_BuscaUmCampo("nome_pro", "tab_produtos", "cod_pro ='" & lblCodPro.caption & "'") _
        , vbYesNo + vbQuestion) = vbYes Then

        strSQL = "delete from tab_produtos where cod_pro = '" & lblCodPro.caption & "'"
        Banco.Execute strSQL
        
        strSQL = "delete from tab_composicao where cod_pro = '" & lblCodPro.caption & "'"
        Banco.Execute strSQL
        
        MsgBox "Produto Excluído!", vbExclamation
        dtPro.Refresh
    End If
End Sub


Private Sub dbgrd_DblClick()
'    cmndbtnExcluir_Click
End Sub

Private Sub dbgrd_GotFocus()
    Shape2.Visible = True
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    Shape2.Visible = False
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
    txtbxExpr.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    dtPro.Connect = Banco.Connect
    P_CentraForm Me
    P_PreencheCateg cmbxCat
    P_PreencheMarca cmbxMarca
    cmbxBusca.ListIndex = 0
    cmbxOrd.ListIndex = 0
    P_PreencheMsflxgrd
    PctSelecionaProd.Visible = False
    Shape2.Visible = False
End Sub

Private Sub PctSelecionaProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        PctSelecionaProd.Visible = False
    End If
End Sub


Private Sub Picture1_Click()
    PctSelecionaProd.Visible = False
End Sub

Private Sub txtbxExpr_Change()
    busca
End Sub

Sub busca()
    Dim strSQL As String
    Dim sOrder As String
    Dim sCampo As String
    If cmbxBusca.ListIndex = 1 Then
        sCampo = "tp.cod_pro"
    ElseIf cmbxBusca.ListIndex = 2 Then
        sCampo = "tp.codfab_pro"
    ElseIf cmbxBusca.ListIndex = 3 Then
        sCampo = "te.nome_mar"
    ElseIf cmbxBusca.ListIndex = 4 Then
        sCampo = "tg.nome_cat"
    Else
        sCampo = "tp.nome_Pro"
    End If
    
    If cmbxOrd.ListIndex = 1 Then
        sOrder = "tp.cod_pro,tp.nome_pro"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrder = "tp.codfab_pro,tp.nome_pro,tp.cod_pro"
    Else
        sOrder = "tp.nome_pro,tp.cod_pro "
    End If
    strSQL = "select tg.nome_cat, te.nome_mar, tq.*, "
    strSQL = strSQL & "tp.nome_pro,tp.unidade_pro,tp.endereco_pro, tp.codbarra_pro,"
    strSQL = strSQL & "tp.codfab_pro, tp.lucro_pro,"
    strSQL = strSQL & "(tq.e2_pro-tq.r_pro) as estoque"
    strSQL = strSQL & " from tab_produtos as tp,tab_categorias as tg, "
    strSQL = strSQL & " tab_composicao as tq,tab_marcas as te  "
    strSQL = strSQL & " where " & sCampo & " like '"
    strSQL = strSQL & txtbxExpr.Text & "%' "
    strSQL = strSQL & " and tq.cod_pro=tp.cod_pro "
    strSQL = strSQL & " and tg.cod_cat=tp.cod_cat and te.cod_mar=tp.cod_mar "
    If cmbxCat.ListIndex <> -1 Then
        strSQL = strSQL & " and tp.cod_cat='" & lblCat.caption & "'  "
    End If
    If cmbxMarca.ListIndex <> -1 Then
        strSQL = strSQL & " and tp.cod_mar='" & lblMar.caption & "'  "
    End If
    strSQL = strSQL & " order by " & sOrder
    dtPro.RecordSource = strSQL
    dtPro.Refresh
    lblList.caption = "[Listados..: " & dtPro.Recordset.RecordCount & _
        "   Cadastrados: " & F_BuscaUmCampo("count(*)", "tab_produtos", "") & "]"
End Sub

Private Sub txtbxExpr_GotFocus()
    P_SelecTexto txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        busca
        If dtPro.Recordset.RecordCount > 0 Then
            dbgrd.SetFocus
        End If
    ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dbgrd.SetFocus
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    P_AjustaTexto txtbxExpr
End Sub

Private Sub cmbxCat_Click()
    lblCat.caption = F_BuscaUmCampo("Cod_Cat", "tab_categorias", "nome_cat='" & cmbxCat.Text & "'")
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
    lblMar.caption = F_BuscaUmCampo("Cod_mar", "tab_marcas", "nome_mar='" & cmbxMarca.Text & "'")
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

Private Sub cmbxSub_Click()
    lblSub.caption = F_BuscaUmCampo("Cod_sub", "tab_subcat", "nome_sub='" & cmbxSub.Text & "'")
    busca
End Sub

Private Sub cmbxSub_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub dtPro_Reposition()
    On Error Resume Next
    With dtPro.Recordset
        If .RecordCount > 0 Then
            lblCodPro.caption = !cod_pro
        Else
            
        End If
    End With
End Sub

Sub P_PreencheMsflxgrd()
    msflxgrd.Rows = 1
    msflxgrd.Cols = 12
    msflxgrd.Col = 0
    msflxgrd.CellAlignment = flexAlignCenterCenter
    msflxgrd.Text = "Itens"
    msflxgrd.Col = 1
    msflxgrd.Text = "Código"
    msflxgrd.Col = 2
    msflxgrd.Text = "Cód. Barra"
    msflxgrd.Col = 3
    msflxgrd.Text = "Nome do Produto"
    msflxgrd.Col = 4
    msflxgrd.Text = "Marca"
    msflxgrd.Col = 5
    msflxgrd.Text = "Categoria"
    msflxgrd.Col = 6
    msflxgrd.Text = "Preço"
    msflxgrd.Col = 7
    msflxgrd.Text = "UN"
    msflxgrd.Col = 8
    msflxgrd.Text = "Cód. Fáb."
    msflxgrd.Col = 9
    msflxgrd.Text = "End."
    msflxgrd.Col = 10
    msflxgrd.Text = "S"
    msflxgrd.Col = 11
    msflxgrd.Text = "Estoque"
    
    msflxgrd.ColWidth(0) = 500
    msflxgrd.ColWidth(2) = 1500
    msflxgrd.ColWidth(3) = 4000
    msflxgrd.ColWidth(4) = 1200
    msflxgrd.ColWidth(5) = 1500
    msflxgrd.ColWidth(6) = 1500
    msflxgrd.ColWidth(7) = 400
    msflxgrd.ColWidth(8) = 1500
    msflxgrd.ColWidth(9) = 600
    msflxgrd.ColWidth(10) = 300
    msflxgrd.ColWidth(11) = 800
    
End Sub
