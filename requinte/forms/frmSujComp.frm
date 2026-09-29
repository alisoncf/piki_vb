VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmSujComp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Processo Eletrônico de Compras"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10515
   Icon            =   "frmSujComp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      Caption         =   "Incluir Produto &Não Sujerido:"
      Height          =   615
      Left            =   6720
      TabIndex        =   32
      Top             =   0
      Width           =   3735
      Begin VB.TextBox txtbxInc_Mer 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   34
         ToolTipText     =   "Informe o nome da mercadoria seguido de <ENTER> para localizá-la."
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton cmndbtnInc_Mer 
         Caption         =   "..."
         Height          =   285
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblMarca_mer 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Marca"
         ForeColor       =   &H8000000C&
         Height          =   195
         Left            =   1560
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Fornecedor:"
      Height          =   615
      Left            =   2040
      TabIndex        =   30
      Top             =   0
      Width           =   4575
      Begin VB.ComboBox cmbxFor 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame frPeriod 
      Caption         =   "Periodicidade:(Dias)"
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   1815
      Begin VB.TextBox txtbxTotalDias 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   26
         Text            =   "10"
         ToolTipText     =   "Tempo de duração desse processo..."
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtbxSustent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   1200
         TabIndex        =   11
         Text            =   "7"
         ToolTipText     =   "Refere-se ao período de sustentação(Tempo em que o estoque deve durar)"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtbxEntrega 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Text            =   "2"
         ToolTipText     =   "Tempo previsto para entrega"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtbxDuracao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Text            =   "1"
         ToolTipText     =   "Tempo de duração desse processo..."
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total de Dias:"
         Height          =   195
         Left            =   165
         TabIndex        =   27
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Suprimento"
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Prazo Entrega:"
         Height          =   195
         Left            =   105
         TabIndex        =   13
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Duração:"
         Height          =   195
         Left            =   495
         TabIndex        =   12
         Top             =   1080
         Width           =   660
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1815
      Begin VB.CommandButton cmndbtnFecha 
         Cancel          =   -1  'True
         Caption         =   "&Fechar Formulário"
         Height          =   615
         Left            =   120
         Picture         =   "frmSujComp.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fecha formulário de Compra"
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnImprList 
         Caption         =   "Imprimir Processo"
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         Picture         =   "frmSujComp.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnCancela 
         Caption         =   "&Cancelar Processo"
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         Picture         =   "frmSujComp.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancelar Processo de Compra"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnFinaliza 
         Caption         =   "&Finalizar Processo"
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         Picture         =   "frmSujComp.frx":06E8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Finalizar Processo de Compra"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnInicia 
         Caption         =   "&Iniciar Processo"
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         Picture         =   "frmSujComp.frx":0832
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Iniciar Processo de Compra"
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnList 
         Caption         =   "&Buscar Produtos"
         Height          =   615
         Left            =   120
         Picture         =   "frmSujComp.frx":0DBC
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Listar Mercadorias"
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar Itens"
         Height          =   615
         Left            =   120
         Picture         =   "frmSujComp.frx":1766
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin MSDBGrid.DBGrid dbgrdCompras 
      Bindings        =   "frmSujComp.frx":1AF0
      Height          =   5295
      Left            =   2040
      OleObjectBlob   =   "frmSujComp.frx":1B08
      TabIndex        =   28
      Top             =   720
      Width           =   8415
   End
   Begin VB.Frame Frame4 
      Caption         =   "Ordem:"
      Height          =   735
      Left            =   2040
      TabIndex        =   22
      Top             =   6120
      Width           =   3615
      Begin VB.OptionButton optnrOrderBy 
         Caption         =   "Código"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optnrOrderBy 
         Caption         =   "Descrição"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optnrOrderBy 
         Caption         =   "&Quantidade"
         Height          =   195
         Index           =   2
         Left            =   2280
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label lblCompra 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "000000"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Itens:"
      Height          =   735
      Left            =   5760
      TabIndex        =   15
      Top             =   6120
      Width           =   4695
      Begin VB.Label lblSelec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   1200
         TabIndex        =   21
         Top             =   360
         Width           =   90
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Selecionados:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Listados:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   1920
         TabIndex        =   19
         Top             =   360
         Width           =   660
      End
      Begin VB.Label lblTotalPrev 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   3840
         TabIndex        =   18
         Top             =   360
         Width           =   315
      End
      Begin VB.Label lblListados 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   2640
         TabIndex        =   17
         Top             =   360
         Width           =   90
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   3360
         TabIndex        =   16
         Top             =   360
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmSujComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Dim WithEvents adocompras As ADODB.Recordset
Attribute adocompras.VB_VarHelpID = -1
Private Sub cmbxCat_Click()
    cmbxCat.tag = F_BuscaUmCampo("cod_cat", "tab_categorias", "nome_cat='" & cmbxCat.text & "'")
End Sub

Private Sub cmbxLoc_Ipc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
    
End Sub

Private Sub cmbxMarca_Click()
    cmbxMarca.tag = F_BuscaUmCampo("Cod_mar", "tab_marcas", "nome_mar='" & cmbxMarca.text & "'")
    P_SalvaComp
    P_BuscaItensCompra lblCompra.caption
End Sub

Private Sub cmbxMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxMarca.ListIndex = -1
    End If
End Sub


Private Sub cmndbtnAddFor_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    If dtForn.Recordset.RecordCount = 0 Then
'        MsgBox "Não há fornecedores listados!", vbCritical
        Exit Sub
    End If
    With dtForn.Recordset
        strSQL = "update tab_fornpcp set sts_fpc='S' "
        strSQL = strSQL & " where cod_ter = '" & !cod_For & "' "
        strSQL = strSQL & " and seq_pcp ='" & lblCompra.caption & "' "
        Conn.Execute strSQL
    End With
    P_buscaForn
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmbxFor_Click()
    cmbxFor.tag = F_BuscaUmCampo("cod_for", "tab_fornecedores", "nome_For='" & cmbxFor.text & "'")
    carrega
End Sub

Private Sub cmndbtnBuscar_Click()
    P_BuscaItensCompra lblCompra.caption
End Sub



Private Sub cmndbtnCancela_Click()
    On Error GoTo trata_Erro
    Dim strSQL As String
    If MsgBox("Deseja realmente excluir o Processo de Compra?", vbYesNo + vbCritical) = vbYes Then
        strSQL = "Delete from tab_procomp where seq_pcp='" & lblCompra.caption & "' "
        Conn.Execute strSQL
        strSQL = "Delete from tab_itenspcp where seq_pcp='" & lblCompra.caption & "' "
        Conn.Execute strSQL
        MsgBox "Processo Cancelado!", vbCritical
        Unload Me
        
    End If
        
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
    
End Sub

Private Sub cmndbtnFecha_Click()
    Unload Me
End Sub

Private Sub cmndbtnFinaliza_Click()
    Dim strSQL As String
    If MsgBox("Confirma Finalização do Processo de Compra: " & lblCompra.caption & " ?", vbYesNo + vbQuestion) = vbYes Then
        strSQL = F_BuscaUmCampo("count(cod_pro)", "TAB_ITENSPCP", "STS_IPC='S' and SEQ_PCP='" & lblCompra.caption & "'")
        If strSQL = "0" Then
            MsgBox "Nenhum item foi selecionado!", vbCritical
            Exit Sub
        End If
        strSQL = F_BuscaUmCampo("count(COD_ter)", "TAB_FORNPCP", " SEQ_PCP='" & lblCompra.caption & "'")
        If strSQL = "0" Then
            strSQL = "Não foi informado nenhum fornecedor!"
            strSQL = strSQL & Chr(13) & "Deseja Continuar mesmo assim?"
            If MsgBox(strSQL, vbYesNo + vbCritical) = vbNo Then
                Exit Sub
            End If
            
        End If
        strSQL = "Update tab_procomp set sts_pcp= '2',COD_usu='" & strCod_Usu & "' "
        strSQL = strSQL & " where seq_pcp='" & lblCompra.caption & "' "
        
        Conn.Execute strSQL
        strSQL = "Update tab_itenspcp set sts_ipc = 'C' where sts_ipc = 'N' and seq_pcp='" & lblCompra.caption & "' "
        Conn.Execute strSQL  'passa de N -não pra C -cancelado
        MsgBox "Processo Finalizado!", vbInformation
        Unload Me
    End If
End Sub

Private Sub cmndbtnForn_Click()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    strSQL = "select "
    strSQL = strSQL & " distinct tt.*"
    strSQL = strSQL & " From"
    strSQL = strSQL & " tab_fornecedores as tt,"
    strSQL = strSQL & " tab_itcmp as ti,"
    strSQL = strSQL & " tab_cmp as te,"
    strSQL = strSQL & " tab_itenspcp As tc"
    strSQL = strSQL & " Where"
    strSQL = strSQL & " ti.seq_not=te.seq_not "
    strSQL = strSQL & " and ti.cod_pro = tc.cod_pro"
    
    strSQL = strSQL & " and te.cod_for = tt.cod_for"
    
    strSQL = strSQL & " and seq_pcp = '" & lblCompra.caption & "'"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        strSQL = "Delete from tab_fornpcp where seq_pcp = '" & lblCompra.caption & "' "
        Conn.Execute strSQL
        Do While Not .EOF
            strSQL = "insert into tab_fornpcp(Cod_Ter,Seq_Pcp,Sts_Fpc)Values("
            strSQL = strSQL & "'" & !cod_For & "', "
            strSQL = strSQL & "'" & lblCompra.caption & "',"
            strSQL = strSQL & "'N')"
            Conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
    P_buscaForn
    'stbCompras.Tab = 1
End Sub
Sub P_buscaForn()
'    strSQL = "select tt.* from tab_fornecedores as tt, tab_fornpcp as tf "
'    strSQL = strSQL & " where tt.cod_for=tf.cod_ter"
'    strSQL = strSQL & " and tf.seq_pcp='" & lblCompra.caption & "' "
'    strSQL = strSQL & " and tf.sts_fpc='N' "
'    strSQL = strSQL & " Order by tt.nome_for "
'    dtForn.RecordSource = strSQL
'    dtForn.Refresh
'    If dtForn.Recordset.RecordCount = 0 Then
'        MsgBox "Não há fornecedores a serem listados!", vbCritical
'    Else
'       cmndbtnFinaliza.Enabled = True
'    End If
'    strSQL = "select tt.* from tab_fornecedores as tt, tab_fornpcp as tf "
'    strSQL = strSQL & " where tt.cod_for=tf.cod_ter "
'    strSQL = strSQL & " and tf.seq_pcp='" & lblCompra.caption & "' "
'    strSQL = strSQL & " and tf.sts_fpc='S' "
'    strSQL = strSQL & " Order by tt.nome_for "
'    dtSelect.RecordSource = strSQL
'    dtSelect.Refresh
End Sub

Private Sub cmndbtnImportar_Click()
    
    
End Sub

Private Sub cmndbtnImprCot_Click()
    Dim iQtd_pag As Integer
    P_ImprPcp lblCompra.caption, frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        P_ImprPcp lblCompra.caption, Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
    
    Exit Sub
    
    
    
    
End Sub

Private Sub cmndbtnImprList_Click()
    imprimeListagem frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprimeListagem Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub imprimeListagem(obj As Object)
    If dtCompras.Recordset.RecordCount = 0 Then
        MsgBox "Não há mercadorias para serem listadas!", vbExclamation
        Exit Sub
    End If
    
    P_CabRelTimbr "Listagem de Mercadorias a Adquirir", obj
    With dtCompras.Recordset
        .MoveFirst
        P_ImprX obj, "Código", 0
        P_ImprX obj, "Marca", 20
        P_ImprX obj, "End", 60
        P_ImprX obj, "Descrição", 70
        'P_ImprX "L", 175
        P_ImprX obj, "A Comprar", 190, False, True
        Do While Not .EOF
            If Printer.CurrentY >= Printer.ScaleHeight - 30 Then
                If TypeName(obj) = "Printer" Then
                    P_Rodape strEmp_Nome & "-" & strEmp_Cidade & "/" & strEmp_UF
                End If
            End If
            P_ImprX obj, !cod_pro, 0
            P_ImprX obj, !nome_mar, 20
            P_ImprX obj, !endereco_pro, 60
            P_ImprX obj, !nome_pro, 70
            'P_ImprX obj, !loc_ipc, 175
            P_ImprX obj, FormatVal(!qfa_ipc), 190, False, True
            
            .MoveNext
        Loop
    End With
    
    If TypeName(obj) = "Printer" Then
        P_FimRel obj, strEmp_Nome & "-" & strEmp_Cidade & "/" & strEmp_UF
        Printer.EndDoc
    End If
End Sub
Private Sub cmndbtnInc_Mer_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    If txtbxInc_Mer.tag = "" Then
        txtbxInc_Mer_DblClick
    End If
    If MsgBox("Confirma inclusão da mercadoria: " & Trim(txtbxInc_Mer.text) & " ?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    strSQL = "Insert into tab_itenspcp(seq_pcp,cod_pro,qfa_ipc,"
    strSQL = strSQL & "sts_ipc,upc_ipc)Values("
    strSQL = strSQL & "'" & lblCompra.caption & "',"
    strSQL = strSQL & "'" & txtbxInc_Mer.tag & "',"
    strSQL = strSQL & "0,"
    strSQL = strSQL & "'N'," 'não selecionado
    strSQL = strSQL & str(F_UltPreco_Mer(txtbxInc_Mer.tag, lblMarca_mer.tag)) & ")"
    
    Conn.Execute strSQL
    txtbxInc_Mer.text = ""
    txtbxInc_Mer.tag = ""
    lblMarca_mer.caption = ""
    lblMarca_mer.tag = ""
    P_BuscaItensCompra lblCompra.caption
Exit Sub

trata_Erro:
    MsgBox Err.Description, vbCritical
    
End Sub

Private Sub cmndbtnInicia_Click()
     If MsgBox("Confirma inicialização do Processo de Compras?", vbYesNo + vbQuestion) = vbYes Then
        P_InsereProc
        cmndbtnInicia.Enabled = False
        cmndbtnList.Visible = False
        cmndbtnBuscar.Visible = True
        cmndbtnList.Visible = False
        cmndbtnFinaliza.Enabled = True
        txtbxInc_Mer.Enabled = True
        cmndbtnInc_Mer.Enabled = True
        cmndbtnImprList.Enabled = True
        cmndbtnCancela.Enabled = True
        'cmndbtnImportar.Enabled = True
        'cmndbtnExport.Enabled = True
     End If
End Sub

Private Sub cmndbtnList_Click()
    Dim strSQL As String
    Dim iNDC As Integer 'numero de dias para consumir
    
    If cmbxFor.tag = "" Then
 '       MsgBox "Informe primeiro o fornecedor!", vbCritical
'        cmbxFor.SetFocus
  '      Exit Sub
    End If
    
    
    iNDC = CInt(txtbxDuracao.text) + Int(txtbxEntrega.text) + Int(txtbxSustent.text)
    frmMediaConsumo_Mer.cmbxForn.ListIndex = cmbxFor.ListIndex
    frmMediaConsumo_Mer.Show 1
    strSQL = "Select tm.*,tc.nome_cat,tm.mmd_pro,"
    strSQL = strSQL & " ((tm.MMD_PRO *" & iNDC & ")-tm.er_pro) as qtd from "
    strSQL = strSQL & " tab_produtos as tm,"
    strSQL = strSQL & "  tab_categorias as tc "
    strSQL = strSQL & " where tc.cod_cat=tm.cod_cat "
    'If cmbxCat.ListIndex <> -1 Then
    '    strSQL = strSQL & " and tm.cod_cat='" & cmbxCat.Tag & "' "
    'End If
    strSQL = strSQL & " and tm.er_pro-tm.er_pro-(tm.MMD_PRO *" & iNDC & ") < 0"
    If MsgBox("Deseja listar os itens em que o estoque seja maior que a sugestão?", _
        vbYesNo + vbQuestion) = vbNo Then
        strSQL = strSQL & " having qtd > 0 "
    End If
    dtCompras.RecordSource = strSQL
    dtCompras.Refresh
    
    If dtCompras.Recordset.RecordCount > 0 Then
        cmndbtnInicia.Enabled = True
    Else
        MsgBox "Não foram encontrados itens com condições de compra!" & Chr(13) & _
                "Reinicie o processo de cálculo de média das mercadorias!", vbExclamation
        cmndbtnInicia.Enabled = False
    End If
End Sub



Private Sub cmndbtnRemFor_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    If dtSelect.Recordset.RecordCount = 0 Then
        MsgBox "Não há fornecedores Selecionados!", vbCritical
        Exit Sub
    End If
    With dtSelect.Recordset
        strSQL = "update tab_fornpcp set sts_fpc='N' "
        strSQL = strSQL & " where cod_ter = '" & !cod_ter & "' "
        strSQL = strSQL & " and seq_pcp ='" & lblCompra.caption & "' "
        Conn.Execute strSQL
    End With
    P_buscaForn
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub



Private Sub dbgrdCompras_DblClick()
    If cmndbtnList.Visible = True Then
        MsgBox "O processo ainda não foi inicializado!", vbCritical
        Exit Sub
    End If
    frmAltProPCP.cmbxForn.ListIndex = cmbxFor.ListIndex
    frmAltProPCP.frAltMer.Visible = True
    With dtCompras.Recordset
        frmAltProPCP.caption = "Incluir/Alterar " & !nome_pro & " - Processo: " & Me.tag
        frmAltProPCP.tag = Me.tag
        frmAltProPCP.lblNome_Mer.tag = !cod_pro
        frmAltProPCP.txtbxQFA_Ipc.text = FormatVal(!qfa_ipc)
        frmAltProPCP.txtbxTotal.text = FormatVal(!total)
        frmAltProPCP.lblNome_Mer.caption = !nome_pro
        
        frmAltProPCP.Show 1
        buscaFor
    End With
    Unload frmAltProPCP
    RefreshDataPosition adocompras
    soma
End Sub
Sub soma()
    
    Dim cTotal As Currency
    Dim iLocal As Integer
    Dim iItem As Integer
    With dtCompras.Recordset
        iLocal = .AbsolutePosition
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        Do While Not .EOF
            
            If !sts_ipc = "S" Then
                cTotal = cTotal + !total
                iItem = iItem + 1
            End If
            .MoveNext
        Loop
        If .RecordCount > 0 Then
            .MoveFirst
            .Move iLocal
        End If
        lblTotalPrev.caption = FormatVal(cTotal)
        lblListados.caption = .RecordCount
        lblSelec.caption = iItem
    End With
End Sub
Private Sub dbgrdCompras_GotFocus()
    
    dbgrdCompras.MarqueeStyle = 3
End Sub

Sub P_SalvaComp()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "Update tab_procomp set cod_cat ='" & cmbxCat.tag & "', "
    
    strSQL = strSQL & " cod_mar='" & cmbxMarca.tag & "', "
    strSQL = strSQL & " nome_pro='" & txtbxNome_Mer.text & "' "
    strSQL = strSQL & " where seq_pcp='" & lblCompra.caption & "'"
    Conn.Execute strSQL
    strSQL = "Update tab_ItensPcp set sts_ipc='N' cod_cat <> '"
    strSQL = strSQL & cmbxCat.tag & "' "
    
    strSQL = strSQL & " and cod_mar <> '" & cmbxMarca.tag & "' "
    'conn.execute strSQL rever isto
    
End Sub

Private Sub dbgrdCompras_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dbgrdCompras_DblClick
    End If
End Sub



Private Sub dbgrdCompras_LostFocus()
    dbgrdCompras.MarqueeStyle = 0
End Sub

Private Sub dtCompras_Reposition()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
   ' On Error GoTo trata_erro
    With dtCompras.Recordset
'        verificaProduto !cod_pro
    End With
    
trata_Erro:
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        carrega
    End If
    bActive = True
End Sub
Sub verificaProduto(cod_pro As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select ti.*,tp.sts_pcp,tp.dti_pcp from tab_itenspcp as ti,tab_procomp as tp "
    strSQL = strSQL & " where tp.seq_pcp=ti.seq_pcp "
    strSQL = strSQL & " and cod_pro='" & cod_pro & "' "
    strSQL = strSQL & " and sts_ipc='S'"
    strSQL = strSQL & " and sts_pcp<3 "
    strSQL = strSQL & " and tp.seq_pcp<>'" & Me.tag & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "Este produto encontra-se marcado como já pedido " & Chr(13)
            strSQL = strSQL & "Processo Nº: " & !seq_pcp & " de: " & !dti_pcp & Chr(13)
            strSQL = strSQL & "Quantidade: " & !qfa_ipc & " Situacao do Processo: " & !sts_pcp
            MsgBox strSQL, vbExclamation
            .MoveNext
        Loop
        .Close
    End With
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Ajuda
    End If
End Sub
Sub Ajuda()
    Dim strSQL As String
    strSQL = "Sujestão de Compra é o processo que tem o objetivo de ajudar o Empresário a "
    strSQL = strSQL & "definir quais os produtos devem ser comprados para que se tenha "
    strSQL = strSQL & "um melhor giro, e conseqüentemente mais lucratividade."
    strSQL = strSQL & Chr(13) & Chr(13)
    strSQL = strSQL & "É composto de três etapas: " & Chr(13)
    strSQL = strSQL & " 1 - cálculo do MMD(Média de Movimentação Diária)" & Chr(13)
    strSQL = strSQL & " 2 - preenchimento periodicidade" & Chr(13)
    strSQL = strSQL & " 3 - filtragem dos produtos " & Chr(13)
    strSQL = strSQL & " 4 - seleção e quantização dos itens a serem comprados" & Chr(13)
    strSQL = strSQL & " 5 - impressão da cotação/processo" & Chr(13)
    strSQL = strSQL & " 6 - finalização " & Chr(13)
    strSQL = strSQL & " 7 - entrada de produtos " & Chr(13)
    MsgBox strSQL, vbInformation, "PCIP - Processo de Compra de Produtos Informatizada"
End Sub
Private Sub Form_Load()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Me.top = 10
    Me.Left = 10
    bActive = False
    dtCompras.Connect = Conn.ConnectionString

    
        
    
    
    

End Sub
Sub buscaFor()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select nome_for from tab_fornecedores as tf, tab_itenspcp as ti "
    strSQL = strSQL & " where ti.cod_for=tf.cod_for "
    strSQL = strSQL & " and ti.seq_pcp='" & Me.tag & "' "
    strSQL = strSQL & " order by nome_for "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
    cmbxFor.Clear
        cmbxFor.AddItem "(Todos)"
        Do While Not .EOF
            cmbxFor.AddItem !nome_for
            .MoveNext
        Loop
        .Close
    End With
    
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select tp.*  from tab_procomp as tp "
    
    strSQL = strSQL & " where seq_pcp='" & Me.tag & "'  "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            P_Status "Buscando Compra em aberto..."
            Me.tag = !seq_pcp
            Me.caption = "PCPI - Processo  de Compra Informatizada de Produtos - Processo Nº: " & Me.tag
            lblCompra.caption = !seq_pcp
            txtbxSustent.text = !pse_pcp
            txtbxEntrega.text = !pec_pcp
            txtbxDuracao.text = !ppc_pcp
           ' P_LocalizaOperacao !nome_for, cmbxFor
            P_BuscaItensCompra !seq_pcp
            
            cmndbtnList.Visible = False
            cmndbtnBuscar.Visible = True
            If dtCompras.Recordset.RecordCount > 0 Then
                cmndbtnInicia.Enabled = False
            End If
            cmndbtnFinaliza.Enabled = True
            txtbxInc_Mer.Enabled = True
            cmndbtnInc_Mer.Enabled = True
            cmndbtnImprList.Enabled = True
            cmndbtnCancela.Enabled = True
        Else
            cmndbtnFinaliza.Enabled = False
            txtbxInc_Mer.Enabled = False
            cmndbtnInc_Mer.Enabled = False
            .Close
        End If
    End With
    buscaFor
    P_Status ""
End Sub
Sub P_InsereProc()
    Dim strSQL As String
    
    Dim sSeq_Pcp As String
    sSeq_Pcp = F_BuscaUmCampo("max(SEQ_PCP)", "TAB_PROCOMP", "")
    If sSeq_Pcp = "" Then
        sSeq_Pcp = "000001"
    Else
        sSeq_Pcp = Format(sSeq_Pcp + 1, "0#####")
    End If
    strSQL = "insert tab_procomp(SEQ_PCP,DTI_PCP,DTF_PCP,STS_PCP,"
    strSQL = strSQL & "PSE_PCP,PEC_PCP,PPC_PCP,cod_For)Values("
    strSQL = strSQL & "'" & sSeq_Pcp & "',"
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & F_GravaData(DateAdd("D", CDbl(txtbxDuracao.text), Date)) & "," ' data prevista para término
    strSQL = strSQL & "'1',"
    strSQL = strSQL & str(txtbxSustent.text) & ","
    strSQL = strSQL & str(txtbxEntrega.text) & ","
    strSQL = strSQL & str(txtbxDuracao.text) & ", "
    strSQL = strSQL & "'" & cmbxFor.tag & "')"
    Conn.Execute strSQL
    P_InsereItens sSeq_Pcp
    lblCompra.caption = sSeq_Pcp
    Me.tag = "PCPI - Processo  de Compra Informatizada de Produtos - Processo Nº" & sSeq_Pcp
    Me.tag = sSeq_Pcp
End Sub

Sub P_InsereItens(sSeq_Pcp As String)
    Dim strSQL As String
    With dtCompras.Recordset
        
        Do While Not .EOF
            strSQL = "Insert into tab_itenspcp(seq_pcp,cod_pro,qfa_ipc,"
            strSQL = strSQL & "sts_ipc,suj_ipc,upc_ipc)Values("
            strSQL = strSQL & "'" & sSeq_Pcp & "',"
            strSQL = strSQL & "'" & !cod_pro & "',"
            strSQL = strSQL & str(Round(!qtd)) & ","
            strSQL = strSQL & "'N'," 'não selecionado
            strSQL = strSQL & F_GravaDec(!qtd) & ","
            strSQL = strSQL & str(F_UltPreco_Mer(!cod_pro, "")) & ")"

            Conn.Execute strSQL
            
            P_Status "Inserindo Itens de Compra: " & Round((.AbsolutePosition + 1) * 100 / .RecordCount, 2) & " %"
            .MoveNext
        Loop
        P_BuscaItensCompra sSeq_Pcp
    End With
    P_Status ""
    
End Sub
Sub P_BuscaItensCompra(sSeq_Pcp As String)
        Dim strSQL As String
        Dim cTotal As Currency
        Dim cSelec As Currency
        'On Error GoTo trata_erro
        strSQL = "select tp.nome_pro,tp.mmd_pro,tm.nome_mar,tp.endereco_pro,tp.pv_pro,tt.nome_cat,"
        strSQL = strSQL & " tp.er_pro,ti.*,ti.suj_ipc as qtd,"
        strSQL = strSQL & "(ti.upc_ipc*ti.qfa_ipc) as total "
        strSQL = strSQL & " from tab_produtos as tp, tab_marcas as tm, "
        strSQL = strSQL & " tab_categorias as tt,  "
        strSQL = strSQL & " tab_itenspcp as ti where "
        strSQL = strSQL & " tp.cod_cat=tt.cod_cat "
        strSQL = strSQL & " and tp.cod_mar=tm.cod_mar "
        strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
        
        If cmbxFor.tag <> "" Then
            strSQL = strSQL & " and ti.cod_for='" & cmbxFor.tag & "' "
        End If
        
        strSQL = strSQL & " and ti.seq_pcp = '" & sSeq_Pcp & "' "
        'If cmbxCat.ListIndex > -1 Then
        '    strSQL = strSQL & " And tp.cod_cat='" & cmbxCat.Tag & "' "
        'End If
        
        'If cmbxMarca.ListIndex > -1 Then
        '    strSQL = strSQL & " And tp.cod_mar='" & cmbxMarca.Tag & "' "
        'End If
        'strSQL = strSQL & " and tp.nome_pro like '" & txtbxNome_Mer.Text & "%' "
        strSQL = strSQL & " order by "
        
        If optnrOrderBy(0).Value = True Then
            strSQL = strSQL & " tp.Cod_pro"
        ElseIf optnrOrderBy(2).Value = True Then
            strSQL = strSQL & "suj_ipc desc,tp.nome_pro"
        Else
            strSQL = strSQL & "tp.nome_pro,tp.Cod_pro"
        End If
        dtCompras.RecordSource = strSQL
        dtCompras.Refresh
        cTotal = 0
        With dtCompras.Recordset
            Do While Not .EOF
                If !sts_ipc = "S" Then
                    cTotal = cTotal + !total
                    cSelec = cSelec + 1
                End If
                
                .MoveNext
            Loop
            lblTotalPrev.caption = FormatVal(cTotal)
            lblListados.caption = .RecordCount
            lblSelec.caption = cSelec
            If .RecordCount > 0 Then
                .MoveFirst
            End If
        End With
trata_Erro:
End Sub
Function F_UltPreco_Mer(sCod_Mer As String, sMarca_Mer As String) As Currency
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sData As String
    Dim cValor As Currency
    'On Error GoTo trata_erro
    strSQL = "Select max(te.entrada_not) as maximo from tab_cmp as te, "
    strSQL = strSQL & "tab_itcmp as ti where ti.seq_not=te.seq_not and "

    strSQL = strSQL & " ti.cod_pro='" & sCod_Mer & "' "
    
    abreObjetoAdo rPesq, strSQL
    If Not IsNull(rPesq!maximo) Then
        sData = rPesq!maximo
        strSQL = "Select max(ti.unit_ite) as maximo "
        strSQL = strSQL & " from tab_cmp as te, "
        strSQL = strSQL & " tab_itcmp as ti where ti.Seq_not=te.Seq_not and "
        
        strSQL = strSQL & " ti.cod_pro='" & sCod_Mer & "' "
        
        strSQL = strSQL & " and te.entrada_not=" & F_PesqData(sData)
        abreObjetoAdo rPesq, strSQL
        F_UltPreco_Mer = rPesq!maximo
    Else
        F_UltPreco_Mer = 0
    End If
    
    
    
    rPesq.Close
    Exit Function
trata_Erro:

    F_UltPreco_Mer = 0
End Function

Private Sub frAltMer_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If State <> 1 Then
    'frAltMer.Top = Y
    'frAltMer.Left = X
    End If
End Sub

Private Sub Label12_Click()

End Sub

Private Sub optnrOrderBy_Click(Index As Integer)
    If lblCompra.caption <> "" Then
        P_Status "Buscando itens..."
        P_BuscaItensCompra lblCompra.caption
        P_Status ""
    End If
End Sub

Private Sub stbCompras_Click(PreviousTab As Integer)
'    If stbCompras.Tab = 0 Then
'        cmndbtnImprCot.Visible = False
'        frPeriod.Visible = True
'    Else
'        cmndbtnImprCot.Visible = True
'        frPeriod.Visible = False
'    End If
End Sub

Private Sub stbCompras_KeyDown(KeyCode As Integer, Shift As Integer)
    frAltMer.Visible = False
    
End Sub

Private Sub txtbxDuracao_GotFocus()
    P_SelecTexto txtbxDuracao
End Sub

Private Sub txtbxDuracao_LostFocus()
        P_ValorLostFocus txtbxDuracao, "#0"
        txtbxTotalDias.text = Int(txtbxSustent.text) + Int(txtbxDuracao.text) + Int(txtbxEntrega.text)
End Sub

Private Sub txtbxEntrega_GotFocus()
    P_SelecTexto txtbxEntrega
End Sub

Private Sub txtbxEntrega_LostFocus()
    P_ValorLostFocus txtbxEntrega, "#0"
    txtbxTotalDias.text = Int(txtbxSustent.text) + Int(txtbxDuracao.text) + Int(txtbxEntrega.text)
End Sub

Private Sub txtbxInc_Mer_Change()
    txtbxInc_Mer.tag = ""
    lblMarca_mer.caption = ""
End Sub

Private Sub txtbxInc_Mer_DblClick()
    Dim p As Produto
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    If txtbxInc_Mer.text = "" Then
        MsgBox "Informe um nome para o produto que deseja incluir!", vbExclamation
        Exit Sub
        Exit Sub
    End If
    p = F_BuscaProdSaida(txtbxInc_Mer.text)
    txtbxInc_Mer.text = p.nome_pro
    'txtbxInc_Mer.tag = p.cod_pro
    'lblMarca_mer.tag = p.cod_mar
    'lblMarca_mer.caption = p.nome_mar
    strSQL = "select * from tab_itenspcp where seq_pcp ='" & lblCompra.caption & "' "
    strSQL = strSQL & " and cod_pro = '" & p.cod_pro & "' "
    
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount > 0 Then
        MsgBox "Este Produto já está relacionado nessa compra!", vbExclamation
        rPesq.Close
        txtbxInc_Mer.tag = ""
        txtbxInc_Mer.SetFocus
    Else
        rPesq.Close
    End If
    cmndbtnInc_Mer_Click
End Sub

Private Sub txtbxInc_Mer_GotFocus()
    P_SelecTexto txtbxInc_Mer
End Sub

Private Sub txtbxInc_Mer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxInc_Mer_DblClick
    End If
End Sub

Private Sub txtbxInc_Mer_LostFocus()
    P_AjustaTexto txtbxInc_Mer
End Sub

Private Sub txtbxNome_Mer_GotFocus()
    
End Sub

Private Sub txtbxNome_Mer_LostFocus()
    
    P_SalvaComp
    P_BuscaItensCompra lblCompra.caption
End Sub


Private Sub txtbxSustent_GotFocus()
    P_SelecTexto txtbxSustent
End Sub

Private Sub txtbxSustent_LostFocus()
    P_ValorLostFocus txtbxSustent, "#0"
    txtbxTotalDias.text = Int(txtbxSustent.text) + Int(txtbxDuracao.text) + Int(txtbxEntrega.text)
End Sub
