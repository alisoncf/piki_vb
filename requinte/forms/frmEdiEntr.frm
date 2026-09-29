VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiEntr 
   Caption         =   "Edição de Compras"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14520
   Icon            =   "frmEdiEntr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   14520
   Begin VB.PictureBox pcLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      ForeColor       =   &H80000008&
      Height          =   8805
      Left            =   0
      ScaleHeight     =   8775
      ScaleWidth      =   2265
      TabIndex        =   11
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton cmdPLanilha 
         Caption         =   "Gerar Planilha"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiEntr.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   7200
         Width           =   2055
      End
      Begin VB.CommandButton cmdBusca 
         Appearance      =   0  'Flat
         Caption         =   "Pesquisar"
         Height          =   525
         Left            =   120
         Picture         =   "frmEdiEntr.frx":0557
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   21
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtProduto 
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   19
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton cmdProd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         Picture         =   "frmEdiEntr.frx":0650
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2640
         Width           =   255
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmEdiEntr.frx":0749
         Left            =   120
         List            =   "frmEdiEntr.frx":0759
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3840
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnTerc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         Picture         =   "frmEdiEntr.frx":0781
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2040
         Width           =   255
      End
      Begin VB.TextBox txtForn 
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   7
         Top             =   2040
         Width           =   1815
      End
      Begin VB.ComboBox cmbxTipo 
         Height          =   315
         ItemData        =   "frmEdiEntr.frx":087A
         Left            =   120
         List            =   "frmEdiEntr.frx":088D
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtbxDataIni 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtbxdataFin 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnCancelar 
         Caption         =   "&Cancelar"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiEntr.frx":08FD
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5760
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiEntr.frx":0A47
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4320
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiEntr.frx":0DD1
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5040
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiEntr.frx":0F1B
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   7920
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiEntr.frx":1131
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6480
         Width           =   2055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   2400
         Width           =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar por:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   3600
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedor:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   4
         Top             =   600
         Width           =   285
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiEntr.frx":153E
      Height          =   8175
      Left            =   2400
      TabIndex        =   17
      Top             =   0
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   14420
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
         DataField       =   "num_not"
         Caption         =   "Número"
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
         DataField       =   "nome_for"
         Caption         =   "Fornecedor"
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
         DataField       =   "emissao_not"
         Caption         =   "Emissão"
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
         DataField       =   "valor_not"
         Caption         =   "Valor"
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
         DataField       =   "status_not"
         Caption         =   "Status"
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
         DataField       =   "seq_not"
         Caption         =   "id"
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
         DataField       =   "data_not"
         Caption         =   "Data"
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
         DataField       =   "seq_not"
         Caption         =   "seq"
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
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3614,74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1260,284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   734,74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2340,284
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiEntr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmbxBuscar_Click()
    
    buscaR
End Sub

Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaR
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxTipo_Click()
    buscaR
End Sub

Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub buscaR()
    Dim strSQL As String
    Dim sOrd As String
    If cmbxOrder.ListIndex = 0 Then
        sOrd = "data_not desc, Num_Not"
    ElseIf cmbxOrder.ListIndex = 1 Then
        sOrd = "Nome_For"
    ElseIf cmbxOrder.ListIndex = 2 Then
        sOrd = "Emissao_Not"
    Else
        sOrd = "Entrada_Not"
    End If
    strSQL = "select tc.*,tf.nome_for,tf.uf_for from tab_fornecedores as tf, tab_compras as tc "
    strSQL = strSQL & " where tc.cod_for=tf.cod_for "
    If txtForn.tag <> "" Then
        strSQL = strSQL & " and tf.cod_for ='" & txtForn.tag & "' "
    End If
    If txtNumero.Text <> "" Then
    
        strSQL = strSQL & " and tc.num_not like '" & txtNumero.Text & "%' "
    End If
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and tc.entrada_not < " & F_PesqData(txtbxdataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and tc.entrada_not > " & F_PesqData(txtbxDataIni.Text)
    End If
    If cmbxTipo.ListIndex > -1 Then
        strSQL = strSQL & " and tc.status_not='" & Left(cmbxTipo.Text, 1) & "' "
    End If
    strSQL = strSQL & " order by " & sOrd
    
    abreObjetoAdo ado, strSQL, dtgrd
End Sub

Private Sub cmdBusca_Click()
    buscaR
End Sub

Private Sub cmdPLanilha_Click()
    ExportaGridToCSV dtgrd, ado
End Sub

Private Sub cmndbtnCancelar_Click()
    Dim strSQL As String
    With ado
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If MsgBox("Deseja Realmente Excluir a Nota Fiscal " & !num_not & _
            " ?", vbYesNo + vbCritical) = vbYes Then
            strSQL = "delete from tab_compras where seq_not = '" & !seq_not & "'"
            
            Conn.Execute strSQL
            strSQL = "delete from tab_itcompras where seq_not = '" & !seq_not & "'"
            
            Conn.Execute strSQL
            
            ado.Requery
        End If
    End With
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmEntradas
    fr.operacao = "I"
    If ado.RecordCount <= 0 Or dtgrd.row = -1 Then
        MsgBox "Não há notas fiscais nesta pesquisa!", vbExclamation
        Exit Sub
    End If
    
    With ado
        If !status_not <> "1" Then
            MsgBox "A Nota Não está Aberta!", vbExclamation
            Exit Sub
        End If
        
        fr.tag = !seq_not
        fr.grid.tag = fr.tag
        fr.caption = "Nota Fiscal: " & !num_not & " de " & !nome_for
        fr.txtbxDesconto.Text = FormatVal(!desconto_not)
        fr.txtbxNum_Not.Text = !num_not
        fr.txtbxNum_Not.tag = !num_not
        fr.txtbxSerie_Not.Text = !serie_not
        fr.txtbxSerie_Not.tag = !serie_not
        fr.txtbxBC_Not.Text = Format(!bc_not, "##,###,##0.00")
        fr.txtbxValorProd.Text = FormatVal(!valorprod_not)
        fr.txtbxDepesas_Not.Text = Format(!despesas_not, "###,###,##0.00")
        fr.txtbxFrete.Text = Format(!frete_not, "###,###,##0.00")
        fr.txtbxEmissao_Not.Text = Format(!emissao_not)
        fr.txtbxEntrada_Not.Text = Format(!entrada_not)
        fr.txtbxICMS_Not.Text = Format(!icms_not, "###,###,##0.00")
        fr.txtbxIPI_Not.Text = Format(!ipi_Not, "###,###,##0.00")
        fr.txtbxNOme_Ter.Text = !nome_for
        fr.txtbxNOme_Ter.tag = !cod_for
        fr.lblUF_Ter.caption = Format(!uf_for)
        
        fr.txtbxValor_Not.Text = FormatVal(!valor_not)
        
        P_PreencheOperacoes fr.cmbxTipo_Op, "E"
        
         
        P_LocalizaOperacao !operacao_not, fr.cmbxTipo_Op
        
        
        fr.Show 0, mdiFrm
    End With
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmEntradas
    fr.operacao = "I"
    fr.Show 0, mdiFrm
End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    ado.Close
    Unload Me
End Sub

Private Sub cmndbtnTerc_Click()
     
    Dim f As fornecedor.fornecedor
    If f.cod_for <> "" Then
        txtForn.Text = f.nome_for
        txtForn.tag = f.cod_for
        
    End If
End Sub


Private Sub dtgrd_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub dtgrd_GotFocus()
    dtgrd.MarqueeStyle = 3
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    End If
End Sub

Private Sub dtgrd_LostFocus()
    dtgrd.MarqueeStyle = 0
End Sub

Private Sub adc_Reposition()
    With ado
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If !status_not = "1" Then
            cmndbtnEdit.Enabled = True
            cmndbtnCancelar.Enabled = True
        ElseIf !status_not = "2" Then
            cmndbtnEdit.Enabled = False
            cmndbtnCancelar.Enabled = True
        Else
            
            cmndbtnEdit.Enabled = False
            cmndbtnCancelar.Enabled = False
        
        End If
        
    End With
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        
        frmNotaEntrada.Show 0
        
    End If
End Sub

Private Sub Form_Load()
    resizeForm Me
    Set ado = New ADODB.Recordset

    cmbxOrder.ListIndex = 0
    cmbxTipo.ListIndex = 0
    
    buscaR
End Sub

Private Sub Form_Resize()
On Error Resume Next
    dtgrd.width = Me.width - dtgrd.Left - 200
    dtgrd.height = Me.height - dtgrd.Top - 300
'    vsc.Height = pcLeft.Height - 100
'    vsc.max = pcLeft.Height
    
End Sub


Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaR
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
        buscaR
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And cmndbtnTerc.Visible = True Then
        cmndbtnTerc_Click
    End If
    buscaR
End Sub

Private Sub vsc_Change()
    'pcLeft.Top = 0 - vsc.Value
End Sub
