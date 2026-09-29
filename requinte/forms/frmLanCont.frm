VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmLanCont 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento de Contagem de Produto"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
   Icon            =   "frmLanCont.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data dtPro 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnExcluir 
      Caption         =   "Excluir"
      Height          =   495
      Left            =   4200
      Picture         =   "frmLanCont.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtbxVenc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   5
      Tag             =   "N"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtbxLote 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   3
      Tag             =   "N"
      Top             =   960
      Width           =   1575
   End
   Begin VB.Data dtEst 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnFEchar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   120
      Picture         =   "frmLanCont.frx":0E14
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton cmndbtnOk 
      Caption         =   "Adicionar"
      Height          =   495
      Left            =   4200
      Picture         =   "frmLanCont.frx":1362
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtbxQtd 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3120
      MaxLength       =   50
      TabIndex        =   7
      Text            =   "0,00"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtbxCodBarra 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "N"
      Top             =   360
      Width           =   5295
   End
   Begin MSDBGrid.DBGrid dbgrdEst 
      Bindings        =   "frmLanCont.frx":14AC
      Height          =   2655
      Left            =   5640
      OleObjectBlob   =   "frmLanCont.frx":14C0
      TabIndex        =   15
      Top             =   360
      Width           =   3735
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmLanCont.frx":2058
      Height          =   1455
      Left            =   3720
      OleObjectBlob   =   "frmLanCont.frx":206C
      TabIndex        =   17
      Top             =   3240
      Width           =   5655
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Vencimento:"
      Height          =   195
      Left            =   1800
      TabIndex        =   4
      Tag             =   "B"
      Top             =   720
      Width           =   885
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Lote:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Tag             =   "B"
      Top             =   720
      Width           =   360
   End
   Begin VB.Label lblCont 
      Caption         =   "Contagem:"
      Height          =   735
      Left            =   240
      TabIndex        =   14
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label lblPrincipio 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Princípio"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   765
   End
   Begin VB.Label lblCodBar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cód Barra:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   915
   End
   Begin VB.Label lblCod 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cód"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Qtd:"
      Height          =   195
      Left            =   3120
      TabIndex        =   6
      Tag             =   "B"
      Top             =   720
      Width           =   300
   End
   Begin VB.Label lblNome_pro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   645
   End
   Begin VB.Label lblExpr 
      AutoSize        =   -1  'True
      Caption         =   "Produto:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Tag             =   "B"
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "frmLanCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmndbtnExcluir_Click()
    Dim strSQL As String
    If MsgBox("deseja realmente excluir o item?", vbExclamation + vbYesNo) = vbYes Then
        strSQL = "delete from tab_itcont where seq_cnt='" & Me.tag & "'"
        strSQL = strSQL & " and cod_pro='" & dtEst.Recordset!cod_pro & "' "
        strSQL = strSQL & " and lote_pro='" & dtEst.Recordset!lote_pro & "' "
        Banco.Execute strSQL
        buscaItensProd
        
    End If
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub cmndbtnOk_Click()
    Dim strSQL As String
    If txtbxCodBarra.tag = "" Then
        MsgBox "Escolha primeiro o produto!", vbExclamation
        txtbxCodBarra.SetFocus
        Exit Sub
    End If
    If MsgBox("Incluir contagem de: " & lblNome_pro.caption & " com " & _
        txtbxQtd.Text & " ?", vbYesNo + vbQuestion) = vbYes Then
        
        strSQL = F_BuscaUmCampo("qtd_itc", "tab_itcont", "seq_cnt='" & Me.tag & _
            "' and cod_pro='" & txtbxCodBarra.tag & "' and lote_pro='" & txtbxLote.Text & "'")
        If strSQL <> "" Then
            If MsgBox("Este item já se encontra lançado com " & strSQL & _
                 " Deseja alterar sua quantidade? ", vbYesNo + vbQuestion) = vbNo Then
                
                Exit Sub
            Else
                strSQL = " delete from tab_itcont where seq_cnt='" & Me.tag & "' "
                strSQL = strSQL & " and cod_pro='" & txtbxCodBarra.tag & "'"
                strSQL = strSQL & " and lote_pro='" & txtbxLote.Text & "' "
                Banco.Execute strSQL
            End If
                
        
            
        End If
        
        
        
        
        strSQL = " insert into tab_itcont(seq_cnt,cod_pro,qtd_itc,data_itc,lote_pro,venc_pro)values( "
        strSQL = strSQL & "'" & Me.tag & "', "
        strSQL = strSQL & "'" & txtbxCodBarra.tag & "', "
        strSQL = strSQL & F_GravaDec(txtbxQtd) & ","
        strSQL = strSQL & F_GravaDataTempo(Now) & ", "
        strSQL = strSQL & "'" & txtbxLote.Text & "',"
        strSQL = strSQL & F_GravaData(txtbxVenc.Text) & ") "
        
        Banco.Execute strSQL
        'txtbxCodBarra.Text = ""
        'txtbxCodBarra.tag = ""
        txtbxQtd.Text = "0,00"
        buscaItensProd
        
        txtbxCodBarra.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Me.tag = F_BuscaUmCampo("max(seq_cnt)", "tab_contagens", "")
    lblCont.caption = "Contagem: " & Me.tag & Chr(13) & _
    F_BuscaUmCampo("data_cnt", "tab_contagens", "seq_cnt='" & Me.tag & "'")
    
    dtEst.Connect = Banco.Connect
    dtPro.Connect = Banco.Connect
    
    
    
    
    
End Sub
Sub buscaItensProd()
    Dim strSQL As String
    strSQL = "select * from tab_itcont where seq_cnt='" & Me.tag & "' "
    strSQL = strSQL & " and cod_pro='" & txtbxCodBarra.tag & "' "
    dtEst.RecordSource = strSQL
    dtEst.Refresh
    
    strSQL = "select lote_pro,venc_pro,er_pro,ef_pro from tab_produtos as tp,tab_estoque as te  "
    strSQL = strSQL & " where tp.cod_pro=te.cod_pro and tp.cod_pro='" & txtbxCodBarra.tag & "' "
    dtPro.RecordSource = strSQL
    dtPro.Refresh
    
End Sub

Private Sub optnBar_Click()
    lblExpr.caption = "C&od. Barra:"
    lblExpr.tag = "B"
    txtbxCodBarra.SetFocus
End Sub

Private Sub optnC_Click()
        lblExpr.caption = "C&od. Prod.:"
        lblExpr.tag = "C"
        txtbxCodBarra.SetFocus
End Sub



Private Sub optnN_Click()
    lblExpr.caption = "N&ome Prod.:"
    lblExpr.tag = "N"
    
End Sub

Private Sub txbxVenc_Change()

End Sub

Private Sub txtbxVenc_GotFocus()
    P_SelecTexto txtbxVenc
End Sub

Private Sub txtbxCodBarra_GotFocus()
    P_SelecTexto txtbxCodBarra
End Sub

Private Sub txtbxCodBarra_LostFocus()
    P_AjustaTexto txtbxCodBarra
    'If optnC.Value = True Then
        'txtbxCodBarra.Text = Format(txtbxCodBarra.Text, "0000000#")
              
    'End If
    
End Sub
Private Sub txtbxCodBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtbxCodBarra.Text = "" Then
            MsgBox "Informe nome ou parte do nome ou codigo de barras do produto e tecle enter.", vbExclamation
            txtbxCodBarra.SetFocus
            Exit Sub
        End If
        busca
        
    End If
End Sub
Sub busca()
    Dim p As Produto
    txtbxCodBarra_LostFocus
    p = F_BuscaProdCabecalho(txtbxCodBarra.Text)
    If p.cod_pro <> "" Then
        txtbxCodBarra.Text = p.nome_pro
        txtbxCodBarra.tag = p.cod_pro
        lblNome_pro.caption = "Prod: " & p.nome_pro
        lblCod.caption = "Cód: " & p.cod_pro
        lblCodBar.caption = "Cód.Barra: " & F_BuscaUmCampo("codbarra_pro", "tab_codbarras", "cod_pro='" & p.cod_pro & "'")
        'lblVencimento.caption = p.est.venc_pro
        'lblLote.caption = p.est.lote_pro
        'lblSeq.caption = p.est.seq_pro
        buscaItensProd
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxLote_GotFocus()
    P_SelecTexto txtbxLote
End Sub

Private Sub txtbxLote_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxLote_LostFocus()
    P_AjustaTexto txtbxLote
End Sub

Private Sub txtbxQtd_GotFocus()
    P_SelecTexto txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnOk_Click
    End If
End Sub

Private Sub txtbxQtd_LostFocus()
    P_ValorLostFocus txtbxQtd
End Sub

Private Sub txtbxVenc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
    
End Sub

Private Sub txtbxVenc_LostFocus()
    P_AjustaData txtbxVenc
End Sub
