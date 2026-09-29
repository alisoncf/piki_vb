VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmArranjos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aranjos"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   Icon            =   "frmArranjos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   9210
   Begin VB.Frame frControle 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin MSDBGrid.DBGrid dbgrdProd 
         Bindings        =   "frmArranjos.frx":030A
         Height          =   3495
         Left            =   2040
         OleObjectBlob   =   "frmArranjos.frx":031F
         TabIndex        =   16
         Top             =   240
         Width           =   6855
      End
      Begin VB.Frame frOrd 
         Caption         =   "Ordem:"
         Height          =   1095
         Left            =   120
         TabIndex        =   12
         Top             =   3600
         Width           =   1815
         Begin VB.OptionButton optnDesc 
            Caption         =   "Descrição"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton optnCod 
            Caption         =   "Códi&go"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optnSeq 
            Caption         =   "Seqüência"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame frBotoes 
         Height          =   975
         Left            =   2040
         TabIndex        =   17
         Top             =   3720
         Width           =   6975
         Begin VB.CommandButton cmndbtnCancelar 
            Caption         =   "Cancelar Arranjo"
            Height          =   735
            Left            =   2400
            Picture         =   "frmArranjos.frx":1078
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   180
            Width           =   2175
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   735
            Left            =   4680
            Picture         =   "frmArranjos.frx":11C2
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   180
            Width           =   2175
         End
         Begin VB.CommandButton cmndbtnSalvar 
            Caption         =   "&Concluir"
            Height          =   735
            Left            =   120
            Picture         =   "frmArranjos.frx":130C
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   180
            Width           =   2175
         End
      End
      Begin VB.Frame frOpt 
         Caption         =   "Produtos:"
         Height          =   3375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1815
         Begin VB.OptionButton optnBar 
            Caption         =   "&B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   120
            Width           =   255
         End
         Begin VB.OptionButton optnN 
            Caption         =   "N"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton optnC 
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   255
         End
         Begin VB.TextBox txtbxCodBarra 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   3
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton cmndbtnEditProd 
            Caption         =   "&Editar"
            Enabled         =   0   'False
            Height          =   615
            Left            =   120
            Picture         =   "frmArranjos.frx":1456
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   2560
            Width           =   1575
         End
         Begin VB.CommandButton cmndbtnExc_Prod 
            Caption         =   "E&xcluir"
            Height          =   615
            Left            =   120
            Picture         =   "frmArranjos.frx":15A0
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1920
            Width           =   1575
         End
         Begin VB.CommandButton cmndbtnAddProd 
            Caption         =   "&Adicionar"
            Height          =   615
            Left            =   120
            Picture         =   "frmArranjos.frx":16EA
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1280
            Width           =   1575
         End
         Begin VB.TextBox txtbxQtd 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   8
            Text            =   "0,0"
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblExpr 
            AutoSize        =   -1  'True
            Caption         =   "&C&ód. Barra:"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Tag             =   "B"
            Top             =   240
            Width           =   795
         End
         Begin VB.Label lblMarca 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   3120
            Width           =   45
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Quantidade:"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Tag             =   "B"
            Top             =   840
            Width           =   870
         End
      End
      Begin VB.Data dtProd 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2640
         Visible         =   0   'False
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmArranjos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Private Sub cmndbtnAddProd_Click()
    Dim strSQL As String
    Dim sItem As String
    
    If Me.Tag = "" Then
        If MsgBox("Confirma Geração de um novo lançamento de arranjo?", vbYesNo _
            + vbQuestion) = vbYes Then
            Me.Tag = F_GeraNovoArranjo
            If Me.Tag = "" Then
                MsgBox "Erro!", vbCritical
                Exit Sub
            End If
            Me.caption = "Arranjo " & Me.Tag
        End If
            
        
    End If
    If txtbxCodBarra.Tag = "" Then
        MsgBox "Escolha o produto!", vbExclamation
        Exit Sub
    End If
    If F_BuscaUmCampo("cod_pro", "tab_itarr", "num_arr='" & Me.Tag & "' and cod_pro='" & _
        txtbxCodBarra.Tag & "'") <> "" Then
        MsgBox "Este item já se encontra lançado!", vbExclamation
        Exit Sub
    End If
    
    strSQL = "insert into tab_itarr(num_arr,cod_pro,qtd_ite,item_arr)values("
    strSQL = strSQL & "'" & Me.Tag & "',"
    strSQL = strSQL & "'" & txtbxCodBarra.Tag & "',"
    strSQL = strSQL & F_GravaDec(txtbxQtd.Text) & ","
    sItem = F_BuscaUmCampo("max(item_arr)", "tab_itarr", "seq_arr='" & Me.Tag & "'")
    If sItem = "" Then
        sItem = "1"
    Else
        sItem = Int(sItem + 1)
    End If
    strSQL = strSQL & sItem & ")"
    banco.Execute strSQL
    txtbxCodBarra.Text = ""
    txtbxCodBarra.Tag = ""
    txtbxQtd.Text = "0,0"
    
    Busca
End Sub

Private Sub cmndbtnCancelar_Click()
    Dim strSQL As String
    On Error Resume Next
    If Me.Tag = "" Then
        Unload Me
        Exit Sub
    End If
    If MsgBox("Tem certeza que deseja cancelar o arranjo" & Me.Tag & "?", vbYesNo + vbCritical) = vbYes Then
        strSQL = "delete from tab_arranjos where num_arr='" & Me.Tag & "' "
        banco.Execute strSQL
        strSQL = "delete from tab_arranjos where num_arr='" & Me.Tag & "' "
        banco.Execute strSQL
        MsgBox "Arranjo excluído!", vbExclamation
        If IsLoad("frmEdiArr") Then
            frmEdiArr.dtVendas.Refresh
        End If
        Unload Me
    End If
    
    
End Sub

Private Sub cmndbtnExc_Prod_Click()
    Dim strSQL As String
    On Error GoTo trata_erro
    With dtProd.Recordset
    If .RecordCount = 0 Then
        MsgBox "Não há o que ser excluído!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma exclusão do lançamento?", vbYesNo + vbCritical) = vbYes Then
        strSQL = "delete from tab_itarr where cod_pro='" & !cod_pro & "' "
        strSQL = strSQL & " and num_arr='" & Me.Tag & "' "
        banco.Execute strSQL
        Busca
    End If
        
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
        If Me.Tag = "" Then
            MsgBox "Não há arranjo lançado!", vbExclamation
            Exit Sub
        End If
        If MsgBox("Confirma conclusão dos lançamentos?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
        
        If dtProd.Recordset.RecordCount = 0 Then
            MsgBox "Não há itens para conclusão!", vbExclamation
            Exit Sub
        End If
        Baixa
        MsgBox "Lançamento de Arranjo concluído!", vbInformation
        imprimePerda Me.Tag
        Unload Me
    
End Sub
Sub Baixa()
    Dim strSQL As String
    Dim cEst As Currency
    strSQL = "update tab_arranjos set situacao_arr='2' where num_Arr='" & Me.Tag & "' "
    banco.Execute strSQL

    With dtProd.Recordset
        Do While Not .EOF
            strSQL = F_BuscaUmCampo("e2_pro", "tab_Composicao", "cod_pro='" & !cod_pro & "'")
            If !qtd_ite > CCur(strSQL) Then
                cEst = 0
            Else
                cEst = CCur(strSQL) - !qtd_ite
            End If
            strSQL = "update tab_composicao set e2_pro=" & F_GravaDec(cEst)
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            
            
            banco.Execute strSQL
            
            .MoveNext
        Loop
        
    End With
End Sub
Private Sub Form_Activate()
    If bActive = True Then Exit Sub
    If Me.Tag = "" Then
        cmndbtnSalvar.caption = "&Abrir Lançamento"
    Else
        cmndbtnSalvar.caption = "&Concluir Lançamento"
    End If
    Busca
    bActive = True
End Sub
Sub Busca()
    Dim strSQL As String
    strSQL = "select ti.*,tp.nome_pro from tab_produtos as tp, "
    strSQL = strSQL & "tab_itarr ti, tab_arranjos as ta where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & "and ta.num_arr=ti.num_arr"
    strSQL = strSQL & " and ti.num_arr='" & Me.Tag & "' "
    If optnCod.Value = True Then
        strSQL = strSQL & "  order by ti.cod_pro"
    ElseIf optnSeq.Value Then
        strSQL = strSQL & "  order by ti.item_arr"
    Else
        strSQL = strSQL & "  order by tp.nome_pro"
    End If
    
    dtProd.RecordSource = strSQL
    dtProd.Refresh
End Sub
Private Sub Form_Load()
    P_CentraForm Me
    dtProd.Connect = banco.Connect
End Sub

Private Sub frCab_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub txtbxCodBarra_GotFocus()
    P_SelecTexto txtbxCodBarra
End Sub

Private Sub txtbxCodBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim p As Produto
    If KeyCode = vbKeyReturn Then
        txtbxCodBarra_LostFocus
        If optnBar.Value Then
            p = F_BuscaProdSaida(txtbxCodBarra.Text, "B")
        ElseIf optnC.Value Then
            p = F_BuscaProdSaida(txtbxCodBarra.Text, "C")
        Else
            p = F_BuscaProdSaida(txtbxCodBarra.Text, "N")
        End If
    End If
    If p.cod_pro <> "" Then
        txtbxCodBarra.Text = p.nome_pro
        txtbxCodBarra.Tag = p.cod_pro
        txtbxQtd.SetFocus
    
    End If
End Sub

Private Sub txtbxCodBarra_LostFocus()
    P_AjustaTexto txtbxCodBarra
End Sub


Private Sub txtbxNOme_Ter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
    End If
End Sub


Private Sub txtbxQtd_GotFocus()
    P_SelecTexto txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxQtd_LostFocus()
    P_ValorLostFocus txtbxQtd, "###,##0.0"
End Sub


