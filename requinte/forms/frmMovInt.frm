VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmMovInt 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Movimento Interno"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13065
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNumero 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdCadMovInt 
      Caption         =   "Novo"
      Height          =   255
      Left            =   5520
      TabIndex        =   35
      Top             =   960
      Width           =   615
   End
   Begin VB.ComboBox cmbxTipo 
      Height          =   315
      ItemData        =   "frmMovInt.frx":0000
      Left            =   240
      List            =   "frmMovInt.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   7320
      Width           =   12615
      Begin VB.TextBox txtbxUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6480
         TabIndex        =   16
         Text            =   "0,00"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   5640
         TabIndex        =   14
         Text            =   "1"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtbxProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   5175
      End
      Begin VB.CommandButton cmndbtnAdd 
         Caption         =   "Adicionar"
         Height          =   375
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edita"
         Height          =   375
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtbxPv 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7440
         TabIndex        =   18
         Text            =   "0,00"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnBuscaProd 
         Height          =   285
         Left            =   5280
         Picture         =   "frmMovInt.frx":0062
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Custo:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6480
         TabIndex        =   15
         Top             =   120
         Width           =   450
      End
      Begin VB.Label lblUltItem 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Produto"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Qtd:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5640
         TabIndex        =   13
         Top             =   120
         Width           =   300
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Preço"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7440
         TabIndex        =   17
         Top             =   120
         Width           =   420
      End
   End
   Begin VB.TextBox txtbxDest 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   4935
   End
   Begin VB.CommandButton cmndbtnConcluir 
      Caption         =   "&Concluir (F7)"
      Height          =   615
      Left            =   8760
      Picture         =   "frmMovInt.frx":03EC
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8520
      Width           =   2055
   End
   Begin VB.TextBox txtbxDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   5175
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   10920
      Picture         =   "frmMovInt.frx":0536
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8520
      Width           =   1935
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "Imprimir"
      Height          =   615
      Left            =   6600
      Picture         =   "frmMovInt.frx":0680
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8520
      Width           =   2055
   End
   Begin VB.PictureBox pcTop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   6360
      ScaleHeight     =   1785
      ScaleWidth      =   6465
      TabIndex        =   25
      Top             =   120
      Width           =   6495
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidades:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblUnid 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   33
         Tag             =   "!"
         Top             =   240
         Width           =   90
      End
      Begin VB.Label lblCusto 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   540
         Left            =   1320
         TabIndex        =   31
         Tag             =   "!"
         Top             =   960
         Width           =   780
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Custo:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   540
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   1110
      End
      Begin VB.Label lblProd 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   29
         Tag             =   "!"
         Top             =   0
         Width           =   90
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produtos:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   28
         Top             =   0
         Width           =   690
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   540
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label lblValor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   540
         Left            =   1320
         TabIndex        =   26
         Tag             =   "!"
         Top             =   480
         Width           =   780
      End
   End
   Begin VB.CommandButton cmndbtnBuscaDest 
      Height          =   285
      Left            =   5160
      Picture         =   "frmMovInt.frx":07CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   255
   End
   Begin MSDBGrid.DBGrid dbgrdProd 
      Bindings        =   "frmMovInt.frx":0B54
      Height          =   5175
      Left            =   240
      OleObjectBlob   =   "frmMovInt.frx":0B69
      TabIndex        =   32
      Top             =   2040
      Width           =   12615
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "N° Documento"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   1050
   End
   Begin VB.Label lblDest 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Destino:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   585
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Tipo:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Observação:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   915
   End
End
Attribute VB_Name = "frmMovInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Public tipo_mov As tipoMovimento
Dim bActive As Boolean
Sub P_InsereProd()
    Dim strSQL As String
    Dim iCount As Integer
    On Error GoTo trata_Erro
    strSQL = "delete from tab_itmov where num_mov='" & Me.tag & "' "
    Conn.Execute strSQL
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            strSQL = "insert into tab_itmov(num_mov,cod_pro,custo_ite,qtd_ite,seq_ite)values("
        
            strSQL = strSQL & "'" & Me.tag & "', "
            .col = 5
            strSQL = strSQL & "'" & .text & "',"
            .col = 3
            strSQL = strSQL & F_GravaDec(.text) & ","
            .col = 2
            strSQL = strSQL & F_GravaDec(.text) & ","
            strSQL = strSQL & iCount & ")"
            Conn.Execute strSQL
        Next
    End With

    'RESERVANDO O ESTOQUE
    cmndbtnImprimir.Enabled = True
    cmndbtnSalvar.Enabled = False
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub P_Conclui()
    Dim strSQL As String
    Dim iCount As Integer
    On Error GoTo trata_Erro
    Dim sProtocolo As String
    If tipo_mov = entrada Then
        If Trim(txtNumero.text) = "" Then
            MsgBox "Informe o número!", vbExclamation
            txtNumero.SetFocus
            Exit Sub
        End If
    End If
    
    If Not F_senha Then Exit Sub
    
    
    If strEfetnota_usu <> "S" Then
        MsgBox "Você não tem permissão para efetivar nota", vbExclamation
        Exit Sub
    End If
    
    
    
    With dtProd.Recordset
        .MoveFirst
        Do While Not .EOF
            '.Row = iCount
            
            strSQL = "update tab_produtos set "
            
            If tipo_mov = saida Then
                
                strSQL = strSQL & " er_pro = er_pro - "
            Else
                strSQL = strSQL & "er_pro = er_pro + "
            End If
            
            strSQL = strSQL & F_GravaDec(!qtd_ite) & "  "
            
            strSQL = strSQL & " where cod_pro="
            
            strSQL = strSQL & "'" & !cod_pro & "' "
            
            Conn.Execute strSQL
            .MoveNext
        Loop
        sProtocolo = F_BuscaProtocolo()
        
        strSQL = "update tab_movint set situacao_mov=2 ,"
        strSQL = strSQL & " cod_usu='" & strCod_Usu & "',"
        strSQL = strSQL & " protocolo_mov='" & sProtocolo & "',"
        strSQL = strSQL & " data_mov=" & F_GravaData(Date)
        strSQL = strSQL & " where num_mov='" & Me.tag & "'  "
        
        Conn.Execute strSQL
        MsgBox "Movimento concluído!", vbInformation
        
    End With
    If MsgBox("Deseja Gerar Arquivo?", vbYesNo + vbQuestion) = vbYes Then
        P_GeraArquivoTransfMovInt Me.tag, ""
    End If
    ImprMovInt Me.tag
    
    
    Unload Me
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub


Sub insere()
    Dim iPos As Integer
    Dim i As Integer
    Dim sSeq As String
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    Dim cQtd As Currency
    
    
    If txtbxProduto.tag = "" Then
        MsgBox "Informe o produto!", vbExclamation
        Exit Sub
    End If
    If CCur(txtbxQtd.text) = 0 Then
        MsgBox "informe uma quantidade válida!", vbExclamation
        txtbxQtd.SetFocus
        Exit Sub
    End If
    If Me.tag = "" Then
        MsgBox "Não existe um movimento interno gerado!", vbExclamation
        Exit Sub
    End If
    
    If cmndbtnAdd.tag = "" Then
        strSQL = "select qtd_ite,seq_ite from tab_itmov where num_mov='" & Me.tag & "' and cod_pro='" & txtbxProduto.tag & "'"
        abreObjetoAdo rPesq, strSQL
        
        If rPesq.RecordCount > 0 Then
            sSeq = rPesq!seq_ite
            cQtd = rPesq!qtd_ite
            FechaObjeto rPesq
            
        End If
        
        If sSeq = "" Then
        
        
            MovInt.P_InsereItMov Me.tag, txtbxProduto.tag, txtbxQtd.text, txtbxUnit.text, txtbxPv.text
        Else
            If MsgBox("O Produto " & txtbxProduto.tag & " já está inserido. Deseja adicionar " & txtbxQtd.text & _
                " à quantidade de " & cQtd & " já gravada?", vbYesNo + vbQuestion) = vbYes Then
                MovInt.P_AlteraItMov Me.tag, sSeq, txtbxProduto.tag, cQtd + CCur(txtbxQtd.text), txtbxUnit.text, txtbxPv.text
            End If
        End If
        
        
        cmndbtnAdd.tag = ""
        
    Else
        MovInt.P_AlteraItMov Me.tag, cmndbtnAdd.tag, txtbxProduto.tag, CCur(txtbxQtd.text), txtbxUnit.text, txtbxPv.text
    End If
    cmndbtnExcl.caption = "Excluir"
    cmndbtnAdd.caption = "Adicionar"
    cmndbtnAdd.tag = ""
    txtbxQtd.text = "0"
    buscaItens
    txtbxProduto.SetFocus
     
End Sub


Private Sub cmbxTipo_Click()
    If cmbxTipo.ListIndex = 3 Then
        tipo_mov = entrada
        lblDest.caption = "Origem/Fornecedor:"
        pcTop.BackColor = vbGreen
        
        
    Else
        lblDest.caption = "Destino:"
        tipo_mov = saida
        pcTop.BackColor = vbWhite
    End If
End Sub

Private Sub cmdCadMovInt_Click()
    P_CadDestMovInt
End Sub

Private Sub cmdEdit_Click()
    On erro GoTo trata_Erro

    With dtProd.Recordset
        txtbxProduto.text = !nome_pro
        txtbxProduto.tag = !cod_pro
        txtbxQtd.text = !qtd_ite
        txtbxUnit.text = !custo_ite
        txtbxPv.text = !unit_ite
        cmndbtnAdd.tag = !seq_ite
        cmndbtnExcl.caption = "Cancelar"
        cmndbtnAdd.caption = "Alterar"
        txtbxProduto.SetFocus
        
    End With


Exit Sub




trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnAdd_Click()
    If Me.tag = "" Then
        If MsgBox("Confirma abertura de um novo movimento interno?", vbYesNo + vbQuestion) = vbYes Then
            salva
        End If
        
    Else
        
    End If
    
    insere
    
End Sub

Private Sub cmndbtnBuscaDest_Click()
    Dim m As DestinoMovInt
    m = F_BuscaDestMovInt(txtbxDest.text)
    If m.cod_Des <> "" Then
        txtbxDest.text = m.nome_des
        txtbxDest.tag = m.cod_Des
        txtbxDesc.SetFocus
    End If
End Sub

Private Sub cmndbtnBuscaProd_Click()
    buscaProd
End Sub

Private Sub cmndbtnConcluir_Click()
        
        If MsgBox("Ao concluir o estoque será alterado. Confirma este processo?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
        
        
        
        If cmbxTipo.ListIndex = -1 Then
            MsgBox "Informe o tipo do movimento interno!", vbCritical
            Exit Sub
        End If
        
        
        
        
        
        P_Conclui
End Sub

Private Sub cmndbtnExcl_Click()
    On Error GoTo trata_Erro
    
    If cmndbtnAdd.tag <> "" Then
        cmndbtnAdd.tag = ""
        cmndbtnExcl.caption = "Excluir"
        cmndbtnAdd.caption = "Adicionar"
        Exit Sub
    End If
    
    With dtProd.Recordset
        If MsgBox("Confirma exclusão do item " & !nome_pro & "?", vbYesNo + vbCritical) = vbNo Then
            Exit Sub
        End If
        P_DeletaItMov Me.tag, !seq_ite
        buscaItens
    End With
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnImprimir_Click()
    ImprMovInt Me.tag
End Sub

Private Sub cmndbtnSair_Click()
    
    
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    
    If strEfetnota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!(Efetiva Nota Fiscal)", vbExclamation
        Exit Sub
    End If
    If Not F_senha(True) Then
        Exit Sub
    End If
        
        If cmbxTipo.text = "" Then
            MsgBox "Primeiro informe o tipo de Movimento Interno!", vbExclamation
            cmbxTipo.SetFocus
            Exit Sub
        End If
        If txtbxDest.tag = "" Then
            MsgBox "Primeiro escolha o destino!", vbExclamation
            txtbxDest.SetFocus
            Exit Sub
        End If
        salva

    
        
    
    
    
End Sub
Sub salva()
    Dim strSQL As String
    Dim iCount As Integer
  
    If Me.tag = "" Then
        strSQL = F_GeraNovoMovInt(Left(cmbxTipo.text, 1), txtbxDest.tag, txtbxDesc.text, lblValor.caption, txtNumero.text)
        Me.tag = strSQL
    Else
        strSQL = "update tab_movint set valor_mov=" & F_GravaDec(lblValor) & ","
        strSQL = strSQL & "tipo_mov='" & Left(cmbxTipo.text, 1) & "',"
        strSQL = strSQL & "desc_mov='" & txtbxDesc.text & "' "
        strSQL = strSQL & " where num_mov='" & Me.tag & "' "
        Conn.Execute strSQL
        strSQL = Me.tag
    End If
    
End Sub



Private Sub dbgrdProd_DblClick()
    cmdEdit_Click
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        carrega
    End If
    bActive = True
End Sub
Sub buscaItens()
    Dim strSQL As String
    
    
    
    strSQL = "select tp.nome_pro,tp.ref_pro,tp.cr_pro,tp.pv_pro,"
    strSQL = strSQL & " (cr_pro*qtd_ite) as custot,(pv_pro*qtd_ite) as total,"
    strSQL = strSQL & " ti.* from tab_itmov as ti,tab_produtos as tp "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and nuM_mov='" & Me.tag & "' "
    strSQL = strSQL & " order by seq_ite "
    
    dtProd.RecordSource = strSQL
    dtProd.Refresh
    soma
End Sub
Sub soma()
    Dim cTotal As Currency
    Dim cCusto As Currency
    Dim cUn As Integer
    With dtProd.Recordset
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            cTotal = cTotal + !total
            cCusto = cCusto + !custot
            cUn = cUn + !qtd_ite
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
        lblProd.caption = .RecordCount
    End With
    lblCusto.caption = FormatVal(cCusto)
    lblUnid.caption = cUn
    lblValor.caption = FormatVal(cTotal)
    
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select * from tab_movint where num_mov='" & Me.tag & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        txtbxDesc.text = !Desc_mov
        
        txtbxDest.tag = !cod_Des
        txtbxDest.text = F_BuscaUmCampo("nome_des", "tab_destmovint", "cod_des='" & !cod_Des & "'")
        P_LocalizaOperacao !tipo_mov, cmbxTipo
    End With
    buscaItens
    
    
trata_Erro:
    P_MostraErro
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        editaItem
    End If
    If KeyCode = vbKeyF7 Then
        cmndbtnConcluir_Click
    End If
End Sub

Private Sub txtbxDesc_GotFocus()
    P_SelecTexto txtbxDesc
End Sub

Private Sub txtbxDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtbxDesc_LostFocus()
    P_AjustaTexto txtbxDesc
End Sub

Private Sub Form_Load()
    bActive = False
    Me.top = 10: Me.Left = 10
    
    Set ado = New ADODB.Recordset
    tipo_mov = saida
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyF4 Then
        exclui
    End If
    If KeyCode = vbKeyReturn Then
        edita
    End If
End Sub
Sub edita()
    With msflxgrd
            If .col = 2 Or .col = 3 Then
                txtbxGrid.text = .text
                txtbxGrid.Left = .CellLeft + msflxgrd.Left
                txtbxGrid.top = .CellTop + msflxgrd.top
                txtbxGrid.height = .CellHeight
                txtbxGrid.width = .CellWidth
                txtbxGrid.Visible = True
                txtbxGrid.SetFocus
                
            End If
        End With
End Sub
Sub exclui()
    With msflxgrd
        If .rows = 2 Then
            .col = 4
            lblValor.caption = FormatVal(CCur(lblValor.caption) - CCur(.text))
            msflxgrd.rows = 1
        ElseIf .rows > 2 Then
            
            .col = 4
            lblValor.caption = FormatVal(CCur(lblValor.caption) - CCur(.text))
            msflxgrd.RemoveItem .row
        Else
            MsgBox "Não há o que excluir!", vbCritical
        End If
    End With
    If msflxgrd.rows > 1 Then
        cmndbtnImprimir.Enabled = True
    End If
    cmndbtnSalvar.Enabled = True
    'ZebraGrid msflxgrd, 2, &HD3E7EB
    lblProd.caption = msflxgrd.rows
End Sub


Private Sub txtbxDest_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscaDest_Click
    End If
End Sub



Sub SomaGrid()
    Dim iCount As Integer
    Dim cTotal As Currency
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 4
            cTotal = cTotal + CCur(.text)
            
        Next
        lblValor.caption = FormatVal(cTotal)
        lblProd.caption = .rows - 1
    End With
End Sub
Sub SomaLinha()
    Dim cQtd As Currency
    Dim cUnit As Currency
    msflxgrd.col = 2
    cQtd = CCur(msflxgrd.text)
    msflxgrd.col = 3
    cUnit = CCur(msflxgrd.text)
    msflxgrd.col = 4
    msflxgrd.text = FormatVal(cUnit * cQtd)
    
End Sub
Private Sub txtbxProduto_Change()
    txtbxProduto.tag = ""
End Sub

Private Sub txtbxProduto_GotFocus()
    P_SelecTexto txtbxProduto
End Sub

Private Sub txtbxProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cmbxTipo.text = "" Then
            MsgBox "Primeiro informe o tipo de Movimento Interno!", vbExclamation
            cmbxTipo.SetFocus
            Exit Sub
        End If
'        If F_Vazio(txtbxProduto.Text) Then
'            MsgBox "Informe o produto!", vbExclamation
'            txtbxProduto.SetFocus
'            Exit Sub
'        End If
        buscaProd
    End If
End Sub
Sub buscaProd()
    Dim p As Produto
    p = F_BuscaProdSaida(txtbxProduto.text, "B")
        If p.cod_pro <> "" Then
                
        txtbxProduto.text = p.nome_pro
        txtbxProduto.tag = p.cod_pro
        txtbxUnit.text = FormatVal(p.cr_pro)
        txtbxPv.text = FormatVal(p.pv_pro)
        txtbxQtd.SetFocus
        
       End If
                
End Sub
Private Sub txtbxQtd_GotFocus()
    
    P_SelecTexto txtbxQtd
End Sub

 Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxQtd_LostFocus()
    P_ValorLostFocus txtbxQtd, "###"
    
    
    
End Sub
Sub editaItem()
    txtbxQtd.Enabled = True
    txtbxQtd.text = "1"
    txtbxQtd.SetFocus
End Sub

Private Sub txtbxUnit_GotFocus()
    P_SelecTexto txtbxUnit
End Sub

Private Sub txtbxUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        insereManual
        
    End If
End Sub
Sub insereManual()
    Dim strSQL As String
    If txtbxProduto.tag = "" Then
        MsgBox "Informe o produto!", vbExclamation
        txtbxProduto.SetFocus
        Exit Sub
    End If
    If CCur(txtbxUnit.text) <= 0 Then
        MsgBox "O valor unitário está zerado!", vbExclamation
        Exit Sub
    End If
    If CCur(txtbxQtd.text) <= 0 Then
        MsgBox "Quantidade está zerada!", vbCritical
        Exit Sub
    End If
    
    strSQL = msflxgrd.rows & Chr(9) & txtbxProduto.text & Chr(9) & txtbxQtd.text & Chr(9)
    strSQL = strSQL & FormatVal(txtbxUnit.text) & Chr(9)
    strSQL = strSQL & FormatVal(CCur(txtbxUnit.text) * CCur(txtbxQtd.text)) & Chr(9) & txtbxProduto.tag
    msflxgrd.AddItem strSQL
                    
    lblValor.caption = FormatVal(CCur(lblValor.caption) + CCur(txtbxUnit.text) * CCur(txtbxQtd.text))
    
    cmndbtnImprimir.Enabled = True
    cmndbtnSalvar.Enabled = True
    txtbxProduto.text = ""
    txtbxQtd.text = "1"
    txtbxProduto.SetFocus
    
End Sub
Private Sub txtbxUnit_LostFocus()
    P_ValorLostFocus txtbxUnit
End Sub



Private Sub txtNumero_GotFocus()
    P_SelecTexto txtNumero
End Sub

Private Sub txtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtNumero_LostFocus()
    P_AjustaTexto txtNumero
End Sub
