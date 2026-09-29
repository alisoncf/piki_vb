VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEditaSaidas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edição de Notas de Saídas"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11355
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   5760
      Width           =   11295
      Begin VB.ComboBox cmbxTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "frmEditaSaidas.frx":0000
         Left            =   2040
         List            =   "frmEditaSaidas.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Selecione o tipo de localização dos dados."
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Height          =   495
         Left            =   10080
         Picture         =   "frmEditaSaidas.frx":002E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Imprime rascunho da Nota"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtbxLocaliza2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   5640
         TabIndex        =   7
         ToolTipText     =   "Informe a série da nota fiscal."
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtbxLocaliza2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   4200
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cmbxLocaliza 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "frmEditaSaidas.frx":08F8
         Left            =   120
         List            =   "frmEditaSaidas.frx":0902
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Selecione o tipo de localização dos dados."
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtbxLocaliza 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   4200
         TabIndex        =   8
         ToolTipText     =   "Informe a expressão a ser localizada."
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton cmndbtnLocalizar 
         Height          =   495
         Left            =   7920
         Picture         =   "frmEditaSaidas.frx":091F
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Localiza a expressão digitada."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmndbtnVisualizaItem 
         Enabled         =   0   'False
         Height          =   495
         Left            =   9480
         Picture         =   "frmEditaSaidas.frx":11E9
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Visualiza os itens da Nota."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmndbtnFechar 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   10680
         Picture         =   "frmEditaSaidas.frx":1AB3
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fecha o formulário."
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "&Tipo:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblLocalizar 
         Appearance      =   0  'Flat
         Caption         =   "&Localizar por:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblExpressao 
         AutoSize        =   -1  'True
         Caption         =   "Expressão:"
         Height          =   195
         Left            =   4200
         TabIndex        =   5
         Top             =   120
         Width           =   780
      End
   End
   Begin VB.Data dtNotaSaida 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   270
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid dbgrdNotaSaida 
      Bindings        =   "frmEditaSaidas.frx":237D
      Height          =   5655
      Left            =   0
      OleObjectBlob   =   "frmEditaSaidas.frx":2397
      TabIndex        =   13
      Top             =   0
      Width           =   11295
   End
End
Attribute VB_Name = "frmEditaSaidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxLocaliza_Click()
'Configurando caixa de textos cfe seleção
With cmbxLocaliza
    If .Text = "Cliente" Then
       txtbxLocaliza2(0).Visible = False
       txtbxLocaliza2(1).Visible = False
       txtbxLocaliza.Visible = True
    ElseIf .Text = "Emissão" Then
       txtbxLocaliza2(0).Visible = True
       txtbxLocaliza2(1).Visible = False
       txtbxLocaliza.Visible = False
       txtbxLocaliza2(0).ToolTipText = "Entre com a data de recepção"
    ElseIf .Text = "Número/Série Nota" Then
       txtbxLocaliza2(0).Visible = True
       txtbxLocaliza2(1).Visible = True
       txtbxLocaliza.Visible = False
       txtbxLocaliza2(0).ToolTipText = _
       "Entre com o número da nota fiscal"
    End If
End With
'Informa a expressão a ser procurada
lblExpressao.caption = cmbxLocaliza.Text & ":"

End Sub

Private Sub cmndbtnEditar_Click()
    If dbgrdNotaSaida.Row = -1 Then
        MsgBox "Escolha a Nota que deseja Editar", vbCritical
        Exit Sub
    End If
    If dtNotaSaida.Recordset!Situacao_Sai <> "1" Then
        MsgBox "Esta nota não Pode ser Editada!", vbCritical
        Exit Sub
    End If
    frmSaidaMercadorias.txtbxNota_Sai.Tag = dtNotaSaida.Recordset!Nota_Sai
    frmSaidaMercadorias.txtbxSerie_Sai.Tag = dtNotaSaida.Recordset!Serie_Sai
    frmSaidaMercadorias.Show 1
End Sub

Private Sub cmndbtnEfetiva_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error GoTo trata_erro
    
    If MsgBox("Confirma Efetivação da Nota Fiscal: " & dtNotaSaida.Recordset!Nota_Sai & " ?", vbYesNo + vbQuestion) = vbYes Then
        With dtNotaSaida.Recordset
            strSQL = "Select * from tab_itensaida where nota_sai = '" & !Nota_Sai & ""
            strSQL = strSQL & " and serie_sai = '" & !Serie_Sai & "' "
            Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
            With rPesq
                If .RecordCount = 0 Then
                    MsgBox "Não há itens na nota!", vbCritical
                    .Close
                    Exit Sub
                End If
                Do While Not .EOF
                    strSQL = "update tab_cmplt01 set estoque_mer = estoque_mer -" & Str(!Quantidade_Its)
                    strSQL = strSQL & " where cod_mer = '" & !COD_MER & "' "
                    strSQL = strSQL & " and marca_mer = '" & !MARCA_MER & "'"
                    banco.Execute strSQL
                Loop
            End With
        End With
    End If
trata_erro:
    
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub cmndbtnImprimir_Click()
    On Error GoTo trata_erro
    If dtNotaSaida.Recordset.RecordCount = 0 Then Exit Sub
    P_RelSaidaMercadoria dtNotaSaida.Recordset!Nota_Sai, dtNotaSaida.Recordset!Serie_Sai
trata_erro:
End Sub

Private Sub cmndbtnLocalizar_Click()
  Dim strSQL As String
  ''On Error  GoTo TrataErro
  P_Status "Localizando expressão..."
  
  'Ativando o Grid e o botão de visualização de itens
  dbgrdNotaSaida.Enabled = True
  cmndbtnVisualizaItem.Enabled = True
  'Focalizando Grid
  dbgrdNotaSaida.SetFocus
  If cmbxLocaliza.Text = "Cliente" Then
     'Configura para localizar por nome do Cliente
     strExpr = "TS.NOME_TER Like '%" & Trim(txtbxLocaliza.Text) & "%'"
  ElseIf cmbxLocaliza.Text = "Número/Série Nota" Then
     'Localiza pelo número da nota
     strExpr = "NOTA_SAI like '" & txtbxLocaliza2(0).Text & "%' And "
     strExpr = strExpr + "SERIE_SAI Like '%" & Trim(txtbxLocaliza2(1).Text) & "%'"
  ElseIf cmbxLocaliza.Text = "Emissão" Then
     'Localiza pela data de recepção
     strExpr = "EMISSAO_SAI >= " & F_PesqData(txtbxLocaliza2(0).Text)
  End If
  If cmbxTipo.ListIndex = 1 Then
    strSQL = " and SERIE_SAI = 'CUPM'"
  ElseIf cmbxTipo.ListIndex = 2 Then
    strSQL = " and SERIE_SAI <> 'CUPM'"
  End If
  
  'Localizando cfe expressões passadas aos controles
  dtNotaSaida.RecordSource = "Select TS.*,TC.NOME_TER as nome   " & _
    " from TAB_SAIDAMERC01 AS TS, TAB_TERCEIROS AS TC " & _
    "WHERE TC.COD_TER = TS.COD_CLI and " & _
    strExpr & strSQL & " " & " Order By " & _
    IIf(Left(strExpr, 3) = "NOM", "NOME_TER, NOTA_SAI", _
    IIf(Left(strExpr, 3) = "EMI", "EMISSAO_SAI, NOTA_SAI", _
    "NOTA_SAI, NOTA_SAI"))
    
    
  dtNotaSaida.Refresh   'Atualizando tabela
  P_Status
  Exit Sub
  
  
TrataErro:
   P_MostraErro Me.Tag
End Sub

Private Sub cmndbtnVisualizaItem_Click()
  Dim strSQL  As String
  Dim rcrdstVerifica As Recordset
  
  With dtNotaSaida.Recordset
     If .RecordCount > 0 Then
        
        'Verificando se há itens a serem editados
        strSQL = "Select OPERACAO_SAI From TAB_SAIDAMERC01 Where "
        strSQL = strSQL & "NOTA_SAI = '" & !Nota_Sai
        strSQL = strSQL & "' And SERIE_SAI = '" & !Serie_Sai & "'"
        
        Set rcrdstVerifica = banco.OpenRecordset(strSQL, dbOpenSnapshot)
        If rcrdstVerifica!Operacao_Sai = "CANC" Then
           Beep
           MsgBox "Esta Nota Fiscal foi cancelada!!!", vbExclamation
        End If
        rcrdstVerifica.Close
        'Chamando o formulário de edição de itens da nota de Saídas
        frmDados_Not.caption = "Edição N.F.: " & _
        !Nota_Sai & " [" & !Nome_Ter & "]"
        
        'Passando argumentos para pesquisa de SQL
        frmDados_Not.Tag = !Nota_Sai & !Serie_Sai
        'Indicando ao formulário que é nota de saída
        frmDados_Not.dbgrdItens.Tag = "S"
        frmDados_Not.Show 1
     End If
  End With
End Sub

Private Sub Command1_Click()

End Sub

Private Sub dbgrdnotasaida_DblClick()
   'Visualiza os itens da nota
   cmndbtnVisualizaItem_Click
End Sub

Private Sub dtNotaSaida_Reposition()
    On Error Resume Next
    If dtNotaSaida.Recordset!Situacao_Sai = 2 Then
        cmndbtnEfetiva.Enabled = True
    Else
        cmndbtnEfetiva.Enabled = False
    End If
    If dtNotaSaida.Recordset!Situacao_Sai = 1 Then
        cmndbtnEditar.Enabled = True
    Else
        cmndbtnEditar.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    ''On Error  GoTo TrataErro
    dtNotaSaida.Connect = conCaminho_Dados & conBanco_Dados
    dtNotaSaida.RecordsetType = 2 'Snapshot
   
    'Preenchendo combo box
    cmbxLocaliza.Clear
    cmbxLocaliza.AddItem "Cliente"
    cmbxLocaliza.AddItem "Número/Série Nota"
    cmbxLocaliza.AddItem "Emissão"
    cmbxLocaliza.Text = "Cliente"
    Exit Sub
TrataErro:
    P_MostraErro
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strMens As String
    If KeyCode = vbKeyF1 Then
        strMens = "Alt + L => Localiza a expressão." & Chr(13)
        strMens = "Alt + V => Visualiza o itens da Nota Fiscal."
        P_MostraTeclaAtalho strMens
    ElseIf KeyCode = vbKeyL And Shift = vbAltMask Then
        'Localiza a expressão
        If F_ControleAtivo(cmndbtnLocalizar) Then
            cmndbtnLocalizar_Click
        End If
    ElseIf KeyCode = vbKeyV And Shift = vbAltMask Then
        'Visualiza os itens da nota
        If F_ControleAtivo(cmndbtnVisualizaItem) Then
            cmndbtnVisualizaItem_Click
        End If
    End If
End Sub

Private Sub txtbxLocaliza_GotFocus()
   'Seleciona o texto
   P_SelecTexto txtbxLocaliza
End Sub

Private Sub txtbxLocaliza_LostFocus()
   'Ajusta o texto para maiúsculo
   P_AjustaTexto txtbxLocaliza
End Sub

Private Sub txtbxLocaliza2_GotFocus(Index As Integer)
   'Seleciona o texto
   P_SelecTexto txtbxLocaliza2(Index)
End Sub

Private Sub txtbxLocaliza2_LostFocus(Index As Integer)
   'Ajusta a expressão digitada
   If Index = 0 Then
        If cmbxLocaliza.Text = "Emissão" Then
           'Ajusta data de recepção
           P_AjustaData txtbxLocaliza2(Index)
        Else
           'Ajusta o número da nota fiscal
           P_AjustaNota txtbxLocaliza2(Index)
        End If
   Else
       P_AjustaTexto txtbxLocaliza2(Index)
   End If
End Sub

