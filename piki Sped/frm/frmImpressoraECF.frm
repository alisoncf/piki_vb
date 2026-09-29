VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmImpressoraECF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle do ECF"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   Icon            =   "frmImpressoraECF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      Picture         =   "frmImpressoraECF.frx":27A2
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   24
      Top             =   4560
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      Picture         =   "frmImpressoraECF.frx":28EC
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   23
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton cmndbtnFechar 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   1935
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Fecha o formulário."
      Top             =   4440
      Width           =   2280
   End
   Begin TabDlg.SSTab sstb 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&Relatórios Fiscais"
      TabPicture(0)   =   "frmImpressoraECF.frx":2A36
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Visualizações"
      TabPicture(1)   =   "frmImpressoraECF.frx":2A52
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Controle Geral ECF"
      TabPicture(2)   =   "frmImpressoraECF.frx":2A6E
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Msg. Promoc."
      TabPicture(3)   =   "frmImpressoraECF.frx":2A8A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(1)=   "Frame6"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame6 
         Caption         =   "Opções de Mensagem:"
         Height          =   1215
         Left            =   -74520
         TabIndex        =   20
         Top             =   2280
         Width           =   4575
         Begin VB.OptionButton optnCumpr 
            Caption         =   "&2 Emitir Mensagem Promocional"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   2655
         End
         Begin VB.OptionButton optnCumpr 
            Caption         =   "&1 Emitir Saudação. (Cumprimento)"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Value           =   -1  'True
            Width           =   3615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Mensa&gem Promocional (até 48 caracteres):"
         Height          =   975
         Left            =   -74520
         TabIndex        =   18
         Top             =   840
         Width           =   4575
         Begin VB.TextBox txtbxMsg 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            MaxLength       =   48
            TabIndex        =   19
            Text            =   "Obrigado Pela Preferência! Volte Sempre!"
            Top             =   360
            Width           =   4335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Visualizar Relatórios"
         Height          =   975
         Left            =   240
         TabIndex        =   12
         Top             =   3240
         Width           =   4935
         Begin VB.CommandButton cmndbtnVisualizarMemoriaFiscal 
            Caption         =   "&Memória Fiscal"
            Height          =   615
            Left            =   120
            Picture         =   "frmImpressoraECF.frx":2AA6
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   240
            Width           =   2295
         End
         Begin VB.CommandButton cmndbtnVisualizarLeituraX 
            Caption         =   "&Leitura X"
            Height          =   615
            Left            =   2520
            Picture         =   "frmImpressoraECF.frx":2BF0
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Visualização:"
         Height          =   3495
         Left            =   -74760
         TabIndex        =   15
         Top             =   600
         Width           =   5775
         Begin VB.TextBox txtbxVisualizacao 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3015
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   240
            Width           =   5415
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Controle Gerais (Visualização)"
         Height          =   2535
         Left            =   270
         TabIndex        =   5
         Top             =   600
         Width           =   4935
         Begin VB.CommandButton cmndbtnHorarioVerao 
            Caption         =   "&Horário de Verão"
            Height          =   615
            Left            =   120
            Picture         =   "frmImpressoraECF.frx":2D3A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Entra ou sai do horário de verão."
            Top             =   240
            Width           =   2295
         End
         Begin VB.CommandButton cmndbtnConfiguracoes 
            Caption         =   "Configurações"
            Height          =   615
            Left            =   120
            Picture         =   "frmImpressoraECF.frx":32C4
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Informa as configurações atuais da impressora."
            Top             =   1680
            Width           =   2295
         End
         Begin VB.CommandButton cmndbtnCancelaUltimoCupom 
            Caption         =   "&Cancela  Cupom"
            Height          =   615
            Left            =   2520
            Picture         =   "frmImpressoraECF.frx":384E
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Cancela o ÚLTIMO cupom fiscal, ou cupom aberto."
            Top             =   240
            Width           =   2295
         End
         Begin VB.CommandButton cmndbtnReseta 
            Caption         =   "&Reseta"
            Height          =   615
            Left            =   2520
            Picture         =   "frmImpressoraECF.frx":3DD8
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Corrige a impressora de eventuais erros."
            Top             =   1680
            Width           =   2295
         End
         Begin VB.CommandButton cmndbtnCaixaInicial 
            Caption         =   "&Controle Caixa"
            Enabled         =   0   'False
            Height          =   615
            Left            =   2520
            Picture         =   "frmImpressoraECF.frx":4362
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Utilizado para informar Entrada(Suprimento)/Saída(Sangria) do Caixa sem vínculo com ICMS."
            Top             =   960
            Width           =   2295
         End
         Begin VB.CommandButton cmndbtnEstadoImpressora 
            Caption         =   "&Estado Impressora"
            Height          =   615
            Left            =   120
            Picture         =   "frmImpressoraECF.frx":48EC
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Informa o estado atual da impressora."
            Top             =   960
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Relatórios Fiscais"
         Height          =   3255
         Left            =   -74520
         TabIndex        =   1
         Top             =   720
         Width           =   5055
         Begin VB.CommandButton cmndbtnLeituraX 
            Caption         =   "Leitura. &X"
            Height          =   615
            Left            =   1320
            Picture         =   "frmImpressoraECF.frx":4E76
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Emite relatório da Leitura X do ECF."
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton cmndbtnReducaoZ 
            Caption         =   "Redução. &Z"
            Height          =   615
            Left            =   1320
            Picture         =   "frmImpressoraECF.frx":5400
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Emite a Redução Z do ECF."
            Top             =   2040
            Width           =   2295
         End
         Begin VB.CommandButton cmndbtnLeituraMemoriaFiscal 
            Caption         =   "Memória. &Fiscal"
            Height          =   615
            Left            =   1320
            Picture         =   "frmImpressoraECF.frx":598A
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Emite relatório da Leitura da Memória Fiscal da Impressora."
            Top             =   600
            Width           =   2295
         End
      End
   End
End
Attribute VB_Name = "frmImpressoraECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub P_LeituraMemoriaFiscal(Optional strTipo As String = "R")
    'Emite a leitura da memória Fiscal
    'Padrão Retorno de Arquivo
    
    ''On Error  GoTo TrataErro
    Dim intRet As Integer
    Dim strExpr As String
    Dim strDataInic As String
    Dim strDataFin As String
    txtbxVisualizacao.Text = ""
    If Not F_AbrePortaECF Then
       'Não é possível estabelecer comunicação com a impressora
       Exit Sub
    End If
    
    
    
    'Entrando com a data inicial
    
       
       'Indica uma data inicial
        strDataInic = F_DataInicialFinal  'Chama a função de datas
       If strDataInic = "" Then
          'O Usuário cancelou a função
          Exit Sub
       End If
       strDataFin = Right(strDataInic, 10) 'Seta data final
       strDataInic = Left(strDataInic, 10) 'Seta data inicial
    
    
    'Entrando com a data final
   
    
    
    'Manda o comando de Leitura da Memória Fiscal
    strExpr = conAbre_Protocolo_Fiscal & "08|"
    strExpr = strExpr & Left(strDataInic, 2)
    strExpr = strExpr & Mid(strDataInic, 4, 2)
    strExpr = strExpr & Right(strDataInic, 2) & "|"
    
    strExpr = strExpr & Left(strDataFin, 2)
    strExpr = strExpr & Mid(strDataFin, 4, 2)
    strExpr = strExpr & Right(strDataFin, 2) & "|"
    strExpr = strExpr & strTipo & conFecha_Protocolo_Fiscal
    
    P_Status "Gerando Visualização..."
    
    intRet = FormataTX(strExpr)
    
    intRet = FechaPorta
    
    If strTipo <> "R" Then
       'É impressão então já foi feita
       P_Status ""
       Exit Sub
    End If
    
    
    
    
    txtbxVisualizacao = F_RetornoArquivoECF
    P_Status ""
    Exit Sub
TrataErro:
    P_Status ""
    P_MostraErro
End Sub

Private Sub P_EmiteLeituraX(Optional strTipo As String = "R")
    'Faz a leitura X padrão visualização
    ''On Error  GoTo TrataErro
    Dim intRet As Integer
    Dim strExpr As String
    Dim strDataInic As String
    Dim strDataFin As String
    
    txtbxVisualizacao.Text = "" 'Apagando caixa de texto
    
    If Not F_AbrePortaECF Then
       'Não é possível estabelecer comunicação com a impressora
       Exit Sub
    End If
          
        
    'Manda o comando de Leitura X
    strExpr = conAbre_Protocolo_Fiscal
    
    If strTipo = "R" Then
       'Retorno pelo arquivo
       strExpr = strExpr & "69"
    Else
       'Retorno pela impressora
       strExpr = strExpr & "06"
    End If
    
    
    strExpr = strExpr & conFecha_Protocolo_Fiscal
    
    P_Status "Gerando Visualização..."
    
    intRet = FormataTX(strExpr)
    
    intRet = FechaPorta
    
    If strTipo <> "R" Then
       'Foi impresso o relatório não precisa lê arquivo
       P_Status ""
       Exit Sub
    End If
    
    txtbxVisualizacao.Text = F_RetornoArquivoECF
    P_Status ""
    Exit Sub
TrataErro:
    P_Status ""
    P_MostraErro
End Sub




Private Sub cmndbtnCancelaUltimoCupom_Click()
    Dim strExpr As String, intReg As Integer
    Dim strSQL As String
    Dim rcrdstVerifica As Recordset
    Dim strNota_Sai As String
    Dim strCod_Mer As String
    Dim strMarca_Mer As String
    Dim sNum_Ven As String
    
    Beep
    If MsgBox("Você está prestes a cancelar o último cupom!" & _
    Chr(13) & "Tem certeza disto? ", vbQuestion & vbYesNo) <> vbYes Then
        Exit Sub
    End If
    
    
    
    'Verifica se é possível cancelar o último cupom fiscal
    
          
    'Verifica senha de confirmação
    
    
    
    P_Status "Cancelando cupom fiscal..."
    
    If InStr(1, F_ECFStatus, UCase("Cupom aberto")) <> 0 Then
       'Indica que há cupom aberto
       'Inicia comunicação com a impressora
       IniPortaStr (strPorta_Com_Fscl)
       intReg = F_VendaItem("000000000", _
       "CANCELANDO CUPOM -ERRO-", "NN", 1, 1)
       FechaPorta
       
    End If
    
    strCod_Mer = F_ECFInformacoesFiscais
    If InStr(1, strCod_Mer, _
    UCase("É possível cancelar o último cupom!")) = 0 Then
        Beep
        MsgBox "Não é possível cancelar o cupom fiscal!", vbExclamation
        P_Status ""
        Exit Sub
    End If
    
    
    IniPortaStr (strPorta_Com_Fscl)
    'Cancela o cupom fiscal
    F_ComandoECF (Chr(27) & "|14|" & Chr(27))
    FechaPorta
    
    'Pegando o número do cupom fiscal
    P_Status "Cancelando registro de cupom fiscal... "
    strNota_Sai = Format(val(F_NumeroCupom) - 2, "0#####")
    
    
    'Verificando se a nota fiscal foi registrada
    strSQL = "Select *  "
    strSQL = strSQL & "From TAB_itsai Where "
    strSQL = strSQL & "Nota_Sai = '" & strNota_Sai & "' And Serie_Sai = 'CUPM' "
    strSQL = strSQL & " and cod_Cai='" & strCod_Cai & "' "
    
    Set rcrdstVerifica = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    sNum_Ven = F_BuscaUmCampo("num_ven", "tab_saidas", "nota_sai='" & _
        strNota_Sai & "' and serie_Sai='CUPM'")
        
    With rcrdstVerifica
        P_Status "Atualizando mercadorias..."
        If .RecordCount > 0 Then
        'O cupom fiscal foi gravado! no arquivo
        'Apagar os registros de itens referente a nota fiscal
        'Recolocando as mercadorias novamente em estoque
                
            
            Do While Not .EOF
               
               
               'Atualizando a tabela de mercadorias
               strSQL = "Update TAB_produtos Set er_pro = er_pro + "
               strSQL = strSQL & Str(!qtd_its) & " Where "
               strSQL = strSQL & "Cod_pro = '" & !cod_pro & "' "
               
               
               Banco.Execute strSQL
               
               
               'Removendo o item do cupom fiscal lançada!
'               strSQL = "Delete from TAB_ITENSSAIDA01 Where Nota_Sai = '"
'               strSQL = strSQL & strNota_Sai & "' And Serie_Sai = "
'               strSQL = strSQL & "'CUPM' And Cod_Mer = '" & !Cod_Mer
'               strSQL = strSQL & "' And Marca_Mer = '" & !MARCA_MER & "'"
'
'               Banco.Execute strSQL
'
               .MoveNext 'Processar o próximo item
'
            Loop
            'Atualizando Tabela de Notas de Saída
            'Passa o campo operacao_sai como CANC indicando que
            'o cupom fiscal foi cancelado
            strSQL = "Update TAB_SAIDAS Set serie_Sai = 'CANC' "
            strSQL = strSQL & "Where Nota_Sai = '" & strNota_Sai & "' "
            strSQL = strSQL & "And Serie_Sai = 'CUPM'"
            strSQL = strSQL & " and cod_Cai='" & strCod_Cai & "' "
            
            Banco.Execute strSQL
                        
            'Removendo da Tabela de ICMS
'            strSQL = "Delete From TAB_ICMS Where Nota_Sai = '"
'            strSQL = strSQL & strNota_Sai & "' And Serie_Sai = 'CUPM' "
            
'            Banco.Execute strSQL
        End If
        .Close
    End With
    strSQL = F_BuscaUmCampo("num_ven", "tab_saidas", "nota_sai='" & strNota_Sai & "'")
    If Not F_CancelaVendaAposCupom(strSQL) Then
        MsgBox "Falha ao cancelar a venda!", vbCritical
    End If
    
    P_Status
    'rcrdstVerifica.Close
    MsgBox "Cupom Cancelado! Venda Cancelada!", vbInformation
    
End Sub

Private Sub cmndbtnConfiguracoes_Click()
    txtbxVisualizacao.Text = F_ECFInformacoesFiscais & _
     "SÍMBOLO DA MOEDA: " & F_ECFSimboloMoeda & "$"
End Sub

Private Sub cmndbtnEstadoImpressora_Click()
    txtbxVisualizacao.Text = F_ECFStatus
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub cmndbtnHorarioVerao_Click()
    
       F_ECFHorarioVerao
    
End Sub

Private Sub cmndbtnLeituraMemoriaFiscal_Click()
   'Alerta sobre o que vai acontecer
   Beep
   If MsgBox("Você está prestes a emitir um Relatório da Memória Fiscal!" & _
   Chr(13) & Chr(10) & "Tem Certeza disto?", vbQuestion & vbYesNo) = vbYes Then
      'Emite a leitura da memória fiscal!!!
      P_LeituraMemoriaFiscal "I"
   End If
   Exit Sub
End Sub

Private Sub cmndbtnLeituraX_Click()
   Beep
   If MsgBox("Você está prestes a emitir um Relatório da Leitura X!!" & _
   Chr(13) & Chr(10) & "Tem Certeza disto?", vbQuestion & vbYesNo) = vbYes Then
      'Emite a leitura Xl!!!
      P_EmiteLeituraX "I"
   End If
End Sub

Private Sub cmndbtnReducaoZ_Click()
   'Verifica a senha para o acesso
   If MsgBox("Você está prestes a emitir um Relatório da Redução Z!!" & _
   Chr(13) & Chr(10) & "Não será possível efetuar mais vendas!" & _
   Chr(13) & Chr(10) & "Tem Certeza disto?", vbQuestion & vbYesNo) <> vbYes Then
      'Sai do procedimento
      Exit Sub
   End If
   
   'Pede a senha para a operação
   ' If Not F_Senha("SIT" & Date) Then
     '  Exit Sub
   ' End If
   If F_HouveReducaoZData Then
      'Já houve Redução Z
        If MsgBox("Já houve Redução Z nesta data?", vbYesNo + vbQuestion) = vbYes Then
            Exit Sub
        End If
   End If
   If Not F_ReducaoZ Then
      MsgBox "Não consigo fechar a Redução Z"
   End If
End Sub

Private Sub cmndbtnVisualizarLeituraX_Click()
   P_EmiteLeituraX
End Sub

Private Sub cmndbtnVisualizarMemoriaFiscal_Click()
   'Faz a leitura da memória fiscal
   P_LeituraMemoriaFiscal
End Sub


Function F_RetornoArquivoECF() As String
    'Retorna o Conteúdo do Arquivo MP20FI.RET
    ''On Error  GoTo TrataErro
    
    F_RetornoArquivoECF = ""
    If Dir("Mp20Fi.Ret") = "" Then
       'Não gerou o arquivo
       Exit Function
    End If
    
    Dim strResto As String, strExpr As String
    
    'Abrindo arquivo
    strResto = ""
    txtbxVisualizacao.Text = "" 'Apagando caixa de texto
    
    Open "Mp20Fi.Ret" For Input As #1
    Do While Not EOF(1)
       Line Input #1, strExpr
       F_RetornoArquivoECF = F_RetornoArquivoECF & _
       strResto & Left(F_ConvDOS(strExpr), 48) & Chr(13) & Chr(10)
       strResto = Mid(F_ConvDOS(strExpr), 49)
    Loop
    Close #1
    'Apaga o arquivo
    Kill "Mp20Fi.Ret"
    Exit Function
    
TrataErro:
    P_MostraErro
End Function

Private Sub Form_Activate()
    txtbxMsg.Text = F_BuscaUmCampo("msgecf_par", "tab_Parametros", "")
    If F_BuscaUmCampo("TIPOmsgecf_par", "tab_Parametros", "") = "1" Then
        optnCumpr(0).Value = True
    Else
        optnCumpr(1).Value = True
    
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    sstb.Tab = 0
End Sub

Private Sub optnCumpr_Click(Index As Integer)
    Dim strSQL As String
    
    On Error GoTo trata_erro
    strSQL = "update tab_parametros set tipomsgecf_par='" & (Index + 1) & "' "
    Banco.Execute strSQL
    If Index = 0 Then
        txtbxMsg.Enabled = False
    Else
        txtbxMsg.Enabled = True
    End If
        Exit Sub
    
trata_erro:
    MsgBox Err.Description, vbCritical
    
    
End Sub

Private Sub sstb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub txtbxMsg_LostFocus()
    Dim strSQL As String
    
    On Error GoTo trata_erro
    strSQL = "update tab_parametros set msgecf_par='" & txtbxMsg.Text & "' "
    Banco.Execute strSQL
        Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
