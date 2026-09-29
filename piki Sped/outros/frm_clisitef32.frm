VERSION 5.00
Begin VB.Form FrmSiTef 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "API TEF Dedicado"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmr 
      Interval        =   1000
      Left            =   2160
      Top             =   5400
   End
   Begin VB.PictureBox pcFaixa 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EF9C7C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   9420
      TabIndex        =   5
      Top             =   7740
      Width           =   9420
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   7800
         ScaleHeight     =   495
         ScaleWidth      =   1455
         TabIndex        =   16
         Top             =   120
         Width           =   1455
         Begin VB.PictureBox pcIco 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   120
            Picture         =   "frm_clisitef32.frx":0000
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   17
            Top             =   120
            Width           =   195
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sair"
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   18
            Top             =   120
            Width           =   270
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   3360
         ScaleHeight     =   495
         ScaleWidth      =   1455
         TabIndex        =   13
         Top             =   120
         Width           =   1455
         Begin VB.PictureBox pcIco 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   120
            Picture         =   "frm_clisitef32.frx":016B
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   14
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preferências"
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   15
            Top             =   120
            Width           =   885
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   1800
         ScaleHeight     =   495
         ScaleWidth      =   1455
         TabIndex        =   10
         Top             =   120
         Width           =   1455
         Begin VB.PictureBox pcIco 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   120
            Picture         =   "frm_clisitef32.frx":0B6D
            ScaleHeight     =   210
            ScaleWidth      =   255
            TabIndex        =   11
            Top             =   120
            Width           =   255
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gerencial"
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   12
            Top             =   120
            Width           =   675
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   1455
         TabIndex        =   7
         Top             =   120
         Width           =   1455
         Begin VB.PictureBox pcIco 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   120
            Picture         =   "frm_clisitef32.frx":0D06
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   8
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Conecta"
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   9
            Top             =   120
            Width           =   600
         End
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   495
         Left            =   40
         TabIndex        =   6
         Top             =   7200
         Width           =   1695
      End
   End
   Begin VB.PictureBox pcStatus 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   9390
      TabIndex        =   3
      Top             =   8535
      Width           =   9420
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3240
         TabIndex        =   4
         Top             =   0
         Width           =   180
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8160
      Left            =   0
      Picture         =   "frm_clisitef32.frx":1090
      ScaleHeight     =   8160
      ScaleWidth      =   9345
      TabIndex        =   1
      Top             =   0
      Width           =   9345
      Begin VB.Label lblVisorOp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "mensagem"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3945
         Left            =   600
         TabIndex        =   2
         Top             =   720
         Width           =   8055
      End
   End
   Begin VB.TextBox TxtDisplay 
      Height          =   495
      Left            =   4320
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "FrmSiTef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub cmndbtnGerencial_Click()
    Dim Retorno        As Long
    Dim Buffer         As String * 20000
    Dim proximocomando As Long
    Dim TipoCampo      As Long
    Dim TamanhoMinimo  As Integer
    Dim TamanhoMaximo  As Integer
    Dim ContinuaNavegacao  As Long
    Dim i As Integer
    Dim sCab As String
    Command1_Click
    Retorno = IniciaFuncaoSiTefInterativo(110, 0, "", CStr(Date), "", "", "")
    
    Do While Retorno = 10000
        
        Retorno = ContinuaFuncaoSiTefInterativo(proximocomando, TipoCampo, TamanhoMinimo, TamanhoMaximo, Buffer, Len(Buffer), ContinuaNavegacao)
        
        P_TrataProximoComando proximocomando, Buffer
        
        
        
    Loop
    
    If TipoCampo = 121 Then
        ecf_LeituraX
        i = Bematech_FI_RelatorioGerencial(Buffer)
        
    End If
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Configura "127.0.0.1"
End Sub

Public Sub Configura(ByVal IP As Variant)

  Dim Retorno        As Long

  Screen.MousePointer = vbHourglass

  Retorno = ConfiguraIntSiTefInterativo(IP & Chr(0), "00000000" & Chr(0), "SE000001" & Chr(0), 0)
  Screen.MousePointer = vbDefault

  If (Retorno = 0) Then
    FrmSiTef.lblVisorOp.caption = "Sitef Ok"
  Else
    MsgBox "Erro: Retorno -> " & CStr(Retorno)
  End If

End Sub

Public Function EfetuaPagamento(Valor As String, cupom As String, atendente As String) As Boolean

  Dim Retorno        As Long
  Dim Buffer         As String * 20000
  Dim proximocomando As Long
  Dim TipoCampo      As Long
  Dim TamanhoMinimo  As Integer
  Dim TamanhoMaximo  As Integer
  Dim ContinuaNavegacao  As Long
  Dim sCab As String
  Dim i As Integer
  Dim c As New Collection
  Dim iCount As Integer
  Command1_Click
  
  Screen.MousePointer = vbHourglass
  Retorno = IniciaFuncaoSiTefInterativo(0, Format(Valor, "###,##0.00") & Chr(0), cupom & Chr(0), Format(Date, "yyyymmdd") & Chr(0), Format(Time, "hhmmss") & Chr(0), atendente & Chr(0), Chr(0))
  Screen.MousePointer = vbDefault
    
  proximocomando = 0
  TipoCampo = 0
  TamanhoMinimo = 0
  TamanhoMaximo = 0
  ContinuaNavegacao = 0
  Resultado = 0
  Buffer = String(20000, 0)

  Do

    Screen.MousePointer = vbHourglass
    Buffer = Buffer & Chr(0)
    
    Retorno = ContinuaFuncaoSiTefInterativo(proximocomando, TipoCampo, TamanhoMinimo, TamanhoMaximo, Buffer, Len(Buffer), Resultado)
    Screen.MousePointer = vbDefault

    If (Retorno = 10000) Then

      FrmSiTef.TxtDisplay.Text = Buffer
        
        P_TrataProximoComando proximocomando, Buffer
    End If

  Loop Until Not (Retorno = 10000)

  If (Retorno = 0) Then
    ecf_TerminaFechamentoCupom "Obrigado! Volte Sempre!"
    Set c = F_SeparaTexto(Buffer, Chr(10), False)
    iCount = ecf_abreComprovanteNaoFinalVinculado("Cartao", Format(Valor, "###,##0.00"), strNum_Cpm)
    For i = 1 To c.count
        iCount = ecf_UsaComprovanteNaoFinalVinculado(c(i))
    Next
    ecf_FechaComprovanteNaoFinalVinculado
  
      
    'MsgBox "Retorno Ok!"
    lblMsg.caption = ""
    P_Msg "Aguardando Transação TEF"
    EfetuaPagamento = True
  Else
    MsgBox "Erro:" & Buffer & " Retorno -> " & CStr(Retorno)
    EfetuaPagamento = False
  End If

End Function

Sub ColetaCampo(ByVal Comando As Integer, _
                ByVal TipoCampo As Long, _
                ByVal TamanhoMinimo As Integer, _
                ByVal TamanhoMaximo As Integer, _
                ByRef pDadosComando As String, _
                ByRef pCampo As String)

  Dim strComando       As String
  Dim strTipoCampo     As String
  Dim strTamanhoMinimo As String
  Dim strTamanhoMaximo As String

  DoEvents

  strComando = "Comando      :" & CStr(Comando)
  strTipoCampo = "TipoCampo    :" & CStr(TipoCampo)
  strTamanhoMinimo = "TamanhoMinimo:" & CStr(TamanhoMinimo)
  strTamanhoMaximo = "TamanhoMaximo:" & CStr(TamanhoMaximo)

End Sub

Sub RecebeResultado(ByVal TipoCampo As Long, _
                    ByRef pResultado As String)

  MsgBox "TipoCampo :" & CStr(TipoCampo)
  MsgBox "pResultado:" & CStr(pResultado)

End Sub


Function efetua_venda(seq_tra As String, num_ven As String, cupom_tra As String) As Boolean
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim i As Integer
        
    strSQL = "update tab_transacoes set status_tra=2"
    strSQL = strSQL & " where seq_tra=" & seq_tra
    Banco.Execute strSQL
    
    
    strSQL = "select tv.*,ti.* from tab_vendas as tv,tab_itvendas as ti "
    strSQL = strSQL & " where tv.num_ven=ti.num_ven "
    strSQL = strSQL & " and tv.num_ven='" & num_ven & "' "
    strSQL = strSQL & " order by seq_ite "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Não foi encontrada a venda, avise o responsável!", vbExclamation
            efetua_venda = False
            Exit Function
        End If
        
        If cupom_tra = "" Then
            P_Status "Abrindo cupom fiscal"
            ecf_AbreCupom !cpf_cli
            P_Status "Buscando número do cupom fiscal"
            i = ecf_NumeroCupom(strNum_Cpm)
            
        Else
            P_Status "Cupom fiscal número " & cupom_tra
            strNum_Cpm = cupom_tra
            i = 1
        End If
        If i = 0 Then
            MsgBox "Não foi possível encontrar o numero do cupom!", vbExclamation
            P_Status ""
            Exit Function
        End If
    
        If cupom_tra = "" Then
            Do While Not .EOF
                P_Status "Imprimindo item " & !nome_ite
                i = ecf_VendeItem(!cod_pro, !nome_ite, !aliquota_ite, _
                    !tipoqtd_ite, !qtd_ite, 3, Format(!unit_ite, ""), _
                    !tipodesc_ite, Format(!desconto_ite, "###,##0.00"))
                
                If F_TrataRetornoVendeItem(i) = False Then
                    Exit Function
                End If
                .MoveNext
            Loop
            .MoveFirst
            
        End If
        P_Status "Iniciando fechamento do Cupom"
        i = ecf_IniciaFechamentoCupom("D", "$", !desconto_ven) 'desconto sempre em dinheiro
        
        P_Status "Efetuando pagamento cartão"
        
        ecf_EfetuaFormaPagamento "Cartao", Format(!cartao_ven, "###,##0.00")
        
        P_Status "Iniciando Transação TEF"
        
        If EfetuaPagamento(!cartao_ven, strNum_Cpm, !atendente_ven) = True Then
            efetua_venda = True
        Else
            efetua_venda = False
            Exit Function
        End If
        
        If !avista_ven > 0 Then
            P_Status "Efetuando pagamento A Vista"
            ecf_EfetuaFormaPagamento "Dinheiro", Format(!avista_ven, "###,##0.00")
        End If
        If !aprazo_ven > 0 Then
            P_Status "Efetuando pagamento A Prazo"
            ecf_EfetuaFormaPagamento "A Prazo", Format(!aprazo_ven, "###,##0.00")
        End If
        
        
        
        'i = ecf_abreComprovanteNaoFinalVinculado
        P_Status "Terminando fechamento do CUPOM"
        
        ecf_TerminaFechamentoCupom "Obrigado!Volte Sempre!"
        
        P_Status ""
        .Close
    End With
    efetua_venda = True
    
End Function

Private Sub lblBotao_Click(Index As Integer)
    pcBotao_Click Index
End Sub

Private Sub pcBotao_Click(Index As Integer)
    pcBotao_GotFocus Index
    If Index = 0 Then
        Command1_Click
    ElseIf Index = 1 Then
        cmndbtnGerencial_Click
    ElseIf Index = 2 Then
        
    ElseIf Index = 3 Then
        cmndbtnSair_Click
    
    
        
    End If
    
End Sub
Sub P_foco(Indice As Integer)
    Dim iCount As Integer
    For iCount = 0 To pcBotao.count - 1
        pcBotao(iCount).BackColor = vbWhite
        
    Next
    pcBotao(Indice).BackColor = &HECE3D0
End Sub
Sub P_focoMouse(Indice As Integer)
    Dim iCount As Integer
    For iCount = 0 To pcBotao.count - 1
        pcBotao(iCount).BorderStyle = 0
        
    Next
    pcBotao(Indice).BorderStyle = 1
    pcBotao(Indice).SetFocus
End Sub

Private Sub pcBotao_GotFocus(Index As Integer)
    P_foco Index
End Sub

Private Sub pcBotao_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    
    If KeyCode = vbKeyLeft Then
        If Index = 0 Then
            pcBotao(pcBotao.count - 1).SetFocus
        Else
            pcBotao(Index - 1).SetFocus
        End If
        
    ElseIf KeyCode = vbKeyRight Then
        If Index = pcBotao.count - 1 Then
            pcBotao(0).SetFocus
        Else
            pcBotao(Index + 1).SetFocus
        End If
    ElseIf KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        pcBotao_Click Index
    End If
    
End Sub

Private Sub pcBotao_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    P_focoMouse Index
End Sub

Private Sub pcIco_Click(Index As Integer)
    pcBotao_Click Index
End Sub

Private Sub pcIco_GotFocus(Index As Integer)
    SendKeys Chr(9)
End Sub
Sub busca_transacao()
    Dim strSQL As String
    Dim rPesq As Recordset
    P_Status "Buscando transação"
    
    strSQL = "select * from tab_transacoes where status_tra=1 order by seq_tra desc limit 1"
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount > 0 Then
            If !tipo_tra = "1" Then 'venda
                If efetua_venda(!seq_tra, !num_ven, !cupom_tra) = True Then
                    strSQL = "update tab_transacoes set status_tra=3"
                    strSQL = strSQL & " where seq_tra=" & !seq_tra
                    Banco.Execute strSQL
                    P_Fim
                Else
                    strSQL = "update tab_transacoes set status_tra=5"
                    strSQL = strSQL & " where seq_tra=" & !seq_tra
                    Banco.Execute strSQL
                    P_Fim
                End If
            
            End If
            
            
        Else
            lblMsg.caption = "Aguardando Transação TEF"
        End If
       
        .Close
    End With
    P_Status ""
End Sub

Private Sub Command2_Click()

End Sub

Private Sub tmr_Timer()
    
'    busca_transacao
    
End Sub
