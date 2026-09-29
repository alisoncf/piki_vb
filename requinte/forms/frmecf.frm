VERSION 5.00
Begin VB.Form frmECF 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Impressora Fiscal - "
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11280
   Icon            =   "frmecf.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnLimpaIntpos 
      Caption         =   "Limpa IntPos"
      Height          =   615
      Left            =   120
      TabIndex        =   27
      Top             =   7320
      Width           =   2055
   End
   Begin VB.PictureBox Picture4 
      Height          =   3735
      Left            =   120
      ScaleHeight     =   3675
      ScaleWidth      =   4395
      TabIndex        =   16
      Top             =   3240
      Width           =   4455
      Begin VB.CommandButton cmndbtnVerificaTotalizadoresNF 
         Caption         =   "Busca Totalizadores Não Fiscais"
         Height          =   615
         Left            =   120
         TabIndex        =   26
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnLinhasEntreCupons 
         Caption         =   "Linhas Entre Cupons"
         Height          =   615
         Left            =   2280
         TabIndex        =   25
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtbxComando 
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   3360
         Width           =   4215
      End
      Begin VB.CommandButton cmndbtnEspacoEntreLinhas 
         Caption         =   "Espaço Entre Linhas"
         Height          =   615
         Left            =   2280
         TabIndex        =   22
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnInfo 
         Caption         =   "&Informações da Impressora"
         Height          =   615
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Imagem de Leitura X"
         Height          =   615
         Left            =   2280
         TabIndex        =   20
         Top             =   120
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Imagem de Memória Fiscal"
         Height          =   615
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnEstado 
         Caption         =   "&Estado da Impressora"
         Height          =   615
         Left            =   2280
         TabIndex        =   18
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnMapaResumo 
         Caption         =   "M&apa Resumo"
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comando:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   3120
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   6840
      ScaleHeight     =   1545
      ScaleWidth      =   4305
      TabIndex        =   13
      Top             =   1440
      Width           =   4335
      Begin VB.CommandButton cmndbtnProgramaAliquota 
         Caption         =   "&Programa Alíquota"
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   2055
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Nomeia Totalizador Não Fiscal"
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00578536&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1545
      ScaleWidth      =   6585
      TabIndex        =   4
      Top             =   1440
      Width           =   6615
      Begin VB.CommandButton cmndbtnCancCup 
         Caption         =   "&Cancela Cupom"
         Height          =   615
         Left            =   2280
         TabIndex        =   10
         Top             =   120
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnReset 
         Caption         =   "&Reseta Impressora"
         Enabled         =   0   'False
         Height          =   615
         Left            =   4440
         TabIndex        =   9
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnHorarioVerao 
         Caption         =   "&Horário de Verão"
         Height          =   615
         Left            =   2280
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnRedZ 
         Caption         =   "&Redução Z"
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnMemFiscal 
         Caption         =   "&Memória Fiscal"
         Height          =   615
         Left            =   4440
         TabIndex        =   6
         Top             =   120
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnLeitX 
         Caption         =   "&Leitura X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Sair (ESC)"
      Height          =   615
      Left            =   9120
      TabIndex        =   3
      Top             =   7200
      Width           =   1935
   End
   Begin VB.TextBox txtbxInfo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   4680
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3240
      Width           =   6495
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   0
      Picture         =   "frmecf.frx":030A
      ScaleHeight     =   1290
      ScaleWidth      =   11280
      TabIndex        =   0
      Top             =   0
      Width           =   11280
      Begin VB.Label lblImprFiscal 
         AutoSize        =   -1  'True
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3120
         TabIndex        =   12
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Impressora ="
         Height          =   195
         Left            =   2160
         TabIndex        =   11
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Utilitário de Impressora Fiscal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   5400
      End
   End
End
Attribute VB_Name = "frmECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnCancCup_Click()
        
   cancela_Cupom
   'frmCancelaCupom.Show 1
   ' Exit Sub
   
End Sub

Private Sub cmndbtnEspacoEntreLinhas_Click()
    Dim i As Integer
    i = Bematech_FI_EspacoEntreLinhas(CInt(txtbxComando.Text))
    If i = 0 Then
        MsgBox "Erro de comunicação!", vbExclamation
    ElseIf i = 1 Then
        MsgBox "Ok!", vbInformation
    ElseIf i = -2 Then
        MsgBox "Parâmetro inválido na função.", vbExclamation
    ElseIf i = -4 Then
        MsgBox "O arquivo de inicialização BemaFI32.ini não foi encontrado no diretório de sistema do Windows.", vbExclamation
        
    ElseIf i = -5 Then
        MsgBox "Erro ao abrir a porta de comunicação. ", vbExclamation

    ElseIf i = -27 Then
        MsgBox "Status da impressora diferente de 6,0,0 (ACK, ST1 e ST2). ", vbExclamation

    ElseIf i = 30 Then
        MsgBox "Função não compatível com a impressora YANCO. ", vbExclamation
        
    End If
    
    
End Sub

Private Sub cmndbtnEstado_Click()
    Dim rImp As retorno_impressora
    rImp = ecf_retornoImprComandos()
    
    txtbxInfo.Text = rImp.msg
    
End Sub

Private Sub cmndbtnHorarioVerao_Click()
    Dim strExpr As String
    strExpr = "A impressora fiscal está funcionando no Horário "
    strExpr = strExpr & IIf(blnVerao, "de Verão!", "Normal!") & Chr(13)
    strExpr = strExpr & "Para poder continuar a operação:" & Chr(13)
    strExpr = strExpr & "Não pode ter ocorrido movimento no dia;" & Chr(13)
    strExpr = strExpr & "Ou logo após a Redução Z, se estiver "
    strExpr = strExpr & "saindo do horário normal;" & Chr(13)
    strExpr = strExpr & "Ou Após uma hora efetuada a Redução Z, "
    strExpr = strExpr & "se estiver saindo do horário de verão!"
    strExpr = strExpr & Chr(13) & Chr(13)
    strExpr = strExpr & F_Centra("D E S E J A   C O N T I N U A R ? ", 80)
    If MsgBox(strExpr, vbYesNo + vbCritical) = vbYes Then
        ecf_HorarioVerao
    End If
End Sub

Private Sub Info_Bema()
    Dim strInfo As String
    Dim i As Integer
    Dim j As Long
    Dim strAux As String
    Dim strAux2 As String
    Dim CrLf As String
    CrLf = Chr(13) + Chr(10)
    strAux = "               "
    strAux2 = "               "
    i = Bematech_FI_NumeroSerie(strAux)
    strInfo = "Nº Série: " & strAux & Chr(9)
    i = Bematech_FI_CGC_IE(strAux, strAux2)
    strInfo = strInfo & CrLf + "CNPJ: " & strAux & " - IE: " & strAux2
    i = Bematech_FI_VersaoFirmware(strAux)
    strInfo = strInfo + CrLf + "Versão Firmware: " & strAux
    strAux = "               "
    i = Bematech_FI_NumeroIntervencoes(strAux)
    
    strInfo = strInfo + CrLf + "Número Intervenções Técnicas: " & strAux
    strAux = "               "
    i = Bematech_FI_NumeroReducoes(strAux)
    strInfo = strInfo + CrLf + "Número de Reduções: " & strAux
    strAux = "               "
    i = Bematech_FI_NumeroSubstituicoesProprietario(strAux)
    
    strInfo = strInfo + CrLf + "Número de Mudanças Proprietários: " & strAux
    
    i = Bematech_FI_MinutosImprimindo(strAux)
    strInfo = strInfo + CrLf + "Tempo Imprimindo (minutos): " & strAux
    i = Bematech_FI_MinutosLigada(strAux)
    strInfo = strInfo + CrLf + "Tempo Ligada (minutos): " & strAux
    
    txtbxInfo.Text = ""
    For i = 1 To Len(strInfo)
        txtbxInfo.Text = txtbxInfo.Text & Mid(strInfo, i, 1)
    Next
    
End Sub
Private Sub Info_Elgin()
    Dim strInfo As String
    Dim i As Integer
    Dim j As Long
    Dim strAux As String
    Dim strAux2 As String
    Dim CrLf As String
    CrLf = Chr(13) + Chr(10)
    strAux = "               "
    strAux2 = "               "
    i = Elgin_NumeroSerie(strAux)
    strInfo = "Nº Série: " & strAux & Chr(9)
    i = Elgin_CGC_IE(strAux, strAux2)
    strInfo = strInfo & CrLf + "CNPJ: " & strAux & " - IE: " & strAux2
    i = Elgin_VersaoFirmware(strAux)
    strInfo = strInfo + CrLf + "Versão Firmware: " & strAux
    strAux = "               "
    i = Elgin_NumeroIntervencoes(strAux)
    
    strInfo = strInfo + CrLf + "Número Intervenções Técnicas: " & strAux
    strAux = "               "
    i = Elgin_NumeroReducoes(strAux)
    strInfo = strInfo + CrLf + "Número de Reduções: " & strAux
    strAux = "               "
    i = Elgin_NumeroSubstituicoesProprietario(strAux)
    
    strInfo = strInfo + CrLf + "Número de Mudanças Proprietários: " & strAux
    
    i = Elgin_MinutosImprimindo(strAux)
    strInfo = strInfo + CrLf + "Tempo Imprimindo (minutos): " & strAux
    i = Elgin_MinutosLigada(strAux)
    strInfo = strInfo + CrLf + "Tempo Ligada (minutos): " & strAux
    
    txtbxInfo.Text = ""
    For i = 1 To Len(strInfo)
        txtbxInfo.Text = txtbxInfo.Text & Mid(strInfo, i, 1)
    Next
    
End Sub

Private Sub Info_Daruma()
    Dim strInfo As String
    Dim i As Integer
    Dim j As Long
    Dim strAux As String
    Dim strAux2 As String
    Dim CrLf As String
    CrLf = Chr(13) + Chr(10)
    strAux = "               "
    strAux2 = "               "
    i = Daruma_FI_NumeroSerie(strAux)
    strInfo = "Nº Série: " & strAux & Chr(9)
    i = Daruma_FI_CGC_IE(strAux, strAux2)
    strInfo = strInfo & CrLf + "CNPJ: " & strAux & " - IE: " & strAux2
    i = Daruma_FI_VersaoFirmware(strAux)
    strInfo = strInfo + CrLf + "Versão Firmware: " & strAux
    strAux = "               "
    i = Daruma_FI_NumeroIntervencoes(strAux)
    
    strInfo = strInfo + CrLf + "Número Intervenções Técnicas: " & strAux
    strAux = "               "
    i = Daruma_FI_NumeroReducoes(strAux)
    strInfo = strInfo + CrLf + "Número de Reduções: " & strAux
    strAux = "               "
    i = Daruma_FI_NumeroSubstituicoesProprietario(strAux)
    
    strInfo = strInfo + CrLf + "Número de Mudanças Proprietários: " & strAux
    
    i = Daruma_FI_MinutosImprimindo(strAux)
    strInfo = strInfo + CrLf + "Tempo Imprimindo (minutos): " & strAux
    i = Daruma_FI_MinutosLigada(strAux)
    strInfo = strInfo + CrLf + "Tempo Ligada (minutos): " & strAux
    
    txtbxInfo.Text = ""
    For i = 1 To Len(strInfo)
        txtbxInfo.Text = txtbxInfo.Text & Mid(strInfo, i, 1)
    Next
    
End Sub

Private Sub cmndbtnInfo_Click()
    If strImprFiscal = "D" Then
        Info_Daruma
    ElseIf strImprFiscal = "E" Then
        Info_Elgin
    Else
        Info_Bema
    End If
End Sub

Private Sub cmndbtnLeitX_Click()
    
    
    ecf_LeituraX
End Sub

Private Sub cmndbtnLimpaIntpos_Click()
    killIntPosResp
    
End Sub

Private Sub cmndbtnLinhasEntreCupons_Click()
     Dim i As Integer
    i = Bematech_FI_LinhasEntreCupons(CInt(txtbxComando.Text))
    If i = 0 Then
        MsgBox "Erro de comunicação!", vbExclamation
    ElseIf i = 1 Then
        MsgBox "Ok!", vbInformation
    ElseIf i = -2 Then
        MsgBox "Parâmetro inválido na função.", vbExclamation
    ElseIf i = -4 Then
        MsgBox "O arquivo de inicialização BemaFI32.ini não foi encontrado no diretório de sistema do Windows.", vbExclamation
        
    ElseIf i = -5 Then
        MsgBox "Erro ao abrir a porta de comunicação. ", vbExclamation

    ElseIf i = -27 Then
        MsgBox "Status da impressora diferente de 6,0,0 (ACK, ST1 e ST2). ", vbExclamation

    ElseIf i = 30 Then
        MsgBox "Função não compatível com a impressora YANCO. ", vbExclamation
        
    End If
End Sub

Private Sub cmndbtnMapaResumo_Click()
    
    ecf_MapaResumo
End Sub

Private Sub Command9_Click()

End Sub

Private Sub cmndbtnMemFiscal_Click()
    Dim sDataIni As String
    Dim sDataFin As String
    Dim s As String
    s = F_DataInicialFinal()
    
        
    ecf_MemoriaFiscal Left(s, 10), Right(s, 10)
    
    
    
    
End Sub

Private Sub cmndbtnProgramaAliquota_Click()
    frmUtilitarioECF.Show 1
End Sub

Private Sub cmndbtnRedZ_Click()
    Dim strSQL As String
    Dim i As Integer
    strSQL = "A Redução Z é a leitura de Fechamento que deve ser feita no final do Dia." & Chr(13)
    strSQL = strSQL & "Uma vez feita a redução Z, até às 23:59:00hs "
    strSQL = strSQL & "a impressora estará bloqueada." & Chr(13)
    strSQL = strSQL & "Deseja realmente emitir a redução Z?"
    If MsgBox(strSQL, vbCritical + vbYesNo) = vbYes Then
        ecf_ReducaoZ
    End If
End Sub

Private Sub cmndbtnReset_Click()
    ecf_ResetaImpressora
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    P_PortaComunicacaoECF
End Sub


Private Sub cmndbtnVerificaTotalizadoresNF_Click()
    Dim i As Integer
    Dim sTot As String
    sTot = Space(179)
    i = Bematech_FI_VerificaTotalizadoresNaoFiscais(sTot)
    If i = 0 Then
        MsgBox "Erro de comunicação!", vbExclamation
    ElseIf i = 1 Then
        MsgBox "Ok!", vbInformation
    ElseIf i = -2 Then
        MsgBox "Parâmetro inválido na função.", vbExclamation
    ElseIf i = -4 Then
        MsgBox "O arquivo de inicialização BemaFI32.ini não foi encontrado no diretório de sistema do Windows.", vbExclamation
        
    ElseIf i = -5 Then
        MsgBox "Erro ao abrir a porta de comunicação. ", vbExclamation

    ElseIf i = -27 Then
        MsgBox "Status da impressora diferente de 6,0,0 (ACK, ST1 e ST2). ", vbExclamation

        
    End If
    txtbxInfo.Text = sTot
End Sub

Private Sub Command10_Click()
    frmTotalizadorNF.Show 1
End Sub

Private Sub Form_Load()
    lblImprFiscal = strImprFiscal
End Sub
