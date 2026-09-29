VERSION 5.00
Begin VB.Form frmECF 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Impressora Fiscal - "
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8415
   Icon            =   "frmECF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Avançado:"
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   5055
      Begin VB.CommandButton Command10 
         Caption         =   "&Nomeia Tot. Não Fiscal"
         Height          =   495
         Left            =   2640
         TabIndex        =   9
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnProgramaAliquota 
         Caption         =   "&Programa Alíquota"
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Diversos:"
      Height          =   3615
      Left            =   5280
      TabIndex        =   10
      Top             =   2640
      Width           =   3015
      Begin VB.CommandButton cmndbtnMapaResumo 
         Caption         =   "M&apa Resumo"
         Height          =   495
         Left            =   480
         TabIndex        =   15
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Estado da Impressora"
         Height          =   495
         Left            =   480
         TabIndex        =   14
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Imagem de Memória Fiscal"
         Height          =   495
         Left            =   480
         TabIndex        =   13
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Imagem de Leitura X"
         Height          =   495
         Left            =   480
         TabIndex        =   12
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnInfo 
         Caption         =   "&Informações da Impressora"
         Height          =   495
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Serial:"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   5055
      Begin VB.CommandButton cmndbtnReset 
         Caption         =   "&Reseta Impressora"
         Height          =   495
         Left            =   2640
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnRedZ 
         Caption         =   "&Redução Z"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnLeitX 
         Caption         =   "&Leitura X"
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnMemFiscal 
         Caption         =   "&Memória Fiscal"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   795
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnHorarioVerao 
         Caption         =   "&Horário de Verão"
         Height          =   495
         Left            =   2640
         TabIndex        =   6
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnCancCup 
         Caption         =   "&Cancela Cupom"
         Height          =   495
         Left            =   2640
         TabIndex        =   4
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Sair (ESC)"
      Height          =   375
      Left            =   6720
      TabIndex        =   19
      Top             =   120
      Width           =   1575
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
      Height          =   1695
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   720
      Width           =   5535
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2295
      Left            =   120
      Picture         =   "frmECF.frx":030A
      ScaleHeight     =   2295
      ScaleWidth      =   2340
      TabIndex        =   16
      Top             =   120
      Width           =   2340
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Utilitário de Impressora Fiscal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      TabIndex        =   17
      Top             =   120
      Width           =   3645
   End
   Begin VB.Line Line1 
      X1              =   2760
      X2              =   8160
      Y1              =   600
      Y2              =   600
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
    Else
        Info_Bema
    End If
End Sub

Private Sub cmndbtnLeitX_Click()
    
    
    ecf_LeituraX
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

