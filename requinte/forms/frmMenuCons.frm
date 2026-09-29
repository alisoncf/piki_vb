VERSION 5.00
Begin VB.Form frmMenuCons 
   Appearance      =   0  'Flat
   BackColor       =   &H00755C33&
   BorderStyle     =   0  'None
   Caption         =   "Área de Trabalho"
   ClientHeight    =   9930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Menu de Opções"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   0
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   36
      Top             =   0
      Width           =   1815
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produtos"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   37
         Top             =   600
         Width           =   645
      End
      Begin VB.Image imgIco 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   600
         MouseIcon       =   "frmMenuCons.frx":0000
         Picture         =   "frmMenuCons.frx":164A
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   1
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   34
      Top             =   1080
      Width           =   1815
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clientes"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   510
         TabIndex        =   35
         Top             =   600
         Width           =   585
      End
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   1
         Left            =   600
         MouseIcon       =   "frmMenuCons.frx":2314
         Picture         =   "frmMenuCons.frx":3C96
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   3
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   32
      Top             =   3240
      Width           =   1815
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendas"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   525
         TabIndex        =   33
         Top             =   600
         Width           =   555
      End
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   3
         Left            =   600
         Picture         =   "frmMenuCons.frx":5618
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   4
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   30
      Top             =   4320
      Width           =   1815
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   31
         Top             =   600
         Width           =   405
      End
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   4
         Left            =   600
         Picture         =   "frmMenuCons.frx":5EE2
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   5
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   28
      Top             =   5400
      Width           =   1815
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   525
         TabIndex        =   29
         Top             =   600
         Width           =   555
      End
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   5
         Left            =   600
         Picture         =   "frmMenuCons.frx":6BAC
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   6
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   26
      Top             =   6480
      Width           =   1815
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sair"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   660
         TabIndex        =   27
         Top             =   600
         Width           =   285
      End
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   6
         Left            =   600
         Picture         =   "frmMenuCons.frx":7876
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   2
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   24
      Top             =   2160
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   2
         Left            =   720
         Picture         =   "frmMenuCons.frx":8540
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Representantes"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   345
         TabIndex        =   25
         Top             =   600
         Width           =   1155
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   7
      Left            =   2160
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   22
      Top             =   0
      Width           =   1815
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consignações"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   405
         TabIndex        =   23
         Top             =   600
         Width           =   1035
      End
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   7
         Left            =   720
         Picture         =   "frmMenuCons.frx":920A
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   8
      Left            =   2160
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   20
      Top             =   1080
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   240
         Index           =   8
         Left            =   720
         Picture         =   "frmMenuCons.frx":9AD4
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transferência"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   21
         Top             =   600
         Width           =   1005
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   9
      Left            =   2160
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   18
      Top             =   2160
      Width           =   1815
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Devolução"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   9
         Left            =   465
         TabIndex        =   19
         Top             =   600
         Width           =   795
      End
      Begin VB.Image imgIco 
         Height          =   240
         Index           =   9
         Left            =   720
         Picture         =   "frmMenuCons.frx":A05E
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   10
      Left            =   2160
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   16
      Top             =   3240
      Width           =   1815
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acerto"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   10
         Left            =   615
         TabIndex        =   17
         Top             =   600
         Width           =   495
      End
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   10
         Left            =   600
         Picture         =   "frmMenuCons.frx":A5E8
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   11
      Left            =   4200
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   14
      Top             =   0
      Width           =   1815
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contas a Receber"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   11
         Left            =   210
         TabIndex        =   15
         Top             =   600
         Width           =   1305
      End
      Begin VB.Image imgIco 
         Height          =   240
         Index           =   11
         Left            =   720
         Picture         =   "frmMenuCons.frx":AEB2
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   14
      Left            =   4200
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   12
      Top             =   3240
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   14
         Left            =   600
         Picture         =   "frmMenuCons.frx":B136
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movimento Financeiro"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   14
         Left            =   75
         TabIndex        =   13
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   12
      Left            =   4200
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   10
      Top             =   1080
      Width           =   1815
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Controle de Cheques"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   12
         Left            =   105
         TabIndex        =   11
         Top             =   600
         Width           =   1515
      End
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   12
         Left            =   615
         Picture         =   "frmMenuCons.frx":BA00
         Top             =   90
         Width           =   480
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   13
      Left            =   4200
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   8
      Top             =   2160
      Width           =   1815
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cobrança"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   13
         Left            =   510
         TabIndex        =   9
         Top             =   600
         Width           =   705
      End
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   13
         Left            =   600
         Picture         =   "frmMenuCons.frx":C6CA
         Top             =   90
         Width           =   480
      End
   End
   Begin VB.PictureBox pcCalend 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9930
      Left            =   8565
      ScaleHeight     =   9930
      ScaleWidth      =   2460
      TabIndex        =   0
      Top             =   0
      Width           =   2460
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   0
         Picture         =   "frmMenuCons.frx":D394
         ScaleHeight     =   2865
         ScaleWidth      =   2385
         TabIndex        =   1
         Top             =   0
         Width           =   2415
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            FillColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   1080
            Picture         =   "frmMenuCons.frx":17743
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   2
            Top             =   120
            Width           =   480
         End
         Begin VB.Label lblDia 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblDia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   885
            TabIndex        =   6
            Top             =   600
            Width           =   645
         End
         Begin VB.Label lblNumDia 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   825
            Left            =   1020
            TabIndex        =   5
            Top             =   795
            Width           =   375
         End
         Begin VB.Label lblMes 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mês"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   4
            Top             =   1560
            Width           =   465
         End
         Begin VB.Label lblAno 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2007"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   360
            Left            =   840
            TabIndex        =   3
            Top             =   1875
            Width           =   735
         End
      End
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   240
      TabIndex        =   7
      Top             =   8880
      Width           =   10095
   End
End
Attribute VB_Name = "frmMenuCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bBlock As Boolean
Dim bActive As Boolean
Sub P_foco(Indice As Integer)
    Dim iCount As Integer
    For iCount = 0 To pcBotao.Count - 1
        pcBotao(iCount).BackColor = vbWhite
        
    Next
    pcBotao(Indice).BackColor = &HECE3D0
End Sub
Sub P_focoMouse(Indice As Integer)
    Dim iCount As Integer
    On Error Resume Next
    For iCount = 0 To pcBotao.Count - 1
        lblBotao(iCount).FontBold = False
        lblBotao(iCount).FontUnderline = False
        
    Next
    lblBotao(Indice).FontBold = True
    lblBotao(Indice).FontUnderline = True
End Sub

Private Sub Form_Activate()
    lblDia.caption = StrConv(WeekdayName(Weekday(Date)), vbProperCase)
    lblNumDia.caption = Day(Date)
    lblMes.caption = StrConv(MonthName(Month(Date)), vbProperCase)
    lblAno.caption = Year(Date)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 And Shift = vbCtrlMask Then
        If mdiFrm.mnRel.Visible = True Then
            mdiFrm.mnRel.Visible = False
        Else
            mdiFrm.mnRel.Visible = True
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.height = mdiFrm.ScaleHeight - mdiFrm.pcStatus.height
    Me.width = mdiFrm.ScaleWidth - 100
    Me.top = 100
    
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    P_mensagem
End Sub

Private Sub imgIco_Click(Index As Integer)
    pcBotao_Click Index
End Sub

Private Sub imgIco_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    pcBotao_MouseMove Index, Button, Shift, X, Y
End Sub

Private Sub lblBotao_Click(Index As Integer)
    pcBotao_Click Index
End Sub

Private Sub pcBotao_Click(Index As Integer)
    If Index = 0 Then
        frmEdiPro.Show 0
        frmEdiPro.SetFocus
    ElseIf Index = 1 Then
        frmEdiClientes.Show 0
        frmEdiClientes.SetFocus
    ElseIf Index = 2 Then
        frmEdiRep.Show 0
        frmEdiRep.SetFocus
    ElseIf Index = 3 Then
        frmEdiVendas.Show 0, mdiFrm
        frmEdiVendas.SetFocus
 
    ElseIf Index = 4 Then
        frmEdiCai.Show 0
    ElseIf Index = 5 Then
        frmOptUsu.Show 0, mdiFrm
    ElseIf Index = 6 Then
    
    
        If MsgBox("Deseja sair do sistema?", vbYesNo + vbQuestion) = vbYes Then
            P_Fim
        End If
    ElseIf Index = 7 Then
        frmConsign.Show 0
        frmConsign.SetFocus
    ElseIf Index = 8 Then
        frmTransfConsign.Show 0
        frmTransfConsign.SetFocus
    ElseIf Index = 9 Then
        frmDevolucaoConsign.Show 0
        frmDevolucaoConsign.SetFocus
    ElseIf Index = 10 Then
        frmAcerto.Show 0
        frmAcerto.SetFocus
    ElseIf Index = 11 Then
        frmEditaMov.Show 0
        frmEditaMov.SetFocus
        
    ElseIf Index = 12 Then
        frmControleChequesReceber.Show 0
        frmControleChequesReceber.SetFocus
    ElseIf Index = 13 Then
        frmCobranca.Show 0
        frmCobranca.SetFocus
    ElseIf Index = 14 Then
        frmEdiFin.Show 0
        frmEdiFin.SetFocus
    End If
End Sub
Sub novaVenda()
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    Dim strSQL As String
    
    strSQL = Vendas.F_GeraNovaVenda("00000")
    
    
    
    frmVenda.caption = strNome_Emp & " - Venda a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
    frmVenda.tag = strSQL
    On Error Resume Next
    
    frmVenda.Show 0, mdiFrm
End Sub

Private Sub pcBotao_GotFocus(Index As Integer)
    P_focoMouse Index
End Sub

Private Sub pcBotao_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If Index < pcBotao.Count - 1 Then
            pcBotao(Index + 1).SetFocus
        End If
    ElseIf KeyCode = vbKeyLeft Then
        
    ElseIf KeyCode = vbKeyRight Then
    
    ElseIf KeyCode = vbKeyUp Then
        If Index > 0 Then
            pcBotao(Index - 1).SetFocus
        End If
    ElseIf KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        pcBotao_Click Index
    End If
End Sub

Private Sub pcBotao_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If bBlock = True Then Exit Sub
    
    P_focoMouse Index
    
    
    
    P_mensagem Index
    
    
    
End Sub

Sub P_mensagem(Optional Index As Integer = -1)
    
    If Index = -1 Then
        lblMsg.caption = ""
        lblMsg.Visible = False
        Exit Sub
    ElseIf Index = 0 Then
        lblMsg.caption = "Abre lista de produtos. Consulta, Alteração, Inclusão de um Novo Produto"
        
    ElseIf Index = 1 Then
        lblMsg.caption = "Caso queira consultar, incluir, alterar um cliente. Clique para entrar."
    ElseIf Index = 2 Then
        lblMsg.caption = "Notas Fiscais de compra. Buscar um nota, concluir um lançamento, lançar um nota Fiscal"
    ElseIf Index = 3 Then
        lblMsg.caption = "Clique para efetuar uma nova Venda."
    ElseIf Index = 4 Then
        lblMsg.caption = "Para concluir uma venda ou pesquisar nas vendas já efetuadas, clique aqui."
    ElseIf Index = 5 Then
        lblMsg.caption = "Lançamentos diários do Caixa."
    ElseIf Index = 6 Then
        lblMsg.caption = "Alterar senha, relogar no sistema, e outras opções pessoais."
    ElseIf Index = 7 Then
        lblMsg.caption = "Sair para o Windows."
    End If
    
  '  pcMsg.Left = Me.Width - pcMsg.Width
  '  pcMsg.Top = Me.Height - pcMsg.Width
    lblMsg.Visible = True
    
    
End Sub
