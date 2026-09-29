VERSION 5.00
Begin VB.Form frmMenuEst 
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
      Index           =   9
      Left            =   2040
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   26
      Top             =   3360
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   9
         Left            =   600
         Picture         =   "frmMenuEstoque.frx":0000
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reposição"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   9
         Left            =   465
         TabIndex        =   27
         Top             =   600
         Width           =   795
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   8
      Left            =   2040
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   24
      Top             =   2280
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   8
         Left            =   600
         Picture         =   "frmMenuEstoque.frx":0937
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perdas"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   600
         TabIndex        =   25
         Top             =   600
         Width           =   525
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   7
      Left            =   2040
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   22
      Top             =   1200
      Width           =   1815
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Configurar Etiquetas"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   23
         Top             =   600
         Width           =   1455
      End
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   7
         Left            =   600
         Picture         =   "frmMenuEstoque.frx":1601
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
      Left            =   2040
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   20
      Top             =   120
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   510
         Index           =   6
         Left            =   600
         Picture         =   "frmMenuEstoque.frx":1EC4
         Top             =   0
         Width           =   510
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gerar Etiquetas"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   21
         Top             =   600
         Width           =   1125
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   2
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   18
      Top             =   2280
      Width           =   1815
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compras"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   19
         Top             =   600
         Width           =   645
      End
      Begin VB.Image imgIco 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmMenuEstoque.frx":2DEE
         Top             =   240
         Width           =   240
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
      TabIndex        =   10
      Top             =   0
      Width           =   2460
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   0
         Picture         =   "frmMenuEstoque.frx":3062
         ScaleHeight     =   2865
         ScaleWidth      =   2385
         TabIndex        =   11
         Top             =   0
         Width           =   2415
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            FillColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1080
            Picture         =   "frmMenuEstoque.frx":D411
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   12
            Top             =   120
            Width           =   240
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
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
            TabIndex        =   13
            Top             =   1875
            Width           =   735
         End
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
      TabIndex        =   8
      Top             =   5520
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   5
         Left            =   600
         Picture         =   "frmMenuEstoque.frx":F10B
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sair"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   660
         TabIndex        =   9
         Top             =   600
         Width           =   285
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
      TabIndex        =   6
      Top             =   4440
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   4
         Left            =   600
         Picture         =   "frmMenuEstoque.frx":FDD5
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   525
         TabIndex        =   7
         Top             =   600
         Width           =   555
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
      TabIndex        =   4
      Top             =   3360
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   510
         Index           =   3
         Left            =   600
         Picture         =   "frmMenuEstoque.frx":10A9F
         Top             =   0
         Width           =   510
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nova Compra"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   5
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
      Index           =   1
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   1
         Left            =   600
         MouseIcon       =   "frmMenuEstoque.frx":114C1
         Picture         =   "frmMenuEstoque.frx":12E43
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedores"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   3
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
      Index           =   0
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.Image imgIco 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   600
         MouseIcon       =   "frmMenuEstoque.frx":13B0D
         Picture         =   "frmMenuEstoque.frx":15157
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produtos"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   645
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
      TabIndex        =   17
      Top             =   8880
      Width           =   10095
   End
End
Attribute VB_Name = "frmMenuEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bBlock As Boolean
Dim bActive As Boolean
Sub P_foco(Indice As Integer)
    Dim iCount As Integer
    For iCount = 0 To pcBotao.count - 1
        pcBotao(iCount).BackColor = vbWhite
        
    Next
    pcBotao(Indice).BackColor = &HECE3D0
End Sub
Sub P_focoMouse(Indice As Integer)
    Dim iCount As Integer
    On Error Resume Next
    For iCount = 0 To pcBotao.count - 1
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

Private Sub Form_Load()
    Me.Height = mdiFrm.ScaleHeight - mdiFrm.pcStatus.Height
    Me.Width = mdiFrm.ScaleWidth - 100
    Me.Top = 100
    
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
    ElseIf Index = 1 Then
        frmEdiFor.Show 0
        frmEdiFor.SetFocus
    ElseIf Index = 2 Then
        Dim fr As New frmEntradas
        fr.operacao = "F"
        frmEntradas.Show 0
    ElseIf Index = 3 Then
        frmEntradas.Show 0
        frmEntradas.SetFocus
    ElseIf Index = 4 Then
        frmOptUsu.Show 0, mdiFrm
    ElseIf Index = 5 Then
        If MsgBox("Deseja sair do sistema?", vbYesNo + vbQuestion) = vbYes Then
            P_Fim
        End If
    ElseIf Index = 6 Then
        frmGeraEtiquetas.Show 0
        frmGeraEtiquetas.SetFocus
    ElseIf Index = 7 Then
        frmEdiEti.Show 0
        frmEdiEti.SetFocus
    ElseIf Index = 8 Then
        frmEdiPerdas.Show 0
        frmEdiPerdas.SetFocus
    ElseIf Index = 9 Then
        frmEdiReposicao.Show 0
        frmEdiReposicao.SetFocus
    
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
        If Index < pcBotao.count - 1 Then
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
