VERSION 5.00
Begin VB.Form frmEscolherVisao 
   BorderStyle     =   0  'None
   Caption         =   "Escolha uma Visão"
   ClientHeight    =   8640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7065
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcVisao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   7
      Left            =   1200
      ScaleHeight     =   825
      ScaleWidth      =   4425
      TabIndex        =   15
      Top             =   7440
      Width           =   4455
      Begin VB.Label lblVisao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Geral"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   1080
         TabIndex        =   16
         Top             =   120
         Width           =   930
      End
   End
   Begin VB.PictureBox pcVisao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   6
      Left            =   1200
      ScaleHeight     =   825
      ScaleWidth      =   4425
      TabIndex        =   13
      Top             =   6480
      Width           =   4455
      Begin VB.Label lblVisao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendas"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   1080
         TabIndex        =   14
         Top             =   120
         Width           =   1275
      End
   End
   Begin VB.PictureBox pcVisao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   5
      Left            =   1200
      ScaleHeight     =   825
      ScaleWidth      =   4425
      TabIndex        =   11
      Top             =   5520
      Width           =   4455
      Begin VB.Label lblVisao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pacote"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   1080
         TabIndex        =   12
         Top             =   120
         Width           =   1125
      End
   End
   Begin VB.PictureBox pcVisao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   4
      Left            =   1200
      ScaleHeight     =   825
      ScaleWidth      =   4425
      TabIndex        =   9
      Top             =   4560
      Width           =   4455
      Begin VB.Label lblVisao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Financeiro"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   1080
         TabIndex        =   10
         Top             =   120
         Width           =   1800
      End
   End
   Begin VB.PictureBox pcVisao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   3
      Left            =   1200
      ScaleHeight     =   825
      ScaleWidth      =   4425
      TabIndex        =   7
      Top             =   3600
      Width           =   4455
      Begin VB.Label lblVisao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estoque"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   1080
         TabIndex        =   8
         Top             =   120
         Width           =   1380
      End
   End
   Begin VB.PictureBox pcVisao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   2
      Left            =   1200
      ScaleHeight     =   825
      ScaleWidth      =   4425
      TabIndex        =   5
      Top             =   2640
      Width           =   4455
      Begin VB.Label lblVisao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crediário"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   1080
         TabIndex        =   6
         Top             =   120
         Width           =   1590
      End
   End
   Begin VB.PictureBox pcVisao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   1
      Left            =   1200
      ScaleHeight     =   825
      ScaleWidth      =   4425
      TabIndex        =   3
      Top             =   1680
      Width           =   4455
      Begin VB.Label lblVisao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consignado"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1080
         TabIndex        =   4
         Top             =   120
         Width           =   1995
      End
   End
   Begin VB.PictureBox pcVisao 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   0
      Left            =   1200
      ScaleHeight     =   825
      ScaleWidth      =   4425
      TabIndex        =   0
      Top             =   720
      Width           =   4455
      Begin VB.Label lblVisao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   1080
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Escolha uma Visão"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   2685
   End
End
Attribute VB_Name = "frmEscolherVisao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim picFoco As Integer
Sub P_selecttexto(Index As Integer)
    
    Dim iCount As Integer
    On Error Resume Next
    For iCount = 0 To pcVisao.count - 1
        pcVisao(iCount).BackColor = &H80000005
    
        
    Next
    If Index > -1 Then
        pcVisao(Index).BackColor = &HFFC0C0
        pcVisao(Index).SetFocus
        picFoco = Index
    End If
    
    
    
End Sub

Private Sub Form_Activate()
    DesabilitaControles
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If picFoco < pcVisao.count Then
            P_selecttexto picFoco + 1
        End If
    End If
    If KeyCode = vbKeyUp Then
        If picFoco > 0 Then
            P_selecttexto picFoco - 1
        End If
    End If
    If KeyCode = vbKeyReturn Then
        If picFoco > -1 Then
            pcVisao_Click picFoco
        End If
    End If
End Sub
Sub DesabilitaControles()
    Dim iCount As Integer
    Dim i As Integer
    Dim j As Integer
    On Error Resume Next
    If strVisaoCaixa_usu <> "S" Then
        pcVisao(0).Enabled = False
        lblVisao(0).Enabled = False
    End If
    If strVisaoConsignado_usu <> "S" Then
        pcVisao(1).Enabled = False
        lblVisao(1).Enabled = False
    End If
    If strVisaoCrediario_usu <> "S" Then
        pcVisao(2).Enabled = False
        lblVisao(2).Enabled = False
    End If
    If strVisaoEstoque_usu <> "S" Then
        pcVisao(3).Enabled = False
        lblVisao(3).Enabled = False
    End If
    If strVisaoFinanceiro_usu <> "S" Then
        pcVisao(4).Enabled = False
        lblVisao(4).Enabled = False
    End If
    If strVisaoPacote_usu <> "S" Then
        pcVisao(5).Enabled = False
        lblVisao(5).Enabled = False
    End If
    If strVisaoVendas_usu <> "S" Then
        pcVisao(6).Enabled = False
        lblVisao(6).Enabled = False
    End If
    If strVisaoGeral_usu <> "S" Then
        pcVisao(7).Enabled = False
        lblVisao(7).Enabled = False
    End If
     For iCount = 0 To pcVisao.count - 1
        If pcVisao(iCount).Enabled = True Then
            pcVisao(iCount).SetFocus
            P_selecttexto iCount
            i = i + 1
            j = iCount
        End If
        
       
        
    Next
     If i = 1 Then
            pcVisao_Click j
        End If
End Sub
Private Sub Form_Load()
    Set ado = New ADODB.Recordset
End Sub

Private Sub lblVisao_Click(Index As Integer)
    pcVisao_Click Index
End Sub

Private Sub pcVisao_Click(Index As Integer)
    If Index = 0 Then
        If strVisaoCaixa_usu = "S" Then
            mdiFrm.muda_visao tCAIXA
        Else
            MsgBox "Você não tem permissão para este acesso!", vbExclamation
            Exit Sub
        End If
        
    ElseIf Index = 1 Then
       If strVisaoConsignado_usu = "S" Then
        mdiFrm.muda_visao tConsignado
        Else
            MsgBox "Você não tem permissão para este acesso!", vbExclamation
            Exit Sub
        End If
        
    ElseIf Index = 2 Then
        If strVisaoCrediario_usu = "S" Then
        mdiFrm.muda_visao tCrediario
        Else
            MsgBox "Você não tem permissão para este acesso!", vbExclamation
            Exit Sub
        End If
    ElseIf Index = 3 Then
        If strVisaoEstoque_usu = "S" Then
        mdiFrm.muda_visao tEstoque
        Else
            MsgBox "Você não tem permissão para este acesso!", vbExclamation
            Exit Sub
        End If
    ElseIf Index = 4 Then
        If strVisaoFinanceiro_usu = "S" Then
        mdiFrm.muda_visao tFinanceiro
        Else
            MsgBox "Você não tem permissão para este acesso!", vbExclamation
            Exit Sub
        End If
    ElseIf Index = 5 Then
        If strVisaoPacote_usu = "S" Then
        mdiFrm.muda_visao tPacote
        Else
            MsgBox "Você não tem permissão para este acesso!", vbExclamation
            Exit Sub
        End If
    ElseIf Index = 6 Then
        If strVisaoVendas_usu = "S" Then
        mdiFrm.muda_visao TVendas
        Else
            MsgBox "Você não tem permissão para este acesso!", vbExclamation
            Exit Sub
        End If
    ElseIf Index = 7 Then
        If strVisaoGeral_usu = "S" Then
        mdiFrm.muda_visao tGeral
        Else
            MsgBox "Você não tem permissão para este acesso!", vbExclamation
            Exit Sub
        End If
    End If
    
    Unload Me
End Sub

Private Sub pcVisao_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        pcVisao_Click Index
        
    End If
    
End Sub
