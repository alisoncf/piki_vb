VERSION 5.00
Begin VB.Form frmEscolherVisao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Escolha Uma Visão de Navegação"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6045
   Icon            =   "frmEscolherVisao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcVisao 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   0
      Left            =   1200
      ScaleHeight     =   825
      ScaleWidth      =   3585
      TabIndex        =   1
      Top             =   720
      Width           =   3615
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
         TabIndex        =   10
         Top             =   120
         Width           =   975
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
      ScaleWidth      =   3585
      TabIndex        =   2
      Top             =   1680
      Width           =   3615
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
         Index           =   1
         Left            =   1080
         TabIndex        =   9
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
      ScaleWidth      =   3585
      TabIndex        =   3
      Top             =   2640
      Width           =   3615
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
         Index           =   2
         Left            =   1080
         TabIndex        =   8
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
      ScaleWidth      =   3585
      TabIndex        =   4
      Top             =   3600
      Width           =   3615
      Begin VB.Label lblVisao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Atacado"
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
         TabIndex        =   7
         Top             =   120
         Width           =   1380
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
      ScaleWidth      =   3585
      TabIndex        =   5
      Top             =   4560
      Width           =   3615
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
         Index           =   4
         Left            =   1080
         TabIndex        =   6
         Top             =   120
         Width           =   930
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
      Left            =   1560
      TabIndex        =   0
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
Sub ganhaFoco(Index As Integer)
    
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
            ganhaFoco picFoco + 1
        End If
    End If
    If KeyCode = vbKeyUp Then
        If picFoco > 0 Then
            ganhaFoco picFoco - 1
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
    If strVisaoEstoque_usu <> "S" Then
        pcVisao(1).Enabled = False
        lblVisao(1).Enabled = False
    End If
    If strVisaoFinanceiro_usu <> "S" Then
        pcVisao(2).Enabled = False
        lblVisao(2).Enabled = False
    End If
    
    
    If strVisaoAtacado_usu <> "S" Then
        pcVisao(3).Enabled = False
        lblVisao(3).Enabled = False
    End If
    If strVisaoGeral_usu <> "S" Then
        pcVisao(4).Enabled = False
        lblVisao(4).Enabled = False
    End If
    
     For iCount = 0 To pcVisao.count - 1
        If pcVisao(iCount).Enabled = True Then
            pcVisao(iCount).SetFocus
            ganhaFoco iCount
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
            Unload Me
        Else
            MsgBox "Você não tem permissão para este acesso!", vbExclamation
            Exit Sub
        End If
        
    ElseIf Index = 1 Then
       If strVisaoEstoque_usu = "S" Then
        mdiFrm.muda_visao tEstoque
        Else
            MsgBox "Você não tem permissão para este acesso!", vbExclamation
            Exit Sub
        End If
        
    ElseIf Index = 2 Then
        If strVisaoFinanceiro_usu = "S" Then
        mdiFrm.muda_visao tFinanceiro
        Else
            MsgBox "Você não tem permissão para este acesso!", vbExclamation
            Exit Sub
        End If
    ElseIf Index = 3 Then
        If strVisaoAtacado_usu = "S" Then
        mdiFrm.muda_visao tAtacado
        Else
            MsgBox "Você não tem permissão para este acesso!", vbExclamation
            Exit Sub
        End If
    ElseIf Index = 4 Then
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

