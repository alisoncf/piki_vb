VERSION 5.00
Begin VB.Form frmCentralRelatorios 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Central de Relatórios"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   615
      Left            =   6240
      TabIndex        =   17
      Top             =   4320
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2355
      ScaleWidth      =   7635
      TabIndex        =   8
      Top             =   960
      Width           =   7695
      Begin VB.CommandButton cmndbtnResumoMovConsData 
         Caption         =   "Resumo de Movimentação de Consignados por Data"
         Height          =   615
         Left            =   1920
         TabIndex        =   19
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnResumoAce 
         Caption         =   "Resumo de Acertos"
         Height          =   615
         Left            =   0
         TabIndex        =   18
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnDetalhar 
         Caption         =   "Detalhar Resumo de Atividades"
         Height          =   615
         Left            =   1920
         TabIndex        =   16
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnDebRepr 
         Caption         =   "Débitos de Representantes"
         Height          =   615
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnAcertosData 
         Caption         =   "Acertos Por Data"
         Height          =   615
         Left            =   0
         TabIndex        =   14
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnComissPerm 
         Caption         =   "Comissões Permanentes"
         Height          =   615
         Left            =   3840
         TabIndex        =   13
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnResumoComissPer 
         Caption         =   "Resumo de Comissões Permantentes"
         Height          =   615
         Left            =   3840
         TabIndex        =   12
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnResumo 
         Caption         =   "Resumo de Movimentações"
         Height          =   615
         Left            =   5760
         TabIndex        =   11
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnRelProdRep 
         Caption         =   "Relatório de Produtos por Representantes"
         Height          =   615
         Left            =   5760
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnRelParcRep 
         Caption         =   "Relatório de Parcelas por Representante"
         Height          =   615
         Left            =   1920
         TabIndex        =   9
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   7995
      TabIndex        =   7
      Top             =   0
      Width           =   8025
      Begin VB.TextBox txtbxDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtbxAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtbxNomeCli 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   4575
      End
      Begin VB.CommandButton cmndbtnCliente 
         Height          =   285
         Left            =   7320
         Picture         =   "frmCentralRelatorios.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "De:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Até:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1440
         TabIndex        =   2
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Representante:"
         Height          =   195
         Left            =   2760
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCentralRelatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnAcertosData_Click()
    Dim d As DataIniFin
    d = F_BuscaDatas(txtbxDe.Text, txtbxAte.Text)
    P_RelAcertosPorRepresentante d.dataI, d.dataF, txtbxNomeCli.tag
    
End Sub

Private Sub cmndbtnCliente_Click()
    buscaRep
End Sub
Sub buscaRep()
    Dim r As representante.representante
    r = F_BuscaRep(txtbxNomeCli.Text)
    If r.cod_rep <> "" Then
        
        txtbxNomeCli.Text = r.nome_rep
        txtbxNomeCli.BackColor = &H80FFFF
        txtbxNomeCli.tag = r.cod_rep
    End If
    
End Sub

Private Sub cmndbtnComissPerm_Click()
    frmComissParc.Show 0, mdiFrm
    frmComissParc.SetFocus
    
End Sub

Private Sub cmndbtnComissProv_Click()
    frmComissProv.Show 0
    
End Sub

Private Sub cmndbtnDebRepr_Click()
    P_RelContasAReceberPorRepresentante txtbxDe.Text, txtbxDe.Text, txtbxNomeCli.tag
End Sub

Private Sub cmndbtnRelParcRep_Click()
    P_RelParcelasPorRepr txtbxNomeCli.tag, txtbxDe.Text, txtbxAte.Text
End Sub

Private Sub cmndbtnRelProdRep_Click()
    frmProdutosPorRepresentante.Show 0
    frmProdutosPorRepresentante.SetFocus
End Sub

Private Sub cmndbtnResumo_Click()
    frmResumoRep.Show 0
    frmResumoRep.SetFocus
End Sub

Private Sub cmndbtnResumoAce_Click()
     Dim d As DataIniFin
    d = F_BuscaDatas(txtbxDe.Text, txtbxAte.Text)
    P_RelResumoAcertos d.dataI, d.dataF, txtbxNomeCli.tag
End Sub

Private Sub cmndbtnResumoComissPer_Click()
    Dim d As DataIniFin
    d = F_BuscaDatas
    P_RelResumoComissPer d.dataI, d.dataF
End Sub

Private Sub cmndbtnDetalhar_Click()
    Dim strSQL As String
    If txtbxNomeCli.tag = "" Then
        If MsgBox("Não foi indicado um representante, deseja continuar", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            buscaRep
            Exit Sub
        End If
        
    End If
    P_RelResumoAtv txtbxNomeCli.tag, txtbxDe.Text, txtbxAte.Text
End Sub

Private Sub cmndbtnResumoMovConsData_Click()
    frmRelResumoConsData.Show 0
    frmRelResumoConsData.SetFocus
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub


Private Sub txtbxDe_GotFocus()
    P_SelecTexto txtbxDe
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDe_LostFocus()
    P_AjustaData txtbxDe
End Sub
Private Sub txtbxAte_GotFocus()
    P_SelecTexto txtbxAte
End Sub

Private Sub txtbxAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxAte_LostFocus()
    P_AjustaData txtbxAte
End Sub

Private Sub txtbxNomeCli_Change()
    txtbxNomeCli.tag = ""
    txtbxNomeCli.BackColor = vbWhite
End Sub

Private Sub txtbxNomeCli_GotFocus()
    P_SelecTexto txtbxNomeCli
End Sub

Private Sub txtbxNomeCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnCliente_Click
    End If
End Sub

Private Sub txtbxNomeCli_LostFocus()
    If txtbxNomeCli.tag = "" Then
        
    End If
End Sub
