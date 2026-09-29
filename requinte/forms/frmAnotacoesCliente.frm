VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAnotacoesCliente 
   Caption         =   "Anotações do Cliente"
   ClientHeight    =   4980
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   7905
   Icon            =   "frmAnotacoesCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRepeteAnot 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3480
      Width           =   7575
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox txtAnot 
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   5415
   End
   Begin VB.CommandButton cmdAddAnot 
      Caption         =   "Adicionar"
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridAnot 
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4260
      _Version        =   393216
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label38 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anotação:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   735
   End
   Begin VB.Menu mnOpt 
      Caption         =   "Opções"
      Visible         =   0   'False
      Begin VB.Menu mnExcluirAnot 
         Caption         =   "Excluir Anotação"
      End
   End
End
Attribute VB_Name = "frmAnotacoesCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddAnot_Click()
    If Me.tag = "" Then
        MsgBox "Insira anotações só depois do cliente estar gravado!", vbExclamation
        Exit Sub
    End If
    Clientes.AdicionaAnot Me.tag, txtAnot
    buscaAnot
    txtAnot.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub
Sub buscaAnot()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
        strSQL = "select * from tab_clientesanot where cod_cli='" & Me.tag & "' and excluido='N' order by id desc "
        abreObjetoAdo rPesq, strSQL
    
        GridUtil.AjustaGridData gridAnot, 4, 1, 0, "id,anotacao,data,usuario", "id,anotacao,datahora,cod_usu", "0,4000,1800,1000", rPesq, "", False, vbActiveTitleBar, False
        
    rPesq.Close
    Exit Sub
trata_Erro:
    P_MostraErro
    
End Sub

Private Sub Form_Activate()
    Me.caption = "Cliente: " & Me.tag & "-" & F_BuscaUmCampo("nome_cli", "tab_clientes", "cod_cli='" & Me.tag & "'")
    
    buscaAnot
End Sub

Private Sub gridAnot_Click()
    If gridAnot.row > 0 Then
        gridAnot.col = 1
        txtRepeteAnot.Text = gridAnot.Text
    End If
End Sub

Private Sub gridAnot_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnOpt
    End If
End Sub

Private Sub mnExcluirAnot_Click()
    If MsgBox("Confirma exclusão da anotação?", vbYesNo + vbQuestion) = vbYes Then
            gridAnot.col = 0
            
            Clientes.MarcaExcluidoAnot gridAnot.Text
            buscaAnot
        End If
End Sub
