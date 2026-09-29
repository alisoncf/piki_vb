VERSION 5.00
Begin VB.Form frmLembrete 
   Caption         =   "Lembrete"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8595
   Icon            =   "frmLembrete.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnExcl 
      Caption         =   "E&xcluir"
      Height          =   735
      Left            =   4920
      Picture         =   "frmLembrete.frx":992A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Esc - &Fecha&r"
      Height          =   735
      Left            =   6720
      Picture         =   "frmLembrete.frx":9A74
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "F7 - Sal&var"
      Height          =   735
      Left            =   3120
      Picture         =   "frmLembrete.frx":9BBE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox txtLembrete 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   8295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Lembrete:"
      Height          =   195
      Left            =   3360
      TabIndex        =   3
      Top             =   360
      Width           =   705
   End
End
Attribute VB_Name = "frmLembrete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tabela As String
Public campo As String
Public id As String
Dim ado As ADODB.Recordset

Private Sub cmndbtnExcl_Click()
    Dim strSQL As String
    strSQL = "update " & tabela & " set " & campo & "='' where " & id & "='" & Me.tag & "' "
    Conn.Execute strSQL
    Unload Me
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    If txtLembrete.Text = "" Then
        MsgBox "informe a mensagem que deseja inserir", vbExclamation
        Exit Sub
    End If
    
    
    strSQL = "update " & tabela & " set " & campo & "='" & txtLembrete.Text & "' where " & id & "='" & Me.tag & "' "
    Conn.Execute strSQL
    
    MsgBox "Informações gravadas!", vbInformation
    Unload Me
    
    
    
End Sub

Private Sub Form_Activate()
    carrega
End Sub
Sub carrega()
    Dim strSQL As String
    txtLembrete.Text = F_BuscaUmCampo(campo, tabela, id & "='" & Me.tag & "'")
    
    
    
    
    
    
    
    
End Sub
Private Sub Form_Load()
    
    Set ado = New ADODB.Recordset
    
    
    
    
End Sub

Private Sub txtLembrete_Change()
    If txtLembrete.Text = "" Then
        cmndbtnExcl.Enabled = False
    Else
        cmndbtnExcl.Enabled = True
    End If
End Sub
