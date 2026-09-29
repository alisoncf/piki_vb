VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBuscaOrc 
   Caption         =   "Localização de Orçamentos"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   9210
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9135
      Begin MSFlexGridLib.MSFlexGrid msflxgrd 
         Height          =   4935
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   8705
         _Version        =   393216
         BackColor       =   -2147483624
         BackColorBkg    =   -2147483624
         GridColor       =   16777215
         GridColorFixed  =   12632256
         AllowUserResizing=   1
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Busca:"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   9015
      Begin VB.TextBox txtbxAte 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   3840
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   7680
         Picture         =   "frmBuscaOrc.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtbxCliente 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtbxValidade 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   2760
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Buscar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   5040
         Picture         =   "frmBuscaOrc.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         Height          =   615
         Left            =   6360
         Picture         =   "frmBuscaOrc.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&E:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3840
         TabIndex        =   5
         Top             =   240
         Width           =   150
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Cliente:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "E&ntre:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2760
         TabIndex        =   3
         Top             =   240
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmBuscaOrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select * from tab_orcamentos where situacao_orc='1' "
    If txtbxCliente.Text <> "" Then
        strSQL = strSQL & " and nomecli_orc like '%" & txtbxCliente.Text & "%'"
        
    End If
    If txtbxValidade.Text <> "" Then
        strSQL = strSQL & " and validade_orc>=" & F_GravaData(txtbxValidade.Text)
    End If
    If txtbxAte.Text <> "" Then
        strSQL = strSQL & " and validade_orc<=" & F_GravaData(txtbxAte.Text)
    End If
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        msflxgrd.rows = 1
        Do While Not .EOF
            .MoveNext
        Loop
    End With
End Sub

Private Sub Form_Load()

    Me.Width = 9330
    Me.Height = 6585
    Me.Top = 10
    Me.Left = 10
End Sub

Private Sub txtbxAte_Change()
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

Private Sub txtbxCliente_GotFocus()
    P_SelecTexto txtbxCliente
End Sub

Private Sub txtbxCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCliente_LostFocus()
    P_AjustaTexto txtbxCliente
End Sub

Private Sub txtbxValidade_GotFocus()
    P_SelecTexto txtbxValidade
End Sub

Private Sub txtbxValidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValidade_LostFocus()
    P_AjustaData txtbxValidade
End Sub
