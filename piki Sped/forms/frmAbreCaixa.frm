VERSION 5.00
Begin VB.Form frmAbreCaixa 
   Caption         =   "Abertura de Caixa"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7695
   Icon            =   "frmAbreCaixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1215
      ScaleWidth      =   7215
      TabIndex        =   10
      Top             =   960
      Width           =   7215
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   120
         Width           =   7095
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   7695
      TabIndex        =   6
      Top             =   2910
      Width           =   7695
      Begin VB.CommandButton cmdAbrir 
         Caption         =   "Abrir Caixa"
         Height          =   855
         Left            =   2280
         Picture         =   "frmAbreCaixa.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdFechar 
         Caption         =   "Fechar"
         Height          =   855
         Left            =   4080
         Picture         =   "frmAbreCaixa.frx":0C54
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   855
         Left            =   5880
         Picture         =   "frmAbreCaixa.frx":115C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.TextBox txtValorFinal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Text            =   "0,00"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtValorInicial 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Text            =   "0,00"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtCaixa 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Final"
      Height          =   195
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Inicial"
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caixa:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   435
   End
End
Attribute VB_Name = "frmAbreCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub abrir()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select a.*,b.data_fechamento,b.hora_fechamento,b.id as id_fecha,"
    strSQL = strSQL & " b.fecha_usu from abrecaixa as a left join fechacaixa as b "
    strSQL = strSQL & " on a.id=b.id "
    strSQL = strSQL & " where cod_cai='" & txtCaixa.Text & "' order by a.id desc limit 1"
    abreObjetoAdo rPesq, strSQL
    
    
    With rPesq
        If .RecordCount > 0 Then
        
            If IsNull(!id_fecha) Then
                MsgBox "Esse caixa está aberto!" & !data & " " & !hora, vbExclamation
                Exit Sub
            End If
            
        End If
        
        If AbreCaixa(txtCaixa.Text, strCod_Usu, txtValorInicial) <> "" Then
            MsgBox "Caixa aberto!", vbInformation
            
            
            
        End If
            
        
    End With
    
    
    
    

End Sub

Private Sub cmdAbrir_Click()
    abrir
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    busca
End Sub

Private Sub Form_Load()
    txtCaixa.Text = strCod_Cai
    
    
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select a.*,b.data_fechamento,b.hora_fechamento,b.id as id_fecha,"
    strSQL = strSQL & " b.fecha_usu from abrecaixa as a left join fechacaixa as b "
    strSQL = strSQL & " on a.id=b.id "
    strSQL = strSQL & " where cod_cai='" & txtCaixa.Text & "' order by a.id desc limit 1"
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount > 0 And IsNull(id_fecha) Then
            
            txtValorInicial.Text = FormatVal(!valor_inicial)
            strSQL = "Caixa Aberto"
            txtInfo.Text = strSQL
            txtValorInicial.Enabled = False
            cmdAbrir.Enabled = False
            cmdFechar.Enabled = True
            
        End If
    End With
    Exit Sub
trata_Erro:
P_MostraErro
End Sub
Private Sub txtValorInicial_GotFocus()
    P_SelecTexto txtValorInicial
End Sub

Private Sub txtValorInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtValorInicial_LostFocus()
    P_ValorLostFocus txtValorInicial
End Sub
Private Sub txtvalorfinal_GotFocus()
    P_SelecTexto txtValorFinal
End Sub

Private Sub txtvalorfinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtvalorfinal_LostFocus()
    P_ValorLostFocus txtValorFinal
End Sub
