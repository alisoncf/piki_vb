VERSION 5.00
Begin VB.Form frmCadPadrao1 
   Caption         =   "Inclusão"
   ClientHeight    =   3300
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6555
   Icon            =   "frmCadPadrao1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtComplemento 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   5895
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   6555
      TabIndex        =   6
      Top             =   2325
      Width           =   6555
      Begin VB.CommandButton cmdNovo 
         Caption         =   "Salvar"
         Height          =   735
         Left            =   3360
         Picture         =   "frmCadPadrao1.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   4920
         Picture         =   "frmCadPadrao1.frx":1351
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.TextBox txtDescricao 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   5895
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblComplemento 
      AutoSize        =   -1  'True
      Caption         =   "Complemento:"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição:"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "frmCadPadrao1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Public campo1 As String
Public campo2 As String
Public campo3 As String
Public tabela As String


Private Sub cmdNovo_Click()
    If Me.tag = "" Then
        grava
    Else
        altera
    End If
End Sub
Sub grava()
    Dim strSQL As String
    strSQL = "insert into " & tabela & "(" & campo2
    If campo3 <> "" Then
        strSQL = strSQL & "," & campo3
    End If
    
    strSQL = strSQL & ")values("
    
    strSQL = strSQL & "'" & txtDescricao.Text & "')"
    conn.Execute strSQL
    MsgBox "Ok", vbInformation
    Unload Me
End Sub
Sub altera()
    Dim strSQL As String
    strSQL = "update " & tabela & " set "
    
    If campo3 <> "" Then
        strSQL = strSQL & campo3 & "='" & txtComplemento & "', "
    End If
    
    strSQL = strSQL & campo2 & "='" & txtDescricao & "' "
    
    
    
    strSQL = strSQL & " where " & campo1 & "='" & Me.tag & "'"
    conn.Execute strSQL
    MsgBox "Alterado", vbInformation
End Sub
Private Sub cmdSair_Click()
     Unload Me
End Sub
Public Sub Inicia(tabela As String, campo1 As String, campo2 As String, campo3 As String, captionForm As String)
    Me.campo1 = campo1
    Me.campo2 = campo2
    Me.tabela = tabela
    If captionForm <> "" Then
        Me.caption = captionForm
    Else
        Me.caption = "Gravar: (" & tabela & ") " & campo1 & "/" & campo2
    End If
End Sub


Private Sub cmdSelecionar_Click()
    Me.tag = ado!codigo
    Me.Hide
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
End Sub
Private Sub txtBuscar_Change()
    carrega
End Sub
Sub carrega()
    Dim strSQL As String
    
    strSQL = "select *  from " & tabela & " where " & campo2 & " like '" & txtBuscar.Text & "%' "
    strSQL = strSQL & " order by " & campo2
    abreObjetoAdo ado, strSQL
    
End Sub

Private Sub txtComplemento_GotFocus()
    P_SelecTexto txtComplemento
End Sub

Private Sub txtComplemento_KeyDown(KeyCode As Integer, Shift As Integer)
    If keyecode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtComplemento_LostFocus()
    P_AjustaTexto txtComplemento, nada_muda
End Sub

Private Sub txtDescricao_GotFocus()
    ganhaFoco txtDescricao
End Sub

Private Sub txtDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    
End Sub

Private Sub txtDescricao_LostFocus()
    AjustaTexto txtDescricao
End Sub
