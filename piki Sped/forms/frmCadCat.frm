VERSION 5.00
Begin VB.Form frmCadCat 
   Caption         =   "Inclusão de Categorias"
   ClientHeight    =   4050
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   6720
   Icon            =   "frmCadCat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   6720
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   6690
      TabIndex        =   8
      Top             =   3075
      Width           =   6720
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Gravar"
         Height          =   735
         Left            =   2880
         Picture         =   "frmCadCat.frx":1084A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   4680
         Picture         =   "frmCadCat.frx":10994
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.TextBox txtQtdAcacado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Text            =   "0"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtDescMax 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Text            =   "0,00"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtbxNome 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   6135
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qtd Atacado:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto(%):"
      Height          =   195
      Left            =   1440
      TabIndex        =   6
      Top             =   1800
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nome da Categoria:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1410
   End
End
Attribute VB_Name = "frmCadCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    If F_Vazio(txtbxNome) Then
        MsgBox "Informe o nome da Categoria", vbExclamation
        txtbxNome.SetFocus
        Exit Sub
    End If
    strSQL = F_BuscaUmCampo("Cod_Cat", TAB_CATEGORIAS, "Nome_Cat='" & txtbxNome.Text & "'")
    If Me.tag = "" Then
        If strSQL <> "" Then
            MsgBox "Já existe esta Categoria!", vbExclamation
            Exit Sub
        End If
        P_Salva
    Else
        
        If strSQL <> "" Then
            If strSQL <> Me.tag Then
                MsgBox "Já existe esta Categoria!", vbExclamation
                Exit Sub
            End If
        End If
        P_Altera
    End If
    Unload Me
End Sub
Sub P_Salva()
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    sCod = F_BuscaUmCampo(" max(cod_cat)", TAB_CATEGORIAS, "")
    If sCod = "" Then
        sCod = "001"
    Else
        
        sCod = Format(Int(sCod) + 1, "00#")
    End If
    strSQL = "Insert into " & TAB_CATEGORIAS & "  (cod_cat,nome_cat,descmax_cat,qtdmin_cat)values("
    strSQL = strSQL & "'" & sCod & "', "
    strSQL = strSQL & "'" & txtbxNome.Text & "',  "
    strSQL = strSQL & F_GravaDec(txtDescMax.Text) & ","
    strSQL = strSQL & F_GravaDec(txtQtdAcacado.Text) & ")"
    conn.Execute strSQL
    
    MsgBox "Categoria Incluída com Sucesso.", vbInformation
    txtbxNome.Text = ""
    txtbxNome.SetFocus
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub P_Altera()
    Dim strSQL As String
    
    On Error GoTo trata_Erro
    
    strSQL = "update " & TAB_CATEGORIAS & "  set "
    
    strSQL = strSQL & " nome_cat='" & txtbxNome.Text & "',  "
    strSQL = strSQL & " qtdmin_cat=" & F_GravaDec(txtQtdAcacado.Text) & ","
    strSQL = strSQL & " descmax_cat=" & F_GravaDec(txtDescMax.Text)
    strSQL = strSQL & " where cod_cat='" & Me.tag & "' "
    conn.Execute strSQL
    
    MsgBox "Categoria Alterada com Sucesso.", vbInformation
    Unload Me
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        Carrega
    End If
    bActive = True
End Sub
Sub Carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select * from " & TAB_CATEGORIAS & " where cod_cat='" & Me.tag & "'"
    abreObjetoAdo rPesq, strSQL
    
    
    With rPesq
        If .RecordCount > 0 Then
            txtCod.Text = !cod_cat
            txtbxNome.Text = !nome_cat
        End If
    End With
    Exit Sub
    
trata_Erro:
    P_MostraErro
    
End Sub
Private Sub txtbxNome_GotFocus()
    ganhaFoco txtbxNome
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxNome_LostFocus()
    AjustaTexto txtbxNome
End Sub

Private Sub txtDescMax_GotFocus()
    P_SelecTexto txtDescMax
End Sub

Private Sub txtDescMax_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDescMax_LostFocus()
    P_ValorLostFocus txtDescMax
    If CCur(txtDescMax.Text) > 50 Then
        MsgBox "Valor muito alto!", vbExclamation
        txtDescMax.Text = "0"
        
    End If
    
End Sub

Private Sub txtQtdAcacado_GotFocus()
    P_SelecTexto txtQtdAcacado
End Sub

Private Sub txtQtdAcacado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtQtdAcacado_LostFocus()
    P_ValorLostFocus txtQtdAcacado
End Sub
