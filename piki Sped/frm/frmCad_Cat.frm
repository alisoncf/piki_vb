VERSION 5.00
Begin VB.Form frmCadCat 
   Caption         =   "Inclusão de Categorias"
   ClientHeight    =   3690
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   6495
   Icon            =   "frmCad_Cat.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   6495
   Begin VB.TextBox txtCod 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtbxNome 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   5655
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   4560
      Picture         =   "frmCad_Cat.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Confirmar"
      Height          =   735
      Left            =   3120
      Picture         =   "frmCad_Cat.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
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
    If F_Vazio(txtbxNOme) Then
        MsgBox "Informe o nome da Categoria", vbExclamation
        txtbxNOme.SetFocus
        Exit Sub
    End If
    strSQL = F_BuscaUmCampo("Cod_Cat", TAB_CATEGORIAS, "Nome_Cat='" & txtbxNOme.Text & "'")
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
    On Error GoTo trata_erro
    sCod = F_BuscaUmCampo(" max(cod_cat)", TAB_CATEGORIAS, "")
    If sCod = "" Then
        sCod = "001"
    Else
        
        sCod = Format(Int(sCod) + 1, "00#")
    End If
    strSQL = "Insert into " & TAB_CATEGORIAS & "  (cod_cat,nome_cat)values("
    strSQL = strSQL & "'" & sCod & "', "
    strSQL = strSQL & "'" & txtbxNOme.Text & "')  "
    conn.Execute strSQL
    
    MsgBox "Categoria Incluída com Sucesso.", vbInformation
    txtbxNOme.Text = ""
    txtbxNOme.SetFocus
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub P_Altera()
    Dim strSQL As String
    
    On Error GoTo trata_erro
    
    strSQL = "update " & TAB_CATEGORIAS & "  set "
    
    strSQL = strSQL & " nome_cat='" & txtbxNOme.Text & "'  "
    strSQL = strSQL & " where cod_cat='" & Me.tag & "' "
    conn.Execute strSQL
    
    MsgBox "Categoria Alterada com Sucesso.", vbInformation
    Unload Me
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        carrega
    End If
    bActive = True
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_erro
    strSQL = "select * from " & TAB_CATEGORIAS & " where cod_cat='" & Me.tag & "'"
    abreObjetoAdo rPesq, strSQL
    
    
    With rPesq
        If .RecordCount > 0 Then
            txtCod.Text = !cod_cat
            txtbxNOme.Text = !nome_cat
        End If
    End With
    Exit Sub
    
trata_erro:
    P_MostraErro
    
End Sub
Private Sub txtbxNome_GotFocus()
    ganhaFoco txtbxNOme
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxNome_LostFocus()
    AjustaTexto txtbxNOme
End Sub
