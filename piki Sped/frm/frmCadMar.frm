VERSION 5.00
Begin VB.Form frmCadMar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de Marcas:"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   Icon            =   "frmCadMar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   7095
   Begin VB.TextBox txtbxNome 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6255
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Confirmar"
      Height          =   735
      Left            =   3720
      Picture         =   "frmCadMar.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   5160
      Picture         =   "frmCadMar.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nome da Marca:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1185
   End
End
Attribute VB_Name = "frmCadMar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    If F_Vazio(txtbxNOme) Then
        MsgBox "Informe o nome da Marca", vbExclamation
        txtbxNOme.SetFocus
        Exit Sub
    End If
    strSQL = F_BuscaUmCampo("Cod_Mar", TAB_MARCAS, "Nome_Mar='" & txtbxNOme.Text & "'")
    If Me.tag = "" Then
        If strSQL <> "" Then
            MsgBox "Já existe esta Marca!", vbExclamation
            Exit Sub
        End If
        P_Salva
    Else
        
        If strSQL <> "" Then
            If strSQL <> Me.tag Then
                MsgBox "Já existe esta Marca!", vbExclamation
                Exit Sub
            End If
        End If
        P_Altera
    End If
    
End Sub
Sub P_Salva()
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    sCod = F_BuscaUmCampo(" max(cod_mar)", TAB_MARCAS, "")
    If sCod = "" Then
        sCod = "001"
    Else
        
        sCod = Format(Int(sCod) + 1, "00#")
    End If
    strSQL = "Insert into  " & TAB_MARCAS & " (cod_mar,nome_mar)values("
    strSQL = strSQL & "'" & sCod & "', "
    strSQL = strSQL & "'" & txtbxNOme.Text & "')  "
    conn.Execute strSQL
    
    MsgBox "Marca Incluída com Sucesso.", vbInformation
    txtbxNOme.Text = ""
    txtbxNOme.SetFocus
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub P_Altera()
    Dim strSQL As String
    
    On Error GoTo trata_Erro
    
    strSQL = "update  " & TAB_MARCAS & "  set "
    
    strSQL = strSQL & " nome_mar='" & txtbxNOme.Text & "'  "
    strSQL = strSQL & " where cod_mar='" & Me.tag & "' "
    conn.Execute strSQL
    
    MsgBox "Marca Alterada com Sucesso.", vbInformation
    Unload Me
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Activate()
    'MudaAparencia Me
End Sub

Private Sub Form_Resize()
    'P_CentraObejct frControle, Me
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

