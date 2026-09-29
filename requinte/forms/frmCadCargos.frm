VERSION 5.00
Begin VB.Form frmCadCargos 
   Caption         =   "Cadastro de Cargos"
   ClientHeight    =   1995
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   4020
   Icon            =   "frmCadCargos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   4020
   Begin VB.Frame frControle 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   3735
         Begin VB.CommandButton cmndbtnSalvar 
            Caption         =   "&Confirmar"
            Height          =   615
            Left            =   120
            Picture         =   "frmCadCargos.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   615
            Left            =   1920
            Picture         =   "frmCadCargos.frx":05F2
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3735
         Begin VB.ComboBox cmbxVenda 
            Height          =   315
            ItemData        =   "frmCadCargos.frx":075D
            Left            =   2880
            List            =   "frmCadCargos.frx":0767
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   345
            Width           =   750
         End
         Begin VB.TextBox txtbxNome 
            Height          =   285
            Left            =   120
            MaxLength       =   40
            TabIndex        =   2
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Venda:"
            Height          =   195
            Left            =   2880
            TabIndex        =   3
            Top             =   120
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Nome do Cargo:"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   1155
         End
      End
   End
End
Attribute VB_Name = "frmCadCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    
    If strCadusu_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
    
    If txtbxNome.Text = "" Then
        MsgBox "Informe o Nome do Cargo!", vbExclamation
         
        Exit Sub
    End If
    
    If Me.tag = "" Then
        
        If F_BuscaUmCampo("Nome_Car", "Tab_Cargos", "Nome_car='" & _
            txtbxNome.Text & "'") <> "" Then
            MsgBox "Já existe um Cargo com este Nome!", vbExclamation
            Exit Sub
        End If
        Grava
    Else
        If F_BuscaUmCampo("Cod_Bco", "Tab_Cargos", "Nome_Car='" & _
            txtbxNome.Text & "' And cod_Car <> '" & Me.tag & "'") <> "" Then
            MsgBox "Já existe um Cargo com este Nome!", vbExclamation
            Exit Sub
        End If
        Altera
    End If
End Sub
Sub Grava()
    Dim strSQL As String
    strSQL = "insert into tab_Cargos(Cod_car,venda_car,Nome_Car)Values("
    strSQL = strSQL & "'" & MaxKeySQL("tab_cargos", "cod_car", "", "00#") & "',"
    strSQL = strSQL & "'" & Left(cmbxVenda.Text, 1) & "', "
    strSQL = strSQL & "'" & txtbxNome.Text & "')"
    Conn.Execute strSQL
    MsgBox "Cargo gravado!", vbInformation
    
    txtbxNome.Text = ""
    
    If IsLoad("frmEdiCargos") Then
        RefreshDataPosition frmEdiCargos.adc
    End If
    txtbxNome.SetFocus
End Sub
Sub Altera()
    Dim strSQL As String
    strSQL = "update tab_cargos set "
    
    strSQL = strSQL & " nome_car = '" & txtbxNome.Text & "',"
    strSQL = strSQL & " venda_car='" & Left(cmbxVenda.Text, 1) & "' "
    strSQL = strSQL & " where cod_car = '" & Me.tag & "' "
    
    Conn.Execute strSQL
    MsgBox "Cargo alterado!", vbInformation
    If IsLoad("frmEdiCargos") Then
        RefreshDataPosition frmEdiCargos.adc
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    P_CentraForm Me
End Sub

Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If IsLoad("frmEdiBan") Then
        'frmEdiBan.adc.Refresh
    End If
End Sub


Private Sub txtbxNome_GotFocus()
    P_SelecTexto txtbxNome
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_LostFocus()
    P_AjustaTexto txtbxNome
End Sub


