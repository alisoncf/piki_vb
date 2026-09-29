VERSION 5.00
Begin VB.Form frmCadImo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Imobilizado"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frControle 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   4695
         Begin VB.CommandButton cmndbtnSalvar 
            Caption         =   "&Confirmar"
            Height          =   615
            Left            =   120
            Picture         =   "frmCadImo.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   615
            Left            =   2640
            Picture         =   "frmCadImo.frx":058A
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   4695
         Begin VB.TextBox txtbxTotal 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "0,00"
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtbxUnit 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2040
            TabIndex        =   9
            Text            =   "0,00"
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtbxQtd 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   960
            TabIndex        =   7
            Text            =   "0"
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtbxNome 
            Height          =   285
            Left            =   120
            MaxLength       =   40
            TabIndex        =   3
            Top             =   360
            Width           =   4215
         End
         Begin VB.TextBox txtbxUni 
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   5
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            Height          =   195
            Left            =   3240
            TabIndex        =   10
            Top             =   840
            Width           =   405
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Uni&tário:"
            Height          =   195
            Left            =   2040
            TabIndex        =   8
            Top             =   840
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Qtd:"
            Height          =   195
            Left            =   960
            TabIndex        =   6
            Top             =   840
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Descrição"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Unidade:"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   645
         End
      End
   End
End
Attribute VB_Name = "frmCadImo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If Me.Tag = "" Then
        grava
    Else
        altera
    End If
    
End Sub
Sub grava()
    Dim strSQL As String
    Dim sCod As String
    sCod = MaxKeySQL("tab_imob", "cod_imo", "", "0000#")
    strSQL = "insert into tab_imob(cod_imo,desc_imo,unidade_imo,qtd_imo,unit_imo)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & txtbxNome.Text & "',"
    strSQL = strSQL & "'" & txtbxUni.Text & "',"
    strSQL = strSQL & F_GravaDec(txtbxQtd.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxUnit.Text) & ")"
    Banco.Execute strSQL
    MsgBox "Imobilizado incluído!", vbInformation
    txtbxNome.Text = ""
    txtbxQtd.Text = "0"
    txtbxUni.Text = "0"
    txtbxNome.SetFocus
    
End Sub
Sub altera()
    Dim strSQL As String
    strSQL = "update tab_imob set desc_imo='" & txtbxNome.Text & "',"
    strSQL = strSQL & "qtd_imo=" & F_GravaDec(txtbxQtd.Text) & ","
    strSQL = strSQL & "unit_imo=" & F_GravaDec(txtbxUnit.Text) & ", "
    strSQL = strSQL & "unidade_imo='" & txtbxUni.Text & "' "
    strSQL = strSQL & " where cod_imo='" & Me.Tag & "' "
    Banco.Execute strSQL
    MsgBox "Imobilizado alterado", vbInformation
    Unload Me
    
    
End Sub
Private Sub Form_Activate()
    MudaAparencia Me
    txtbxUnit_LostFocus
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

Private Sub txtbxQtd_GotFocus()
    P_SelecTexto txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxQtd_LostFocus()
    P_ValorLostFocus txtbxQtd
    txtbxTotal.Text = FormatVal(CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text))
End Sub

Private Sub txtbxTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxUni_GotFocus()
    P_SelecTexto txtbxUni
End Sub

Private Sub txtbxUni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxUni_LostFocus()
    P_AjustaTexto txtbxUni
End Sub

Private Sub txtbxUnit_GotFocus()
    P_SelecTexto txtbxUnit
End Sub

Private Sub txtbxUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxUnit_LostFocus()
    P_ValorLostFocus txtbxUnit
    txtbxTotal.Text = FormatVal(CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text))
End Sub

