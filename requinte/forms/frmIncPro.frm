VERSION 5.00
Begin VB.Form frmIncPro 
   BackColor       =   &H8000000C&
   Caption         =   "Cadastro de Produtos:"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4995
   Icon            =   "frmIncPro.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   4995
   Begin VB.Frame frControle 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.Frame Frame2 
         Caption         =   "Dados do Produto:"
         Height          =   2055
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   4695
         Begin VB.TextBox txtbxNome_Pro 
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   4335
         End
         Begin VB.TextBox txtbxUnidade 
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtbxCodFab 
            Height          =   285
            Left            =   2520
            TabIndex        =   15
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox txtbxLocal 
            Height          =   285
            Left            =   1080
            TabIndex        =   13
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtbxDescricao 
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox txtbxMargem 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   19
            Text            =   "0"
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "&Nome do Produto:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Unidade:"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   645
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Cód. Fábricante:"
            Height          =   195
            Left            =   2520
            TabIndex        =   14
            Top             =   840
            Width           =   1170
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Localização:"
            Height          =   195
            Left            =   1080
            TabIndex        =   12
            Top             =   840
            Width           =   900
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Tipo:"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   1440
            Width           =   360
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&M.Lucro(%)"
            Height          =   195
            Left            =   1080
            TabIndex        =   18
            Top             =   1440
            Width           =   795
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Composição:"
         Height          =   1215
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4695
         Begin VB.ComboBox cmbxGenero 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   3225
         End
         Begin VB.ComboBox cmbxDiv 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   600
            Width           =   3225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Categoria"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Marca:"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   495
         End
         Begin VB.Label lblGen 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "000"
            ForeColor       =   &H00004000&
            Height          =   315
            Left            =   4050
            TabIndex        =   32
            Tag             =   "!"
            Top             =   240
            Width           =   435
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "000"
            ForeColor       =   &H00004000&
            Height          =   315
            Left            =   4050
            TabIndex        =   31
            Tag             =   "!"
            Top             =   600
            Width           =   435
         End
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   3240
         Picture         =   "frmIncPro.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Confirmar"
         Height          =   615
         Left            =   3240
         Picture         =   "frmIncPro.frx":0C54
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   120
         TabIndex        =   20
         Top             =   3600
         Width           =   3015
         Begin VB.TextBox txtbxCusto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Tag             =   "!"
            Text            =   "0,00"
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtbxEstoque 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Tag             =   "!"
            Text            =   "0,00"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtbxPreco 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1560
            TabIndex        =   28
            Tag             =   "!"
            Text            =   "0,00"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtbxUltComp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1560
            TabIndex        =   24
            Tag             =   "!"
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Custo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   21
            Top             =   120
            Width           =   555
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Esto&que:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   795
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Preço:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1560
            TabIndex        =   27
            Top             =   720
            Width           =   585
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Últ. Compra:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1560
            TabIndex        =   23
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Label lblCaptionFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cadastro de Produto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   240
         Left            =   0
         TabIndex        =   1
         Tag             =   "!"
         Top             =   0
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmIncPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxDiv_Click()
    lblDiv.caption = F_BuscaUmCampo("Cod_div", "tab_divisoes", "nome_div='" & cmbxDiv.Text & "'")
End Sub

Private Sub cmbxEsp_Click()
    lblEsp.caption = F_BuscaUmCampo("Cod_Esp", "tab_Especies", "nome_Esp='" & cmbxEsp.Text & "'")
End Sub

Private Sub cmbxGenero_Click()
    lblGen.caption = F_BuscaUmCampo("Cod_Gen", "tab_Generos", "nome_gen='" & cmbxGenero.Text & "'")
    P_PreencheDiv cmbxDiv, lblGen.caption
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If Me.Tag = "" Then
        P_grava
    Else
        P_Altera
    End If
    
End Sub
Sub P_Altera()
    Dim strSQL As String
    strSQL = "update tab_produtos set "
    strSQL = strSQL & "nome_pro='" & txtbxNome_Pro.Text & "',"
    strSQL = strSQL & "unidade_pro='" & txtbxUnidade.Text & "',"
    strSQL = strSQL & "descricao_pro='" & txtbxDescricao.Text & "',"
    strSQL = strSQL & "codfab_pro='" & txtbxCodFab.Text & "', "
    strSQL = strSQL & "custo_pro=" & f_GravaDec(txtbxCusto.Text) & ","
    strSQL = strSQL & "preco_pro=" & f_GravaDec(txtbxPreco.Text) & ","
    strSQL = strSQL & "lucro_pro=" & f_GravaDec(txtbxMargem.Text) & ", "
    strSQL = strSQL & "estoque_pro=" & f_GravaDec(txtbxEstoque.Text) & ", "
    strSQL = strSQL & "estmin_pro=" & f_GravaDec(txtbxEstMin.Text)
    strSQL = strSQL & " where cod_pro='" & Me.Tag & "' "
    banco.Execute strSQL
    MsgBox "Produto Alterado com Sucesso!", vbInformation
    Unload Me
    On Error Resume Next
    If IsLoad("frmEdiPro") Then
        RefreshDataPosition frmEdiPro.dtPro
    End If
End Sub
Sub P_grava()
    Dim strSQL As String
    Dim sCod As String
     sCod = F_BuscaUmCampo(" max(cod_pro)", "tab_produtos", " cod_esp='" & lblEsp.caption & "'")
    If sCod = "" Then
        sCod = "001"
    Else
        sCod = Mid(sCod, 4, 3)
        sCod = Format(Int(sCod) + 1, "00#")
    End If
    sCod = F_GeraDig(sCod)
    sCod = lblEsp.caption & sCod
    strSQL = "insert into tab_produtos(cod_pro,nome_pro,unidade_pro,"
    strSQL = strSQL & "cod_gen,cod_div,cod_esp,descricao_pro,codfab_pro,"
    strSQL = strSQL & "custo_pro,preco_pro,estoque_pro,lucro_pro,estmin_pro"
    strSQL = strSQL & ")Values("
    strSQL = strSQL & "'" & sCod & "' ,"
    strSQL = strSQL & "'" & txtbxNome_Pro.Text & "',"
    strSQL = strSQL & "'" & txtbxUnidade.Text & "',"
    strSQL = strSQL & "'" & lblGen.caption & "',"
    strSQL = strSQL & "'" & lblDiv.caption & "', "
    strSQL = strSQL & "'" & lblEsp.caption & "', "
    strSQL = strSQL & "'" & txtbxDescricao.Text & "',"
    strSQL = strSQL & "'" & txtbxCodFab.Text & "',"
    strSQL = strSQL & f_GravaDec(txtbxCusto.Text) & ","
    strSQL = strSQL & f_GravaDec(txtbxPreco.Text) & ","
    strSQL = strSQL & f_GravaDec(txtbxEstoque.Text) & ","
    strSQL = strSQL & f_GravaDec(txtbxMargem.Text) & ","
    strSQL = strSQL & f_GravaDec(txtbxEstMin.Text) & ")"
    banco.Execute strSQL
    MsgBox "Produto Gravado com Sucesso", vbInformation
    
    
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
   
End Sub

Private Sub Form_Load()
    Me.Height = 5280
    Me.Width = 4935
    
    P_CentraForm Me
    
    P_PreencheDiv cmbxDiv
    P_PreencheEsp cmbxEsp
    P_PreencheGen cmbxGenero
    
End Sub

Private Sub Form_Resize()
     P_CentraObejct frControle, Me
End Sub

Private Sub txtbxCodFab_GotFocus()
    P_SelecTexto txtbxCodFab
End Sub

Private Sub txtbxCodFab_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCodFab_LostFocus()
    P_AjustaTexto txtbxCodFab
End Sub

Private Sub txtbxDescricao_GotFocus()
    P_SelecTexto txtbxDescricao
End Sub

Private Sub txtbxDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDescricao_LostFocus()
    P_AjustaTexto txtbxDescricao
End Sub

Private Sub txtbxEstMin_GotFocus()
    P_SelecTexto txtbxEstMin
End Sub

Private Sub txtbxEstMin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxEstMin_LostFocus()
    P_ValorLostFocus txtbxEstMin
End Sub

Private Sub txtbxLocal_GotFocus()
    P_SelecTexto txtbxLocal
End Sub

Private Sub txtbxLocal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxLocal_LostFocus()
    P_AjustaTexto txtbxLocal
End Sub

Private Sub txtbxMargem_GotFocus()
2    P_SelecTexto txtbxMargem
End Sub

Private Sub txtbxMargem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxMargem_LostFocus()
    P_ValorLostFocus txtbxMargem
End Sub

Private Sub txtbxNome_Pro_GotFocus()
    P_SelecTexto txtbxNome_Pro
End Sub

Private Sub txtbxNome_Pro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_Pro_LostFocus()
    P_AjustaTexto txtbxNome_Pro
End Sub

Private Sub txtbxUnidade_GotFocus()
    P_SelecTexto txtbxUnidade
End Sub

Private Sub txtbxUnidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxUnidade_LostFocus()
    P_AjustaTexto txtbxUnidade
End Sub
