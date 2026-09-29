VERSION 5.00
Begin VB.Form frmMesclaComp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mesclar Componentes:"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   2640
      TabIndex        =   8
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmndbtnOk 
      Caption         =   "Mesclar"
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Selecione o novo coponente:"
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   4335
      Begin VB.ComboBox cmbxNovo 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selecione o Componente que será Mesclado:"
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   4335
      Begin VB.ComboBox cmbxMesc 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opção:"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.OptionButton optnCat 
         Caption         =   "Categoria"
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optnMarca 
         Caption         =   "Marca"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMesclaComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxMesc_Click()
    If optnCat.Value Then
        cmbxMesc.tag = F_BuscaUmCampo("cod_cat", "tab_categorias", "nome_cat='" & cmbxMesc.Text & "'")
    Else
        cmbxMesc.tag = F_BuscaUmCampo("cod_mar", "tab_marcas", "nome_mar='" & cmbxMesc.Text & "'")
    End If
End Sub

Private Sub cmbxNovo_Click()
    If optnCat.Value Then
        cmbxNovo.tag = F_BuscaUmCampo("cod_cat", "tab_categorias", "nome_cat='" & cmbxNovo.Text & "'")
    Else
        cmbxNovo.tag = F_BuscaUmCampo("cod_mar", "tab_marcas", "nome_mar='" & cmbxNovo.Text & "'")
    End If
End Sub

Private Sub cmndbtnOk_Click()
    If cmbxMesc.tag = "" Then
        MsgBox "Informe o componente a ser mesclado!", vbExclamation
        cmbxMesc.SetFocus
        Exit Sub
    End If
    If cmbxNovo.tag = "" Then
        MsgBox "Informe o novo componente!", vbExclamation
        cmbxNovo.SetFocus
        Exit Sub
    End If
    If MsgBox("Deseja realmente mesclar os coponentes?", vbYesNo + vbCritical) = vbNo Then
        Exit Sub
    End If
    If optnCat.Value Then
        cat
    Else
        marca
    End If
    MsgBox "Componente mesclado com sucesso!", vbInformation
End Sub
Sub cat()
    Dim strSQL As String
    strSQL = "update " & TAB_PRODUTOS & " set cod_cat='" & cmbxNovo.tag & "'"
    strSQL = strSQL & " where cod_cat='" & cmbxMesc.tag & "' "
    conn.Execute strSQL
    strSQL = "delete from tab_categorias where cod_Cat='" & cmbxMesc.tag & "' "
    conn.Execute strSQL
End Sub
Sub marca()
strSQL = "update " & TAB_PRODUTOS & " set cod_mar='" & cmbxNovo.tag & "'"
    strSQL = strSQL & " where cod_mar='" & cmbxMesc.tag & "' "
    conn.Execute strSQL
    strSQL = "delete from tab_marcas where cod_mar='" & cmbxMesc.tag & "' "
    conn.Execute strSQL
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub optnCat_Click()
    P_PreencheCateg cmbxMesc
    P_PreencheCateg cmbxNovo
End Sub

Private Sub optnMarca_Click()
    P_PreencheMarca cmbxMesc
    P_PreencheMarca cmbxNovo
End Sub
