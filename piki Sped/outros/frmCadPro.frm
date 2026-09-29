VERSION 5.00
Begin VB.Form frmCadPro 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Produtos:"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7905
   Icon            =   "frmCadPro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frControle 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.CommandButton cmndbtnFoto 
         Caption         =   "Imagem"
         Enabled         =   0   'False
         Height          =   315
         Left            =   143
         TabIndex        =   56
         Top             =   240
         Width           =   1125
      End
      Begin VB.PictureBox pcInfo 
         Height          =   855
         Left            =   120
         ScaleHeight     =   795
         ScaleWidth      =   2475
         TabIndex        =   51
         Top             =   5280
         Width           =   2535
         Begin VB.PictureBox pcEnt 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            Picture         =   "frmCadPro.frx":0ECA
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   53
            Top             =   67
            Width           =   240
         End
         Begin VB.PictureBox pcSai 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            Picture         =   "frmCadPro.frx":1014
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   52
            Top             =   420
            Width           =   240
         End
         Begin VB.Label lblInfoEnt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Informações de Entrada"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            TabIndex        =   55
            Top             =   82
            Width           =   1755
         End
         Begin VB.Label lblInfoSai 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Informações de Saída"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            TabIndex        =   54
            Top             =   435
            Width           =   1605
         End
      End
      Begin VB.CommandButton cmndbtnPro 
         Caption         =   "Pro&moção"
         Enabled         =   0   'False
         Height          =   855
         Left            =   2760
         Picture         =   "frmCadPro.frx":115E
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Frame Frame6 
         Caption         =   "Embalagens Existentes:"
         Height          =   2295
         Left            =   5640
         TabIndex        =   38
         Top             =   2520
         Width           =   2055
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   600
            TabIndex        =   41
            Text            =   "Un/embalagem"
            Top             =   1440
            Width           =   1335
         End
         Begin VB.ListBox lstEmb 
            Height          =   840
            ItemData        =   "frmCadPro.frx":1252
            Left            =   120
            List            =   "frmCadPro.frx":1254
            TabIndex        =   44
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtbxEmb 
            Height          =   285
            Left            =   120
            TabIndex        =   40
            Text            =   "0"
            Top             =   1440
            Width           =   495
         End
         Begin VB.CommandButton cmndbtnAddEmb 
            Caption         =   "Adicionar"
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   1800
            Width           =   855
         End
         Begin VB.CommandButton cmndbtnExclEmb 
            Caption         =   "E&xcluir"
            Height          =   375
            Left            =   1080
            TabIndex        =   43
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Embala&gens a Adicionar"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   1200
            Width           =   1710
         End
      End
      Begin VB.PictureBox pcFoto 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1125
         Left            =   143
         Picture         =   "frmCadPro.frx":1256
         ScaleHeight     =   1125
         ScaleWidth      =   1125
         TabIndex        =   49
         Top             =   525
         Width           =   1125
      End
      Begin VB.Frame Frame1 
         Caption         =   "Composição:"
         Height          =   1575
         Left            =   1320
         TabIndex        =   1
         Top             =   120
         Width           =   4215
         Begin VB.ComboBox cmbxCat 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   480
            Width           =   3585
         End
         Begin VB.ComboBox cmbxMarca 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1080
            Width           =   3585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Categoria"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Marca:"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lblCat 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "000"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   3720
            TabIndex        =   48
            Tag             =   "!"
            Top             =   480
            Width           =   435
         End
         Begin VB.Label lblMar 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "000"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   3720
            TabIndex        =   47
            Tag             =   "!"
            Top             =   1080
            Width           =   435
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1935
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   5415
         Begin VB.TextBox txtbxEstMin 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3480
            TabIndex        =   16
            Text            =   "0"
            Top             =   960
            Width           =   975
         End
         Begin VB.ComboBox cmbxICMS 
            Height          =   315
            ItemData        =   "frmCadPro.frx":3D95
            Left            =   2520
            List            =   "frmCadPro.frx":3DA2
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1560
            Width           =   2775
         End
         Begin VB.ComboBox cmbxSt 
            Height          =   315
            ItemData        =   "frmCadPro.frx":3DC1
            Left            =   120
            List            =   "frmCadPro.frx":3DCE
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1560
            Width           =   2295
         End
         Begin VB.TextBox txtbxNome_Pro 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   4335
         End
         Begin VB.TextBox txtbxUnidade 
            Height          =   285
            Left            =   4560
            TabIndex        =   10
            Text            =   "UN"
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtbxCodFab 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   14
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox txtbxLocal 
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtbxMargem 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4560
            TabIndex        =   18
            Text            =   "30"
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Est. &Mín."
            Height          =   195
            Left            =   3480
            TabIndex        =   15
            Top             =   720
            Width           =   645
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&ICMS:"
            Height          =   195
            Left            =   2520
            TabIndex        =   21
            Top             =   1320
            Width           =   435
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Situação Tributária:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   1320
            Width           =   1380
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "&Nome do Produto:"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Unidade:"
            Height          =   195
            Left            =   4560
            TabIndex        =   9
            Top             =   120
            Width           =   645
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Referência.:"
            Height          =   195
            Left            =   1440
            TabIndex        =   13
            Top             =   720
            Width           =   870
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Localiz.:"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   585
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "M.&Luc(%)"
            Height          =   195
            Left            =   4560
            TabIndex        =   17
            Top             =   720
            Width           =   660
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Códigos de &Barra:"
         Height          =   2295
         Left            =   5640
         TabIndex        =   32
         Top             =   120
         Width           =   2055
         Begin VB.CommandButton cmndbtnExclBarcode 
            Caption         =   "E&xcluir"
            Height          =   375
            Left            =   1080
            TabIndex        =   36
            Top             =   1800
            Width           =   855
         End
         Begin VB.CommandButton cmndbtnAddBarcode 
            Caption         =   "Adicionar"
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox txtbxCodBarra 
            Height          =   285
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            Width           =   1815
         End
         Begin VB.ListBox lstCodBarra 
            Height          =   840
            ItemData        =   "frmCadPro.frx":3DED
            Left            =   120
            List            =   "frmCadPro.frx":3DEF
            TabIndex        =   37
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Barcodes a Adicionar"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   1200
            Width           =   1515
         End
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar (ESC)"
         Height          =   855
         Left            =   6120
         Picture         =   "frmCadPro.frx":3DF1
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Confirmar F7"
         Height          =   855
         Left            =   4440
         Picture         =   "frmCadPro.frx":433F
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   4200
         Width           =   5415
         Begin VB.TextBox txtbxCusto 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4320
            TabIndex        =   31
            Tag             =   "!"
            Text            =   "0,00"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtbxEstoque 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            TabIndex        =   27
            Tag             =   "!"
            Text            =   "0,00"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtbxPreco 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2640
            TabIndex        =   29
            Tag             =   "!"
            Text            =   "0,00"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblCusCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C&usto:"
            Height          =   195
            Left            =   3840
            TabIndex        =   30
            Top             =   285
            Width           =   450
         End
         Begin VB.Label lblestCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Esto&que:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   285
            Width           =   630
         End
         Begin VB.Label lblPrecoCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Preço:"
            Height          =   195
            Left            =   2160
            TabIndex        =   28
            Top             =   285
            Width           =   465
         End
      End
      Begin VB.Frame Frame5 
         Height          =   615
         Left            =   120
         TabIndex        =   23
         Top             =   3600
         Width           =   5415
         Begin VB.CheckBox chkProd 
            Caption         =   "Produto de &Balança:"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "frmCadPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Dim cPreco As String 'para carregar o preço atual no caso de mudança de preço


Private Sub chkProd_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyF5 Then
        P_PreencheCateg cmbxCat
    End If
    
End Sub

Private Sub cmbxICMS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxMarca_Click()
    lblMar.caption = F_BuscaUmCampo("Cod_MAR", "tab_marcas", "nome_mar='" & cmbxMarca.Text & "'")
End Sub

Private Sub cmbxEsp_Click()
    lblEsp.caption = F_BuscaUmCampo("Cod_Esp", "tab_Especies", "nome_Esp='" & cmbxEsp.Text & "'")
End Sub

Private Sub cmbxCat_Click()
    lblCat.caption = F_BuscaUmCampo("Cod_Cat", "tab_Categorias", "nome_Cat='" & cmbxCat.Text & "'")
    
End Sub

Private Sub cmbxMarca_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyF5 Then
        P_PreencheMarca cmbxMarca
    End If
End Sub

Private Sub cmbxSt_Click()
    If cmbxSt.ListIndex <> 0 Then
        cmbxICMS.ListIndex = 0
        cmbxICMS.Enabled = False
        P_LocalizaOperacao "I05", cmbxICMS
    Else
        cmbxICMS.Enabled = True
    End If
End Sub

Private Sub cmbxSt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnAddBarcode_Click()
    If F_Vazio(txtbxCodBarra.Text) Then
        MsgBox "Informe o código de barras!", vbCritical
        Exit Sub
        txtbxCodBarra.SetFocus
        Exit Sub
    End If
    If F_InList(lstCodBarra, txtbxCodBarra.Text) >= 0 Then
        MsgBox "Este código de barras já está inserido!", vbCritical
    Else
        lstCodBarra.AddItem txtbxCodBarra.Text
        txtbxCodBarra.Text = ""
        txtbxCodBarra.SetFocus
        
    End If
End Sub

Private Sub cmndbtnCat_Click()
    BuscaCat
End Sub

Private Sub cmndbtnAddEmb_Click()
     If F_Vazio(txtbxEmb.Text) Then
        MsgBox "Informe a quantidade de unidades por embalagem!", vbCritical
        Exit Sub
        txtbxEmb.SetFocus
        
    End If
    If CCur(txtbxEmb.Text) = 0 Then
        MsgBox "Informe a quantidade de unidades por embalagem!", vbCritical
        Exit Sub
        txtbxEmb.SetFocus
        
    End If
    If F_InList(lstEmb, txtbxEmb.Text) >= 0 Then
        MsgBox "Esta embalagem já está inserida!", vbCritical
    Else
        lstEmb.AddItem txtbxEmb.Text
        
        txtbxEmb.SetFocus
    End If
End Sub

Private Sub cmndbtnExclBarcode_Click()
    If lstCodBarra.ListIndex = -1 Then
        MsgBox "É necessário que se escolha um código de barras!", vbCritical
        lstCodBarra.SetFocus
        Exit Sub
    End If
    If lstCodBarra.ListCount = 0 Then
        MsgBox "Não há o que excluir!", vbExclamation
        Exit Sub
    End If
    If lstCodBarra.ListCount = 1 Then
        MsgBox "Este código de barras não pode ser excluído por ser o único!", vbCritical
        Exit Sub
    End If
    If MsgBox("Deseja realmente excluir o código de barras?", vbYesNo + vbCritical) = vbYes Then
        lstCodBarra.RemoveItem lstCodBarra.ListIndex
    End If
End Sub

Private Sub cmndbtnMarcas_Click()
    buscaMarca
End Sub

Private Sub cmndbtnRefresh_Click()
    RefresheComp
End Sub

Private Sub cmndbtnExclEmb_Click()
    If lstEmb.ListCount = 0 Then
        MsgBox "Não há o que excluir!", vbExclamation
        Exit Sub
    End If
    If lstEmb.ListIndex = -1 Then
        MsgBox "É necessário que se escolha m item!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Deseja realmente excluir a embalagem?", vbYesNo + vbCritical) = vbYes Then
        lstEmb.RemoveItem lstEmb.ListIndex
    End If
End Sub


Private Sub cmndbtnFoto_Click()
    frmFoto.caption = txtbxNome_pro.Text
    frmFoto.Tag = Me.Tag
    frmFoto.Show 1
End Sub

Private Sub cmndbtnPro_Click()
    frmProdProm.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    
    If lstCodBarra.ListCount < 1 And Me.Tag <> "" Then
        MsgBox "Não é possível continuar sem um código de barras!", vbCritical
        Exit Sub
    End If
   If cmbxSt.ListIndex > 0 Then
        If cmbxICMS.Text = "" Then
            MsgBox "Informe o ICMS do produto!", vbExclamation
            If cmbxICMS.Enabled Then
                cmbxICMS.SetFocus
            End If
        End If
   End If
    If Me.Tag = "" Then
        strSQL = F_BuscaUmCampo("nome_pro", "tab_produtos", "NOME_PRO='" & _
            txtbxNome_pro.Text & "'")
        If strSQL <> "" Then
            MsgBox "Já existe um produto com este nome!", vbCritical
            txtbxNome_pro.SetFocus
            Exit Sub
        End If
        If txtbxCodFab.Text <> "" Then
        
            strSQL = F_BuscaUmCampo("nome_pro", "tab_produtos", "codfab_pro='" & _
            txtbxCodFab.Text & "'")
            
        End If
            If strSQL <> "" Then
                If MsgBox("Já existe um produto com esta referência!" & _
                     Chr(13) & strSQL & Chr(13) & "Deseja continuar?", vbYesNo + vbCritical) = vbNo Then
                    Exit Sub
                End If
            End If
        If cmbxSt.Text = "" Then
            MsgBox "Informe a Situacação Tributária!", vbCritical
            cmbxSt.SetFocus
            Exit Sub
        End If
        If MsgBox("confirma inclusão do Produto?", vbYesNo + vbQuestion) = vbYes Then
            P_Grava
        End If
    Else
        If MsgBox("Confirma alteração do Produto?", vbYesNo + vbQuestion) = vbYes Then
            P_Altera
        End If
    End If
        
End Sub
Sub P_Altera()
    On Error GoTo trata_erro
    Dim strSQL As String
    strSQL = "update tab_produtos set "
    strSQL = strSQL & "nome_pro='" & txtbxNome_pro.Text & "',"
    strSQL = strSQL & "unidade_pro='" & txtbxUnidade.Text & "',"
    strSQL = strSQL & "codfab_pro='" & txtbxCodFab.Text & "', "
    strSQL = strSQL & "balanca_pro='" & IIf(chkProd.Value = 1, "S", "N") & "',"
    strSQL = strSQL & "cod_mar='" & lblMar.caption & "', "
    strSQL = strSQL & "cod_cat='" & lblCat.caption & "', "
    
    strSQL = strSQL & "endereco_pro='" & txtbxLocal.Text & "',"
    strSQL = strSQL & "lucro_pro=" & F_GravaDec(txtbxMargem.Text) & ",  "
    strSQL = strSQL & "pv_pro=" & F_GravaDec(txtbxPreco.Text) & ","
    strSQL = strSQL & "cr_pro=" & F_GravaDec(txtbxCusto.Text) & ","
   ' If strAltpreco_usu = "S" Then
  '      strSQL = strSQL & "er_pro=" & F_GravaDec(txtbxEstoque.Text) & ","
   ' End If
    strSQL = strSQL & "estmin_pro=" & F_GravaDec(txtbxEstMin.Text) & ","
    strSQL = strSQL & "st_pro='" & Left(cmbxSt.Text, 1) & "' "  ' situacao trib"
    If cmbxICMS.Text <> "" Then
        strSQL = strSQL & ",cod_imp='" & Left(cmbxICMS.Text, 3) & "'  "
    End If
    strSQL = strSQL & " where cod_pro='" & Me.Tag & "' "
    Banco.Execute strSQL
    P_AltPreco Me.Tag, CCur(txtbxPreco.Text), CCur(cPreco)
    strSQL = "delete from tab_codbarras where cod_pro='" & Me.Tag & "' "
    Banco.Execute strSQL
    With lstCodBarra
        For iCount = 0 To .ListCount - 1
            .ListIndex = iCount
            strSQL = "insert into tab_codbarras(cod_pro,codbarra_pro)values("
            strSQL = strSQL & "'" & Me.Tag & "',"
            strSQL = strSQL & "'" & .Text & "')"
            Banco.Execute strSQL
        Next
    End With
    gravaEmb
    
    
   ' strSQL = F_GeraLog(strSQL, "tab_produtos")
    MsgBox "Produto Alterado com Sucesso!", vbInformation
    Unload Me
    On Error Resume Next
    If IsLoad("frmEdiPro") Then
        RefreshDataPosition frmEdiPro.dtPro
    End If
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub gravaEmb()
    Dim strSQL As String
    strSQL = "delete from tab_embalagens where cod_pro='" & Me.Tag & "' "
    Banco.Execute strSQL
    With lstEmb
        For iCount = 0 To .ListCount - 1
            .ListIndex = iCount
            strSQL = "insert into tab_embalagens(cod_pro,qtd_emb)values("
            strSQL = strSQL & "'" & Me.Tag & "',"
            strSQL = strSQL & .Text & ")"
            Banco.Execute strSQL
        Next
    End With
End Sub
Sub P_Grava()
On Error GoTo trata_erro

    Dim strSQL As String
    Dim sCod As String
    Dim sMar As String
    Dim iCount As Integer
    
    sCod = F_BuscaUmCampo(" max(cod_pro)", "tab_produtos", "cod_pro like '" & lblCat.caption & "%'")
    If sCod = "" Then
        sCod = "0001"
    Else
        
        sCod = Format(Int(Mid(sCod, 4, 4)) + 1, "000#")
    End If
    sCod = F_Ean8(lblCat.caption & sCod)
    If lstCodBarra.ListCount = 0 Then
'        If MsgBox("Este produto não possui um código de barras!" & _
            Chr(13) & "Deseja que o Sistema gere um automaticamente?", vbYesNo + _
                vbCritical) = vbYes Then
            lstCodBarra.AddItem sCod
 '       End If
    End If
    strSQL = "insert into tab_produtos(cod_pro,nome_pro,unidade_pro,"
    strSQL = strSQL & "cod_cat,cod_mar,codfab_pro,"
    strSQL = strSQL & "lucro_pro,"
    strSQL = strSQL & "endereco_pro,er_pro,ef_pro,st_pro,cr_pro,cf_pro,cm_pro,pv_pro,balanca_pro,cod_imp,estmin_pro)Values("
    strSQL = strSQL & "'" & sCod & "' ,"
    
    strSQL = strSQL & "'" & txtbxNome_pro.Text & "',"
    strSQL = strSQL & "'" & txtbxUnidade.Text & "',"
    strSQL = strSQL & "'" & lblCat.caption & "',"
    strSQL = strSQL & "'" & lblMar.caption & "', "
    strSQL = strSQL & "'" & txtbxCodFab.Text & "',"
    strSQL = strSQL & F_GravaDec(txtbxMargem.Text) & ","
    strSQL = strSQL & "'" & txtbxLocal.Text & "', "
    strSQL = strSQL & F_GravaDec(txtbxEstoque.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxEstoque.Text) & ","
    strSQL = strSQL & "'" & Left(cmbxSt.Text, 1) & "',"
    strSQL = strSQL & F_GravaDec(txtbxCusto.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxCusto.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxCusto.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxPreco.Text) & ","
    strSQL = strSQL & "'" & IIf(chkProd.Value = 1, "S", "N") & "',"
    If cmbxICMS.Text <> "" Then
    
        strSQL = strSQL & "'" & Left(cmbxICMS.Text, 3) & "', "
    Else
        strSQL = strSQL & "'I01',"
    End If
    strSQL = strSQL & F_GravaDec(txtbxEstMin.Text) & ") "

    
    Banco.Execute strSQL
    
    
    
    
    strSQL = "delete from tab_codbarras where cod_pro='" & sCod & "' "
    Banco.Execute strSQL
    With lstCodBarra
        For iCount = 0 To .ListCount - 1
            .ListIndex = iCount
            strSQL = "insert into tab_codbarras(cod_pro,codbarra_pro)values("
            strSQL = strSQL & "'" & sCod & "',"
            strSQL = strSQL & "'" & .Text & "')"
            Banco.Execute strSQL
            
        Next
    End With
    gravaEmb
    MsgBox "Produto Gravado com Sucesso", vbInformation
    txtbxCodFab.Text = ""
    txtbxCusto.Text = "0,00"
    txtbxEstoque.Text = "0,00"
    cmbxSt.ListIndex = -1
    cmbxCat.ListIndex = -1
    cmbxMarca.ListIndex = -1
    txtbxUnidade.Text = ""
    txtbxLocal.Text = ""
    txtbxMargem.Text = "0,00"
    txtbxPreco.Text = "0,00"
    txtbxNome_pro.Text = ""
    txtbxCodFab.Text = ""
    txtbxCodBarra.Text = ""
    lstCodBarra.Clear
    If IsLoad("frmEdiPro") Then
        RefreshDataPosition frmEdiPro.dtPro
    End If
    cmbxCat.SetFocus
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmnxdbtnBuscaFor_Click()
    frmSelectCli.txtbxNome.Tag = "F"
    frmSelectCli.txtbxNome.Text = txtbxFornecedor.Text
    
    On Error Resume Next
    
    frmSelectCli.Show 1
    txtbxFornecedor.Text = frmSelectCli.dtBusca.Recordset!Nome
    txtbxFornecedor.Tag = frmSelectCli.dtBusca.Recordset!COD
    Unload frmSelectCli
End Sub


Sub BuscaUltEntrada()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select max(data_not) as data, "
End Sub

Private Sub Form_Activate()
    On Error GoTo trata_erro
    MudaAparencia Me
    If bActive Then Exit Sub
    bActive = True
    cPreco = txtbxPreco.Text
    
    If strAltpreco_usu = "S" Then
        lblPrecoCap.Enabled = True
        lblCusCap.Enabled = True
        txtbxPreco.Enabled = True
        txtbxCusto.Enabled = True
        
    Else
        lblPrecoCap.Enabled = False
        lblCusCap.Enabled = False
        txtbxPreco.Enabled = False
        txtbxCusto.Enabled = False
        txtbxCustoFiscal.Enabled = False
    End If
    If Me.Tag <> "" Then
        cmndbtnFoto.Enabled = True
    End If
    'P_PreencheCateg cmbxCat
    'P_PreencheMarca cmbxMarca
    Exit Sub
    
trata_erro:
P_MostraErro
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        BuscaCat
    End If
    If KeyCode = vbKeyF3 Then
       buscaMarca
    End If
    If KeyCode = vbKeyF4 Then
        
    End If
    If KeyCode = vbKeyF7 Then
        cmndbtnSalvar_Click
    End If
    
End Sub
Sub BuscaCat()
    frmEdiCat.Show 0
    frmEdiCat.SetFocus
End Sub
Sub buscaMarca()
     frmEdiMar.Show 0
        frmEdiMar.SetFocus
End Sub
Sub RefresheComp()
    P_PreencheCateg cmbxCat
        P_PreencheMarca cmbxMarca
End Sub
Private Sub Form_Load()
'    Me.Height = 5280
'    Me.Width = 4935
    bActive = False
    P_CentraForm Me
    
    
    P_PreencheMarca cmbxMarca
    P_PreencheCateg cmbxCat
    P_PreencheImpostos cmbxICMS
    cmbxSt.ListIndex = 1
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblInfoEnt.FontBold = False
    lblInfoEnt.FontUnderline = False
    pcEnt.BorderStyle = 0
End Sub

Private Sub Form_Resize()
'     P_CentraObejct frControle, Me
End Sub

Private Sub frControle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblInfoEnt.FontBold = False
    lblInfoEnt.FontUnderline = False
    pcEnt.BorderStyle = 0
    lblInfoSai.FontBold = False
    lblInfoSai.FontUnderline = False
    pcSai.BorderStyle = 0
End Sub

Private Sub lblInfoEnt_Click()
    pcEnt_Click
End Sub

Private Sub lblInfoEnt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblInfoEnt.FontBold = True
    lblInfoEnt.FontUnderline = True
    pcEnt.BorderStyle = 1
End Sub

Private Sub lblInfoSai_Click()
    pcSai_Click
End Sub

Private Sub lblInfoSai_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblInfoSai.FontBold = True
    lblInfoSai.FontUnderline = True
    pcSai.BorderStyle = 1
End Sub

Private Sub lstCodBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub lstEmb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub pcEnt_Click()
    If Me.Tag <> "" Then
        P_InfoEntradasPro Me.Tag
    Else
        MsgBox "Não há produto gravado!", vbCritical
    End If
End Sub

Private Sub pcEnt_GotFocus()
    lblInfoEnt_MouseMove 0, 0, 0, 0
End Sub

Private Sub pcEnt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        pcEnt_Click
    End If
End Sub

Private Sub pcEnt_LostFocus()
    frControle_MouseMove 0, 0, 0, 0
End Sub

Private Sub pcInfo_Click()
    lblInfoEnt.FontBold = False
    lblInfoEnt.FontUnderline = False
    pcEnt.BorderStyle = 0
    lblInfoSai.FontBold = False
    lblInfoSai.FontUnderline = False
    pcSai.BorderStyle = 0
End Sub

Private Sub pcInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblInfoEnt.FontBold = False
    lblInfoEnt.FontUnderline = False
    pcEnt.BorderStyle = 0
    lblInfoSai.FontBold = False
    lblInfoSai.FontUnderline = False
    pcSai.BorderStyle = 0
End Sub

Private Sub pcSai_Click()
    If Me.Tag = "" Then Exit Sub
    P_InfoSaidasPro Me.Tag
End Sub

Private Sub pcSai_GotFocus()
    lblInfoSai_MouseMove 0, 0, 0, 0
End Sub

Private Sub pcSai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        pcSai_Click
    End If
End Sub

Private Sub pcSai_LostFocus()
    frControle_MouseMove 0, 0, 0, 0
End Sub

Private Sub txtbxCodBarra_GotFocus()
    P_SelecTexto txtbxCodBarra
End Sub

Private Sub txtbxCodBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnAddBarcode_Click
    End If
End Sub

Private Sub txtbxCodBarra_LostFocus()
    P_AjustaTexto txtbxCodBarra
End Sub

Private Sub txtbxPrecoMin_GotFocus()
    
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
    'P_SelecTexto txtbxDescricao
End Sub

Private Sub txtbxDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDescricao_LostFocus()
    'P_AjustaTexto txtbxDescricao
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

Private Sub txtbxCusto_GotFocus()
    P_SelecTexto txtbxCusto
End Sub

Private Sub txtbxCusto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCusto_LostFocus()
    P_ValorLostFocus txtbxCusto
End Sub



Private Sub txtbxEmb_GotFocus()
    P_SelecTexto txtbxEmb
End Sub

Private Sub txtbxEmb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxEmb_LostFocus()
    P_ValorLostFocus txtbxEmb, "###,##0"
End Sub

Private Sub txtbxEstoque_GotFocus()
    P_SelecTexto txtbxEstoque
End Sub

Private Sub txtbxEstoque_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
     End If
End Sub

Private Sub txtbxEstoque_LostFocus()
    P_ValorLostFocus txtbxEstoque
End Sub

Private Sub txtbxFornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmnxdbtnBuscaFor_Click
    End If
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
    P_SelecTexto txtbxNome_pro
End Sub

Private Sub txtbxNome_Pro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_Pro_LostFocus()
    P_AjustaTexto txtbxNome_pro
End Sub

Private Sub txtbxPei_GotFocus()
    
End Sub


Private Sub txtbxPreco_GotFocus()
    P_SelecTexto txtbxPreco
End Sub

Private Sub txtbxPreco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxPreco_LostFocus()
    P_ValorLostFocus txtbxPreco
End Sub

Private Sub txtbxPrecoMin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxPrecoMin_LostFocus()
    
End Sub


Private Sub txtbxUltComp_Change()

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

