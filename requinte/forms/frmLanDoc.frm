VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLanDoc 
   Caption         =   "Lançamento de Documentos"
   ClientHeight    =   7740
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   11340
   DrawStyle       =   5  'Transparent
   Icon            =   "frmLanDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   11340
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "F7 - &Confirmar"
      Height          =   735
      Left            =   7440
      Picture         =   "frmLanDoc.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "ESC - &Fechar"
      Height          =   735
      Left            =   9360
      Picture         =   "frmLanDoc.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   6600
      Width           =   1815
   End
   Begin VB.PictureBox pcAtalhos 
      Appearance      =   0  'Flat
      BackColor       =   &H00C6AE86&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1320
      ScaleHeight     =   705
      ScaleWidth      =   3225
      TabIndex        =   43
      Top             =   6600
      Width           =   3255
      Begin VB.Label lblBuscaPro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Del> Excluir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1800
         TabIndex        =   50
         Top             =   45
         Width           =   1020
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F7 - Concluir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   9000
         TabIndex        =   49
         Top             =   45
         Width           =   1020
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F6 - Cancelar Venda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   7080
         TabIndex        =   48
         Top             =   45
         Width           =   1620
      End
      Begin VB.Label lblMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<F5> Navegar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1800
         TabIndex        =   47
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label lblMsg2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<+> Adicionar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F12 - Sair"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   10440
         TabIndex        =   45
         Top             =   45
         Width           =   735
      End
      Begin VB.Label lblTef 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Enter> Editar "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   44
         Top             =   45
         Width           =   1170
      End
   End
   Begin VB.Frame frControle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.ComboBox cmbxTipoTer 
         Height          =   315
         ItemData        =   "frmLanDoc.frx":09A2
         Left            =   3000
         List            =   "frmLanDoc.frx":09B2
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtbxNOme 
         Appearance      =   0  'Flat
         BackColor       =   &H00CF614E&
         Height          =   285
         Left            =   5280
         TabIndex        =   54
         Top             =   360
         Width           =   5415
      End
      Begin VB.TextBox txtbxClass 
         Appearance      =   0  'Flat
         BackColor       =   &H0005C9EB&
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Tag             =   "000"
         Text            =   "NÃO CLASSIFICADO"
         Top             =   2280
         Width           =   4815
      End
      Begin VB.TextBox txtbxBoleto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txtbxVenIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   18
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtbxValor_Doc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
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
         Left            =   9720
         TabIndex        =   14
         Text            =   "0,00"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtbxHist 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1080
         Width           =   6615
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tipo:"
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   1560
         TabIndex        =   4
         Top             =   0
         Width           =   1335
         Begin VB.OptionButton optnDup 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "D&uplicata"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optnCh 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "C&heque:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox txtbxRef 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtbxNum_Doc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Natureza"
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   1335
         Begin VB.OptionButton optnReceb 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "A &Receber"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optnPag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "A &Pagar"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame frCab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   5880
         TabIndex        =   22
         Top             =   1440
         Width           =   5175
         Begin VB.CheckBox chk 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Considerar Data Base p/ 1ª Parcela"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   840
            Value           =   1  'Checked
            Width           =   3135
         End
         Begin VB.ComboBox cmbxDias 
            Height          =   315
            ItemData        =   "frmLanDoc.frx":09F0
            Left            =   1080
            List            =   "frmLanDoc.frx":09FA
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtbxParc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            TabIndex        =   24
            Text            =   "1"
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtbxInterv 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   600
            TabIndex        =   26
            Text            =   "1"
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton cmndbtnDistr 
            Caption         =   "Distribuir"
            Height          =   375
            Left            =   2400
            TabIndex        =   28
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "&Parc.:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Tag             =   "!"
            Top             =   120
            Width           =   420
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "&Intervalo:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   600
            TabIndex        =   25
            Tag             =   "!"
            Top             =   120
            Width           =   660
         End
      End
      Begin VB.Frame frVal 
         Height          =   3735
         Left            =   120
         TabIndex        =   29
         Top             =   2640
         Width           =   11055
         Begin VB.CommandButton cmndbtnLancarCheque 
            Caption         =   "Lançar Cheque"
            Height          =   375
            Left            =   5760
            TabIndex        =   53
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox cmbxOpc 
            Height          =   315
            ItemData        =   "frmLanDoc.frx":0A0F
            Left            =   2160
            List            =   "frmLanDoc.frx":0A19
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   1920
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.ComboBox cmbxItem 
            Height          =   315
            ItemData        =   "frmLanDoc.frx":0A3C
            Left            =   120
            List            =   "frmLanDoc.frx":0A43
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtbxData 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   3960
            TabIndex        =   37
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox cmbxEspecie 
            Height          =   315
            ItemData        =   "frmLanDoc.frx":0A4A
            Left            =   960
            List            =   "frmLanDoc.frx":0A54
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox txtbxValor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   2880
            TabIndex        =   35
            Text            =   "0,00"
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmndbtnCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   8040
            TabIndex        =   38
            Top             =   240
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid msflxgrd 
            Height          =   2895
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   5106
            _Version        =   393216
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "E&spécie:"
            Height          =   255
            Left            =   960
            TabIndex        =   32
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "&Item:"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   3960
            TabIndex        =   36
            Top             =   120
            Width           =   390
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2880
            TabIndex        =   34
            Top             =   120
            Width           =   405
         End
      End
      Begin VB.CommandButton cmndbtnPlanoContas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         Picture         =   "frmLanDoc.frx":0A77
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10680
         Picture         =   "frmLanDoc.frx":0E01
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terceiro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5280
         TabIndex        =   60
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Terceiro:"
         Height          =   255
         Left            =   3000
         TabIndex        =   59
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblCPFCap 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "CPF/CNPJ:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4680
         TabIndex        =   57
         Top             =   720
         Width           =   825
      End
      Begin VB.Label lblCPF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00000000000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5640
         TabIndex        =   56
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Classificação:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&Boleto:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Tag             =   "!"
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Data &Base:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3000
         TabIndex        =   17
         Tag             =   "!"
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor:"
         Height          =   195
         Left            =   9720
         TabIndex        =   13
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Histórico:"
         Height          =   195
         Left            =   3000
         TabIndex        =   11
         Top             =   840
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Referência:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   825
      End
      Begin VB.Label lblTipo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "R - Contas a Receber"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   40
         Tag             =   "!"
         Top             =   6120
         Width           =   1440
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Documento:"
         Height          =   195
         Left            =   1440
         TabIndex        =   9
         Top             =   840
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmLanDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bCheque As Boolean
Function F_Soma() As Currency
    Dim cVal As Currency
    Dim iCount As Integer
    Dim iRow As Integer
    iRow = msflxgrd.row
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 1
            cVal = cVal + CCur(.Text)
        Next
        
    End With
    F_Soma = cVal
    msflxgrd.row = iRow
End Function

Private Sub cmbxDias_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxOpc_LostFocus()
    cmbxOpc.Visible = False
    
End Sub

Private Sub cmndbtnAdd_Click()
    Dim p As parcela
    Dim strSQL As String
    If F_Soma = CCur(txtbxValor_Doc.Text) Then
        MsgBox "Você não pode adicionar parcelas quando o Valor total está fechado!", vbExclamation
        Exit Sub
    End If
    p = F_DadosPar
    If p.valor <> 0 Then
        strSQL = p.vencimento & Chr(9) & FormatVal(p.valor)
        msflxgrd.AddItem strSQL
    Else
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    If Left(lblTipo.caption, 1) = "R" Then
        If cmbxTipoTer.ListIndex = 0 Then
            buscaCli
        ElseIf cmbxTipoTer.ListIndex = 1 Then
            buscaRep
        Else
            buscaFun
        End If
    Else
        
        
            buscaFor
        
    End If
        
End Sub
Sub buscaCli()
    Dim c As Clientes.cliente
    c = Clientes.F_BuscaCliente(txtbxNome.Text)
    If c.cod_cli <> "" Then
        txtbxNome.Text = c.nome_cli
        txtbxNome.tag = c.cod_cli
        lblcpf.caption = c.cnpj_Cli
        txtbxNum_Doc.SetFocus
    End If
End Sub
Sub buscaRep()
    Dim c As representante.representante
    c = representante.F_BuscaRep(txtbxNome.Text)
    If c.cod_Rep <> "" Then
        txtbxNome.Text = c.nome_rep
        txtbxNome.tag = c.cod_Rep
        lblcpf.caption = c.cpf_rep
        txtbxNum_Doc.SetFocus
    End If
End Sub

Sub buscaFor()
    Dim f As fornecedor.fornecedor
    f = fornecedor.F_SelectForn(txtbxNome.Text)
    If f.cod_For <> "" Then
        txtbxNome.Text = f.nome_for
        txtbxNome.tag = f.cod_For
        lblcpf.caption = f.cnpj_for
        txtbxNum_Doc.SetFocus
    End If
End Sub
Sub buscaFun()
    Dim f As funcionario.funcionario
    f = funcionario.F_SelectFuncionario(txtbxNome.Text)
    If f.cod_fun <> "" Then
        txtbxNome.Text = f.nome_fun
        txtbxNome.tag = f.cod_fun
        lblcpf.caption = f.cpf_fun
        txtbxNum_Doc.SetFocus
    End If
End Sub

Private Sub cmndbtnCancelar_Click()
    If MsgBox("Deseja Realmente Cancelar a Distribuição das parcelas? ", _
                vbYesNo + vbCritical) = vbYes Then
        msflxgrd.rows = 1
        frCab.Enabled = True
        frVal.Enabled = False
    End If
End Sub

Private Sub cmndbtnDistr_Click()
    Dim iCount As Integer
    Dim strSQL As String
    Dim iDias As Integer
    Dim cTotal As Currency
    If optnCh.Value Then bCheque = True Else bCheque = False
    If cmbxDias.ListIndex = 1 And CCur(txtbxInterv.Text) > 2 Then
        MsgBox "Intervalo muito grande!", vbExclamation
        cmbxDias.SetFocus
        Exit Sub
    End If
    
    If txtbxNome.tag = "" Then
        MsgBox "Escolha o Terceiro primeiramente!", vbExclamation
        txtbxNome.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Confirma distribuição das parcelas?", vbYesNo + vbQuestion) = vbYes Then
        'iDias = Int(txtbxInterv.Text)
        msflxgrd.rows = 1
        If CCur(txtbxValor_Doc.Text) = 0 Then
            MsgBox "Informe um valor válido!", vbExclamation
            txtbxValor_Doc.SetFocus
            Exit Sub
        End If
        
        
        
        
        For iCount = 1 To Int(txtbxParc.Text)
            If cmbxDias.ListIndex = 0 Then
                iDias = Int(txtbxInterv) * (iCount - IIf(chk.Value = 1, 1, 0))
                strSQL = DateAdd("D", iDias, txtbxVenIni.Text) & Chr(9)
            Else
                iDias = Int(txtbxInterv) * iCount
                strSQL = DateAdd("M", iDias, txtbxVenIni.Text) & Chr(9)
            End If
            strSQL = strSQL & FormatVal(CCur(txtbxValor_Doc.Text) / CCur(txtbxParc.Text)) & Chr(9)
            If bCheque Then
                strSQL = strSQL & "2 - Cheque pré"
            Else
                strSQL = strSQL & "1 - Duplicata"
            End If
            msflxgrd.AddItem strSQL
            msflxgrd.row = msflxgrd.rows - 1
            msflxgrd.col = 11
            msflxgrd.Text = txtbxBoleto.Text
            cTotal = cTotal + Round((CCur(txtbxValor_Doc.Text) / CCur(txtbxParc.Text)), 2)
            
        Next
            If cTotal <> CCur(txtbxValor_Doc.Text) Then
                msflxgrd.row = msflxgrd.rows - 1
                msflxgrd.col = 1
                msflxgrd.Text = FormatVal(CCur(msflxgrd.Text) + CCur(txtbxValor_Doc.Text) - cTotal)
            End If
    End If
    frCab.Enabled = False
    frVal.Enabled = True
End Sub

Private Sub cmndbtnEdit_Click()
    Dim sVen As String
    Dim sVal As String
    Dim sBol As String
    If msflxgrd.rows = 1 Then
        MsgBox "Não há o que ser Editado!", vbExclamation
    Else
        With msflxgrd
            .col = 0
            sVen = .Text
            .col = 1
            sVal = .Text
            .col = 2
            .col = 11
            sBol = .Text
            If Left(cmbxEspecie.Text, 1) = "1" Then
                Dim p As parcela
                p = F_DadosPar(sVen, sVal, sBol)
                
                If p.valor <> 0 Then
                    If (F_Soma - CCur(sVal) + p.valor) > CCur(txtbxValor_Doc.Text) Then
                        MsgBox "O Valor excede o valor total!", vbExclamation
                        Exit Sub
                    End If
                    .col = 0
                    .Text = p.vencimento
                    .col = 1
                    .Text = FormatVal(p.valor)
                    .col = 2
                    .Text = cmbxEspecie.Text
                    .col = 3
                    .Text = ""
                    .col = 4
                    .Text = ""
                    .col = 5
                    .Text = ""
                    .col = 6
                    .Text = ""
                    .col = 7
                    .Text = ""
                    .col = 8
                    .Text = ""
                    .col = 9
                    .Text = ""
                End If
            End If
        End With
    End If
End Sub

Private Sub cmndbtnRem_Click()
    If msflxgrd.rows = 1 Then
        MsgBox "Não há o que ser excluído!", vbExclamation
    Else
        If MsgBox("Deseja excluir Item? ", vbYesNo + vbCritical) = vbYes Then
            If msflxgrd.rows > 2 Then
                msflxgrd.RemoveItem msflxgrd.row
            Else
                msflxgrd.rows = 1
            End If
        
            
        End If
        
    End If
End Sub

Private Sub cmndbtnPlanoContas_Click()
    Dim p As PlanoContas
    p = F_BuscaPlanoContas(Left(lblTipo.caption, 1), "")
    If p.cod_pla <> "" Then
        txtbxClass.Text = p.nome_pla
        txtbxClass.tag = p.cod_pla
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Function F_confMsflx() As Boolean
     Dim iCount As Integer
     Dim sTipo As String
     Dim retorno As Boolean
     Dim strSQL As String
     retorno = True
     strSQL = "os seguintes dados estão faltando: " & Chr(13) & Chr(13)
     With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            If optnPag.Value Then
                .col = 2
                sTipo = Left(.Text, 1)
                If sTipo = "2" Then
                    .col = 3
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & "  número do cheque" & Chr(13)
                        retorno = False
                    End If
                    .col = 4
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " banco" & Chr(13)
                        retorno = False
                    End If
                    
                    
                    
                    
                    
                End If
            
            Else
                .col = 2
                sTipo = Left(.Text, 1)
                If sTipo = "2" Then
                    .col = 3
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & "  número do cheque" & Chr(13)
                        retorno = False
                    End If
                    .col = 4
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " banco" & Chr(13)
                        retorno = False
                    End If
                    .col = 5
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " agência" & Chr(13)
                        retorno = False
                    End If
                    .col = 6
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " conta" & Chr(13)
                        retorno = False
                    End If
                    .col = 9
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " nome do titular" & Chr(13)
                        retorno = False
                    End If
                    .col = 10
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " CPF do titular" & Chr(13)
                        retorno = False
                    End If
                End If
            End If
        Next
     End With
     F_confMsflx = retorno
     If retorno = False Then
        MsgBox strSQL, vbExclamation
        
     End If
End Function
Private Sub cmndbtnSalvar_Click()
    Dim sSEQ_DOC As String
    Dim sSEQ_ITD As String
    Dim sSeq_che As String
    Dim iCount As Integer
    Dim sVen As String, sVal As String, sEsp As String, sNum_Che As String, sBco As String
    Dim sAge As String, sCCO As String, sNome_che As String
    Dim sCPF As String
    Dim sBol As String
    Dim iDoc As String
    If Not F_confMsflx Then Exit Sub
    If msflxgrd.rows = 1 Then
        MsgBox "Os valores não estão distribuídos!", vbExclamation
        Exit Sub
    End If
    
    If F_Soma <> CCur(txtbxValor_Doc.Text) Then
        MsgBox "Os valores não estão corretamente distribuídos!", vbExclamation
        Exit Sub
    ElseIf F_Vazio(txtbxNum_Doc.Text) Then
        MsgBox "Informe o número do Documento!", vbExclamation
        txtbxNum_Doc.SetFocus
        Exit Sub
    End If
    sSEQ_DOC = doc.F_GravaDoc(txtbxHist.Text, txtbxNum_Doc.Text, txtbxValor_Doc.Text, _
        CStr(txtbxRef.Text), txtbxNome.tag, Left(lblTipo.caption, 1), "", msflxgrd.rows - 1, _
         "F", Left(cmbxTipoTer.Text, 1), txtbxClass.tag)
    
    
    If sSEQ_DOC <> "" Then
        With msflxgrd
            For iCount = 1 To .rows - 1
                .row = iCount
                msflxgrd.col = 0
                sVen = msflxgrd.Text
                msflxgrd.col = 1
                sVal = msflxgrd.Text
                msflxgrd.col = 2
                sEsp = Left(msflxgrd.Text, 1)
                msflxgrd.col = 3
                sNum_Che = msflxgrd.Text
                msflxgrd.col = 4
                sBco = msflxgrd.Text
                msflxgrd.col = 5
                sAge = msflxgrd.Text
                msflxgrd.col = 6
                sCCO = msflxgrd.Text
                msflxgrd.col = 9
                sNome_che = msflxgrd.Text
                msflxgrd.col = 10
                sCPF = msflxgrd.Text
                msflxgrd.col = 11
                sBol = msflxgrd.Text
                If sEsp = "1" Then
                    sSEQ_ITD = doc.F_GravaItDoc(sSEQ_DOC, sVal, sVen, sVen, CStr(iCount), a_receber_avulso, sBol)
                ElseIf sEsp = "2" Then
                    If optnReceb.Value = True Then 'cheque a receber
                        sSeq_che = F_GravaCh(sSEQ_DOC, sNum_Che, sBco, sAge, sCCO, _
                        sNome_che, sVal, sVen, txtbxNome.tag, Left(lblTipo.caption, 1), _
                        "", "", CStr(iCount), "", sCPF)
                    Else 'cheque a pagar
                        sSeq_che = F_GravaChePag(sSEQ_DOC, sSEQ_DOC, sNum_Che, sBco, sVal, sVen, _
                        txtbxNome.tag, CStr(iCount))
                    End If
                End If
                
            Next
        End With
    Else
        MsgBox "Erro na gravação!", vbExclamation
        Exit Sub
    End If
    APIInserirDoc sSEQ_DOC
    MsgBox "Documento Gravado!", vbInformation
    msflxgrd.rows = 1
    txtbxValor_Doc.Text = "0,00"
    txtbxParc.Text = "1"
    Form_Load
    frCab.Enabled = True
    frVal.Enabled = False
    txtbxNome.SetFocus
    
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        cmndbtnSalvar_Click
    ElseIf KeyCode = vbKeyF5 Then
        msflxgrd.SetFocus
    End If
End Sub

Private Sub Form_Load()
    
    bCheque = False
    txtbxRef.Text = Date
    cmbxDias.ListIndex = 1
    txtbxNum_Doc.Text = MaxKeySQL("tab_doc", "num_doc", "num_doc like '_AV%'", "_AV0######")
    P_CentraForm Me, 600
    msflxgrd.rows = 1
    msflxgrd.Cols = 12
    msflxgrd.col = 0
    msflxgrd.Text = "Vencimento"
    msflxgrd.col = 1
    msflxgrd.Text = "Valor"
    msflxgrd.col = 2
    msflxgrd.Text = "Espécie"
    msflxgrd.col = 3
    msflxgrd.Text = "Cheque Nº"
    msflxgrd.col = 4
    msflxgrd.Text = "Ban/Con"
    msflxgrd.col = 5
    msflxgrd.Text = "Agencia"
    msflxgrd.col = 6
    msflxgrd.Text = "Conta C."
    msflxgrd.col = 7
    msflxgrd.Text = "Valor"
    msflxgrd.col = 8
    msflxgrd.Text = "Vencimento"
    msflxgrd.col = 9
    msflxgrd.Text = "Titular"
    msflxgrd.col = 10
    msflxgrd.Text = "CPF"
    msflxgrd.col = 11
    msflxgrd.Text = "Boleto"
    msflxgrd.ColAlignment(2) = 1
    txtbxVenIni.Text = Date
    msflxgrd.ColWidth(0) = 1000
    
    msflxgrd.ColWidth(1) = 1000
    msflxgrd.ColWidth(2) = 1500
End Sub

Private Sub msflxGrd_GotFocus()
    msflxgrd.BorderStyle = flexBorderSingle
    
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
        With msflxgrd
            If .col = 2 Then
                If KeyCode = vbKeySpace Then
                    cmbxOpc.Left = .Left + .CellLeft
                    cmbxOpc.top = .top + .CellTop
                    cmbxOpc.width = .CellWidth
                    cmbxOpc.Text = msflxgrd.Text
                    msflxgrd.SetFocus
                    cmbxOpc.Visible = True
                
                End If
                
            End If
            
        End With
        If KeyCode = vbKeyReturn Then
            msflxgrd_DblClick
        ElseIf KeyCode = vbKey1 Or KeyCode = 97 Then
            msflxgrd.col = 2
            msflxgrd.Text = "1 - Duplicata"
            msflxgrd_DblClick
        ElseIf KeyCode = vbKey2 Or KeyCode = 98 Then
            msflxgrd.col = 2
            msflxgrd.Text = "2 - Cheque Pré"
            msflxgrd_DblClick
        ElseIf KeyCode = vbKeyAdd Or KeyCode = 187 Then
            cmndbtnAdd_Click
        ElseIf KeyCode = vbKeyF5 Then
            msflxgrd.SetFocus
        ElseIf KeyCode = vbKeyDelete Then
            cmndbtnRem_Click
        End If
        
    
    
    Exit Sub
    
End Sub

Private Sub msflxGrd_LostFocus()
    msflxgrd.BorderStyle = flexBorderNone
End Sub

Private Sub optnCH_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub optnDup_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub optnPag_Click()
    lblTipo.caption = "P - Contas a Pagar"
    txtbxNome.tag = ""
    
    txtbxNome.Text = ""
    cmbxTipoTer.Clear
    cmbxTipoTer.AddItem "Fornecedor"
    cmbxTipoTer.ListIndex = 0
End Sub

Private Sub optnPag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub optnReceb_Click()
    lblTipo.caption = "R - Contas a Receber"
    txtbxNome.tag = ""
    txtbxNome.Text = ""
    cmbxTipoTer.Clear
    cmbxTipoTer.AddItem "Cliente"
    cmbxTipoTer.AddItem "Representante"
    cmbxTipoTer.ListIndex = 0
    
End Sub

Private Sub optnReceb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxBoleto_GotFocus()
    P_SelecTexto txtbxBoleto
End Sub

Private Sub txtbxBoleto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxBoleto_LostFocus()
    P_AjustaTexto txtbxBoleto
End Sub

Private Sub txtbxData_GotFocus()
    P_SelecTexto txtbxData
End Sub

Private Sub txtbxData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxData_LostFocus()
    P_AjustaData txtbxData
End Sub

Private Sub txtbxHist_GotFocus()
    P_SelecTexto txtbxHist
End Sub


Private Sub txtbxHist_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxHist_LostFocus()
    P_AjustaTexto txtbxHist
End Sub

Private Sub txtbxInterv_GotFocus()
    P_SelecTexto txtbxInterv
End Sub

Private Sub txtbxInterv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        txtbxNome_LostFocus
        Sendkeys Chr(9)
    End If

End Sub

Private Sub txtbxInterv_LostFocus()
    P_ValorLostFocus txtbxInterv, "###"
    If Int(txtbxInterv.Text) < 1 Then
        MsgBox "Informe o intervalo corretamente!", vbExclamation
        txtbxInterv.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub txtbxNome_GotFocus()
    P_SelecTexto txtbxNome
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxNome_LostFocus
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxNome_LostFocus()
    P_AjustaTexto txtbxNome
End Sub

Private Sub txtbxNum_Doc_GotFocus()
    P_SelecTexto txtbxNum_Doc
End Sub

Private Sub txtbxNum_Doc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxNum_Doc_LostFocus()
    P_AjustaTexto txtbxNum_Doc
    If Left(txtbxNum_Doc, 3) = "_AV" Then
        txtbxNum_Doc.Text = MaxKeySQL("tab_doc", "num_doc", "num_doc like '_AV%'", "_AV0######")
    End If
    buscaDoc
End Sub
Sub buscaDoc()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    If txtbxNum_Doc.Text <> "" Then
        strSQL = "select * from tab_Doc where num_doc='" & txtbxNum_Doc.Text & "' "
        strSQL = strSQL & " and cod_Ter='" & txtbxNome.tag & "' "
        abreObjetoAdo rPesq, strSQL
        With rPesq
            If .RecordCount > 0 Then
                strSQL = "Já existe (m) " & .RecordCount & " documento (s) para este "
                
                strSQL = strSQL & IIf(optnPag.Value, "Fornecedor", "Cliente") & " com este Número!"
                strSQL = strSQL & Chr(13) & "Deseja continuar?"
                If MsgBox(strSQL, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                    txtbxNum_Doc.Text = ""
                    txtbxNum_Doc.SetFocus
                End If
                
            End If
            .Close
        End With
    
    End If
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Private Sub txtbxParc_GotFocus()
    P_SelecTexto txtbxParc
End Sub

Private Sub txtbxParc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxNome_LostFocus
        Sendkeys Chr(9)
    End If

End Sub

Private Sub txtbxParc_LostFocus()
    P_ValorLostFocus txtbxParc, "###"
    If Int(txtbxParc.Text) < 1 Then
        MsgBox "informe um valor decente", vbExclamation
        txtbxParc.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub txtbxRef_GotFocus()
    P_SelecTexto txtbxRef
End Sub

Private Sub txtbxRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxRef_LostFocus()
    P_AjustaData txtbxRef
End Sub

Private Sub txtbxValor_GotFocus()
    P_SelecTexto txtbxValor
End Sub

Private Sub txtbxValor_Doc_GotFocus()
    P_SelecTexto txtbxValor_Doc
End Sub

Private Sub txtbxValor_Doc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxNome_LostFocus
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_Doc_LostFocus()
    P_ValorLostFocus txtbxValor_Doc
End Sub

Private Sub txtbxValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_LostFocus()
    P_ValorLostFocus txtbxValor
End Sub

Private Sub txtbxVenIni_GotFocus()
    P_SelecTexto txtbxVenIni
End Sub

Private Sub txtbxVenIni_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSQL As String
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        txtbxNome_LostFocus
        Sendkeys Chr(9)
    End If
    If KeyCode = vbKeyF2 Then
        strSQL = InputBox("Quantos dias deseja adicionar?", "Soma de Dias", "0")
        strSQL = F_SoNumero(strSQL)
        txtbxVenIni.Text = DateAdd("D", CDbl(strSQL), txtbxVenIni.Text)
        
    End If
End Sub

Private Sub txtbxVenIni_LostFocus()
    P_AjustaData txtbxVenIni
End Sub
Private Sub msflxgrd_DblClick()
    cmbxItem_GotFocus
    cmbxItem.Text = Format(msflxgrd.row, "00#")
    If optnReceb.Value = True Then
        cmbxEspecie_Click
    End If
    msflxgrd.col = 2
    cmndbtnEdit_Click
    
End Sub

Private Sub cmbxEspecie_Click()
    Dim c As cheque
    Dim sVal As String
    Dim sNum As String
    Dim sVen As String
    Dim sNome As String
    Dim sBan As String
    Dim sCCO As String
    Dim sAge As String
    Dim cpp As chePag
    Dim sCPF As String
    Dim sBol As String
    If msflxgrd.row = -1 Or msflxgrd.rows <= 1 Then
        Exit Sub
    End If
    
    If cmbxEspecie.ListIndex = 0 Then
        With msflxgrd
            .col = 0
            If .Text <> txtbxVenIni.Text Then
                
            End If
        End With
    ElseIf cmbxEspecie.ListIndex = 0 Then
        cmndbtnEdit_Click
    Else
        If cmbxEspecie.ListIndex = 1 Then
            
            
        
            
            msflxgrd.col = 0
            sVen = msflxgrd.Text
            msflxgrd.col = 1
            sVal = msflxgrd.Text
            msflxgrd.col = 3
            sNum = msflxgrd.Text
            msflxgrd.col = 4
            sBan = msflxgrd.Text
            msflxgrd.col = 5
            sAge = msflxgrd.Text
            msflxgrd.col = 6
            sCCO = msflxgrd.Text
            msflxgrd.col = 7
            sNome = msflxgrd.Text
            msflxgrd.col = 10
            sCPF = msflxgrd.Text
            If optnReceb.Value = True Then
                c = F_DadosCheque(True, sVen, txtbxNome.Text, sNum, sCCO, sAge, sBan, sVal, sCPF)
            Else
                cpp = F_DadosChePag(sVen, sNum, sCCO, sVal)
            End If
            If c.num_che = "" And cpp.num_che = "" Then
                MsgBox "Não foi informado um Cheque Válido!", vbExclamation
                Exit Sub
            Else
                msflxgrd.col = 2
                msflxgrd.Text = cmbxEspecie.Text
            End If
            
        End If
        
        With msflxgrd
           If c.num_che <> "" Then
                .col = 0
                .Text = c.Venc_che
                .col = 1
                .Text = FormatVal(c.Valor_che)
                
                
                .col = 3
                .Text = c.num_che
                .col = 4
                .Text = c.cod_bco
                .col = 5
                .Text = c.age_che
                .col = 6
                .Text = c.cco_che
                .col = 7
                .Text = c.Valor_che
                .col = 8
                .Text = c.Venc_che
                .col = 9
                .Text = c.nome_che
                .col = 10
                .Text = c.cpf_che
            End If
            If cpp.num_che <> "" Then
                .col = 0
                .Text = cpp.Venc_che
                .col = 1
                .Text = FormatVal(cpp.Valor_che)
                
                
                .col = 3
                .Text = cpp.num_che
                .col = 4
                .Text = cpp.cod_con
                
            End If
        End With
    End If
    
End Sub
Private Sub cmbxItem_Click()
    msflxgrd.row = Int(cmbxItem.Text)
    msflxgrd.col = 2
    cmbxEspecie.ListIndex = Int(Left(msflxgrd.Text, 1)) - 1
    msflxgrd.col = 0
    txtbxData.Text = msflxgrd.Text
    msflxgrd.col = 1
    txtbxValor.Text = msflxgrd.Text
End Sub

Private Sub cmbxItem_GotFocus()
    cmbxItem.Clear
    For iCount = 1 To msflxgrd.rows - 1
        cmbxItem.AddItem Format(iCount, "0##")
    Next
    
End Sub

Private Sub cmbxNat_Click()
    If cmbxNat.ListIndex = 1 Then
        Dim c As cheque
        c = F_DadosCheque(False, Date, lblNome_Che.caption, lblNum_Che.caption)
        If c.num_che <> "" Then
            lblAge_Che.caption = c.age_che
            lblBco.caption = c.cod_bco
            lblCco_Che.caption = c.cco_che
            lblNome_Che.caption = c.nome_che
            lblNum_Che.caption = c.num_che
            lblValor_Che.caption = c.Valor_che
            lblVenc_Che.caption = c.Venc_che
            frCheque.Visible = True
        End If
    End If
End Sub



