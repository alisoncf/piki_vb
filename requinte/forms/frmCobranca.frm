VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCobranca 
   Caption         =   "Cobrança"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   Icon            =   "frmCobranca.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   9915
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   960
      ScaleHeight     =   585
      ScaleWidth      =   16545
      TabIndex        =   61
      Top             =   1320
      Width           =   16575
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   13680
         ScaleHeight     =   585
         ScaleWidth      =   2745
         TabIndex        =   74
         Top             =   0
         Width           =   2775
         Begin VB.CommandButton cmdPredatarTodos 
            Caption         =   "Confirmar"
            Height          =   495
            Left            =   1560
            TabIndex        =   77
            Top             =   0
            Width           =   975
         End
         Begin VB.TextBox txtPredata 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   75
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pré-Datar todos"
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Tag             =   "!"
            Top             =   0
            Width           =   1110
         End
      End
      Begin VB.CheckBox chkDataIni 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Data Inicial"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4200
         TabIndex        =   73
         Top             =   120
         Width           =   1215
      End
      Begin VB.CheckBox chkDesm 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Desmarcados"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6840
         TabIndex        =   69
         Top             =   120
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnDesmarcar 
         Caption         =   "Desmascar"
         Height          =   375
         Left            =   12000
         TabIndex        =   68
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnMarcar 
         Caption         =   "Marcar"
         Height          =   375
         Left            =   10440
         TabIndex        =   67
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnDesmarc 
         Caption         =   "Desmarcar Todos"
         Height          =   375
         Left            =   8880
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chkCarnes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Marcados"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5520
         TabIndex        =   65
         Top             =   120
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.OptionButton optnAlfabet 
         Caption         =   "Alfabético"
         Height          =   195
         Left            =   2400
         TabIndex        =   64
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optnPorDoc 
         Caption         =   "Doc"
         Height          =   195
         Left            =   1440
         TabIndex        =   63
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optnPorVenc 
         Caption         =   "Vencimento"
         Height          =   195
         Left            =   120
         TabIndex        =   62
         Top             =   120
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   6255
      Left            =   0
      TabIndex        =   41
      Top             =   1920
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   11033
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   9885
      TabIndex        =   40
      Top             =   9360
      Width           =   9915
      Begin VB.CommandButton cmdLib 
         Caption         =   "Liberar Todos"
         Height          =   375
         Left            =   7080
         TabIndex        =   72
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdBloq 
         Caption         =   "Bloquear Todos"
         Height          =   375
         Left            =   7080
         TabIndex        =   71
         Top             =   0
         Width           =   1335
      End
      Begin VB.CheckBox chkTodas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Todas as Parcelas do Documento"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   600
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CommandButton cmndbrnTransf 
         Caption         =   "Transferir"
         Height          =   495
         Left            =   3600
         Picture         =   "frmCobranca.frx":1CCA
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtbxTransfCob 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   2775
      End
      Begin VB.PictureBox pcBotoes 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   8640
         ScaleHeight     =   825
         ScaleWidth      =   8625
         TabIndex        =   43
         Top             =   0
         Width           =   8655
         Begin VB.CommandButton cmndbtnImpr 
            Caption         =   "Im&primir"
            Height          =   810
            Left            =   3960
            Picture         =   "frmCobranca.frx":2254
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   0
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnEditar 
            Caption         =   "Editar"
            Height          =   810
            Left            =   120
            Picture         =   "frmCobranca.frx":2661
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   0
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnFechar 
            Cancel          =   -1  'True
            Caption         =   "&Sair"
            Height          =   810
            Left            =   5880
            Picture         =   "frmCobranca.frx":3063
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   0
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnImprimir 
            Caption         =   "Im&primir HTML"
            Height          =   810
            Left            =   2040
            Picture         =   "frmCobranca.frx":31CE
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmndbtnBuscaCobTransf 
         Height          =   285
         Left            =   2880
         Picture         =   "frmCobranca.frx":35DB
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkProrr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Apenas Parcelas Prorrogadas"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3000
         TabIndex        =   60
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label lblListados 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   6720
         TabIndex        =   54
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listados"
         Height          =   195
         Left            =   5520
         TabIndex        =   53
         Top             =   360
         Width           =   585
      End
      Begin VB.Label lblSelect 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   6720
         TabIndex        =   52
         Top             =   120
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selecionados:"
         Height          =   195
         Left            =   5520
         TabIndex        =   51
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transferir para Conta Cobrança:"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   0
         Width           =   2265
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1305
      ScaleWidth      =   9885
      TabIndex        =   39
      Top             =   0
      Width           =   9915
      Begin VB.ComboBox cmbxCreditoSuspenso 
         Height          =   315
         ItemData        =   "frmCobranca.frx":3965
         Left            =   14160
         List            =   "frmCobranca.frx":3972
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cmbxStatus 
         Height          =   315
         ItemData        =   "frmCobranca.frx":3982
         Left            =   14160
         List            =   "frmCobranca.frx":398F
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10680
         TabIndex        =   32
         Top             =   840
         Width           =   1575
      End
      Begin VB.PictureBox pcLista 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7920
         Picture         =   "frmCobranca.frx":39AA
         ScaleHeight     =   195
         ScaleWidth      =   180
         TabIndex        =   70
         Top             =   600
         Width           =   180
      End
      Begin VB.TextBox txtbxOrgao 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Height          =   285
         Left            =   8760
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtbxContaSug 
         Appearance      =   0  'Flat
         BackColor       =   &H00EF9C7C&
         Height          =   285
         Left            =   4560
         TabIndex        =   55
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Pesquisar"
         Height          =   570
         Left            =   15600
         Picture         =   "frmCobranca.frx":3AA3
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9480
         MaxLength       =   10
         TabIndex        =   30
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8280
         MaxLength       =   10
         TabIndex        =   28
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtbxTerceiro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtbxSeqItd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         MaxLength       =   60
         TabIndex        =   24
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtbxSeqDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         MaxLength       =   60
         TabIndex        =   22
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtbxNumDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6960
         MaxLength       =   60
         TabIndex        =   26
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtbxGrupoCli 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6960
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cmbxOrdenarPor 
         Height          =   315
         ItemData        =   "frmCobranca.frx":3B9C
         Left            =   12600
         List            =   "frmCobranca.frx":3BAC
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cmbxTipoTer 
         Height          =   315
         ItemData        =   "frmCobranca.frx":3BDF
         Left            =   12600
         List            =   "frmCobranca.frx":3BE9
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cmbxEnd 
         Height          =   315
         ItemData        =   "frmCobranca.frx":3C05
         Left            =   10680
         List            =   "frmCobranca.frx":3C15
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtbxContaCob 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnContaCob 
         Height          =   285
         Left            =   4200
         Picture         =   "frmCobranca.frx":3C4E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtbxTipoCob 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnCob 
         Height          =   285
         Left            =   1800
         Picture         =   "frmCobranca.frx":3FD8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnGrupo 
         Height          =   285
         Left            =   8400
         Picture         =   "frmCobranca.frx":4362
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnTerc 
         Height          =   285
         Left            =   4200
         Picture         =   "frmCobranca.frx":46EC
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnBuscaContaSuj 
         Height          =   285
         Left            =   6600
         Picture         =   "frmCobranca.frx":4A76
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnOrg 
         Height          =   285
         Left            =   10320
         Picture         =   "frmCobranca.frx":4E00
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdCidade 
         Height          =   285
         Left            =   12240
         Picture         =   "frmCobranca.frx":518A
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crédito Suspenso:"
         Height          =   195
         Left            =   14160
         TabIndex        =   36
         Top             =   600
         Width           =   1290
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   195
         Left            =   14160
         TabIndex        =   16
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         Height          =   195
         Left            =   10680
         TabIndex        =   31
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lblOrgao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Órgão:"
         Height          =   195
         Left            =   8760
         TabIndex        =   9
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conta Cobrança Sugerida:"
         Height          =   195
         Left            =   4560
         TabIndex        =   57
         Top             =   0
         Width           =   1875
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   9480
         TabIndex        =   29
         Tag             =   "!"
         Top             =   600
         Width           =   285
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   8280
         TabIndex        =   27
         Tag             =   "!"
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id. Parcela:"
         Height          =   195
         Left            =   5760
         TabIndex        =   23
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id. Doc:"
         Height          =   195
         Left            =   4560
         TabIndex        =   21
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terceiro:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documento:"
         Height          =   195
         Left            =   6960
         TabIndex        =   25
         Top             =   600
         Width           =   870
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo de Cientes:"
         Height          =   195
         Left            =   6960
         TabIndex        =   6
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenar Por:"
         Height          =   195
         Left            =   12600
         TabIndex        =   34
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terceiro:"
         Height          =   195
         Left            =   12600
         TabIndex        =   14
         Top             =   0
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listar Endereços:"
         Height          =   195
         Left            =   10680
         TabIndex        =   12
         Top             =   0
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conta Cobrança Atual:"
         Height          =   195
         Left            =   2160
         TabIndex        =   3
         Top             =   0
         Width           =   1605
      End
      Begin VB.Label lblTipoCob 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cobrança:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmndbtnTodos 
      Caption         =   "Todos"
      Height          =   615
      Left            =   0
      TabIndex        =   58
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "frmCobranca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rPesq As New ADODB.Recordset
Dim sListaDoc As String
Public Sub mnEdiParc()

End Sub
Public Sub mnSelec()

End Sub
Public Sub mnSelecTodos()
    
End Sub

Private Sub chkCarnes_Click()
    busca
End Sub

Private Sub chkDesm_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub chkProrr_Click()
    busca
End Sub

Private Sub cmbxOrdenarPor_Click()
    busca
End Sub

Private Sub cmdBloq_Click()
     Dim strSQL As String
     If strLiberarCli_usu <> "S" Then
        MsgBox "Não há permissão a este acesso!", vbExclamation
        Exit Sub
     End If
     If MsgBox("Confirma bloqueio de todos os clientes listados abaixo?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
        Exit Sub
     End If
     With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 1
            If .Text = "S" Then
                .col = 9
                Clientes.F_AlteraStatusCli .Text, "2", ""
                
                
            End If
        Next
        
    End With
End Sub

Private Sub cmdCidade_Click()
    Dim c As Cidade
    c = Cidades.F_BuscaCidades(txtCidade.Text)
    txtCidade.Text = c.nome_cid
    txtCidade.tag = c.cod_cid
    
End Sub

Private Sub cmdLib_Click()
    Dim strSQL As String
     If strLiberarCli_usu <> "S" Then
        MsgBox "Não há permissão a este acesso!", vbExclamation
        Exit Sub
     End If
     If MsgBox("Confirma desbloqueio de todos os clientes listados abaixo?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
        Exit Sub
     End If
     With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 1
            If .Text = "S" Then
                .col = 9
                Clientes.F_AlteraStatusCli .Text, "1", ""
                
                
            End If
        Next
    End With
End Sub

Private Sub cmdPredatarTodos_Click()
    Predata
End Sub

Private Sub cmndbrnTransf_Click()
    If txtbxTransfCob.tag = "" Then
        MsgBox "Escolha a conta que deseja Transferir", vbExclamation
        txtbxTransfCob.SetFocus
        Exit Sub
        
    End If
    If MsgBox("Confirma Transferência dos " & lblSelect.caption & " para " & txtbxTransfCob.Text & "?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    transfere
    txtbxTransfCob.Text = ""
    txtbxTransfCob.tag = ""
    txtbxTransfCob.SetFocus
End Sub
Sub Predata()
    Dim strSQL As String
    Dim iCount As Integer
    Dim Anterior As String
    Dim sSeqItd As String
    Dim venc As String
    Dim doc As String
    
    
    Dim tipo As String
    If strBaixadoc_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbCritical
        Exit Sub
    End If
    
    If MsgBox("Confirma a alteração do vencimento de todas as parcelas anteriores a hoje para o dia " & txtPredata.Text & "?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 6
            venc = .Text
            .col = 4
            doc = .Text
            .col = 3
            sSeqItd = .Text
            .col = 1
            
            If .Text = "S" Then
            
                 
                 
                 'P_GravaItDocPre sSeqItd, CStr(Date), txtPredata, "Predatou na cobrança"
                 
                 P_AlteraDocsAnteriores doc, txtPredata.Text, txtPredata.Text
                
            End If
   
        Next
        
    End With
    busca
    
    
End Sub

Sub transfere()
    Dim strSQL As String
    Dim iCount As Integer
    Dim Anterior As String
    Dim sSeqItd As String
    Dim tipo As String
    Dim id_doc As String
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 1
            If .Text = "S" Then
                If chkTodas.Value = 0 Then
                    strSQL = "update tab_itdoc set cod_cob='" & txtbxTransfCob.tag & "' "
                    .col = 3
                    strSQL = strSQL & "where seq_itd='" & .Text & "' "
                    Conn.Execute strSQL
                    
                Else
                    strSQL = "update tab_itdoc set cod_cob='" & txtbxTransfCob.tag & "' "
                    .col = 4
                    strSQL = strSQL & "where seq_doc='" & .Text & "' "
                    strSQL = strSQL & " and reman_itd>0"
                    Conn.Execute strSQL
                                        
                    
                    
                End If
                .col = 22
                id_doc = .Text
                .col = 16
                tipo = .Text
                .col = 20
                Anterior = .Text
                .col = 4
                doc.P_gravaDocHist id_doc, Anterior, txtbxTransfCob.tag, tipo, tipo
            End If
        Next
        
    End With
    busca
    
    
End Sub
Sub Bloqueia(opcao As Boolean)
    Dim strSQL As String
    Dim iCount As Integer
    Dim Anterior As String
    Dim sSeqItd As String
    Dim tipo As String
    
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 1
            
            
            
        Next
        
    End With
    busca
    
    
End Sub

Private Sub cmndbtnBuscaCobTransf_Click()
        Dim c As ContaCob
    c = F_BuscaContaCob(txtbxTransfCob.Text)
    txtbxTransfCob.Text = c.nome_cob
    txtbxTransfCob.tag = c.cod_cob

End Sub

Private Sub cmndbtnBuscaContaSuj_Click()
    Dim c As ContaCob
    c = F_BuscaContaCob(txtbxContaSug.Text)
    txtbxContaSug.Text = c.nome_cob
    txtbxContaSug.tag = c.cod_cob
    
End Sub

Private Sub cmndbtnBuscar_Click()
    P_Status "Aguarde"
    busca
    
    P_Status ""
    
End Sub
Sub busca()
    Dim strSQL As String
    Dim s As setor
    Dim o As orgao.orgao
    Dim sTab As String
    Dim sCod As String
    Dim sNome As String
    Dim sTipoTer As String
    Dim sEnd As String
    Dim sNumDoc As String
    'On Error GoTo trata_erro
    msflxgrd.Visible = False
    
    
    
    
    msflxgrd.rows = 1
    If cmbxOrdenarPor.ListIndex = 1 Then
        sOrderBy = "num_doc"
    ElseIf cmbxOrdenarPor.ListIndex = 2 Then
        sOrderBy = "vencini_Itd,nome"
    ElseIf cmbxOrdenarPor.ListIndex = 3 Then
        sOrderBy = "nome,venc_itd"
    Else
        sOrderBy = "venc_itd,nome"
    End If
        
    
    
    lblSelect.caption = "0"
    
    If cmbxTipoTer.ListIndex = 0 Then
        sTab = "tab_clientes"
        sCod = "cod_cli"
        sNome = "nome_cli"
        sTipoTer = "C"
        sEnd = "concat(logradouro_cli,' ',numero_cli,' ',complemento_cli)"
    Else
        sTab = "tab_representantes"
        sNome = "nome_rep"
        sCod = "cod_rep"
        sEnd = "endereco_rep"
        sTipoTer = "R"
    End If
    
    strSQL = "select (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as dias, tcs.nome_cob as nomecobsuj,tcb.nome_cob,"
    strSQL = strSQL & "tc." & sCod & " as codigo, " & sEnd & " as endereco,  tc.cod_org, tc.cobranca_cli,ts.nome_set,td.num_ven,tc." & sNome & " as nome,td.historico_doc,td.parcelas_doc,marcado_itd, "
    strSQL = strSQL & "td.parcelas_doc,td.num_doc,td.data_doc, td.valor_doc,tci.nome_cid, "
    strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as valor_juro, "
    strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) + reman_itd as sugerido, ti.* "
    
    strSQL = strSQL & " from tab_doc as td inner join " & sTab & " as tc  on td.cod_ter=tc." & sCod
    strSQL = strSQL & " inner join tab_itdoc as ti on td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " left join tab_setores as ts on tc.cod_set=ts.cod_set "
    strSQL = strSQL & " left join tab_cidades as tci on tci.cod_cid=ts.cod_cid "
    strSQL = strSQL & " left join tab_contacob as tcb  on ti.cod_cob=tcb.cod_cob "
    strSQL = strSQL & " left join tab_contacob as tcs on ts.cobrador_set=tcs.cod_cob "
    strSQL = strSQL & " left join tab_orgaos as torg on torg.cod_org=tc.cod_org"
    
    strSQL = strSQL & " ,tab_parametros as tp"
    
    
    
   ' strSQL = strSQL & " where tc." & sCod & "=td.cod_ter and td.seq_doc=ti.seq_doc "
    
    
    
    
    strSQL = strSQL & " WHERE td.tipo_doc='R' "
    strSQL = strSQL & " and td.tipo_ter='" & sTipoTer & "' "
    
    If Left(cmbxEnd.Text, 1) = "1" Then
        strSQL = strSQL & " and tc.cod_set=ts.cod_set "
        strSQL = strSQL & " and cobranca_cli='1'"
    ElseIf Left(cmbxEnd.Text, 1) = "2" Then
        strSQL = strSQL & " and torg.cod_set=ts.cod_set "
        strSQL = strSQL & " and cobranca_cli='2'"
    
    Else
        If txtbxOrgao.tag <> "" Then
          '  strSQL = strSQL & " and if(tc.cobranca_cli='2',torg.cod_set,tc.cod_set)=ts.cod_set "
        Else
          '  strSQL = strSQL & " and if(tc.cobranca_cli='2',setortrab_cli,tc.cod_set)=ts.cod_set "
        End If
    End If
    
    
    If txtbxGrupoCli.tag <> "" And cmbxTipoTer.ListIndex = 0 Then
        
        strSQL = strSQL & " and tc.cod_gru='" & txtbxGrupoCli.tag & "' "
        
    End If
    If chkProrr.Value = 1 Then
        strSQL = strSQL & " and venc_itd<>vencini_itd "
    End If
    
    strSQL = strSQL & " and ti.reman_itd > 0 "
    If sListaDoc <> "" Then
        strSQL = strSQL & " and num_doc in (" & sListaDoc & ")"
    End If
    If txtbxContaCob.tag <> "" Then
        strSQL = strSQL & " and ti.cod_cob='" & txtbxContaCob.tag & "' "
    End If
    If txtbxContaSug.tag <> "" Then
        strSQL = strSQL & " and ts.cobrador_set='" & txtbxContaSug.tag & "' "
    End If
    If txtbxOrgao.tag <> "" Then
        strSQL = strSQL & " and tc.cod_org='" & txtbxOrgao.tag & "' "
    End If
    
    If chkDataIni.Value = 1 Then
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and ti.vencini_itd <= " & F_GravaData(txtbxdataFin.Text)
        End If
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and ti.vencini_itd >= " & F_GravaData(txtbxDataIni.Text)
        End If
    Else
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and ti.venc_itd <= " & F_GravaData(txtbxdataFin.Text)
        End If
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and ti.venc_itd >= " & F_GravaData(txtbxDataIni.Text)
        End If
    End If
    If txtbxTipoCob.tag <> "" Then
        strSQL = strSQL & " and ti.cod_tip='" & txtbxTipoCob.tag & "' "
    End If
    
    If txtCidade.tag <> "" Then
        strSQL = strSQL & " and tci.cod_cid='" & txtCidade.tag & "' "
    End If
    
    
    If txtbxNumDoc.Text <> "" Then
        strSQL = strSQL & " and num_doc='" & txtbxNumDoc.Text & "' "
    End If
    If txtbxSeqDoc.Text <> "" Then
        strSQL = strSQL & " and td.seq_doc='" & txtbxSeqDoc.Text & "' "
    End If
    If txtbxSeqItd.Text <> "" Then
        strSQL = strSQL & " and ti.seq_itd='" & txtbxSeqDoc.Text & "' "
    End If
    If txtbxTerceiro.tag <> "" Then
        strSQL = strSQL & " and td.cod_ter='" & txtbxTerceiro.tag & "' "
    End If
    
    If chkCarnes.Value = 1 Then
        strSQL = strSQL & " and (marcado_itd<>'' "
    Else
        strSQL = strSQL & " and (marcado_itd='' "
    End If
    If chkDesm.Value = 1 Then
        strSQL = strSQL & " or marcado_itd='' or marcado_itd is null "
    
    Else
        
    End If
    strSQL = strSQL & ")"
    
    If cmbxTipoTer.ListIndex = 0 Then
        If cmbxStatus.Text <> "" Then
            strSQL = strSQL & " and status_cli='" & cmbxStatus.ListIndex & "'"
        End If
        
         If cmbxCreditoSuspenso.Text <> "" Then
            strSQL = strSQL & " and creditosuspenso_cli='" & Left(cmbxStatus.ListIndex, 1) & "'"
        End If
    End If
    
    
    strSQL = strSQL & " order by " & sOrderBy
        
        

    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            If .AbsolutePosition = 37 Then
            '    MsgBox ""
            End If
            strSQL = .AbsolutePosition + 1 & Chr(9)
            strSQL = strSQL & Chr(9)
            strSQL = strSQL & !num_doc & Chr(9)
            strSQL = strSQL & !seq_itd & Chr(9)
            strSQL = strSQL & !seq_doc & Chr(9)
            strSQL = strSQL & !parcela_itd & "/" & !parcelas_doc & Chr(9)
            strSQL = strSQL & !venc_itd & Chr(9)
            strSQL = strSQL & FormatVal(!valor_itd) & Chr(9)
            strSQL = strSQL & FormatVal(!reman_itd) & Chr(9)
            strSQL = strSQL & !codigo & Chr(9)
            strSQL = strSQL & !nome & Chr(9)
            '
            
            If cmbxEnd.ListIndex = 0 Then
                If !cobranca_cli = "2" Then
                    s = F_RetornaSetorTrab(!codigo)
                    o = orgao.F_BuscaOrgaoPorCod(!cod_org)
                    'strSQL = strSQL & s.Cidade & Chr(9)
                    'strSQL = strSQL & s.nome & Chr(9)
                    'strSQL = strSQL & o.endereco_org & Chr(9)
                    strSQL = strSQL & !nome_cid & Chr(9)
                    strSQL = strSQL & !nome_set & Chr(9)
                    strSQL = strSQL & !Endereco & Chr(9)
                    strSQL = strSQL & "T" & Chr(9)
                Else
                    strSQL = strSQL & !nome_cid & Chr(9)
                    strSQL = strSQL & !nome_set & Chr(9)
                    strSQL = strSQL & !Endereco & Chr(9)
                    strSQL = strSQL & "R" & Chr(9)
                End If
            ElseIf cmbxEnd.ListIndex = 1 Then
                strSQL = strSQL & !Endereco & Chr(9)
                strSQL = strSQL & !nome_set & Chr(9)
                strSQL = strSQL & !Endereco & Chr(9)
                strSQL = strSQL & "R" & Chr(9)
            Else
                
                s = F_RetornaSetorTrab(!codigo)
                strSQL = strSQL & s.Cidade & Chr(9)
                strSQL = strSQL & s.nome & Chr(9)
                strSQL = strSQL & !Endereco & Chr(9)
                strSQL = strSQL & "T" & Chr(9)
                
            End If
            
            strSQL = strSQL & F_BuscaTel(!codigo, sTipoTer) & Chr(9)
            strSQL = strSQL & !cod_tip & Chr(9)
            strSQL = strSQL & !nome_cob & Chr(9) & !nomecobsuj & Chr(9) & !cobranca_cli & Chr(9) & !cod_cob
            strSQL = strSQL & Chr(9) & !marcado_itd & Chr(9) & !seq_doc
            
            msflxgrd.AddItem strSQL
            msflxgrd.row = msflxgrd.rows - 1
            msflxgrd.col = 0
            msflxgrd.CellFontSize = 6
            If .RecordCount > 50 Then
                P_Status "Registro " & .AbsolutePosition & "/" & .RecordCount
            End If
            .MoveNext
        Loop
        .Close
    End With
    lblListados.caption = msflxgrd.rows - 1
    ZebraGrid msflxgrd, 1
    MarcaMarcados
    msflxgrd.Visible = True
    P_Status
    Exit Sub
trata_Erro:
'msgbox err.desc
    msflxgrd.Visible = True
    P_Status
    P_MostraErro
End Sub
Sub MarcaMarcados()
    Dim iCount As Integer
    With msflxgrd
        
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 21
            If .Text <> "" Then
                ZebraLinha msflxgrd, iCount, vbRed
            End If
            
        Next
    End With
End Sub

Private Sub cmndbtnCob_Click()
    Dim t As tipocob
    t = F_BuscaTipoCob(txtbxTipoCob.Text)
    txtbxTipoCob.Text = t.nome_tip
    txtbxTipoCob.tag = t.cod_tip
    
    
End Sub

Private Sub cmndbtnContaCob_Click()
    Dim c As ContaCob
    c = F_BuscaContaCob(txtbxContaCob.Text)
    txtbxContaCob.Text = c.nome_cob
    txtbxContaCob.tag = c.cod_cob
    
    
End Sub

Private Sub cmndbtnDesmarc_Click()
    Dim strSQL As String
    
    
    
    
    If txtbxContaCob.tag = "" Then
        MsgBox "Para desmarcar é necessário informar a conta cobrança!", vbExclamation
        Exit Sub
    End If
    If strBaixadoc_usu <> "S" Then
        MsgBox "Você não tem autorização para esse acesso!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Deseja realmente desmarcar todos os carnês desse cobrador?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    strSQL = "update tab_itdoc set marcado_itd='' where cod_cob='" & txtbxContaCob.tag & "' "
    Conn.Execute strSQL
    
    busca
    
End Sub

Private Sub cmndbtnDesmarcar_Click()
      Dim strSQL As String
    Dim iCount As Integer
   ' If txtbxContaCob.tag = "" Then
   '     MsgBox "Para marcar é necessário informar a conta cobrança!", vbExclamation
   '     Exit Sub
   ' End If
    If strBaixadoc_usu <> "S" Then
        MsgBox "Você não tem autorização para esse acesso!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Deseja marcar os carnês selecionados desse cobrador?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 1
            If .Text = "S" Then
                .col = 4
                strSQL = "update tab_itdoc set marcado_itd='' "
                strSQL = strSQL & " where seq_doc='" & .Text & "' "
                Conn.Execute strSQL
            End If
        Next
    End With
    busca
End Sub

Private Sub cmndbtnEditar_Click()
    Dim fr As New frmEdiLan
    On Error GoTo trata_Erro
    If strBaixadoc_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbCritical
        Exit Sub
    End If
    
        msflxgrd.col = 3
        
        fr.tag = msflxgrd.Text
        fr.Show 1
        busca
        
    
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical

End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub



Private Sub cmndbtnGrupo_Click()
    Dim g As grupoCli
    g = F_BuscaGrupoCli()
    
    txtbxGrupoCli.Text = g.nome_gru
    txtbxGrupoCli.tag = g.cod_gru
End Sub

Private Sub cmndbtnImpr_Click()
    imprime frmVisualiza.Pic
    
    frmVisualiza.Show 1
       If frmVisualiza.tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub imprime(obj As Object)
    Dim iCount As Integer
    Dim yAux As Single
    Dim yFin As Single
    P_CabRelTimbr "Listagem de Cobrança", obj
    P_ImprX obj, "Conta Sugerida: " & txtbxContaSug.tag & "-" & txtbxContaSug.Text, 5, True
    P_ImprX obj, "Atual: " & txtbxContaCob.tag & "-" & txtbxContaCob.Text, 75, True
    P_ImprX obj, "De: " & txtbxDataIni.Text & " Até:  " & txtbxdataFin.Text, 150, False
    
    
    yAux = obj.CurrentY
                
    P_line1 obj
    P_ImprX obj, "Ord", 2, True
    P_ImprX obj, "Nº Doc", 10, True
    P_ImprX obj, "Id.Parc.", 25, True
    P_ImprX obj, "Id. Doc", 40, True
    P_ImprX obj, "Nome", 60, True
    P_ImprX obj, "Valor", 170, True, direita
    P_ImprX obj, "Restam", 190, False, direita
    P_line1 obj
    obj.DrawWidth = 1
    
    
    
    With msflxgrd
        For iCount = 1 To .rows - 1
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                yFin = obj.CurrentY
                P_Grades obj, "0,9,24,39,59,150,171,191", yAux, yFin, vertical, vbSolid
                
                
                
                
                P_RodapeVis obj, "Continua"
                P_AddPicVis obj
                
                            
                obj.DrawWidth = 2
                
                yAux = obj.CurrentY
                
                P_line1 obj
                P_ImprX obj, "Ord", 2, True
                P_ImprX obj, "Nº Doc", 10, True
                P_ImprX obj, "Id.Parc.", 25, True
                P_ImprX obj, "Id. Doc", 40, True
                P_ImprX obj, "Nome", 60, True
                P_ImprX obj, "Valor", 170, True, direita
                P_ImprX obj, "Restam", 190, False, direita
                P_line1 obj
                obj.DrawWidth = 1
                            
                            
            End If
            
            
            
            .row = iCount
            
            P_ImprX obj, .row, 2, True
            .col = 2
            P_ImprX obj, .Text, 10, True
            .col = 3
            P_ImprX obj, .Text, 25, True
            .col = 4
            P_ImprX obj, .Text, 40, True
            
            
            .col = 10
            P_ImprX obj, .Text, 60, True
            
            .col = 8
            P_ImprX obj, FormatVal(.Text), 170, True, direita
            .col = 7
            
            P_ImprX obj, FormatVal(.Text), 190, False, direita
            
            P_line1 obj
                
            
        Next
        
    End With
    P_AddPicVis obj
    
End Sub
Private Sub cmndbtnImprimir_Click()
    Dim strSQL As String
    strSQL = "Tipo:[" & txtbxTipoCob.Text & "], "
    strSQL = strSQL & "Conta:[" & txtbxContaCob.Text & "], "
    strSQL = strSQL & "Grupo:[" & txtbxGrupoCli.Text & "]," & Chr(9)
    strSQL = strSQL & "Endereço:[" & cmbxEnd.Text & "], " & Chr(9)
    strSQL = strSQL & "Tipo Terc.:[" & cmbxTipoTer.Text & "],"
    strSQL = strSQL & "De:[" & txtbxDataIni.Text & "], "
    strSQL = strSQL & "Até:[" & txtbxdataFin.Text & "]"
    
    P_MsflxToHtml "Listagem de Cobrança", msflxgrd, strSQL
End Sub

Private Sub cmndbtnMarcar_Click()
    Dim strSQL As String
    Dim iCount As Integer
'    If txtbxTransfCob.tag = "" Then
'        MsgBox "Para marcar é necessário informar a conta cobrança!", vbExclamation
'        Exit Sub
'    End If
    If strBaixadoc_usu <> "S" Then
        MsgBox "Você não tem autorização para esse acesso!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Deseja marcar os carnês selecionados desse cobrador?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 1
            If .Text = "S" Then
                .col = 4
                strSQL = "update tab_itdoc set marcado_itd='" & strCod_Usu & F_SoNumero(F_GravaData(Date)) & "'"
                strSQL = strSQL & " where seq_doc='" & .Text & "' "
                Conn.Execute strSQL
            End If
        Next
    End With
    busca
End Sub

Private Sub cmndbtnOrg_Click()
    Dim o As orgao.orgao
    o = F_BuscaOrgao(txtbxOrgao.Text, "")
    
    txtbxOrgao.Text = o.nome_org
    txtbxOrgao.tag = o.cod_org
    
    
    
End Sub


Private Sub cmndbtnTodos_Click()
    If cmndbtnTodos.tag = "" Then
        marcaTodos True
        cmndbtnTodos.tag = "T"
        cmndbtnTodos.caption = "&Nenhum"
    Else
        marcaTodos False
        cmndbtnTodos.tag = ""
        cmndbtnTodos.caption = "&Todos"
    End If
End Sub
Sub marcaTodos(opcao As Boolean)
    Dim iCount As Integer
    
    
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 1
            If opcao = False Then
                msflxgrd.Text = "S"
            Else
                msflxgrd.Text = ""
            End If
            
        Next
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 1
            msflxgrd_DblClick
            
        Next
        
    If opcao = False Then
        lblSelect.caption = "0"
    Else
        lblSelect.caption = msflxgrd.rows - 1
    End If
        
        
    End With
End Sub

Private Sub Form_Load()
    cmbxTipoTer.ListIndex = 0
    cmbxEnd.ListIndex = 0
    resizeForm Me
    txtbxdataFin.Text = Date
    txtbxDataIni.Text = Date
    ajustaGrid
End Sub

Private Sub Form_Resize()
On Error Resume Next
    msflxgrd.width = Me.width - msflxgrd.Left - 200
    msflxgrd.height = pcBot.top - msflxgrd.top - 100
    pcBotoes.Left = pcBot.width - pcBotoes.width - 100
    
End Sub
Sub ajustaGrid()
    With msflxgrd
        .rows = 1
        .Cols = 23
        .FixedCols = 2
        .col = 0
        
        .col = 0: .Text = "..."
        .col = 1: .Text = "S"
        .col = 2: .Text = "Doc"
        .col = 3: .Text = "Id.Parc."
        .col = 4: .Text = "Seq.Doc."
        .col = 5: .Text = "Parcela"
        .col = 6: .Text = "Vencimento"
        .col = 7: .Text = "Valor"
        .col = 8: .Text = "Restam"
        .col = 9: .Text = "Código"
        .col = 10: .Text = "Cliente"
        .col = 11: .Text = "Cidade"
        .col = 12: .Text = "Setor"
        
        .col = 13: .Text = "Endereço"
        .col = 14: .Text = "T"
        .col = 15: .Text = "Fones"
        .col = 16: .Text = "Cob."
        .col = 17: .Text = "Conta Cob. Atual"
        .col = 18: .Text = "Conta Cob Sug."
        .col = 19: .Text = "Pref.Cobr."
        .col = 20: .Text = "Cod.Cob."
        .col = 21: .Text = "Marcado"
        .col = 22: .Text = "docId"
        
        
        .ColWidth(0) = 500
        .ColWidth(1) = 300
        .ColWidth(2) = 700
        .ColWidth(3) = 700
        .ColWidth(5) = 600
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 1000
        .ColWidth(9) = 600
        .ColWidth(10) = 2600
        .ColWidth(11) = 1000
        .ColWidth(12) = 2600
        .ColWidth(13) = 2600
        .ColWidth(14) = 500
        .ColWidth(15) = 2600
        .ColWidth(16) = 300
        .ColWidth(17) = 1500
        .ColWidth(18) = 1500
        .ColWidth(18) = 800
        .ColWidth(18) = 1200
        .ColAlignment(15) = 1
        
    End With
End Sub

Private Sub msflxgrd_DblClick()
    
    With msflxgrd
        .col = 1
        If .Text = "" Then
            
            ZebraLinha msflxgrd, msflxgrd.row
            .col = 1
            .CellBackColor = vbBlue
            .CellForeColor = vbBlue
            .Text = "S"
            lblSelect.caption = Int(lblSelect.caption) + 1
        Else
            .Text = ""
            ZebraLinha msflxgrd, msflxgrd.row, vbWhite
            .CellBackColor = vbWhite
            lblSelect.caption = Int(lblSelect.caption) - 1
        End If
        
    End With
End Sub

Private Sub msflxgrd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'Me.PopupMenu mdiFrm.mnCobranca, , X + msflxgrd.Left, Y + msflxgrd.Top + Me.Top
        
    End If
End Sub

Private Sub msflxgrd_RowColChange()
   ' Dim s As String
    'With msflxgrd
    
   '     .Col = 10
        
    
   ' End With
End Sub

Private Sub optnAlfabet_Click()
    busca
End Sub

Private Sub optnPorDoc_Click()
    busca
End Sub

Private Sub optnPorVenc_Click()
    busca
End Sub

Private Sub pcLista_Click()
    sListaDoc = BuscaListaDoc
    
    
End Sub

Private Sub txtbxContaCob_Change()
    txtbxContaCob.tag = ""
End Sub

Private Sub txtbxContaSug_Change()
    txtbxContaSug.tag = ""
End Sub

Private Sub txtbxContaSug_GotFocus()
    P_SelecTexto txtbxContaSug
End Sub

Private Sub txtbxContaSug_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscaContaSuj_Click
    End If
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
       
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxdataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
       
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxGrupoCli_Change()
    txtbxGrupoCli.tag = ""
End Sub

Private Sub cmndbtnTerc_Click()
    
    
        If cmbxTipoTer.ListIndex = 0 Then
            buscaCli
        Else
            buscaRep
        End If
    
    
    cmndbtnBuscar_Click
End Sub
Sub buscaCli()
    Dim c As cliente
    c = F_BuscaCliente(txtbxTerceiro.Text)
        If c.cod_cli <> "" Then
            txtbxTerceiro.Text = c.nome_cli
            txtbxTerceiro.tag = c.cod_cli
        End If
End Sub
Sub buscaFor()
    Dim f As fornecedor.fornecedor
        f = F_SelectForn(txtbxTerceiro.Text)
        If f.cod_For <> "" Then
            txtbxTerceiro.Text = f.nome_for
            txtbxTerceiro.tag = f.cod_For
        End If
End Sub
Sub buscaRep()
    Dim r As representante.representante
        r = representante.F_BuscaRep(txtbxTerceiro.Text)
        If r.cod_Rep <> "" Then
            txtbxTerceiro.Text = r.nome_rep
            txtbxTerceiro.tag = r.cod_Rep
        End If
    
End Sub

Private Sub txtbxNumDoc_GotFocus()
    P_SelecTexto txtbxNumDoc
End Sub

Private Sub txtbxNumDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxNumDoc_LostFocus()
    
    
    
    P_AjustaTexto txtbxNumDoc
    
    
    
    
End Sub
Sub ListaDocs()
    Dim c As New Collection
    
    Set c = F_SeparaTexto(txtbxNumDoc, ",; ", False, "000000#", "'")
    sListaDoc = CollectionToNotIn(c)
    txtbxNumDoc.Text = sListaDoc
    
    
    
    
    
End Sub
Private Sub txtbxOrgao_GotFocus()
    P_SelecTexto txtbxOrgao
End Sub

Private Sub txtbxOrgao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnOrg_Click
    End If
End Sub

Private Sub txtbxSeqDoc_GotFocus()
    P_SelecTexto txtbxSeqDoc
End Sub

Private Sub txtbxSeqDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
        busca
    End If
End Sub

Private Sub txtbxSeqDoc_LostFocus()
    P_AjustaNota txtbxSeqDoc, 7
End Sub

Private Sub txtbxSeqItd_GotFocus()
    P_SelecTexto txtbxSeqItd
End Sub

Private Sub txtbxSeqItd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
        busca
    End If
End Sub

Private Sub txtbxSeqItd_LostFocus()
    P_AjustaNota txtbxSeqItd, 7
End Sub

Private Sub txtbxTerceiro_Change()
    txtbxTerceiro.tag = ""
End Sub

Private Sub txtbxTerceiro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxTipoCob_Change()
    txtbxTipoCob.tag = ""
End Sub

Private Sub txtbxTipoCob_LostFocus()
    If txtbxTipoCob.Text = "" Then txtbxTipoCob.tag = ""
End Sub

Private Sub txtCidade_Change()
    txtCidade.tag = ""
End Sub

Private Sub txtCidade_GotFocus()
    P_SelecTexto txtCidade
End Sub

Private Sub txtCidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdCidade_Click
    End If
End Sub

Private Sub txtCidade_LostFocus()
    P_AjustaTexto txtCidade
End Sub

Private Sub txtPredata_LostFocus()
    P_AjustaData txtPredata
End Sub
