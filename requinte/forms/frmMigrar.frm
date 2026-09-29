VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMigrar 
   Caption         =   "Migrar"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnImport 
      Caption         =   "Importar Arquivos"
      Height          =   495
      Left            =   1560
      TabIndex        =   9
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   495
      Left            =   7560
      TabIndex        =   7
      Top             =   2280
      Width           =   1455
   End
   Begin TabDlg.SSTab sstb 
      Height          =   2145
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   3784
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Exportar"
      TabPicture(0)   =   "frmMigrar.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmndbtnGerar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Importar"
      TabPicture(1)   =   "frmMigrar.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmndbtnBuscaArq"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton cmndbtnBuscaArq 
         Caption         =   "Buscar Arquivo"
         Height          =   495
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnGerar 
         Caption         =   "Gerar"
         Height          =   495
         Left            =   4920
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Opções:"
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   4695
         Begin VB.OptionButton optnProd 
            Caption         =   "Produtos"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optnCli 
            Caption         =   "Clientes"
            Height          =   195
            Left            =   1440
            TabIndex        =   4
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optnForn 
            Caption         =   "Fornecedores"
            Height          =   195
            Left            =   2880
            TabIndex        =   3
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   8400
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmndbtnInfo 
      Caption         =   "Informações"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   1455
   End
End
Attribute VB_Name = "frmMigrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnDestino_Click()
    Dim sPath As String
    frmEscolhePasta.Show 1
    If frmEscolhePasta.tag <> "" Then
        txtbxDest.Text = frmEscolhePasta.Dir1.Path
        
    End If
    Unload frmEscolhePasta
    
End Sub

Private Sub cmndbtnGerar_Click()
    If optnProd.Value = True Then
        P_GeraArquivoTransfProdutos ""
    ElseIf optnCli.Value = True Then
   
    Else
    End If
End Sub

Private Sub cmndbtnImport_Click()
    frmImportArquivo.Show 1
End Sub

Private Sub cmndbtnInfo_Click()
    frmInfo.Show 1
End Sub

Private Sub cmndbtnMigrar_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
        
    
    
    Banco.Execute "drop table if exists " & lblEntDest.caption
    FileCopy txtbxOrigem.Text & "\" & lblEntDest.caption & ".frm", "c:\mysql\data\" & sNomeBanco & "\" & lblEntDest.caption & ".frm"
    FileCopy txtbxOrigem.Text & "\" & lblEntDest.caption & ".myi", "c:\mysql\data\" & sNomeBanco & "\" & lblEntDest.caption & ".MYI"
    FileCopy txtbxOrigem.Text & "\" & lblEntDest.caption & ".myd", "c:\mysql\data\" & sNomeBanco & "\" & lblEntDest.caption & ".MYD"
        

    
    strSQL = "insert into " & lblEntOrig & "  select * from " & lblEntDest
    Banco.Execute strSQL
    
    
    MsgBox "Migrado com sucesso!", vbInformation
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
    
    
    
End Sub

Private Sub Command1_Click()
        Dim sPath As String
    frmEscolhePasta.Show 1
    If frmEscolhePasta.tag <> "" Then
        txtbxOrigem.Text = frmEscolhePasta.Dir1.Path
        
    End If
    Unload frmEscolhePasta

End Sub

Private Sub Command2_Click()
    sNomeBanco = txtbxBanco.Text
    If Not F_AbreBanco() Then
        Exit Sub
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    optnFotos_Click
End Sub

Private Sub optnAnot_Click()
    
End Sub

Private Sub optnFotos_Click()
    
End Sub

Private Sub optnPol_Click()
    
End Sub

Private Sub txtbxCodBan_Change()

End Sub
