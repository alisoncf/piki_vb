VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAlterarPrecosEmLote 
   Caption         =   "Alterar Preços em Lote"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   11775
      TabIndex        =   6
      Top             =   6015
      Width           =   11775
      Begin VB.CommandButton cmndbtnNao 
         Caption         =   "Sair"
         Height          =   615
         Left            =   10320
         Picture         =   "frmAlterarPrecosEmLote.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnSim 
         Caption         =   "Alterar preços"
         Height          =   615
         Left            =   9000
         Picture         =   "frmAlterarPrecosEmLote.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgrid 
      Height          =   4335
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   7646
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   11745
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.CheckBox chkCab 
         Caption         =   "A primeira linha contém cabeçalho"
         Height          =   195
         Left            =   2160
         TabIndex        =   5
         Top             =   840
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CommandButton cmdArquivo 
         Caption         =   "Buscar CSV"
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblArquivo 
         AutoSize        =   -1  'True
         Caption         =   "arquivo"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Arquivo com código do produto na primeira coluna e novo preço na segunda separados por pontoe vírgula e quebra de linha com enter."
         Height          =   615
         Left            =   2160
         TabIndex        =   2
         Top             =   120
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmAlterarPrecosEmLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdArquivo_Click()
    mdiFrm.cmndlg.FileName = "*.csv"
    mdiFrm.cmndlg.ShowOpen
    lblArquivo.caption = mdiFrm.cmndlg.FileName
    Carrega
End Sub
Sub Carrega()
    On Error GoTo trata_Erro
    Dim linha As String
    Dim c As New Collection
    Dim i As Integer
    Dim p As Produto
    hgrid.rows = 1
    AbreArquivoExternoLeitura lblArquivo.caption, 1
    hgrid.Visible = False
    Do While Not EOF(1)
         Line Input #1, linha
         Set c = F_SeparaTexto(linha, ";", False)
         If (i > 0) Or (i = 0 And chkCab.Value = 0) Then
            hgrid.AddItem c(1) & Chr(9) & c(2)
            
            p = Produtos.buscaProduto(c(1))
            hgrid.row = hgrid.rows - 1
            hgrid.col = 2: hgrid.Text = p.codfab_pro
            hgrid.col = 3: hgrid.Text = p.nome_pro
            hgrid.col = 4: hgrid.Text = FormatVal(p.pv_pro)
            hgrid.col = 0: hgrid.Text = p.cod_pro
            If p.pv_pro <> "" Then
                If CCur(p.pv_pro) > c(2) Then
                    ColoreLinhaGrid hgrid, hgrid.rows - 1, vbMagenta
                ElseIf CCur(p.pv_pro) < c(2) Then
                    ColoreLinhaGrid hgrid, hgrid.rows - 1, vbYellow
                End If
            End If
            If p.cod_pro = "" Then
                ColoreLinhaGrid hgrid, hgrid.rows - 1
            End If
         End If
         i = 1 + 1
    Loop
    hgrid.Visible = True
    FechaArquivoExterno 1
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub ajustaGrid()
    Dim o As Object
    AjustaGridData hgrid, 5, 1, 2, "Código,Preço,CodFab,Produto,Preço_Atual", "", "800,1200,2000,3500,1200", o, "", True
     
End Sub

Private Sub cmndbtnSim_Click()
    If Not F_senha Then Exit Sub
    Dim i As Integer
    Dim cod As String
    Dim preco As String
    Dim strSQL As String
    Dim precoAntigo As String
    If strAltpreco_usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Confirma alteração de preços em lote?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    With hgrid
        If .rows <= 1 Then
            Exit Sub
        End If
        For i = 1 To .rows - 1
            .row = i
            .col = 0
            cod = .Text
            .col = 1
            preco = .Text
            .col = 4
            precoAntigo = .Text
            Produtos.P_AltPreco cod, CCur(preco), CCur(precoAntigo), True
            
        Next
    End With
    MsgBox "Processamento concluído!", vbInformation
    
    
End Sub

Private Sub Form_Activate()
    Form_Resize
End Sub

Private Sub Form_Load()
    ajustaGrid
End Sub

Private Sub Form_Resize()
    PosicionaGrade hgrid, pcBot, Me, 500, True
    hgrid.height = pcBot.Top - hgrid.Top - 200
End Sub

