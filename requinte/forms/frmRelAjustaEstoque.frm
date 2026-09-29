VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRelAjustaEstoque 
   Caption         =   "Ajuste de Estoque"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   14550
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctop 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   14490
      TabIndex        =   17
      Top             =   8895
      Width           =   14550
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2655
      Left            =   0
      TabIndex        =   15
      Top             =   960
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   4683
      _Version        =   393216
   End
   Begin VB.PictureBox pcBusca 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   14520
      TabIndex        =   0
      Top             =   0
      Width           =   14550
      Begin VB.CommandButton cmdPreencher 
         Caption         =   "Preencher"
         Height          =   495
         Left            =   9360
         TabIndex        =   16
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdListar 
         Caption         =   "Listar"
         Height          =   495
         Left            =   8040
         TabIndex        =   14
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
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
         MaxLength       =   40
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtbxSecao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         MaxLength       =   40
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtbxSubSec 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         MaxLength       =   40
         TabIndex        =   5
         Top             =   360
         Width           =   2295
      End
      Begin VB.PictureBox pcEstCons 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   14520
         ScaleHeight     =   855
         ScaleWidth      =   1575
         TabIndex        =   1
         Top             =   0
         Width           =   1575
         Begin VB.CommandButton cmndbtnEstCons 
            Height          =   255
            Left            =   1200
            Picture         =   "frmRelAjustaEstoque.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   120
            Width           =   255
         End
         Begin VB.Label lblEstCons 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   600
            TabIndex        =   4
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Est. Cons:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   3
            Top             =   120
            Width           =   930
         End
      End
      Begin VB.CommandButton cmndbtnSubSec 
         Height          =   285
         Left            =   7200
         Picture         =   "frmRelAjustaEstoque.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnSecao 
         Height          =   285
         Left            =   4560
         Picture         =   "frmRelAjustaEstoque.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblSubCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Subseção:"
         Height          =   195
         Left            =   4920
         TabIndex        =   13
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lblSecCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Seção:"
         Height          =   195
         Left            =   2400
         TabIndex        =   12
         Top             =   120
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   600
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   14760
         TabIndex        =   10
         Tag             =   "!"
         Top             =   0
         Width           =   195
      End
   End
End
Attribute VB_Name = "frmRelAjustaEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ado As New ADODB.Recordset

Private Sub cmdListar_Click()
    busca
End Sub

Private Sub cmndbtnSecao_Click()
    Dim s As secao.secao
    s = F_BuscaSec(txtbxSecao.Text)
    If s.cod_sec <> "" Then
        txtbxSecao.Text = s.Nome_Sec
        txtbxSecao.tag = s.cod_sec
        txtbxSubSec.Text = ""
        txtbxSubSec.tag = ""
        busca
        txtbxSubSec.SetFocus
    End If
    
    
    
End Sub

Private Sub cmndbtnSubSec_Click()
    Dim s As subsec
    s = F_BuscaSubSec(txtbxSecao.tag, txtbxSubSec.Text)
    If s.cod_sub <> "" Then
        txtbxSubSec.Text = s.nome_sub
        txtbxSubSec.tag = s.cod_sub
        busca
        txtbxFab.SetFocus
    End If
End Sub
Sub busca()
    Dim strSQL As String
    grid.rows = 1
    Dim rPesq As New ADODB.Recordset
    
    
    strSQL = "select tp.nome_pro,tp.er_pro,tp.cod_pro,tp.er_pro,tp.ef_pro from tab_produtos as tp "
    strSQL = strSQL & " left join tab_subsec as tu on tu.cod_sub=tp.cod_sub "
    strSQL = strSQL & " left join tab_secoes as ts on ts.cod_sec=tu.cod_sec "
    strSQL = strSQL & " where 1>0 "
    If txtbxSecao.tag <> "" Then
        strSQL = strSQL & " and ts.cod_sec='" & txtbxSecao.tag & "' "
    End If
    If txtbxSubSec.tag <> "" Then
        strSQL = strSQL & " and tp.cod_sub='" & txtbxSubSec.tag & "' "
    End If
    
    strSQL = strSQL & " and tp.nome_pro like '" & txtbxExpr.Text & "%'"
    
    strSQL = strSQL & " order by nome_pro,cod_pro  "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "" & Chr(9) & !cod_pro & Chr(9) & !nome_pro & Chr(9) & !er_pro
            
            grid.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
    
    
End Sub
Private Sub Form_Load()
    ajustaGrid
End Sub
Sub ajustaGrid()
    With grid
        
        .FixedRows = 1
        .rows = 1
        .FixedCols = 0
        .Cols = 11
        .Col = 0: .Text = "*"
        .Col = 1: .Text = "Código"
        .Col = 2: .Text = "Nome"
        .Col = 3: .Text = "Estoque"
        .Col = 4: .Text = "Reservado"
        .Col = 5: .Text = "Contagem"
        .Col = 6: .Text = "Qtd Contagem"
        .Col = 7: .Text = "Vendido"
        .Col = 8: .Text = "Comprado"
        .Col = 9: .Text = "Acerto"
        .Col = 10: .Text = "Sugestão"
        .ColWidth(2) = 3000
        
    End With
End Sub

Private Sub Form_Resize()
    PosicionaGrade grid, Me, pctop
End Sub
