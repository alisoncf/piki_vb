VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTrocaAvulsa 
   Caption         =   "Troca Avulsa"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13500
   Icon            =   "frmTrocaAvulsa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   13470
      TabIndex        =   55
      Top             =   0
      Width           =   13500
      Begin VB.TextBox txtTelefone 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3960
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtObs 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7800
         TabIndex        =   8
         Top             =   240
         Width           =   5175
      End
      Begin VB.TextBox txtVendedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5520
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3735
      End
      Begin VB.CommandButton cmdVendedor 
         Height          =   315
         Left            =   7320
         Picture         =   "frmTrocaAvulsa.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone"
         Height          =   195
         Left            =   3960
         TabIndex        =   2
         Top             =   0
         Width           =   630
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observações:"
         Height          =   195
         Left            =   7800
         TabIndex        =   7
         Top             =   0
         Width           =   990
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5520
         TabIndex        =   4
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   525
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1185
      ScaleWidth      =   13470
      TabIndex        =   37
      Top             =   6825
      Width           =   13500
      Begin VB.PictureBox pcDesconto 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   2520
         ScaleHeight     =   1065
         ScaleWidth      =   3825
         TabIndex        =   44
         Top             =   0
         Width           =   3855
         Begin VB.TextBox txtbxDesconto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   46
            Text            =   "0,00"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtbxDescPerc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   48
            Text            =   "0"
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton cmndbtnDesconto 
            Caption         =   "Desconto"
            Height          =   525
            Left            =   1920
            Picture         =   "frmTrocaAvulsa.frx":1054
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   240
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Desconto Gerencial"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   90
            TabIndex        =   52
            Top             =   -30
            Width           =   1860
         End
         Begin VB.Label lblAutorizaDesconto 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   345
            Left            =   1920
            TabIndex        =   51
            Top             =   720
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Autorizado Por:"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   840
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desconto:"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   195
            Left            =   1200
            TabIndex        =   47
            Top             =   240
            Width           =   120
         End
      End
      Begin VB.CommandButton cmndbtnCancel 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "Ca&ncelar (F12)"
         Height          =   735
         Left            =   11640
         Picture         =   "frmTrocaAvulsa.frx":17BE
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnOk 
         Appearance      =   0  'Flat
         Caption         =   "C&onfirmar (F7)"
         Height          =   735
         Left            =   9960
         Picture         =   "frmTrocaAvulsa.frx":1D0C
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblDiferenca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1560
         TabIndex        =   59
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diferença:"
         Height          =   195
         Left            =   120
         TabIndex        =   58
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Final Previsto:"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   840
         Width           =   990
      End
      Begin VB.Label lblFinalPrev 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   1560
         TabIndex        =   53
         Top             =   840
         Width           =   360
      End
      Begin VB.Label lblTotalSaida 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1560
         TabIndex        =   43
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Saida:"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblTotalEntr 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1560
         TabIndex        =   41
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Entrada:"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   120
         Width           =   1005
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   0
      ScaleHeight     =   2745
      ScaleWidth      =   13425
      TabIndex        =   34
      Top             =   3480
      Width           =   13455
      Begin VB.CommandButton cmdAddOut 
         Caption         =   "Adicionar"
         Height          =   375
         Left            =   10200
         TabIndex        =   29
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdDelOut 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   11640
         TabIndex        =   30
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtTotalSai 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   8880
         TabIndex        =   28
         Text            =   "0,00"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtUnitSai 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   7800
         TabIndex        =   26
         Text            =   "0,00"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtProdSai 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   6015
      End
      Begin VB.TextBox txtQtdSai 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   6960
         TabIndex        =   24
         Text            =   "1"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdProdSai 
         Height          =   315
         Left            =   6120
         Picture         =   "frmTrocaAvulsa.frx":1E56
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   480
         Width           =   255
      End
      Begin MSFlexGridLib.MSFlexGrid msfSai 
         Height          =   1455
         Left            =   120
         TabIndex        =   57
         Top             =   1200
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   2566
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8880
         TabIndex        =   27
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unitário"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7800
         TabIndex        =   25
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Qtd:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6960
         TabIndex        =   23
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Produto"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produtos que Saem:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   36
         Top             =   0
         Width           =   2355
      End
      Begin VB.Label lblQueimaSai 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Queima"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   9240
         TabIndex        =   35
         Top             =   120
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2505
      ScaleWidth      =   13425
      TabIndex        =   31
      Top             =   840
      Width           =   13455
      Begin VB.CommandButton cmdExclIn 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   11640
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddIn 
         Caption         =   "Adicionar"
         Height          =   375
         Left            =   10200
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   6960
         TabIndex        =   13
         Text            =   "1"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   6015
      End
      Begin VB.TextBox txtbxUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   7800
         TabIndex        =   15
         Text            =   "0,00"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtbxTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   8880
         TabIndex        =   17
         Text            =   "0,00"
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnLocaliza 
         Height          =   315
         Left            =   6120
         Picture         =   "frmTrocaAvulsa.frx":21E0
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   480
         Width           =   255
      End
      Begin MSFlexGridLib.MSFlexGrid msfEntra 
         Height          =   1575
         Left            =   120
         TabIndex        =   56
         Top             =   840
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   2778
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblQueima 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Queima"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   9240
         TabIndex        =   33
         Top             =   120
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produtos que Entram:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   32
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Produto"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Qtd:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6960
         TabIndex        =   12
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unitário"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7800
         TabIndex        =   14
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8880
         TabIndex        =   16
         Top             =   240
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmTrocaAvulsa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddIn_Click()
    Dim strSQL As String
    Dim c As Currency
    Dim q As Currency
    If F_InFlxgrd(msfEntra, txtProduto.tag, 0) > -1 Then
        If MsgBox("Este produto já está lançado. Deseja alterar quantidade?", vbYesNo + vbQuestion) = vbNo Then
        
            Exit Sub
        Else
            msfEntra.col = 2
            msfEntra.Text = CCur(msfEntra.Text) + CCur(txtbxQtd.Text)
            q = msfEntra.Text
            msfEntra.col = 3
            c = CCur(msfEntra.Text)
            msfEntra.col = 4
            msfEntra.Text = FormatVal(q * c)
            soma
            
            
            
            Exit Sub
        End If
    End If
    strSQL = txtProduto.tag & Chr(9) & txtProduto.Text & Chr(9)
    strSQL = strSQL & txtbxQtd.Text & Chr(9)
    strSQL = strSQL & txtbxUnit.Text & Chr(9)
    strSQL = strSQL & txtbxTotal.Text & Chr(9)
    If lblQueima.Visible = True Then
        strSQL = strSQL & "Sim"
    End If
    With msfEntra
        .AddItem strSQL
    End With
    soma
    txtProduto.SetFocus
End Sub
Sub AjustaGrid(grid As MSFlexGrid)
    On Error Resume Next
    With grid
        .rows = 1
        .FixedCols = 0
        .Cols = 6
        .col = 0: .Text = "Código"
        .col = 1: .Text = "Descrição"
        .col = 2: .Text = "Qtd"
        .col = 3: .Text = "Unitário"
        .col = 4: .Text = "Total"
        .col = 5: .Text = "Queima"
        .ColWidth(1) = 3000
    End With
End Sub

Private Sub cmdAddOut_Click()
    Dim strSQL As String
    If F_InFlxgrd(msfSai, txtProdSai.tag, 0) > -1 Then
        MsgBox "Este produto já está lançado.", vbExclamation
        Exit Sub
    End If
    strSQL = txtProdSai.tag & Chr(9) & txtProdSai.Text & Chr(9)
    strSQL = strSQL & txtQtdSai.Text & Chr(9)
    strSQL = strSQL & txtUnitSai.Text & Chr(9)
    strSQL = strSQL & txtTotalSai.Text & Chr(9)
    If lblQueimaSai.Visible = True Then
        strSQL = strSQL & "Sim"
    End If
    With msfSai
        .AddItem strSQL
    End With
    soma
End Sub

Private Sub cmdDelOut_Click()
    If msfSai.rows = 2 Then
        msfSai.rows = 1
    Else
        msfSai.RemoveItem msfSai.row
    End If
    soma
End Sub

Private Sub cmdExclIn_Click()
    If msfEntra.rows = 2 Then
        msfEntra.rows = 1
    Else
        msfEntra.RemoveItem msfEntra.row
    End If
    soma
End Sub

Private Sub cmdProdSai_Click()
    buscaProdutoSai
End Sub

Private Sub cmdVendedor_Click()
    Dim f As funcionario.funcionario
    f = funcionario.F_SelectVendedor(txtVendedor.Text)
    If f.cod_fun <> "" Then
        txtVendedor.Text = f.guerra_fun
        txtVendedor.tag = f.cod_fun
        txtObs.SetFocus
    End If
    
End Sub

Private Sub cmndbtnCancel_Click()
    Unload Me
End Sub

Private Sub cmndbtnDesconto_Click()
    Dim sUsu As String
    Dim sDes As String
    If txtbxDesconto.Text = 0 Then
        Exit Sub
    End If
    sUsu = F_SenhaCodBar
    
    If lblAutorizaDesconto.tag <> "" Then
        MsgBox "Já foi autorizado um desconto para esta venda!", vbExclamation
        Exit Sub
    End If
    
    If sUsu = "" Then
        MsgBox "Não autenticado!", vbExclamation
        Exit Sub
    End If
    
    sDes = F_BuscaUmCampo("desconto_usu", "tab_defusu", "cod_usu='" & sUsu & "'")
    If sDes = "" Then
        sDes = "0"
    End If
    If CCur(sDes) < CCur(txtbxDescPerc.Text) Then
        MsgBox "O Desconto possível para este usuário é de " & sDes & "%", vbExclamation
        txtbxDesconto.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Confirma desconto de " & txtbxDesconto.Text & "?", vbYesNo + vbCritical) = vbYes Then
        lblFinalPrev.caption = FormatVal(CCur(lblTotalSaida.caption) - CCur(lblTotalEntr.caption) - CCur(txtbxDesconto.Text))
        lblDiferenca.caption = lblFinalPrev.caption
    End If
    
    If strExclVenda_usu <> "S" Then
        MsgBox "Para fazer este tipo de operação, o usuário deve ter permissão para excluir vendas!", vbExclamation
        Exit Sub
    End If
    If Not F_senha Then
        Exit Sub
    End If
    If MsgBox("Confirma a Troca avulsa dos produtos estipulados àcima?", vbYesNo + vbQuestion) = vbYes Then
               
    End If
    
End Sub

Private Sub cmndbtnLocaliza_Click()
    buscaProduto
End Sub
Sub buscaProduto()
    Dim p As Produto
    p = F_BuscaProdSaida(txtProduto.Text)
    lblQueima.Visible = False
    If p.cod_pro <> "" Then
        txtProduto.Text = p.nome_pro
        txtProduto.tag = p.cod_pro
        txtbxUnit.Text = FormatVal(p.pv_pro)
        txtbxTotal.Text = FormatVal(p.pv_pro)
        lblQueima.Visible = False
        If p.pq_pro > 0 Then
            If MsgBox("O produto está marcado como 'queima', deseja lançar o preço promocional?", vbYesNo + vbQuestion) = vbYes Then
            
                txtbxUnit.Text = FormatVal(p.pq_pro)
                txtbxTotal.Text = FormatVal(p.pq_pro)
            End If
            
            
            lblQueima.Visible = True
            soma
        Else
            soma
        End If
        
        
        
        txtbxQtd.SetFocus
        
        
        
        
    End If
    
End Sub
Sub somaEntra()
    Dim iCount As Integer
    Dim cTotal As Currency
    With msfEntra
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 4
            cTotal = cTotal + CCur(.Text)
        Next
    End With
    lblTotalEntr.caption = FormatVal(cTotal)
End Sub
Sub somaSai()
    Dim iCount As Integer
    Dim cTotal As Currency
    With msfSai
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 4
            cTotal = cTotal + CCur(.Text)
        Next
    End With
    lblTotalSaida.caption = FormatVal(cTotal)
End Sub
Sub soma()
    somaEntra
    somaSai
    lblFinalPrev.caption = FormatVal(CCur(lblTotalSaida.caption) - CCur(lblTotalEntr.caption))
    lblDiferenca.caption = lblFinalPrev.caption
End Sub
Sub buscaProdutoSai()
    Dim p As Produto
    p = F_BuscaProdSaida(txtProdSai.Text)
    If p.cod_pro <> "" Then
        txtProdSai.Text = p.nome_pro
        txtProdSai.tag = p.cod_pro
        txtUnitSai.Text = FormatVal(p.pv_pro)
        txtTotalSai.Text = FormatVal(p.pv_pro)
        
        lblQueimaSai.Visible = False
        If p.pq_pro > 0 Then
            If p.pq_pro > 0 Then
            If MsgBox("O produto está marcado como 'queima', deseja lançar o preço promocional?", vbYesNo + vbQuestion) = vbYes Then
            
                txtUnitSai.Text = FormatVal(p.pq_pro)
                txtTotalSai.Text = FormatVal(p.pq_pro)
            End If
            
            
            lblQueima.Visible = True
            soma
        Else
            soma
        End If
 
        End If
        
        
        
        txtQtdSai.SetFocus
        
        
        
        
    End If
    
End Sub


Private Sub cmndbtnLocaliza_GotFocus()
    P_SelecTexto txtProduto
End Sub

Private Sub cmndbtnOk_Click()
    Dim strSQL As String
    Dim sTipo As String
    Dim iCount As Integer
    Dim sSeq As String
    Dim sCod As String
    Dim sUnit As String
    Dim sTotal As String
    Dim sQtd As String
    If msfSai.rows <= 1 Then
        If MsgBox("Deseja continuar sem produtos que saem!", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If
    If msfEntra.rows <= 1 Then
        MsgBox "Informe o produto que entra!", vbExclamation
        Exit Sub
    End If
    If txtCliente.Text = "" Then
        MsgBox "Informe o nome do cliente!", vbExclamation
        Exit Sub
    End If
    If txtTelefone.Text = "" Then
        MsgBox "Informe o telefone nome do cliente!", vbExclamation
        Exit Sub
    End If
    If strEfetivarVenda_Usu <> "S" Then
        MsgBox "Para esta operação você precisa ter permissão para excluir vendas!", vbExclamation
        Exit Sub
    End If
    If txtVendedor.tag = "" Then
        MsgBox "Informe um vendedor", vbInformation
        txtVendedor.SetFocus
        Exit Sub
    End If
    If Not F_senha Then
        Exit Sub
    End If
    
    
    If MsgBox("Confirma conclusão da troca avulsa no valor de " & lblFinalPrev.caption & "?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    sSeq = GravaTrocaAvulsa(txtCliente.Text, txtVendedor.tag, txtObs.Text, lblTotalEntr.caption, lblTotalSaida.caption, _
        txtbxDesconto.Text, txtTelefone.Text)
    Me.tag = sSeq
    If sSeq <> "" Then
        If CCur(lblFinalPrev.caption) < 0 Then
            sTipo = "D"
        Else
            sTipo = "C"
        End If
        
        With msfEntra
            For iCount = 1 To .rows - 1
                .row = iCount
                .col = 0: sCod = .Text
                .col = 2: sQtd = .Text
                .col = 3: sUnit = .Text
                .col = 4: sTotal = .Text
                GravaTrocaAvulsaIt "I", sSeq, sCod, sQtd, sUnit, sTotal
                If Produtos.F_AlteraEstoquePro(txtProduto.tag, txtbxQtd.Text) Then
            
                End If
            Next
        End With
        With msfSai
            For iCount = 1 To .rows - 1
                .row = iCount
                .col = 0: sCod = .Text
                .col = 2: sQtd = .Text
                .col = 3: sUnit = .Text
                .col = 4: sTotal = .Text
                GravaTrocaAvulsaIt "O", sSeq, sCod, sQtd, sUnit, sTotal
                If Produtos.F_AlteraEstoquePro(txtProdSai.tag, "-" & txtQtdSai.Text) Then
                End If
            Next
        End With
        
        
        
        
        
        
        If F_GravaCai(sTipo, Date, Abs(lblFinalPrev.caption), "TROCA AVULSA " & Now, "", "DIN", strCod_Cai, _
            0, 0, strCod_Cai, strCod_Tur, "", "N") Then
        
        End If
            
        imprTrocaAvulsa Me.tag
        
        MsgBox "Produtos alterados com sucesso!", vbExclamation
        Unload Me
        Exit Sub
        
    End If
    
    
End Sub



Private Sub Form_Load()
    AjustaGrid msfEntra
    AjustaGrid msfSai
    
End Sub

Private Sub txtbxDesconto_GotFocus()
    P_SelecTexto txtbxDesconto
End Sub

Private Sub txtbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDesconto_LostFocus()
    On Error Resume Next
    P_ValorLostFocus txtbxDesconto
   txtbxDescPerc.Text = FormatVal(CCur(txtbxDesconto.Text) * 100 / CCur(lblDiferenca.caption), "###,##0.0000")

    lblFinalPrev.caption = FormatVal(CCur(lblDiferenca.caption) - CCur(txtbxDesconto.Text))
    
End Sub

Private Sub txtbxDescPerc_GotFocus()
    P_SelecTexto txtbxDescPerc
End Sub

Private Sub txtbxDescPerc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub




Private Sub txtbxDescPerc_LostFocus()
    P_ValorLostFocus txtbxDescPerc, "###,##0.0000"
    On Error Resume Next
    txtbxDesconto.Text = FormatVal(CCur(txtbxDescPerc.Text) / 100 * (CCur(lblValor.caption) - CCur(lblDiferenca.caption)))
    lblFinalPrev.caption = FormatVal(CCur(lblValor.caption) - CCur(txtbxDesconto.Text))
    
    
End Sub



Private Sub txtProduto_GotFocus()
    P_SelecTexto txtProduto
End Sub

Private Sub txtProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnLocaliza_Click
    End If
End Sub

Private Sub txtbxQtd_GotFocus()
    P_SelecTexto txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxQtd_LostFocus()
    P_ValorLostFocus txtbxQtd
    totalEntrada
End Sub
Sub totalEntrada()
    txtbxTotal.Text = FormatVal(CCur(txtbxQtd.Text) * CCur(txtbxUnit.Text))
    
End Sub
Sub totalSaida()
    txtTotalSai.Text = FormatVal(CCur(txtQtdSai.Text) * CCur(txtUnitSai.Text))
    
End Sub
Private Sub txtProdSai_GotFocus()
    P_SelecTexto txtProdSai
End Sub

Private Sub txtProdSai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaProdutoSai
    End If
End Sub

Private Sub txtQtdSai_GotFocus()
    P_SelecTexto txtQtdSai
End Sub

Private Sub txtQtdSai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtQtdSai_LostFocus()
    P_ValorLostFocus txtQtdSai
    totalSaida
End Sub

Private Sub txtTelefone_GotFocus()
    P_SelecTexto txtTelefone
End Sub

Private Sub txtTelefone_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtTelefone_LostFocus()
    P_AjustaTexto txtTelefone
End Sub

Private Sub txtVendedor_Change()
    txtVendedor.tag = ""
End Sub

Private Sub txtVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdVendedor_Click
    End If
End Sub
