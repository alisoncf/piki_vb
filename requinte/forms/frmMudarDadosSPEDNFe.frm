VERSION 5.00
Begin VB.Form frmMudarDadosSPEDNFe 
   Caption         =   "Mudar Dados da NFE para os produtos selecionados"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7635
   Icon            =   "frmMudarDadosSPEDNFe.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   6000
      Picture         =   "frmMudarDadosSPEDNFe.frx":1084A
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Confirmar"
      Height          =   735
      Left            =   4440
      Picture         =   "frmMudarDadosSPEDNFe.frx":109B5
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5280
      Width           =   1455
   End
   Begin VB.PictureBox pcCST 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   7305
      TabIndex        =   0
      Top             =   240
      Width           =   7335
      Begin VB.TextBox txtCst_Icms 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   6735
      End
      Begin VB.PictureBox pcCstIcms 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6840
         Picture         =   "frmMudarDadosSPEDNFe.frx":10E5D
         ScaleHeight     =   255
         ScaleWidth      =   225
         TabIndex        =   4
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtCst_Ipi 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   3000
         Width           =   6735
      End
      Begin VB.PictureBox pcCstIpi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6840
         Picture         =   "frmMudarDadosSPEDNFe.frx":11208
         ScaleHeight     =   255
         ScaleWidth      =   225
         TabIndex        =   16
         Top             =   3000
         Width           =   255
      End
      Begin VB.TextBox txtCst_Pis 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   6735
      End
      Begin VB.PictureBox pcCstPis 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6840
         Picture         =   "frmMudarDadosSPEDNFe.frx":115B3
         ScaleHeight     =   255
         ScaleWidth      =   225
         TabIndex        =   10
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox txtCst_Cofins 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   6735
      End
      Begin VB.PictureBox pcCstCofins 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6840
         Picture         =   "frmMudarDadosSPEDNFe.frx":1195E
         ScaleHeight     =   255
         ScaleWidth      =   225
         TabIndex        =   13
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox txtbxICMS 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   2775
      End
      Begin VB.PictureBox pcIcms 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2880
         Picture         =   "frmMudarDadosSPEDNFe.frx":11D09
         ScaleHeight     =   255
         ScaleWidth      =   225
         TabIndex        =   7
         Top             =   1200
         Width           =   255
      End
      Begin VB.PictureBox pcGenero 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6840
         Picture         =   "frmMudarDadosSPEDNFe.frx":120B4
         ScaleHeight     =   255
         ScaleWidth      =   225
         TabIndex        =   19
         Top             =   3600
         Width           =   255
      End
      Begin VB.TextBox txtGenero 
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
         TabIndex        =   18
         Top             =   3600
         Width           =   6735
      End
      Begin VB.TextBox txtOrigem 
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
         TabIndex        =   21
         Top             =   4200
         Width           =   6735
      End
      Begin VB.PictureBox pcOrigem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6840
         Picture         =   "frmMudarDadosSPEDNFe.frx":1245F
         ScaleHeight     =   255
         ScaleWidth      =   225
         TabIndex        =   22
         Top             =   4200
         Width           =   255
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cst ICMS:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cst IPI:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   510
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cst PIS:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cst COFINS:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   900
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Códigos de Situação tributária em operações de saída (vendas):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   5265
      End
      Begin VB.Label Label32 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "&ICMS:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gênero: (Apenas para produtos primários ex.: insumos)"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   3360
         Width           =   3840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Origem:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   3960
         Width           =   540
      End
      Begin VB.Label lblIcms 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Situação Trib."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5760
         TabIndex        =   23
         Top             =   1200
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmMudarDadosSPEDNFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnSair_Click()
    Me.tag = ""
    Me.Hide
End Sub

Private Sub cmndbtnSalvar_Click()
    Me.tag = "x"
    Me.Hide
End Sub



Private Sub pcCstCofins_Click()
    Dim p As Padrao1
    Dim expr As String
    
    expr = txtCst_Cofins.Text
    If F_SoNumero(expr) <> "" Then
        expr = F_SoNumero(expr)
           
    End If
    
    p = Sped_Lista_Cst_Cofins(expr)
    If p.codigo <> "" Then
        txtCst_Cofins.Text = p.codigo & " - " & p.nome
        txtCst_Cofins.tag = p.codigo
        
    End If
End Sub

Private Sub pcCstIcms_Click()
    Dim p As Padrao1
    Dim expr As String
    
    expr = txtCst_Icms.Text
    If F_SoNumero(expr) <> "" Then
        expr = F_SoNumero(expr)
           
    End If
    
    p = Sped_Lista_Cst_Icms(expr)
    If p.codigo <> "" Then
        txtCst_Icms.Text = p.codigo & " - " & p.nome
        txtCst_Icms.tag = p.codigo
        txtCst_Ipi.SetFocus
    End If
    
    
End Sub

Private Sub pcCstIpi_Click()
    Dim p As Padrao1
    Dim expr As String
    
    expr = txtCst_Ipi.Text
    If F_SoNumero(expr) <> "" Then
        expr = F_SoNumero(expr)
           
    End If
    
    p = Sped_Lista_Cst_IPI(expr)
    If p.codigo <> "" Then
        txtCst_Ipi.Text = p.codigo & " - " & p.nome
        txtCst_Ipi.tag = p.codigo
        txtCst_Pis.SetFocus
    End If
End Sub

Private Sub pcCstPis_Click()
    Dim p As Padrao1
    Dim expr As String
    
    expr = txtCst_Pis.Text
    If F_SoNumero(expr) <> "" Then
        expr = F_SoNumero(expr)
           
    End If
    
    p = Sped_Lista_Cst_Pis(expr)
    If p.codigo <> "" Then
        txtCst_Pis.Text = p.codigo & " - " & p.nome
        txtCst_Pis.tag = p.codigo
        txtCst_Cofins.SetFocus
    End If
End Sub

Private Sub pcEnt_Click()
    If Me.tag <> "" Then
        P_InfoEntradasPro Me.tag
    Else
        MsgBox "Não há produto gravado!", vbCritical
    End If
End Sub

Sub buscaGenero()
    Dim p As Padrao1
    p = Sped_Genero(txtGenero.Text)
    If p.codigo <> "" Then
        txtGenero.Text = p.nome
        txtGenero.tag = p.codigo
        txtOrigem.SetFocus
    End If
End Sub
Sub buscaOrigem()
    Dim p As Padrao1
    p = Sped_Origem(txtOrigem.Text)
    If p.codigo <> "" Then
        txtOrigem.Text = p.nome
        txtOrigem.tag = p.codigo
        txtbxUnidade.SetFocus
    End If
End Sub


Private Sub pcGenero_Click()
    buscaGenero
End Sub

Private Sub pcIcms_Click()
    Dim i As imposto.imposto
    i = imposto.F_SelectImp("")
    If i.cod_imp <> "" Then
        txtbxICMS.Text = i.desc_imp
        txtbxICMS.tag = i.cod_imp
        
        
    End If
    
    
End Sub

Private Sub pcOrigem_Click()
    buscaOrigem
End Sub

Private Sub pcSubCat_Click()
    
End Sub

