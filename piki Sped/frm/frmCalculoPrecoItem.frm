VERSION 5.00
Begin VB.Form frmCalculoPrecoItem 
   Caption         =   "Cálculo de Preço do Produto"
   ClientHeight    =   3180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11715
   Icon            =   "frmCalculoPrecoItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   11715
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1545
      ScaleWidth      =   9465
      TabIndex        =   31
      Top             =   1560
      Width           =   9495
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   7680
         ScaleHeight     =   825
         ScaleWidth      =   1065
         TabIndex        =   60
         Top             =   120
         Width           =   1095
         Begin VB.TextBox txtCustoNovo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Height          =   285
            Left            =   120
            TabIndex        =   61
            Text            =   "0,00"
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H0000C000&
            Caption         =   "Custo Novo"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   63
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor ($)"
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   6600
         ScaleHeight     =   825
         ScaleWidth      =   1065
         TabIndex        =   56
         Top             =   120
         Width           =   1095
         Begin VB.TextBox txtCAnterior 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   57
            Text            =   "0,00"
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "Custo Anterior"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   59
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor ($)"
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.PictureBox pcCofins 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   4440
         ScaleHeight     =   825
         ScaleWidth      =   1065
         TabIndex        =   52
         Top             =   120
         Width           =   1095
         Begin VB.TextBox txtVlCofinsItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   53
            Text            =   "0,00"
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label74 
            Alignment       =   2  'Center
            Caption         =   "COFINS"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -120
            TabIndex        =   55
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor ($)"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   5520
         ScaleHeight     =   825
         ScaleWidth      =   1065
         TabIndex        =   48
         Top             =   120
         Width           =   1095
         Begin VB.TextBox txtDesconto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   49
            Text            =   "0,00"
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor ($)"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "Desconto:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -120
            TabIndex        =   50
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.PictureBox pcIpi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   2280
         ScaleHeight     =   825
         ScaleWidth      =   1065
         TabIndex        =   44
         Top             =   120
         Width           =   1095
         Begin VB.TextBox txtVlIpiItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   45
            Text            =   "0,00"
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label66 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            Caption         =   "IPI"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   47
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor ($)"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.PictureBox pcIcms 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         ScaleHeight     =   825
         ScaleWidth      =   1065
         TabIndex        =   40
         Top             =   120
         Width           =   1095
         Begin VB.TextBox txtVlICMSItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   41
            Text            =   "0,00"
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label77 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor ($)"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label75 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Caption         =   "ICMS"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.PictureBox pcPis 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   3360
         ScaleHeight     =   825
         ScaleWidth      =   1065
         TabIndex        =   36
         Top             =   120
         Width           =   1095
         Begin VB.TextBox txtVlPisItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   37
            Text            =   "0,00"
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor ($)"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Caption         =   "PIS"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.PictureBox pcIcmsSt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   1200
         ScaleHeight     =   825
         ScaleWidth      =   1065
         TabIndex        =   32
         Top             =   120
         Width           =   1095
         Begin VB.TextBox txtVlIcmsStItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   33
            Text            =   "0,00"
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label33 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Caption         =   "ICMS ST"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -120
            TabIndex        =   35
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor ($)"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   585
         End
      End
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   615
      Left            =   9720
      Picture         =   "frmCalculoPrecoItem.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   855
      Left            =   9720
      Picture         =   "frmCalculoPrecoItem.frx":2258
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1560
      Width           =   1815
   End
   Begin VB.PictureBox pcProd 
      Align           =   1  'Align Top
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   11655
      TabIndex        =   0
      Top             =   0
      Width           =   11715
      Begin VB.TextBox txtUnitConvert 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   18
         Text            =   "0,0000"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtQtdConvert 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Text            =   "0"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtCean 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         MaxLength       =   50
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtPF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9600
         TabIndex        =   28
         Text            =   "0,00"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtPG 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "0,00"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtPV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   6000
         TabIndex        =   22
         Text            =   "0,00"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtLucro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7200
         TabIndex        =   24
         Text            =   "0,00"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtbxUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F8AC67&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   16
         Text            =   "0,0000"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FEF4D8&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Text            =   "0"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtbxTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4800
         TabIndex        =   20
         Text            =   "0,00"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtProd 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   2
         Top             =   240
         Width           =   5415
      End
      Begin VB.TextBox txtCst 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   8160
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtUnidEntr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   9600
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtEstoque 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7200
         TabIndex        =   6
         Text            =   "0,00"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit. Convert."
         Height          =   195
         Left            =   3600
         TabIndex        =   17
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd Convertida"
         Height          =   195
         Left            =   1200
         TabIndex        =   13
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço Futuro:"
         Height          =   195
         Left            =   9600
         TabIndex        =   27
         Top             =   600
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço Sugerido"
         Height          =   195
         Left            =   8160
         TabIndex        =   25
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço Atual"
         Height          =   195
         Left            =   6000
         TabIndex        =   21
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M. Lucro(%)"
         Height          =   195
         Left            =   7200
         TabIndex        =   23
         Top             =   600
         Width           =   840
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   4800
         TabIndex        =   19
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Qtd Entrada:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Item"
         Height          =   195
         Left            =   2400
         TabIndex        =   15
         Top             =   600
         Width           =   705
      End
      Begin VB.Label lblExpr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Produto"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Tag             =   "B"
         Top             =   0
         Width           =   555
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "Em estoque:"
         Height          =   195
         Left            =   7200
         TabIndex        =   5
         Top             =   0
         Width           =   885
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EAN:"
         Height          =   195
         Left            =   5640
         TabIndex        =   3
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CST:"
         Height          =   195
         Left            =   8160
         TabIndex        =   7
         Top             =   0
         Width           =   360
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Und Entr."
         Height          =   195
         Left            =   9600
         TabIndex        =   9
         Top             =   0
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmCalculoPrecoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Id_nf As String
Public id_item As String
Dim ado As New ADODB.Recordset

Sub busca()
    Dim strSQL As String
    Dim i As Integer
    
    On Error GoTo trata_erro:
    strSQL = "select tp.*,tp.cod_pro as codigo, tp.nome_pro as produto,nome_mar as marca, nome_cat as categoria, tp.ncm_pro as ncm, ti.*,"
    
    strSQL = strSQL & " tp.cr_pro as cr, tp.pv_pro as pv "
    strSQL = strSQL & " from sped_nf_itens as ti left join " & TAB_PRODUTOS & " as tp on tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " left join " & TAB_MARCAS & " as tm on tm.cod_mar=tp.cod_mar"
    strSQL = strSQL & " left join tab_categorias as tc on tc.cod_cat=tp.cod_cat "
    strSQL = strSQL & " where ti.id_nf ='" & Id_nf & "' "
    strSQL = strSQL & " and ti.id='" & id_item & "' "
    
    
    
    abreObjetoAdo ado, strSQL
    With ado
        If .RecordCount > 0 Then
            txtProd.Text = !nome_pro
            txtProd.tag = !cod_pro
            txtbxQtd.Text = !qtd
            txtUnidEntr.Text = !unid
            txtCean.Text = !cEAN
            
            If !qtd_convert > 0 Then
                txtUnitConvert.Text = FormatVal(!vl_item / !qtd_convert)
            Else
                txtUnitConvert.Text = "1"
            End If
            
            txtPV.Text = FormatVal(!pv_pro)
            txtLucro.Text = FormatVal(!lucro_pro)
            txtbxUnit.Text = FormatVal(!vl_item)
            txtEstoque.Text = FormatVal(!er_pro)
            txtCst.Text = !cst_icms
            txtVlCofinsItem.Text = FormatVal(!vl_cofins / CCur(txtUnitConvert.Text), "###,##0.0000")
            txtVlICMSItem.Text = FormatVal(!vl_icms / CCur(txtUnitConvert.Text), "###,##0.0000")
            txtVlIcmsStItem.Text = FormatVal(!VL_ICMS_ST / CCur(txtUnitConvert.Text), "###,##0.0000")
            txtVlIpiItem.Text = FormatVal(!VL_IPI / CCur(txtUnitConvert.Text), "###,##0.0000")
            txtVlPisItem.Text = FormatVal(!vl_pis / CCur(txtUnitConvert.Text), "###,##0.0000")
            txtVlCofinsItem.Text = FormatVal(!vl_cofins / CCur(txtUnitConvert.Text), "###,##0.0000")
            txtDesconto.Text = FormatVal(!VL_DESC / CCur(txtUnitConvert.Text), "###,##0.0000")
            txtQtdConvert.Text = FormatVal(!qtd_convert, "###,##0.00")
            txtCAnterior.Text = FormatVal(!cr_pro)
            txtPF.Text = FormatVal(!pv_pro)
            
            calcula_preco
            
            
        End If
    End With
    Exit Sub
trata_erro:
    MostraErro
    cmdSair_Click
End Sub
Sub calcula_preco()
    Dim gastos As Currency
    gastos = CCur(txtVlCofinsItem.Text) + CCur(txtVlICMSItem.Text) + CCur(txtVlIcmsStItem.Text) + CCur(txtVlIpiItem.Text) + _
        CCur(txtVlPisItem.Text) - CCur(txtDesconto.Text)
    txtCustoNovo.Text = FormatVal(gastos + CCur(txtUnitConvert.Text))
    
    gastos = gastos / CCur(txtQtdConvert.Text)
    
    txtPG.Text = FormatVal(CCur(txtCustoNovo.Text) + (CCur(txtLucro.Text) / 100 * CCur(txtCustoNovo.Text)))
    
        
        
End Sub

Private Sub cmdConfirmar_Click()
    Dim strSQL As String
    GravaNfeItensPreco id_item, txtPG.Text, txtPF.Text, Now
    
    If MsgBox("Deseja atribuir valores e custos do item para o PRODUTO: " & txtProd.Text & " agora?", vbYesNo + vbQuestion) = vbYes Then
        strSQL = "update " & TAB_PRODUTOS & " set cr_pro=" & F_GravaDec(txtCustoNovo.Text) & ","
        strSQL = strSQL & "pv_pro=" & F_GravaDec(txtPF.Text)
        strSQL = strSQL & " where cod_pro='" & txtProd.tag & "' "
        conn.Execute strSQL
        MsgBox "Produto alterado com sucesso!", vbInformation
        Unload Me
    End If
    
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    busca
End Sub

Private Sub txtLucro_GotFocus()
    ganhaFoco txtLucro
End Sub

Private Sub txtLucro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtLucro_LostFocus()
    PerdeFoco txtLucro
    calcula_preco
End Sub

Private Sub txtPF_GotFocus()
    txtPF.Text = txtPG.Text
    ganhaFoco txtPF
End Sub

Private Sub txtPF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtPF_LostFocus()
    PerdeFoco txtPF
End Sub

Private Sub txtPG_GotFocus()
    ganhaFoco txtPG
End Sub

Private Sub txtPG_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtPG_LostFocus()
    PerdeFoco txtPG
End Sub
