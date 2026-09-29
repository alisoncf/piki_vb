VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmInventario 
   Caption         =   "Inventário"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   14475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   14475
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1335
      ScaleWidth      =   14475
      TabIndex        =   10
      Top             =   7350
      Width           =   14475
      Begin VB.PictureBox Picture6 
         Height          =   1095
         Left            =   2040
         ScaleHeight     =   1035
         ScaleWidth      =   2475
         TabIndex        =   47
         Top             =   120
         Width           =   2535
         Begin VB.CommandButton cmdProp 
            Caption         =   "Ok"
            Height          =   255
            Left            =   1560
            TabIndex        =   56
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtProp 
            Height          =   285
            Left            =   960
            TabIndex        =   55
            Text            =   "1/1"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtY 
            BackColor       =   &H0000FFFF&
            Height          =   285
            Left            =   960
            TabIndex        =   52
            Text            =   "1000"
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmdDiminuirQtd 
            Caption         =   "-"
            Height          =   255
            Left            =   1560
            TabIndex        =   51
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtX 
            Height          =   285
            Left            =   120
            TabIndex        =   49
            Text            =   "1"
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton Command6 
            Caption         =   "+"
            Height          =   255
            Left            =   2040
            TabIndex        =   48
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label12 
            Caption         =   "Proporção"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Em"
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
            Left            =   600
            TabIndex        =   53
            Top             =   240
            Width           =   270
         End
         Begin VB.Label Label10 
            Caption         =   "Fator Qtd"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture5 
         Height          =   1095
         Left            =   11160
         ScaleHeight     =   1035
         ScaleWidth      =   1635
         TabIndex        =   42
         Top             =   120
         Width           =   1695
         Begin VB.TextBox txtAlterarEst 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Command4 
            Caption         =   "..."
            Height          =   285
            Left            =   1080
            TabIndex        =   44
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Alterar estoque para inventáro"
            Height          =   555
            Left            =   120
            TabIndex        =   45
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inventário:"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   0
            Width           =   750
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Corrigir Custo Inv"
         Height          =   495
         Left            =   13080
         TabIndex        =   39
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Estoque inventário"
         Height          =   495
         Left            =   13080
         TabIndex        =   38
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdDuplicar 
         Caption         =   "Duplicar"
         Height          =   255
         Left            =   1200
         TabIndex        =   34
         Top             =   1080
         Width           =   855
      End
      Begin VB.PictureBox Picture4 
         Height          =   1095
         Left            =   9360
         ScaleHeight     =   1035
         ScaleWidth      =   1635
         TabIndex        =   32
         Top             =   120
         Width           =   1695
         Begin VB.TextBox txtInvAj 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   285
            Left            =   1080
            TabIndex        =   36
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdAj 
            Caption         =   "Ajustar Est a partir do Inv"
            Height          =   555
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inventário:"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   0
            Width           =   750
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   1095
         Left            =   8040
         ScaleHeight     =   1035
         ScaleWidth      =   1155
         TabIndex        =   29
         Top             =   120
         Width           =   1215
         Begin VB.TextBox txtRetornar 
            Height          =   285
            Left            =   120
            TabIndex        =   31
            Top             =   120
            Width           =   975
         End
         Begin VB.CommandButton cmdRetornar 
            Caption         =   "Retornar à"
            Height          =   315
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H0000FF00&
         Height          =   1095
         Left            =   6240
         ScaleHeight     =   1035
         ScaleWidth      =   1635
         TabIndex        =   25
         Top             =   120
         Width           =   1695
         Begin VB.TextBox txtAno 
            Height          =   285
            Left            =   600
            TabIndex        =   28
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton cmdSintegra 
            Caption         =   "Gerar Sintegra"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ano:"
            Height          =   195
            Left            =   240
            TabIndex        =   26
            Top             =   120
            Width           =   330
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   1095
         Left            =   4560
         ScaleHeight     =   1035
         ScaleWidth      =   1395
         TabIndex        =   21
         Top             =   120
         Width           =   1455
         Begin VB.CommandButton cmdMudar 
            Caption         =   "Alterar"
            Height          =   255
            Left            =   600
            TabIndex        =   24
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtCusto 
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Text            =   "1"
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "Fator Custo:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkT2 
         Caption         =   "Total 2"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox chkNeg 
         Caption         =   "Listar Negativas"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T1:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lt1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
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
         Left            =   600
         TabIndex        =   15
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T2:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   240
      End
      Begin VB.Label lt2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
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
         Left            =   600
         TabIndex        =   13
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Itens:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   390
      End
      Begin VB.Label lItens 
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
         Left            =   600
         TabIndex        =   11
         Top             =   600
         Width           =   120
      End
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   6255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   11033
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "codigo"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nome"
         Caption         =   "Nome"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "custo"
         Caption         =   "Custo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "preco"
         Caption         =   "Preço"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "q1"
         Caption         =   "Q1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "q2"
         Caption         =   "Q2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "total1"
         Caption         =   "Total1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "total2"
         Caption         =   "Total2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "r"
         Caption         =   "reservado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "ultqtd"
         Caption         =   "último qtd"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "ultcusto"
         Caption         =   "Últim Custo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1184,882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3525,166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1124,787
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   975,118
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1065,26
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   14475
      TabIndex        =   0
      Top             =   0
      Width           =   14475
      Begin VB.CommandButton cmdAnalise 
         Caption         =   "Análise"
         Height          =   495
         Left            =   12000
         TabIndex        =   57
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdInfoMov 
         Caption         =   "Info. Mov"
         Height          =   495
         Left            =   11040
         TabIndex        =   41
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdEdiProd 
         Caption         =   "Editar Produto"
         Height          =   495
         Left            =   9720
         TabIndex        =   40
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbOrd 
         Height          =   315
         ItemData        =   "frmInventario.frx":0000
         Left            =   2280
         List            =   "frmInventario.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   495
         Left            =   12840
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "Excluir"
         Height          =   495
         Left            =   7080
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdImpr 
         Caption         =   "Imprmir"
         Height          =   495
         Left            =   8400
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Editar Item"
         Height          =   495
         Left            =   5760
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "Novo"
         Height          =   495
         Left            =   4440
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtInv 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem:"
         Height          =   195
         Left            =   2280
         TabIndex        =   20
         Top             =   120
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inventário:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmbOrd_Click()
    
    busca ""
End Sub

Private Sub cmdAj_Click()
    
    
    Dim strSQL As String
    Dim rIt As New ADODB.Recordset
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    If txtInvAj.tag = "" Then
        MsgBox "Informe um inv!", vbExclamation
        Exit Sub
    End If
    
    If strNome_Usu <> "ALISON" Then
        MsgBox "Erro@", vbExclamation
        Exit Sub
    End If
    
    
    
    If MsgBox("Deseja realmente alterar o estoque do inventário para a data?", vbYesNo + vbQuestion) = vbYes Then
        P_Status "Aguarde..."
        grid.Visible = False
        With ado
            .MoveFirst
            Do While Not .EOF
                
                
                strSQL = "select * from inventario_it as ti where inventario=" & txtInvAj.tag
                
                
                strSQL = strSQL & " and ti.codigo='" & !codigo & "'"
                
                abreObjetoAdo rIt, strSQL
                
                With rIt
                    
                    If .RecordCount > 0 Then
                        strSQL = "update inventario_it set q1=" & F_GravaDec(!q1) & ","
                        strSQL = strSQL & " q2=" & F_GravaDec(!q2) & " "
                        strSQL = strSQL & " where id=" & ado!id
                        Conn.Execute strSQL
                        
                    End If
                    
                    .Close
                End With
                
                
                
                
                
                strSQL = "select sum(ti.qtd_ite) as total from tab_itvendas as ti, tab_vendas as tv "
                
                strSQL = strSQL & " where tv.dataefet_ven>=" & F_GravaDataTempo("01/01/" & Year(Date) - 1)
                strSQL = strSQL & " and tv.num_ven=ti.num_ven "
                strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaDataTempo("31/12/" & Year(Date) - 1 & " 23:59:00")
                strSQL = strSQL & " and ti.cod_pro='" & !codigo & "'"
                strSQL = strSQL & " and tipo_operacao=1"
                abreObjetoAdo rIt, strSQL
                
                With rIt
                    
                    If .RecordCount > 0 And Not IsNull(!total) Then
                        strSQL = "update inventario_it set q1=q1-" & F_GravaDec(!total)
                        strSQL = strSQL & " where id=" & ado!id
                        Conn.Execute strSQL
                        
                    End If
                    
                    .Close
                End With
                
                
                strSQL = "select sum(ti.qtd_ite) as total from tab_itcompras as ti, tab_compras as tc "
                strSQL = strSQL & " where tc.emissao_not>=" & F_GravaData("01/01/" & Year(Date) - 1)
                strSQL = strSQL & " and tc.seq_not=ti.seq_not "
                strSQL = strSQL & " and tc.emissao_not<=" & F_GravaData("31/12/" & Year(Date) - 1)
                strSQL = strSQL & " and ti.cod_pro='" & !codigo & "'"
                
                abreObjetoAdo rIt, strSQL
                
                With rIt
                    
                    If .RecordCount > 0 And Not IsNull(!total) Then
                        strSQL = "update inventario_it set q1=q1+" & F_GravaDec(!total)
                        strSQL = strSQL & " where id=" & ado!id
                        Conn.Execute strSQL
                        
                    End If
                    
                    .Close
                End With
                P_Status .AbsolutePosition & " de " & .RecordCount
                .MoveNext
            Loop
            .MoveFirst
        End With
        
        
        
        
        
        
        P_Status ""
        
        
    End If
    grid.Visible = True
    P_Status ""
    Exit Sub
trata_Erro:
    P_MostraErro
    P_Status ""
grid.Visible = True
End Sub

Private Sub cmdAnalise_Click()
    frmInventarioAnalise.txtInv.text = txtInv.text
    frmInventarioAnalise.txtInv.tag = txtInv.tag
    frmInventarioAnalise.Show 1
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo trata_Erro
    Dim i As inventario.inventario
    Dim ord As String
    If cmbOrd.ListIndex = 1 Then
        ord = "nome_pro"
    ElseIf cmbOrd.ListIndex = 2 Then
    Else
    
    End If
    i = inventario.SelecionaInventario
    If i.id <> "" Then
        txtInv.text = i.data
        
        txtInv.tag = i.id
        busca ord
    End If
    Exit Sub
trata_Erro:
    P_MostraErro
    
End Sub
Sub busca(Optional ordem As String = "nome")
    Dim strSQL As String
    strSQL = "select it.*,it.custo*it.q1 as total1,i.de,i.ate,p.cf_pro  "
    
    If ordem = "" Then
        If cmbOrd.ListIndex = 1 Then
        
        ElseIf cmbOrd.ListIndex = 2 Then
            ordem = "total1 desc"
         ElseIf cmbOrd.ListIndex = 3 Then
            ordem = "it.q1 desc, total1 desc"
         ElseIf cmbOrd.ListIndex = 4 Then
            ordem = "total1 desc,it.q1 desc"
        
        End If
    End If
    
    If chkT2.Value = 1 Then
        strSQL = strSQL & ", custo*q2 as total2 "
        grid.Columns(7).Visible = True
    Else
        grid.Columns(7).Visible = False
    End If
    strSQL = strSQL & " from inventario_it as it inner join inventario as i "
    strSQL = strSQL & " on i.id=it.inventario "
    'strSQL = strSQL & " left join inventario as i2 on i.id_ultinv=i2.id"
    'strSQL = strSQL & " left join inventario_it as it2 on i2.id=it2.inventario and it2.codigo=it.codigo"
    strSQL = strSQL & " inner join tab_produtos as p on p.cod_pro=it.codigo "
    
    strSQL = strSQL & " where it.inventario=" & txtInv.tag
    If chkNeg.Value = 0 Then
        strSQL = strSQL & " and it.q1>0 "
    End If
    
    If ordem <> "" Then
        strSQL = strSQL & " order by " & ordem
    End If
    abreObjetoAdo ado, strSQL, grid
    total
    On Error Resume Next
    grid.caption = "Inventário " & ado!inventario & " - " & ado.RecordCount & " itens"
    
End Sub

Private Sub cmdDiminuirQtd_Click()
    Dim strSQL As String
    If MsgBox("Confirma diminuição de  " & txtX.text & " para todos os itens que tiverem mais que " & txtY.text & "?", vbYesNo + vbQuestion) = vbYes Then
        strSQL = " update inventario_it set q1=q1-" & F_GravaDec(txtX.text)
        strSQL = strSQL & " where inventario=" & txtInv.tag
        strSQL = strSQL & " and q1>" & F_GravaDec(txtY.text)
        Conn.Execute strSQL
        busca
    End If
End Sub

Private Sub cmdDuplicar_Click()
    If MsgBox("Confirma Duplicação do Inventário para Data Atual?", vbYesNo + vbQuestion) = vbYes Then
        duplica txtInv.tag
    End If
End Sub
Sub duplica(id_inv As String)
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    
    If txtInv.tag = "" Then
        MsgBox "Escolha um inventário!", vbExclamation
        Exit Sub
    End If
    id = inventario.geraInventario(Now, Date, Date)
    If id <> "" Then
        
        
        strSQL = strSQL & "select '" & id & "' as inventario, codigo,nome,custo,preco,q1,q2 from inventario_it as i where i.inventario=" & id_inv
        abreObjetoAdo rPesq, strSQL
        
        With rPesq
            Do While Not .EOF
                strSQL = "insert into inventario_it(inventario,codigo,nome,custo,preco,q1,q2) values("
                strSQL = strSQL & !inventario & ","
                strSQL = strSQL & "'" & !codigo & "',"
                strSQL = strSQL & "'" & !nome & "',"
                strSQL = strSQL & F_GravaDec(!custo) & ","
                strSQL = strSQL & F_GravaDec(!preco) & ","
                strSQL = strSQL & F_GravaDec(!q1) & ","
                strSQL = strSQL & F_GravaDec(!q2) & ")"
                Conn.Execute strSQL
                P_Status "Gravando item " & .AbsolutePosition & " de " & .RecordCount
                .MoveNext
            Loop
            .Close
        End With
        
        
        
        P_Status
        MsgBox "Duplicado!", vbExclamation
    End If
    
    
    
End Sub

Private Sub cmdEdiProd_Click()
    On Error Resume Next
    Dim frm As New frmCadPro
    frm.tag = ado!codigo
    frm.Show 1
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo trata_Erro
    With ado
        Dim i As Integer
        frmAlteraEstInv.lblId.caption = !id
        frmAlteraEstInv.txtCusto.text = FormatVal(!custo)
        frmAlteraEstInv.txtQ1.text = FormatVal(!q1)
        frmAlteraEstInv.txtQ2.text = FormatVal(!q2)
        frmAlteraEstInv.lblProduto.tag = !codigo
        frmAlteraEstInv.lblProduto.caption = !codigo & "-" & !nome
     '   frmAlteraEstInv.txtDe.Text = !de
      '  frmAlteraEstInv.txtAte.Text = !ate
        frmAlteraEstInv.Show 1
        
        i = ado.AbsolutePosition
        ado.Requery
        total
        ado.Move i
        
        
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmdImpr_Click()
    P_DbgrdToHtml "Inventário " & txtInv.text, grid, ado, "Total: " & lt1.caption, "6,7"
End Sub
Sub total()
    Dim i As Currency
    Dim t1 As Currency
    Dim t2 As Currency
    On Error GoTo trata_Erro
    grid.Visible = False
    lt1.caption = "0"
    lt2.caption = "0"
    lItens.caption = 0
    With ado
        .MoveFirst
        lItens.caption = .RecordCount
        Do While Not .EOF
            t1 = t1 + !q1 * !custo
            t2 = t2 + !q2 * !custo
            .MoveNext
        Loop
        lt1.caption = FormatVal(t1)
        lt2.caption = FormatVal(t2)
        .MoveFirst
    End With
grid.Visible = True
trata_Erro:
grid.Visible = True
End Sub

Private Sub cmdInfoMov_Click()
    P_HistProd ado!codigo
End Sub

Private Sub cmdMudar_Click()
    Dim strSQL As String
    If MsgBox("Deseja multiplicar o valor de todos os produtos por " & txtCusto.text & "?", vbYesNo + vbQuestion) = vbYes Then
        strSQL = " update inventario_it set custo=custo*" & F_GravaDec(txtCusto.text)
        strSQL = strSQL & " where inventario=" & txtInv.tag
        Conn.Execute strSQL
        busca
    End If
End Sub

Private Sub cmdNovo_Click()
    frmBuscaInv.Show 1
End Sub

Private Sub cmdProp_Click()
    Dim strSQL As String
    If MsgBox("Confirma a alteração do estoque para proporção  " & txtProp.text & " para todos os itens?", vbYesNo + vbQuestion) = vbYes Then
        strSQL = " update inventario_it set q1=(" & txtProp.text & "*q1)"
        'strSQL = " update inventario_it set q1=cast(" & txtProp.Text & "*q1  as UNSIGNED)"
        strSQL = strSQL & " where inventario=" & txtInv.tag
        strSQL = strSQL & " and q1>0"
        Conn.Execute strSQL
        busca
    End If
End Sub

Private Sub cmdRetornar_Click()
    Dim strSQL As String
    Dim rIt As New ADODB.Recordset
    Dim rPesq As New ADODB.Recordset
    grid.Visible = False
    On Error GoTo trata_Erro
    If txtRetornar.text = "" Then
        MsgBox "Informe uma data!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Deseja realmente alterar o estoque do inventário para a data?", vbYesNo + vbQuestion) = vbYes Then
        
        
        With ado
            .MoveFirst
            Do While Not .EOF
            
            
            
            
            
            
                strSQL = "select sum(ti.qtd_ite) as total from tab_itvendas as ti, tab_vendas as tv "
                strSQL = strSQL & " where ti.num_ven=tv.num_ven "
                strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaDataTempo(txtRetornar.text)
                strSQL = strSQL & " and ti.cod_pro='" & !codigo & "'"
                strSQL = strSQL & " and tipo_operacao=1"
                abreObjetoAdo rIt, strSQL
                
                With rIt
                    
                    If .RecordCount > 0 And Not IsNull(!total) Then
                        strSQL = "update inventario_it set q2=q2+" & F_GravaDec(!total)
                        strSQL = strSQL & " where id=" & ado!id
                        Conn.Execute strSQL
                        
                    End If
                    
                    .Close
                End With
                
                
                strSQL = "select sum(ti.qtd_ite) as total from tab_itacertos as ti, tab_acertos as tv "
                strSQL = strSQL & " where ti.seq_ace=tv.seq_ace "
                strSQL = strSQL & " and tv.data_ace>" & F_GravaData(txtRetornar.text)
                strSQL = strSQL & " and status_ace=2"
                strSQL = strSQL & " and ti.cod_pro='" & !codigo & "'"
                
                abreObjetoAdo rIt, strSQL
                
                With rIt
                    
                    If .RecordCount > 0 And Not IsNull(!total) Then
                        strSQL = "update inventario_it set q2=q2+" & F_GravaDec(!total)
                        strSQL = strSQL & " where id=" & ado!id
                       ' Conn.Execute strSQL
                        
                    End If
                    
                    .Close
                End With
                
                
                strSQL = "select sum(ti.qtd_its) as total from tab_itsai as ti, tab_saidas as tv "
                strSQL = strSQL & " where ti.nota_sai=tv.nota_sai "
                strSQL = strSQL & " and tv.emissao_sai>" & F_GravaData(txtRetornar.text)
                strSQL = strSQL & " and ti.cod_pro='" & !codigo & "'"
                
                abreObjetoAdo rIt, strSQL
                
                With rIt
                    
                    If .RecordCount > 0 And Not IsNull(!total) Then
                        strSQL = "update inventario_it set q1=q1+" & F_GravaDec(!total)
                        strSQL = strSQL & " where id=" & !id
                        Conn.Execute strSQL
                        
                    End If
                    
                    .Close
                End With
                
                P_Status ado.AbsolutePosition
                .MoveNext
            Loop
            .MoveFirst
            grid.Visible = True
            P_Status ""
        End With
    End If
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub



Private Sub cmdSintegra_Click()
    If txtAno.text = "" Then
        txtAno = Year(Date)
        
    End If
    
    
    GeraInventarioSintegra ado!inventario, "01/12/" & txtAno.text, "31/12/" & txtAno.text
    
    
    
    
    
    
End Sub

Private Sub Command1_Click()
        Dim i As inventario.inventario
    Dim ord As String
    If cmbOrd.ListIndex = 1 Then
        ord = "nome_pro"
    ElseIf cmbOrd.ListIndex = 2 Then
    Else
    
    End If
    i = inventario.SelecionaInventario
    If i.id <> "" Then
        txtInvAj.text = i.data
        
        txtInvAj.tag = i.id
        
    End If

End Sub

Private Sub Command2_Click()
    Dim q1 As Currency
    Dim q2 As Currency
    Dim e1 As Currency
    Dim e2 As Currency
    Dim i1 As Currency
    Dim i2 As Currency
    Dim v1 As Currency
    Dim v2 As Currency
    Dim c1 As Currency
    Dim c2 As Currency
    Dim Qa As Currency
    Dim strSQL As String
    Dim rIt As New ADODB.Recordset
    If MsgBox("Deseja transformar o estoque no inventário?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    If Not F_senha Then Exit Sub
    If strNome_Usu <> "ALISON" Then
        MsgBox "Usuário inválido!", vbExclamation
        Exit Sub
    End If
    

    
    With ado
            .MoveFirst
            Do While Not .EOF
                
                strSQL = "select sum(ti.qtd_ite) as total from tab_itvendas as ti, tab_vendas as tv "
                strSQL = strSQL & " where ti.num_ven=tv.num_ven "
                strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtRetornar.text)
                strSQL = strSQL & " and ti.cod_pro='" & !codigo & "'"
                strSQL = strSQL & " and tipo_operacao=1"
                abreObjetoAdo rIt, strSQL
                
                With rIt
                    
                    If .RecordCount > 0 And Not IsNull(!total) Then
                        v2 = !total
                    Else
                        v2 = 0
                    End If
                    
                    .Close
                End With
                
                 strSQL = "select sum(ti.qtd_its) as total from tab_itsai as ti, tab_saidas as tv "
                strSQL = strSQL & " where ti.nota_sai=tv.nota_sai and ti.cod_cai=tv.cod_cai "
                strSQL = strSQL & " and tv.emissao_sai>" & F_GravaData(txtRetornar.text)
                strSQL = strSQL & " and ti.cod_pro='" & !codigo & "'"
                
                abreObjetoAdo rIt, strSQL
                
                With rIt
                    
                    If .RecordCount > 0 And Not IsNull(!total) Then
                        v1 = !total
                    Else
                        v1 = 0
                    End If
                    
                    .Close
                End With
                
                
                 strSQL = "select sum(ti.qtd_ite) as total from tab_itcompras as ti, tab_compras as tv "
                strSQL = strSQL & " where ti.seq_not=tv.seq_not"
                strSQL = strSQL & " and tv.emissao_not>" & F_GravaData(txtRetornar.text)
                strSQL = strSQL & " and ti.cod_pro='" & !codigo & "'"
                
                abreObjetoAdo rIt, strSQL
                
                With rIt
                    
                    If .RecordCount > 0 And Not IsNull(!total) Then
                        c1 = !total
                    Else
                        c1 = 0
                    End If
                    
                    .Close
                End With
                
                strSQL = "select sum(ti.qtd_ite) as total from tab_itcmp as ti, tab_cmp as tv "
                strSQL = strSQL & " where ti.seq_not=tv.seq_not"
                strSQL = strSQL & " and tv.emissao_not>" & F_GravaData(txtRetornar.text)
                strSQL = strSQL & " and ti.cod_pro='" & !codigo & "'"
                
                abreObjetoAdo rIt, strSQL
                
                With rIt
                    
                    If .RecordCount > 0 And Not IsNull(!total) Then
                        c2 = !total
                    Else
                        c2 = 0
                    End If
                    
                    .Close
                End With
                
                
                strSQL = "select sum(ti.qtd_ite) as total from tab_itacertos as ti, tab_acertos as tv "
                strSQL = strSQL & " where ti.seq_ace=tv.seq_ace"
                strSQL = strSQL & " and tv.data_ace>" & F_GravaData(txtRetornar.text)
                strSQL = strSQL & " and ti.cod_pro='" & !codigo & "'"
                strSQL = strSQL & " and status_ace=2 "
                abreObjetoAdo rIt, strSQL
                
                With rIt
                    
                    If .RecordCount > 0 And Not IsNull(!total) Then
                        c1 = !total
                    Else
                        c1 = 0
                    End If
                    
                    .Close
                End With
                
                e2 = CCur(F_BuscaUmCampo("er_pro", "tab_produtos", "cod_pro='" & !codigo & "'"))
                e1 = CCur(F_BuscaUmCampo("ef_pro", "tab_produtos", "cod_pro='" & !codigo & "'"))
                strSQL = !codigo & "-" & !nome & Chr(13)
                strSQL = strSQL & "Er=" & e2 & Chr(13) & "ef=" & e1 & Chr(13)
                strSQL = strSQL & "q1: " & !q1 & Chr(13)
                strSQL = strSQL & "q2: " & !q2 & Chr(13)
                strSQL = strSQL & "comprou (1): " & c1 & Chr(13)
                strSQL = strSQL & "comprou (2): " & c2 & Chr(13)
                strSQL = strSQL & "vendeu (1): " & v1 & Chr(13)
                strSQL = strSQL & "vendeu (2): " & v2 & Chr(13)
                strSQL = strSQL & "acertou (1): " & Qa & Chr(13)
                
                strSQL = strSQL & " er = (q2+c2)-(v2+qa) : " & !q2 + c2 - v2 - Qa & Chr(13)
                strSQL = strSQL & " ef = (q1+c1)-(v1) : " & !q1 + c1 - v1 & Chr(13)
                
                
                
                
                
'                MsgBox strSQL
                strSQL = "update tab_produtos set er_pro=" & F_GravaDec(!q2 + c2 - v2 - Qa) & ","
                strSQL = strSQL & " ef_pro=" & F_GravaDec(!q1 + c1 - v1) & " where cod_pro='" & !codigo & "'"
                Conn.Execute strSQL
                P_Status .AbsolutePosition & "/" & .RecordCount
            .MoveNext
            Loop
        '.Close
    End With
    P_Status
End Sub

Private Sub Command3_Click()
    Dim strSQL As String
    With ado
        Do While Not .EOF
            strSQL = "update inventario_it set custo=" & F_GravaDec(!cf_pro) & " where codigo='" & !codigo & "' "
            strSQL = strSQL & " and id=" & !id
            Conn.Execute strSQL
            .MoveNext
        Loop
    End With
End Sub

Private Sub Command4_Click()
    
    Dim i As inventario.inventario
    Dim ord As String
    If cmbOrd.ListIndex = 1 Then
        ord = "nome_pro"
    ElseIf cmbOrd.ListIndex = 2 Then
    Else
    
    End If
    i = inventario.SelecionaInventario
    If i.id <> "" Then
        txtAlterarEst.text = i.data
        
        txtAlterarEst.tag = i.id
        
    End If

End Sub

Private Sub Command5_Click()
    Dim strSQL As String
    Dim rIt As New ADODB.Recordset
    Dim rPesq As New ADODB.Recordset
    Dim i As Integer
    On Error GoTo trata_Erro
    If txtAlterarEst.tag = "" Then
        MsgBox "Informe um inv!", vbExclamation
        Exit Sub
    End If
    
    If strNome_Usu <> "ALISON" Then
        MsgBox "Erro@", vbExclamation
        Exit Sub
    End If
    
    
    
    If MsgBox("Deseja realmente alterar o estoque para este inventário?", vbYesNo + vbQuestion) = vbYes Then
        
        grid.Visible = False
        
        
            
            
            
            strSQL = "select * from inventario_it where inventario=" & txtAlterarEst.tag
           ' strSQL = strSQL & " and q1 > 0 "
                       
            abreObjetoAdo rPesq, strSQL
            If rPesq.RecordCount = 0 Then
                Exit Sub
            End If
            
            
            strSQL = "update tab_produtos set  ef_pro=0,er_pro=0 "
            Conn.Execute strSQL
            
            With rPesq
                Do While Not .EOF
                    
                        
                        If !q1 > 0 Then
                            strSQL = "update tab_produtos set ef_pro=" & F_GravaDec(!q1) & ","
                            strSQL = strSQL & " er_pro=" & F_GravaDec(!q1) & ","
                            strSQL = strSQL & " cr_pro=" & F_GravaDec(!custo)
                            strSQL = strSQL & " where cod_pro='" & rPesq!codigo & "' "
                            Conn.Execute strSQL
                            
                        End If
                        
                        
                    Me.caption = .AbsolutePosition & " de " & .RecordCount
                    If .AbsolutePosition Mod 50 = 0 Then
                        Me.Refresh
                    End If
                    .MoveNext
                Loop
                
            
            End With
        

        
            strSQL = "select distinct ti.seq_not,ti.seq_ite,ti.cod_pro,ti.custo_ite,tp.cr_pro,tp.pv_pro,status_not,"
            strSQL = strSQL & " (ti.qtd_ite) as total"
            strSQL = strSQL & " from tab_compras as tc  inner join tab_itcompras as ti on tc.seq_not=ti.seq_not "
            strSQL = strSQL & " inner join tab_produtos as tp on tp.cod_pro=ti.cod_pro"
            strSQL = strSQL & " and status_not in (2,3,4)"
            strSQL = strSQL & " where tc.emissao_not>=" & F_GravaData("01/01/" & Year(Date) - 1)
            strSQL = strSQL & " and tc.emissao_not<=" & F_GravaData("31/12/" & Year(Date) - 1)
            strSQL = strSQL & "order by cod_pro "
            
            
            GravaArquivoExternoUTF8 App.path & "\inv.sql", strSQL
            'strSQL = strSQL & " and ti.cod_pro='" & !codigo & "'"
            
            abreObjetoAdo rIt, strSQL
            
                With rIt
                    
                    Do While Not .EOF
                        strSQL = "update tab_produtos set ef_pro=ef_pro+" & F_GravaDec(!total) & ","
                        strSQL = strSQL & " er_pro=er_pro+" & F_GravaDec(!total) & " "
                        strSQL = strSQL & " where cod_pro=" & !cod_pro
                        Conn.Execute strSQL
                  '
                        If .AbsolutePosition Mod 10 = 0 Then
                            Me.caption = "Compras..." & .AbsolutePosition & " de " & .RecordCount
                            Me.Refresh
                        End If
                        .MoveNext
                    Loop
                    
                    .Close
                End With
        
        
                
        
        strSQL = "select cprod,sum(qcom) as total from nfe as n inner join nfe_item as i on n.id=i.id_nfe"
                strSQL = strSQL & " where status=2 "
                strSQL = strSQL & " and left(dhemi,10) >=" & F_GravaData("01/01/" & Year(Date) - 1)
                strSQL = strSQL & " and left(dhemi,10)<=" & F_GravaData("31/12/" & Year(Date) - 1)
                strSQL = strSQL & " group by cprod"
                
                abreObjetoAdo rIt, strSQL
                
                With rIt
                    
                    Do While Not .EOF
                        
                        
                            
                            strSQL = "update tab_produtos set  "
                            strSQL = strSQL & " ef_pro=ef_pro-" & F_GravaDec(!total)
                            strSQL = strSQL & " where cod_pro='" & !cProd & "' "
                            Conn.Execute strSQL, i
                            If .AbsolutePosition = 1 Then
                              '  MsgBox "primeiro registro"
                            End If
                            
                        If .AbsolutePosition Mod 10 = 0 Then
                            Me.caption = "nfe..." & .AbsolutePosition & " de " & .RecordCount
                        End If
                        .MoveNext
                    Loop
                    
                    .Close
                End With
        
        
        
        P_Status ""
        
        
    End If
    grid.Visible = True
    P_Status ""
    Exit Sub
trata_Erro:
    P_MostraErro
    P_Status ""
grid.Visible = True
End Sub





Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    'txtApartir.Text = "01/01/" & Year(Date) - 1
    'txtAte.Text = "31/12/" & Year(Date) - 1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    PosicionaGrade grid, Me, pcBot
End Sub

Private Sub grid_DblClick()
    cmdEdit_Click
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdEdit_Click
    End If
End Sub

Private Sub txtRetornar_GotFocus()
    P_SelecTexto txtRetornar
End Sub

Private Sub txtRetornar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtRetornar_LostFocus()
    P_AjustaData txtRetornar
End Sub
