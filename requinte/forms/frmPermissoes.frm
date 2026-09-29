VERSION 5.00
Begin VB.Form frmPermissoes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Permissões de:"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame9 
      Caption         =   "Estoque"
      Height          =   1335
      Left            =   120
      TabIndex        =   56
      Top             =   6000
      Width           =   3135
      Begin VB.CheckBox Check51 
         Caption         =   "Cadastrar Produtos"
         Height          =   195
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   2175
      End
      Begin VB.CheckBox Check50 
         Caption         =   "Controlar Preços"
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox Check49 
         Caption         =   "Lançar Notas Fiscais"
         Height          =   195
         Left            =   120
         TabIndex        =   58
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox Check48 
         Caption         =   "Efetivar Notas Fiscais"
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   960
         Width           =   2895
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Consignação"
      Height          =   3135
      Left            =   6600
      TabIndex        =   43
      Top             =   2040
      Width           =   3135
      Begin VB.CheckBox Check47 
         Caption         =   "Imprimir Movimentação"
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   2640
         Width           =   2775
      End
      Begin VB.CheckBox Check46 
         Caption         =   "Visualizar Relatórios Consignado"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   2400
         Width           =   2775
      End
      Begin VB.CheckBox Check44 
         Caption         =   "Movimentar Caixa Consignado"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   2160
         Width           =   2775
      End
      Begin VB.CheckBox Check43 
         Caption         =   "Excluir Lançamento Repr"
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CheckBox Check42 
         Caption         =   "Movimentar Conta Repr."
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CheckBox Check41 
         Caption         =   "Efetuar Ajuste Consignado"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox Check40 
         Caption         =   "Pagar Comissão Repr."
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CheckBox Check39 
         Caption         =   "Excluir Acerto"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   960
         Width           =   2775
      End
      Begin VB.CheckBox Check38 
         Caption         =   "Efetuar Transferência"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox Check37 
         Caption         =   "Efetuar Devolução"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   480
         Width           =   2175
      End
      Begin VB.CheckBox Check36 
         Caption         =   "Efetuar Acerto"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Cliente:"
      Height          =   1935
      Left            =   6600
      TabIndex        =   33
      Top             =   120
      Width           =   3135
      Begin VB.CheckBox Check34 
         Caption         =   "Manipular SPC"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox Check32 
         Caption         =   "Liberar Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox Check31 
         Caption         =   "Bloquear Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox Check30 
         Caption         =   "Alterar Limite"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   2175
      End
      Begin VB.CheckBox Check29 
         Caption         =   "Cadastrar Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   2295
      End
      Begin VB.CheckBox Check27 
         Caption         =   "Visualizar Informações"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   1200
         Width           =   2175
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Financeiro:"
      Height          =   3135
      Left            =   3360
      TabIndex        =   24
      Top             =   2040
      Width           =   3135
      Begin VB.CheckBox Check35 
         Caption         =   "Editar Documento"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox Check26 
         Caption         =   "Baixar Documento"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   2175
      End
      Begin VB.CheckBox Check25 
         Caption         =   "Lançar Documento"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox Check24 
         Caption         =   "Lançar Cheque"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox Check23 
         Caption         =   "Baixar Cheque"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CheckBox Check22 
         Caption         =   "Excluir Documento"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox Check21 
         Caption         =   "Excluir Cheque"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CheckBox Check15 
         Caption         =   "Fechar Mês"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CheckBox Check14 
         Caption         =   "Cadastrar Formas de Pagamento"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   2775
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Representante:"
      Height          =   1935
      Left            =   3360
      TabIndex        =   17
      Top             =   120
      Width           =   3135
      Begin VB.CheckBox Check13 
         Caption         =   "Visualizar Informações"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox Check20 
         Caption         =   "Alterar Comissão"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   2175
      End
      Begin VB.CheckBox Check19 
         Caption         =   "Cadastrar Representante"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   2295
      End
      Begin VB.CheckBox Check18 
         Caption         =   "Alterar Limite"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox Check17 
         Caption         =   "Bloquear Representante"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox Check16 
         Caption         =   "Liberar Representante"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Venda:"
      Height          =   2655
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   3135
      Begin VB.CheckBox Check45 
         Caption         =   "Efetuar Troca"
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CheckBox Check12 
         Caption         =   "Imprimir Venda"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Cancelar Item Atual"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Cancelar Ultimo Item"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Cancelar Item Genérico"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Cancelar Venda"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Finalizar Venda"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Abrir Venda"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Fechar Venda"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Crediário:"
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   3135
      Begin VB.CheckBox Check4 
         Caption         =   "Cancelar Crediário"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2175
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Resolver Pendências"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controle de Usuários:"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.CheckBox chkUsu 
         Caption         =   "Cadastrar Usuários"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Funcionário:"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   3135
      Begin VB.CheckBox Check33 
         Caption         =   "Baixar Parc/Funcionários"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   960
         Width           =   2895
      End
      Begin VB.CheckBox Check28 
         Caption         =   "Lançar Vales/Parc de Funcionários"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Visualizar Dados de Funcionários"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Cadastrar Funcionário"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmPermissoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
