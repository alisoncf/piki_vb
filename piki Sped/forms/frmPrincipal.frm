VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "Piki - SPED"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   11160
   Icon            =   "frmPrincipal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnGerarSintegra 
         Caption         =   "Gerar Sintegra"
      End
      Begin VB.Menu div1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSair 
         Caption         =   "&Sair"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnECF 
      Caption         =   "ECF"
      Begin VB.Menu mnRedZ 
         Caption         =   "Registros de Redução Z"
      End
      Begin VB.Menu mnCodTotRedZ 
         Caption         =   "Códigos de Totalizadores de Redução Z"
      End
   End
   Begin VB.Menu mnEmpresa 
      Caption         =   "&Empresa"
      Begin VB.Menu mnAtivoImob 
         Caption         =   "Ativo Imobilizado"
      End
      Begin VB.Menu mnDados 
         Caption         =   "Dados da Empresa"
      End
      Begin VB.Menu mnParticipantes 
         Caption         =   "Participantes"
      End
      Begin VB.Menu mnContabilista 
         Caption         =   "Contabilista"
      End
      Begin VB.Menu mnUnidades 
         Caption         =   "Unidades"
      End
      Begin VB.Menu mnGeneros 
         Caption         =   "Gêneros"
      End
      Begin VB.Menu mn411 
         Caption         =   "Tipo de Documento ICMS (Tab 4.1.1)"
      End
      Begin VB.Menu mnTab412 
         Caption         =   "Tabela Situação Documento (Tab 4.1.2)"
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mn411_Click()
    frmTipoDocIcms.Show 0
End Sub

Private Sub mnAtivoImob_Click()
    frmAtivoImob.Show 1
End Sub

Private Sub mnCodTotRedZ_Click()
    frmCodTotRedZ.Show 1
End Sub

Private Sub mnContabilista_Click()
    frmContabilista.Show 1
End Sub

Private Sub mnDados_Click()
    frmEmpresa.Show 1
End Sub

Private Sub mnGeneros_Click()
    frmGeneroLista.Show 1
End Sub

Private Sub mnGerarSintegra_Click()
    frmSped.Show 1
End Sub

Private Sub mnParticipantes_Click()
    frmParticipantes.Show 0
End Sub

Private Sub mnSair_Click()
    P_Fim
End Sub

Private Sub mnTab412_Click()
    frmSituacaoDoc.Show 1
End Sub

Private Sub mnUnidades_Click()
    '
End Sub
