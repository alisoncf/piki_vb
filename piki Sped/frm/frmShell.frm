VERSION 5.00
Begin VB.Form frmShell 
   Caption         =   "Shell"
   ClientHeight    =   765
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11865
   Icon            =   "frmShell.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   765
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtComando 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtComando_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        p_shell
    End If
End Sub
Sub p_shell()
    Dim c As New Collection
    
    Set c = F_SeparaTexto(txtComando.Text, , False)
    
    If LCase(c(1)) = "sair" Then
        Unload Me
        Exit Sub
    End If
    
    If LCase(c(1)) = "bkp" Or LCase(c(1)) = "backup" Then
        frmBackup.Show 1
        Exit Sub
    End If
    
    
    
    If c(1) <> "do" Then
        MsgBox "Comando inválido"
        Exit Sub
    End If
    If LCase(c(2)) = "mudar_icms" Or LCase(c(2)) = "alterar_icms" Then
        mudar_icms
    End If
    
    If LCase(c(2)) = "baixafin" Or LCase(c(2)) = "baixacartao" Then
        frmBaixaFinanceirasLote.Show 1
    End If
    
    
    If LCase(c(2)) = "migrar" Then
        frmMigrarEstoque.Show 1
    End If
    If LCase(c(2)) = "vendas" Then
        frmEdiVendas.Show 1
    End If
    If LCase(c(2)) = "ajuste" And strNome_Usu = "ALISON" Then
        frmAjustes.Show 1
    End If
    If LCase(c(2)) = "parametros" And strNome_Usu = "ALISON" Then
        frmParametrosRegras.Show 1
        
    End If


End Sub
Sub mudar_icms()
    'do mudar_icms(empresa=01, cfop_revenda=5102, picms=19
    
    
    

End Sub
