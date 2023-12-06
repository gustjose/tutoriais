VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ativar 
   Caption         =   "Ativar Planilha"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6810
   OleObjectBlob   =   "frm_ativar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_ativar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_ativar_Click()
    Call enviarDados
End Sub

Private Sub btn_cancelar_Click()
    ThisWorkbook.Close SaveChanges:=False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        ' Exibir uma MsgBox para confirmar o fechamento
        Dim resposta As VbMsgBoxResult
        resposta = MsgBox("Tem certeza que deseja fechar a planilha?", vbQuestion + vbYesNo, "Fechar Planilha")
        
        ' Verificar a resposta do usuário
        If resposta = vbNo Then
            Cancel = True ' Cancelar o fechamento
        Else
            ThisWorkbook.Close SaveChanges:=False
        End If
    End If
End Sub
