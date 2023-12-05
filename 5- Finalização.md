# Finalização
Por fim, basta fazer alguns pequenos ajustes de modo que a verificação on-line seja feita sempre quando a planilha for aberta, e outros pequenos detalhes.

## Formulário de Ativação
Abra o formulário de ativação dentro do editor VBA, clique com botão direito dentro do form e selecione "Exibir código".

### Impedindo o Fechamento durante a ativação
Cole o código abaixo:

```vbnet
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
```

### Atribuindo a função ao botão "Cancelar"
Cole o código abaixo:
```vbnet
Private Sub btn_cancelar_Click()
    ThisWorkbook.Close SaveChanges:=False
End Sub
```

## Ativando a verificação on-line
> [!WARNING]  
> Salve e faça backup dos arquivos antes de prosseguir para esta etapa.

Para enfim ativar a verificação on-line, basta inserir o código abaixo no evento "Open" da pasta de trabalho.

```vbnet
Private Sub Workbook_Open()
    If ThisWorkbook.Sheets("cadastro").Range("B1").Value = Empty Then
        frm_ativar.Show
    Else
        Call consultarAPI
    End If
End Sub
```
Deve ficar assim:

![](https://imgur.com/rAVznkE.jpg)

## Ocultando acesso a aba "cadastro"

Para que nenhum usuário tenha acesso as informações vitais do sistema, é necessário ocultar a aba "cadastro" dentro do VBA.
1. Selecione a aba "cadastro"
2. Na janela de propriedades, no atributo "Visible", escolha a opção "2- XlSheetVeryHidden"

![](https://imgur.com/hNApGO1.jpg)

Pronto, chegamos ao final do projeto. Espero que essas informações tenham alguma serventia para o seu trabalho.