# Ativação da Planilha
Agora nós desenvolveremos o script que conectará o Formulário do Google e o nosso arquivo do Excel. Quando o usuário abrir a Planilha, a página de ativação que desenvolvemos será mostrada e ele deverá preencher com seus dados pessoais. Ao final do processo, quando ele clicar em "Enviar", a resposta será automaticamente registrada na tabela do Google Planilha.

## 1. Criando o Módulo:
O primeiro passo é criar o módulo no Excel que carregará todo o script que vamos desenvolver. 
1. Para isso, abra sua planilha, abra o editor de códigos VBA e adicione um módulo em branco. Por fim, o renomei para "mdl_sistema".

    ![](https://imgur.com/pGRXRt4.jpg)

## 2. Armazenando dados na Planilha:
Eu recomendo que você crie uma aba em sua planilha do Excel para armazenar os dados de cadastro. Também é possível armazeá-los de outras formas como no gerenciador de nomes, mas assim elas poderão ser acessadas por usuários que possuírem maior conhecimento. Aproveite e renomeia a aba para "cadastro".

![](https://imgur.com/ZTIgk3U.jpg)

> [!NOTE] 
> Na coluna "A" eu escrevi os itens que serão armazenados apenas para melhor organização, enquanto a coluna "B" ficará reservada para que o script adicione automaticamente as informações.

## 3. Ajustando URL:
Lembra-se daquela URL do formulário que eu pedi para você salvar? Pegue ela, pois precisaremos fazer algumas modificações para que tudo funcione corretamente.

>https://docs.google.com/forms/d/e/1FAIpQLScuUP0fJ5J2CBxp3CihenzL0y0PXyhTo0A8mwEiOD9BYoe4Rg/viewform?usp=pp_url&entry.1328736906=123456&entry.421279956=123456&entry.1412787280=123456&entry.578858415=123456&entry.893614914=123456&entry.465661081=123456

1. Procure com atenção no texto da URL a expressão "viewform" e a altere para "formResponse". 

<p style="text-decoration: none;">https://docs.google.com/forms/d/e/1FAIpQLScuUP0fJ5J2CBxp3CihenzL0y0PXyhTo0A8mwEiOD9BYoe4Rg/<span style="color: green">formResponse</span>?usp=pp_url&entry.1328736906=123456&entry.421279956=123456&entry.1412787280=123456&entry.578858415=123456&entry.893614914=123456&entry.465661081=123456</p>

2. Agora é uma parte mais complicada. Vamos ter que converter o link para o formato que usaremos no VBA. 

    1. Coloque o link entre aspas;
        ```
        "https://docs.google.com/forms/d/e/1FAIpQLScuUP0fJ5J2CBxp3CihenzL0y0PXyhTo0A8mwEiOD9BYoe4Rg/formResponse?usp=pp_url&entry.1328736906=123456&entry.421279956=123456&entry.1412787280=123456&entry.578858415=123456&entry.893614914=123456&entry.465661081=123456"
        ```
    2. Identidique todos os "123456" que foram usados, já que eles correspondem aos itens na ordem colocada no Google Formulário. Basta você substituir o "123456" por: __" & id & "__ (Incluindo as aspas e para todos os itens: id, pc_name, email, nome, cpf).

    <p style="text-decoration: none;">https://docs.google.com/forms/d/e/1FAIpQLScuUP0fJ5J2CBxp3CihenzL0y0PXyhTo0A8mwEiOD9BYoe4Rg/<span style="color: green">formResponse</span>?usp=pp_url&entry.1328736906=<span style="color: red">123456</span>&entry.421279956=<span style="color: red">123456</span>&entry.1412787280=<span style="color: red">123456</span>&entry.578858415=<span style="color: red">123456</span>&entry.893614914=<span style="color: red">123456</span>&entry.465661081=<span style="color: red">123456</span></p>

    Aqui está como deve ficar no final:
  
        "https://docs.google.com/forms/d/e/1FAIpQLScuUP0fJ5J2CBxp3CihenzL0y0PXyhTo0A8mwEiOD9BYoe4Rg/formResponse?usp=pp_url&entry.1328736906=" & id & "&entry.421279956=" & status & "&entry.1412787280=" & pc-name & "&entry.578858415=" & nome & "&entry.893614914=" & email & "&entry.465661081=" & cpf

3. Por fim, basta adicionar __& "&submit=Submit"__ ao final. Ficando com esse texto:
    ```
        "https://docs.google.com/forms/d/e/1FAIpQLScuUP0fJ5J2CBxp3CihenzL0y0PXyhTo0A8mwEiOD9BYoe4Rg/formResponse?usp=pp_url&entry.1328736906=" & id & "&entry.421279956=" & status & "&entry.1412787280=" & pc_name & "&entry.578858415=" & nome & "&entry.893614914=" & email & "&entry.465661081=" & cpf & "&submit=Submit"
    ```

## 4. Script de Cadastro:
Vamos voltar ao módulo para criar uma rotina que será responsável por fazer uma requisição no link e cadastrar a planilha no sistema. Também criaremos uma rotina para gerar um ID alfanumérico aleatório automaticamente.

> [!CAUTION]
> Para usar o código abaixo como está, é necessário que o formulário de cadastro no Excel esteja nomeado como "frm_ativar" e as caixas de texto estejam respectivamente com os nomes: "txt_nome", "txt_email" e "txt_cpf". Além disso, a aba da planilha que armazenará o cadastro no Excel deve ter o nome "cadastro".

```vbnet
'Para impedir o usuário acesse as funções desse módulo
Option Private Module

'Função para gerar o ID aleatório
Function GerarCodigoAlfanumerico() As String
    Dim caracteres As String
    Dim codigo As String
    Dim i As Integer
    
    ' Definir os caracteres permitidos para o código
    caracteres = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    
    ' Inicializar o gerador de números aleatórios
    Randomize
    
    ' Gerar o código alfanumérico de 8 caracteres
    For i = 1 To 15
        codigo = codigo & Mid(caracteres, Int((Len(caracteres) * Rnd) + 1), 1)
    Next i
    
    ' Retornar o código gerado
    GerarCodigoAlfanumerico = codigo
End Function

'Função para enviar a requisão ao google formulários
Sub enviarDados()

'Declarando as variáveis
Dim id As String
Dim status As String
Dim pc_name As String
Dim nome As String
Dim email As String
Dim cpf As String
Dim chave As String
Dim link As String

'Gerar id
chave = GerarCodigoAlfanumerico()

'Pegando as respostas do usuário e convertando para URL
id = WorksheetFunction.EncodeURL(chave)
status = WorksheetFunction.EncodeURL("ativo")
pc_name = WorksheetFunction.EncodeURL(Environ("COMPUTERNAME"))
nome = WorksheetFunction.EncodeURL(frm_ativar.txt_nome)
email = WorksheetFunction.EncodeURL(frm_ativar.txt_email)
cpf = WorksheetFunction.EncodeURL(frm_ativar.txt_cpf)

'Aqui você deve colocar o link que foi ajustado
link = 

'Salvando dados na aba de cadastro 
Thisworkbook.Sheets("cadastro").Range("B1").Value = id
Thisworkbook.Sheets("cadastro").Range("B2").Value = pc_name
Thisworkbook.Sheets("cadastro").Range("B3").Value = nome
Thisworkbook.Sheets("cadastro").Range("B4").Value = email
Thisworkbook.Sheets("cadastro").Range("B5").Value = cpf

On Error GoTo TratamentoConexão
'Enviando requisição ao link
Set http = CreateObject("MSXML2.ServerXMLHTTP")
    http.Open "Get", link, False
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.Send
    

Application.ThisWorkbook.Save
MsgBox "Planilha ativada com sucesso!", vbInformation
Unload frm_ativar
Exit Sub
    
TratamentoConexão:
    MsgBox "Error! Por favor, confira se você está conectado a Internet e tente novamente.", vbCritical, "Acesso Negado"
    Unload form_validation
    ThisWorkbook.Close SaveChanges:=False

End Sub
```

## 5. Adicionando script ao formulário de cadastro:
Agora basta adicionar a chamada da função ao botão de ativar do formulário de cadastro da planilha.
1. No editor VBA, entre no formulário de cadastro e clique duas vezes no botão "Ativar" para a abrir o evento "click".
2. Adicione a instrução "Call enviarDados".

Deve ficar assim:
```vbnet
Private Sub btn_ativar_Click()
    Call enviarDados
End Sub
```

Agora que você seguiu todos os passos, siga para o próximo capítulo do tutorial.
[Próximo Capítulo](/4-%20API.md)