# API
Agora, nós precisaremos transformar a nossa tabela do Google Planilhas em um sistema que seja capaz de receber uma requisição e retornar o status referente ao id da planilha do usuário. Caso o status seja "inativo", a planilha irá fechar-se automaticamente.

## 1. Criando script da API
1. Acesse a sua planilha do Google Planilhas, clique em extensões e selecione "Apps Script". Se o site solicitar alguma autorização, conceda.

    ![](https://imgur.com/mavOCin.jpg)
2. Ao entrar no painel do Google Script, você verá que existe apenas uma função vazia. Basta excluir ela e inserir o código abaixo.
> [!CAUTION]
> Para usar o código abaixo como está, é necessário que a aba da sua planilha vinculada ao seu formulário esteja com o nome "database".

```js
function doGet(req) {
  var id = req.parameter.id;
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getSheetByName('database');
  var values = sheet.getDataRange().getValues();

  var output = [];
  for(var i = 1; i < values.length; i++) {
    var row = {};
    row['Id'] = values[i][1];
    row['status'] = values[i][2];
    row['pc-name'] = values[i][3];
    output.push(row);
  }

  if(id != null) {
    var outputRetorno = output.filter(obj => obj.Id.includes(id));
    return ContentService.createTextOutput(JSON.stringify(outputRetorno[0])).setMimeType(ContentService.MimeType.JSON);
  } else {
    return null
  } 
}
```
## 2. Implementando API:
1. Clique em botão "Implementar" no topo da página.
2. Na janela, clique na engrenagem ao lado de "Selecionar o tipo" e selecione "App da Web".

    ![](https://imgur.com/iRkwx5A.jpg)
3. Na página de configuração, garanta que "Executar como" esteja como "Eu" e "Quem pode acessar" como "Qualquer pessoa" 

    ![](https://imgur.com/gEFZvpk.jpg)
4. Será solicitado que você conceda acesso ao script, basta seguir as orientações descritas. Caso você se depare com a tela de que o site não é seguro, clicar em "Hide Advanced" e em "Go to".

    ![](https://imgur.com/5dN1DJW.jpg)
5. Copie o link gerado, pois nós o utilizaremos na próxima etapa.

## 3. Integrando API com o VBA:
Agora nós vamos trabalhar para fazer a requisição ao Google Script e lidar com a resposta do servidor dentro da planilha do Excel.

### JSON
O primeiro passo é adicionar um código no VBA para que ele possa lidar com a resposta do servidor. Para isso, vamos precisar importar um módulo para dentro do nosso projeto VBA. Na pasta "Arquivo" deste repositório [(clique aqui)](/Arquivos/JsonConverter.bas), baixe o arquivo "JsonConverter.bas".

> Esse módulo foi criado por [Tim Hall](https://github.com/VBA-tools/VBA-JSON)

### Adicionando Script
No nosso projeto no editor VBA, entre no módulo "mdl_sistema" e adicione a rotina abaixo.

```vbnet
Sub consultarAPI()

'Delacarando as variáveis
Dim requisicao As New WinHttpRequest
Dim resposta As Object
Dim id As String
Dim pc_name As String
Dim link As String

'Coloque aqui o link do Google Script
link = ""

'Consultando ID e Nome do PC armazenado
id = ThisWorkbook.Sheets("cadastro").Range("B1").Value
pc_name = ThisWorkbook.Sheets("cadastro").Range("B2").Value

'Fazendo a requisição na API
On Error GoTo TratamentoConexao
requisicao.Open "Get", link & "?id=" & id
requisicao.Send
On Error GoTo 0

    If requisicao.status <> 200 Then
        MsgBox "Erro de conexão! Por favor, verfifique se o seu dispositivo está conectado a internet.", vbCritical, "Acesso Negado"
        Exit Sub
    End If

    Set resposta = JsonConverter.ParseJson(requisicao.responseText)

    'Conferindo Status

    If resposta("status") = "inativo" Then
        MsgBox "Sua sessão não foi validada!", vbCritical, "Acesso Negado"
        ThisWorkbook.Close SaveChanges:=False
    End If

    'Conferindo PC NAME
    If resposta("pc-name") = pc_name Then
        Exit Sub
    Else
        MsgBox "Dipositivo não autorizado!", vbCritical, "Acesso Negado"
        ThisWorkbook.Close SaveChanges:=False
    End If
    

Exit Sub

TratamentoConexao:
    MsgBox "Erro de conexão! Por favor, verfifique se o seu dispositivo está conectado a internet.", vbCritical, "Acesso Negado"
    Exit Sub

TratamentoJson:
    MsgBox "Erro crítico (ERRO-JSON)", vbCritical
    Exit Sub

End Sub
```

Agora que você seguiu todos os passos, siga para o próximo capítulo do tutorial.
[Próximo Capítulo](/5-%20Finalização.md)
