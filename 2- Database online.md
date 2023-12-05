# Database on-line
Chegou a hora de criarmos a database on-line do nossa planilha, onde ficarão os dados de ativação e onde poderemos desativar uma planilha caso haja alguma violação. Já que queremos fazer um sistema totalmente gratuíto, utilizaremos serviços do Google.

> [!IMPORTANT] 
> Esse tipo de sistema não é adequado para projetos que possuem um grande volume de dados, pois existem limites quanto a transferência de informações.

## 1. Construíndo o database:
1. Entre no link https://docs.google.com/forms/u/0/ e logue com sua conta do Google.
2. Clique em "Criar Formulário em Branco".
3. Altere o nome do formulário (para melhor organização) e comece a adicionar os itens abaixo utilizando a modalidade "Resposta Curta":
    * ID;
    * STATUS;
    * PC-NAME;
    * NOME;
    * EMAIL;
    * CPF.
4. Na guia respostas, clique em "Link para o app Google Planilha". Aqui você poderá criar uma nova planilha para armazenar os dados ou vincular o formulário a uma planilha já existente.
5. Na planilha gerada pelo Google Planilhas, altere o nome da aba para "database" (esse passo é importante para que o código funcione corretamente).

    ![](https://imgur.com/75tEM4p.jpg)

## 2. Configurando as repostas
O Google Formulário será usado como meio para ativar novas planilhas e, para isso, é necessário gerar um link com uma resposta padrão que será utilizada pelo script VBA.
1. No menu três pontinhos, selecione a opção "Gerar link preenchido automaticamente".
2. Na aba que foi aberta, preencha todos os campos com "123456" e clique em "Gerar Link" no final da página.
3. Ao final do processo você ter um link parecido com esse: 
>https://docs.google.com/forms/d/e/1FAIpQLScuUP0fJ5J2CBxp3CihenzL0y0PXyhTo0A8mwEiOD9BYoe4Rg/viewform?usp=pp_url&entry.1328736906=123456&entry.421279956=123456&entry.1412787280=123456&entry.578858415=123456&entry.893614914=123456&entry.465661081=123456
4. Salve-o em algum lugar, pois o usaremos no futuro.

Agora que você seguiu todos os passos, siga para o próximo capítulo do tutorial.
[Próximo Capítulo](/3-%20Ativação%20da%20Planilha.md)