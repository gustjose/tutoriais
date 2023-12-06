# Preparando a Planilha
A princípio, precisamos preparar o arquivo para receber o nosso código. Já que o nosso projeto é baseado em VBA, poderemos escolher qualquer programa do Pacote Office para utilizá-lo e, por escolha própria, decidi utilizar o **Excel**.

## 1. Ativando as Opções de Desenvolvedor:
Após criar o arquivo da sua planilha, entre no excel para fazer algumas configurações necessárias. As opções abaixo podem mudar de acordo com a versão do pacote office utilizado.

1. Procure a opção "Arquivo" no canto esquerdo superior:

    ![](https://imgur.com/z8joyr7.jpg)
2. Após abrir o menu lateral, entre em Opções (última opção da lista).
3. Aguarde a abetura da janela, selecione a opção "Personalizer Faixa de Opções" e, por fim, marque a caixa que diz "Desenvolvedor":

    ![](https://imgur.com/EtP0zuc.jpg)
4. Marque "OK" e verifique se a aba "Desenvolver" apareceu no menu da planilha.

## 2. Ativando as bibliotecas necessárias:
Considerando que o modo desenvolver está ativo no seu Excel, utilize o atalho **Alt + F11** no seu teclado para entrar na janela de desenvolvimento do VBA.

1. No menu superiror, localize a opção "Ferramentas" e, dentro do sub-menu, selecione "Referências":

    ![](https://imgur.com/BmTV4fP.jpg)
2. Na lista de referências, busque e ative as bibliotecas abaixo:

    * Microsoft Scripting Runtime
    * Microsoft WinHTTP Services

    ![](https://imgur.com/O30YwH6.jpg)

## 3. Escolhendo os dados solicitados na Ativação:
Está é a hora de você decidir quais dados o usuário precisará fornecer e quais informações o software vai capturar durante a ativação online do arquivo. Eu escolhi as seguintes:
* Nome completo;
* E-mail;
* Nome do Computador;
* 4 primeiros dígitos do CPF.

Lembrando que a escolha dos dados fica a sua escolha, podendo ser personalizada da maneira que for preciso.

## 4. Criando a janela de ativação:
Agora você precisa criar um formulário dentro do VBA para que o usuário o preencha com seus dados. Criei um template para facilitar o nosso tutorial, você o encontra na pasta "Arquivos" com o nome "frm_ativar.frm" e "frm_ativar.frx" ou [clicando aqui](/Arquivos/frm_ativar.frm) (Atenção! Os dois arquivos devem ser baixados juntos para que o formulário seja importado corretamente). Fique a vontade para adaptá-lo de acordo com as necessidades do seu projeto.

![](https://imgur.com/bYjS354.jpg)

> [!NOTE] 
> Não há caixa de texto perguntando o nome do computador, já que essa informação será capturada diretamente pelo script. Você deve fazer isso para quaisquer outras informações que o usuário não terá acesso.

Lembre-se que o nome das caixas de texto e dos botões fazem diferença no momento de criar o script.

### Como importar o seu modelo para a minha planilha?
1. Baixe o template.
2. No VBA, clique com o botão direito do mouse na pagina que contém os arquivos do projeto. Busque a opção "Importar arquivo".

    ![](https://imgur.com/6G5ctF7.jpg)

3. Vai abrir uma janela para que você possa selecionar o arquivo baixado, basta seguir as instruções.


Agora que você seguiu todos os passos, siga para o próximo capítulo do tutorial.
[Próximo Capítulo](/2-%20Database%20online.md)

