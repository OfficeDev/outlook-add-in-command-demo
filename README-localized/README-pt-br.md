---
page_type: sample
products:
- office-outlook
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 9/10/2015 1:04:20 PM
---
# Suplemento do Outlook para demonstração de comando de suplemento

O suplemento de Demonstração de Comando de Suplemento usa o modelo de comandos para suplementos do Outlook para adicionar botões à faixa de opções.

## Pré-requisitos

Para executar este exemplo, você precisará do seguinte:

- Um servidor Web para hospedar os arquivos de exemplo. O servidor deve ser capaz de aceitar solicitações protegidas por SSL (https) e ter um certificado SSL válido.
- Uma conta de e-mail do Office 365 **ou** uma conta de e-mail do Outlook.com.
- O Outlook 2016, que faz parte do[Office 2016 Preview](https://products.office.com/en-us/office-2016-preview).

## Configurar e instalar o exemplo

1. Baixar ou bifurcar o repositório.
1. Copie os arquivos do suplemento para um servidor da Web. Você tem algumas opções:
    1. Carregar manualmente em um servidor:
        1. Carregue o `AllPropsView`, `Assets`, `Functionfile`, `InsertTextPane`, `,`e os diretórios`RestCaller` a um diretório no servidor da Web.
        1. Abra `command-demo-manifest.xml` em um editor de texto. Substitua todas as instâncias de `https://localhost:8443` com a URL HTTPS do diretório em que você carregou os arquivos na etapa anterior. Salve suas alterações.
    1. Use `gulp-webserver` (exige o NPM):
        1. Abra o seu aviso de comando no diretório onde o arquivo de `pacote.json` está instalado e execute`instalar o npm`.
        1. Execute `gulp serve-static` para iniciar um servidor Web na pasta atual.
        1. Para que o Outlook carregue o suplemento, o certificado SSL usado pelo `gulp-webserver` deve ser confiável. Abra o navegador e vá para `https://localhost:8443/AllPropsView/AllProps.html`. Se você for avisado de que "há um problema com o certificado de segurança do site" (IE ou Edge) ou que "o certificado de segurança do site não é confiável" (Chrome), você precisará adicionar o certificado às suas autoridades de certificação raiz confiáveis. Se você continuar a página no navegador, a maioria dos navegadores permite que você exiba e instale o certificado. Depois de instalar e reiniciar seu navegador, você poderá navegar em `https://localhost:8443/AllPropsView/AllProps.html` sem erros.
1. Faça logon na sua conta de e-mail com um navegador no https://outlook.office365.com (para Office 365) ou https://www.outlook.com (para Outlook.com). Clique no ícone de engrenagem no canto superior direito.

    - Se houver um item de menu denominado **Gerenciar integrações**, siga estas etapas:
        1. Clique em **Gerenciar integrações**.

            ![O item de menu Gerenciar integrações no https://www.outlook.com](./readme-images/outlook-manage-integrations.PNG)

        1. Clique no texto **Clique aqui para adicionar um suplemento personalizado**, em seguida, escolha **Adicionar do arquivo...**.

            ![O menu de suplemento personalizado no https://www.outlook.com](./readme-images/integrations-add-from-file.PNG)

        1. Navegue até o arquivo `command-demo-manifest.xml` em seu computador de desenvolvimento. Clique em **Abrir**.

        1. Examine o aviso e clique em **Instalar**.

    - Se não houver um item de menu denominado **Gerenciar integrações**, siga estas etapas:
        1. Clique em **Opções**.
            
            ![O item de menu opções em https://www.outlook.com](./readme-images/outlook-manage-addins.PNG)

        1. Na navegação à esquerda, expanda **Geral** e, em seguida, clique em **Gerenciar suplementos**.
            
        1. Na lista de suplementos, clique no ícone de **+** e escolha **Adicionar de um arquivo**.

            ![O item de menu Adicionar do arquivo na lista de suplementos](./readme-images/addin-list.PNG)

        1. Clique em **Procurar** e navegue até o arquivo `command-demo-manifest.xml` em sua máquina de desenvolvimento. Clique em **Avançar**.

            ![Caixa de diálogo Adicionar suplemento de um arquivo](./readme-images/browse-manifest.PNG)

        1. Na tela de confirmação, você verá um aviso informando que o suplemento não é da Office Store e ainda não foi verificado pela Microsoft. Clique em **Instalar**.
        1. Você deverá ver uma mensagem de sucesso: **Você adicionou um suplemento do Outlook**. Clique em OK.

## Execução da amostra ##

1. Abra o Outlook 2016 e conecte-se à conta de e-mail na qual você instalou o suplemento.
1. Abra uma mensagem existente (no painel de leitura ou em uma janela separada). Observe que o suplemento colocou novos botões na faixa de opções de comando.
  
  ![Os botões de suplemento em um formulário de leitura de e-mail no Outlook](./readme-images/read-mail.PNG)
  
1. Crie um novo e-mail. Observe que o suplemento colocou novos botões na faixa de opções de comando.

  ![Os botões de suplemento em um novo formulário de e-mail no Outlook](./readme-images/new-mail.PNG)

## Componentes principais do exemplo

- [```command-demo-manifest.xml```](command-demo-manifest.xml): O arquivo de manifesto para o suplemento.
- [```FunctionFile/Functions.html```](FunctionFile/Functions.html): Um arquivo HTML vazio para carregar `Funções.js` para clientes que oferecem suporte aos comandos de suplemento.
- [```FunctionFile/Functions.js```](FunctionFile/Functions.js): O código que é chamado quando os botões de comando do suplemento são clicados.
- [```InsertTextPane/InsertText.html```](InsertTextPane/InsertText.html): A marcação HTML para a interface do usuário no painel de tarefas exibida pelo botão **inserir mensagem personalizada**.
- [```InsertTextPane/InsertText.js```](InsertTextPane/InsertText.js): Código usado pela IU do painel de tarefas exibida pelo botão **Inserir mensagem personalizada**.
- [```AllPropsView/AllProps.html```](AllPropsView/AllProps.html): A marcação HTML para a interface do usuário do painel de tarefas exibida pelo botão **Exibir todas as propriedades**. Isso também é exibido por clientes em modo de leitura que não têm suporte para comandos de suplemento.
- [```AllPropsView/AllProps.js```](AllPropsView/AllProps.js): Código usado pelo painel de tarefas IU exibida pelo botão **Exibir todas as propriedades**.
- [```NoCommands/NoCommands.html```](NoCommands/NoCommands.html): O arquivo HTML que é carregado e exibido por clientes em modo de redação não têm suporte para comandos de suplemento.
- [```nocommands/nocommands. js```](NoCommands/NoCommands.js): O código invocado por clientes em modo de redação que não têm suporte para comandos de suplemento.

## Como isso funciona?

A parte fundamental da amostra é a estrutura do arquivo de manifesto. O manifesto usa o mesmo esquema de versão 1.1 que qualquer suplemento do Office. No entanto, há uma nova seção do manifesto chamada `VersionOverrides`. Esta seção contém todas as informações que os clientes que oferecem suporte aos comandos do suplemento precisam para invocar o suplemento de um botão da faixa de opções. Colocando isso em uma seção completamente separada, o manifesto também pode incluir a marcação original para habilitar o suplemento a ser carregado por clientes que não têm suporte para o modelo de comando do suplemento. Você pode ver isso em ação carregando o suplemento no Outlook 2013 ou no Outlook na Web.

### O Suplemento de Demonstração de Comando do suplemento carregado no Outlook na Web ###

#### Ler formulários de e-mails ####

![O suplemento carregado no Outlook no formulário de leitura de e-mails da Web](./readme-images/outlook-on-web-read.PNG)

#### Formulário Redigir e-mail ####

![O suplemento carregado no Outlook no formulário Redigir e-mails da Web](./readme-images/outlook-on-web.PNG)

No elemento `VersionOverrides`, há três elementos filho, `Requisitos`, `Recursos`e `Hosts`. O elemento `Requisitos` especifica a versão mínima da API exigida pelo suplemento quando carregado por clientes que oferecem suporte ao modelo de suplemento. O elemento `Recursos` contém informações sobre ícones, cadeias de caracteres e qual arquivo HTML carregar para o suplemento. A seção `Hosts` especifica como e quando o suplemento é carregado.

Neste exemplo, há apenas um host especificado (Outlook):

```xml
<Host xsi:type="MailHost">
```
    
Dentro desse elemento estão as especificações de configuração para a versão de área de trabalho do Outlook:

```xml
<DesktopFormFactor>
```
    
A URL para o arquivo HTML com todo o código JavaScript para o botão é especificado no elemento `Functionfile`, (observe que ele usa a ID do recurso especificado no elemento `Recursos`):

```xml
<FunctionFile resid="functionFile" />
```

O manifesto especifica todos os quatro pontos de extensão disponíveis:

```xml
<!-- Message compose form -->
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
<!-- Appointment compose form -->
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
<!-- Message read form -->
<ExtensionPoint xsi:type="MessageReadCommandSurface">
<!-- Appointment read form -->
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
```
    
Within each extension point, there is an example of each type of button.

### A button that executes a function ###

This is created by setting the `xsi:type` attribute of a `Control` element to `Button`, and adding an `Action` child element with an `xsi:type` attribute set to `ExecuteFunction`. For example, look at the **Insert default message** button:

```xml
<!-- Function (UI-less) button -->
<Control xsi:type="Button" id="msgComposeFunctionButton">
  <Label resid="funcComposeButtonLabel" />
  <Supertip>
    <Title resid="funcComposeSuperTipTitle" />
    <Description resid="funcComposeSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>addDefaultMsgToBody</FunctionName>
  </Action>
</Control>
```
    
### A drop-down menu button ###

This is created by setting the `xsi:type` attribute of a `Control` element to `Menu`, and adding an `Items` child element that contains the items to appear on the menu. For example, look at the **Insert message** button:

```xml
<!-- Menu (dropdown) button -->
<Control xsi:type="Menu" id="msgComposeMenuButton">
  <Label resid="menuComposeButtonLabel" />
  <Supertip>
    <Title resid="menuComposeSuperTipTitle" />
    <Description resid="menuComposeSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="red-icon-16" />
    <bt:Image size="32" resid="red-icon-32" />
    <bt:Image size="80" resid="red-icon-80" />
  </Icon>
  <Items>
    <Item id="msgComposeMenuItem1">
      <Label resid="menuItem1ComposeLabel" />
      <Supertip>
        <Title resid="menuItem1ComposeLabel" />
        <Description resid="menuItem1ComposeTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>addMsg1ToBody</FunctionName>
      </Action>
    </Item>
    <Item id="msgComposeMenuItem2">
      <Label resid="menuItem2ComposeLabel" />
      <Supertip>
        <Title resid="menuItem2ComposeLabel" />
        <Description resid="menuItem2ComposeTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>addMsg2ToBody</FunctionName>
      </Action>
    </Item>
    <Item id="msgComposeMenuItem3">
      <Label resid="menuItem3ComposeLabel" />
      <Supertip>
        <Title resid="menuItem3ComposeLabel" />
        <Description resid="menuItem3ComposeTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>addMsg3ToBody</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```
    
### A button that opens a task pane ###

This is created by setting the `xsi:type` attribute of a `Control` element to `Button`, and adding an `Action` child element with an `xsi:type` attribute set to `ShowTaskPane`. For example, look at the **Insert custom message** button:

```xml
<!-- Task pane button -->
<Control xsi:type="Button" id="msgComposeOpenPaneButton">
  <Label resid="paneComposeButtonLabel" />
  <Supertip>
    <Title resid="paneComposeSuperTipTitle" />
    <Description resid="paneComposeSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="composeTaskPaneUrl" />
  </Action>
</Control>
```

## Questions and comments

- If you have any trouble running this sample, please [log an issue](https://github.com/jasonjoh/command-demo/issues).
- Questions about Office Add-in development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Make sure that your questions or comments are tagged with `office-addins`.

## Additional resources

- [Outlook Dev Center](https://developer.microsoft.com/en-us/outlook/)
- [Office Add-ins](https://msdn.microsoft.com/library/office/jj220060.aspx) documentation on MSDN
- [More Add-in samples](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)

## Copyright

Copyright (c) 2015 Microsoft. All rights reserved.


----------
Connect with me on Twitter [@JasonJohMSFT](https://twitter.com/JasonJohMSFT)

Follow the [Outlook Dev Blog](https://blogs.msdn.microsoft.com/exchangedev)


This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
