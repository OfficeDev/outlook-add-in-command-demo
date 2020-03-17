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
# Пример команд надстройки Outlook

Эта надстройка добавляет кнопки на ленту с помощью модели команд для надстроек Outlook.

## Условия

Для работы с этим образцом необходимо выполнить указанные ниже действия.

- Веб-сервер для размещения примеров файлов. Сервер должен иметь возможность принимать запросы, защищенные SSL (https), и иметь действительный сертификат SSL.
- Учетная запись электронной почты Office 365 **или** учетная запись электронной почты Outlook.com.
- Outlook 2016, который является частью [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview).

## Настройка и установка образца

1. Загрузите или разветвите репозиторий.
1. Скопируйте файлы надстроек на веб-сервер. У вас есть несколько вариантов:
    1. Вручную загрузить на сервер:
        1. Передайте `AllPropsView`, `активы`, `FunctionFile`, `InsertTextPane``NoCommands``RestCaller` каталогов в каталог на веб-сервере.
        1. Откройте `command-demo-manifest.xml` в текстовом редакторе. Замените все экземпляры `https://localhost:8443` URL-адресом HTTPS каталога, в который вы загрузили файлы на предыдущем шаге. Сохраните изменения.
    1. Используйте `gulp-webserver` (требуется NPM):
        1. Откройте командную строку в каталоге, где установлен файл `package.json`, и запустите `npm install`.
        1. Запустите `gulp serve-static`, чтобы запустить веб-сервер в текущем каталоге.
        1. Чтобы Outlook мог загрузить надстройку, сертификат SSL, используемый `gulp-webserver`, должен быть доверенным. Откройте браузер и перейдите по адресу `https://localhost:8443/AllPropsView/AllProps.html`. Если вам будет предложено указать, что «существует проблема с сертификатом безопасности этого веб-сайта» (IE или Edge) или «сертификат безопасности сайта не является доверенным» (Chrome), вам необходимо добавить сертификат в доверенные корневые центры сертификации. Если вы переходите на страницу в браузере, большинство браузеров позволяют просматривать и устанавливать сертификат. После того, как вы установите и перезапустите свой браузер, вы сможете без проблем перейти по адресу`https://localhost:8443/AllPropsView/AllProps.html`.
1. Войдите в свою учетную запись электронной почты с помощью браузера по адресу https://outlook.office365.com (для Office 365) или https://www.outlook.com (для Outlook.com). Нажмите на значок шестеренки в правом верхнем углу.

    - Если есть пункт меню под названием **Управление интеграциями**, выполните следующие действия:
        1. Щелкните **Управление интеграцией**.

            ![Пункт меню Управление интеграциями на https://www.outlook.com](./readme-images/outlook-manage-integrations.PNG)

        1. Нажмите на текст. **Нажмите здесь, чтобы добавить пользовательскую надстройку**, затем выберите **Добавить из файла ...**.

            ![Настраиваемое меню надстроек в https://www.outlook.com](./readme-images/integrations-add-from-file.PNG)

        1. Перейдите к файлу `command-demo-manifest.xml` на компьютере разработчика. Нажмите кнопку **Open** (Открыть).

        1. Просмотрите предупреждение и нажмите **Установить**.

    - Если нет пункта меню под названием **Управление интеграциями**, выполните следующие действия:
        1. Нажмите кнопку **Параметры**.
            
            ![Пункт меню Параметры на https://www.outlook.com](./readme-images/outlook-manage-addins.PNG)

        1. В области навигации слева разверните раздел **общие**, а затем выберите пункт **Управление надстройками**.
            
        1. В списке надстроек щелкните значок **+** и выберите **добавить из файла**.

            ![Пункт "Добавить из файла" в списке надстроек](./readme-images/addin-list.PNG)

        1. Нажмите кнопку **обзор** и перейдите к файлу `command-demo-manifest.xml` на компьютере разработчика. Нажмите **Далее**.

            ![Диалоговое окно "Добавление надстройки из файла"](./readme-images/browse-manifest.PNG)

        1. На экране подтверждения вы увидите предупреждение о том, что надстройка не из Office Store и не была проверена Microsoft. Нажмите **Установить**.
        1. Должно отобразиться сообщение об успешном выполнении: Удаление**вы добавили надстройку для Outlook**. надстройки Outlook Нажмите кнопку ОК.

## Запуск приложения ##

1. Откройте Outlook 2016 и подключитесь к учетной записи электронной почты, в которой вы установили надстройку.
1. Откройте существующее сообщение (в области чтения или в отдельном окне). Обратите внимание, что надстройка поместила новые кнопки на командную ленту.
  
  ![Кнопки надстройки в форме для просмотра почты в Outlook](./readme-images/read-mail.PNG)
  
1. Создайте новое сообщение. Обратите внимание, что надстройка поместила новые кнопки на командную ленту.

  ![Кнопки надстройки на новой почтовой форме в Outlook](./readme-images/new-mail.PNG)

## Ключевые компоненты примера

- [```command-demo-manifest.xml```](command-demo-manifest.xml): Файл манифеста для надстройки.
- [```FunctionFile/Functions.html```](FunctionFile/Functions.html): Пустой HTML-файл для загрузки `функции. js` для клиентов, которые поддерживают команды Add-in.
- [```FunctionFile/Functions.js```](FunctionFile/Functions.js): Код, который вызывается при нажатии кнопок команд надстройки.
- [```InsertTextPane/InsertText.html```](InsertTextPane/InsertText.html): HTML-разметка для пользовательского интерфейса области задач, отображаемая в разделе **Вставка настраиваемого сообщения** ".
- [```InsertTextPane/InsertText.js```](InsertTextPane/InsertText.js): Код, используемый пользовательским интерфейсом области задач, отображаемым кнопкой **Вставка настраиваемого сообщения**
- [```AllPropsView/AllProps.html```](AllPropsView/AllProps.html): HTML-разметка для пользовательского интерфейса области задач, отображаемая кнопкой **Показать все свойства**. Это также отображается клиентами в режиме чтения, которые не поддерживают команды надстроек.
- [```AllPropsView/AllProps.js```](AllPropsView/AllProps.js): Код, используемый в пользовательском интерфейсе области задач, отображаемом кнопкой **Показать все свойства**.
- [```NoCommands/NoCommands.html```](NoCommands/NoCommands.html): Файл HTML, который загружается и отображается клиентами в режиме создания, которые не поддерживают команды надстроек.
- [```NoCommands/NoCommands.js```](NoCommands/NoCommands.js): Код, который вызывается клиентами в режиме компоновки, которые не поддерживают команды надстроек.

## Как все это работает?

Ключевой частью примера является структура файла манифеста. В манифесте используется та же схема версии 1.1, что и в манифесте любой надстройки Office. Однако есть новый раздел манифеста под названием `VersionOverrides`. В этом разделе содержится вся информация, которая необходима клиентам, поддерживающим команды надстройки, для вызова надстройки с помощью кнопки ленты. Поместив это в совершенно отдельный раздел, манифест может также включать исходную разметку, чтобы позволить загрузке надстройки клиентами, которые не поддерживают модель команд надстройки. Вы можете увидеть это в действии, загрузив надстройку в Outlook 2013 или Outlook в Интернете.

### Демонстрационная надстройка «Команда надстройки», загруженная в Outlook в Интернете ###

#### Читать почтовую форму ####

![Надстройка загружена в Outlook в веб-форме для чтения почты](./readme-images/outlook-on-web-read.PNG)

#### Написать письмо ####

![Надстройка загружена в Outlook в веб-форме для создания почты](./readme-images/outlook-on-web.PNG)

В элементе `VersionOverrides` есть три дочерних элемента: `требования`, `ресурсы` и `хосты`. Элемент `Requirements` указывает минимальную версию API, требуемую надстройкой при загрузке клиентами, которые поддерживают модель надстройки. Элемент `Ресурсы` содержит информацию о значках, строках и HTML-файле, который нужно загрузить для надстройки. В разделе `Хосты` указывается, как и когда загружается надстройка.

В этом примере указан только один хост (Outlook):

```xml
<Host xsi:type="MailHost">
```
    
Внутри этого элемента находятся особенности конфигурации для настольной версии Outlook:

```xml
<DesktopFormFactor>
```
    
URL-адрес HTML-файла со всем кодом JavaScript для кнопки указывается в элементе `FunctionFile` (обратите внимание, что он использует идентификатор ресурса, указанный в элементе `Ресурсы`):

```xml
<FunctionFile resid="functionFile" />
```

В манифесте указаны все четыре доступные точки расширения:

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
