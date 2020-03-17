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
# 加载项命令演示 Outlook 加载项

加载项命令演示加载项将命令模型用于 Outlook 加载项以便向功能区添加按钮。

## 先决条件

若要运行此示例，需要具备以下条件：

- 用于托管示例文件的 Web 服务器。服务器必须能够接受受 SSL 保护的请求 (https)，并且具备有效的 SSL 证书。
- Office 365 电子邮件帐户**或** Outlook.com 电子邮件帐户。
- Outlook 2016（[Office 2016 预览版的一部分](https://products.office.com/en-us/office-2016-preview)）。

## 配置和安装示例

1. 下载或为存储库创建分支。
1. 将加载项文件复制到 Web 服务器。你拥有几个选项：
    1. 手动上传到服务器：
        1. 将 `AllPropsView`、`Assets`、`FunctionFile`、`InsertTextPane`、`NoCommands` 和 `RestCaller` 目录上传到 Web 服务器上的目录。
        1. 在文本编辑器中打开 `command-demo-manifest.xml`。将 `https://localhost:8443` 的所有实例替换为在上一步中上传的文件所在目录的 HTTPS URL。保存所做的更改。
    1. 使用 `gulp-webserver`（需要 NPM）：
        1. 在安装了 `package.json` 文件的目录中打开命令提示符，然后运行 `npm install`。
        1. 运行 `gulp serve-static` 以在当前目录中启动 Web 服务器。
        1. 为了让 Outlook 加载加载项，必须信任 `gulp-webserver` 使用的 SSL 证书。打开浏览器并转到 `https://localhost:8443/AllPropsView/AllProps.html`。如果系统提示“此网站的安全证书有问题”（IE 或Edge），或“该网站的安全证书不受信任”(Chrome)，则需要将该证书添加到受信任的根证书颁发机构。如果你继续在浏览器中浏览页面，则大多数浏览器都允许你查看和安装证书。安装并重启浏览器后，你应该能够浏览 `https://localhost:8443/AllPropsView/AllProps.html`，而无任何错误。
1. 使用浏览器在 https://outlook.office365.com（对于 Office 365）或 https://www.outlook.com（对于 Outlook.com）上登录电子邮件帐户。单击右上角的齿轮图标。

    - 如果存在名为“**管理集成**”的菜单项，请按照以下步骤操作：
        1. 单击“**管理集成**”。

            ![https://www.outlook.com 上的“管理集成”菜单项](./readme-images/outlook-manage-integrations.PNG)

        1. 单击文本“**单击此处添加自定义外接程序**”，然后选择“**从文件添加...**”。

            ![https://www.outlook.com 上的自定义外接程序菜单](./readme-images/integrations-add-from-file.PNG)

        1. 浏览到开发计算机上的 `command-demo-manifest.xml` 文件。单击“**打开**”。

        1. 查看警告，然后单击“**安装**”。

    - 如果不存在名为“**管理集成**”的菜单项，请按照以下步骤操作：
        1. 单击“**选项**”。
            
            ![https://www.outlook.com 上的“选项”菜单项](./readme-images/outlook-manage-addins.PNG)

        1. 在左侧导航中，展开“**常规**”，然后单击“**管理加载项**”。
            
        1. 在加载项列表中，单击 **+** 图标并选择“**从文件添加**”。

            ![加载项列表中的“从文件添加”菜单项](./readme-images/addin-list.PNG)

        1. 单击“**浏览**”并浏览到开发计算机上的 `command-demo-manifest.xml` 文件。单击“**下一步**”。

            ![“从文件添加加载项”对话框](./readme-images/browse-manifest.PNG)

        1. 在确认屏幕上，你将看到一条警告，指出该加载项不是来自 Office 应用商店，并且尚未得到 Microsoft 的验证。单击“**安装**”。
        1. 应看到一条成功消息：**你已添加一个 Outlook 相关加载项**。单击“确定”。

## 运行示例 ##

1. 打开 Outlook 2016 并连接到安装了该加载项的电子邮件帐户。
1. 打开现有邮件（在阅读窗格中或在单独的窗口中）。请注意，该加载项已在命令功能区上放置了新按钮。
  
  ![Outlook 中的“读取邮件”窗体上的加载项按钮。](./readme-images/read-mail.PNG)
  
1. 新建一封电子邮件。请注意，该加载项已在命令功能区上放置了新按钮。

  ![Outlook 中的“新建邮件”窗体上的加载项按钮](./readme-images/new-mail.PNG)

## 示例的主要组件

- [```command-demo-manifest.xml```](command-demo-manifest.xml)：加载项的清单文件。
- [```FunctionFile/Functions.html```](FunctionFile/Functions.html)：空 HTML 文件，用于为支持加载项命令的客户端加载 `Functions.js`。
- [```FunctionFile/Functions.js```](FunctionFile/Functions.js)：单击加载项命令按钮时调用的代码。
- [```InsertTextPane/InsertText.html```](InsertTextPane/InsertText.html)：“**插入自定义邮件**”按钮显示的任务窗格 UI 的 HTML 标记。
- [```InsertTextPane/InsertText.js```](InsertTextPane/InsertText.js)：“**插入自定义邮件**”按钮显示的任务窗格 UI 使用的代码。
- [```AllPropsView/AllProps.html```](AllPropsView/AllProps.html)：“**显示所有属性**”按钮显示的任务窗格 UI 的 HTML 标记。客户端在不支持加载项命令的读取模式下也会显示它。
- [```AllPropsView/AllProps.js```](AllPropsView/AllProps.js)：“**显示所有属性**”按钮显示的任务窗格 UI 使用的代码。
- [```NoCommands/NoCommands.html```](NoCommands/NoCommands.html)：由客户端在不支持加载项命令的撰写模式中加载和显示的 HTML 文件。
- [```NoCommands/NoCommands.js```](NoCommands/NoCommands.js)：由客户端在不支持加载项命令的撰写模式中调用的代码。

## 它是如何工作的？

该示例的关键部分是清单文件的结构。该清单使用与任何 Office 加载项清单相同的版本 1.1 架构。然而，该清单中有一个称为 `VersionOverrides` 的新部分。此部分包含支持加载项命令的客户端从功能区按钮调用加载项所需的所有信息。通过将其置于完全独立的部分中，该清单还可包含原始标记，以允许不支持加载项命令模型的客户端加载此加载项。你可以通过在 Outlook 2013 或 Outlook 网页版中加载该加载项来了解工作方式。

### Outlook 网页版中加载的“加载项命令演示”加载项 ###

#### “读取邮件”窗体 ####

![Outlook 网页版的“读取邮件”窗体中加载的加载项](./readme-images/outlook-on-web-read.PNG)

#### “撰写邮件”窗体 ####

![Outlook 网页版的“撰写邮件”窗体中加载的加载项](./readme-images/outlook-on-web.PNG)

在 `VersionOverrides` 元素中，有三个子元素，即`要求`、`资源`和`主机`。`要求`元素指定由支持加载项模型的客户端加载此加载项时所需的最低 API 版本。`资源`元素包含有关图标、字符串以及要为加载项加载的 HTML 文件的信息。`主机`部分指定加载加载项的方式和时间。

在此示例中，仅指定了一个主机 (Outlook)：

```xml
<Host xsi:type="MailHost">
```
    
此元素包含 Outlook 桌面版本的配置详细信息：

```xml
<DesktopFormFactor>
```
    
HTML 文件的 URL，该文件包含在 `FunctionFile` 元素（请注意，它使用在`资源`元素中指定的资源 ID）中指定的按钮的所有 JavaScript 代码：

```xml
<FunctionFile resid="functionFile" />
```

该清单指定所有四个可用的扩展点：

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
