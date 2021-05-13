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
# アドイン コマンド デモの Outlook アドイン

アドイン コマンド デモのアドインでは、Outlook アドインのコマンド モデルを使用してリボンにボタンを追加します。

## 前提条件

このサンプルを実行するには、次のものが必要です。

- サンプル ファイルをホストする Web サーバー。サーバーは、SSL で保護された要求 (https) を受け入れることが可能で、有効な SSL 証明書を所有している必要があります。
- Office 365 メール アカウント **または** Outlook.com メール アカウント。
- Office 2016 の一部である [Outlook 2016 プレビュー](https://products.office.com/en-us/office-2016-preview)。

## サンプルの構成とインストール

1. レポジトリをダウンロードまたはフォークします。
1. アドイン ファイルを Web サーバーにコピーします。2 つのオプションがあります。
    1. 手動でサーバーにアップロードする:
        1. `AllPropsView`、`Assets`、`FunctionFile`、`InsertTextPane`、`NoCommands`、および `RestCaller` の各ディレクトリを Web サーバーのディレクトリにアップロードします。
        1. テキスト エディターで [`command-demo-manifest.xml`] を開きます。`https://localhost:8443` のすべてのインスタンスを、前の手順でファイルをアップロードしたディレクトリの HTTPS URL で置き換えます。変更内容を保存します。
    1. `gulp-webserver` を使用する (NPM が必要です):
        1. [`package.json`] ファイルがインストールされているディレクトリでコマンド プロンプトを開き、`npm install` を実行します。
        1. `gulp serve-static` を実行して、現在のディレクトリで Web サーバーを開始します。
        1. Outlook がアドインを読み込むには、`gulp-webserver` で使用する SSL 証明書 が信頼される必要があります。ブラウザーを開いて、`https://localhost:8443/AllPropsView/AllProps.html` に移動します。"この Web サイトのセキュリティ証明書には問題があります" (IE または Microsoft Edge) または "このサイトのセキュリティ証明書は信頼できません" (Chrome)　と表示される場合は、証明書を信頼されたルート証明機関に追加する必要があります。ブラウザーでページに進んだ場合、ほとんどのブラウザーでは証明書を表示してインストールできます。インストールを行ってブラウザーを再起動すると、エラーが表示されることなく `https://localhost:8443/AllPropsView/AllProps.html` に移動できるはずです。
1. ブラウザーを使用して、https://outlook.office365.com (Office 365 用) または https://www.outlook.com (Outlook.com 用) のいずれかでメール アカウントにログオンします。右上隅にある歯車アイコンをクリックします。

    - [**統合の管理**] というメニュー項目が表示されている場合は、次の手順を実行します。
        1. [**統合の管理**] をクリックします。

            ![https://www.outlook.com の [統合の管理] メニュー項目](./readme-images/outlook-manage-integrations.PNG)

        1. [**カスタム アドインを追加するには、ここをクリックします**] というテキストをクリックし、次に [**ファイルから追加...**] を選択します。

            ![https://www.outlook.com のカスタム アドイン メニュー](./readme-images/integrations-add-from-file.PNG)

        1. 展開用コンピューター上の [`command-demo-manifest.xml`] ファイルを参照します。[**開く**] をクリックします。

        1. 警告を確認し、[**インストール**] をクリックします。

    - [**統合の管理**] というメニュー項目が表示されていない場合は、次の手順を実行します。
        1. [**オプション**] をクリックします。
            
            ![https://www.outlook.com のオプション メニュー](./readme-images/outlook-manage-addins.PNG)

        1. 左側のナビゲーションで、[**全般**] を展開し、[**アドインの管理**] をクリックします。
            
        1. アドインの一覧で [**+**] アイコンをクリックし、[**ファイルから追加**] を選択します。

            ![アドインの一覧の [ファイルから追加] メニュー項目](./readme-images/addin-list.PNG)

        1. [**参照**] をクリックし、展開用コンピューター上の [`command-demo-manifest.xml`] ファイルを参照します。[**次へ**] をクリックします。

            ![[ファイル からアドインを追加] ダイアログ](./readme-images/browse-manifest.PNG)

        1. アドインが Office Store からのものではなく、Microsoft により確認されていないという警告が確認画面に表示されます。[**インストール**] をクリックします。
        1. 次の成功メッセージが表示されます:**Outlook 用のアドインを追加しました**。[OK] をクリックします。

## サンプルの実行 ##

1. Outlook 2016 を開き、アドインをインストールしたメール アカウントに接続します。
1. (閲覧ウィンドウまたは別のウィンドウのいずれかに) 既存のメッセージを 1 つ開きます。アドインにより新しいボタンがコマンド リボンに配置されたことを確認してください。
  
  ![Outlook のメール閲覧形式のアドイン ボタン](./readme-images/read-mail.PNG)
  
1. 新しいメール メッセージを作成します。アドインにより新しいボタンがコマンド リボンに配置されたことを確認してください。

  ![Outlook の新しいメール形式の [アドイン] ボタン](./readme-images/new-mail.PNG)

## サンプルの主要なコンポーネント

- [```command-demo-manifest.xml```](command-demo-manifest.xml):アドイン用のマニフェスト ファイル。
- [```FunctionFile/Functions.html```](FunctionFile/Functions.html):アドイン コマンドをサポートするクライアント用に `Functions.js` を読み込むための空の HTML ファイル。
- [```FunctionFile/Functions.js```](FunctionFile/Functions.js):アドイン コマンド ボタンがクリックされたときに呼び出されるコード。
- [```InsertTextPane/InsertText.html```](InsertTextPane/InsertText.html):[**カスタム メッセージを挿入**] ボタンにより表示される作業ウィンドウ UI の HTML マークアップ。
- [```InsertTextPane/InsertText.js```](InsertTextPane/InsertText.js):[**カスタム メッセージを挿入**] ボタンにより表示される作業ウィンドウ UI が使用するコード。
- [```AllPropsView/AllProps.html```](AllPropsView/AllProps.html):[**すべてのプロパティを表示**] ボタンにより表示される作業ウィンドウ UI の HTML マークアップ。これは、アドイン コマンドをサポートしていないクライアントの閲覧モードでも表示されます。
- [```AllPropsView/AllProps.js```](AllPropsView/AllProps.js):[**すべてのプロパティを表示**] ボタンにより表示される作業ウィンドウ UI が使用するコード。
- [```NoCommands/NoCommands.html```](NoCommands/NoCommands.html):アドイン コマンドをサポートしていないクライアントにより新規作成モードで読み込まれて表示される HTML ファイル。
- [```NoCommands/NoCommands.js```](NoCommands/NoCommands.js):アドイン コマンドをサポートしていないクライアントにより新規作成モードで呼び出されるコード。

## 動作の仕組み

このサンプルの重要な部分は、マニフェスト ファイルの構造です。Office アドインのすべてのマニフェストと同様、このマニフェストではバージョン 1.1 のスキーマが使用されています。ただし、このマニフェストには `VersionOverrides` という新しいセクションがあります。このセクションには、アドイン コマンドをサポートしているクライアントがリボン ボタンからアドインを呼び出すのに必要なすべての情報が含まれています。この情報を完全に別のセクションに含めることにより、アドイン コマンド モデルをサポートしていないクライアントがアドインを読み込むことを可能にする元のマークアップをマニフェストに含めることもできます。アドインを Outlook 2013 または Outlook on the web に読み込むと、この動作を確認することができます。

### Outlook on the web に読み込まれたアドイン コマンド デモ アドイン ###

#### メール閲覧形式 ####

![Outlook on the web のメール閲覧形式に読み込まれたアドイン](./readme-images/outlook-on-web-read.PNG)

#### メールの新規作成形式 ####

![Outlook on the web のメールの新規作成形式に読み込まれたアドイン](./readme-images/outlook-on-web.PNG)

`VersionOverrides` 要素内には、`Requirements`、`Resources`、`Hosts` という 3 つの子要素があります。`Requirements` 要素は、アドイン モデルをサポートするクライアントにより読み込まれる際にアドインで要求される最小 API バージョンを指定します。`Resources` 要素には、アドイン用のアイコン、文字列、および読み込む HTML ファイルに関する情報が含まれています。`Hosts` セクションは、アドインが読み込まれる方法とタイミングを指定します。

このサンプルでは、ホストは 1 つだけ指定されています (Outlook)。

```xml
<Host xsi:type="MailHost">
```
    
この要素内には、デスクトップ版 Outlook の構成の詳細が含まれています。

```xml
<DesktopFormFactor>
```
    
ボタン用のすべての JavaScript コードを含む HTML ファイルへの URL は、`FunctionFile` 要素で指定されています (`Resources` 要素が指定するリソース ID が使用される点に注意してください):

```xml
<FunctionFile resid="functionFile" />
```

マニフェストでは、利用可能な 4 つすべての拡張点が指定されます。

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
