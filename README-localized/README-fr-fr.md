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
# Démonstration de la commande du complément Outlook

Le complément de démonstration de la commande du complément utilise le modèle de commandes pour que les compléments Outlook ajoutent des boutons au ruban.

## Conditions préalables

Pour exécuter cet exemple, vous devez disposer des éléments ci-après :

- Un serveur web pour héberger les fichiers de l'exemple. Le serveur doit pouvoir accepter des demandes protégées par SSL (https) et disposer d’un certificat SSL valide.
- Un compte de messagerie Office 365 **ou** un compte de messagerie Outlook.com.
- Outlook 2016, qui fait partie de la [préversion Office 2016](https://products.office.com/en-us/office-2016-preview).

## Configuration et installation de l'exemple

1. Téléchargez ou dérivez par le référentiel.
1. Copiez les fichiers de compléments sur un serveur web. Deux options s’offrent à vous :
    1. Charger manuellement sur un serveur :
        1. Téléchargez les répertoires `AllPropsView`, `Assets`, `FunctionFile`, `InsertTextPane`, `NoCommands`et `RestCaller` vers un annuaire sur votre serveur web.
        1. Ouvrez `command-demo-manifest.xml` dans un éditeur de texte. Remplacez toutes les instances de `https://localhost:8443` par l’URL de HTTP du répertoire dans lequel vous avez téléchargé les fichiers au cours de l’étape précédente. Enregistrez vos modifications.
    1. Utilisez `gulp-webserver` (nécessite NPM) :
        1. Ouvrez votre invite de commande dans le répertoire dans lequel le fichier `package.json` est installé et exécute `npm install`.
        1. Exécutez `gulp serve-static` pour démarrer un serveur web dans le répertoire actif.
        1. Pour qu’Outlook charge le complément, le certificat SSL utilisé par `gulp-webserver` doit être approuvé. Ouvrez votre navigateur et accédez à `https://localhost:8443/AllPropsView/AllProps.html`. Si vous recevez un message indiquant qu'« il existe un problème avec le certificat de sécurité de ce site web » (par ex., Internet Explorer ou Microsoft Edge), ou que « le certificat de sécurité du site n’est pas approuvé » (Chrome), vous devez ajouter le certificat à vos autorités de certification racines de confiance. Si vous continuez vers la page du navigateur, la plupart des navigateurs vous permettent d’afficher et d’installer le certificat. Une fois que vous avez effectué l'installation et redémarré votre navigateur, vous devez être en mesure d’accéder à `https://localhost:8443/AllPropsView/AllProps.html` sans aucune erreur.
1. Connectez-vous à l'aide d'un navigateur à votre compte de messagerie sur https://outlook.office365.com (pour Office 365) ou https://www.outlook.com (pour Outlook.com). Cliquez sur l’icône d’engrenage dans le coin supérieur droit.

    - S'il existe un élément de menu appelé **Gérer des intégrations**, procédez comme suit :
        1. Cliquez sur **Gérer des intégrations**.

            ![L'élément de menu Gérer des intégrations sur https://outlook.com](./readme-images/outlook-manage-integrations.PNG)

        1. Cliquez sur le texte **Cliquez ici pour ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir du fichier....**.

            ![Le menu du complément personnalisé sur https://www.outlook.com](./readme-images/integrations-add-from-file.PNG)

        1. Parcourez le fichier `command-demo-manifest.xml` sur votre ordinateur de développement. Cliquez sur **Ouvrir**.

        1. Lisez le message d'alerte, puis cliquez sur **Installer**.

    - Si l'élément de menu appelé **Gérer des intégrations** est inexistant, procédez comme suit :
        1. Cliquez sur **Options**.
            
            ![L'élément de menu Options sur https://outlook.com](./readme-images/outlook-manage-addins.PNG)

        1. Dans le volet de navigation gauche, développez **Général**, puis cliquez sur **Gérer des compléments**.
            
        1. Dans la liste de compléments, cliquez sur l’icône **+**, puis sélectionnez **Ajouter à partir d’un fichier**.

            ![Élément de menu Ajouter à partir du fichier dans la liste des compléments](./readme-images/addin-list.PNG)

        1. Cliquex sur **Parcourir** et accédez au fichier `command-demo-manifest.xml` sur votre ordinateur de développement. Cliquez sur **Suivant**.

            ![Ajouter un complément à partir de la boîte de dialogue Fichier](./readme-images/browse-manifest.PNG)

        1. L’écran de confirmation affiche un message d'alerte indiquant que le complément ne provient pas du Store Office et n’a pas été vérifié par Microsoft. Cliquez sur **Installer**.
        1. Un message de réussite doit s'afficher : **Vous avez ajouté un complément pour Outlook**. Cliquez sur OK.

## Exécution de l’exemple ##

1. Ouvrez Outlook 2016 et connectez-vous au compte de messagerie sur lequel vous avez installé le complément.
1. Ouvrez un message existant (soit dans le volet de lecture ou dans une fenêtre séparée). Veuillez noter que le complément a installé de nouveaux boutons sur le ruban de commandes.
  
  ![Boutons de compléments sur un formulaire de courrier lu dans Outlook](./readme-images/read-mail.PNG)
  
1. Créez un courrier électronique. Veuillez noter que le complément a installé de nouveaux boutons sur le ruban de commandes.

  ![Boutons de compléments sur un nouveau formulaire de courrier dans Outlook](./readme-images/new-mail.PNG)

## Principaux composants de l’exemple

- [```command-demo-manifest.xml```](command-demo-manifest.xml) : Fichier manifeste pour le complément.
- [```FunctionFile/Functions.html```](FunctionFile/Functions.html) : Fichier HTML vide pour charger `functions.js` pour les clients prenant en charge les commandes de complément.
- [```FunctionFile/Functions.js```](FunctionFile/Functions.js) : Code appelé lorsqu'un utilisateur clique sur les boutons de commande du complément.
- [```InsertTextPane/InsertText.html```](InsertTextPane/InsertText.html) : Balise HTML pour l’interface utilisateur du volet de tâches affichée par le bouton **Insérer un message personnalisé**.
- [```InsertTextPane/InsertText.js```](InsertTextPane/InsertText.js) : Code utilisé par l’interface utilisateur du volet de tâches affichée par le bouton **Insérer un message personnalisé**.
- [```AllPropsView/AllProps.html```](AllPropsView/AllProps.html) : Balise HTML pour l’interface utilisateur du volet de tâches affichée par le bouton **Afficher toutes les propriétés**. Elle est également affichée par les clients en mode lecture qui ne prennent pas en charge les commandes de complément.
- [```AllPropsView/AllProps.js```](AllPropsView/AllProps.js) : Code utilisé pour l’interface utilisateur du volet de tâches affichée par le bouton **Afficher toutes les propriétés**.
- [```NoCommands/NoCommands.html```](NoCommands/NoCommands.html) : Fichier HTML chargé et affiché par les clients en mode composition qui ne prennent pas en charge les commandes de complément.
- [```NoCommands/NoCommands.js```](NoCommands/NoCommands.js) : Code qui est appelé par les clients en mode composition qui ne prennent pas en charge les commandes de complément.

## Comment fonctionne tout cela ?

L'élément clé de l’exemple est la structure du fichier manifeste. Le manifeste utilise le même schéma de version 1.1 que tout manifeste de complément Office. Cependant, une nouvelle section du manifeste appelée `VersionOverrides` existe. Cette section comprend toutes les informations que les clients prenant en charge les commandes de complément doivent suivre pour appeler le complément à partir d’un bouton du ruban. En plaçant celui-ci dans une section totalement distincte, le manifeste peut également inclure la balise d’origine permettant le chargement du complément par des clients qui ne prennent pas en charge le modèle de commande du complément. Vous pouvez afficher ceci en action en chargeant le complément dans Outlook 2013 ou Outlook sur le web.

### Le complément de démo de la commande de complément est chargé dans Outlook sur le web ###

#### Lecture du formulaire courrier ####

![Le complément chargé dans le formulaire de lecture de courrier d'Outlook sur le web](./readme-images/outlook-on-web-read.PNG)

#### Composer un formulaire de courrier ####

![Le complément chargé dans le formulaire de rédaction de courrier d'Outlook sur le web](./readme-images/outlook-on-web.PNG)

Dans l'élément `VersionOverrides`, il existe trois éléments enfants : `Configuration requise`, `Ressources` et `Hôtes`. L'élément `Configuration requise` spécifie la version minimale d'API exigée par le complément lors du chargement par des clients prenant en charge le modèle de complément. L’élément `Ressources` se compose des informations sur les icônes, les chaînes et le fichier HTML à charger pour le complément. La section `Hôtes` spécifie la modalité et la période de chargement du complément.

Dans cet exemple, un seul hôte est mentionné (Outlook) :

```xml
<Host xsi:type="MailHost">
```
    
Les détails de la configuration de la version de bureau d’Outlook sont inclus dans cet élément :

```xml
<DesktopFormFactor>
```
    
L’URL du fichier HTML contenant la totalité du code JavaScript pour le bouton est spécifiée dans l'élément `FunctionFile` (veuillez noter qu’il utilise l’ID de ressource spécifié dans l’élément `Ressources`) :

```xml
<FunctionFile resid="functionFile" />
```

Le manifeste précise les quatre points d’extension disponibles :

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
