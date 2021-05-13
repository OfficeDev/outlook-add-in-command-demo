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
# Complemento de Outlook de demostración de comando de complemento

El complemento de demostración de comando de complemento usa el modelo de comandos de los complementos de Outlook para agregar botones a la cinta de opciones.

## Requisitos previos

Para ejecutar este ejemplo, necesitará lo siguiente:

- Un servidor web para hospedar los archivos de ejemplo. El servidor debe poder aceptar solicitudes protegidas por SSL (https) y tener un certificado SSL válido.
- Una cuenta de correo electrónico de Office 365 **o** una cuenta de correo electrónico de Outlook.com.
- Outlook 2016, que forma parte de la [versión preliminar de Office 2016](https://products.office.com/en-us/office-2016-preview).

## Configurar e instalar el ejemplo

1. Descargue o bifurque el repositorio.
1. Copie los archivos de complemento en un servidor web. Tiene dos opciones:
    1. Cargarlos manualmente a un servidor:
        1. Cargue los directorios `AllPropsView`, `Assets`, `FunctionFile`, `InsertTextPane`, `NoCommands` y `RestCaller` en un directorio de servidor web.
        1. Abra `command-demo-manifest.xml` en un editor de texto. Reemplace todas las instancias de `https://localhost:8443` por la URL HTTPS del directorio donde se encuentran los archivos cargados en el paso anterior. Guarde los cambios.
    1. Use `Gulp-WebServer` (requiere NPM):
        1. Abra un símbolo del sistema en el directorio donde se encuentra instalado el archivo `package.json` y ejecute `npm install`.
        1. Ejecute `gulp serve-static` para iniciar un servidor web en el directorio actual.
        1. Para que Outlook pueda cargar el complemento, el certificado SSL usado por `gulp-webserver` debe ser de confianza. Abra el explorador y vaya a `https://localhost:8443/AllPropsView/AllProps.html`. Si se le indica que "existe un problema con el certificado de seguridad de este sitio web" (Internet Explorer o Microsoft Edge) o que "el certificado de seguridad del sitio no es de confianza" (Chrome), tendrá que agregar el certificado a las entidades de certificación raíz de confianza. Si va a la página en el explorador, la mayoría de los exploradores le permitirán ver e instalar el certificado. Una vez que haya instalado el certificado y reiniciado el explorador, debería poder examinar `https://localhost:8443/AllPropsView/AllProps.html` sin errores.
1. Inicie sesión en su cuenta de correo electrónico con un explorador en https://outlook.office365.com (para Office 365) o https://www.outlook.com (para Outlook.com). Haga clic en el icono de engranaje de la esquina superior derecha.

    - Si hay un elemento de menú llamado **Administrar integraciones**, siga estos pasos:
        1. Haga clic en **Administrar integraciones**.

            ![Elemento de menú Administrar integraciones en https://www.outlook.com](./readme-images/outlook-manage-integrations.PNG)

        1. Haga clic en el texto **Haga clic aquí para agregar un complemento personalizado** y, después, elija **Agregar desde archivo...**.

            ![Menú de complemento personalizado en https://www.outlook.com](./readme-images/integrations-add-from-file.PNG)

        1. Vaya al archivo `command-demo-manifest.xml` en el equipo de desarrollo. Haga clic en **Abrir**.

        1. Revise la advertencia y haga clic en **Instalar**.

    - Si no hay un elemento de menú llamado **Administrar integraciones**, siga estos pasos:
        1. Haga clic en **Opciones**.
            
            ![Menú Opciones en https://www.outlook.com](./readme-images/outlook-manage-addins.PNG)

        1. En el panel de navegación de la parte izquierda, expanda **General** y, después, haga clic en **Administrar complementos**.
            
        1. En la lista de complementos, haga clic en el icono **+** y elija **Agregar desde un archivo**.

            ![Menú Agregar desde archivo de la lista de complementos](./readme-images/addin-list.PNG)

        1. Haga clic en **Examinar** y vaya al archivo `command-demo-manifest.xml` en el equipo de desarrollo. Haga clic en **Siguiente**.

            ![Diálogo Agregar complemento desde archivo](./readme-images/browse-manifest.PNG)

        1. En la pantalla de confirmación, verá una advertencia que indica que el complemento no es de la Tienda Office y no está comprobado por Microsoft. Haga clic en **Instalar**.
        1. Debería aparecer un mensaje que indica que todo ha ido bien: **Ha agregado un complemento para Outlook**. Haga clic en Aceptar.

## Ejecutar el ejemplo ##

1. Abra Outlook 2016 y conéctese a la cuenta de correo electrónico en la que instaló el complemento.
1. Abra un mensaje existente (ya sea en el panel de lectura o en una ventana independiente). Observe que el complemento ha colocado nuevos botones en la cinta de comandos.
  
  ![Botones del complemento en un formulario de lectura de correo de Outlook](./readme-images/read-mail.PNG)
  
1. Cree un nuevo correo electrónico. Observe que el complemento ha colocado nuevos botones en la cinta de comandos.

  ![Botones del complemento en un formulario de nuevo correo de Outlook](./readme-images/new-mail.PNG)

## Componentes clave del ejemplo

- [```command-demo-manifest.xml```](command-demo-manifest.xml): archivo de manifiesto para el complemento.
- [```FunctionFile/Functions.html```](FunctionFile/Functions.html): archivo HTML vacío para cargar `Functions.js` para los clientes que admiten los comandos de complemento.
- [```FunctionFile/Functions.js```](FunctionFile/Functions.js): el código que se llama cuando se hace clic en los botones del comando de complemento.
- [```InsertTextPane/InsertText.html```](InsertTextPane/InsertText.html): el formato HTML de la interfaz de usuario del panel de tareas que se muestra mediante el botón **Insertar mensaje personalizado**.
- [```InsertTextPane/InsertText.js```](InsertTextPane/InsertText.js): el código que usa la interfaz de usuario del panel de tareas que se muestra mediante el botón **Insertar mensaje personalizado**.
- [```AllPropsView/AllProps.html```](AllPropsView/AllProps.html): el formato HTML de la interfaz de usuario del panel de tareas que se muestra mediante el botón **Mostrar todas las propiedades**. También se muestra en modo lectura en los clientes que no admiten los comandos de complemento.
- [```AllPropsView/AllProps.js```](AllPropsView/AllProps.js): el código que usa la interfaz de usuario del panel de tareas que se muestra mediante el botón **Mostrar todas las propiedades**.
- [```NoCommands/NoCommands.html```](NoCommands/NoCommands.html): archivo HTML que los clientes que no admiten los comandos de complemento cargan y muestran en el modo de redacción.
- [```NoCommands/NoCommands.js```](NoCommands/NoCommands.js): código que los clientes que no admiten los comandos de complemento llaman en el modo de redacción.

## ¿Cómo funciona?

La parte esencial del ejemplo es la estructura del archivo de manifiesto. El manifiesto usa el mismo esquema de la versión 1.1 que el manifiesto de cualquier complemento de Office. Sin embargo, hay una nueva sección del manifiesto llamada `VersionOverrides`. En esta sección se incluye toda la información que los clientes que admiten los comandos de complemento necesitan para llamar al complemento desde un botón de la cinta de opciones. Al colocar esto en una sección completamente independiente, el manifiesto también puede incluir el formato original para habilitar el complemento para que lo carguen los clientes que no admiten el modelo de comandos de complemento. Puede cargar el complemento en Outlook 2013 o en Outlook en la Web para ver cómo funciona.

### El complemento de demostración de comandos de complemento cargado en Outlook en la Web ###

#### Formulario de lectura de correo ####

![Formulario de lectura de correo del complemento cargado en Outlook en la Web](./readme-images/outlook-on-web-read.PNG)

#### Formulario de redacción de correo ####

![Formulario de redacción de correo del complemento cargado en Outlook en la Web](./readme-images/outlook-on-web.PNG)

En el elemento `VersionOverrides`, hay tres elementos secundarios: `Requirements`, `Resources` y `Hosts`. El elemento `Requirements` especifica la versión de API mínima que necesita el complemento cuando lo cargan los clientes que no admiten el modelo de complementos. El elemento `Resources` contiene información sobre iconos y cadenas, así como sobre el archivo HTML que hay que cargar para el complemento. La sección `Hosts` especifica cómo y cuándo se carga el complemento.

En este ejemplo, solo se especifica un host (Outlook):

```xml
<Host xsi:type="MailHost">
```
    
En este elemento, se especifican las opciones de configuración para la versión de escritorio de Outlook:

```xml
<DesktopFormFactor>
```
    
La URL del archivo HTML que contiene todo el código JavaScript para el botón se especifica en el elemento `FunctionFile` (tenga en cuenta que usa el Id. de recurso que se especifica en el elemento `Resources`):

```xml
<FunctionFile resid="functionFile" />
```

El manifiesto especifica los cuatro puntos de extensión disponibles:

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
