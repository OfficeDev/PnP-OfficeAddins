---
page_type: sample
products:
- office-outlook
- office
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 04/02/2021 10:00:00 AM
description: "Use Outlook event-based activation to set the signature."
---

# Use Outlook event-based activation to set the signature

**Applies to:** Outlook on Windows | Outlook on the web

## Summary

This sample uses event-based activation to run an Outlook add-in when the user creates a new message or appointment. The add-in can respond to events, even when the task pane is not open. It also uses the [setSignatureAsync API](https://docs.microsoft.com/javascript/api/outlook/office.body?view=outlook-js-preview#setSignatureAsync_data__options__callback_). If no signature is set, the add-in prompts the user to set a signature, and can then open the task pane for the user.

For documentation related to this sample, see [Configure your Outlook add-in for event-based activation](https://docs.microsoft.com/office/dev/add-ins/outlook/autolaunch)

## Features

- Use event-based activation to respond to events when the task pane is not open.
- Set a signature for Outlook to use in messages and appointments.

## Applies to

- Outlook on Windows, and on the web.

## Prerequisites

- Before running this sample, you need a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org/en/) installed on your computer. To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.

## Solution

| Solution | Author(s) |
|---------|----------|
| Use Outlook event-based activation to set the signature | Microsoft |

## Version history

Version  | Date | Comments
|---------|------|---------|
1.0 | 4-01-2021 | Initial release
1.1 | 6-1-2021 | Update for GA of setSignature API

## Scenario: Event-based activation

In this scenario, the add-in helps the user manage their email signature, even when the task pane is not open. When the user sends a new email message, or creates a new appointment, the add-in displays an information bar prompting the user to create a signature. If the user chooses to set a signature, the add-in opens the task pane for the user to continue setting their signature.

## Build and run the solution

1. Clone or download this repository.
2. In the command line, go to the **outlook-set-signature** folder from your root directory.
3. Run the following command to download the dependencies required to run the sample.
    
    ```command&nbsp;line
    $ npm install
    ```
4. Run the following command to start the localhost web server.
    
    ```command&nbsp;line
    $ npm run dev-server
    ```

5. Sideload the add-in to Outlook on Windows, or Outlook on the web by following the manual instructions in the article [Sideload Outlook add-ins for testing](https://docs.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).

Once the add-in is loaded use the following steps to try out the functionality.

1. Open Outlook on Windows or in a browser.
2. Create a new message or appointment.
    
    You should see a notification at the top of the message that reads: **Please set your signature with the PnP sample add-in.**
    
3. Choose **Set signatures**. This will open the task pane for the add-in.
4. In the task pane fill out the fields for your signature data. Then choose **Save**.
5. The task pane will load a page of sample templates. You can assign the templates to a **New Mail**, **Reply**, or **Forward** action. Once you've assign the templates you want to use, choose **Save**.

The next time you create a message or appointment, you'll see the signature you selected applied by the add-in.

## Key parts of this sample

### Configure event-based activation in the manifest

The manifest configures a runtime that is loaded specifically to handle event-based activation. The following `<Runtime>` element specifies an HTML page resource id that loads the runtime on Outlook on the web. The `<Override>` element specifies the JavaScript file instead, to load the runtime for Outlook on Windows. Outlook on Windows doesn't use the HTML page to load the runtime.

```xml
<Runtime resid="Autorun">
  <Override type="javascript" resid="runtimeJs"/>
...
<bt:Url id="Autorun" DefaultValue="https://localhost:3000/src/runtime/HTML/autorunweb.html"></bt:Url>
<bt:Url id="runtimeJs" DefaultValue="https://localhost:3000/src/runtime/Js/autorunshared.js"></bt:Url>
```

The add-in handles two events that are mapped to the `checkSignature()` function.

```xml
<LaunchEvents>
  <LaunchEvent Type="OnNewMessageCompose" FunctionName="checkSignature" />
  <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="checkSignature" />
</LaunchEvents>
```

### Handling the events and using the setSignatureAsync API

When the user creates a new message or appointment, Outlook will load the files specified in the manifest to handle the `OnNewMessageCompose` and `OnNewAppointmentOrganizer` events. Outlook on the web will load the `autorunweb.html` page, which then also loads `autorunweb.js` and `autorunshared.js`.

The `autorunweb.js` file contains a version of the `insert_auto_signature` function used specifically when running on Outlook on the web. The [setSignatureAsync() API cannot be used in Outlook on the web for appointments](https://docs.microsoft.com/javascript/api/outlook/office.body?view=outlook-js-preview#setSignatureAsync_data__options__callback_). Therefore, `insert_auto_signature` inserts the signature into a new appointment by directly writing to the body text of the appointment.

The `autorunshared.js` file contains the `checkSignature` function that handles the events from Outlook. It also contains additional code that is shared and loaded when the add-in is used in Outlook on the web and Outlook on Windows. On Outlook on Windows, this file is loaded directly and `autorunweb.html` and `autorunweb.js` are not loaded.

The `autorunshared.js` file contains a version of the `insert_auto_signature` function that uses the `setSignatureAsync()` API to set the signature for both messages and appointments.

Note that you can use a similar pattern when handling events. If you need code that only applies to Outlook on the web, you can load it in a separate file like `autorunweb.js`. And for code that applies to both Outlook on the web and Outlook on Windows, you can load it in a shared file like `autorunshared.js`.

### Embedding images with the signature

Template A shows how to insert an image by embedding it in the signature. This will avoid the image being downloaded from your server when the signature is inserted into new mail items. The HTML uses the following `<img>` tag format with the **src** set to **cid:*imageFileName*** to embed the image.

```xml
str +=
    "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='cid:" +
    logoFileName +
    "' alt='MS Logo' width='24' height='24' /></td>";
```

In the **addTemplateSignature** function, if template A is used, it will attach the image by calling the **addFileAttachmentFromBase64Async()** API. Then it calls the **setSignatureAsync()** API.

### Referencing images from the signature

Template B shows how to reference an image from the HTML. It uses the `<img>` tag and references the web location.

```xml
 str +=
    "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://localhost:3000/assets/sample-logo.png' alt='Logo' /></td>";
```

This is a simpler approach as you don't need to attach the image. Although your web server will need to provide the image anytime Outlook needs it for a signature.

### Task pane code

The task pane code is located under the `taskpane` folder of this project. The task pane HTML and JavaScript files only provide UI and functionality to let the user specify and save a signature.

- `editsignature.html` is loaded when the task pane first opens. It lets the user enter details such as name and title for their signature.
- `assignsignature.html` is loaded when the user saves their details from the `editsignature.html` page. It lets the user assign the signature to actions such as "new email", "reply", and "forward.

## Security notes

In the webpack.config.js file, a header is set to `"Access-Control-Allow-Origin": "*"`. This is only for development purposes. In production code, you should list the allowed domains and not leave this header open to all domains.

You'll be prompted to install certificates for trusted access to https://localhost. The certificates are intended only for running and studying this code sample. Do not reuse them in your own code solutions or in production environments.

Install or uninstall the certificates by running the following commands in the project folder.

```
npx office-addin-dev-certs install
npx office-addin-dev-certs uninstall
```

## Copyright

Copyright (c) 2021 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/samples/outlook-autorun-set-signature" />
