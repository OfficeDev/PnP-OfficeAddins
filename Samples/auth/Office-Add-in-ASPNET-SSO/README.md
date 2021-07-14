---
page_type: sample
products:
- office-excel
- office-powerpoint
- office-word
- microsoft-365
languages:
- javascript
- aspx
extensions:
  contentType: samples
  technologies:
  - Add-ins
  - Microsoft Graph
  services:
  - Excel
  - Microsoft 365
  createdDate: 5/1/2017 2:09:09 PM
---
# Office Add-in that supports Single Sign-on to Office, the add-in, and Microsoft Graph

This sample implements an Office Add-in that uses the `getAccessToken` API in Office.js to give the add-in access to Microsoft Graph data. This sample is built on ASP.NET and Microsoft Authentication Library (MSAL) .NET.

There are two versions of the sample in this repo, one of which has its own README file:

- **Before** folder. The starting point for the SSO walkthrough at [Create an ASP.NET Office Add-in that uses single sign-on](https://docs.microsoft.com/office/dev/add-ins/develop/create-sso-office-add-ins-aspnet). Please follow the instructions in the article.
- **Complete** folder. The completed sample. To use this version, follow the instructions in the article [Create an ASP.NET Office Add-in that uses single sign-on](https://docs.microsoft.com/office/dev/add-ins/develop/create-sso-office-add-ins-aspnet), but substitute "Complete" for "Before" in those instructions and skip the sections **Code the client-side** and **Code the server-side**.

## Features

Integrating data from online service providers increases the value and adoption of your add-ins. This code sample shows you how to connect your add-in to Microsoft Graph. Use this code sample to:

* See how to use the Single Sign-on (SSO) API
* Connect to Microsoft Graph from an Office Add-in
* Build an add-in using ASP.NET MVC, MSAL 4.x.x for .NET, and Office.js 
* Use the MSAL.NET Library to implement the OAuth 2.0 authorization framework in an add-in
* Use the OneDrive REST APIs from Microsoft Graph
* See how an add-in can fall back to an interactive sign-in in scenarios where SSO is not supported
* Show a dialog using the Office UI namespace in scenarios where SSO is not supported
* Use add-in commands

## Applies to

-  Any platform and Office host combination that supports the [IdentityAPI 1.3 requirement set](https://docs.microsoft.com/en-us/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).

## Prerequisites

To run this code sample, the following are required:

* Visual Studio 2019 or later
* [Office Developer Tools](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)
* A Microsoft 365 account. If you don't have one, join the [Microsoft 365 developer program](https://aka.ms/devprogramsignup) which includes a free 1 year subscription to Microsoft 365
* At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription
* A Microsoft Azure Tenant. This add-in requires Azure Active Directory (AD). Azure AD provides identity services that applications use for authentication and authorization. Sign up for a trial subscription at [Microsoft Azure Signup](https://account.windowsazure.com/SignUp)

## Solution

Solution | Author(s)
---------|----------
Office Add-in Microsoft Graph ASP.NET | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | May 10, 2017| Initial release
1.0 | September 15, 2017 | Added support for 2FA.
1.0 | December 8, 2017 | Added extensive error handling.
1.0 | January 7, 2019 | Added information about web application security practices.
2.0 | November 5, 2019 | Added Display Dialog API fall back and use new version of SSO API.
2.1 | August 11, 2020 | Removed preview note because the APIs have released.
2.2 | June 15, 2021 | Updated NuGet packages and adjust code for breaking changes.


## Security note

The sample sends a hardcoded query parameter on the URL for the Microsoft Graph REST API. If you modify this code in a production add-in and any part of query parameter comes from user input, be sure that it is sanitized so that it cannot be used in a Response header injection attack.

## Questions and comments

We'd love to get your feedback about this sample. You can send your feedback to us in the *Issues* section of this repository.
Questions about developing Office Add-ins should be posted to [Microsoft Q&A](https://docs.microsoft.com/en-us/answers/topics/office-js-dev.html).

## Join the Microsoft 365 Developer Program
Get a free sandbox, tools, and other resources you need to build solutions for the Microsoft 365 platform.
- [Free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Get a free, renewable 90-day Microsoft 365 E5 developer subscription.
- [Sample data packs](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Automatically configure your sandbox by installing user data and content to help you build your solutions.
- [Access to experts](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Access community events to learn from Microsoft 365 experts.
- [Personalized recommendations](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) Find developer resources quickly from your personalized dashboard.

## Additional resources

* [Microsoft Graph documentation](https://docs.microsoft.com/graph/)
* [Office Add-ins documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## Copyright
Copyright (c) 2019 - 2021 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/auth/Office-Add-in-ASPNET-SSO" />
