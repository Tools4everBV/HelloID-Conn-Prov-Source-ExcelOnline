# HelloID-Conn-Prov-Source-ExcelOnline

| :warning: Warning |
|:---------------------------|
| Note that this HelloID connector has not been tested in a production environment!      |

| :information_source: Information |
|:---------------------------|
| This repository contains the connector and configuration code only. The implementer is responsible to acquire the connection details such as username, password, certificate, etc. You might even need to sign a contract or agreement with the supplier before implementing this connector. Please contact the client's application manager to coordinate the connector requirements.       |
<br />
<p align="center">
  <img src="https://tools4ever.nl/connector-logos/excel-logo.png">
</p>

<!-- TABLE OF CONTENTS -->
## Table of Contents
- [Table of Contents](#table-of-contents)
- [Introduction](#introduction)
- [Getting the Azure AD graph API access](#getting-the-azure-ad-graph-api-access)
  - [Application Registration](#application-registration)
  - [Configuring App Permissions](#configuring-app-permissions)
  - [Authentication and Authorization](#authentication-and-authorization)
  - [Connection settings](#connection-settings)
- [Getting help](#getting-help)
- [HelloID Docs](#helloid-docs)

## Introduction

This connector retrieves data from an Excel Online Sheet

It now supports that it is located in a Onedrive Folder from a User or a Sharepoint Site

<!-- GETTING STARTED -->
## Getting the Azure AD graph API access

By using this connector you will have the ability to get data from an Excel Online sheet.

### Application Registration
The first step to connect to Graph API and make requests, is to register a new <b>Azure Active Directory Application</b>. The application is used to connect to the API and to manage permissions.

* Navigate to <b>App Registrations</b> in Azure, and select “New Registration” (<b>Azure Portal > Azure Active Directory > App Registration > New Application Registration</b>).
* Next, give the application a name. In this example we are using “<b>HelloID PowerShell</b>” as application name.
* Specify who can use this application (<b>Accounts in this organizational directory only</b>).
* Specify the Redirect URI. You can enter any url as a redirect URI value. In this example we used http://localhost because it doesn't have to resolve.
* Click the “<b>Register</b>” button to finally create your new application.

Some key items regarding the application are the Application ID (which is the Client ID), the Directory ID (which is the Tenant ID) and Client Secret.

### Configuring App Permissions
The [Microsoft Graph documentation](https://docs.microsoft.com/en-us/graph) provides details on which permission are required for each permission type.

To assign your application the right permissions, navigate to <b>Azure Portal > Azure Active Directory >App Registrations</b>.
Select the application we created before, and select “<b>API Permissions</b>” or “<b>View API Permissions</b>”.
To assign a new permission to your application, click the “<b>Add a permission</b>” button.
From the “<b>Request API Permissions</b>” screen click “<b>Microsoft Graph</b>”.
For this connector the following permissions are used as <b>Application permissions</b>:
*	Read files in all site collections (Onedrive) by using <b><i>Files.Read.All</i></b>
*	Read items in all site collections (Sharepoint) by using <b><i>Sites.Read.All</i></b>

Some high-privilege permissions can be set to admin-restricted and require an administrators consent to be granted.

To grant admin consent to our application press the “<b>Grant admin consent for TENANT</b>” button.

### Authentication and Authorization
There are multiple ways to authenticate to the Graph API with each has its own pros and cons, in this example we are using the client_credentials grant type.

*	First we need to get the <b>Client ID</b>, go to the <b>Azure Portal > Azure Active Directory > App Registrations</b>.
*	Select your application and copy the Application (client) ID value.
*	After we have the Client ID we also have to create a <b>Client Secret</b>.
*	From the Azure Portal, go to <b>Azure Active Directory > App Registrations</b>.
*	Select the application we have created before, and select "<b>Certificates and Secrets</b>". 
*	Under “Client Secrets” click on the “<b>New Client Secret</b>” button to create a new secret.
*	Provide a logical name for your secret in the Description field, and select the expiration date for your secret.
*	It's IMPORTANT to copy the newly generated client secret, because you cannot see the value anymore after you close the page.
*	At last we need to get is the <b>Tenant ID</b>. This can be found in the Azure Portal by going to <b>Azure Active Directory > Custom Domain Names</b>, and then finding the .onmicrosoft.com domain.

### Connection settings
The following settings are required to connect to the API.

| Setting     | Description |
| ------------ | ----------- |
| Client ID | Id of the Azure app |
| Client Secret | Secret of the Azure app |
| Tenant ID | Id of the Azure tenant |
| Use Sharepoint (instead of Onedrive) | By default the Script searches for the file in a users onedrive folder - with this switch you can select to search a Sharepoint Site|
| Site Name (Sharepoint) | Name of the Sharepoint Site where the file is located|
| List Name (Sharepoint) | If the File is not located in the Default Documents Folder - Name the List where it is located|
| User ID (Onedrive) | Id of the Azure User where the Sheet is located - example: 12345678-1234-1234-1234-12345678901234|
| Document Path | Path to the document - Replace "/" with %2F - example: sheet.xlsx if it is located in the root of your documents - folder%2Fsheet.xlsx if it is in a subfolder |
| Table Name | Name of the Sheet in the Document - example: Sheet1 or Tabelle1 |

Please correct the column numbers in the powershell script

Please make sure that you clicked on "Insert -> Table" in Excel or it will not work

## Getting help
> _For more information on how to configure a HelloID PowerShell connector, please refer to our [documentation](https://docs.helloid.com/hc/en-us/articles/360012518799-How-to-add-a-target-system) pages_

> _If you need help, feel free to ask questions on our [forum](https://forum.helloid.com)_

## HelloID Docs
The official HelloID documentation can be found at: https://docs.helloid.com/
