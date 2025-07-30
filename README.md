# Microsoft Graph API transport for Nodemailer

This module is a transport plugin for [Nodemailer](https://github.com/andris9/Nodemailer)
that use [Microsoft Graph API](https://learn.microsoft.com/en-us/graph)
to send e-mails from a tenant.

## Usage
### Creation of an application on Microsoft Entra
1. Go to [Microsoft Entra admin center](https://entra.microsoft.com)
2. Click on **App registrations** then **New registration**
3. Give a name, leave other parameters as default (single tenant, no redirect
   URL)
4. Register the new app
5. Keep the **Application (Client) ID** and **Directory (Tenant) ID**

#### Create a secret
1. Go to Manage -> Certificates and secrets
2. **New client secret**, give it a name
3. Keep the client secret **value**

#### Add graph permissions
1. Go to Manage -> API permissions
2. **Add a permission**
3. Select **Microsoft Graph** then **Application permissions**
4. Search for **Mail.Send**, select it and click on **Add permissions**
5. Click on **Grant admin consent for *"your tenant"***

### Install this transport
Install via npm

    npm install nodemailer-msgraph-transport


Require the module and intialize it with the Microsoft application credentials

```javascript
const nodemailer = require('nodemailer')
const MsGraphTransport = required('nodemailer-msgraph-transport')

const msGraphConfig = {
  clientId: 'MS_GRAPH_CLIENT_ID',
  tenantId: 'MS_GRAPH_TENANT_ID',
  clientSecret: 'MS_GRAPH_TENANT_SECRET_VALUE',
  userPrincipalName: 'id | userPrincipalName'
  saveToSentItems: true
}

const msGraphTransport = new MsgraphTransport(msGraphConfig)

const transporter = nodemailer.createTransport(msGraphTransport)
```

**userPrincipalName** and **saveToSentItems** can be overridden for each mail
in the message object passed to ```sendMail()```

When sending a batch of e-mails, you can request a token before hand to avoid
unnecessary requests to Microsoft graph api authentication server. The token
will be automatically cached.

```javascript
await msGraphTransport.getAccessToken()

transporter.sendMail(mailData)
transporter.sendMail(mailData)
...
```
## Limitations
Attachements are currently not supported.

## References
- SendMail graph function: https://learn.microsoft.com/en-us/graph/api/user-sendmail
- graph message: https://learn.microsoft.com/en-us/graph/api/resources/message

## Licence
Licensed under the MIT Licence.

