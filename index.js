const msal = require('@azure/msal-node')
const addressParser = require('nodemailer/lib/addressparser')

class MsgraphTransport {
  name = 'msGraph'
  version = require("./package.json").version

  constructor (config) {
    this.config = config
    this.auth = {
      clientId: this.config.clientId,
      authority: this.config.authority || `https://login.microsoftonline.com/${this.config.tenantId}`,
      clientSecret: this.config.clientSecret
    }

    this.cca = new msal.ConfidentialClientApplication({ auth: this.auth })
  }

  async getAccessToken () {
    const clientCredentialRequest = {
      scopes: ['https://graph.microsoft.com/.default']
    }

    return this.cca.acquireTokenByClientCredential(clientCredentialRequest)
  }

  async send (mail, done) {
    const envelope = mail.message.getEnvelope()
    const messageId = mail.message.messageId()

    const message = {
      subject: mail.data.subject
    }
    message.body = mail.data.html ? {
      contentType: 'html',
      content: mail.data.html
    } : {
      contentType: 'text',
      content: mail.data.text
    }
    if (mail.data.from) {
      message.from = { emailAddress: addressParser(mail.data.from)[0] }
    }
    if (mail.data.to) {
      message.toRecipients = addressParser(mail.data.to).map(recipient => ({ emailAddress: recipient}))
    }
    if (mail.data.cc) {
      message.ccRecipients = addressParser(mail.data.cc).map(recipient => ({ emailAddress: recipient}))
    }
    if (mail.data.bcc) {
      message.bccRecipients = addressParser(mail.data.bcc).map(recipient => ({ emailAddress: recipient}))
    }

    const userPrincipalName = mail.data.userPrincipalName || this.config.userPrincipalName
    const saveToSentItems = mail.data.saveToSentItems !== undefined ? mail.data.saveToSentItems : this.config.saveToSentItems

    const token = await this.getAccessToken()
    const response = await fetch(`https://graph.microsoft.com/v1.0/users/${userPrincipalName}/sendMail`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token.accessToken}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ message, saveToSentItems })
    })

    if (response.status !== 202) {
      const errorText = await response.text()
      done(new Error(`Failed to send email: ${response.status} ${errorText}`))
    } else {
      done(null, { envelope, messageId })
    }
  }
}

module.exports = MsgraphTransport
