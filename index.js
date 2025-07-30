const addressParser = require('nodemailer/lib/addressparser')
const msal = require('@azure/msal-node')

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
    const token = await this.getAccessToken()

    const from = addressParser(mail.data.from)[0].address
    const envelope = mail.message.getEnvelope()
    const messageId = mail.message.messageId()

    mail.message.build(async (err, data) => {
      if (err) {
        done(err)
        return
      }

      const response = await fetch(`https://graph.microsoft.com/v1.0/users/${from}/sendMail`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token.accessToken}`,
          'Content-Type': 'text/plain'
        },
        body: data.toString('base64')
      })

      if (response.status !== 202) {
        const errorText = await response.text()
        done(new Error(`Failed to send email: ${response.status} ${errorText}`))
      } else {
        done(null, { envelope, messageId })
      }
    })
  }
}

module.exports = MsgraphTransport
