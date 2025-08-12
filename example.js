require('dotenv').config()
const nodemailer = require('nodemailer')
const MsgraphTransport = require('./index.js')

async function sendTestEmails () {
  const msGraphConfig = {
    clientId: process.env.MS_GRAPH_CLIENT_ID,
    tenantId: process.env.MS_GRAPH_TENANT_ID,
    clientSecret: process.env.MS_GRAPH_CLIENT_SECRET,
    userPrincipalName: process.env.MS_GRAPH_PRINCIPAL,
    saveToSentItems: false
  }

  const msGraphTransport = new MsgraphTransport(msGraphConfig)
  const transporter = nodemailer.createTransport(msGraphTransport)

  // optionnal: allow sending multiple mails in a batch without getting a new access token each time
  // the token will be stored in cache and reused until it expires
  await msGraphTransport.getAccessToken()

  transporter.sendMail({
    from: process.env.MAIL_FROM,
    to: process.env.MAIL_TO,
    subject: 'hello',
    text: 'This is a test email sent using **nodemailer** with MSAL authentication and some long line that need to be cut when transformed, so some equal sign a added at the end of line, do not know why but it is causing some issue with the delivery',
    html: '<p>This is a test email sent using <strong>nodemailer</strong> with MSAL authentication and some long line that need to be cut when transformed, so some equal sign a added at the end of line, do not know why but it is causing some issue with the delivery.</p>',
    saveToSentItems: false
  })
}

sendTestEmails ()
