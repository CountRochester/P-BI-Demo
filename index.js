import * as pbi from 'powerbi-client'

const reportContainer = document.querySelector('#report-container')

// Initialize iframe for embedding report
window.powerbi.bootstrap(reportContainer, { type: 'report' })

const models = pbi.models;
const reportLoadConfig = {
  type: 'report',
  tokenType: models.TokenType.Embed,

  // Enable this setting to remove gray shoulders from embedded report
  // settings: {
  //     background: models.BackgroundType.Transparent
  // }
}

async function fetchReport (url) {
  try {
    const data = await (await fetch(url)).json()
    const embedData = JSON.parse(JSON.stringify(data))
    reportLoadConfig.accessToken = embedData.accessToken

    // You can embed different reports as per your need
    reportLoadConfig.embedUrl = embedData.reportConfig[0].embedUrl

    // Use the token expiry to regenerate Embed token for seamless end user experience
    // Refer https://aka.ms/RefreshEmbedToken
    const tokenExpiry = embedData.tokenExpiry

    // Embed Power BI report when Access token and Embed URL are available
    const report = window.powerbi.embed(reportContainer, reportLoadConfig);

    // Triggers when a report schema is successfully loaded
    report.on('loaded', () => {
      console.log('Report load successful')
    })

    // Triggers when a report is successfully embedded in UI
    report.on('rendered', () => {
      console.log('Report render successful')
    })

    // Clear any other error handler event
    report.off('error')

    // Below patch of code is for handling errors that occur during embedding
    report.on('error', event => {
      const errorMsg = event.detail

      // Use errorMsg variable to log error in any destination of choice
      console.error(errorMsg)
      return
    })
  } catch (err) {
    const errorContainer = document.querySelector('.error-container')
    const embedContainer = document.querySelector('.embed-container')
    embedContainer.style.display = 'none'
    errorContainer.style.display = 'block'

    // Format error message
    const errMessageHtml = '<strong> Error Details: </strong> <br/>' + JSON.parse(err?.responseText)?.errorMsg
    errMessageHtml = errMessageHtml.split('\n').join('<br/>')

    // Show error message on UI
    errorContainer.innerHTML = errMessageHtml
  }
}

fetchReport('/getembedinfo')
