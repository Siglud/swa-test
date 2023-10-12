param resourceBaseName string
param storageSKU string
param functionStorageSKU string
param functionAppSKU string

param aadAppClientId string
param aadAppTenantId string
param aadAppOauthAuthorityHost string
@secure()
param aadAppClientSecret string
param staticWebAppName string = resourceBaseName

// Azure Static Web Apps that hosts your static web site
resource swa 'Microsoft.Web/staticSites@2022-09-01' = {
  name: staticWebAppName
  // SWA do not need location setting
  location: 'westus2'
  sku: {
    name: 'Free'
    tier: 'Free'
  }
  properties:{}
}

resource symbolicname 'Microsoft.Web/staticSites/config@2022-09-01' = {
  name: 'appsettings'
  kind: 'string'
  parent: swa
  properties: {
    M365_AUTHORITY_HOST: aadAppOauthAuthorityHost
    M365_TENANT_ID: aadAppTenantId
    M365_CLIENT_ID: aadAppClientId
    M365_CLIENT_SECRET: aadAppClientSecret
  }
}


var siteDomain = swa.properties.defaultHostname
var tabEndpoint = 'https://${siteDomain}'

var apiEndpoint = 'https://${siteDomain}'

// The output will be persisted in .env.{envName}. Visit https://aka.ms/teamsfx-actions/arm-deploy for more details.
output TAB_DOMAIN string = siteDomain
output TAB_ENDPOINT string = tabEndpoint
output API_FUNCTION_ENDPOINT string = apiEndpoint
output SECRET_TAB_SWA_DEPLOY_TOKEN string = swa.listSecrets().properties.apiKey
