
# Requires MSAL.PS PowerShell modules
# https://www.powershellgallery.com/packages/MSAL.PS

# NOTE! Client ID used in the API call is the sample Client ID not meant for production use
# See more: https://learn.microsoft.com/en-us/power-apps/developer/data-platform/xrm-tooling/use-connection-strings-xrm-tooling-connect#connection-string-parameters

# API endpoint used is still in Preview https://learn.microsoft.com/en-us/power-platform/admin/create-dataverseapplicationuser
param
(
    [Parameter(Mandatory=$true)]
    [string]$TenantName = "????????????.onmicrosoft.com", 
    
    [Parameter(Mandatory=$true)]
    [string]$EnvironmentID = "",

    [Parameter(Mandatory=$true)]
    [string]$ApplicationID = ""
)

$connectionDetails = @{
    'TenantId'     = $TenantName
    'ClientId'     = '51f81489-12ee-4a9e-aaae-a2591f45987d' # This Client ID is meant only for development and testing purposes. See more: https://learn.microsoft.com/en-us/power-apps/developer/data-platform/xrm-tooling/use-connection-strings-xrm-tooling-connect#connection-string-parameters
    'Interactive'  = $true
    'RedirectUri'  = 'https://localhost'
    'Scopes'    = 'https://api.bap.microsoft.com/.default'
}

$t = Get-MsalToken @connectionDetails

$authHeader = @{
    "Authorization" = $t.CreateAuthorizationHeader()
    "Content-type" = "application/json"
}

$body = @{
    servicePrincipalAppId = $ApplicationID
} | ConvertTo-Json

$uri = "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/$EnvironmentID/addAppUser?api-version=2020-10-01"

Invoke-RestMethod -Uri $uri -Method Post -Headers $authHeader -Body $body
