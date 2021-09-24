<#
.SYNOPSIS
Add AAD Application and SPN to Dynamics365 AAD and configure Dynamics365 to accept the SPN as tenant admin user.

.DESCRIPTION
This script assists in creating and configuring the ServicePrincipal to be used with
the Power Platform Build Tools AzureDevOps task library.

Registers an Application object and corresponding ServicePrincipalName (SPN) with the Dynamics365 AAD instance.
This Application is then added as admin user to the Dynamics365 tenant itself.
NOTE: This script will prompt *TWICE* with the AAD login dialogs:
    1. time: to login as admin to the AAD instance associated with the Dynamics365 tenant
    2. time: to login as tenant admin to the Dynamics365 tenant itself

.INPUTS
None

.OUTPUTS
Object with D365 TenantId, ApplicationId and client secret (in clear text);
use this triple to configure the AzureDevOps ServiceConnection

.LINK
https://marketplace.visualstudio.com/items?itemName=microsoft-IsvExpTools.PowerPlatform-BuildTools

.EXAMPLE
> New-CrmServicePrincipal
> New-CrmServicePrincipal -TenantLocation "Europe"
> New-CrmServicePrincipal -AdminUrl "https://admin.services.crm4.dynamics.com"
#>
[CmdletBinding()]
Param(
    # gather permission requests but don't create any AppId nor ServicePrincipal
    [switch] $DryRun = $false,
    # other possible Azure environments, see: https://docs.microsoft.com/en-us/powershell/module/azuread/connect-azuread?view=azureadps-2.0#parameters
    [string] $AzureEnvironment = "AzureCloud",

    [ValidateSet(
        "UnitedStates",
        "Preview(UnitedStates)",
        "Europe",
        "EMEA",
        "Asia",
        "Australia",
        "Japan",
        "SouthAmerica",
        "India",
        "Canada",
        "UnitedKingdom",
        "France"
    )]
    [string] $TenantLocation = "UnitedStates",
    [string] $AdminUrl
)

$adminUrls = @{
    "UnitedStates"	            =	"https://admin.services.crm.dynamics.com"
    "Preview(UnitedStates)"	    =	"https://admin.services.crm9.dynamics.com"
    "Europe"		            =	"https://admin.services.crm4.dynamics.com"
    "EMEA"	                    =	"https://admin.services.crm4.dynamics.com"
    "Asia"	                    =	"https://admin.services.crm5.dynamics.com"
    "Australia"	                =	"https://admin.services.crm6.dynamics.com"
    "Japan"		                =	"https://admin.services.crm7.dynamics.com"
    "SouthAmerica"	            =	"https://admin.services.crm2.dynamics.com"
    "India"		                =	"https://admin.services.crm8.dynamics.com"
    "Canada"		            =	"https://admin.services.crm3.dynamics.com"
    "UnitedKingdom"	            =	"https://admin.services.crm11.dynamics.com"
    "France"		            =	"https://admin.services.crm12.dynamics.com"
    }

    function ensureModules {
    $dependencies = @(
        # the more general and modern "Az" a "AzureRM" do not have proper support to manage permissions
        @{ Name = "AzureAD"; Version = [Version]"2.0.2.76"; "InstallWith" = "Install-Module -Name AzureAD -AllowClobber -Scope CurrentUser" },
        @{ Name = "Microsoft.Xrm.OnlineManagementAPI"; Version = [Version]"1.2.0.1"; "InstallWith" = "Install-Module -Name Microsoft.Xrm.OnlineManagementAPI -AllowClobber -Scope CurrentUser" }
    )
    $missingDependencies = $false
    $dependencies | ForEach-Object -Process {
        $moduleName = $_.Name
        $deps = (Get-Module -ListAvailable -Name $moduleName `
            | Sort-Object -Descending -Property Version)
        if ($deps -eq $null) {
            Write-Error "Required module not installed; install from PowerShell prompt with: '$($_.InstallWith)'"
            $missingDependencies = $true
            return
        }
        $dep = $deps[0]
        if ($dep.Version -lt $_.Version) {
            Write-Error "Required module installed but does not meet minimal required version: found: $($dep.Version), required: >= $($_.Version); run: 'Update-Module '$($_.Name)'"
            $missingDependencies = $true
            return
        }
        Import-Module $moduleName
    }
    if ($missingDependencies) {
        throw "Missing required dependencies!"
    }
}

function connectAAD {
    Write-Host @"

Connecting to AzureAD: Please log in, using your Dynamics365 / Power Platform tenant ADMIN credentials:

"@
    try {
        Connect-AzureAD -AzureEnvironmentName $AzureEnvironment -ErrorAction Stop | Out-Null
    }
    catch {
        throw "Failed to login: $($_.Exception.Message)"
    }
    return Get-AzureADCurrentSessionInfo
}

function reconnectAAD {
    # for tenantID, see DirectoryID here: https://aad.portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/Overview
    try {
        $session = Get-AzureADCurrentSessionInfo -ErrorAction SilentlyContinue
        if ($session.Environment.Name -ne $AzureEnvironment) {
            Disconnect-AzureAd
            $session = connectAAD
        }
    }
    catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException] {
        $session = connectAAD
    }
    $tenantId = $session.TenantId
    Write-Host @"
Connected to AAD tenant: $($session.TenantDomain) ($($tenantId)) in $($session.Environment.Name)

"@
    return $tenantId
}

function addRequiredAccess {
    param(
        [System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.RequiredResourceAccess]] $requestList,
        [Microsoft.Open.AzureAD.Model.ServicePrincipal[]] $spns,
        [string] $spnDisplayName,
        [string] $permissionName
    )
    Write-Host "  - requiredAccess for $spnDisplayName - $permissionName"
    $selectedSpns = $spns | Where-Object { $_.DisplayName -eq $spnDisplayName }

    # have to build the List<ResourceAccess> item by item since PS doesn't deal well with generic lists (which is the signature for .ResourceAccess)
    $selectedSpns | ForEach-Object -process {
        $spn = $_
        $accessList = New-Object -TypeName 'System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.ResourceAccess]'
        ( $spn.OAuth2Permissions `
        | Where-Object { $_.Value -eq $permissionName } `
        | ForEach-Object -process {
            $acc = New-Object -TypeName 'Microsoft.Open.AzureAD.Model.ResourceAccess'
            $acc.Id = $_.Id
            $acc.Type = "Scope"
            $accessList.Add($acc)
        } )
        Write-Verbose "accessList: $accessList"

        # TODO: filter out the now-obsoleted SPN for CDS user_impersonation: id = 9f7cb6a3-2591-431e-b80d-385fce1f93aa (PowerApps Runtime), see once granted admin consent in SPN permissions
        $req  = New-Object -TypeName 'Microsoft.Open.AzureAD.Model.RequiredResourceAccess'
        $req.ResourceAppId = $spn.AppId
        $req.ResourceAccess = $accessList
        $requestList.Add($req)
    }
}

function calculateSecretKey {
    param (
        [int] $length = 32
    )
    $secret = [System.Byte[]]::new($length)
    $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider

    # restrict to printable alpha-numeric characters
    $validCharSet = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
    function getRandomChar {
        param (
            [uint32] $min = 0,
            [uint32] $max = $validCharSet.length - 1
        )
        $diff = $max - $min + 1
        [Byte[]] $bytes = 1..4
        $rng.getbytes($bytes)
        $number = [System.BitConverter]::ToUInt32(($bytes), 0)
        $index = [char] ($number % $diff + $min)
        return $validCharSet[$index]
    }
    for ($i = 0; $i -lt $length; $i++) {
        $secret[$i] = getRandomChar
    }
    return $secret
}

ensureModules
$ErrorActionPreference = "Stop"
$tenantId = reconnectAAD

$allSPN = Get-AzureADServicePrincipal -All $true

$requiredAccess = New-Object -TypeName 'System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.RequiredResourceAccess]'

addRequiredAccess $requiredAccess $allSPN "Microsoft Graph" "User.Read"
addRequiredAccess $requiredAccess $allSPN "PowerApps-Advisor" "Analysis.All"
addRequiredAccess $requiredAccess $allSPN "Common Data Service" "user_impersonation"

$appBaseName = "$((Get-AzureADTenantDetail).VerifiedDomains.Name)-$(get-date -Format "yyyyMMdd-HHmmss")"
$spnDisplayName = "App-$($appBaseName)"

Write-Verbose "Creating AAD Application: '$spnDisplayName'..."
$appId = "<dryrun-no-app-created>"
$spnId = "<dryrun-no-spn-created>"
if (!$DryRun) {
    # https://docs.microsoft.com/en-us/azure/active-directory/develop/app-objects-and-service-principals
    $app = New-AzureADApplication -DisplayName $spnDisplayName -PublicClient $true -ReplyUrls "urn:ietf:wg:oauth:2.0:oob" -RequiredResourceAccess $requiredAccess
    $appId = $app.AppId
}
Write-Host "Created AAD Application: '$spnDisplayName' with appID $appId (objectId: $($app.ObjectId)"

$secretText = [System.Text.Encoding]::UTF8.GetString((calculateSecretKey))

Write-Verbose "Creating Service Principal Name (SPN): '$spnDisplayName'..."
if (!$DryRun) {
    # display name of SPN must be same as for the App itself
    # https://docs.microsoft.com/en-us/powershell/module/azuread/new-azureadserviceprincipal?view=azureadps-2.0
    $spn = New-AzureADServicePrincipal -AccountEnabled $true -AppId $appId -AppRoleAssignmentRequired $true -DisplayName $spnDisplayName -Tags {WindowsAzureActiveDirectoryIntegratedApp}
    $spnId = $spn.ObjectId

    $spnKey = New-AzureADServicePrincipalPasswordCredential -ObjectId $spn.ObjectId -StartDate (get-date).AddHours(-1) -EndDate (get-date).AddYears(1) -Value $secretText
    Set-AzureADServicePrincipal -ObjectId $spn.ObjectId -PasswordCredentials @($spnKey)
}
Write-Host "Created SPN '$spnDisplayName' with objectId: $spnId"

Write-Host @"

Connecting to Dynamics365 CRM managment API and adding appID to Dynamics365 tenant:
    Please log in, using your Dynamics365 / Power Platform tenant ADMIN credentials:
"@

if (!$DryRun) {
    if ($PSBoundParameters.ContainsKey("AdminUrl")) {
        $adminApi = $AdminUrl
    } else {
        $adminApi = $adminUrls[$TenantLocation]
    }
    Write-Host "Admin Api is: $adminApi"
    $mgmtApp = New-CrmManagementApp -ApiUrl $adminApi -AppId $appId -TenantId $tenantId -Enable
    Write-Host @"

Added appId $($appId) to D365 tenant ($($tenantId))

"@
}
$result = [PSCustomObject] @{
    TenantId = $tenantId;
    ApplicationId = $appId;
    ClientSecret = $secretText
}
Write-Output $result

# SIG # Begin signature block
# MIIdhAYJKoZIhvcNAQcCoIIddTCCHXECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUAtY2Spj/QkdCFzn3bo+wNcue
# 9P2gghhuMIIE3jCCA8agAwIBAgITMwAAAVVgYcb45V+AeQAAAAABVTANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTkxMjE5MDExMzAx
# WhcNMjEwMzE3MDExMzAxWjCBzjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEpMCcGA1UECxMgTWljcm9zb2Z0IE9wZXJhdGlvbnMgUHVlcnRvIFJp
# Y28xJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOjc4ODAtRTM5MC04MDE0MSUwIwYD
# VQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNlMIIBIjANBgkqhkiG9w0B
# AQEFAAOCAQ8AMIIBCgKCAQEAjkmUktwnUA7Bn3t3ifGSV/oYP/4ZVNKlMJRCFsEL
# apHJSidS6N8r5WnrCdafgB+ISduNic7YvVHc75kmfFukvyqOAPpDDChYONWYdIHp
# WBCGpzuO5RRYhNRYyPk3MLH7Ti5uPMxrpe7DTINFYaFJNTNk2O7+C23mR3t8sVSD
# IfZAP2NhoKieZl2EBiR4EC47NPvHD66xW/TZblQkmbdynCF3RTed30rH4E74iky1
# b4RbH9/BJMABJZ4Ul2Zbf6SGpRQ07c9SbytPhvIPtA8+0HtOp2hUjV7m2TN8C5+U
# WaJXqFxFJb7Mj8pDjYPHAJYh7VlceDhiYiuPs1OfXsWnlQIDAQABo4IBCTCCAQUw
# HQYDVR0OBBYEFDoSsdt37lJKTyUVh2uvEVwtsQrrMB8GA1UdIwQYMBaAFCM0+NlS
# RnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly9jcmwubWlj
# cm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY3Jvc29mdFRpbWVTdGFtcFBD
# QS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsGAQUFBzAChjxodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcnQw
# EwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZIhvcNAQEFBQADggEBAFJZAoNaXqOx
# QBW25ktP2JIIrgnJrjxkXF9ECI+jDNJcFyGjx2FR2b/lpA/IWz40+ISVPlhxxkHI
# Ghn6M28RwFlRoq1Sh7u/+ApENoVDSzdCPtwKRmYustcHN9hQUVRQmcbC7vJiBiHm
# BVbgU9Rz8leNlqNWnxfbkGy/uS32khDXbyhDb2kyiK5n/J8xEyFx+og8jxBHTcb4
# Wetkjl6qp5XKlAX/Lp+ATn8JtCklBjfTTn+G/9iyEB/0q7+/80lbvh1DIuJi1Hbf
# +GXtbpn9U7ejOyp3xKPVZInFs/ZJyXZmctvS9LbOfwcnhZz0zH3uRKeqMIeGCS46
# QdwwZ9lK1PwwggX/MIID56ADAgECAhMzAAABh3IXchVZQMcJAAAAAAGHMA0GCSqG
# SIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
# DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
# KDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTEwHhcNMjAw
# MzA0MTgzOTQ3WhcNMjEwMzAzMTgzOTQ3WjB0MQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29ycG9yYXRpb24w
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDOt8kLc7P3T7MKIhouYHew
# MFmnq8Ayu7FOhZCQabVwBp2VS4WyB2Qe4TQBT8aBznANDEPjHKNdPT8Xz5cNali6
# XHefS8i/WXtF0vSsP8NEv6mBHuA2p1fw2wB/F0dHsJ3GfZ5c0sPJjklsiYqPw59x
# J54kM91IOgiO2OUzjNAljPibjCWfH7UzQ1TPHc4dweils8GEIrbBRb7IWwiObL12
# jWT4Yh71NQgvJ9Fn6+UhD9x2uk3dLj84vwt1NuFQitKJxIV0fVsRNR3abQVOLqpD
# ugbr0SzNL6o8xzOHL5OXiGGwg6ekiXA1/2XXY7yVFc39tledDtZjSjNbex1zzwSX
# AgMBAAGjggF+MIIBejAfBgNVHSUEGDAWBgorBgEEAYI3TAgBBggrBgEFBQcDAzAd
# BgNVHQ4EFgQUhov4ZyO96axkJdMjpzu2zVXOJcswUAYDVR0RBEkwR6RFMEMxKTAn
# BgNVBAsTIE1pY3Jvc29mdCBPcGVyYXRpb25zIFB1ZXJ0byBSaWNvMRYwFAYDVQQF
# Ew0yMzAwMTIrNDU4Mzg1MB8GA1UdIwQYMBaAFEhuZOVQBdOCqhc3NyK1bajKdQKV
# MFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lv
# cHMvY3JsL01pY0NvZFNpZ1BDQTIwMTFfMjAxMS0wNy0wOC5jcmwwYQYIKwYBBQUH
# AQEEVTBTMFEGCCsGAQUFBzAChkVodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# b3BzL2NlcnRzL01pY0NvZFNpZ1BDQTIwMTFfMjAxMS0wNy0wOC5jcnQwDAYDVR0T
# AQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEAixmyS6E6vprWD9KFNIB9G5zyMuIj
# ZAOuUJ1EK/Vlg6Fb3ZHXjjUwATKIcXbFuFC6Wr4KNrU4DY/sBVqmab5AC/je3bpU
# pjtxpEyqUqtPc30wEg/rO9vmKmqKoLPT37svc2NVBmGNl+85qO4fV/w7Cx7J0Bbq
# k19KcRNdjt6eKoTnTPHBHlVHQIHZpMxacbFOAkJrqAVkYZdz7ikNXTxV+GRb36tC
# 4ByMNxE2DF7vFdvaiZP0CVZ5ByJ2gAhXMdK9+usxzVk913qKde1OAuWdv+rndqkA
# Im8fUlRnr4saSCg7cIbUwCCf116wUJ7EuJDg0vHeyhnCeHnBbyH3RZkHEi2ofmfg
# nFISJZDdMAeVZGVOh20Jp50XBzqokpPzeZ6zc1/gyILNyiVgE+RPkjnUQshd1f1P
# Mgn3tns2Cz7bJiVUaqEO3n9qRFgy5JuLae6UweGfAeOo3dgLZxikKzYs3hDMaEtJ
# q8IP71cX7QXe6lnMmXU/Hdfz2p897Zd+kU+vZvKI3cwLfuVQgK2RZ2z+Kc3K3dRP
# z2rXycK5XCuRZmvGab/WbrZiC7wJQapgBodltMI5GMdFrBg9IeF7/rP4EqVQXeKt
# evTlZXjpuNhhjuR+2DMt/dWufjXpiW91bo3aH6EajOALXmoxgltCp1K7hrS6gmsv
# j94cLRf50QQ4U8QwggYHMIID76ADAgECAgphFmg0AAAAAAAcMA0GCSqGSIb3DQEB
# BQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAXBgoJkiaJk/IsZAEZFgltaWNy
# b3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhv
# cml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0MDMxMzAzMDlaMHcxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFBDQTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJ+h
# bLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn0UytdDAgEesH1VSVFUmUG0KS
# rphcMCbaAGvoe73siQcP9w4EmPCJzB/LMySHnfL0Zxws/HvniB3q506jocEjU8qN
# +kXPCdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4nrIZPVVIM5AMs+2qQkDBuh/NZ
# MJ36ftaXs+ghl3740hPzCLdTbVK0RZCfSABKR2YRJylmqJfk0waBSqL5hKcRRxQJ
# gp+E7VV4/gGaHVAIhQAQMEbtt94jRrvELVSfrx54QTF3zJvfO4OToWECtR0Nsfz3
# m7IBziJLVP/5BcPCIAsCAwEAAaOCAaswggGnMA8GA1UdEwEB/wQFMAMBAf8wHQYD
# VR0OBBYEFCM0+NlSRnAK7UD7dvuzK7DDNbMPMAsGA1UdDwQEAwIBhjAQBgkrBgEE
# AYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQOrIJgQFYnl+UlE/wq4QpTlVnkpKFj
# pGEwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jv
# c29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9y
# aXR5ghB5rRahSqClrUxzWPQHEy5lMFAGA1UdHwRJMEcwRaBDoEGGP2h0dHA6Ly9j
# cmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL21pY3Jvc29mdHJvb3Rj
# ZXJ0LmNybDBUBggrBgEFBQcBAQRIMEYwRAYIKwYBBQUHMAKGOGh0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0Um9vdENlcnQuY3J0MBMG
# A1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEBBQUAA4ICAQAQl4rDXANENt3p
# tK132855UU0BsS50cVttDBOrzr57j7gu1BKijG1iuFcCy04gE1CZ3XpA4le7r1ia
# HOEdAYasu3jyi9DsOwHu4r6PCgXIjUji8FMV3U+rkuTnjWrVgMHmlPIGL4UD6ZEq
# JCJw+/b85HiZLg33B+JwvBhOnY5rCnKVuKE5nGctxVEO6mJcPxaYiyA/4gcaMvnM
# MUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tOi3/FNSteo7/rvH0LQnvUU3Ih7jDK
# u3hlXFsBFwoUDtLaFJj1PLlmWLMtL+f5hYbMUVbonXCUbKw5TNT2eb+qGHpiKe+i
# myk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0NbhOxXEjEiZ2CzxSjHFaRkMU
# vLOzsE1nyJ9C/4B5IYCeFTBm6EISXhrIniIh0EPpK+m79EjMLNTYMoBMJipIJF9a
# 6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2JoXZhtG6hE6a/qkfwEm/9ijJs
# sv7fUciMI8lmvZ0dhxJkAj0tr1mPuOQh5bWwymO0eFQF1EEuUKyUsKV4q7OglnUa
# 2ZKHE3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng9wFlb4kLfchpyOZu6qeXzjEp
# /w7FW1zYTRuh2Povnj8uVRZryROj/TCCB3owggVioAMCAQICCmEOkNIAAAAAAAMw
# DQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5n
# dG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
# YXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhv
# cml0eSAyMDExMB4XDTExMDcwODIwNTkwOVoXDTI2MDcwODIxMDkwOVowfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBAKvw+nIQHC6t2G6qghBNNLrytlghn0IbKmvpWlCquAY4GgRJun/D
# DB7dN2vGEtgL8DjCmQawyDnVARQxQtOJDXlkh36UYCRsr55JnOloXtLfm1OyCizD
# r9mpK656Ca/XllnKYBoF6WZ26DJSJhIv56sIUM+zRLdd2MQuA3WraPPLbfM6XKEW
# 9Ea64DhkrG5kNXimoGMPLdNAk/jj3gcN1Vx5pUkp5w2+oBN3vpQ97/vjK1oQH01W
# KKJ6cuASOrdJXtjt7UORg9l7snuGG9k+sYxd6IlPhBryoS9Z5JA7La4zWMW3Pv4y
# 07MDPbGyr5I4ftKdgCz1TlaRITUlwzluZH9TupwPrRkjhMv0ugOGjfdf8NBSv4yU
# h7zAIXQlXxgotswnKDglmDlKNs98sZKuHCOnqWbsYR9q4ShJnV+I4iVd0yFLPlLE
# tVc/JAPw0XpbL9Uj43BdD1FGd7P4AOG8rAKCX9vAFbO9G9RVS+c5oQ/pI0m8GLhE
# fEXkwcNyeuBy5yTfv0aZxe/CHFfbg43sTUkwp6uO3+xbn6/83bBm4sGXgXvt1u1L
# 50kppxMopqd9Z4DmimJ4X7IvhNdXnFy/dygo8e1twyiPLI9AN0/B4YVEicQJTMXU
# pUMvdJX3bvh4IFgsE11glZo+TzOE2rCIF96eTvSWsLxGoGyY0uDWiIwLAgMBAAGj
# ggHtMIIB6TAQBgkrBgEEAYI3FQEEAwIBADAdBgNVHQ4EFgQUSG5k5VAF04KqFzc3
# IrVtqMp1ApUwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGG
# MA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAUci06AjGQQ7kUBU7h6qfHMdEj
# iTQwWgYDVR0fBFMwUTBPoE2gS4ZJaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3Br
# aS9jcmwvcHJvZHVjdHMvTWljUm9vQ2VyQXV0MjAxMV8yMDExXzAzXzIyLmNybDBe
# BggrBgEFBQcBAQRSMFAwTgYIKwYBBQUHMAKGQmh0dHA6Ly93d3cubWljcm9zb2Z0
# LmNvbS9wa2kvY2VydHMvTWljUm9vQ2VyQXV0MjAxMV8yMDExXzAzXzIyLmNydDCB
# nwYDVR0gBIGXMIGUMIGRBgkrBgEEAYI3LgMwgYMwPwYIKwYBBQUHAgEWM2h0dHA6
# Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvZG9jcy9wcmltYXJ5Y3BzLmh0bTBA
# BggrBgEFBQcCAjA0HjIgHQBMAGUAZwBhAGwAXwBwAG8AbABpAGMAeQBfAHMAdABh
# AHQAZQBtAGUAbgB0AC4gHTANBgkqhkiG9w0BAQsFAAOCAgEAZ/KGpZjgVHkaLtPY
# dGcimwuWEeFjkplCln3SeQyQwWVfLiw++MNy0W2D/r4/6ArKO79HqaPzadtjvyI1
# pZddZYSQfYtGUFXYDJJ80hpLHPM8QotS0LD9a+M+By4pm+Y9G6XUtR13lDni6WTJ
# RD14eiPzE32mkHSDjfTLJgJGKsKKELukqQUMm+1o+mgulaAqPyprWEljHwlpblqY
# luSD9MCP80Yr3vw70L01724lruWvJ+3Q3fMOr5kol5hNDj0L8giJ1h/DMhji8MUt
# zluetEk5CsYKwsatruWy2dsViFFFWDgycScaf7H0J/jeLDogaZiyWYlobm+nt3TD
# QAUGpgEqKD6CPxNNZgvAs0314Y9/HG8VfUWnduVAKmWjw11SYobDHWM2l4bf2vP4
# 8hahmifhzaWX0O5dY0HjWwechz4GdwbRBrF1HxS+YWG18NzGGwS+30HHDiju3mUv
# 7Jf2oVyW2ADWoUa9WfOXpQlLSBCZgB/QACnFsZulP0V3HjXG0qKin3p6IvpIlR+r
# +0cjgPWe+L9rt0uX4ut1eBrs6jeZeRhL/9azI2h15q/6/IvrC4DqaTuv/DDtBEyO
# 3991bWORPdGdVk5Pv4BXIqF4ETIheu9BCrE/+6jMpF3BoYibV3FWTkhFwELJm3Zb
# CoBIa/15n8G9bW1qyVJzEw16UM0xggSAMIIEfAIBATCBlTB+MQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29k
# ZSBTaWduaW5nIFBDQSAyMDExAhMzAAABh3IXchVZQMcJAAAAAAGHMAkGBSsOAwIa
# BQCggZQwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFCCPe8Uud2qfEa36wXUUqCKa
# t8aQMDQGCisGAQQBgjcCAQwxJjAkoBKAEABUAGUAcwB0AFMAaQBnAG6hDoAMaHR0
# cDovL3Rlc3QgMA0GCSqGSIb3DQEBAQUABIIBAIkBKldSTVuwIViBFiwvxeJsUT8L
# 98D+AduyajEpx1BpVT+IYaNdguKIdP2UU7CIQZO2S8IxrUpmv7N/mtw9i+fNqKEB
# buzplwR3zdoci7uRJKaBJenxBZbZIE3m8ic8AXJmHx20p/EnOH+IWA83QIoh7tDK
# 1IhejKZMA1Ixx1sv+qowzBt11xwHqqOxikmxcMxwPIYxzknWBPdFZMZJDEkkB/m/
# QvjxEo3U6ccp9E6c0nqdqGeaJgXbKM4WPKnKL/0Vmsykne97vaklk/dFV6q9NEH5
# y2Q1fDkcWkExSlX2g2AhuvH+8PVPnhhBJAjLiGm1WnAHmkDbcBarboaNk8ahggIo
# MIICJAYJKoZIhvcNAQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgUENBAhMzAAABVWBhxvjlX4B5AAAAAAFVMAkGBSsOAwIaBQCgXTAYBgkqhkiG
# 9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yMDA1MzAwNDQxNTZa
# MCMGCSqGSIb3DQEJBDEWBBQe2ScTQYDhPHXtmybrVfieFI0lnzANBgkqhkiG9w0B
# AQUFAASCAQBF7P9GqcQXIZilBq5P50NGAmbPZ/zrSXoC/++2F0aZQaWhVSOVJTSs
# v2sl+BKxYeg577LWYJy4Q4YnDjHQYNj5RUBlM4Z+3IcZzuTokqk7CoElihsmZLww
# MR6j8Uiue8c2WYp6smT1QoFlbGMyqfb/flzeU+O+GvXirVFEgtYPZGwF3FO/7s5w
# RQ0fQKJljjlEWmnaCu5wpdkSsPPeLpv9qKlD0LTb/9yjPz0BUf5U6SVO8rX9bbkL
# 663I+7jSD4SKukzbCjwlAL6Y2DFyurDc3GVpREuQOMGhPNJyyc+SfRSuAzGoHMWQ
# M8e6Bk/Hx+qxY92u7LZbxXF4lAXkBpiw
# SIG # End signature block
