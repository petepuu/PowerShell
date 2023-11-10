
# NOTE: Requires PAC CLI installed https://learn.microsoft.com/en-us/power-platform/developer/cli/introduction#install-power-platform-cli-for-windows

param(
    [Parameter(Mandatory=$false)]
    [string]$EnvironmentUrl = "", # E.g. https://org12345678.crm.dynamics.com/

    [Parameter(Mandatory=$false)]
    [string]$CoECorePackagePath = "", # e.g. C:\temp\CoECore\CenterofExcellenceCoreComponents_4.17_managed.zip

    # msedge, chrome, firefox
    [Parameter(Mandatory=$false)]
    [ValidateSet("msedge","chrome","firefox")]
    [string]$Browser = "firefox"
)

$WorkingDirectory = Split-Path -Parent $CoECorePackagePath

$CoECorePackage = (Get-ChildItem -path $CoECorePackagePath).Name

$global:CoEConnections = @("shared_dataflows ","shared_commondataservice ","shared_commondataserviceforapps ","shared_flowmanagement ","shared_microsoftflowforadmins ","shared_office365 ","shared_office365groups ","shared_office365users ","shared_powerappsforadmins ","shared_powerappsforappmakers ","shared_powerplatformforadmins ","shared_rss ","shared_teams ","shared_webcontents ")

# Check are CoE Core Component connections created
function CheckCoEConnections($env)
{
    $global:Connections = pac connection list -env $env.EnvironmentUrl

    $allConnectionsOK = $true

    foreach ($coeCon in $global:CoEConnections)
    {
        if ($global:Connections | ? { $_ -like "*$coeCon*" -and $_ -like "*Connected" })
        {
            write-host "$coeCon - " -NoNewline
            Write-Host "OK" -ForegroundColor Green 
        }
        else
        {
            write-host "$coeCon - " -NoNewline
            Write-Host "NOT FOUND" -ForegroundColor Red

            $allConnectionsOK = $false
        }
    }

    return $allConnectionsOK
}

function CreateMissingConnections($env)
{
    $global:Connections = pac connection list -env $env.EnvironmentUrl

    $allConnectionsOK = $true

    foreach ($coeCon in $global:CoEConnections)
    {
        if (!($global:Connections | ? { $_ -like "*$coeCon*" -and $_ -like "*Connected" }))
        {
            # Open browser for creating CoE connections
            [system.Diagnostics.Process]::Start($Browser, "https://make.powerapps.com/environments/$($env.EnvironmentId)/connections/available?apiName=$coeCon") | Out-Null
        }
    }
}

$stopwatch =  [system.diagnostics.stopwatch]::StartNew()

$global:Connections = $null

[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

$Credentials = Get-Credential -Message "Provide CoE Service Account credentials having Power Platform Administrator role"

# PAC AUTH
pac auth create -u $EnvironmentUrl -un $Credentials.UserName -p $Credentials.GetNetworkCredential().Password

# Get CoE environment
$env = (pac admin list -env $EnvironmentUrl --json | convertfrom-json)

write-host

while(!(CheckCoEConnections $env))
{
    write-host

    Read-Host -Prompt "Did not find all required CoE Core Component connections. Click any key to open connection creation page in selected browser"

    CreateMissingConnections $env

    Read-Host -Prompt "Waiting for connections created. When ready, click any key to continue"
} 

Write-host

Write-Host "Installing CoE Core Components... "

pac solution create-settings --solution-zip "$WorkingDirectory\$CoECorePackage" --settings-file "$WorkingDirectory\CenterofExcellenceCoreComponents_settings.json"

$settings = Get-Content -Path "$WorkingDirectory\CenterofExcellenceCoreComponents_settings.json" | ConvertFrom-Json
#>
foreach ($cf in $settings.ConnectionReferences)
{
    $cType = $cf.ConnectorId.Substring($cf.ConnectorId.LastIndexOf('/')+1) + " "

    $c = $global:Connections | ? { $_ -like "*$cType*" }

    if ($c -ne $null)
    {
        if ($c.Count -gt 1)
        {
            $cf.ConnectionId = $c[0].Substring(0, $c[0].IndexOf(" "))
        }
        else 
        {
            $cf.ConnectionId = $c.Substring(0, $c.IndexOf(" "))
        }
    }
}

# clear environment variables from the settings
$settings.EnvironmentVariables = @()

$settings | ConvertTo-Json | Out-File "$WorkingDirectory\CenterofExcellenceCoreComponents_settings.json"

pac solution import -p "$WorkingDirectory\$CoECorePackage" -wt 45 -a -ap --settings-file "$WorkingDirectory\CenterofExcellenceCoreComponents_settings.json"


$stopwatch.Stop()

$e = $stopwatch.Elapsed

write-host "Script execution time: $($e.Hours)h $($e.Minutes)m $($e.Seconds)s" -ForegroundColor Green

Write-Host
