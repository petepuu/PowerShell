
# Uses EXO cmdlet Search-UnifiedAuditLog to get LaunchPowerApp events from Audit Log
# You need to have either Global or EXO Admin permissions and also permission to update CoE AuditLog table to run this 

# NOTE!!! This might not work in large environments as is because paging does not work correctly
# At the moment runs one query per 24h between defined start and end dates or using "dummy" paging on every 3h of the day if 24h query returns 5000 items
# CMDLET is using the recommended "-SessionCommand ReturnLargeSet" which should support paging up to 50000 returned events 
# but for some reason I was not able to get that command to return more than 5000 events i.e. was not able to get paging working


# Updates to CoE DB are done using Microsoft.Xrm.Data.PowerShell module commands

param
(   
    [Parameter(Mandatory=$false)]
    [datetime]$StartDate = (Get-Date).AddDays(-90), # Start date of the queries
    
    [Parameter(Mandatory=$false)]
    [datetime]$EndDate = (Get-Date).AddDays(-1), # End date until which we want fetch events

    [Parameter(Mandatory=$false)]
    [boolean]$UpdateCoE = $false,   # $false = only output the number of LaunchPowerApp events for given timeframe
                                    # $true = also save them to CoE AuditLog table

    [Parameter(Mandatory=$false)]
    [string]$EnvironmentUrl = "https://????.crm?.dynamics.com/" # Need to be set IF $UpdateCoE = $true
    
)


function UpdateCoE($entries)
{
    Write-Host " - Saving to AuditLog table"

    foreach ($e in $entries)
    {
        $entry = ConvertFrom-Json $e.AuditData
  
        $id = [guid]$entry.Id

        # check does the app exist in Power Apps table. If yes then get reference for the lookup and also update admin_applastlaunchedon
        $rApp = Get-CrmRecord -conn $conn -EntityLogicalName admin_app -Id $entry.AppName -Fields admin_appid,admin_applastlaunchedon -ErrorAction SilentlyContinue              

        if ($rApp.original -ne $null)
        {
            if ($rApp.admin_applastlaunchedon -eq $null)
            {
                # Set App admin_applastlaunchedon info
                Set-CrmRecord -conn $conn -EntityLogicalName admin_app -Id $entry.AppName -Fields @{ "admin_applastlaunchedon" = [datetime]$entry.CreationTime }
            }
            else
            {
                # Update admin_applastlaunchedon is less than current auditlog entry
                if (([datetime]$rApp.admin_applastlaunchedon) -lt ([datetime]$entry.CreationTime))
                {
                    # Update App admin_applastlaunchedon info
                    Set-CrmRecord -conn $conn -EntityLogicalName admin_app -Id $entry.AppName -Fields @{ "admin_applastlaunchedon" = [datetime]$entry.CreationTime }
                }
            }
        }

        # check does the event already exist in Audit Log table. If yes then skip to next
        $r = Get-CrmRecord -conn $conn -EntityLogicalName admin_auditlog -Id $id -Fields admin_auditlogid -ErrorAction SilentlyContinue
            
        if ($r.original -eq $null)
        {
            #Write-Host "." -NoNewline

            $lookupObject = $null

            # create lookup object for auditlog item
            if ($rApp.original -ne $null)
            {
                $lookupObject = New-Object -TypeName Microsoft.Xrm.Sdk.EntityReference;
                $lookupObject.LogicalName = "admin_app";
                $lookupObject.Id = $entry.AppName;
            }

            if ($lookupObject -ne $null)
            {
                New-CrmRecord -conn $conn -EntityLogicalName admin_auditlog -Fields @{
                    "admin_title"="Power App Launch $($entry.AppName) - $($entry.CreationTime)";
                    "admin_applookup"=$lookupObject;
                    "admin_operation"=$entry.Operation;
                    "admin_appid"=$entry.AppName;
                    "admin_workload"=$entry.Workload;
                    "admin_userupn"=$entry.UserId;
                    "admin_creationtime"=[datetime]$entry.CreationTime;
                    "admin_auditlogid"=$id 
                } | Out-Null
            }
            else
            {
                New-CrmRecord -conn $conn -EntityLogicalName admin_auditlog -Fields @{
                    "admin_title"="Power App Launch $($entry.AppName) - $($entry.CreationTime)";
                    "admin_operation"=$entry.Operation;
                    "admin_appid"=$entry.AppName;
                    "admin_workload"=$entry.Workload;
                    "admin_userupn"=$entry.UserId;
                    "admin_creationtime"=[datetime]$entry.CreationTime;
                    "admin_auditlogid"=$id
                } | Out-Null
            }
        }
  
    }
}

$currentDate = Get-Date

# Install ExchangeOnlineManagement module if not installed already 
$m = Get-InstalledModule -Name ExchangeOnlineManagement -MinimumVersion 3.0.0 -ErrorAction Ignore

if ($m -eq $null)
{
    $title    = 'Install ExchangeOnlineManagement PowerShell Module'
    $question = 'Unable to find ExchangeOnlineManagement PowerShell module version 3.0.0 which is needed for this script to run. Do you want to install this module?'
    $choices  = '&Yes', '&No'

    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)

    if ($decision -eq 0) 
    {
        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Confirm:$false
    } 
    else 
    {
        Exit
    }
}

if ($UpdateCoE)
{
    # Install Microsoft.Xrm.Data.PowerShell module if not installed already 
    $m2 = Get-InstalledModule -Name Microsoft.Xrm.Data.PowerShell -ErrorAction Ignore

    if ($m2 -eq $null)
    {
        $title    = 'Install Microsoft.Xrm.Data.PowerShell Module'
        $question = 'Unable to find Microsoft.Xrm.Data.PowerShell module version which is needed for this script to run. Do you want to install this module?'
        $choices  = '&Yes', '&No'

        $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)

        if ($decision -eq 0) 
        {
            Install-Module -Name Microsoft.Xrm.Data.PowerShell -Force -Confirm:$false

            Import-Module Microsoft.Xrm.Data.PowerShell -Force
        } 
        else 
        {
            Exit
        }
    }
}


Connect-ExchangeOnline

if ($UpdateCoE)
{
    $conn = Connect-CrmOnline -ServerUrl $EnvironmentUrl -ForceOAuth
}

$tmpEntries = $null

$StartDate = $StartDate.AddHours(-$StartDate.Hour).AddMinutes(-$StartDate.Minute).AddSeconds(-$StartDate.Second)
        
do
{
    $entries = $null

    $entries = Search-UnifiedAuditLog -StartDate $StartDate.ToLocalTime() -EndDate $StartDate.AddDays(1).ToLocalTime() -SessionCommand ReturnLargeSet -ResultSize 5000 -RecordType PowerAppsApp -Operations LaunchPowerApp

    $cnt = @($entries).Count
    
   
    if ($cnt -eq 5000)
    {
        write-host "Count of daily event entries more than max 5000. Will run query for $($StartDate) in 3 hour slots" -ForegroundColor Red

        $tmpStartDate = $StartDate

        for ($hh=3; $hh -le 24; $hh=$hh+3)
        {
            $cnt = 0

            $entries = Search-UnifiedAuditLog -StartDate $tmpStartDate.ToLocalTime() -EndDate $tmpStartDate.AddHours(3).ToLocalTime() -SessionCommand ReturnLargeSet -ResultSize 5000 -RecordType PowerAppsApp -Operations LaunchPowerApp

            $cnt = @($entries).Count

            write-host "Found $($cnt) events for StartTime: $($tmpStartDate.ToString("MM/dd/yyyy HH:mm:ss")) - EndTime: $($tmpStartDate.AddHours(3).ToString("MM/dd/yyyy HH:mm:ss"))" -NoNewline

            $tmpStartDate = $StartDate.AddHours($hh)

            $tmpEntries += $cnt

            if ($UpdateCoE -and $cnt -gt 0)
            {
                UpdateCoE $entries     
            }
            else
            {
                Write-Host
            }
        }
    }
    else
    {   
        write-host "Found $($cnt) events for StartTime: $($StartDate.ToString("MM/dd/yyyy HH:mm:ss")) - EndTime: $($StartDate.AddDays(1).ToString("MM/dd/yyyy HH:mm:ss"))" -NoNewline
               
        $tmpEntries += $cnt

        if ($UpdateCoE -and $cnt -gt 0)
        {
            UpdateCoE $entries 
        }
        else
        {
            Write-Host
        }
    }        
                

    $StartDate = $StartDate.AddDays(1) 

} Until($StartDate -ge $EndDate) 

write-host    
write-host "Processed total $($tmpEntries) events"