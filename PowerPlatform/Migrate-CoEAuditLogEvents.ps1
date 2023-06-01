﻿# Requires Microsoft.Xrm.Data.PowerShell module which will be installed if not found

param
(   
    [Parameter(Mandatory=$true)]
    [string]$SourceEnvironmentURL,

    [Parameter(Mandatory=$true)]
    [string]$DestinationEnvironmentURL
)

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

$connSource = Connect-CrmOnline -ServerUrl $SourceEnvironmentUrl -ForceOAuth

$connDest = Connect-CrmOnline -ServerUrl $DestinationEnvironmentUrl -ForceOAuth

$sourceRecords = Get-CrmRecords -conn $connSource -EntityLogicalName admin_auditlog -Fields admin_title,admin_applookup,admin_operation,admin_appid,admin_workload,admin_userupn,admin_creationtime,admin_auditlogid -ErrorAction SilentlyContinue   

write-host
Write-Host "MIGRATING COE AUDIT LOG EVENTS"
Write-Host "Found $($sourceRecords.CrmRecords.Count) events from source AuditLog table in $SourceEnvironmentURL"

$cntMigrated = 0

foreach ($e in $sourceRecords.CrmRecords)
{
    # check does the app exist in Power Apps table. If yes then get reference for the lookup and also update admin_applastlaunchedon
    $rApp = Get-CrmRecord -conn $connDest -EntityLogicalName admin_app -Id $e.admin_appid -Fields admin_appid,admin_applastlaunchedon -ErrorAction SilentlyContinue              
    
    if ($rApp.original -ne $null)
    {
        # Update admin_applastlaunchedon is less than current auditlog entry
        if (($rApp.admin_applastlaunchedon -eq $null) -or ([datetime]$rApp.admin_applastlaunchedon) -lt ([datetime]$e.admin_creationtime))
        {
            # Update App admin_applastlaunchedon info
            Set-CrmRecord -conn $conn -EntityLogicalName admin_app -Id $rApp.admin_appid -Fields @{ "admin_applastlaunchedon" = [datetime]$e.admin_creationtime }
        }

        $lookupObject = $null

        # create lookup object for auditlog item
        if ($rApp.original -ne $null)
        {
            $lookupObject = New-Object -TypeName Microsoft.Xrm.Sdk.EntityReference;
            $lookupObject.LogicalName = "admin_app";
            $lookupObject.Id = $e.admin_appid;
        }

        # Insert only if app exists
        if ($lookupObject -ne $null)
        {
            Update-CrmRecord -conn $conn -EntityLogicalName admin_auditlog -Id $e.admin_auditlogid -Upsert -Fields @{
                "admin_title"= $e.admin_title;
                "admin_applookup"=$lookupObject;
                "admin_operation"=$e.admin_operation;
                "admin_appid"=$e.admin_appid;
                "admin_workload"=$e.admin_workload;
                "admin_userupn"=$e.admin_userupn;
                "admin_creationtime"=[datetime]$e.admin_creationtime;
            } | Out-Null

            $cntMigrated++
        }
    }  
}

write-host "Migrated total $cntMigrated events to $DestinationEnvironmentURL (only apps which exists in CoE PowerApps App table)"
write-host