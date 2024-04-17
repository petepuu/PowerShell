
# Requires Microsoft.Xrm.Data.PowerShell module which will be installed if not found

param
(   
    [Parameter(Mandatory=$true)]
    [string]$SourceEnvironmentURL,

    [Parameter(Mandatory=$true)]
    [string]$DestinationEnvironmentURL,

    # !!!NOTE!!! If you want to migrate ALL events in one run then do NOT set these parameters
    [Parameter(Mandatory=$false)]
    [string]$StartDate, # yyyy-MM-dd

    [Parameter(Mandatory=$false)]
    [string]$EndDate # yyyy-MM-dd
)

# Install Microsoft.Xrm.Data.PowerShell module if not installed already 
$m2 = Get-InstalledModule -Name Microsoft.Xrm.Data.PowerShell -ErrorAction Ignore

if ($null -eq $m2)
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

# Get events from the AuditLog table of the source environment

if ([string]::IsNullOrEmpty($StartDate) -or [string]::IsNullOrEmpty($EndDate))
{
    $sourceRecords = Get-CrmRecords -conn $connSource -AllRows -EntityLogicalName admin_auditlog -Fields admin_title,admin_applookup,admin_operation,admin_appid,admin_workload,admin_userupn,admin_creationtime,admin_auditlogid -ErrorAction SilentlyContinue   
}
else 
{
    $fetchxml = @"
    <fetch version="1.0" distinct="false">
      <entity name="admin_auditlog">
        <attribute name="admin_title" />
        <attribute name="admin_applookup" />
        <attribute name="admin_operation" />
        <attribute name="admin_appid" />
        <attribute name="admin_workload" />
        <attribute name="admin_userupn" />
        <attribute name="admin_creationtime" />
        <attribute name="admin_auditlogid" />
        <filter type="and">
            <condition attribute="admin_creationtime" operator="ge" value="$($StartDate)" />
            <condition attribute="admin_creationtime" operator="le" value="$($EndDate)" />
        </filter>
      </entity>
    </fetch>
"@
    $sourceRecords = Get-CrmRecordsByFetch -conn $connSource -Fetch $fetchxml -ErrorAction SilentlyContinue   
}



write-host
Write-Host "MIGRATING COE AUDIT LOG EVENTS"
Write-Host "Found $($sourceRecords.CrmRecords.Count) events from source AuditLog table in $SourceEnvironmentURL"

$cntMigrated = 0

foreach ($e in $sourceRecords.CrmRecords)
{
    # Check does the app exist in Power Apps table
    $rApp = Get-CrmRecord -conn $connDest -EntityLogicalName admin_app -Id $e.admin_appid -Fields admin_appid,admin_applastlaunchedon -ErrorAction SilentlyContinue              
    
    # !!!! Only process if app is found in admin_app table !!!!
    if ($null -ne $rApp.original)
    {
        # Update admin_applastlaunchedon is less than current auditlog entry or NULL
        if (($null -eq $rApp.admin_applastlaunchedon) -or ([datetime]$rApp.admin_applastlaunchedon) -lt ([datetime]$e.admin_creationtime))
        {
            # Update App admin_applastlaunchedon info
            Set-CrmRecord -conn $conn -EntityLogicalName admin_app -Id $rApp.admin_appid -Fields @{ "admin_applastlaunchedon" = [datetime]$e.admin_creationtime }
        }

        $lookupObject = $null

        # create lookup object for auditlog item
        if ($null -ne $rApp.original)
        {
            $lookupObject = New-Object -TypeName Microsoft.Xrm.Sdk.EntityReference;
            $lookupObject.LogicalName = "admin_app";
            $lookupObject.Id = $e.admin_appid;
        }

        # Insert only if app exists
        if ($null -ne $lookupObject)
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

write-host
write-host "Migrated total $cntMigrated events to $DestinationEnvironmentURL (only apps which exists in CoE PowerApps App table)"
write-host
