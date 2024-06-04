
# Required Microsoft.Xrm.Data.Powershell module

# !!! NOTE !!! Ensure that before running the script you have set the correct home domains in 'Host Domains' environment value. 
# It need to be comma-separated list of FQDNs like contoso.com,fabrikam.com

param
(    
    # Get environment URL https://org??????.crm.dynamics.com from PPAC for example 
    [Parameter(Mandatory=$true)]
    [string]$EnvironmentURL
)

Import-Module Microsoft.Xrm.Data.Powershell

$hostDomains = $null

Connect-CrmOnline -ServerUrl $EnvironmentURL -ForceOAuth | Out-Null

$envVar = Get-CrmRecords -TopCount 1 -EntityLogicalName environmentvariabledefinition -FilterAttribute schemaname -FilterOperator eq -FilterValue admin_HostDomains

if ($envVar.CrmRecords.Count -eq 1)
{
    $envVarValue = Get-CrmRecords -TopCount 1 -EntityLogicalName environmentvariablevalue -Fields * -FilterAttribute environmentvariabledefinitionid -FilterOperator eq -FilterValue $envVar.CrmRecords[0].environmentvariabledefinitionid

    if ($envVarValue.CrmRecords.Count -gt 0)
    {
        $value = $envVarValue.CrmRecords[0].value

        $hostDomains = $value.split(",")
    }
    else 
    {
        Write-Host "Host Domains environment variable not set!" -ForegroundColor Red
        break
    }
}

if ($hostDomains.Count -gt 0)
{
    $conns = Get-CrmRecords -EntityLogicalName admin_connectionreferenceidentity -Fields admin_noneorcrosstenantidentity,admin_accountname -AllRows

    $records = $conns.CrmRecords

    foreach ($r in $records)
    {
        $acct = $r.admin_accountname

        $domain = $acct.Substring($acct.IndexOf("@")+1)

        if ($r.admin_noneorcrosstenantidentity -eq "Yes" -and $domain -in $hostDomains)
        {
            Update-CrmRecord -EntityLogicalName admin_connectionreferenceidentity -Id $r.admin_connectionreferenceidentityid -Fields @{ "admin_noneorcrosstenantidentity" = $false } 
            
            Write-Host "$($acct) set to NO"
        }
        elseif($r.admin_noneorcrosstenantidentity -eq "No" -and $domain -notin $hostDomains) 
        {
            Update-CrmRecord -EntityLogicalName admin_connectionreferenceidentity -Id $r.admin_connectionreferenceidentityid -Fields @{ "admin_noneorcrosstenantidentity" = $true }

            Write-Host "$($acct) set to YES"
        }
    }
}
