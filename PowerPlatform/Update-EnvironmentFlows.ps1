
param(
    [Parameter(Mandatory=$false)]
    [string]$EnvironmentID = "Default-d5ff2245-5ff7-4335-9930-08f74f2e8418", #

    [Parameter(Mandatory=$false)]
    [string]$CoEEnvironment = "Power Platform CoE Toolkit" #Personal Productivity"
)

if(-not (Get-Item -Path "$PSScriptRoot\Logs" -ErrorAction Ignore).Exists)
{
    New-Item -Path "$PSScriptRoot\Logs" -ItemType Directory | Out-Null
}

$stopwatch =  [system.diagnostics.stopwatch]::StartNew()


$logFilePath = "$PSScriptRoot\Logs\Flows-$((Get-Date).ToString("ddMMyyyyhhmmss")).log"
$logErrorFilePath = "$PSScriptRoot\Logs\Flows-$((Get-Date).ToString("ddMMyyyyhhmmss"))-errors.log"

New-Item -Path $logFilePath -Force | Out-Null
New-Item -Path $logErrorFilePath -Force | Out-Null


# Login to CoE environment
$Cred = Get-Credential -UserName "peetteri@woodstinen.com" -Message "Creds" # 

Add-PowerAppsAccount -Username $Cred.UserName -Password $Cred.Password
 
# Get CoE environment
$envCOE = Get-AdminPowerAppEnvironment | ? { $_.DisplayName -like "$($CoEEnvironment)*" }

$envDefault = Get-AdminPowerAppEnvironment -Default


$CRMConn = Connect-CrmOnline -Credential $Cred -ForceOAuth -ServerUrl $envCOE.Internal.properties.linkedEnvironmentMetadata.instanceUrl

$existingConRefs = Get-CrmRecords -conn $CRMConn -EntityLogicalName admin_connectionreference -Fields admin_connectionreferenceid -AllRows -WarningAction SilentlyContinue

<#
foreach ($existCF in $existingConRefs.CrmRecords)
{
    Remove-CrmRecord -conn $CRMConn -EntityLogicalName admin_connectionreferenceid -Id $existCF.admin_connectionreferenceid
}
#>

if (![string]::IsNullOrEmpty($EnvironmentID))
{
    # Get environment for processing
    $environment = Get-AdminPowerAppEnvironment -EnvironmentName $EnvironmentID

    # Get all flows in default environment
    $flows = Get-AdminFlow -EnvironmentName $environment.EnvironmentName -FlowName 9c20f983-b78b-4e77-97f5-eeb3035d40f9

    # Get all existing flows from Dataverse
    $records = Get-CrmRecords -AllRows -EntityLogicalName admin_flow -conn $CRMConn -Fields admin_flowid -FilterAttribute admin_flowenvironment -FilterOperator eq -FilterValue $Environment.EnvironmentName.Replace("Default-", "") # d5ff2245-5ff7-4335-9930-08f74f2e8418
}
else
{
    $environment = Get-AdminPowerAppEnvironment 

    # Get all flows in all environments
    $flows = Get-AdminFlow

    # Get all existing flows from Dataverse
    $records = Get-CrmRecords -AllRows -EntityLogicalName admin_flow -conn $CRMConn -Fields admin_flowid 
}



$tenantId = $envDefault.EnvironmentName.Replace("Default-","")

Connect-AzureAD -TenantId $tenantId -Credential $Cred  | Out-Null



foreach ($flow in $flows)
{
    if ($flow.FlowName -ne $null)
    {
        $aadUser = $null
        $resMaker = $null
        $env = $null
        $country = ""
        $city = ""
        $department = ""
        $isOrphan = "no"
        $maker = $null
        $isError = $false
        
        Write-Host "Processing flow: $($flow.DisplayName)" -NoNewline

        try
        {
            $aadUser = Get-AzureADUser -ObjectId $flow.CreatedBy.objectId
        }
        catch
        {}

        if ($aadUser -eq $null) 
        { 
            $isOrphan = "yes" 
        }
        else
        {
            if (-not [string]::IsNullOrEmpty($aadUser.Country)) { $country = $aadUser.Country }
            if (-not [string]::IsNullOrEmpty($aadUser.City)) { $city = $aadUser.City }
            if (-not [string]::IsNullOrEmpty($aadUser.Department)) { $department = $aadUser.Department }
        }
    
        if ($aadUser -ne $null)
        {        
            $resMaker = Get-CrmRecord -EntityLogicalName admin_maker -Fields admin_makerid -Id $aadUser.ObjectId -ErrorAction Ignore

            if ($resMaker -eq $null)
            {
                try
                {
                    New-CrmRecord -conn $CRMConn -EntityLogicalName admin_maker -Id $aadUser.ObjectId -ErrorAction ignore `
                    -Fields @{
                                "admin_makerid"=[guid]::Parse($aadUser.ObjectId);
                                "admin_city"=$city;
                                "admin_country"=$country;
                                "admin_department"=$department;
                                "admin_displayname"=$aadUser.DisplayName;
                                "admin_userprincipalname"=$aadUser.UserPrincipalName
                            }
                }
                catch
                {
                    "MAKER INSERT ERROR (FlowID: $($flow.FlowName)) - Exception: $($_.Exception.Message)" | Out-File -FilePath $logErrorFilePath -Append

                    $isError = $true
                }
            }
            else
            {
                try
                {
                    Update-CrmRecord -conn $CRMConn -EntityLogicalName admin_maker -Id $aadUser.ObjectId -ErrorAction Ignore `
                    -Fields @{
                                "admin_makerid"=[guid]::Parse($aadUser.ObjectId);
                                "admin_city"=$city;
                                "admin_country"=$country;
                                "admin_department"=$department;
                                "admin_displayname"=$aadUser.DisplayName;
                                "admin_userprincipalname"=$aadUser.UserPrincipalName
                            }
                }
                catch
                {
                    "MAKER UPDATE ERROR (FlowID: $($flow.FlowName)) - Exception: $($_.Exception.Message)" | Out-File -FilePath $logErrorFilePath -Append

                    $isError = $true
                }
            }

            try
            {
                # Check does Maker exists and create reference (lookup) if exists
                $resMaker = Get-CrmRecord -conn $CRMConn -EntityLogicalName admin_maker -Id $flow.CreatedBy.userId -Fields ownerid,admin_userprincipalname -ErrorAction Ignore
            }
            catch {}


            if ($resMaker -ne $null)
            {
                $maker = New-CrmEntityReference -EntityLogicalName admin_maker -Id $resMaker.admin_makerid -ErrorAction Ignore
            }
            else
            {
                "ERROR (FlowID: $($flow.FlowName)) - Unable to create Maker reference with ID: $($resMaker.admin_makerid)" | Out-File -FilePath $logErrorFilePath -Append
            
                $isError = $true
            }
        }   


        try
        {
            $env = New-CrmEntityReference -EntityLogicalName admin_environment -id ([guid]::Parse($flow.EnvironmentName.Replace("Default-", "")))
        }
        catch
        {
            "ERROR (FlowID: $($flow.FlowName)) - Unable to create Environment reference with ID: $($environment.EnvironmentName.Replace('Default-', ''))" | Out-File -FilePath $logErrorFilePath -Append
        
            $isError = $true
        }

        $flowExists = $records.CrmRecords | ? { $_.admin_flowid -eq $flow.FlowName}

        if ($flowExists -eq $null)
        {
            try
            {
                New-CrmRecord -conn $CRMConn -EntityLogicalName admin_flow `
                -Fields @{
                    "admin_flowid"=[guid]::Parse($flow.FlowName);
                    "admin_city"=$city;
                    "admin_country"=$country;
                    "admin_department"=$department;
                    "admin_displayname"=$flow.DisplayName;
                    "admin_flowcreatedon"=[datetime]::Parse($flow.CreatedTime);
                    "admin_flowmodifiedon"=[datetime]::Parse($flow.LastModifiedTime);
                    "admin_flowcreatorupn"=$flow.CreatedBy.objectId;
                    "admin_flowstate"=$flow.Internal.properties.state;
                    "admin_flowmakerdisplayname"=if($aadUser -ne $null) { $aadUser.DisplayName };
                    "admin_flowcreator"=if ($maker -ne $null) { $maker };
                    "cr5d5_flowisorphaned"=$isOrphan;
                    "admin_flowenvironment"=$env;
                    "admin_flowenvironmentdisplayname"=$envDefault.DisplayName;
                    "admin_flowenvironmentid"=$flow.EnvironmentName 
                } | out-null

                "INSERT: $($flow.FlowName) - $($flow.DisplayName)" | Out-File -FilePath $logFilePath -Append
            }
            catch
            {
                "INSERT ERROR (FlowID: $($flow.FlowName)) - Exception: $($_.Exception.Message)" | Out-File -FilePath $logErrorFilePath -Append

                $isError = $true
            }
        }
        else
        {
            try
            {
                Update-CrmRecord -conn $CRMConn -EntityLogicalName admin_flow -Id $flow.FlowName `
                -Fields @{
                    "admin_flowid"=[guid]::Parse($flow.FlowName);
                    "admin_city"=$city;
                    "admin_country"=$country;
                    "admin_department"=$department;
                    "admin_displayname"=$flow.DisplayName;
                    "admin_flowcreatedon"=[datetime]::Parse($flow.CreatedTime);
                    "admin_flowmodifiedon"=[datetime]::Parse($flow.LastModifiedTime);
                    "admin_flowcreatorupn"=$flow.CreatedBy.objectId;
                    "admin_flowstate"=$flow.Internal.properties.state;
                    "admin_flowmakerdisplayname"=if($aadUser -ne $null) { $aadUser.DisplayName };
                    "admin_flowcreator"=if ($maker -ne $null) { $maker };
                    "cr5d5_flowisorphaned"=$isOrphan;
                    "admin_flowenvironment"=$env;
                    "admin_flowenvironmentdisplayname"=$envDefault.DisplayName;
                    "admin_flowenvironmentid"=$flow.EnvironmentName 
                } | out-null

                "UPDATE: $($flow.FlowName) - $($flow.DisplayName)" | Out-File -FilePath $logFilePath -Append
            }
            catch
            {
                "UPDATE ERROR (FlowID: $($flow.FlowName)) - Exception: $($_.Exception.Message)" | Out-File -FilePath $logErrorFilePath -Append

                $isError = $true
            }
        }

        if ($isError)
        {
            Write-Host " -> ERROR, see details from error log" -ForegroundColor Red
        }
        else
        {
            Write-Host " -> OK" -ForegroundColor Green
        }

        <#
        # Get flow details for connection references
        $f = Get-AdminFlow -FlowName $flow.FlowName

        $j = ConvertTo-Json -InputObject $f.Internal.properties.connectionReferences

        $conRefs = $f.Internal.properties.connectionReferences | gm -MemberType NoteProperty

        foreach ($cf in $conRefs)
        {
            $params = $cf.Definition.Split(';')

            $tier = $params[$params.Count-1]

            $tier = $tier.Substring($tier.IndexOf('=')+1).replace('}', '')

            Write-Host "$($cf.Name) - $tier"
        }
        #>

        

        Write-Host
    }
}


$stopwatch.Stop()

"" | Out-File -FilePath $logFilePath -Append

"Time taken: " + $stopwatch.Elapsed.Hours + "h " + $stopwatch.Elapsed.Minutes + "m " + $stopwatch.Elapsed.Seconds + "s" | Out-File -FilePath $logFilePath -Append

#>