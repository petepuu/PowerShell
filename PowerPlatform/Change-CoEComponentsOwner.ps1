
# Requires MSAL.PS and Power Platform Admin PowerShell modules 
# https://www.powershellgallery.com/packages/Microsoft.PowerApps.Administration.PowerShell
# https://www.powershellgallery.com/packages/MSAL.PS

# Run script using current CoE service account

param
(
    [Parameter(Mandatory=$false)]
    [string]$TenantName = ".onmicrosoft.com", 
    
    [Parameter(Mandatory=$false)]
    [string]$EnvironmentID = "", # GUID
    
    [Parameter(Mandatory=$false)]
    [string]$SolutionName = "CenterofExcellenceCoreComponents",

    [Parameter(Mandatory=$false)]
    [string]$NewOwnerUPN = "newaccount@contoso.com"
)

Import-Module MSAL.PS -ea 0
Import-Module Microsoft.PowerApps.Administration.PowerShell -ea 0

Add-PowerAppsAccount

$environment = Get-AdminPowerAppEnvironment -EnvironmentName $EnvironmentID

$EnvironmentUrl = $environment.Internal.properties.linkedEnvironmentMetadata.instanceUrl

$connectionDetails = @{
    'TenantId'     = $TenantName
    'ClientId'     = '51f81489-12ee-4a9e-aaae-a2591f45987d'
    'Interactive'  = $true
    'RedirectUri'  = 'https://localhost'
    'Scopes'    = $EnvironmentUrl + '.default'
}

$t = Get-MsalToken @connectionDetails

$authHeader = @{
    "Authorization" = $t.CreateAuthorizationHeader()
    "Content-type" = "application/json"
}

$authHeaderSOAP = @{
            "Authorization" = $t.CreateAuthorizationHeader()
            "Content-type" = "text/xml; charset=UTF-8"
            "SOAPAction" = "http://schemas.microsoft.com/xrm/2011/Contracts/Services/IOrganizationService/Execute"
        }


# get nw owner from systemusers
$resNewOwner = Invoke-RestMethod -Uri "$($EnvironmentUrl)api/data/v9.2/systemusers?`$filter=(domainname eq '$NewOwnerUPN')" -Method Get -Headers $authHeader

# if owner exists then continue
if ($resNewOwner.value.Count -eq 1)
{
    $newOwner = $resNewOwner.value[0]

    $resSolution = Invoke-RestMethod -Uri "$($EnvironmentUrl)api/data/v9.2/solutions?`$filter=(uniquename eq '$SolutionName')" -Method Get -Headers $authHeader

    $solution = $resSolution.value
    
    $resObjects = Invoke-RestMethod -Uri "$($EnvironmentUrl)api/data/v9.2/msdyn_solutioncomponentsummaries?`$filter=(msdyn_solutionid eq $($solution.solutionid) and (msdyn_componentlogicalname eq 'connectionreference' or msdyn_componentlogicalname eq 'workflow' or msdyn_componentlogicalname eq 'environmentvariabledefinition'))" -Method Get -Headers $authHeader 
    
    $objects = $resObjects.value | sort msdyn_componenttype
    
    foreach ($o in $objects)
    {
        #$objectType = ""

        Write-Host "'$($o.msdyn_displayname)'... " -NoNewline

        if ($o.msdyn_componentlogicalname -eq "workflow")
        {
            #$objectType = "workflow"

            # Current owner
            $reqFlowOwnerID = "$($EnvironmentUrl)api/data/v9.2/workflows($($o.msdyn_objectid))?`$select=_ownerid_value"
            $resFlowOwnerID = Invoke-RestMethod -Uri $reqFlowOwnerID -Headers $authHeader -Method Get

            
            $resCurrentOwnerUPN = Invoke-RestMethod -Uri "$($EnvironmentUrl)api/data/v9.2/systemusers($($resFlowOwnerID._ownerid_value))?`$select=domainname" -Method Get -Headers $authHeader
            $currentOwnerUPN = $resCurrentOwnerUPN.domainname
        }
        <# Custom Connector 
        elseif ($o.msdyn_componenttype -eq 372)
        {
            $objectType = "connector"
        }
        
        elseif ($o.msdyn_componenttype -eq 10088)
        {
            $objectType = "connectionreference"
        }
        elseif ($o.msdyn_componenttype -eq 380)
        {
            $objectType = "environmentvariabledefinition"
        }
        #>

        try
        {
            # Request body for the change owner web service call
            $xmlOwner = '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
                        <s:Header>
                            <SdkClientVersion xmlns="http://schemas.microsoft.com/xrm/2011/Contracts">9.0</SdkClientVersion>
                        </s:Header>
                        <s:Body>
                            <Execute xmlns="http://schemas.microsoft.com/xrm/2011/Contracts/Services"
                                xmlns:i="http://www.w3.org/2001/XMLSchema-instance">
                                <request i:type="b:AssignRequest" xmlns:a="http://schemas.microsoft.com/xrm/2011/Contracts"
                                    xmlns:b="http://schemas.microsoft.com/crm/2011/Contracts">
                                    <a:Parameters xmlns:b="http://schemas.datacontract.org/2004/07/System.Collections.Generic">
                                        <a:KeyValuePairOfstringanyType>
                                            <b:key>Target</b:key>
                                            <b:value i:type="a:EntityReference">
                                                <a:Id>' + $o.msdyn_objectid + '</a:Id> 
                                                <a:LogicalName>' + $o.msdyn_componentlogicalname + '</a:LogicalName>
                                            </b:value>
                                        </a:KeyValuePairOfstringanyType>
                                        <a:KeyValuePairOfstringanyType>
                                            <b:key>Assignee</b:key>
                                            <b:value i:type="a:EntityReference">
                                                <a:Id>' + $newOwner.systemuserid + '</a:Id>
                                                <a:LogicalName>systemuser</a:LogicalName>
                                            </b:value>
                                        </a:KeyValuePairOfstringanyType>
                                    </a:Parameters>
                                    <a:RequestId i:nil="true" />
                                    <a:RequestName>Assign</a:RequestName>
                                </request>
                            </Execute>
                        </s:Body>
                    </s:Envelope>'


            # Change the owner using 'XRMServices/2011/Organization.svc' web service
            $resNewOwner = Invoke-RestMethod -Uri "$($EnvironmentUrl)XRMServices/2011/Organization.svc/web" -Method Post -Headers $authHeaderSOAP -Body $xmlOwner
    

            if ($resNewOwner.Envelope.Body.ExecuteResponse.ExecuteResult.ResponseName -eq "Assign")
            {
                # If modern flow then remove old owner permissions
                if ($o.msdyn_componenttype -eq 29 -and $o.msdyn_workflowcategory -eq 5) 
                {
                    $bodyRevoke = '{
                        "Target": {
                            "workflowid": "' + $o.msdyn_objectid + '",
                            "@odata.type": "Microsoft.Dynamics.CRM.workflow"
                        },
                        "Revokee": {
                            "ownerid": "' + $resFlowOwnerID._ownerid_value + '",
                            "@odata.type": "Microsoft.Dynamics.CRM.systemuser"
                        }
                    }'

                    # Remove old account permissions
                    $reqRevoke = "$($EnvironmentUrl)api/data/v9.2/RevokeAccess"
                    $resRevoke = Invoke-RestMethod -Uri $reqRevoke -Headers $authHeader -Method Post -Body $bodyRevoke -ErrorAction SilentlyContinue
                }

                Write-Host "UPDATED" -ForegroundColor Green
            }
        }
        catch
        { 
            write-host

            $ResponseBody = "UNKNOWN"

            $resError = $_.Exception.Response

            $Reader = New-Object System.IO.StreamReader($resError.GetResponseStream())
            $Reader.BaseStream.Position = 0
            $Reader.DiscardBufferedData()
            $ResponseBody = $Reader.ReadToEnd()

            if ($ResponseBody.StartsWith('{')) 
            {
                $ResponseBody = $ResponseBody | ConvertFrom-Json

                Write-Host "ERROR - $($ResponseBody.error.message)" -ForegroundColor Red
            }
            elseif ($ResponseBody.StartsWith('<s:Envelope'))
            {
                $ResponseBody = [xml]$ResponseBody

                Write-Host "ERROR - $($ResponseBody.Envelope.Body.Fault.faultstring.InnerText)" -ForegroundColor Red
            }
        }        

        write-host
    }
    
    
    Write-Host
    Write-Host

    Write-Host "UPDATING CANVAS APPS"
    Write-Host

    $resCanvasApps = Invoke-RestMethod -Uri "$($EnvironmentUrl)api/data/v9.2/msdyn_solutioncomponentsummaries?`$filter=(msdyn_solutionid eq $($solution.solutionid) and msdyn_componenttype eq 300)" -Method Get -Headers $authHeader

    foreach ($app in $resCanvasApps.value) 
    {
        write-host "'$($app.msdyn_displayname)'... " -NoNewline

        $res = Set-AdminPowerAppOwner -AppName $app.msdyn_objectid -EnvironmentName $EnvironmentID -AppOwner $newOwner.azureactivedirectoryobjectid

        if ($res.Error -eq $null)
        {
            Write-Host "UPDATED" -ForegroundColor Green 
        }

        Write-Host
    }
    
}
