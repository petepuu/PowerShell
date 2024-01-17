
# Requires MSAL.PS module for authentication
# https://www.powershellgallery.com/packages/MSAL.PS

# NOTE! Client ID used in the API call is the sample Client ID not meant for production use
# See more: https://learn.microsoft.com/en-us/power-apps/developer/data-platform/xrm-tooling/use-connection-strings-xrm-tooling-connect#connection-string-parameters


param
(
    [Parameter(Mandatory=$false)]
    [string]$TenantName = "????????.onmicrosoft.com", 
    
    [Parameter(Mandatory=$false)]
    [string]$EnvironmentURL = "https://????????.crm?.dynamics.com/", 
    
    [Parameter(Mandatory=$false)]
    [string]$SolutionName = "", # Solution internal name e.g. CenterofExcellenceCoreComponents

    [Parameter(Mandatory=$false)]
    [string]$NewOwnerUPN = ""
)

$connectionDetails = @{
    'TenantId'     = $TenantName
    'ClientId'     = '51f81489-12ee-4a9e-aaae-a2591f45987d' # This Client ID is meant only for development and testing purposes. See more: https://learn.microsoft.com/en-us/power-apps/developer/data-platform/xrm-tooling/use-connection-strings-xrm-tooling-connect#connection-string-parameters
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


# get new owneer from systemusers
$resNewOwner = Invoke-RestMethod -Uri "$($EnvironmentUrl)api/data/v9.2/systemusers?`$filter=(domainname eq '$NewOwnerUPN')" -Method Get -Headers $authHeader

# if owner exists then continue
if ($resNewOwner.value.Count -eq 1)
{
    $newOwner = $resNewOwner.value[0]

    $resSolution = Invoke-RestMethod -Uri "$($EnvironmentUrl)api/data/v9.2/solutions?`$filter=(uniquename eq '$SolutionName')" -Method Get -Headers $authHeader

    $solution = $resSolution.value

    $resObjects = Invoke-RestMethod -Uri "$($EnvironmentUrl)api/data/v9.2/msdyn_solutioncomponentsummaries?`$filter=(msdyn_solutionid eq $($solution.solutionid) and msdyn_componenttypename eq 'Connection Reference')" -Method Get -Headers $authHeader

    $objects = $resObjects.value
   
    foreach ($o in $objects)
    {
        Write-Host "'$($o.msdyn_displayname)'... " -NoNewline

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
                                                <a:LogicalName>connectionreference</a:LogicalName>
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
}
