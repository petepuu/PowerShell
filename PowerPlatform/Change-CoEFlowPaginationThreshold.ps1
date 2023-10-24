param
(
    [Parameter(Mandatory=$false)]
    [string]$TenantName = ".onmicrosoft.com",
   
    [Parameter(Mandatory=$false)]
    [string]$EnvironmentUrl = "", # https://???????.crm.dynamics.com/
   
    [Parameter(Mandatory=$false)]
    [string]$SolutionName = "CenterofExcellenceCoreComponents"
)

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
    "Accept" = "application/json"
}

$resSolution = Invoke-RestMethod -Uri "$($EnvironmentUrl)api/data/v9.2/solutions?`$filter=(uniquename eq '$SolutionName')" -Method Get -Headers $authHeader

$solution = $resSolution.value

$resObjects = Invoke-RestMethod -Uri "$($EnvironmentUrl)api/data/v9.2/msdyn_solutioncomponentsummaries?`$filter=(msdyn_solutionid eq $($solution.solutionid) and msdyn_componenttype eq 29)" -Method Get -Headers $authHeader

$objects = $resObjects.value | sort msdyn_componenttype

foreach($o in $objects)
{
    $reqFlow = "$($EnvironmentUrl)api/data/v9.2/workflows($($o.msdyn_objectid))"
    $resFlow = Invoke-RestMethod -Uri $reqFlow -Headers $authHeader -Method Get  

    $clientData = $resFlow.clientdata

    if ($clientData -like "*`"minimumItemCount`":10000*")
    {
        if ($resFlow.name -ne "Admin | Sync Template v3 CoE Solution Metadata")
        {
            if ($resFlow.name -like "Admin | Sync Template*")
            {
                Write-Host $resFlow.name

                $clientData = $clientData.Replace("`"minimumItemCount`":100000", "`"minimumItemCount`":5000")

                $clientData = $clientData | ConvertTo-Json -Depth 100

                $b = '{
                    "statecode": 1,
                    "clientdata": ' + $clientData + '
                }'

                $res = Invoke-RestMethod -Uri "$($EnvironmentUrl)api/data/v9.2/workflows($($o.msdyn_objectid))" -Method Patch -Headers $authHeader -Body $b
            }
        }
    }
}