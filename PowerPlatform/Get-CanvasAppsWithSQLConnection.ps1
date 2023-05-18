
# Script lists all canvas apps in every environment having SQL connection with SQL server, database and connected tables (datasources) info
# Output is saved to CSV file into same directory where script is located

Add-PowerAppsAccount

$apps = $null
$output = @()

$envs = Get-AdminPowerAppEnvironment

Write-Host "Canvas apps having SQL Server connection(s)"

foreach ($env in $envs)
{
    $apps = Get-AdminPowerApp -EnvironmentName $env.EnvironmentName

    foreach ($a in $apps)
    {  
        if ($a.Internal.properties.connectionReferences -ne $null)
        {
            $connRefs = $a.Internal.properties.connectionReferences | gm -MemberType NoteProperty

            $sqls = $connRefs | ? { $_.Definition -like "*id=/providers/Microsoft.PowerApps/scopes/admin/apis/shared_sql*" }
        
            $cnt = Measure-Object -InputObject $sqls

            if ($cnt.Count -gt 0)
            {
                write-host "APP:'$($a.DisplayName)' ($($a.AppName)) - ENV:$($a.EnvironmentName)"

                foreach ($sql in $sqls)
                {
                    $c = $a.Internal.properties.connectionReferences."$($sql.Name)"

                    $isImplicit = "False"

                    if ($c.authenticationType -ne $null)
                    {
                        if ($c.authenticationType -notin ("oAuth", "windowsAuthenticationNonShared"))
                        {
                            $isImplicit = "True"
                        }
                    }
                    
                    $dataSources = ""

                    if ($c.dataSources -ne $null)
                    {
                        $dataSources = [system.String]::Join(",", $c.dataSources)
                    }

                    $endpoints = ""

                    if ($c.endpoints -ne $null)
                    {
                        $endpoints = [system.String]::Join(",", $c.endpoints)
                    }

                    $out = New-Object PSObject
                    $out | Add-Member -MemberType NoteProperty -Name AppId -Value $a.AppName
                    $out | Add-Member -MemberType NoteProperty -Name DisplayName -Value $a.DisplayName
                    $out | Add-Member -MemberType NoteProperty -Name AppEnvId -Value $a.EnvironmentName
                    $out | Add-Member -MemberType NoteProperty -Name DataSources -Value $dataSources
                    $out | Add-Member -MemberType NoteProperty -Name Endpoints -Value $endpoints
                    $out | Add-Member -MemberType NoteProperty -Name AuthenticationType -Value $c.authenticationType
                    $out | Add-Member -MemberType NoteProperty -Name IsImplicit -Value $isImplicit
                    $out | Add-Member -MemberType NoteProperty -Name IsOnPrem -Value $c.isOnPremiseConnection
                    $out | Add-Member -MemberType NoteProperty -Name AppSharedusers -Value $a.Internal.properties.sharedUsersCount
                    $out | Add-Member -MemberType NoteProperty -Name AppSharedGroups -Value $a.Internal.properties.sharedGroupsCount

                    $output  += $out
                }
            }  
        }    
    }
}

$output | Export-Csv -NoTypeInformation $PSScriptRoot\appswithsqlconnections.csv -Force

Write-Host
Write-Host "Output saved to $("$PSScriptRoot\appswithsqlconnections.csv")"
Write-Host
