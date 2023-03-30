
Add-PowerAppsAccount

$ft = @{Expression={$_.Environment};Label="Environment"}, 
	    @{Expression={$_.App};Label="App"},
        @{Expression={$_.SiteUrl};Label="SiteUrl"},
        @{Expression={$_.List};Label="List"}

$output = @()

$envs = Get-AdminPowerAppEnvironment

foreach ($e in $envs)
{
    $apps = Get-AdminPowerApp -EnvironmentName $e.EnvironmentName -ApiVersion 2021-02-01

    foreach ($a in $apps)
    {
        if ($a.Internal.properties.connectionReferences)
        {
            $connRefs = $a.Internal.properties.connectionReferences | gm -MemberType NoteProperty 

            $spCons = $connRefs | ? { $_.Definition -like "*/apis/shared_sharepointonline*" }

            if ($spCons.Count -gt 0)
            {
                foreach ($spCon in $spCons)
                {
                    $c = $a.Internal.properties.connectionReferences."$($spCon.Name)"  

                    if ($c.dataSets)
                    {
                        $site = $c.dataSets | gm -MemberType NoteProperty 
                        $siteUrl = $site.Name
                    }
                    else
                    {
                        $siteUrl = ""
                    }

                    $out = New-Object PSObject
                    $out | Add-Member -MemberType NoteProperty -Name Environment -Value $e.DisplayName
                    $out | Add-Member -MemberType NoteProperty -Name App -Value $a.DisplayName
                    $out | Add-Member -MemberType NoteProperty -Name SiteUrl -Value $siteUrl
                    $out | Add-Member -MemberType NoteProperty -Name List -Value $c.dataSources

                    $output += $out
                }   
            }  
        }
    }
}

$output | Format-Table $ft -AutoSize



