
Add-PowerAppsAccount

$ft = @{Expression={$_.Environment};Label="Environment"}, 
	    @{Expression={$_.Flow};Label="Flow"},
        @{Expression={$_.TriggerAction};Label="TriggerAction"},
        @{Expression={$_.SiteUrl};Label="SiteUrl"},
        @{Expression={$_.ListId};Label="ListId"},
        @{Expression={$_.Operation};Label="Operation"}

$output = @()

$envs = Get-AdminPowerAppEnvironment

foreach ($e in $envs)
{
    $flows = Get-AdminFlow -EnvironmentName $e.EnvironmentName

    foreach ($f in $flows)
    {
        $flow = Get-AdminFlow -EnvironmentName $e.EnvironmentName -FlowName $f.FlowName 

        $spTriggersActions = $flow.Internal.properties.referencedResources | ? { $_.service -eq "sharepoint" }
    
        $cnt = ($spTriggersActions | measure).Count

        if ($cnt.Count -gt 0)
        {
            foreach ($ta in $spTriggersActions)
            {
                $referencers = $ta.referencers

                foreach ($ref in $referencers)
                {
                    $out = New-Object PSObject
                    $out | Add-Member -MemberType NoteProperty -Name Environment -Value $e.DisplayName
                    $out | Add-Member -MemberType NoteProperty -Name Flow -Value $flow.DisplayName
                    $out | Add-Member -MemberType NoteProperty -Name TriggerAction -Value $ref.referenceSourceType
                    $out | Add-Member -MemberType NoteProperty -Name SiteUrl -Value $ta.resource.site
                    $out | Add-Member -MemberType NoteProperty -Name ListId -Value $ta.resource.list
                    $out | Add-Member -MemberType NoteProperty -Name Operation -Value $ref.operationId.Substring($ref.operationId.LastIndexOf('/')+1)

                    $output += $out
        
                }
            }
        }        
    }
}

$output | Format-Table $ft -AutoSize



