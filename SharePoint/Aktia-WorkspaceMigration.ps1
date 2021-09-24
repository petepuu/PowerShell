
Import-Module Microsoft.SharePoint.MigrationTool.PowerShell
Import-Module MicrosoftTeams
Import-Module SharePointPnpPowerShellOnline


$tenantName = "peteiam"

$teamOwnerUPN = "peetteri@iampete.net"

$teamChannelName = "Työtilan Dokumentit"

$spOnpremURL = "http://teams.iampete.net/sites/aktia2"



$Global:SPOUrl = "https://$tenantname.sharepoint.com"
$Global:UserName = ""
#$Global:PassWord = ConvertTo-SecureString -String "" -AsPlainText -Force
$Global:SPOCredential = Get-Credential -UserName $Global:UserName -Message "SharePoint Online Credentials" #
#$Global:SPOCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Global:UserName, $Global:PassWord

$Global:SPUrl = $spOnpremURL
$Global:SPUserName = ""
#$Global:SPPassWord = ConvertTo-SecureString -String "" -AsPlainText -Force
$Global:SPCredential = Get-Credential -UserName $Global:SPUserName -Message "SharePoint On-Premise Credentials" #
#$Global:SPCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Global:SPUserName, $Global:SPPassWord


Connect-MicrosoftTeams -Credential $Global:SPOCredential



Connect-PnPOnline https://$tenantName-admin.sharepoint.com -Credentials $Global:SPOCredential


#$ExcludeLists = @("Images","Pages","Workflow Tasks","Master Page Gallery","Composed Looks")


#$s = Get-SPSite $spOnpremURL

#$webs = $s.AllWebs


#foreach ($web in $webs)
#{
    Register-SPMTMigration -SPOCredential $Global:SPOCredential -MigrateSiteSettings:$false -Force

    #if ($web.url -ne $siteUrl)
    #{
        Write-Host "Processing site $($spOnpremURL)"
        write-host

        $teamName = "Migrated_$($spOnpremURL.Substring($spOnpremURL.LastIndexOf('/')+1))"
        
                
        $team = Get-Team -MailNickName $teamName

        if ($team -eq $null)
        {
            Write-Host "Provisioning new Team with name -> $teamName"

            $team = New-Team -DisplayName $teamName -MailNickName $teamName -Visibility Private -Owner $teamOwnerUPN
        }
        else
        {
            Write-Host "Team '$($teamName)' already exists" -ForegroundColor Yellow
        }


        $channels = Get-TeamChannel -GroupId $team.GroupId

        $channel = $channels | ? {$_.DisplayName -eq $teamChannelName}

        if ($channel -eq $null)
        {
            Write-Host "Create new Channel -> 'Työtilan Dokumentit'"
            New-TeamChannel -GroupId $team.groupId -DisplayName "Työtilan Dokumentit" | Out-Null
        }
        else
        {
            Write-Host "Channel '$($teamChannelName)' already exists" -ForegroundColor Yellow
        }

        

        write-host "Sleeping for 10 seconds to wait Team provisioning to complete"
        Start-Sleep -Seconds 10

        
        
        # get sharepoint site url for the new team
        $i = Get-PnPListItem -List 'DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS' -Query "<View><Query><Where><Eq><FieldRef Name='GroupId'/><Value Type='Guid'>$($team.groupId)</Value></Eq></Where></Query></View>"
        $teamSPSiteUrl = $i.FieldValues.SiteUrl

        Add-SPMTTask    -SharePointSourceCredential $Global:SPCredential `
                                -SharePointSourceSiteUrl $spOnpremURL `
                                -SourceList $teamChannelName `
                                -TargetSiteUrl $teamSPSiteUrl `
                                -TargetList "Documents" `
                                -TargetListRelativePath $teamChannelName

        <#
        foreach ($l in $web.Lists)
        { 
            if ($l.title -inotin $ExcludeLists)
            {
                #$l.Title
                #$l.ForceCheckout

                
            }
        }
        #>
        
      
        Start-SPMTMigration -Verbose
    #}
#}


