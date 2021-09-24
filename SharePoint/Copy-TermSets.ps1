
param(
    [Parameter(Mandatory=$false)]
    [string]$TermGroup = "My2",

    [Parameter(Mandatory=$false)]
    [string]$TermSet = "Test",

    [Parameter(Mandatory=$false)]
    [string]$TenantSite = "https://???????-admin.sharepoint.com",

    [Parameter(Mandatory=$false)]
    [string]$TargetSite = "https://???????.sharepoint.com/sites/????"
)


function SaveChildTerm($parent, $childTerm)
{
    $id = [guid]::NewGuid().tostring()

    $parent.CreateTerm($childTerm.Name, 1033, $id)

    Start-Sleep -Seconds 2

    $t2 = Get-PnPTerm -Identity $id

    $t2.SetCustomProperty("OriginalID", $childTerm.Id)

    if ($childTerm.TermsCount -gt 0)
    {
        foreach ($t3 in $childTerm.Terms)
        {
            SaveChildTerm $t2 $t3
        }
    }
}



$cTenant = Connect-PnPOnline -Url $TenantSite -Interactive
$cSite = Connect-PnPOnline -Url $TargetSite -Interactive


$terms = Get-PnPTerm -TermSet $TermSet -TermGroup $TermGroup -Recursive -IncludeChildTerms -Connection $cTenant

$tsTenant = Get-PnPTaxonomySession -Connection $cTenant

$gDest = Get-PnPSiteCollectionTermStore

if ($gDest -eq $null)
{
    $gDest = New-PnPSiteCollectionTermStore -Connection $cSite
}

$tsDest = Get-PnPTermSet -TermGroup $gDest -Identity $TermSet -ErrorAction Ignore

if ($tsDest -eq $null)
{
    $tsDest = New-PnPTermSet -Name Test -TermGroup $gDest -Lcid 1033
}

foreach ($term in $terms)
{
    $id = [guid]::NewGuid().tostring()

    New-PnPTerm -Name $term.Name -Lcid 1033 -Id $id -TermSet $tsDest -TermGroup $gDest -ErrorAction Ignore

    Start-Sleep -Seconds 2

    $parent = Get-PnPTerm -Identity $id

    $parent.SetCustomProperty("OriginalID", $term.Id)

    if ($term.TermsCount -gt 0)
    {
        foreach ($t2 in $term.Terms)
        {  
            SaveChildTerm $parent $t2
        }
    }

    #Start-Sleep -Seconds 2
}

<#
$cSite = Connect-PnPOnline -Url https://woodstinen.sharepoint.com/sites/nokia -Credentials $credObject

$localTerms = Get-PnPSiteCollectionTermStore -Connection $cSite

$items = Get-PnPListItem -List Termit -Fields "Title","Managed"

$i=$items[0].FieldValues



$localTermSet = Get-PnPTermSet -TermGroup $localTerms -Identity Test

$localTermSet.GetTermsWithCustomPropert ("OriginalID", $i.Managed.TermGuid.ToString())
#>

