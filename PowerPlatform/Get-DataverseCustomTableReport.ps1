param
(
    [Parameter(Mandatory=$true)]
    [string]$instanceUrl,

    [Parameter(Mandatory=$true)]
    [string]$username,

    [Parameter(Mandatory=$true)]
    [string]$tablePrefix
)

Add-Type -Path (Join-Path (Split-Path $script:MyInvocation.MyCommand.Path) "Microsoft.IdentityModel.Clients.ActiveDirectory.WindowsForms.dll")
Add-Type -Path (Join-Path (Split-Path $script:MyInvocation.MyCommand.Path) "Microsoft.IdentityModel.Clients.ActiveDirectory.dll")

$Password = Read-Host -AsSecureString -Prompt "Password"

[Hashtable] $Headers = @{}

$version = "9.1"
$webapiurl = "$instanceUrl/api/data/v$version/"


[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
 
$authContext = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext("https://login.microsoftonline.com/common");
$credential = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.UserCredential($username, $Password)
$authResult = $authContext.AcquireToken($instanceUrl, "51f81489-12ee-4a9e-aaae-a2591f45987d", $credential);

$token = $authResult.AccessToken #Get-JwtToken -Audience $audience
$Headers["Authorization"] = "Bearer $token";
$Headers["Accept"] = "application/json";
$Headers["OData-MaxVersion"] = "4.0";
$Headers["OData-Version"] = "4.0";
$Headers["If-None-Match"] = "null";
$Headers["Content-Type"] = "application/json"
$Headers["User-Agent"] = "PowerShell cmdlets 1.0";


$reqCustomTablesParams =
@{
    URI = "$($instanceUrl)/api/data/v9.2/entities?`$select=originallocalizedcollectionname&`$filter=startswith(entitysetname,'$tablePrefix')"
    Headers = @{
        "Authorization" = "$($authResult.AccessTokenType) $token" 
    }
    Method = 'GET'
}

$reqViewsParams =
@{
    URI = "$($instanceUrl)/api/data/v9.2/savedqueries?`$select=name,_createdby_value"
    Headers = @{
        "Authorization" = "$($authResult.AccessTokenType) $token" 
    }
    Method = 'GET'
}

$reqUsersParams =
@{
    URI = "$($instanceUrl)/api/data/v9.2/systemusers?`$select=azureactivedirectoryobjectid,internalemailaddress,fullname&`$filter=issyncwithdirectory eq true"
    Headers = @{
        "Authorization" = "$($authResult.AccessTokenType) $token"  
    }
    Method = 'GET'
}

# Get custom tables
$resCustomTables = Invoke-RestMethod @reqCustomTablesParams -ErrorAction Stop

$tables = $resCustomTables.value

# Get all AAD synched users (AAD->Dataverse) which includes cloud accounts as well
$resUsers = Invoke-RestMethod @reqUsersParams -ErrorAction Stop

$users = $resUsers.value

# Get all Views (savedqueries)
$resViews = Invoke-RestMethod @reqViewsParams -ErrorAction Stop

$views = $resViews.value

$ft = @{Expression={$_.Table};Label="Table"}, 
	    @{Expression={$_.CreatedByEmail};Label="CreatedBy-Email"},
        @{Expression={$_.CreatedByObjectId};Label="CreatedBy-ObjectId"}
    
    $output = @()


foreach ($t in $tables)
{
    $v = $views | ? { $_.name -eq "Active $($t.originallocalizedcollectionname)" } | Select -First 1

    $createdBy = $users | ? { $_.systemuserid -eq $v._createdby_value }
    
    $out = New-Object PSObject
    $out | Add-Member -MemberType NoteProperty -Name Table -Value $t.originallocalizedcollectionname
    $out | Add-Member -MemberType NoteProperty -Name CreatedByEmail -Value $createdBy.internalemailaddress
    $out | Add-Member -MemberType NoteProperty -Name CreatedByObjectId -Value $createdBy.azureactivedirectoryobjectid

    $output += $out
}


$output | Format-Table $ft -AutoSize