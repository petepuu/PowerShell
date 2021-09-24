
param
(
    [Parameter(Mandatory=$true)]
    [string]$WebApplication,

    [Parameter(Mandatory=$false)]
    [Int]$Days = 14,

    [Parameter(Mandatory=$false)]
    [Int]$Months = 6
)


Add-PSSnapin microsoft.sharepoint.powershell -ea 0

try
{
    $webApp = Get-SPWebApplication $WebApplication -ErrorAction 1
}
catch [Exception]
{
    Write-Host
    Write-Host "Error: " $_.Exception.Message -ForegroundColor Red
    Exit 0
}

$d = (Get-Date).ToString("ddMMyyyy")

$outputDir = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)

$outputFileDaily = "$($outputDir)\DailyUsageReport_$($d).csv"
$outputFileMonthly = "$($outputDir)\MonthlyUsageReport_$($d).csv"


$ftDaily = @{Expression={$_.Site};Label="Site"}, 
	        @{Expression={$_.Date};Label="Date"},
            @{Expression={$_.Hits};Label="Hits"},
            @{Expression={$_.UniqueUsers};Label="UniqueUsers"}

$ftMonthly = @{Expression={$_.Site};Label="Site"}, 
	        @{Expression={$_.Month};Label="Month"},
            @{Expression={$_.Hits};Label="Hits"},
            @{Expression={$_.UniqueUsers};Label="UniqueUsers"}

$outputDaily = @()
$outputMonthly = @()


$searchApp = Get-SPEnterpriseSearchServiceApplication

$curDate = Get-Date

foreach ($site in $webApp.Sites)
{
    $usage = $searchApp.GetRollupAnalyticsItemData(1, [System.Guid]::Empty, $site.ID, [System.Guid]::Empty)
    $i = 0

    # Daily usage
    do
    {
        $date = $curDate.AddDays(-($i))
    
        $hits = 0
        $uniqueUsers = 0
    

        try
        {
            $usage.GetDailyData($date, [ref]$hits, [ref]$uniqueUsers)
        }
        catch [Exception] {}

     

        $out = New-Object PSObject
        $out | Add-Member -MemberType NoteProperty -Name Site -Value $site.Url
        $out | Add-Member -MemberType NoteProperty -Name Date -Value $date.ToShortDateString()
        $out | Add-Member -MemberType NoteProperty -Name Hits -Value $hits
        $out | Add-Member -MemberType NoteProperty -Name UniqueUsers -Value $uniqueUsers

        $outputDaily += $out

        $i = $i+1

    } while ($i -lt $Days)


    $m = 0

    # Monthly usage
    do
    {
        $date = $curDate.AddMonths(-($m))
    
        $hits = 0
        $uniqueUsers = 0
    

        try
        {
            $usage.GetMonthlyData($date, [ref]$hits, [ref]$uniqueUsers)
        }
        catch [Exception] {}


        $out = New-Object PSObject
        $out | Add-Member -MemberType NoteProperty -Name Site -Value $site.Url
        $out | Add-Member -MemberType NoteProperty -Name Month -Value "$($date.Month)/$($date.Year)"
        $out | Add-Member -MemberType NoteProperty -Name Hits -Value $hits
        $out | Add-Member -MemberType NoteProperty -Name UniqueUsers -Value $uniqueUsers

        $outputMonthly += $out

        $m = $m+1

    } while ($m -le $Months)
    
}


$outputDaily | Select-Object Site, Date, Hits, UniqueUsers | Export-Csv $outputFileDaily -NoTypeInformation

$outputMonthly | Select-Object Site, Month, Hits, UniqueUsers | Export-Csv $outputFileMonthly -NoTypeInformation
