# Input bindings are passed in via param block.
param($Timer)

# Get the current universal time in the default string format.
$currentUTCtime = (Get-Date).ToUniversalTime()

# The 'IsPastDue' property is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}

function Get-Report{
    [cmdletbinding()]
        Param(
            [Parameter(Mandatory = $true, Position = 0)]
            [string]$report,
            [parameter(Mandatory = $true, Position = 1)]
            [string]$parameterset
        )

    $reportRoot = @()
    $headerParams = Get-AuthToken $env:clientID $env:clientSecret $env:tenantID
    $uri = "https://graph.microsoft.com/beta/reports/{0}({1})?`$format=application/json" -f $report, $parameterset
    
    do 
    {
      $response = Invoke-RestMethod -Headers $headerParams -Uri $uri -UseBasicParsing -Method Get 
      $reportRoot += $response.value
      $uri = $response.'@odata.nextlink'
    } while ($null -ne $url)

    return $reportRoot
}

function Get-AuthToken{
    [cmdletbinding()]
        Param(
            [Parameter(Mandatory = $true, Position = 0)]
            [string]$ClientID,
            [parameter(Mandatory = $true, Position = 1)]
            [string]$ClientSecret,
            [Parameter(Mandatory = $true, Position = 2)]
            [string]$TenantID
        )
    # Create app of type Web app / API in Azure AD, generate a Client Secret, and update the client id and client secret here
    $loginURL = "https://login.microsoftonline.com/"
    # Get the tenant ID from Properties | Directory ID under the Azure Active Directory section
    $resource = "https://graph.microsoft.com"
    # auth
    $body = @{grant_type="client_credentials";resource=$resource;client_id=$ClientID;client_secret=$ClientSecret}
    $oauth = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantID/oauth2/token?api-version=1.0 -Body $body
    $headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
    return $headerParams 
}

#Enumerators and object to wrap the incoming request

$date = (Get-Date).AddDays(-5).ToString('yyyy-MM-dd')
$period = "D180"
$result = @()
Write-Host "Processing items from $date"

$report = "getSharePointActivityUserDetail"
#$parameterset = "date={0}" -f $date
$parameterset = "period='{0}'" -f $period
$result += Get-Report $report $parameterset
##
$report = "getSharePointActivityFileCounts"
#$parameterset = "date={0}" -f $date
$parameterset = "period='{0}'" -f $period
$result += Get-Report $report $parameterset
##
$report = "getSharePointActivityUserCounts"
#$parameterset = "date={0}" -f $date
$parameterset = "period='{0}'" -f $period
$result += Get-Report $report $parameterset
##
$report = "getSharePointActivityPages"
#$parameterset = "date={0}" -f $date
$parameterset = "period='{0}'" -f $period
$result += Get-Report $report $parameterset
Push-OutputBinding -Name outputDocument -Value $result -clobber 
Write-Host "Pushed $($result.Count) items to CosmosDB"

# Write an information log with the current time.
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"
