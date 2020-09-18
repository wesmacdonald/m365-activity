using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

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
    $resource = "https://manage.office.com"
    # auth
    $body = @{grant_type="client_credentials";resource=$resource;client_id=$ClientID;client_secret=$ClientSecret}
    $oauth = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantID/oauth2/token?api-version=1.0 -Body $body
    $headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
    return $headerParams 
}

#Validate AuthId from Headers in request
if ($env:AuthId -eq $request.headers["WebHook-AuthID"])
{
    #Enumerators and object to wrap the incoming request
    $pageArray = @()
    $rawreq = @()
    $rawreq = New-Object -TypeName psobject
    $rawreq | Add-Member -name Content -value Content -membertype noteproperty

    #Retrieve the content URI
    #$requestbody = Get-Content $Request -Raw | ConvertFrom-Json
    $requestbody = $request.body
    $rawreq.content  = $requestbody | convertto-json
    #Write-Host "Raw $($rawreq.content)"

    
    $headerParams = Get-AuthToken $env:clientID $env:clientSecret $env:tenantID
      
    #If more than one page is returned capture and return in pageArray
    if ($REQ_HEADERS_NextPageUri) 
    {     
        $pageTracker = $true
        $pagedReq = $REQ_HEADERS_NextPageUri

        while ($pageTracker) 
        {   
            $CurrentPage = Invoke-WebRequest -Headers $headerParams -Uri $pagedReq -UseBasicParsing
            $pageArray += $CurrentPage

            if ($CurrentPage.Headers.NextPageUri)
            {
              $pageTracker = $true    
              $pagedReq = $CurrentPage.Headers.NextPageUri
            }
            else
            {
              $pageTracker = $false
            }    
        }
    }     
    
    $pageArray += $rawreq

    foreach ($page in $pageArray)
    {

      $req = $page.content | ConvertFrom-Json                                       

      foreach ($content in $req)     
      {
        $uri = $content.contentUri + "?PublisherIdentifier=" + $env:TenantID  
        $records = Invoke-WebRequest -UseBasicParsing -Headers $headerParams -Uri $uri
        Push-OutputBinding -Name outputDocument -Value $records.content -clobber 
      }
    }
}

Push-OutputBinding -Name response -Value ([HttpResponseContext]@{
    StatusCode = [System.Net.HttpStatusCode]::OK
    Body = ""
}) 
