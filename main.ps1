.\variables.ps1


function Get-GraphToken {
  param (
    [Parameter(Mandatory = $true)]
    [String]  $appID,
    [Parameter(Mandatory = $true)]
    [String]   $clientSecret,
    [Parameter(Mandatory = $true)]
    [String]    $tenantID
  )

#Prepare token request
$url = 'https://login.microsoftonline.com/' + $tenantId + '/oauth2/v2.0/token'

$body = @{
    grant_type = "client_credentials"
    client_id = $appID
    client_secret = $clientSecret
    scope = "https://graph.microsoft.com/.default"
}

#Obtain the token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $url -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing -ErrorAction Stop
($tokenRequest.Content | ConvertFrom-Json).access_token
}

function Get-SPSites {
  param (
    [Parameter(Mandatory = $true)]
    [String] $token,
    [Parameter(Mandatory = $false)]
    [String] $siteName
  )

$authHeader = @{
   'Content-Type'='application\json'
   'Authorization'="Bearer $token"
}
$sites = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites"  -Headers $authHeader).value

try {
  if ($siteName) {
    $sites = $sites | Where-Object {$_.displayName -eq $siteName}
    }
}
catch {
  Write-Error  "No site found with the name $siteName"
}

If (!$sites.count -gt 0) {
  Return "No site found with the name $siteName"
  }
  $sites

}

function Get-SPLists {
 param (
  [Parameter(Mandatory = $true)]
  [String] $token,
  [Parameter(Mandatory = $false)]
  [Object] $sites,
  [Parameter(Mandatory = $false)]
  [String] $siteName

 )
$siteslists = @()
if ($sites) {
foreach ($site in $sites) {
  $authHeader = @{
    'Content-Type'='application\json'
    'Authorization'="Bearer $token"
  }
  if ($siteName) {
    if ($siteName -eq $site.Name) {
    (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($site.Id)/lists" -Headers $authHeader).value
    Break
  }
}else{
  $sitelists = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($site.Id)/lists" -Headers $authHeader).value
    $siteslists += $sitelists
}
$siteslists
}
}elseif ($siteName) {
  $sites = Get-SPSites  -token $token -siteName $siteName
  Get-SPLists $token -sites $sites
}
}

function New-SPListFromCSV {
  param(
    [Parameter(Mandatory = $true)]
    [String] $token,
    [Parameter(Mandatory = $true)]
    [String] $siteName,
    [Parameter(Mandatory = $false)]
    [String] $listName,
    [Parameter(Mandatory = $True)]
    [String] $csvFilePath
  )
    $list = Import-CSV -Path $csvFilePath -Encoding UTF8

    if (!$listname)
    {
      $filenameWithoutExtension = [IO.Path]::GetFileNameWithoutExtension($csvFilePath)
      $listname = $filenameWithoutExtension
    }
    $body = '{
      "displayName": "' + $listname + '",
      "columns": ['
    foreach ($column in $list[0].psobject.properties) {
        $body += '{ "name" : "' + $column.Name + '", "text" : {} },'
    }
    $body += '],
      }'

    $authHeader = @{
      'Content-Type'='application\json'
      'Authorization'="Bearer $token"
    }
    $siteId = (Get-SPSites  -token $token  -siteName $siteName).id
    $endpointURI = "https://graph.microsoft.com/v1.0/sites/" + $siteId + "/lists"
    $createListRequest = Invoke-RestMethod -Uri $endpointURI  -Method Post -Headers $authHeader -Body $body -ContentType 'application/json'
  }

$token = Get-GraphToken -appID $appID -clientSecret $clientSecret -tenantID $tenantID
$sites = Get-SPSites -token $token 
$sites
#-siteName "Team Site"
$sitesLists = Get-SPLists -token $token -siteName "Team Site"
$sitesLists.Count
$body = New-SPListFromCSV -token $token -siteName "Team Site" -csvFilePath "C:\Users\ludov\Downloads\Security Onion - DNS - Query.csv"

<#

# Get all the lists of a specific site and list their id, names and webUrl
$sites.value | ForEach-Object {
    
            $siteId = $_."id"
            $endpointURI = "https://graph.microsoft.com/v1.0/sites/" + $siteId + "/lists"
            $lists = Invoke-RestMethod -Uri $endpointURI -Headers $authHeader
            $lists.value | Select-Object -Property Id, Name, webUrl
    
<#for ($i = 31; $i -lt 60; $i++) {
    Write-Host "Iteration number: $i"
    $listTitle = "List" + $i
    $body = '{
        "displayName": "' + $listTitle + '",
        "description": "Discover teams to join in Office 365 for IT Pros",
        "columns": [
          {
            "name": "Deeplink",
            "description": "Link to access the team",
            "text": { }
        },{
            "name": "Description",
            "description": "Purpose of the team",
            "text": { }
          },
          {
            "name": "Owner",
            "description": "Team owner",
            "text": { }
          },      
          {
            "name": "OwnerSMTP",
            "description": "Primary SMTP address for owner",
            "text": { }
          },
          {
            "name": "Members",
            "description": "Number of tenant menbers",
            "number": { }
          },
          {
            "name": "ExternalGuests",
            "description": "Number of external guest menbers",
            "number": { }
          },
          {
            "name": "Access",
            "description": "Public or Private access",
            "text": { }
          },
        ],
      }'

# Create a new SharePoint list using the Microsoft Graph API


$createListRequest = Invoke-RestMethod -Uri $endpointURI  -Method Post -Headers $authHeader -Body $body -ContentType 'application/json'
$createListRequest
}#>



