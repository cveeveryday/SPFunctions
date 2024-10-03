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

function Grant-SPSelectedSitePermissions {
  param (
    [Parameter(Mandatory = $true)]
    [String] $token,
    [Parameter(Mandatory = $true)]
    [String] $siteId,
    [Parameter(Mandatory = $true)]
    [string] $appClientId,
    [Parameter(Mandatory = $true)]
    [String] $appDisplayName
  )
$authHeader = @{
  'Content-Type'='application\json'
  'Authorization'="Bearer $token"
}


$url = "https://graph.microsoft.com/v1.0/sites/$siteId/permissions"

$permissions = @(
  @{Permission="Sites.Read"},
  @{Permission="Sites.Write"},
  @{Permission="Sites.Manage"},
  @{Permission="Sites.FullControl"}
)
ForEach ($permission in $permissions) {

$body = '{
  "roles": [' + $permission.Permission + '],
  "grantedToIdentities": [{
     "application": {
       "id": "' + $appClientId + '",
       "displayName": "' + $appDisplayName + '"
       }
       }]
       }'
  
      $response = Invoke-RestMethod -Uri $url -Headers $authHeader -Method Post -Body $body -ContentType 'application\json'
      $response.Value
}

}


function Get-SPLists {
 param (
  [Parameter(Mandatory = $true)]
  [String] $token,
  [Parameter(Mandatory = $false)]
  [Object] $sites,
  [Parameter(Mandatory = $false)]
  [String] $siteName,
  [Parameter(Mandatory = $false)]
  [String] $listName
 )

$siteslists = @()
$authHeader = @{
  'Content-Type'='application\json'
  'Authorization'="Bearer $token"
}
if ($sites) {
foreach ($site in $sites) {
  if ($siteName) {
    if ($siteName -eq $site.Name) {
    (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($site.Id)/lists" -Headers $authHeader).value
    Break
  }
}else{
  $sitelists = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($site.Id)/lists" -Headers $authHeader).value
  if ($listName){
    $sitelists | Where-Object {$_.displayName -eq $listName}
    break
  }
    $siteslists += $sitelists
}
$siteslists
}
}elseif ($siteName) {
  $sites = Get-SPSites  -token $token -siteName $siteName
  if ($listName) {
    Get-SPLists $token -sites $sites -listName $listName
  }else{
    Get-SPLists $token -sites $sites
  }
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
      $listname = Get-ListNameFromCSVFileName -csvFilePath $csvFilePath
    }
    $body = '{
      "displayName": "' + ($listname -replace '\s', '') + '",
      "columns": ['
    foreach ($column in $list[0].psobject.properties) {
        $body += '{ "name" : "' + ($column.Name -replace '\s', '') + '", "text" : {} },'
    }
    $body += '],
      }'

    $authHeader = @{
      'Content-Type'='application\json'
      'Authorization'="Bearer $token"
    }
    $siteId = (Get-SPSites  -token $token  -siteName $siteName).id
    $endpointURI = "https://graph.microsoft.com/v1.0/sites/" + $siteId + "/lists"
    Invoke-RestMethod -Uri $endpointURI  -Method Post -Headers $authHeader -Body $body -ContentType 'application/json'
  }

  function New-SPListFromObject {
    param(
      [Parameter(Mandatory = $true)]
      [String] $token,
      [Parameter(Mandatory = $true)]
      [String] $siteName,
      [Parameter(Mandatory = $true)]
      [String] $listName,
      [Parameter(Mandatory = $True)]
      [String[]] $colunmns
      )
    $body = '{
      "displayName": "' + ($listname -replace '\s', '') + '",
      "columns": ['
      foreach ($column in $colunmns) {
      $body  += '{ "name" : "' + ($column -replace '\s', '') + '", "text": {} },'
      }
        $body += '],
        }'
        $authHeader = @{
          'Content-Type'='application\json'
          'Authorization'="Bearer $token"
        }
        $siteId = (Get-SPSites  -token $token  -siteName $siteName).id
        $endpointURI = "https://graph.microsoft.com/v1.0/sites/" + $siteId + "/lists"
        Invoke-RestMethod -Uri $endpointURI  -Method Post -Headers $authHeader -Body $body -ContentType 'application/json'
  }

function Add-CSVToSPList {
  param(
    [Parameter(Mandatory = $true)]
    [String] $token,
    [Parameter(Mandatory = $true)]
    [String] $siteName,
    [Parameter(Mandatory = $true)]
    [String] $csvFilePath
    )
    $listname = Get-ListNameFromCSVFileName -csvFilePath $csvFilePath
    $list = Get-SPLists -token $token  -siteName $siteName -listName $listName
    If ($list)
      {
      $listId = $list.id
      $authHeader = @{
        'Content-Type'='application\json'
        'Authorization'="Bearer $token"
      }
      $siteId = (Get-SPSites  -token $token  -siteName $siteName).id
      $endpointURI = "https://graph.microsoft.com/v1.0/sites/" + $siteId + "/lists/" + $listId + "/items"
      $csvFile = Import-Csv -Path $csvFilePath -Encoding 'UTF8'
      foreach ($row in $csvFile)
      {
        $body = '{ "fields": {'
        foreach ($column in $row.psobject.properties)
        {
          $value = $column.Value -replace('"', '\"')
          $body += '"' + ($column.Name -replace '\s', '')  + '": "' + $column.Value + '",'
        }
        $body = $body.TrimEnd(',') + '} }'
        Invoke-RestMethod -Uri $endpointURI   -Method Post  -Headers $authHeader  -Body $body  -ContentType 'application/json'
      }
    }
}

function Get-ListNameFromCSVFileName {
  param (
    [Parameter(Mandatory = $true)]
    [String] $csvFilePath
  )
  
  $filenameWithoutExtension = [IO.Path]::GetFileNameWithoutExtension($csvFilePath)
  ($filenameWithoutExtension -replace '[\W_]+', '')

}


$token = Get-GraphToken -appID $appID -clientSecret $clientSecret -tenantID $tenantID
$sites = Get-SPSites -token $token 
$sites.count
#$sites
#$sitesLists = Get-SPLists -token $token -siteName "Team Site"
#$sitesLists.Count
#$siteId = (Get-SPSites -name "Team Site").Id
#$response = Grant-SPSelectedSitePermissions -token $token -siteId $siteId -appClientId $appClientId -appDisplayName $appDisplayName
#$body = New-SPListFromCSV -token $token -siteName "Team Site" -csvFilePath "C:\Users\ludov\Downloads\1810000402_MetaData.csv"
New-SPListFromObject -token $token -siteName "Team Site" -listName "Patty's Emails" -colunmns @("From","To","DateReceived","Subject","Body")
#$body
#$list = Get-SPLists  -token $token  -siteName "Team Site" -listName "Security Onion - DNS - Query"
#$list.Count
#$list
#Add-CSVToSPList -token $token -siteName "Team Site" -csvFilePath "C:\Users\ludov\Downloads\1810000402_MetaData.csv"



