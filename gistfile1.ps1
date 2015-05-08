$clientId = "CLIENT_ID_HERE"
$codeUrl = "https://www.yammer.com/dialog/oauth?client_id=$clientId"
$clientSecret = "CLIENT_SECRET_HERE"
$SleepInterval = 1

$IE = New-Object -ComObject InternetExplorer.Application;
$IE.Navigate($codeUrl);
$IE.Visible = $true;

while ($IE.LocationUrl -notmatch ‘code=’) {
Write-Debug -Message (‘Sleeping {0} seconds for access URL’ -f $SleepInterval);
Start-Sleep -Seconds $SleepInterval;
}

Write-Debug -Message (‘Callback URL is: {0}’ -f $IE.LocationUrl);
[Void]($IE.LocationUrl -match ‘=([\w\.]+)’);
$tempCode = $Matches[1];

$IE.Quit();

$r = Invoke-WebRequest https://www.yammer.com/oauth2/access_token.json?client_id=$clientId"&"client_secret=$clientSecret"&"code=$tempCode | ConvertFrom-Json

$tempCode
$accessToken = $r.access_token.token

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization", 'Bearer '+ $accessToken)

$headers

$currentUser = Invoke-RestMethod 'https://www.yammer.com/api/v1/users/current.json' -Headers $headers
$allTokens = Invoke-RestMethod 'https://www.yammer.com/api/v1/oauth/tokens.json'-Headers $headers
 

Write-Host $currentUser.name
Write-Host $currentUser.network_name
Write-Host $currentUser.email


foreach ($token in $allTokens) {
   $token | Format-Table user_id, network_name,token -autosize
 }


 