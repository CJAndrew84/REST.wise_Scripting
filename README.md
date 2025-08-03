# {REST:wise} Scripting
## Scripting ProjectWise and other systems RESTfully

No Fluff. No “Hello World.” Just Real Solutions.

This repo is for anyone who’s:

Stared blankly at WSG’s empty JSON responses
Fought PowerShell’s curly brace escaping (and lost, then won, then lost again)
Watched a working script break because a document version changed
If that’s you — welcome. You’ve found the right corner of the internet.

Why This Exists

Inspired by Brian M. Flaherty, Dave Brumbaugh, P.E., and the ProjectWise PowerShell Blog — but built from my own day-to-day battles with ProjectWise, PowerShell, and Bentley’s iTwin APIs. I’m here to share the scripts, tools, and “wish I’d known this sooner” moments that actually help you get real work done.

What You’ll Find Here

PowerShell snippets that actually run in production
Power Automate tips that save hours of clicking
Postman collections for reverse-engineering undocumented APIs
Complete scripts, walkthroughs, and logic explanations
Real-world automation insights for ProjectWise and iTwin
Coming Soon

Using PowerShell to talk to the WSG API (and get something useful back)
Auth flows and access tokens — the real-world version
Connecting ProjectWise to Power Automate — the right way (well, my way)
Do Something Useful Right Now

How to List ProjectWise Datasources via WSG

Endpoint:
GET /WS/version/repositories

Option 1: Basic Auth (ProjectWise Logical Account)

PowerShell
$wsgUrl = "https://yourserver-pw-ws.bentley.com/ws/version/repositories"
$username = "youruser"
$password = "yourpassword"

$securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$creds = New-Object System.Management.Automation.PSCredential ($username, $securePassword)
$encodedCreds = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($creds.UserName + ":" + $creds.GetNetworkCredential().Password))

$headers = @{
  "Authorization" = "Basic $encodedCreds"
  "Accept"        = "application/json"
}

$response = Invoke-WebRequest -Uri $wsgUrl -Headers $headers -Method Get -ContentType "application/json" -UseBasicParsing
$response
Option 2: Token Auth via pwps_dab

PowerShell
Import-Module pwps_dab
New-PWLogin -BentleyIMS

$token = "$(ConvertTo-EncodedToken $(Get-PWConnectionClientToken -UsePWRelyingParty))"

$headers = @{
  "Authorization" = "Bearer $token"
  "Accept"        = "application/json"
}

$wsgUrl = "https://yourserver-pw-ws.bentley.com/ws/version/repositories"

$response = Invoke-WebRequest -Uri $wsgUrl -Headers $headers -Method Get -ContentType "application/json" -UseBasicParsing
$response
If you see a list of datasources (with id and displayName), congrats — you’re scripting ProjectWise RESTfully.

TL;DR

This repo is your source for real-world scripting and automation for ProjectWise and Bentley APIs — using PowerShell, Power Platform, and REST APIs. No nonsense, just solutions that work.

Want more working scripts, reusable functions, and fewer silent failures?
Subscribe to the newsletter. Let’s start scripting ProjectWise RESTfully.
