# Azure Functions profile.ps1
#
# This profile.ps1 will get executed every "cold start" of your Function App.
# "cold start" occurs when:
#
# * A Function App starts up for the very first time
# * A Function App starts up after being de-allocated due to inactivity
#
# You can define helper functions, run commands, or specify environment variables
# NOTE: any variables defined that are not environment variables will get reset after the first execution

# Authenticate with Azure PowerShell using MSI.
# Remove this if you are not planning on using MSI or Azure PowerShell.
if ($env:MSI_SECRET) {
    Disable-AzContextAutosave -Scope Process | Out-Null
    Connect-AzAccount -Identity
}

# Microsoft To Do Exporter Helper Functions
function Get-AccessToken {
    param($Code)

    Add-Type -AssemblyName System.Web
    $bodyString = "grant_type=authorization_code&client_id=$([System.Web.HttpUtility]::UrlEncode($env:CLIENT_ID))&client_secret=$([System.Web.HttpUtility]::UrlEncode($env:CLIENT_SECRET))&code=$([System.Web.HttpUtility]::UrlEncode($Code))&redirect_uri=$([System.Web.HttpUtility]::UrlEncode($env:REDIRECT_URI))&scope=$([System.Web.HttpUtility]::UrlEncode('Tasks.Read'))"

    Write-Host "Body string: $bodyString"

    $token = Invoke-RestMethod `
        -Uri "https://login.microsoftonline.com/common/oauth2/v2.0/token" `
        -Method POST `
        -Body $bodyString `
        -ContentType "application/x-www-form-urlencoded"

    return $token.access_token
}

function Invoke-Graph {
    param($Uri, $Token)

    Invoke-RestMethod `
        -Uri $Uri `
        -Headers @{ Authorization = "Bearer $Token" }
}
function Get-HtmlTemplate {
    param($TemplateName, $Replacements = @{})

    $templatePath = Join-Path $PSScriptRoot "templates\$TemplateName.html"

    if (-not (Test-Path $templatePath)) {
        throw "Template not found: $templatePath"
    }

    $html = Get-Content $templatePath -Raw

    foreach ($key in $Replacements.Keys) {
        $placeholder = "{{$key}}"
        $value = $Replacements[$key]
        $html = $html -replace [regex]::Escape($placeholder), $value
    }

    return $html
}

function Render-ErrorResponse {
    param(
        [System.Net.HttpStatusCode]$StatusCode = [System.Net.HttpStatusCode]::InternalServerError,
        $Title = "Error",
        $Message = "An unexpected error occurred.",
        $Details = ""
    )

    $detailsClass = ''
    if (-not $Details) { $detailsClass = 'hidden' }

    $replacements = @{
        "TITLE" = $Title
        "MESSAGE" = $Message
        "DETAILS" = $Details
        "DETAILS_CLASS" = $detailsClass
    }

    try {
        $html = Get-HtmlTemplate -TemplateName "Error" -Replacements $replacements
    }
    catch {
        $html = "<h1>$Title</h1><p>$Message</p><pre>$Details</pre>"
    }

    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = $StatusCode
        ContentType = "text/html; charset=utf-8"
        Body        = $html
    })
}