using namespace System.Net

param($Request, $TriggerMetadata)

# Check for errors from OAuth provider
if ($Request.Query.error) {
    $errorDescription = $Request.Query.error_description
    if (-not $errorDescription) {
        $errorDescription = "Authorization failed: $($Request.Query.error)"
    }

    Render-ErrorResponse -StatusCode ([HttpStatusCode]::BadRequest) -Title "Authorization Error" -Message $errorDescription
    return
}

# Get authorization code
$code = $Request.Query.code
if (-not $code) {
    Render-ErrorResponse -StatusCode ([HttpStatusCode]::BadRequest) -Title "Missing Authorization Code" -Message "Missing authorization code"
    return
}

# Exchange code for access token
try {
    $token = Get-AccessToken $code

    # Show intermediate "exporting" page that submits token via POST
    $replacements = @{
        "ACCESS_TOKEN" = $token
    }

    $html = Get-HtmlTemplate -TemplateName "Callback" -Replacements $replacements

    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        ContentType = "text/html; charset=utf-8"
        Body       = $html
    })
}
catch {
    Write-Host "Error getting access token: $_"
    Render-ErrorResponse -StatusCode ([HttpStatusCode]::InternalServerError) -Title "Token Error" -Message "Failed to obtain access token. Please try again." -Details "$_"
}
