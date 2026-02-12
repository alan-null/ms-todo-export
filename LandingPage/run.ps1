using namespace System.Net

param($Request, $TriggerMetadata)

$authorizeUrl = @(
    $env:AUTHORIZE_URL
    "?client_id=$($env:CLIENT_ID)"
    "&response_type=code"
    "&redirect_uri=$($env:REDIRECT_URI)"
    "&scope=Tasks.Read"
    "&response_mode=query"
) -join ''

$replacements = @{
    "AUTHORIZE_URL" = $authorizeUrl
}

$html = Get-HtmlTemplate -TemplateName "LandingPage" -Replacements $replacements

Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = [HttpStatusCode]::OK
        ContentType = "text/html; charset=utf-8"
        Body        = $html
    })
