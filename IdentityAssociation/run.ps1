using namespace System.Net

param($Request, $TriggerMetadata)

# Return Microsoft identity association JSON
$identityAssociation = @{
    associatedApplications = @(
        @{
            applicationId = $env:CLIENT_ID
        }
    )
} | ConvertTo-Json

Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    ContentType = "application/json; charset=utf-8"
    Body       = $identityAssociation
})