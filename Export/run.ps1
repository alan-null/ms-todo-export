using namespace System.Net

param($Request, $TriggerMetadata)

# Get access token from query parameter (GET) or form body (POST)
$token = $Request.Query.token
if (-not $token -and $Request.Body) {
    # Parse form data from POST request body
    $bodyString = $Request.Body
    if ($bodyString -and $bodyString -match "token=([^&]+)") {
        $token = [System.Net.WebUtility]::UrlDecode($matches[1])
    }
}
if (-not $token) {
    Render-ErrorResponse -StatusCode ([HttpStatusCode]::Unauthorized) -Title "Missing Token" -Message "Missing access token"
    return
}

try {
    $base = "https://graph.microsoft.com/v1.0"

    Write-Host "Fetching Microsoft To Do lists..."
    $lists = Invoke-GraphGet -Token $token -Url "$base/me/todo/lists/delta"

    # Fetch tasks for each list and embed them
    foreach ($list in $lists) {
        Write-Host "Fetching tasks for list: $($list.displayName) ($($list.id))"
        Start-Sleep -Milliseconds 100  # Rate limiting

        $list | Add-Member -NotePropertyName 'tasks' -NotePropertyValue @() -Force
        $list.tasks = Invoke-GraphGet -Token $token -Url "$base/me/todo/lists/$($list.id)/tasks"
        Write-Host "Got $($list.tasks.Count) tasks for list '$($list.displayName)'"
    }

    Write-Host "Export completed successfully. Total lists: $($lists.Count)"

    # Convert to JSON
    $json = $lists | ConvertTo-Json -Depth 100

    # Base64 encode the JSON for safe embedding
    $jsonBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($json))

    # Generate filename with timestamp
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $filename = "microsoft-todo-backup-$timestamp.json"

    # Calculate stats
    $listCount = $lists.Count
    $totalTasks = ($lists | ForEach-Object { $_.tasks.Count } | Measure-Object -Sum).Sum
    $exportDate = (Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC")

    # Create HTML response with embedded download
    $replacements = @{
        "LIST_COUNT"  = $listCount
        "TOTAL_TASKS" = $totalTasks
        "EXPORT_DATE" = $exportDate
        "FILENAME"    = $filename
        "JSON_BASE64" = $jsonBase64
    }

    $html = Get-HtmlTemplate -TemplateName "Export" -Replacements $replacements

    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode  = [HttpStatusCode]::OK
            ContentType = "text/html; charset=utf-8"
            Body        = $html
        })
}
catch {
    Write-Host "Error during export: $_"

    $errorMessage = $_.Exception.Message
    if ($_.ErrorDetails.Message) {
        $errorMessage += ": $($_.ErrorDetails.Message)"
    }

    Render-ErrorResponse -StatusCode ([HttpStatusCode]::InternalServerError) -Title "Export Error" -Message "Failed to export tasks" -Details $errorMessage
}
