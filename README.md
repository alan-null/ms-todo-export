# Microsoft To Do Exporter

Azure Functions app to export Microsoft To Do lists and tasks.

Application available at:

**https://ms-todo-export.azurewebsites.net**

## Setup

### 1. Azure App Registration (Entra ID)

1. Go to [Azure Portal](https://portal.azure.com) → Microsoft Entra ID → App registrations
2. Create a new registration:
   - **Name**: Microsoft To Do Exporter
   - **Account types**: Multitenant (or single tenant)
   - **Redirect URI**:
     - Local: `http://localhost:7071/api/callback`
     - Production: `https://<your-function-app>.azurewebsites.net/api/callback`
3. After creation, note the **Application (client) ID**
4. Go to **Certificates & secrets** → New client secret → Note the **Value**
5. Go to **API permissions** → Add permission → Microsoft Graph → Delegated:
   - `Tasks.Read`

### 2. Local Development

1. Install [Azure Functions Core Tools](https://learn.microsoft.com/en-us/azure/azure-functions/functions-run-local)
2. Create or update `local.settings.json` at the project root with the following example. Do NOT commit secrets to source control.

```json
{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "UseDevelopmentStorage=true",
    "FUNCTIONS_WORKER_RUNTIME_VERSION": "7.4",
    "FUNCTIONS_WORKER_RUNTIME": "powershell",
    "CLIENT_ID": "your-client-id",
    "CLIENT_SECRET": "your-client-secret",
    "REDIRECT_URI": "http://localhost:7071/api/callback",
    "AUTHORIZE_URL": "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
  }
}
```

3. Run locally:

```powershell
func start
```

4. Navigate to `http://localhost:7071`


### 4. Domain Verification (Optional)

For custom domains, you may need to verify domain ownership with Microsoft.

The app includes an endpoint at `/.well-known/microsoft-identity-association.json` that returns the application association data.

## Development Notes

- Built with Azure Functions PowerShell runtime
- Uses Microsoft Graph API for To Do data access
- OAuth 2.0 Authorization Code flow with Entra ID
- Base64 encoding for safe JSON embedding in HTML
- Automatic download functionality with JavaScript

## Usage

1. Navigate to `http://localhost:7071/`
2. Click "Authorize and Export"
3. Sign in with Microsoft account
4. Grant permissions
5. Download the JSON backup file

## Architecture

- **LandingPage** (`/`): Landing page with authorization link
- **Callback** (`/api/callback`): OAuth callback handler, exchanges code for token, redirects via POST
- **Export** (`/api/export`): Fetches all lists and tasks, returns JSON with auto-download
- **IdentityAssociation** (`/.well-known/microsoft-identity-association.json`): Microsoft identity association endpoint (uses `CLIENT_ID` env var)

- **Local vs Azure settings**: `local.settings.json` is only for local development. In Azure, set equivalent values as Function App Settings (Application Settings) — do not rely on `local.settings.json` in production.

## Troubleshooting

### 404 Errors
- Ensure function host is running: `func host start`
- Check that all functions are loaded without errors
- Verify route configurations in `function.json` files

### OAuth Issues
- Confirm redirect URI matches exactly in App Registration
- Check that API permissions are granted (`Tasks.Read`)
- Verify client ID and secret are correct in app settings

## Exported Data Format

```json
{
  "exportDate": "2026-02-11T00:00:00Z",
  "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#me/todo",
  "lists": {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#me/todo/lists",
    "@odata.count": 1,
    "value": [
      {
        "@odata.etag": "W/\"etag-list-example-1\"",
        "id": "example-list-id-1",
        "displayName": "Example List",
        "isOwner": true,
        "isShared": false,
        "wellknownListName": "defaultList",
        "createdDateTime": "2026-02-01T00:00:00Z",
        "lastModifiedDateTime": "2026-02-11T00:00:00Z"
      }
    ]
  },
  "tasks": {
    "example-list-id-1": {
      "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#me/todo/lists('example-list-id-1')/tasks",
      "@odata.count": 1,
      "value": [
        {
          "@odata.etag": "W/\"etag-task-example-1\"",
          "id": "example-task-id-1",
          "title": "Example Task 1",
          "body": {
            "content": "Example task details.",
            "contentType": "text"
          },
          "bodyLastModifiedDateTime": "2026-02-02T08:00:00Z",
          "status": "notStarted",
          "importance": "normal",
          "createdDateTime": "2026-02-01T12:00:00Z",
          "lastModifiedDateTime": "2026-02-11T12:00:00Z",
          "completedDateTime": null,
          "dueDateTime": {
            "dateTime": "2026-02-15T17:00:00",
            "timeZone": "UTC"
          },
          "reminderDateTime": null,
          "isReminderOn": false,
          "recurrence": null,
          "categories": [],
          "linkedResources": [],
          "attachments": []
        }
      ]
    }
  }
}
```

[![License](https://img.shields.io/badge/License-GNU-blue.svg)](LICENSE.md)
[![Privacy Policy](https://img.shields.io/badge/Privacy-Policy-green.svg)](PRIVACY_POLICY.md)