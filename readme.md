# Exchange Group Manager HTTP Server

The goal of this project is to create an alternative for managing members of `Mail-enabled security groups` and `Distribution groups` in Exchange Online. According to Microsoft Graph documentation, these group types cannot have their members managed through the API.

This PowerShell script sets up a simple HTTP server that allows managing members of these groups through POST requests. It supports adding or removing multiple members in a single request, with basic authentication through an authorization header.

## Features
- Handles "add" and "remove" actions for group members
- Supports `DistributionGroup` and `UnifiedGroup` types
- Validates group existence before attempting operations
- Validates group type (rejects dynamic groups and unsupported types)
- Validates member email existence before processing
- Returns structured JSON responses with appropriate HTTP status codes
- Automatic connection and disconnection from Exchange Online

## Environment variables
  - `CLIENT_ID`: The Application ID of the registered app in Azure AD
  - `ORGANIZATION`: The tenant organization (e.g., "yourdomain.onmicrosoft.com")
  - `CERT_PATH`: Full path to the certificate file (.pfx) for authentication inside the container
  - `API_TOKEN`: The secret token used for Authorization header (e.g., "my-secret-key" or "Bearer xyz123")

## Usage
- Send a POST request to `http://localhost:8080/` (or the appropriate host/port)
- Include the `Authorization` header with the value matching your `API_TOKEN`
- The request body must be JSON with the following structure:
  - `action`: "add" or "remove" (string)
  - `group`: The identity of the group (e.g., email address)
  - `members`: An array of member identities (e.g., email addresses)

### Example Request (using curl)
```bash
curl -X POST http://localhost:8080/ \
-H "Authorization: your-secret-token" \
-H "Content-Type: application/json" \
-d '{
  "action": "add",
  "group": "group@example.com",
  "members": ["user1@example.com", "user2@example.com"]
}'
```

## HTTP Response Codes

| Code | Description |
|------|-------------|
| **200** | Operation successful |
| **400** | Bad request (missing parameters, invalid JSON) |
| **401** | Unauthorized (invalid or missing token) |
| **405** | Method not allowed (non-POST request) |
| **422** | Validation error (group not found, unsupported type, recipient not found) |
| **500** | Internal server error (connection failure, unexpected error) |

## Response Examples

### Success (200 OK)
```json
{
  "success": true,
  "results": [
    {
      "member": "user1@example.com",
      "action": "add",
      "status": "success",
      "message": "Member added to distribution group successfully"
    },
    {
      "member": "user2@example.com",
      "action": "add",
      "status": "success",
      "message": "Member added to distribution group successfully"
    }
  ]
}
```

### Group Not Found (422 Unprocessable Entity)
```json
{
  "success": false,
  "error": "Group not found",
  "group": "nonexistent@example.com"
}
```

### Unsupported Group Type (422 Unprocessable Entity)
```json
{
  "success": false,
  "error": "Unsupported group type",
  "group": "dynamicgroup@example.com",
  "type": "DynamicDistributionGroup"
}
```

### Recipient Not Found (422 Unprocessable Entity)
```json
{
  "success": false,
  "error": "Recipient not found",
  "email": "invalid@example.com"
}
```

### Unauthorized (401 Unauthorized)
```json
{
  "success": false,
  "error": "Unauthorized: Invalid or missing authorization token"
}
```

### Invalid Parameters (400 Bad Request)
```json
{
  "success": false,
  "error": "Incomplete parameters in JSON: action, members (array), and group are required"
}
```

### Invalid Action (400 Bad Request)
```json
{
  "success": false,
  "error": "Invalid action: must be 'add' or 'remove'"
}
```

### Connection Error (500 Internal Server Error)
```json
{
  "success": false,
  "error": "Failed to connect to Exchange Online: <error details>"
}
```

## Validation Flow

The server performs the following validations in order:

1. **Parameter Validation** (before connecting to Exchange)
   - Checks if `action`, `members`, and `group` are provided
   - Validates `action` is either "add" or "remove"
   - Validates `members` is a non-empty array

2. **Exchange Connection**
   - Connects to Exchange Online using certificate authentication
   - Returns 500 error if connection fails

3. **Group Validation** (requires Exchange connection)
   - Checks if the group exists
   - Returns 422 error if group not found

4. **Group Type Validation**
   - Verifies the group is either `DistributionGroup` or `UnifiedGroup`
   - Returns 422 error if group type is not supported (e.g., dynamic groups)

5. **Member Validation** (requires Exchange connection)
   - Checks if each member email exists as a recipient
   - Returns 422 error if any member is not found

6. **Operation Processing**
   - Adds or removes each member from the group
   - Returns individual results for each member

7. **Cleanup**
   - Automatically disconnects from Exchange Online (always executed)