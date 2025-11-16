# Exchange Group Manager HTTP Server

The goal of this project is to create an alternative for managing members of `Mail-enabled security groups` and `Distribution groups` in Exchange Online. According to Microsoft Graph documentation, these group types cannot have their members managed through the API.

This PowerShell script sets up a simple HTTP server that allows managing members of these groups through POST requests. It supports adding or removing multiple members in a single request, with basic authentication through an authorization header.

## Features
- Handles "add" and "remove" actions for group members
- Supports `Mail-enabled security groups` and `Distribution groups`
- Returns results as plain text in the HTTP response

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

### Response
- **200 OK**: Plain text with success messages for each member (one per line). Errors for individual members are included if they occur
- **401 Unauthorized**: If the Authorization header is missing or incorrect
- **405 Method Not Allowed**: For non-POST requests
- **500 Internal Server Error**: For validation failures or connection issues

Example success response:
```
User user1@example.com added to Distribution group group@example.com successfully.
User user2@example.com added to Distribution group group@example.com successfully.
```