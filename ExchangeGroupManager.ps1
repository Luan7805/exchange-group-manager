# --- Connection and Security ---------------------------------
$CLIENT_ID    = $env:CLIENT_ID
$ORGANIZATION = $env:ORGANIZATION
$CERT_PATH    = $env:CERT_PATH
$API_TOKEN    = $env:API_TOKEN

# Check if all required environment variables are set
$missing = @()
if (-not $CLIENT_ID)    { $missing += "CLIENT_ID" }
if (-not $ORGANIZATION) { $missing += "ORGANIZATION" }
if (-not $CERT_PATH)    { $missing += "CERT_PATH" }
if (-not $API_TOKEN)    { $missing += "API_TOKEN" }

if ($missing.Count -gt 0) {
    Write-Error "Missing environment variable(s): $($missing -join ', '). Please set them before running the script."
    exit 1
}

# --- Web Server ----------------------------------------------
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://+:8080/")
$listener.Start()
Write-Host "Web server started. Listening on http://localhost:8080/" -ForegroundColor Yellow

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response

        if ($request.HttpMethod -eq "POST") {
            # Check authorization header
            $authHeader = $request.Headers["Authorization"]
            if ($authHeader -ne $API_TOKEN) {
                $response.StatusCode = 401
                $response.ContentType = "text/plain"
                $buffer = [System.Text.Encoding]::UTF8.GetBytes("Unauthorized: Invalid or missing authorization token.")
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
                continue  # Proceed to the next request
            }

            # Read the POST request body
            $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
            $body = $reader.ReadToEnd()
            $reader.Close()

            # Assume the body is JSON with keys: action, members (array), group
            try {
                $params = $body | ConvertFrom-Json
                $action = $params.action
                $members = $params.members
                $group = $params.group

                # Validate parameters
                if (-not $action -or -not $members -or -not $group) {
                    throw "Incomplete parameters in JSON: action, members (array), and group are required."
                }
                if ($action -notin @("add", "remove")) {
                    throw "Invalid action: must be 'add' or 'remove'."
                }
                if ($members -isnot [array] -or $members.Count -eq 0) {
                    throw "Members must be a non-empty array of strings."
                }

                # Connect to Exchange Online
                Connect-ExchangeOnline -CertificateFilePath $CERT_PATH -AppID $CLIENT_ID -Organization $ORGANIZATION -ShowBanner:$false

                # Try to get the group as Distribution Group
                $groupObj = Get-DistributionGroup -Identity $group -ErrorAction SilentlyContinue

                if ($groupObj) {
                    $groupType = "DistributionGroup"
                } else {
                    # If not Distribution Group, try as Unified Group
                    $groupObj = Get-UnifiedGroup -Identity $group -ErrorAction SilentlyContinue
                    if ($groupObj) {
                        $groupType = "UnifiedGroup"
                    } else {
                        throw "Group not found or unsupported: $group"
                    }
                }

                # Collect results
                $results = @()

                foreach ($member in $members) {
                    try {
                        if ($groupType -eq "DistributionGroup") {
                            if ($action -eq "add") {
                                Add-DistributionGroupMember -Identity $group -Member $member -Confirm:$false -BypassSecurityGroupManagerCheck -ErrorAction Stop
                                $results += "User $member added to Distribution group $group successfully."
                            } elseif ($action -eq "remove") {
                                Remove-DistributionGroupMember -Identity $group -Member $member -Confirm:$false -BypassSecurityGroupManagerCheck -ErrorAction Stop
                                $results += "User $member removed from Distribution group $group successfully."
                            }
                        } elseif ($groupType -eq "UnifiedGroup") {
                            if ($action -eq "add") {
                                Add-UnifiedGroupLinks -Identity $group -LinkType "Members" -Links $member -Confirm:$false -ErrorAction Stop
                                $results += "User $member added to Unified group $group successfully."
                            } elseif ($action -eq "remove") {
                                Remove-UnifiedGroupLinks -Identity $group -LinkType "Members" -Links $member -Confirm:$false -ErrorAction Stop
                                $results += "User $member removed from Unified group $group successfully."
                            }
                        }
                    } catch {
                        $results += "Error processing ${member}: $_"
                    }
                }

                # Disconnect
                Disconnect-ExchangeOnline -Confirm:$false

                # Send success response with all results
                $result = $results -join "`n"
                $response.StatusCode = 200
                $response.ContentType = "text/plain"
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($result)
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
            } catch {
                # Send error response
                $errorMsg = "Error: $_"
                $response.StatusCode = 500
                $response.ContentType = "text/plain"
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($errorMsg)
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
            }
        } else {
            # Method not supported
            $response.StatusCode = 405
            $buffer = [System.Text.Encoding]::UTF8.GetBytes("Method not allowed. Use POST.")
            $response.OutputStream.Write($buffer, 0, $buffer.Length)
        }

        $response.OutputStream.Close()
    }
} finally {
    $listener.Stop()
    Write-Host "Web server stopped." -ForegroundColor Yellow
}
