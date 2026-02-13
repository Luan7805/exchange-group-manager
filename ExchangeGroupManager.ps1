# --- Connection and Security ---------------------------------
$CLIENT_ID = $env:CLIENT_ID
$ORGANIZATION = $env:ORGANIZATION
$CERT_PATH = $env:CERT_PATH
$API_TOKEN = $env:API_TOKEN

# Check if all required environment variables are set
$missing = @()
if (-not $CLIENT_ID) { $missing += "CLIENT_ID" }
if (-not $ORGANIZATION) { $missing += "ORGANIZATION" }
if (-not $CERT_PATH) { $missing += "CERT_PATH" }
if (-not $API_TOKEN) { $missing += "API_TOKEN" }

if ($missing.Count -gt 0) {
    Write-Error "Missing environment variable(s): $($missing -join ', '). Please set them before running the script."
    exit 1
}

# --- Functions -----------------------------------------------

function Send-JsonResponse {
    param(
        [System.Net.HttpListenerResponse]$Response,
        [int]$StatusCode,
        [hashtable]$Data
    )
    
    $Response.StatusCode = $StatusCode
    $Response.ContentType = "application/json; charset=utf-8"
    
    $json = $Data | ConvertTo-Json -Depth 10
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
    $Response.OutputStream.Write($buffer, 0, $buffer.Length)
}

function Test-GroupExists {
    param(
        [string]$GroupIdentity
    )
    
    $result = @{
        Exists = $false
        Type   = $null
        Object = $null
    }
    
    # Try Distribution Group
    $groupObj = Get-DistributionGroup -Identity $GroupIdentity -ErrorAction SilentlyContinue
    if ($groupObj) {
        $result.Exists = $true
        $result.Type = "DistributionGroup"
        $result.Object = $groupObj
        return $result
    }
    
    # Try Unified Group
    $groupObj = Get-UnifiedGroup -Identity $GroupIdentity -ErrorAction SilentlyContinue
    if ($groupObj) {
        $result.Exists = $true
        $result.Type = "UnifiedGroup"
        $result.Object = $groupObj
        return $result
    }
    
    # Try Dynamic Distribution Group (not supported, but we need to detect it)
    $groupObj = Get-DynamicDistributionGroup -Identity $GroupIdentity -ErrorAction SilentlyContinue
    if ($groupObj) {
        $result.Exists = $true
        $result.Type = "DynamicDistributionGroup"
        $result.Object = $groupObj
        return $result
    }
    
    return $result
}

function Test-RecipientExists {
    param(
        [string]$Email
    )
    
    $recipient = Get-Recipient -Identity $Email -ErrorAction SilentlyContinue
    return ($null -ne $recipient)
}

function Invoke-GroupMemberOperation {
    param(
        [string]$GroupType,
        [string]$Action,
        [string]$Group,
        [string]$Member
    )
    
    $result = @{
        member  = $Member
        action  = $Action
        status  = "success"
        message = ""
    }
    
    try {
        if ($GroupType -eq "DistributionGroup") {
            if ($Action -eq "add") {
                Add-DistributionGroupMember -Identity $Group -Member $Member -Confirm:$false -BypassSecurityGroupManagerCheck -ErrorAction Stop
                $result.message = "Member added to distribution group successfully"
            }
            elseif ($Action -eq "remove") {
                Remove-DistributionGroupMember -Identity $Group -Member $Member -Confirm:$false -BypassSecurityGroupManagerCheck -ErrorAction Stop
                $result.message = "Member removed from distribution group successfully"
            }
        }
        elseif ($GroupType -eq "UnifiedGroup") {
            if ($Action -eq "add") {
                Add-UnifiedGroupLinks -Identity $Group -LinkType "Members" -Links $Member -Confirm:$false -ErrorAction Stop
                $result.message = "Member added to unified group successfully"
            }
            elseif ($Action -eq "remove") {
                Remove-UnifiedGroupLinks -Identity $Group -LinkType "Members" -Links $Member -Confirm:$false -ErrorAction Stop
                $result.message = "Member removed from unified group successfully"
            }
        }
    }
    catch {
        $result.status = "error"
        $result.message = $_.Exception.Message
    }
    
    return $result
}

function Invoke-GroupRequest {
    param(
        [PSCustomObject]$Params,
        [System.Net.HttpListenerResponse]$Response
    )
    
    $action = $Params.action
    $members = $Params.members
    $group = $Params.group
    
    # Validate parameters (before connecting to Exchange)
    if (-not $action -or -not $members -or -not $group) {
        Send-JsonResponse -Response $Response -StatusCode 400 -Data @{
            success = $false
            error   = "Incomplete parameters in JSON: action, members (array), and group are required"
        }
        return
    }
    
    if ($action -notin @("add", "remove")) {
        Send-JsonResponse -Response $Response -StatusCode 400 -Data @{
            success = $false
            error   = "Invalid action: must be 'add' or 'remove'"
        }
        return
    }
    
    if ($members -isnot [array] -or $members.Count -eq 0) {
        Send-JsonResponse -Response $Response -StatusCode 400 -Data @{
            success = $false
            error   = "Members must be a non-empty array of strings"
        }
        return
    }
    
    # Connect to Exchange Online BEFORE any validation that uses Exchange cmdlets
    try {
        Connect-ExchangeOnline -CertificateFilePath $CERT_PATH -AppID $CLIENT_ID -Organization $ORGANIZATION -ShowBanner:$false -ErrorAction Stop
    }
    catch {
        Send-JsonResponse -Response $Response -StatusCode 500 -Data @{
            success = $false
            error   = "Failed to connect to Exchange Online: $($_.Exception.Message)"
        }
        return
    }
    
    try {
        # Validate group exists (requires Exchange connection)
        $groupInfo = Test-GroupExists -GroupIdentity $group
        
        if (-not $groupInfo.Exists) {
            Send-JsonResponse -Response $Response -StatusCode 422 -Data @{
                success = $false
                error   = "Group not found"
                group   = $group
            }
            return
        }
        
        # Validate group type is supported
        if ($groupInfo.Type -notin @("DistributionGroup", "UnifiedGroup")) {
            Send-JsonResponse -Response $Response -StatusCode 422 -Data @{
                success = $false
                error   = "Unsupported group type"
                group   = $group
                type    = $groupInfo.Type
            }
            return
        }
        
        # Validate all members exist (requires Exchange connection)
        foreach ($member in $members) {
            if (-not (Test-RecipientExists -Email $member)) {
                Send-JsonResponse -Response $Response -StatusCode 422 -Data @{
                    success = $false
                    error   = "Recipient not found"
                    email   = $member
                }
                return
            }
        }
        
        # Process all members
        $results = @()
        foreach ($member in $members) {
            $opResult = Invoke-GroupMemberOperation -GroupType $groupInfo.Type -Action $action -Group $group -Member $member
            $results += $opResult
        }
        
        # Send success response
        Send-JsonResponse -Response $Response -StatusCode 200 -Data @{
            success = $true
            results = $results
        }
        
    }
    catch {
        # Catch any unexpected errors
        Send-JsonResponse -Response $Response -StatusCode 500 -Data @{
            success = $false
            error   = "Internal server error: $($_.Exception.Message)"
        }
    }
    finally {
        # Always disconnect
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
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
                Send-JsonResponse -Response $response -StatusCode 401 -Data @{
                    success = $false
                    error   = "Unauthorized: Invalid or missing authorization token"
                }
                $response.OutputStream.Close()
                continue
            }

            # Read the POST request body
            $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
            $body = $reader.ReadToEnd()
            $reader.Close()

            try {
                $params = $body | ConvertFrom-Json
                Invoke-GroupRequest -Params $params -Response $response
            }
            catch {
                Send-JsonResponse -Response $response -StatusCode 500 -Data @{
                    success = $false
                    error   = "Internal server error: $($_.Exception.Message)"
                }
            }
        }
        else {
            # Method not supported
            Send-JsonResponse -Response $response -StatusCode 405 -Data @{
                success = $false
                error   = "Method not allowed. Use POST"
            }
        }

        $response.OutputStream.Close()
    }
}
finally {
    $listener.Stop()
    Write-Host "Web server stopped." -ForegroundColor Yellow
}
