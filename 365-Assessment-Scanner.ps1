# Function usage examples and documentation
<#
.SYNOPSIS
    Comprehensive Microsoft 365 Assessment and Azure App Registration Tool

.DESCRIPTION
    This script contains multiple functions for:
    
    1. New-Cert: Creates a self-signed certificate and exports it to PFX and CER files
    2. New-AzureAppRegistration: Creates an Azure app registration with specific permissions using a certificate
    3. New-CertificateAndAzureApp: Combines both operations in a single workflow
    4. Invoke-M365Assessment: Runs Microsoft 365 Assessment tool with various assessment types
    5. Show-AssessmentMenu: Interactive menu for running assessments

.ASSESSMENT TYPES SUPPORTED
    - Workflow 2013: Assess SharePoint 2013 workflows for Power Automate migration
    - InfoPath: Assess InfoPath forms for modernization
    - Add-ins and ACS: Assess SharePoint Add-ins and Azure ACS dependencies
    - Alerts: Assess SharePoint alerts usage

.ASSESSMENT OPERATIONS
    - Execute: Start a new assessment
    - Status: Check status of running assessments
    - Report: Generate Power BI or CSV reports from completed assessments

.PERMISSIONS GRANTED TO CREATED APPLICATION
    Microsoft Graph:
    - Sites.Read.All (Application)
    - Application.Read.All (Application)
    
    SharePoint:
    - Sites.Read.All (Application)
    - Sites.Manage.All (Application)
    - Sites.FullControl.All (Application)

.EXAMPLE
    # Create certificate and M365 Assessment app in one step (recommended)
    $result = New-CertificateAndAzureApp -AppDisplayName "My M365 Assessment App"
    
.EXAMPLE
    # Manual approach: Create certificate first, then app, then configure
    $certInfo = New-Cert
    $appResult = New-AzureAppRegistration -AppDisplayName "My Assessment App" -CertificateFilePath $certInfo.CerFilePath
    Set-M365AssessmentApp -AppId $appResult.ApplicationId -CertThumbprint $certInfo.Thumbprint
    
.EXAMPLE
    # Run Workflow 2013 assessment
    Invoke-M365Assessment -AssessmentType "Workflow" -Operation "Execute"
    
.EXAMPLE
    # Check assessment status
    Invoke-M365Assessment -AssessmentType "Workflow" -Operation "Status"
    
.EXAMPLE
    # Generate report for specific assessment
    Invoke-M365Assessment -AssessmentType "Workflow" -Operation "Report" -AssessmentId "22989c75-f08f-4af9-8857-6f19e333d6d3"

.EXAMPLE
    # Show interactive assessment menu
    Show-AssessmentMenu

.PREREQUISITES
    - Microsoft 365 Assessment tool downloaded from https://github.com/pnp/pnpassessment/releases
    - Tool placed at path specified in $assessmentToolPath variable
    - Global Administrator or Application Administrator role in Azure AD
    - Appropriate permissions granted to the Azure application

Required Permissions for App Creation:
Graph: Application: AppRoleAssignment.ReadWrite.All 
Graph: Application: Directory.ReadWrite.All
Graph: Application: Application.ReadWrite.All

.NOTES
    - Assessment tool path can be configured via $assessmentToolPath variable
    - Reports are saved to path specified in $assessmentReportsPath variable
    - The workflow creates TWO applications:
      1. Existing app with permissions to create other apps (uses variables $appID, $thumbprint, etc.)
      2. New M365 Assessment app (dynamically created, uses $global:m365AssessmentAppID, $global:m365AssessmentThumbprint)
    - Use New-CertificateAndAzureApp for automatic setup of M365 Assessment app
    - Use Set-M365AssessmentApp to manually configure if app was created separately
    - New certificates are stored in CurrentUser store, not LocalMachine
    - Some permissions may require additional admin consent in the Azure portal

 AUTHOR
    Michael Lee
    Date: 10/2/25
#>

# --------------------------------------------
# Set Variables
# ----------------------------------------------

# Tenant Information
$tenantname = "m365cpi13246019.onmicrosoft.com"                 # This is your tenant name
$sharepointTenantUrl = "m365cpi13246019.sharepoint.com"        # This is your SharePoint tenant URL
$tenantid = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"              # This is your Tenant ID

# EXISTING App with Permissions to Create New Apps (for New-AzureAppRegistration function)
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"                 # Existing App ID with creation permissions
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082"        # Existing app certificate thumbprint
$CertStoreLocation = "LocalMachine"                             # Where the existing app certificate is stored

# NEW M365 Assessment App Variables (will be populated dynamically when app is created)
$global:m365AssessmentAppID = $null                             # New M365 Assessment App ID (populated after creation)
$global:m365AssessmentThumbprint = $null                        # New M365 Assessment app certificate thumbprint (populated after creation)
$global:m365AssessmentCertStoreLocation = "CurrentUser"          # Where the new M365 Assessment certificate is stored

#App and Cert Variables (for creating new apps)
$appname = "M365 Assessment App"
$certname = "m365assessmentcert"
$certpwd = "Pass@word1" # Password for exporting PFX file
$certexportpath = "c:\temp"

# Microsoft 365 Assessment Tool Variables
$assessmentToolPath = "c:\temp\microsoft365-assessment.exe"     # Path to Microsoft 365 Assessment tool executable
$assessmentReportsPath = "c:\temp\assessmentreports"            # Path where assessment reports will be stored


# Script Variables
$currentDateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$logPath = "$env:TEMP\Create-AzureApp_$currentDateTime.log"

# Initialize global variables for the Graph token
$global:graphToken = $null

# Function to write log entries
Function Write-LogEntry {
    param(
        [string] $LogName,
        [string] $LogEntryText,
        [string] $LogLevel = "INFO"
    )
    
    if ($LogName -ne $null) {
        "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToShortTimeString()) : [$LogLevel] $LogEntryText" | Out-File -FilePath $LogName -Append
    }
}

# Function to authenticate with Microsoft Graph API and retrieve an access token
function Get-GraphAccessToken {
    try {
        # Get the certificate from the local certificate store using the thumbprint
        $certificate = Get-Item Cert:\$CertStoreLocation\My\$Thumbprint -ErrorAction Stop

        # Define the URI for authentication
        $uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        
        # Create the client assertion
        $jwtHeader = @{
            alg = "RS256"
            typ = "JWT"
            x5t = [System.Convert]::ToBase64String($certificate.GetCertHash())
        }
        
        $now = [DateTime]::UtcNow
        $jwtExpiry = [Math]::Floor(([DateTimeOffset]$now.AddMinutes(10)).ToUnixTimeSeconds())
        $jwtNbf = [Math]::Floor(([DateTimeOffset]$now).ToUnixTimeSeconds())
        $jwtIssueTime = [Math]::Floor(([DateTimeOffset]$now).ToUnixTimeSeconds())
        $jwtPayload = @{
            aud = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
            exp = $jwtExpiry
            iss = $AppId
            jti = [guid]::NewGuid().ToString()
            nbf = $jwtNbf
            sub = $AppId
            iat = $jwtIssueTime
        }
        
        # Convert to Base64
        $jwtHeaderBase64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(($jwtHeader | ConvertTo-Json -Compress)))
        $jwtHeaderBase64 = $jwtHeaderBase64.TrimEnd('=').Replace('+', '-').Replace('/', '_')
        
        $jwtPayloadBase64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(($jwtPayload | ConvertTo-Json -Compress)))
        $jwtPayloadBase64 = $jwtPayloadBase64.TrimEnd('=').Replace('+', '-').Replace('/', '_')
        
        # Sign the JWT
        $toSign = [System.Text.Encoding]::UTF8.GetBytes($jwtHeaderBase64 + "." + $jwtPayloadBase64)
        $rsa = $certificate.PrivateKey
        $signature = $rsa.SignData($toSign, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        
        # Convert signature to Base64
        $signatureBase64 = [System.Convert]::ToBase64String($signature)
        $signatureBase64 = $signatureBase64.TrimEnd('=').Replace('+', '-').Replace('/', '_')
        
        # Create the complete JWT token
        $jwt = $jwtHeaderBase64 + "." + $jwtPayloadBase64 + "." + $signatureBase64
        
        # Define the body for the authentication request
        $body = @{
            client_id             = $AppId
            client_assertion      = $jwt
            client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
            scope                 = "https://graph.microsoft.com/.default"
            grant_type            = "client_credentials"
        }
        
        # Send the authentication request and extract the token
        $loginResponse = Invoke-RestMethod -Method Post -Uri $uri -Body $body -ContentType 'application/x-www-form-urlencoded'
        $global:graphToken = $loginResponse.access_token
        Write-LogEntry -LogName $logPath -LogEntryText "Successfully authenticated with Microsoft Graph API using certificate." -LogLevel "INFO"
        Write-Host -ForegroundColor Green "Successfully authenticated with Microsoft Graph API using certificate."
        return $global:graphToken
    }
    catch {
        Write-LogEntry -LogName $logPath -LogEntryText "Authentication failed with Microsoft Graph API: $_" -LogLevel "ERROR"
        Write-Host -ForegroundColor Red "Authentication failed with Microsoft Graph API: $_"
        throw
    }
}

# Function to get OneDrive URL for a specific user by UPN
function New-Cert {
    #Create Cert
    $currentdate = Get-Date
    $enddate = $currentdate.AddYears(5)
    $notafter = $enddate.AddYears(5)
    $cert = (New-SelfSignedCertificate -CertStoreLocation "Cert:\CurrentUser\My" -DnsName $tenantname -KeyExportPolicy Exportable -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" -NotAfter $notafter)
    $certpwd = ConvertTo-SecureString -String $certpwd -Force -AsPlainText

    #export to PFX
    $certpfxfile = Export-PfxCertificate -Cert $cert -FilePath "$certexportpath\$certname.pfx" -Password $certpwd

    #Export to CER
    $certcerfile = Export-Certificate -Cert $cert -FilePath "$certexportpath\$certname.cer"

    Write-Host "This is your ThumbPrint: $($cert.Thumbprint)" -ForegroundColor Green
    Write-Host "PFX file $certpfxfile exported to $certexportpath\$certname.pfx" -ForegroundColor Green
    Write-Host "CER file $certcerfile exported to $certexportpath\$certname.cer" -ForegroundColor Green

    # Return certificate information for use in app registration
    return @{
        Certificate = $cert
        Thumbprint  = $cert.Thumbprint
        CerFilePath = "$certexportpath\$certname.cer"
        PfxFilePath = "$certexportpath\$certname.pfx"
    }
}

# Function to run Microsoft 365 Assessment Tool
function Invoke-M365Assessment {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Workflow", "InfoPath", "AddInsACS", "Alerts")]
        [string]$AssessmentType,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet("Execute", "Status", "Report")]
        [string]$Operation,
        
        [Parameter(Mandatory = $false)]
        [string]$AssessmentId = "",
        
        [Parameter(Mandatory = $false)]
        [string]$SitesList = "",
        
        [Parameter(Mandatory = $false)]
        [string]$CustomReportPath = "",
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("PowerBI", "CsvOnly")]
        [string]$ReportMode = "PowerBI",
        
        [Parameter(Mandatory = $false)]
        [string]$AppId = $global:m365AssessmentAppID,
        
        [Parameter(Mandatory = $false)]
        [string]$CertThumbprint = $global:m365AssessmentThumbprint,
        
        [Parameter(Mandatory = $false)]
        [string]$TenantName = $global:sharepointTenantUrl
    )
    
    try {
        Write-LogEntry -LogName $logPath -LogEntryText "Starting Microsoft 365 Assessment - Type: $AssessmentType, Operation: $Operation" -LogLevel "INFO"
        
        # Validate that M365 Assessment app has been created
        if (-not $AppId -or -not $CertThumbprint) {
            Write-Host -ForegroundColor Red "❌ M365 Assessment application not configured!"
            Write-Host -ForegroundColor Yellow "💡 Please create the M365 Assessment application first using:"
            Write-Host -ForegroundColor Cyan "   1. New-Cert (to create certificate)"
            Write-Host -ForegroundColor Cyan "   2. New-AzureAppRegistration (to create app)"
            Write-Host -ForegroundColor Cyan "   3. Update the global variables with the new app details"
            return @{ Success = $false; Error = "M365 Assessment app not configured" }
        }
        
        # Check if assessment tool exists
        if (-not (Test-Path $assessmentToolPath)) {
            Write-Host -ForegroundColor Red "❌ Microsoft 365 Assessment tool not found at: $assessmentToolPath"
            Write-Host -ForegroundColor Yellow "Please download from: https://github.com/pnp/pnpassessment/releases"
            return @{ Success = $false; Error = "Assessment tool not found" }
        }
        
        # Create reports directory if it doesn't exist
        if (-not (Test-Path $assessmentReportsPath)) {
            New-Item -ItemType Directory -Path $assessmentReportsPath -Force | Out-Null
            Write-Host -ForegroundColor Green "Created reports directory: $assessmentReportsPath"
        }
        
        # Build the command based on operation
        $command = ""
        
        switch ($Operation) {
            "Execute" {
                Write-Host -ForegroundColor Yellow "🚀 Starting $AssessmentType assessment..."
                
                # Build the start command
                $command = "& `"$assessmentToolPath`" start --mode $AssessmentType --authmode application --tenant $TenantName --applicationid $AppId --certpath `"My|CurrentUser|$CertThumbprint`""
                
                # Add sites list if provided
                if ($SitesList) {
                    $command += " --siteslist `"$SitesList`""
                    Write-Host -ForegroundColor Cyan "🎯 Targeting specific sites: $SitesList"
                }
                else {
                    Write-Host -ForegroundColor Cyan "🎯 Targeting entire tenant: $TenantName"
                }
                
                Write-Host -ForegroundColor White "Using Application ID: $AppId"
                Write-Host -ForegroundColor White "Using Certificate Thumbprint: $CertThumbprint"
                
                Write-Host -ForegroundColor White "Command: $command"
                Write-Host -ForegroundColor Yellow "⏳ This may take a while depending on tenant size..."
                
                try {
                    # Use Start-Process for better console handling
                    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
                    $processInfo.FileName = $assessmentToolPath
                    $arguments = "start --mode $AssessmentType --authmode application --tenant $TenantName --applicationid $AppId --certpath `"My|CurrentUser|$CertThumbprint`""
                    if ($SitesList) {
                        $arguments += " --siteslist `"$SitesList`""
                    }
                    $processInfo.Arguments = $arguments
                    $processInfo.RedirectStandardOutput = $true
                    $processInfo.RedirectStandardError = $true
                    $processInfo.UseShellExecute = $false
                    $processInfo.CreateNoWindow = $true
                    
                    $process = New-Object System.Diagnostics.Process
                    $process.StartInfo = $processInfo
                    $process.Start() | Out-Null
                    
                    # Read output
                    $stdout = $process.StandardOutput.ReadToEnd()
                    $stderr = $process.StandardError.ReadToEnd()
                    $process.WaitForExit()
                    
                    $exitCode = $process.ExitCode
                    
                    if ($exitCode -eq 0) {
                        Write-Host -ForegroundColor Green "✅ $AssessmentType assessment started successfully!"
                        Write-Host -ForegroundColor Cyan "📊 Use 'Status' operation to monitor progress"
                        Write-LogEntry -LogName $logPath -LogEntryText "$AssessmentType assessment started successfully" -LogLevel "INFO"
                        
                        if ($stdout) {
                            Write-Host -ForegroundColor Cyan "Output:"
                            Write-Host $stdout
                        }
                        
                        return @{ 
                            Success = $true
                            Message = "$AssessmentType assessment started"
                            Output  = $stdout
                        }
                    }
                    else {
                        $errorMessage = "Assessment failed with exit code: $exitCode"
                        if ($stderr) {
                            $errorMessage += "`nError output: $stderr"
                        }
                        throw $errorMessage
                    }
                }
                catch {
                    # Fallback to cmd.exe execution if process approach fails
                    Write-Host -ForegroundColor Yellow "⚠️  Trying alternative execution method..."
                    try {
                        $cmdArgs = "start --mode $AssessmentType --authmode application --tenant $TenantName --applicationid $AppId --certpath `"My|CurrentUser|$CertThumbprint`""
                        if ($SitesList) {
                            $cmdArgs += " --siteslist `"$SitesList`""
                        }
                        
                        $result = cmd.exe /c "`"$assessmentToolPath`" $cmdArgs 2>&1"
                        
                        if ($LASTEXITCODE -eq 0) {
                            Write-Host -ForegroundColor Green "✅ $AssessmentType assessment started successfully!"
                            Write-Host -ForegroundColor Cyan "📊 Use 'Status' operation to monitor progress"
                            Write-LogEntry -LogName $logPath -LogEntryText "$AssessmentType assessment started successfully" -LogLevel "INFO"
                            
                            return @{ 
                                Success = $true
                                Message = "$AssessmentType assessment started"
                                Output  = $result
                            }
                        }
                        else {
                            throw "Assessment failed with exit code: $LASTEXITCODE. Output: $result"
                        }
                    }
                    catch {
                        throw "Assessment execution failed: $_"
                    }
                }
            }
            
            "Status" {
                Write-Host -ForegroundColor Yellow "📊 Checking assessment status..."
                
                Write-Host -ForegroundColor Cyan "Select status display method:"
                Write-Host -ForegroundColor White "1. Interactive Live Status (opens in new window)"
                Write-Host -ForegroundColor White "2. Quick Status Check (non-interactive)"
                Write-Host -ForegroundColor White "3. List All Assessments"
                $statusChoice = Read-Host "Enter your choice (1-3)"
                
                switch ($statusChoice) {
                    "1" {
                        Write-Host -ForegroundColor Yellow "🔄 Opening interactive status display..."
                        Write-Host -ForegroundColor Cyan "💡 A new command window will open. Press ESC in that window to exit the status display."
                        
                        try {
                            # Start the status command in a new window for interactive display
                            Start-Process -FilePath $assessmentToolPath -ArgumentList "status" -Wait
                            
                            Write-Host -ForegroundColor Green "✅ Interactive status display completed!"
                            return @{ 
                                Success = $true
                                Message = "Interactive status displayed"
                                Output  = "Status displayed in separate window"
                            }
                        }
                        catch {
                            Write-Host -ForegroundColor Red "❌ Failed to open interactive status: $_"
                            return @{ Success = $false; Error = "Failed to open interactive status: $_" }
                        }
                    }
                    
                    "2" {
                        Write-Host -ForegroundColor Yellow "⚡ Performing quick status check..."
                        
                        try {
                            # First try to get the list of assessments to determine current state
                            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
                            $processInfo.FileName = $assessmentToolPath
                            $processInfo.Arguments = "list"
                            $processInfo.RedirectStandardOutput = $true
                            $processInfo.RedirectStandardError = $true
                            $processInfo.UseShellExecute = $false
                            $processInfo.CreateNoWindow = $true
                            
                            $process = New-Object System.Diagnostics.Process
                            $process.StartInfo = $processInfo
                            $process.Start() | Out-Null
                            $process.WaitForExit()
                            
                            $listOutput = $process.StandardOutput.ReadToEnd()
                            $exitCode = $process.ExitCode
                            
                            if ($exitCode -eq 0 -and $listOutput) {
                                Write-Host -ForegroundColor Green "✅ Quick status check completed!"
                                Write-Host -ForegroundColor Cyan "Assessment Status Summary:"
                                
                                # Parse the list output to find running assessments
                                $lines = $listOutput -split "`n"
                                $runningAssessments = @()
                                $finishedAssessments = @()
                                $inTable = $false
                                $currentAssessment = $null
                                
                                foreach ($line in $lines) {
                                    $trimmedLine = $line.Trim()
                                    
                                    # Skip header and formatting lines
                                    if ($trimmedLine -like "*─*" -or $trimmedLine -like "*┌*" -or $trimmedLine -like "*└*" -or 
                                        $trimmedLine -like "*│ Id*" -or $trimmedLine -like "*├*" -or $trimmedLine -eq "") {
                                        if ($trimmedLine -like "*│ Id*") { $inTable = $true }
                                        continue
                                    }
                                    
                                    # Parse assessment data lines
                                    if ($inTable -and $trimmedLine -like "*│*") {
                                        $parts = $trimmedLine -split "│"
                                        if ($parts.Count -ge 4) {
                                            $id = $parts[1].Trim()
                                            $mode = $parts[2].Trim()
                                            $status = $parts[3].Trim()
                                            $progress = if ($parts.Count -ge 5) { $parts[4].Trim() } else { "" }
                                            $startTime = if ($parts.Count -ge 6) { $parts[5].Trim() } else { "" }
                                            $endTime = if ($parts.Count -ge 7) { $parts[6].Trim() } else { "" }
                                            
                                            # Check if this is a new assessment row (has mode and status) or continuation of previous ID
                                            if ($mode -ne "" -and $status -ne "") {
                                                # This is a new assessment row
                                                $currentAssessment = @{
                                                    Id        = $id
                                                    Mode      = $mode
                                                    Status    = $status
                                                    Progress  = $progress
                                                    StartTime = $startTime
                                                    EndTime   = $endTime
                                                }
                                                
                                                # Add to appropriate collection based on status
                                                if ($status -like "*Running*" -or $status -like "*In Progress*") {
                                                    $runningAssessments += $currentAssessment
                                                }
                                                elseif ($status -like "*Finished*" -or $status -like "*Completed*") {
                                                    $finishedAssessments += $currentAssessment
                                                }
                                            }
                                            elseif ($currentAssessment -ne $null -and $id -ne "") {
                                                # This is a continuation line with more ID parts
                                                $currentAssessment.Id += $id
                                                
                                                # Update progress if it contains more information
                                                if ($progress -ne "" -and $progress -ne ")") {
                                                    $currentAssessment.Progress += $progress
                                                }
                                            }
                                        }
                                    }
                                    
                                    # Show version and connection info
                                    if ($trimmedLine -like "*version*" -or $trimmedLine -like "*Connecting*" -or $trimmedLine -eq "OK") {
                                        Write-Host -ForegroundColor White "  $trimmedLine"
                                    }
                                }
                                
                                # Display status summary
                                Write-Host ""
                                if ($runningAssessments.Count -gt 0) {
                                    Write-Host -ForegroundColor Green "🔄 Currently Running Assessments:"
                                    foreach ($assessment in $runningAssessments) {
                                        Write-Host -ForegroundColor Cyan "  • $($assessment.Mode) - $($assessment.Status) - $($assessment.Progress)"
                                        Write-Host -ForegroundColor Yellow "    Assessment ID: " -NoNewline
                                        Write-Host -ForegroundColor White "[$($assessment.Id)]"
                                    }
                                }
                                else {
                                    Write-Host -ForegroundColor Yellow "⏸️  No assessments currently running"
                                }
                                
                                if ($finishedAssessments.Count -gt 0) {
                                    Write-Host -ForegroundColor Cyan "✅ Recently Completed Assessments:"
                                    foreach ($assessment in $finishedAssessments) {
                                        Write-Host -ForegroundColor White "  • $($assessment.Mode) - $($assessment.Status) - $($assessment.Progress)"
                                        Write-Host -ForegroundColor Yellow "    Assessment ID: " -NoNewline
                                        Write-Host -ForegroundColor White "[$($assessment.Id)]"
                                    }
                                    Write-Host ""
                                    Write-Host -ForegroundColor Yellow "💡 Copy Assessment IDs (including brackets) for report generation"
                                    Write-Host -ForegroundColor Cyan "📋 Quick Copy Format:"
                                    foreach ($assessment in $finishedAssessments) {
                                        Write-Host -ForegroundColor Green "   $($assessment.Id)"
                                    }
                                }
                                
                                $statusMessage = if ($runningAssessments.Count -gt 0) {
                                    "$($runningAssessments.Count) running, $($finishedAssessments.Count) completed"
                                }
                                else {
                                    "No active assessments, $($finishedAssessments.Count) completed"
                                }
                                
                                return @{ 
                                    Success              = $true
                                    Message              = "Quick status retrieved"
                                    Output               = $statusMessage
                                    RunningAssessments   = $runningAssessments.Count
                                    CompletedAssessments = $finishedAssessments.Count
                                }
                            }
                            else {
                                Write-Host -ForegroundColor Yellow "⚠️ Unable to retrieve assessment status"
                                Write-Host -ForegroundColor Gray "This might indicate no assessments have been run yet or the assessment service is not running"
                                
                                return @{ 
                                    Success = $true
                                    Message = "No assessments found"
                                    Output  = "No assessments detected"
                                }
                            }
                        }
                        catch {
                            Write-Host -ForegroundColor Red "❌ Quick status check failed: $_"
                            return @{ Success = $false; Error = "Quick status check failed: $_" }
                        }
                    }
                    
                    "3" {
                        Write-Host -ForegroundColor Yellow "📋 Listing all assessments..."
                        
                        try {
                            # Use the list command to show all assessments
                            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
                            $processInfo.FileName = $assessmentToolPath
                            $processInfo.Arguments = "list"
                            $processInfo.RedirectStandardOutput = $true
                            $processInfo.RedirectStandardError = $true
                            $processInfo.UseShellExecute = $false
                            $processInfo.CreateNoWindow = $true
                            
                            $process = New-Object System.Diagnostics.Process
                            $process.StartInfo = $processInfo
                            $process.Start() | Out-Null
                            
                            # Read output
                            $listOutput = $process.StandardOutput.ReadToEnd()
                            $stderr = $process.StandardError.ReadToEnd()
                            $process.WaitForExit()
                            
                            $exitCode = $process.ExitCode
                            
                            if ($exitCode -eq 0 -and $listOutput) {
                                Write-Host -ForegroundColor Green "✅ Assessment list retrieved successfully!"
                                
                                # Parse the list output to extract assessments with complete IDs
                                $lines = $listOutput -split "`n"
                                $allAssessments = @()
                                $inTable = $false
                                $currentAssessment = $null
                                
                                foreach ($line in $lines) {
                                    $trimmedLine = $line.Trim()
                                    
                                    # Show version and connection info first
                                    if ($trimmedLine -like "*version*" -or $trimmedLine -like "*Connecting*" -or $trimmedLine -eq "OK") {
                                        Write-Host -ForegroundColor White "  $trimmedLine"
                                        continue
                                    }
                                    
                                    # Skip header and formatting lines
                                    if ($trimmedLine -like "*─*" -or $trimmedLine -like "*┌*" -or $trimmedLine -like "*└*" -or 
                                        $trimmedLine -like "*│ Id*" -or $trimmedLine -like "*├*" -or $trimmedLine -eq "") {
                                        if ($trimmedLine -like "*│ Id*") { $inTable = $true }
                                        continue
                                    }
                                    
                                    # Parse assessment data lines
                                    if ($inTable -and $trimmedLine -like "*│*") {
                                        $parts = $trimmedLine -split "│"
                                        if ($parts.Count -ge 4) {
                                            $id = $parts[1].Trim()
                                            $mode = $parts[2].Trim()
                                            $status = $parts[3].Trim()
                                            $progress = if ($parts.Count -ge 5) { $parts[4].Trim() } else { "" }
                                            $startTime = if ($parts.Count -ge 6) { $parts[5].Trim() } else { "" }
                                            $endTime = if ($parts.Count -ge 7) { $parts[6].Trim() } else { "" }
                                            
                                            # Check if this is a new assessment row or continuation
                                            if ($mode -ne "" -and $status -ne "") {
                                                # This is a new assessment row
                                                $currentAssessment = @{
                                                    Id        = $id
                                                    Mode      = $mode
                                                    Status    = $status
                                                    Progress  = $progress
                                                    StartTime = $startTime
                                                    EndTime   = $endTime
                                                }
                                                $allAssessments += $currentAssessment
                                            }
                                            elseif ($currentAssessment -ne $null -and $id -ne "") {
                                                # This is a continuation line with more ID parts
                                                $currentAssessment.Id += $id
                                                
                                                # Update progress if it contains more information
                                                if ($progress -ne "" -and $progress -ne ")") {
                                                    $currentAssessment.Progress += $progress
                                                }
                                                
                                                # Update time information if present
                                                if ($startTime -ne "") {
                                                    $currentAssessment.StartTime += " " + $startTime
                                                }
                                                if ($endTime -ne "") {
                                                    $currentAssessment.EndTime += " " + $endTime
                                                }
                                            }
                                        }
                                    }
                                }
                                
                                # Display formatted assessment information
                                Write-Host ""
                                if ($allAssessments.Count -gt 0) {
                                    Write-Host -ForegroundColor Cyan "📊 All Assessments Summary:"
                                    Write-Host ""
                                    
                                    $runningCount = 0
                                    $completedCount = 0
                                    
                                    foreach ($assessment in $allAssessments) {
                                        # Determine status color and category
                                        if ($assessment.Status -like "*Running*" -or $assessment.Status -like "*In Progress*") {
                                            Write-Host -ForegroundColor Yellow "🔄 RUNNING ASSESSMENT"
                                            $runningCount++
                                        }
                                        elseif ($assessment.Status -like "*Finished*" -or $assessment.Status -like "*Completed*") {
                                            Write-Host -ForegroundColor Green "✅ COMPLETED ASSESSMENT"
                                            $completedCount++
                                        }
                                        else {
                                            Write-Host -ForegroundColor Gray "❓ OTHER STATUS"
                                        }
                                        
                                        Write-Host -ForegroundColor White "   Type: $($assessment.Mode)"
                                        Write-Host -ForegroundColor White "   Status: $($assessment.Status)"
                                        Write-Host -ForegroundColor White "   Progress: $($assessment.Progress)"
                                        if ($assessment.StartTime -and $assessment.StartTime.Trim() -ne "") {
                                            Write-Host -ForegroundColor White "   Started: $($assessment.StartTime.Trim())"
                                        }
                                        if ($assessment.EndTime -and $assessment.EndTime.Trim() -ne "") {
                                            Write-Host -ForegroundColor White "   Ended: $($assessment.EndTime.Trim())"
                                        }
                                        Write-Host -ForegroundColor Yellow "   Assessment ID: " -NoNewline
                                        Write-Host -ForegroundColor White "[$($assessment.Id)]"
                                        Write-Host ""
                                    }
                                    
                                    # Show summary
                                    Write-Host -ForegroundColor Cyan "📈 Summary: $($allAssessments.Count) total assessments ($runningCount running, $completedCount completed)"
                                    
                                    # Show copy-paste section for all assessments
                                    Write-Host ""
                                    Write-Host -ForegroundColor Yellow "💡 Copy Assessment IDs for report generation:"
                                    Write-Host -ForegroundColor Cyan "📋 Quick Copy Format (Assessment ID only):"
                                    foreach ($assessment in $allAssessments) {
                                        Write-Host -ForegroundColor Green "   $($assessment.Id)"
                                    }
                                }
                                else {
                                    Write-Host -ForegroundColor Yellow "⚠️ No assessments found"
                                    Write-Host -ForegroundColor Gray "This might indicate no assessments have been run yet"
                                }
                                
                                return @{ 
                                    Success         = $true
                                    Message         = "Assessment list retrieved"
                                    Output          = "Found $($allAssessments.Count) assessments"
                                    AssessmentCount = $allAssessments.Count
                                }
                            }
                            else {
                                Write-Host -ForegroundColor Yellow "⚠️ No assessments found or unable to retrieve list"
                                return @{ 
                                    Success = $true
                                    Message = "No assessments found"
                                    Output  = "No assessments detected"
                                }
                            }
                        }
                        catch {
                            Write-Host -ForegroundColor Red "❌ Failed to list assessments: $_"
                            return @{ Success = $false; Error = "Failed to list assessments: $_" }
                        }
                    }
                    
                    default {
                        Write-Host -ForegroundColor Red "Invalid selection. Returning to menu..."
                        return @{ Success = $false; Error = "Invalid selection" }
                    }
                }
            }
            
            "Report" {
                if (-not $AssessmentId) {
                    Write-Host -ForegroundColor Red "❌ Assessment ID is required for report generation"
                    Write-Host -ForegroundColor Yellow "💡 Use 'microsoft365-assessment.exe list' to get assessment IDs"
                    return @{ Success = $false; Error = "Assessment ID required" }
                }
                
                Write-Host -ForegroundColor Yellow "📋 Generating $AssessmentType assessment report..."
                Write-Host -ForegroundColor Cyan "Assessment ID: $AssessmentId"
                Write-Host -ForegroundColor Cyan "Report Mode: $ReportMode"
                
                # Build the report command arguments
                $arguments = "report --id $AssessmentId"
                
                # Add report mode if CSV only
                if ($ReportMode -eq "CsvOnly") {
                    $arguments += " --mode CsvOnly"
                }
                
                # Clean up Assessment ID (remove brackets if present)
                $cleanAssessmentId = $AssessmentId.Trim('[', ']')
                
                # Create assessment-specific subfolder
                if ($CustomReportPath) {
                    $reportPath = Join-Path $CustomReportPath $cleanAssessmentId
                }
                else {
                    $reportPath = Join-Path $assessmentReportsPath $cleanAssessmentId
                }
                
                # Create the assessment-specific directory if it doesn't exist
                if (-not (Test-Path $reportPath)) {
                    New-Item -ItemType Directory -Path $reportPath -Force | Out-Null
                    Write-Host -ForegroundColor Green "📁 Created assessment report directory: $reportPath"
                }
                
                $arguments += " --path `"$reportPath`""
                
                Write-Host -ForegroundColor White "Command: $assessmentToolPath $arguments"
                Write-Host -ForegroundColor Cyan "Report will be saved to: $reportPath"
                
                try {
                    # Use Start-Process for better console handling
                    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
                    $processInfo.FileName = $assessmentToolPath
                    $processInfo.Arguments = $arguments
                    $processInfo.RedirectStandardOutput = $true
                    $processInfo.RedirectStandardError = $true
                    $processInfo.UseShellExecute = $false
                    $processInfo.CreateNoWindow = $true
                    
                    $process = New-Object System.Diagnostics.Process
                    $process.StartInfo = $processInfo
                    $process.Start() | Out-Null
                    
                    # Read output
                    $stdout = $process.StandardOutput.ReadToEnd()
                    $stderr = $process.StandardError.ReadToEnd()
                    $process.WaitForExit()
                    
                    $exitCode = $process.ExitCode
                    
                    if ($exitCode -eq 0) {
                        Write-Host -ForegroundColor Green "✅ $AssessmentType report generated successfully!"
                        Write-Host -ForegroundColor Cyan "📁 Report location: $reportPath"
                        
                        if ($ReportMode -eq "PowerBI") {
                            Write-Host -ForegroundColor Yellow "📊 Power BI report and CSV files generated"
                        }
                        else {
                            Write-Host -ForegroundColor Yellow "📊 CSV files generated"
                        }
                        
                        Write-LogEntry -LogName $logPath -LogEntryText "$AssessmentType report generated at $reportPath" -LogLevel "INFO"
                        
                        if ($stdout) {
                            Write-Host -ForegroundColor Cyan "Output:"
                            Write-Host $stdout
                        }
                        
                        return @{ 
                            Success    = $true
                            Message    = "Report generated"
                            ReportPath = $reportPath
                            Output     = $stdout
                        }
                    }
                    else {
                        $errorMessage = "Report generation failed with exit code: $exitCode"
                        if ($stderr) {
                            $errorMessage += "`nError output: $stderr"
                        }
                        throw $errorMessage
                    }
                }
                catch {
                    # Fallback to cmd.exe execution if process approach fails
                    Write-Host -ForegroundColor Yellow "⚠️  Trying alternative execution method..."
                    try {
                        $result = cmd.exe /c "`"$assessmentToolPath`" $arguments 2>&1"
                        
                        if ($LASTEXITCODE -eq 0) {
                            Write-Host -ForegroundColor Green "✅ $AssessmentType report generated successfully!"
                            Write-Host -ForegroundColor Cyan "📁 Report location: $reportPath"
                            
                            if ($ReportMode -eq "PowerBI") {
                                Write-Host -ForegroundColor Yellow "📊 Power BI report and CSV files generated"
                            }
                            else {
                                Write-Host -ForegroundColor Yellow "📊 CSV files generated"
                            }
                            
                            Write-LogEntry -LogName $logPath -LogEntryText "$AssessmentType report generated at $reportPath" -LogLevel "INFO"
                            
                            return @{ 
                                Success    = $true
                                Message    = "Report generated"
                                ReportPath = $reportPath
                                Output     = $result
                            }
                        }
                        else {
                            throw "Report generation failed with exit code: $LASTEXITCODE. Output: $result"
                        }
                    }
                    catch {
                        throw "Report generation failed: $_"
                    }
                }
            }
        }
    }
    catch {
        $errorMessage = "Microsoft 365 Assessment failed: $_"
        Write-LogEntry -LogName $logPath -LogEntryText $errorMessage -LogLevel "ERROR"
        Write-Host -ForegroundColor Red $errorMessage
        
        return @{
            Success = $false
            Error   = $errorMessage
        }
    }
}

# Function to show assessment menu and handle user selection
function Show-AssessmentMenu {
    param(
        [Parameter(Mandatory = $false)]
        [string]$QuickMode = "",
        
        [Parameter(Mandatory = $false)]
        [string]$AppId = $global:m365AssessmentAppID,
        
        [Parameter(Mandatory = $false)]
        [string]$CertThumbprint = $global:m365AssessmentThumbprint
    )
    
    if (-not $QuickMode) {
        Write-Host -ForegroundColor Cyan "`n🔍 Microsoft 365 Assessment Tool"
        Write-Host -ForegroundColor Cyan "================================="
        Write-Host -ForegroundColor White "Select Assessment Type:"
        Write-Host -ForegroundColor White "1. Workflow 2013"
        Write-Host -ForegroundColor White "2. InfoPath"
        Write-Host -ForegroundColor White "3. Add-ins and ACS"
        Write-Host -ForegroundColor White "4. Alerts"
        Write-Host -ForegroundColor White "5. Exit"
        Write-Host ""
        
        $assessmentChoice = Read-Host "Enter your choice (1-5)"
        
        $assessmentTypes = @{
            "1" = "Workflow"
            "2" = "InfoPath" 
            "3" = "AddInsACS"
            "4" = "Alerts"
        }
        
        if ($assessmentChoice -eq "5") {
            Write-Host -ForegroundColor Yellow "Exiting assessment menu..."
            return
        }
        
        if (-not $assessmentTypes.ContainsKey($assessmentChoice)) {
            Write-Host -ForegroundColor Red "Invalid selection. Please try again."
            return
        }
        
        $selectedAssessment = $assessmentTypes[$assessmentChoice]
    }
    else {
        $selectedAssessment = $QuickMode
    }
    
    # Show operation menu
    Write-Host -ForegroundColor Cyan "`n⚙️ Select Operation for $selectedAssessment Assessment:"
    Write-Host -ForegroundColor White "1. Execute Assessment"
    Write-Host -ForegroundColor White "2. Check Status"
    Write-Host -ForegroundColor White "3. Generate Report"
    Write-Host -ForegroundColor White "4. Back to Assessment Menu"
    Write-Host ""
    
    $operationChoice = Read-Host "Enter your choice (1-4)"
    
    switch ($operationChoice) {
        "1" {
            Write-Host -ForegroundColor Yellow "`n🎯 Execute $selectedAssessment Assessment"
            $sitesList = Read-Host "Enter specific sites (comma-separated URLs) or press Enter for entire tenant"
            
            $result = Invoke-M365Assessment -AssessmentType $selectedAssessment -Operation "Execute" -SitesList $sitesList -AppId $AppId -CertThumbprint $CertThumbprint
            
            if ($result.Success) {
                Write-Host -ForegroundColor Green "`n✅ Assessment started successfully!"
                Write-Host -ForegroundColor Cyan "💡 Next steps:"
                Write-Host -ForegroundColor White "1. Use 'Check Status' to monitor progress"
                Write-Host -ForegroundColor White "2. Use 'Generate Report' when completed"
            }
        }
        
        "2" {
            Write-Host -ForegroundColor Yellow "`n📊 Checking Assessment Status"
            $result = Invoke-M365Assessment -AssessmentType $selectedAssessment -Operation "Status" -AppId $AppId -CertThumbprint $CertThumbprint
        }
        
        "3" {
            Write-Host -ForegroundColor Yellow "`n📋 Generate $selectedAssessment Report"
            $assessmentId = Read-Host "Enter Assessment ID (use 'microsoft365-assessment.exe list' to get IDs)"
            
            if ($assessmentId) {
                Write-Host -ForegroundColor Cyan "Select Report Type:"
                Write-Host -ForegroundColor White "1. Power BI Report (includes CSV)"
                Write-Host -ForegroundColor White "2. CSV Files Only"
                $reportChoice = Read-Host "Enter choice (1-2)"
                
                $reportMode = if ($reportChoice -eq "2") { "CsvOnly" } else { "PowerBI" }
                $customPath = Read-Host "Enter custom report path or press Enter for default ($assessmentReportsPath)"
                
                $result = Invoke-M365Assessment -AssessmentType $selectedAssessment -Operation "Report" -AssessmentId $assessmentId -ReportMode $reportMode -CustomReportPath $customPath -AppId $AppId -CertThumbprint $CertThumbprint
                
                if ($result.Success) {
                    Write-Host -ForegroundColor Green "`n✅ Report generated successfully!"
                    Write-Host -ForegroundColor Cyan "📁 Location: $($result.ReportPath)"
                }
            }
        }
        
        "4" {
            Show-AssessmentMenu -AppId $AppId -CertThumbprint $CertThumbprint
        }
        
        default {
            Write-Host -ForegroundColor Red "Invalid selection. Please try again."
        }
    }
}

# Function to get SharePoint permission IDs for verification
function Get-SharePointPermissionIds {
    param(
        [Parameter(Mandatory = $false)]
        [string]$AccessToken = $null
    )
    
    try {
        $token = if ($AccessToken) { $AccessToken } else { Get-GraphAccessToken }
        
        $headers = @{
            'Authorization' = "Bearer $token"
            'Content-Type'  = 'application/json'
        }
        
        Write-Host -ForegroundColor Yellow "Retrieving SharePoint service principal and permission IDs..."
        
        # Get SharePoint service principal
        $sharePointSP = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0ff1-ce00-000000000000'" -Method GET -Headers $headers
        
        if ($sharePointSP.value.Count -gt 0) {
            $spObject = $sharePointSP.value[0]
            Write-Host -ForegroundColor Cyan "`nSharePoint Service Principal App Roles (Permissions):"
            
            foreach ($role in $spObject.appRoles) {
                if ($role.value -like "*Sites*") {
                    Write-Host -ForegroundColor White "Permission: $($role.value)"
                    Write-Host -ForegroundColor Gray "  ID: $($role.id)"
                    Write-Host -ForegroundColor Gray "  Description: $($role.description)"
                    Write-Host ""
                }
            }
            
            # Find specific permissions we need
            $sitesReadAll = $spObject.appRoles | Where-Object { $_.value -eq "Sites.Read.All" }
            $sitesManageAll = $spObject.appRoles | Where-Object { $_.value -eq "Sites.Manage.All" }
            $sitesFullControlAll = $spObject.appRoles | Where-Object { $_.value -eq "Sites.FullControl.All" }
            
            Write-Host -ForegroundColor Cyan "Correct Permission IDs for Script:"
            if ($sitesReadAll) { Write-Host -ForegroundColor Green "Sites.Read.All: $($sitesReadAll.id)" }
            if ($sitesManageAll) { Write-Host -ForegroundColor Green "Sites.Manage.All: $($sitesManageAll.id)" }
            if ($sitesFullControlAll) { Write-Host -ForegroundColor Green "Sites.FullControl.All: $($sitesFullControlAll.id)" }
            
            return @{
                SitesReadAll        = $sitesReadAll.id
                SitesManageAll      = $sitesManageAll.id
                SitesFullControlAll = $sitesFullControlAll.id
            }
        }
        else {
            Write-Host -ForegroundColor Red "SharePoint service principal not found!"
            return $null
        }
    }
    catch {
        Write-Host -ForegroundColor Red "Error retrieving SharePoint permissions: $_"
        return $null
    }
}

# Function to check if current app has required permissions to create applications
function Test-AppCreationPermissions {
    param(
        [Parameter(Mandatory = $false)]
        [string]$AccessToken = $null
    )
    
    try {
        $token = if ($AccessToken) { $AccessToken } else { Get-GraphAccessToken }
        
        $headers = @{
            'Authorization' = "Bearer $token"
            'Content-Type'  = 'application/json'
        }
        
        Write-Host -ForegroundColor Yellow "Checking current application permissions..."
        
        # Get current app's service principal and permissions
        $currentAppSP = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$appID'" -Method GET -Headers $headers
        
        if ($currentAppSP.value.Count -eq 0) {
            Write-Host -ForegroundColor Red "Current application service principal not found!"
            return $false
        }
        
        $spId = $currentAppSP.value[0].id
        
        # Get app role assignments (application permissions)
        $appRoleAssignments = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$spId/appRoleAssignments" -Method GET -Headers $headers
        
        # Get Microsoft Graph service principal
        $graphSP = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'" -Method GET -Headers $headers
        $graphSpId = $graphSP.value[0].id
        
        # Get Microsoft Graph app roles
        $graphAppRoles = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$graphSpId" -Method GET -Headers $headers
        
        # Check for required permissions
        $requiredPermissions = @{
            "Application.ReadWrite.All"       = "1bfefb4e-e0b5-418b-a88f-73c46d2cc8e9"
            "Directory.ReadWrite.All"         = "19dbc75e-c2e2-444c-a770-ec69d8559fc7"
            "AppRoleAssignment.ReadWrite.All" = "06b708a9-e830-4db3-a914-8e69da51d44f"
        }
        
        $hasAllPermissions = $true
        $currentPermissions = @()
        
        Write-Host -ForegroundColor Cyan "`nCurrent Microsoft Graph Application Permissions:"
        
        foreach ($assignment in $appRoleAssignments.value) {
            if ($assignment.resourceId -eq $graphSpId) {
                $roleInfo = $graphAppRoles.appRoles | Where-Object { $_.id -eq $assignment.appRoleId }
                if ($roleInfo) {
                    $currentPermissions += $roleInfo.value
                    Write-Host -ForegroundColor White "✓ $($roleInfo.value)"
                }
            }
        }
        
        Write-Host -ForegroundColor Cyan "`nRequired Permissions for App Creation:"
        foreach ($permName in $requiredPermissions.Keys) {
            if ($currentPermissions -contains $permName) {
                Write-Host -ForegroundColor Green "✓ $permName (GRANTED)"
            }
            else {
                Write-Host -ForegroundColor Red "✗ $permName (MISSING)"
                $hasAllPermissions = $false
            }
        }
        
        if (-not $hasAllPermissions) {
            Write-Host -ForegroundColor Yellow "`n⚠️  SOLUTION REQUIRED:"
            Write-Host -ForegroundColor White "Your current application needs additional permissions to create Azure applications."
            Write-Host -ForegroundColor White "`nTo fix this, you have two options:"
            Write-Host -ForegroundColor Cyan "`nOption 1 - Add permissions to current app:"
            Write-Host -ForegroundColor White "1. Go to Azure Portal > Azure Active Directory > App registrations"
            Write-Host -ForegroundColor White "2. Find your app (ID: $appID)"
            Write-Host -ForegroundColor White "3. Go to API permissions > Add a permission > Microsoft Graph"
            Write-Host -ForegroundColor White "4. Select Application permissions and add:"
            Write-Host -ForegroundColor White "   - Application.ReadWrite.All"
            Write-Host -ForegroundColor White "   - Directory.ReadWrite.All"
            Write-Host -ForegroundColor White "   - AppRoleAssignment.ReadWrite.All"
            Write-Host -ForegroundColor White "5. Click 'Grant admin consent for [your tenant]'"
            Write-Host -ForegroundColor Cyan "`nOption 2 - Use manual app creation:"
            Write-Host -ForegroundColor White "1. Use the certificate created by New-Cert function"
            Write-Host -ForegroundColor White "2. Manually create the app in Azure Portal"
            Write-Host -ForegroundColor White "3. Upload the certificate and configure permissions"
        }
        
        return $hasAllPermissions
    }
    catch {
        Write-Host -ForegroundColor Red "Error checking permissions: $_"
        return $false
    }
}

# Function to create Azure application with Graph permissions using certificate authentication
function New-AzureAppRegistration {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppDisplayName,
        
        [Parameter(Mandatory = $true)]
        [string]$CertificateFilePath,
        
        [Parameter(Mandatory = $false)]
        [string]$ExistingAccessToken = $null
    )
    
    try {
        # Use existing token or get a new one
        $accessToken = if ($ExistingAccessToken) { $ExistingAccessToken } else { Get-GraphAccessToken }
        
        Write-LogEntry -LogName $logPath -LogEntryText "Starting Azure application registration: $AppDisplayName" -LogLevel "INFO"
        Write-Host -ForegroundColor Yellow "Creating Azure application registration: $AppDisplayName"
        
        # Read the certificate file
        if (-not (Test-Path $CertificateFilePath)) {
            throw "Certificate file not found: $CertificateFilePath"
        }
        
        $certBytes = [System.IO.File]::ReadAllBytes($CertificateFilePath)
        $certBase64 = [System.Convert]::ToBase64String($certBytes)
        
        # Define Microsoft Graph resource IDs
        $microsoftGraphResourceId = "00000003-0000-0000-c000-000000000000"
        $sharePointResourceId = "00000003-0000-0ff1-ce00-000000000000"
        
        # Define the required resource access (permissions)
        $requiredResourceAccess = @(
            @{
                resourceAppId  = $microsoftGraphResourceId
                resourceAccess = @(
                    @{
                        id   = "332a536c-c7ef-4017-ab91-336970924f0d"  # Sites.Read.All
                        type = "Role"  # Application permission
                    },
                    @{
                        id   = "9a5d68dd-52b0-4cc2-bd40-abcf44ac3a30"  # Application.Read.All
                        type = "Role"  # Application permission
                    }
                )
            },
            @{
                resourceAppId  = $sharePointResourceId
                resourceAccess = @(
                    @{
                        id   = "d13f72ca-a275-4b96-b789-48ebcc4da984"  # Sites.Read.All
                        type = "Role"  # Application permission
                    },
                    @{
                        id   = "9bff6588-13f2-4c48-bbf2-ddab62256b36"  # Sites.Manage.All
                        type = "Role"  # Application permission
                    },
                    @{
                        id   = "678536fe-1083-478a-9c59-b99265e6b0d3"  # Sites.FullControl.All
                        type = "Role"  # Application permission
                    }
                )
            }
        )
        
        # Define the application payload
        $appPayload = @{
            displayName            = $AppDisplayName
            signInAudience         = "AzureADMyOrg"
            requiredResourceAccess = $requiredResourceAccess
            keyCredentials         = @(
                @{
                    type  = "AsymmetricX509Cert"
                    usage = "Verify"
                    key   = $certBase64
                }
            )
            publicClient           = @{
                redirectUris = @(
                    "https://login.microsoftonline.com/common/oauth2/nativeclient",
                    "http://localhost",
                    "https://login.live.com/oauth20_desktop.srf",
                    "ms-appx-web://microsoft.aad.brokerplugin/$AppID"
                )
            }
            isFallbackPublicClient = $true
        } | ConvertTo-Json -Depth 10
        
        # Headers for Graph API call
        $headers = @{
            'Authorization' = "Bearer $accessToken"
            'Content-Type'  = 'application/json'
        }
        
        # Create the application
        Write-Host -ForegroundColor Yellow "Creating application registration..."
        $appResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/applications" -Method POST -Headers $headers -Body $appPayload
        
        Write-LogEntry -LogName $logPath -LogEntryText "Application created successfully. App ID: $($appResponse.appId)" -LogLevel "INFO"
        Write-Host -ForegroundColor Green "Application created successfully!"
        Write-Host -ForegroundColor Cyan "Application ID (Client ID): $($appResponse.appId)"
        Write-Host -ForegroundColor Cyan "Object ID: $($appResponse.id)"
        
        # Create a service principal for the application
        Write-Host -ForegroundColor Yellow "Creating service principal..."
        $spPayload = @{
            appId = $appResponse.appId
        } | ConvertTo-Json
        
        $spResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals" -Method POST -Headers $headers -Body $spPayload
        
        Write-LogEntry -LogName $logPath -LogEntryText "Service principal created. SP Object ID: $($spResponse.id)" -LogLevel "INFO"
        Write-Host -ForegroundColor Green "Service principal created successfully!"
        Write-Host -ForegroundColor Cyan "Service Principal Object ID: $($spResponse.id)"
        
        # Configure Mobile and Desktop platform
        Write-Host -ForegroundColor Yellow "Configuring Mobile and Desktop platform..."
        try {
            $platformPayload = @{
                publicClient           = @{
                    redirectUris = @(
                        "https://login.microsoftonline.com/common/oauth2/nativeclient",
                        "msal://redirect",
                        "ms-appx-web://microsoft.aad.brokerplugin/$($appResponse.appId)"
                    )
                }
                isFallbackPublicClient = $true
            } | ConvertTo-Json -Depth 10
            
            # Update the application with platform configuration
            Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/applications/$($appResponse.id)" -Method PATCH -Headers $headers -Body $platformPayload | Out-Null
            
            Write-Host -ForegroundColor Green "✅ Mobile and Desktop platform configured successfully!"
            Write-LogEntry -LogName $logPath -LogEntryText "Mobile and Desktop platform configured for application." -LogLevel "INFO"
        }
        catch {
            Write-Host -ForegroundColor Yellow "⚠️  Note: Mobile and Desktop platform may need to be configured manually."
            Write-LogEntry -LogName $logPath -LogEntryText "Failed to configure Mobile and Desktop platform: $_" -LogLevel "WARNING"
        }
        
        # Grant admin consent for the permissions using appRoleAssignments
        Write-Host -ForegroundColor Yellow "Granting admin consent for application permissions..."
        
        # Get the Microsoft Graph service principal
        $graphSpResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$microsoftGraphResourceId'" -Method GET -Headers $headers
        $graphSpId = $graphSpResponse.value[0].id
        
        # Get the SharePoint service principal
        $sharePointSpResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$sharePointResourceId'" -Method GET -Headers $headers
        $sharePointSpId = $sharePointSpResponse.value[0].id
        
        $successfulGrants = 0
        $totalGrants = 0
        
        # Grant Microsoft Graph permissions using appRoleAssignedTo
        Write-Host -ForegroundColor Cyan "Granting Microsoft Graph permissions..."
        foreach ($permission in $requiredResourceAccess[0].resourceAccess) {
            try {
                $totalGrants++
                $grantPayload = @{
                    principalId = $spResponse.id
                    resourceId  = $graphSpId
                    appRoleId   = $permission.id
                } | ConvertTo-Json
                
                Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$graphSpId/appRoleAssignedTo" -Method POST -Headers $headers -Body $grantPayload | Out-Null
                
                Write-Host -ForegroundColor Green "✅ Granted Microsoft Graph permission: $($permission.id)"
                Write-LogEntry -LogName $logPath -LogEntryText "Granted Microsoft Graph permission: $($permission.id)" -LogLevel "INFO"
                $successfulGrants++
            }
            catch {
                Write-Host -ForegroundColor Yellow "⚠️  Could not grant Microsoft Graph permission $($permission.id): $_"
                Write-LogEntry -LogName $logPath -LogEntryText "Failed to grant Microsoft Graph permission $($permission.id): $_" -LogLevel "WARNING"
            }
        }
        
        # Grant SharePoint permissions using appRoleAssignedTo
        Write-Host -ForegroundColor Cyan "Granting SharePoint permissions..."
        foreach ($permission in $requiredResourceAccess[1].resourceAccess) {
            try {
                $totalGrants++
                $grantPayload = @{
                    principalId = $spResponse.id
                    resourceId  = $sharePointSpId
                    appRoleId   = $permission.id
                } | ConvertTo-Json
                
                Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$sharePointSpId/appRoleAssignedTo" -Method POST -Headers $headers -Body $grantPayload | Out-Null
                
                Write-Host -ForegroundColor Green "✅ Granted SharePoint permission: $($permission.id)"
                Write-LogEntry -LogName $logPath -LogEntryText "Granted SharePoint permission: $($permission.id)" -LogLevel "INFO"
                $successfulGrants++
            }
            catch {
                Write-Host -ForegroundColor Yellow "⚠️  Could not grant SharePoint permission $($permission.id): $_"
                Write-LogEntry -LogName $logPath -LogEntryText "Failed to grant SharePoint permission $($permission.id): $_" -LogLevel "WARNING"
            }
        }
        
        # Summary of permission grants
        Write-Host -ForegroundColor Cyan "`n📊 Permission Grant Summary:"
        Write-Host -ForegroundColor White "✅ Successfully granted: $successfulGrants/$totalGrants permissions"
        
        if ($successfulGrants -eq $totalGrants) {
            Write-Host -ForegroundColor Green "🎉 All permissions granted successfully!"
        }
        elseif ($successfulGrants -gt 0) {
            Write-Host -ForegroundColor Yellow "⚠️  Some permissions granted. Check Azure portal for any missing permissions."
        }
        else {
            Write-Host -ForegroundColor Red "❌ No permissions were granted automatically. Manual admin consent required."
        }
        
        Write-Host -ForegroundColor Green "`nApplication registration completed successfully!"
        Write-Host -ForegroundColor Cyan "==============================================="
        Write-Host -ForegroundColor White "Application Details:"
        Write-Host -ForegroundColor White "- Display Name: $($appResponse.displayName)"
        Write-Host -ForegroundColor White "- Application ID: $($appResponse.appId)"
        Write-Host -ForegroundColor White "- Object ID: $($appResponse.id)"
        Write-Host -ForegroundColor White "- Service Principal Object ID: $($spResponse.id)"
        Write-Host -ForegroundColor White "- Certificate configured: Yes"
        Write-Host -ForegroundColor White "- Permissions granted: $successfulGrants/$totalGrants"
        Write-Host -ForegroundColor Cyan "==============================================="
        
        if ($successfulGrants -lt $totalGrants) {
            Write-Host -ForegroundColor Yellow "`n⚠️  Some permissions may require manual admin consent:"
            Write-Host -ForegroundColor White "1. Go to the Azure portal (portal.azure.com)"
            Write-Host -ForegroundColor White "2. Navigate to Azure Active Directory > App registrations"
            Write-Host -ForegroundColor White "3. Find your application: $($appResponse.displayName)"
            Write-Host -ForegroundColor White "4. Go to API permissions and grant admin consent if needed"
        }
        else {
            Write-Host -ForegroundColor Green "`n🎉 All permissions granted successfully! Ready to use."
        }
        
        # Return application information
        return @{
            ApplicationId      = $appResponse.appId
            ObjectId           = $appResponse.id
            ServicePrincipalId = $spResponse.id
            DisplayName        = $appResponse.displayName
            PermissionsGranted = $successfulGrants
            TotalPermissions   = $totalGrants
            Success            = $true
        }
    }
    catch {
        $errorMessage = "Failed to create Azure application registration: $_"
        Write-LogEntry -LogName $logPath -LogEntryText $errorMessage -LogLevel "ERROR"
        Write-Host -ForegroundColor Red $errorMessage
        
        return @{
            Success = $false
            Error   = $errorMessage
        }
    }
}

# Function to verify granted application permissions
function Test-GrantedPermissions {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ServicePrincipalId,
        
        [Parameter(Mandatory = $false)]
        [string]$AccessToken = $null
    )
    
    try {
        $token = if ($AccessToken) { $AccessToken } else { Get-GraphAccessToken }
        
        $headers = @{
            'Authorization' = "Bearer $token"
            'Content-Type'  = 'application/json'
        }
        
        Write-Host -ForegroundColor Yellow "🔍 Verifying granted permissions..."
        
        # Get all app role assignments for the service principal
        $appRoleAssignments = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$ServicePrincipalId/appRoleAssignments" -Method GET -Headers $headers
        
        Write-Host -ForegroundColor Cyan "`n📋 Granted Application Permissions:"
        
        $microsoftGraphResourceId = "00000003-0000-0000-c000-000000000000"
        $sharePointResourceId = "00000003-0000-0ff1-ce00-000000000000"
        
        $graphPermissions = @()
        $sharePointPermissions = @()
        
        foreach ($assignment in $appRoleAssignments.value) {
            # Get the resource service principal to determine the service
            $resourceSP = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($assignment.resourceId)" -Method GET -Headers $headers
            
            if ($resourceSP.appId -eq $microsoftGraphResourceId) {
                # Find the app role name
                $appRole = $resourceSP.appRoles | Where-Object { $_.id -eq $assignment.appRoleId }
                if ($appRole) {
                    Write-Host -ForegroundColor Green "✅ Microsoft Graph: $($appRole.value)"
                    $graphPermissions += $appRole.value
                }
            }
            elseif ($resourceSP.appId -eq $sharePointResourceId) {
                # Find the app role name  
                $appRole = $resourceSP.appRoles | Where-Object { $_.id -eq $assignment.appRoleId }
                if ($appRole) {
                    Write-Host -ForegroundColor Green "✅ SharePoint: $($appRole.value)"
                    $sharePointPermissions += $appRole.value
                }
            }
        }
        
        Write-Host -ForegroundColor Cyan "`n📊 Permission Summary:"
        Write-Host -ForegroundColor White "Microsoft Graph permissions: $($graphPermissions.Count)"
        Write-Host -ForegroundColor White "SharePoint permissions: $($sharePointPermissions.Count)"
        Write-Host -ForegroundColor White "Total permissions: $($appRoleAssignments.value.Count)"
        
        return @{
            Success               = $true
            GraphPermissions      = $graphPermissions
            SharePointPermissions = $sharePointPermissions
            TotalPermissions      = $appRoleAssignments.value.Count
        }
    }
    catch {
        Write-Host -ForegroundColor Red "Error verifying permissions: $_"
        return @{ Success = $false; Error = $_ }
    }
}

# Function to create certificate and Azure application in one workflow
function New-CertificateAndAzureApp {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppDisplayName
    )
    
    try {
        Write-Host -ForegroundColor Cyan "Starting certificate creation and Azure application registration workflow..."
        
        # Step 1: Create certificate
        Write-Host -ForegroundColor Yellow "`nStep 1: Creating certificate..."
        $certInfo = New-Cert
        
        if ($certInfo -and $certInfo.CerFilePath) {
            Write-Host -ForegroundColor Green "Certificate created successfully!"
            
            # Step 2: Create Azure application with the certificate
            Write-Host -ForegroundColor Yellow "`nStep 2: Creating Azure application registration..."
            $appResult = New-AzureAppRegistration -AppDisplayName $AppDisplayName -CertificateFilePath $certInfo.CerFilePath
            
            if ($appResult.Success) {
                Write-Host -ForegroundColor Green "`nWorkflow completed successfully!"
                Write-Host -ForegroundColor Cyan "==============================================="
                Write-Host -ForegroundColor White "Summary:"
                Write-Host -ForegroundColor White "- Certificate Thumbprint: $($certInfo.Thumbprint)"
                Write-Host -ForegroundColor White "- Certificate Files:"
                Write-Host -ForegroundColor White "  * CER: $($certInfo.CerFilePath)"
                Write-Host -ForegroundColor White "  * PFX: $($certInfo.PfxFilePath)"
                Write-Host -ForegroundColor White "- Application ID: $($appResult.ApplicationId)"
                Write-Host -ForegroundColor White "- Application Name: $($appResult.DisplayName)"
                Write-Host -ForegroundColor Cyan "==============================================="
                
                # Set global variables for M365 Assessment operations
                Write-Host -ForegroundColor Yellow "`nStep 3: Configuring M365 Assessment variables..."
                $global:m365AssessmentAppID = $appResult.ApplicationId
                $global:m365AssessmentThumbprint = $certInfo.Thumbprint
                
                Write-Host -ForegroundColor Green "✅ M365 Assessment app configured successfully!"
                Write-Host -ForegroundColor Cyan "💡 You can now use the assessment functions with the new app:"
                Write-Host -ForegroundColor White "   - App ID: $global:m365AssessmentAppID"
                Write-Host -ForegroundColor White "   - Thumbprint: $global:m365AssessmentThumbprint"
                Write-Host -ForegroundColor White "   - Certificate Store: CurrentUser"
                
                return @{
                    Certificate = $certInfo
                    Application = $appResult
                    Success     = $true
                }
            }
            else {
                throw "Application registration failed: $($appResult.Error)"
            }
        }
        else {
            throw "Certificate creation failed"
        }
    }
    catch {
        $errorMessage = "Workflow failed: $_"
        Write-LogEntry -LogName $logPath -LogEntryText $errorMessage -LogLevel "ERROR"
        Write-Host -ForegroundColor Red $errorMessage
        
        return @{
            Success = $false
            Error   = $errorMessage
        }
    }
}

# Function to automatically discover existing M365 Assessment app configuration
# Function to add certificate to existing Azure app registration
function Add-CertificateToApp {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppId,
        
        [Parameter(Mandatory = $true)]
        [string]$CertificateFilePath,
        
        [Parameter(Mandatory = $false)]
        [string]$AccessToken = $null
    )
    
    try {
        # Use existing token or get a new one
        $token = if ($AccessToken) { $AccessToken } else { Get-GraphAccessToken }
        
        Write-Host -ForegroundColor Yellow "Adding certificate to app registration: $AppId"
        
        # Read the certificate file
        if (-not (Test-Path $CertificateFilePath)) {
            throw "Certificate file not found: $CertificateFilePath"
        }
        
        $certBytes = [System.IO.File]::ReadAllBytes($CertificateFilePath)
        $certBase64 = [System.Convert]::ToBase64String($certBytes)
        
        # Create certificate object to get details
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($CertificateFilePath)
        
        # Get the app's current configuration
        $headers = @{
            'Authorization' = "Bearer $token"
            'Content-Type'  = 'application/json'
        }
        
        # First, get the app by appId to find the object ID
        $appFilter = "appId eq '$AppId'"
        $appUri = "https://graph.microsoft.com/v1.0/applications?`$filter=$appFilter"
        $appResponse = Invoke-RestMethod -Uri $appUri -Method GET -Headers $headers
        
        if ($appResponse.value.Count -eq 0) {
            throw "Application not found with App ID: $AppId"
        }
        
        $appObjectId = $appResponse.value[0].id
        $currentKeyCredentials = $appResponse.value[0].keyCredentials
        
        # Create new key credential
        $newKeyCredential = @{
            type          = "AsymmetricX509Cert"
            usage         = "Verify"
            key           = $certBase64
            displayName   = "Auto-uploaded certificate - $($cert.Thumbprint)"
            startDateTime = $cert.NotBefore.ToString("yyyy-MM-ddTHH:mm:ssZ")
            endDateTime   = $cert.NotAfter.ToString("yyyy-MM-ddTHH:mm:ssZ")
        }
        
        # Add to existing key credentials (don't replace, append)
        $updatedKeyCredentials = @()
        if ($currentKeyCredentials) {
            $updatedKeyCredentials += $currentKeyCredentials
        }
        $updatedKeyCredentials += $newKeyCredential
        
        # Update the application
        $updatePayload = @{
            keyCredentials = $updatedKeyCredentials
        } | ConvertTo-Json -Depth 10
        
        $updateUri = "https://graph.microsoft.com/v1.0/applications/$appObjectId"
        $updateResponse = Invoke-RestMethod -Uri $updateUri -Method PATCH -Headers $headers -Body $updatePayload
        
        Write-Host -ForegroundColor Green "✅ Certificate added successfully to app registration!"
        Write-Host -ForegroundColor Cyan "Certificate Thumbprint: $($cert.Thumbprint)"
        Write-Host -ForegroundColor Cyan "Certificate Subject: $($cert.Subject)"
        
        return @{
            Success    = $true
            Thumbprint = $cert.Thumbprint
            Subject    = $cert.Subject
        }
    }
    catch {
        Write-Host -ForegroundColor Red "❌ Failed to add certificate to app: $_"
        return @{
            Success = $false
            Error   = $_.Exception.Message
        }
    }
}

# Function to auto-discover existing M365 Assessment app configuration
function Get-M365AssessmentAppConfig {
    param(
        [Parameter(Mandatory = $false)]
        [string]$AppDisplayName = $global:appname,
        
        [Parameter(Mandatory = $false)]
        [string]$AccessToken = $null
    )
    
    try {
        Write-Host -ForegroundColor Yellow "🔍 Auto-discovering M365 Assessment app configuration..."
        
        # Get access token using the existing app
        $token = if ($AccessToken) { $AccessToken } else { Get-GraphAccessToken }
        
        $headers = @{
            'Authorization' = "Bearer $token"
            'Content-Type'  = 'application/json'
        }
        
        # Search for applications by display name
        $filter = "displayName eq '$AppDisplayName'"
        $uri = "https://graph.microsoft.com/v1.0/applications?`$filter=$filter"
        
        Write-Host -ForegroundColor Cyan "Searching for app: '$AppDisplayName'..."
        $appResponse = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers
        
        if ($appResponse.value.Count -gt 0) {
            $app = $appResponse.value[0]  # Take the first match
            Write-Host -ForegroundColor Green "✅ Found M365 Assessment app!"
            Write-Host -ForegroundColor Cyan "   App ID: $($app.appId)"
            Write-Host -ForegroundColor Cyan "   Display Name: $($app.displayName)"
            
            # Get the certificate information from the app
            if ($app.keyCredentials -and $app.keyCredentials.Count -gt 0) {
                Write-Host -ForegroundColor Cyan "🔍 Found $($app.keyCredentials.Count) certificate(s) in app registration..."
                
                # Debug: Show the structure of keyCredentials
                Write-Host -ForegroundColor Gray "Debug: keyCredentials structure:"
                for ($i = 0; $i -lt $app.keyCredentials.Count; $i++) {
                    $keyCredential = $app.keyCredentials[$i]
                    Write-Host -ForegroundColor Gray "  Certificate $($i + 1):"
                    Write-Host -ForegroundColor Gray "    Type: $($keyCredential.type)"
                    Write-Host -ForegroundColor Gray "    Usage: $($keyCredential.usage)"
                    Write-Host -ForegroundColor Gray "    DisplayName: $($keyCredential.displayName)"
                    Write-Host -ForegroundColor Gray "    StartDateTime: $($keyCredential.startDateTime)"
                    Write-Host -ForegroundColor Gray "    EndDateTime: $($keyCredential.endDateTime)"
                    Write-Host -ForegroundColor Gray "    Key present: $($keyCredential.key -ne $null)"
                    if ($keyCredential.key) {
                        Write-Host -ForegroundColor Gray "    Key length: $($keyCredential.key.Length) characters"
                        Write-Host -ForegroundColor Gray "    Key starts with: $($keyCredential.key.Substring(0, [Math]::Min(50, $keyCredential.key.Length)))..."
                    }
                    Write-Host -ForegroundColor Gray "    CustomKeyIdentifier present: $($keyCredential.customKeyIdentifier -ne $null)"
                    if ($keyCredential.customKeyIdentifier) {
                        Write-Host -ForegroundColor Gray "    CustomKeyIdentifier: $($keyCredential.customKeyIdentifier)"
                    }
                }
                
                # Check all certificates in the app to find the right one
                $foundMatchingCert = $false
                $certIndex = 0
                $discoveredThumbprint = $null  # Initialize in outer scope
                
                foreach ($keyCredential in $app.keyCredentials) {
                    Write-Host -ForegroundColor Gray "   Checking certificate $($certIndex + 1)..."
                    
                    try {
                        # Extract the certificate details
                        Write-Host -ForegroundColor Gray "   Extracting certificate data..."
                        
                        # Try multiple approaches to get the thumbprint
                        $currentThumbprint = $null
                        
                        # Method 1: Try customKeyIdentifier (sometimes contains the thumbprint)
                        if ($keyCredential.customKeyIdentifier) {
                            Write-Host -ForegroundColor Gray "   Trying customKeyIdentifier approach..."
                            try {
                                # customKeyIdentifier is often the SHA1 hash (thumbprint) directly as hex string
                                $currentThumbprint = $keyCredential.customKeyIdentifier
                                Write-Host -ForegroundColor Green "   ✅ Got thumbprint from customKeyIdentifier: $currentThumbprint"
                            }
                            catch {
                                Write-Host -ForegroundColor Yellow "   customKeyIdentifier method failed: $_"
                            }
                        }
                        
                        # Method 2: Try extracting from certificate data
                        if (-not $currentThumbprint) {
                            $certBase64 = $keyCredential.key
                            
                            if (-not $certBase64) {
                                Write-Host -ForegroundColor Red "   No certificate data found in keyCredential"
                                continue
                            }
                            
                            Write-Host -ForegroundColor Gray "   Certificate data length: $($certBase64.Length) characters"
                            Write-Host -ForegroundColor Gray "   Converting from Base64..."
                            
                            $certBytes = [System.Convert]::FromBase64String($certBase64)
                            Write-Host -ForegroundColor Gray "   Certificate bytes length: $($certBytes.Length)"
                            
                            # Create a temporary certificate object to get the thumbprint
                            Write-Host -ForegroundColor Gray "   Creating X509Certificate2 object..."
                            $tempCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certBytes)
                            $currentThumbprint = $tempCert.Thumbprint
                            
                            Write-Host -ForegroundColor Green "   ✅ Got thumbprint from certificate data: $currentThumbprint"
                            Write-Host -ForegroundColor Gray "   Certificate Subject: $($tempCert.Subject)"
                            Write-Host -ForegroundColor Gray "   Certificate Issuer: $($tempCert.Issuer)"
                            Write-Host -ForegroundColor Gray "   Certificate Valid From: $($tempCert.NotBefore)"
                            Write-Host -ForegroundColor Gray "   Certificate Valid To: $($tempCert.NotAfter)"
                        }
                        
                        # Now check if this certificate exists in CurrentUser store
                        if ($currentThumbprint) {
                            Write-Host -ForegroundColor Gray "   Checking CurrentUser certificate store for: $currentThumbprint"
                            $localCert = Get-Item "Cert:\CurrentUser\My\$currentThumbprint" -ErrorAction SilentlyContinue
                            if ($localCert) {
                                Write-Host -ForegroundColor Green "✅ Found matching certificate in CurrentUser store!"
                                Write-Host -ForegroundColor Cyan "   Certificate Thumbprint: $currentThumbprint"
                                $discoveredThumbprint = $currentThumbprint  # Set in outer scope
                                $foundMatchingCert = $true
                                break
                            }
                            else {
                                Write-Host -ForegroundColor Yellow "   Certificate not found in CurrentUser store"
                                Write-Host -ForegroundColor Gray "   Looking for certificates with similar subject..."
                                $similarCerts = Get-ChildItem "Cert:\CurrentUser\My" | Where-Object { 
                                    $_.Thumbprint -eq $currentThumbprint
                                }
                                if ($similarCerts) {
                                    Write-Host -ForegroundColor Gray "   Found exact thumbprint match:"
                                    foreach ($cert in $similarCerts) {
                                        Write-Host -ForegroundColor Gray "     • $($cert.Thumbprint) - $($cert.Subject)"
                                    }
                                }
                                else {
                                    # Check for any certificates with the expected thumbprint
                                    $expectedCert = Get-Item "Cert:\CurrentUser\My\$expectedThumbprint" -ErrorAction SilentlyContinue
                                    if ($expectedCert) {
                                        Write-Host -ForegroundColor Yellow "   Found expected certificate in store but thumbprints don't match:"
                                        Write-Host -ForegroundColor Yellow "     Azure app thumbprint: $currentThumbprint"
                                        Write-Host -ForegroundColor Yellow "     Local cert thumbprint: $expectedThumbprint"
                                    }
                                }
                            }
                        }
                        else {
                            Write-Host -ForegroundColor Red "   Failed to extract thumbprint from certificate"
                        }
                    }
                    catch {
                        Write-Host -ForegroundColor Red "   ❌ Error processing certificate $($certIndex + 1): $_"
                        Write-Host -ForegroundColor Red "   Error details: $($_.Exception.Message)"
                        if ($_.Exception.InnerException) {
                            Write-Host -ForegroundColor Red "   Inner exception: $($_.Exception.InnerException.Message)"
                        }
                    }
                    
                    $certIndex++
                }
                
                if ($foundMatchingCert) {
                    Write-Host -ForegroundColor Green "✅ Certificate verified in CurrentUser certificate store!"
                    
                    # Set global variables
                    $global:m365AssessmentAppID = $app.appId
                    $global:m365AssessmentThumbprint = $discoveredThumbprint
                    
                    Write-Host -ForegroundColor Green "🎉 M365 Assessment app auto-configured successfully!"
                    Write-Host -ForegroundColor Cyan "Configuration details:"
                    Write-Host -ForegroundColor White "   - App ID: $global:m365AssessmentAppID"
                    Write-Host -ForegroundColor White "   - Thumbprint: $global:m365AssessmentThumbprint"
                    Write-Host -ForegroundColor White "   - Certificate Store: CurrentUser"
                    
                    return @{
                        Success        = $true
                        AppId          = $global:m365AssessmentAppID
                        Thumbprint     = $global:m365AssessmentThumbprint
                        DisplayName    = $app.displayName
                        AutoDiscovered = $true
                    }
                }
                else {
                    Write-Host -ForegroundColor Yellow "⚠️  No matching certificates found in CurrentUser store"
                    Write-Host -ForegroundColor Cyan "💡 Expected certificate thumbprint: 7A145A41D29D7A90F208DE33E61E82AAD7DF06AC"
                    Write-Host -ForegroundColor Cyan "💡 Certificates found in app registration:"
                    
                    # List all certificates found for troubleshooting
                    $certIndex = 0
                    foreach ($keyCredential in $app.keyCredentials) {
                        try {
                            $certBase64 = $keyCredential.key
                            $certBytes = [System.Convert]::FromBase64String($certBase64)
                            $tempCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certBytes)
                            Write-Host -ForegroundColor White "   $($certIndex + 1). $($tempCert.Thumbprint)"
                        }
                        catch {
                            Write-Host -ForegroundColor White "   $($certIndex + 1). [Error reading certificate]"
                        }
                        $certIndex++
                    }
                    
                    # Check what certificates are in CurrentUser store
                    Write-Host -ForegroundColor Cyan "💡 Certificates in CurrentUser\My store:"
                    try {
                        $localCerts = Get-ChildItem "Cert:\CurrentUser\My" | Where-Object { $_.Subject -like "*$tenantname*" -or $_.Subject -like "*$certname*" }
                        if ($localCerts) {
                            foreach ($cert in $localCerts) {
                                Write-Host -ForegroundColor White "   • $($cert.Thumbprint) - $($cert.Subject)"
                            }
                        }
                        else {
                            Write-Host -ForegroundColor Yellow "   No certificates found matching tenant name or cert name"
                        }
                    }
                    catch {
                        Write-Host -ForegroundColor Red "   Error listing CurrentUser certificates: $_"
                    }
                    
                    return @{
                        Success               = $false
                        Error                 = "No certificate found in app registration"
                        AppId                 = $app.appId
                        RequiresCertUpload    = $true
                        FoundLocalCertificate = $true
                        LocalCertThumbprint   = "7A145A41D29D7A90F208DE33E61E82AAD7DF06AC"
                    }
                }
            }
            else {
                Write-Host -ForegroundColor Yellow "⚠️  App found but no certificate configured"
                Write-Host -ForegroundColor Cyan "   App ID: $($app.appId)"
                Write-Host -ForegroundColor Cyan "   Checking if local certificate can be uploaded..."
                
                # Check if we have the expected certificate locally
                $expectedThumbprint = "7A145A41D29D7A90F208DE33E61E82AAD7DF06AC"
                $localCert = Get-Item "Cert:\CurrentUser\My\$expectedThumbprint" -ErrorAction SilentlyContinue
                
                if ($localCert) {
                    Write-Host -ForegroundColor Green "✅ Found expected certificate in local store!"
                    Write-Host -ForegroundColor Cyan "   Thumbprint: $expectedThumbprint"
                    Write-Host -ForegroundColor Yellow "💡 Would you like to upload this certificate to the app registration? (Y/N)"
                    
                    $uploadChoice = Read-Host "Upload certificate"
                    if ($uploadChoice -eq "Y" -or $uploadChoice -eq "y") {
                        Write-Host -ForegroundColor Yellow "🔄 Uploading certificate to app registration..."
                        
                        try {
                            # Export the certificate to upload
                            $certPath = "$env:TEMP\temp_cert_upload.cer"
                            Export-Certificate -Cert $localCert -FilePath $certPath -Force | Out-Null
                            
                            # Upload certificate to the app
                            $uploadResult = Add-CertificateToApp -AppId $app.appId -CertificateFilePath $certPath -AccessToken $token
                            
                            if ($uploadResult.Success) {
                                Write-Host -ForegroundColor Green "✅ Certificate uploaded successfully!"
                                
                                # Clean up temporary file
                                Remove-Item $certPath -Force -ErrorAction SilentlyContinue
                                
                                # Set global variables
                                $global:m365AssessmentAppID = $app.appId
                                $global:m365AssessmentThumbprint = $expectedThumbprint
                                
                                Write-Host -ForegroundColor Green "🎉 M365 Assessment app auto-configured successfully!"
                                Write-Host -ForegroundColor Cyan "Configuration details:"
                                Write-Host -ForegroundColor White "   - App ID: $global:m365AssessmentAppID"
                                Write-Host -ForegroundColor White "   - Thumbprint: $global:m365AssessmentThumbprint"
                                Write-Host -ForegroundColor White "   - Certificate Store: CurrentUser"
                                
                                return @{
                                    Success        = $true
                                    AppId          = $global:m365AssessmentAppID
                                    Thumbprint     = $global:m365AssessmentThumbprint
                                    DisplayName    = $app.displayName
                                    AutoDiscovered = $true
                                    CertUploaded   = $true
                                }
                            }
                            else {
                                Write-Host -ForegroundColor Red "❌ Failed to upload certificate: $($uploadResult.Error)"
                                return @{
                                    Success            = $false
                                    Error              = "Certificate upload failed: $($uploadResult.Error)"
                                    AppId              = $app.appId
                                    RequiresCertUpload = $true
                                }
                            }
                        }
                        catch {
                            Write-Host -ForegroundColor Red "❌ Error uploading certificate: $_"
                            return @{
                                Success            = $false
                                Error              = "Certificate upload error: $_"
                                AppId              = $app.appId
                                RequiresCertUpload = $true
                            }
                        }
                    }
                    else {
                        Write-Host -ForegroundColor Yellow "Certificate upload cancelled by user"
                        return @{
                            Success            = $false
                            Error              = "Certificate upload cancelled"
                            AppId              = $app.appId
                            RequiresCertUpload = $true
                        }
                    }
                }
                else {
                    return @{
                        Success            = $false
                        Error              = "No certificate found in app registration"
                        AppId              = $app.appId
                        RequiresCertUpload = $true
                    }
                }
            }
        }
        else {
            Write-Host -ForegroundColor Yellow "⚠️  No M365 Assessment app found with name: '$AppDisplayName'"
            Write-Host -ForegroundColor Cyan "💡 You can create one using: New-CertificateAndAzureApp"
            
            return @{
                Success             = $false
                Error               = "M365 Assessment app not found"
                RequiresAppCreation = $true
            }
        }
    }
    catch {
        Write-Host -ForegroundColor Red "❌ Error during app discovery: $_"
        Write-LogEntry -LogName $logPath -LogEntryText "M365 Assessment app discovery failed: $_" -LogLevel "ERROR"
        
        return @{
            Success = $false
            Error   = "Discovery failed: $_"
        }
    }
}

# Function to configure M365 Assessment app variables manually
function Set-M365AssessmentApp {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppId,
        
        [Parameter(Mandatory = $true)]
        [string]$CertThumbprint
    )
    
    try {
        Write-Host -ForegroundColor Yellow "Setting M365 Assessment app configuration..."
        
        # Validate that the certificate exists in CurrentUser store
        $cert = Get-Item "Cert:\CurrentUser\My\$CertThumbprint" -ErrorAction SilentlyContinue
        if (-not $cert) {
            throw "Certificate with thumbprint '$CertThumbprint' not found in CurrentUser\My store"
        }
        
        # Set global variables
        $global:m365AssessmentAppID = $AppId
        $global:m365AssessmentThumbprint = $CertThumbprint
        
        Write-Host -ForegroundColor Green "✅ M365 Assessment app configured successfully!"
        Write-Host -ForegroundColor Cyan "Configuration details:"
        Write-Host -ForegroundColor White "   - App ID: $global:m365AssessmentAppID"
        Write-Host -ForegroundColor White "   - Thumbprint: $global:m365AssessmentThumbprint"
        Write-Host -ForegroundColor White "   - Certificate Store: CurrentUser"
        Write-Host -ForegroundColor Yellow "💡 You can now use Invoke-M365Assessment and Show-AssessmentMenu functions"
        
        return @{
            Success    = $true
            AppId      = $global:m365AssessmentAppID
            Thumbprint = $global:m365AssessmentThumbprint
        }
    }
    catch {
        $errorMessage = "Failed to configure M365 Assessment app: $_"
        Write-Host -ForegroundColor Red $errorMessage
        return @{
            Success = $false
            Error   = $errorMessage
        }
    }
}

# Function to show main menu
function Show-MainMenu {
    Write-Host -ForegroundColor Cyan "`n🚀 Microsoft 365 Assessment Scanner"
    Write-Host -ForegroundColor Cyan "======================================"
    Write-Host -ForegroundColor White "Current Configuration:"
    Write-Host -ForegroundColor Gray "  • Tenant: $tenantname"
    Write-Host -ForegroundColor Gray "  • SharePoint URL: $sharepointTenantUrl"
    Write-Host -ForegroundColor Gray "  • Creation App ID: $appID"
    Write-Host -ForegroundColor Gray "  • Creation App Thumbprint: $thumbprint"
    Write-Host -ForegroundColor Gray "  • Assessment Tool: $assessmentToolPath"
    Write-Host ""
    Write-Host -ForegroundColor White "M365 Assessment App Status:"
    if ($global:m365AssessmentAppID -and $global:m365AssessmentThumbprint) {
        Write-Host -ForegroundColor Green "  ✅ M365 Assessment App Configured"
        Write-Host -ForegroundColor Gray "     • App ID: $global:m365AssessmentAppID"
        Write-Host -ForegroundColor Gray "     • Thumbprint: $global:m365AssessmentThumbprint"
    }
    else {
        Write-Host -ForegroundColor Yellow "  ⚠️  M365 Assessment App Not Configured"
        Write-Host -ForegroundColor Gray "     • Use option 1 to create or option 3 to auto-discover"
    }
    Write-Host ""
    Write-Host -ForegroundColor White "Choose an option:"
    Write-Host -ForegroundColor White "1. 🔧 Create New M365 Assessment App (with certificate)"
    Write-Host -ForegroundColor White "2. 📊 Run Assessment with M365 Assessment App"
    Write-Host -ForegroundColor White "3. 🔍 Auto-discover Existing M365 Assessment App"
    Write-Host -ForegroundColor White "4. ❌ Exit"
    Write-Host ""
    
    $choice = Read-Host "Enter your choice (1-4)"
    
    switch ($choice) {
        "1" {
            Write-Host -ForegroundColor Yellow "`n🔧 Starting M365 Assessment App Creation Process..."
            Start-AppCreationFlow
        }
        "2" {
            Write-Host -ForegroundColor Yellow "`n📊 Starting Assessment with M365 Assessment App..."
            Start-AssessmentFlow
        }
        "3" {
            Write-Host -ForegroundColor Yellow "`n🔍 Auto-discovering M365 Assessment App..."
            $discoveryResult = Get-M365AssessmentAppConfig
            if ($discoveryResult.Success) {
                Write-Host -ForegroundColor Green "`n✅ M365 Assessment app configured successfully!"
                Write-Host -ForegroundColor Cyan "You can now use option 2 to run assessments."
            }
            else {
                Write-Host -ForegroundColor Yellow "`n⚠️  Auto-discovery completed but app not configured."
                Write-Host -ForegroundColor Cyan "Consider using option 1 to create a new app."
            }
            Write-Host -ForegroundColor Yellow "`nReturning to main menu..."
            Show-MainMenu
        }
        "4" {
            Write-Host -ForegroundColor Yellow "Exiting..."
            return
        }
        default {
            Write-Host -ForegroundColor Red "Invalid selection. Please try again."
            Show-MainMenu
        }
    }
}

# Function to handle app creation flow
function Start-AppCreationFlow {
    try {
        # Authenticate with Microsoft Graph
        Write-Host -ForegroundColor Yellow "Authenticating with Microsoft Graph..."
        $accessToken = Get-GraphAccessToken
        
        if ($accessToken) {
            Write-Host -ForegroundColor Green "Authentication successful!"
            
            # Check if current app has permissions to create applications
            $hasRequiredPermissions = Test-AppCreationPermissions -AccessToken $accessToken
            
            if ($hasRequiredPermissions) {
                Write-Host -ForegroundColor Green "`n✅ Permissions verified! Proceeding with app creation..."
                
                # Get correct SharePoint permission IDs
                Write-Host -ForegroundColor Yellow "`nVerifying SharePoint permission IDs..."
                $spPermissions = Get-SharePointPermissionIds -AccessToken $accessToken
                
                if ($spPermissions -and $spPermissions.SitesManageAll) {
                    Write-Host -ForegroundColor Green "✅ SharePoint permission IDs verified!"
                    
                    # Create certificate
                    Write-Host -ForegroundColor Yellow "`nCreating certificate..."
                    $certInfo = New-Cert
                    
                    # Create Azure application
                    $existingCertPath = "$certexportpath\$certname.cer"
                    if (Test-Path $existingCertPath) {
                        $appResult = New-AzureAppRegistration -AppDisplayName $appname -CertificateFilePath $existingCertPath -ExistingAccessToken $accessToken
                        if ($appResult.Success) {
                            Write-Host -ForegroundColor Green "✅ Application created successfully!"
                            Write-Host -ForegroundColor Cyan "Application ID: $($appResult.ApplicationId)"
                            Write-Host -ForegroundColor Cyan "Certificate Thumbprint: $($certInfo.Thumbprint)"
                            Write-Host -ForegroundColor Cyan "Permissions Granted: $($appResult.PermissionsGranted)/$($appResult.TotalPermissions)"
                            
                            # Update global variables with new app details
                            $global:newAppId = $appResult.ApplicationId
                            $global:newThumbprint = $certInfo.Thumbprint
                            
                            Write-Host -ForegroundColor Yellow "`n⏳ Waiting 30 seconds for app permissions to propagate..."
                            Start-Sleep -Seconds 30
                            
                            # Verify permissions were granted properly
                            $permissionCheck = Test-GrantedPermissions -ServicePrincipalId $appResult.ServicePrincipalId -AccessToken $accessToken
                            
                            if ($permissionCheck.Success -and $permissionCheck.TotalPermissions -gt 0) {
                                Write-Host -ForegroundColor Green "`n🎉 Application is ready with proper permissions!"
                            }
                            else {
                                Write-Host -ForegroundColor Yellow "`n⚠️  Application created but some permissions may need manual consent."
                            }
                            
                            # Show next steps
                            Write-Host -ForegroundColor Green "`n🎯 Application Creation Complete!"
                            Write-Host -ForegroundColor Cyan "📋 New Application Details:"
                            Write-Host -ForegroundColor White "  • Application ID: $($appResult.ApplicationId)"
                            Write-Host -ForegroundColor White "  • Certificate Thumbprint: $($certInfo.Thumbprint)"
                            Write-Host -ForegroundColor White "  • Certificate Files:"
                            Write-Host -ForegroundColor White "    - CER: $($certInfo.CerFilePath)"
                            Write-Host -ForegroundColor White "    - PFX: $($certInfo.PfxFilePath)"
                            Write-Host ""
                            Write-Host -ForegroundColor Yellow "💡 Next Steps:"
                            Write-Host -ForegroundColor White "Run assessments using the newly created application"
                            Write-Host ""
                            Write-Host -ForegroundColor White "What would you like to do now?"
                            Write-Host -ForegroundColor White "1. Run Assessment with New Application"
                            Write-Host -ForegroundColor White "2. Return to Main Menu"
                            Write-Host -ForegroundColor White "3. Exit"
                            
                            $nextChoice = Read-Host "Enter your choice (1-3)"
                            
                            switch ($nextChoice) {
                                "1" {
                                    # Pass the new app credentials to the assessment menu
                                    Show-AssessmentMenu -AppId $global:newAppId -CertThumbprint $global:newThumbprint
                                }
                                "2" {
                                    Show-MainMenu
                                }
                                "3" {
                                    Write-Host -ForegroundColor Yellow "Exiting..."
                                    return
                                }
                                default {
                                    Write-Host -ForegroundColor Yellow "Invalid choice. Returning to main menu..."
                                    Show-MainMenu
                                }
                            }
                        }
                        else {
                            Write-Host -ForegroundColor Red "Failed to create application: $($appResult.Error)"
                            Write-Host -ForegroundColor Yellow "Returning to main menu..."
                            Show-MainMenu
                        }
                    }
                }
                else {
                    Write-Host -ForegroundColor Red "❌ Could not retrieve correct SharePoint permission IDs!"
                    Write-Host -ForegroundColor Yellow "Please check the SharePoint service principal configuration."
                    Write-Host -ForegroundColor Yellow "Returning to main menu..."
                    Show-MainMenu
                }
            }
            else {
                Write-Host -ForegroundColor Yellow "`n⚠️  Cannot create Azure application due to insufficient permissions."
                Write-Host -ForegroundColor Cyan "`nCreating certificate only (you can use this for manual app creation):"
                
                # Still create the certificate for manual use
                $certInfo = New-Cert
                
                Write-Host -ForegroundColor Cyan "`n📋 Manual App Creation Instructions:"
                Write-Host -ForegroundColor White "1. Go to Azure Portal: https://portal.azure.com"
                Write-Host -ForegroundColor White "2. Navigate to: Azure Active Directory > App registrations > New registration"
                Write-Host -ForegroundColor White "3. Enter name: '$appname'"
                Write-Host -ForegroundColor White "4. Select 'Accounts in this organizational directory only'"
                Write-Host -ForegroundColor White "5. Click 'Register'"
                Write-Host -ForegroundColor White "6. Go to 'Certificates & secrets' > 'Certificates' > 'Upload certificate'"
                Write-Host -ForegroundColor White "7. Upload: $($certInfo.CerFilePath)"
                Write-Host -ForegroundColor White "8. Go to 'API permissions' and add these Application permissions:"
                Write-Host -ForegroundColor White "   Microsoft Graph:"
                Write-Host -ForegroundColor White "   - Sites.Read.All"
                Write-Host -ForegroundColor White "   - Application.Read.All"
                Write-Host -ForegroundColor White "   SharePoint:"
                Write-Host -ForegroundColor White "   - Sites.Read.All"
                Write-Host -ForegroundColor White "   - Sites.Manage.All"
                Write-Host -ForegroundColor White "   - Sites.FullControl.All"
                Write-Host -ForegroundColor White "9. Click 'Grant admin consent'"
                Write-Host -ForegroundColor White "10. Copy the Application (client) ID and update your script variables"
                
                Write-Host -ForegroundColor Yellow "`nReturning to main menu..."
                Show-MainMenu
            }
        }
    }
    catch {
        Write-LogEntry -LogName $logPath -LogEntryText "App creation flow failed: $_" -LogLevel "ERROR"
        Write-Host -ForegroundColor Red "App creation flow failed: $_"
        Write-Host -ForegroundColor Yellow "Returning to main menu..."
        Show-MainMenu
    }
}

# Function to handle assessment flow with existing settings
function Start-AssessmentFlow {
    try {
        # Validate M365 Assessment app configuration
        Write-Host -ForegroundColor Yellow "🔍 Validating M365 Assessment app configuration..."
        
        $configValid = $true
        $validationErrors = @()
        
        # Check if assessment tool exists
        if (-not (Test-Path $assessmentToolPath)) {
            $configValid = $false
            $validationErrors += "❌ Microsoft 365 Assessment tool not found at: $assessmentToolPath"
        }
        
        # Check if we have M365 Assessment app ID and thumbprint
        if ([string]::IsNullOrWhiteSpace($global:m365AssessmentAppID)) {
            $configValid = $false
            $validationErrors += "❌ M365 Assessment App ID is not configured"
        }
        
        if ([string]::IsNullOrWhiteSpace($global:m365AssessmentThumbprint)) {
            $configValid = $false
            $validationErrors += "❌ M365 Assessment certificate thumbprint is not configured"
        }
        
        # Check if certificate exists in store
        if ($global:m365AssessmentThumbprint) {
            try {
                Get-Item "Cert:\CurrentUser\My\$global:m365AssessmentThumbprint" -ErrorAction Stop | Out-Null
                Write-Host -ForegroundColor Green "✅ M365 Assessment certificate found in certificate store"
            }
            catch {
                $configValid = $false
                $validationErrors += "❌ M365 Assessment certificate with thumbprint '$global:m365AssessmentThumbprint' not found in CurrentUser\My store"
            }
        }
        
        if (-not $configValid) {
            Write-Host -ForegroundColor Red "`n❌ M365 Assessment app configuration validation failed:"
            foreach ($validationError in $validationErrors) {
                Write-Host -ForegroundColor Red "   $validationError"
            }
            Write-Host -ForegroundColor Yellow "`n💡 Suggestions:"
            Write-Host -ForegroundColor White "1. Download the Microsoft 365 Assessment tool from: https://github.com/pnp/pnpassessment/releases"
            Write-Host -ForegroundColor White "2. Use option 1 in main menu to create a new M365 Assessment app"
            Write-Host -ForegroundColor White "3. Use option 3 in main menu to auto-discover existing M365 Assessment app"
            Write-Host -ForegroundColor White "4. Use Set-M365AssessmentApp to manually configure app variables"
            Write-Host -ForegroundColor Yellow "`nReturning to main menu..."
            Show-MainMenu
            return
        }
        
        Write-Host -ForegroundColor Green "✅ M365 Assessment app configuration validation passed!"
        Write-Host -ForegroundColor Cyan "📋 Using M365 Assessment App:"
        Write-Host -ForegroundColor White "  • Tenant: $tenantname"
        Write-Host -ForegroundColor White "  • SharePoint URL: $sharepointTenantUrl"
        Write-Host -ForegroundColor White "  • M365 Assessment App ID: $global:m365AssessmentAppID"
        Write-Host -ForegroundColor White "  • M365 Assessment Certificate Thumbprint: $global:m365AssessmentThumbprint"
        Write-Host ""
        
        # Show assessment menu with M365 Assessment app settings
        Show-AssessmentMenu -AppId $global:m365AssessmentAppID -CertThumbprint $global:m365AssessmentThumbprint
        
        # After assessment menu, return to main menu
        Write-Host -ForegroundColor Yellow "`nReturning to main menu..."
        Show-MainMenu
    }
    catch {
        Write-LogEntry -LogName $logPath -LogEntryText "Assessment flow failed: $_" -LogLevel "ERROR"
        Write-Host -ForegroundColor Red "Assessment flow failed: $_"
        Write-Host -ForegroundColor Yellow "Returning to main menu..."
        Show-MainMenu
    }
}

# Main execution
try {
    # Display welcome message
    Write-Host -ForegroundColor Cyan "`n" + "=" * 60
    Write-Host -ForegroundColor Cyan "    Microsoft 365 Assessment Scanner"
    Write-Host -ForegroundColor Cyan "    Automated Assessment Tool with Azure App Integration"
    Write-Host -ForegroundColor Cyan "=" * 60
    
    # Auto-discover existing M365 Assessment app configuration
    Write-Host -ForegroundColor Cyan "`n🔍 Checking for existing M365 Assessment app configuration..."
    try {
        $discoveryResult = Get-M365AssessmentAppConfig
        
        if ($discoveryResult.Success) {
            Write-Host -ForegroundColor Green "✅ M365 Assessment app auto-configured!"
            Write-Host -ForegroundColor White "   Ready to run assessments with App ID: $($discoveryResult.AppId)"
        }
        else {
            Write-Host -ForegroundColor Yellow "⚠️  M365 Assessment app not auto-configured"
            if ($discoveryResult.RequiresAppCreation) {
                Write-Host -ForegroundColor Cyan "💡 Use option 1 in the menu to create a new M365 Assessment app"
            }
            elseif ($discoveryResult.RequiresCertInstall) {
                Write-Host -ForegroundColor Cyan "💡 App found but certificate needs to be installed locally"
                Write-Host -ForegroundColor Cyan "   Thumbprint: $($discoveryResult.Thumbprint)"
            }
            elseif ($discoveryResult.RequiresCertUpload) {
                Write-Host -ForegroundColor Cyan "💡 App found but needs certificate configuration"
            }
        }
    }
    catch {
        Write-Host -ForegroundColor Yellow "⚠️  Could not auto-discover M365 Assessment app (this is normal on first run)"
        Write-LogEntry -LogName $logPath -LogEntryText "Auto-discovery failed: $_" -LogLevel "INFO"
    }
    
    # Start with main menu
    Show-MainMenu
}
catch {
    Write-LogEntry -LogName $logPath -LogEntryText "Script execution failed: $_" -LogLevel "ERROR"
    Write-Host -ForegroundColor Red "Script execution failed: $_"
}

Write-Host ""
Write-Host -ForegroundColor Cyan "Script completed. Check log file for details: $logPath"
