# Microsoft 365 Assessment Scanner

A comprehensive PowerShell script for automated Microsoft 365 assessments using Azure App Registration with certificate-based authentication. This tool streamlines the process of creating Azure applications, managing certificates, and running Microsoft 365 assessments for migration planning and modernization efforts.

## 🚀 Features

### Core Functionality
- **Automated Azure App Registration**: Creates Azure AD applications with proper permissions
- **Certificate Management**: Generates, exports, and manages X.509 certificates for authentication
- **Microsoft 365 Assessment Integration**: Seamless integration with the Microsoft 365 Assessment tool
- **Auto-Discovery**: Automatically discovers and configures existing M365 Assessment applications
- **Interactive Menus**: User-friendly console interface for all operations
- **Comprehensive Logging**: Detailed logging for troubleshooting and audit purposes

### Assessment Types Supported
- **SharePoint 2013 Workflows**: Assess workflows for Power Automate migration
- **InfoPath Forms**: Analyze InfoPath forms for modernization opportunities
- **Add-ins and ACS**: Evaluate SharePoint Add-ins and Azure ACS dependencies
- **SharePoint Alerts**: Review alerts usage and configuration

### Assessment Operations
- **Execute**: Start new assessments with customizable scope
- **Status**: Monitor assessment progress with real-time updates
- **Report**: Generate Power BI reports and CSV exports with organized folder structure

## 📋 Prerequisites

### Software Requirements
- **PowerShell 5.1** or later
- **Microsoft 365 Assessment Tool** - Download from [GitHub Releases](https://github.com/pnp/pnpassessment/releases)
- **Azure PowerShell modules** (automatically imported if available)

### Azure Requirements
- **Azure AD Tenant** with administrative access
- **Global Administrator** or **Application Administrator** role
- **Existing Azure Application** with the following permissions:
  - `Application.ReadWrite.All`
  - `Directory.ReadWrite.All`
  - `AppRoleAssignment.ReadWrite.All`

### Microsoft 365 Requirements
- **SharePoint Online** tenant
- **Sites to assess** (can be specific sites or entire tenant)

## 🛠️ Installation

### 1. Download Prerequisites
```powershell
# Download Microsoft 365 Assessment Tool
# Visit: https://github.com/pnp/pnpassessment/releases
# Extract to: C:\temp\microsoft365-assessment.exe
```

### 2. Clone or Download Script
```bash
git clone https://github.com/yourusername/m365-assessment-scanner.git
# OR download 365-Assessment-Scanner.ps1 directly
```

### 3. Configure Variables
Edit the script variables at the top of `365-Assessment-Scanner.ps1`:

```powershell
# Tenant Information
$tenantname = "yourtenant.onmicrosoft.com"
$sharepointTenantUrl = "yourtenant.sharepoint.com"
$tenantid = "your-tenant-id-guid"

# EXISTING App with Permissions to Create New Apps
$appID = "your-existing-app-id"
$thumbprint = "your-existing-certificate-thumbprint"
$CertStoreLocation = "LocalMachine"  # or "CurrentUser"

# Tool Paths
$assessmentToolPath = "c:\temp\microsoft365-assessment.exe"
$assessmentReportsPath = "c:\temp\assessmentreports"
$certexportpath = "c:\temp"
```

## 🏃‍♂️ Quick Start

### Method 1: Automated Setup (Recommended)
```powershell
# Load the script
. .\365-Assessment-Scanner.ps1

# Create certificate and app in one step
$result = New-CertificateAndAzureApp -AppDisplayName "My M365 Assessment App"

# Start assessment
Show-AssessmentMenu
```

### Method 2: Manual Setup
```powershell
# Load the script
. .\365-Assessment-Scanner.ps1

# Step 1: Create certificate
$certInfo = New-Cert

# Step 2: Create Azure application
$appResult = New-AzureAppRegistration -AppDisplayName "My Assessment App" -CertificateFilePath $certInfo.CerFilePath

# Step 3: Configure for assessments
Set-M365AssessmentApp -AppId $appResult.ApplicationId -CertThumbprint $certInfo.Thumbprint

# Step 4: Run assessments
Show-AssessmentMenu
```

### Method 3: Auto-Discovery (Existing Apps)
```powershell
# Load the script
. .\365-Assessment-Scanner.ps1

# Auto-discover existing configuration
$discoveryResult = Get-M365AssessmentAppConfig
```

## 📖 Detailed Usage

### Main Menu Options

When you run the script, you'll see an interactive menu:

```
🚀 Microsoft 365 Assessment Scanner
======================================
Current Configuration:
  • Tenant: yourtenant.onmicrosoft.com
  • SharePoint URL: yourtenant.sharepoint.com
  • Creation App ID: 1e488dc4-1977-48ef-8d4d-9856f4e04536
  • Creation App Thumbprint: 5EAD7303A5C7E27DB4245878AD554642940BA082
  • Assessment Tool: c:\temp\microsoft365-assessment.exe

M365 Assessment App Status:
  ✅ M365 Assessment App Configured
     • App ID: a9891c1e-188b-45a7-a455-1a04c5e60a1a
     • Thumbprint: 7A145A41D29D7A90F208DE33E61E82AAD7DF06AC

Choose an option:
1. 🔧 Create New M365 Assessment App (with certificate)
2. 📊 Run Assessment with M365 Assessment App
3. 🔍 Auto-discover Existing M365 Assessment App
4. ❌ Exit
```

### Assessment Workflow

#### 1. Execute Assessment
```powershell
# Select assessment type (Workflow, InfoPath, AddInsACS, Alerts)
# Choose scope (entire tenant or specific sites)
Invoke-M365Assessment -AssessmentType "Workflow" -Operation "Execute"
```

#### 2. Monitor Progress
```powershell
# Check assessment status
Invoke-M365Assessment -AssessmentType "Workflow" -Operation "Status"
```

#### 3. Generate Reports
```powershell
# Generate Power BI and CSV reports
Invoke-M365Assessment -AssessmentType "Workflow" -Operation "Report" -AssessmentId "40e5fe55-108e-4cee-92ee-f052e008534a"
```

### Report Organization

Reports are automatically organized by assessment ID:
```
c:\temp\assessmentreports\
├── 40e5fe55-108e-4cee-92ee-f052e008534a\
│   ├── PowerBI_Report.pbit
│   ├── Workflow_Data.csv
│   └── Summary.csv
├── 22989c75-f08f-4af9-8857-6f19e333d6d3\
│   ├── InfoPath_Analysis.csv
│   └── PowerBI_Report.pbit
```

## 🔧 Core Functions

### Certificate Management
```powershell
# Create new certificate
New-Cert

# Returns certificate information:
# - Thumbprint
# - CER file path
# - PFX file path
```

### Azure App Registration
```powershell
# Create Azure application with certificate
New-AzureAppRegistration -AppDisplayName "My App" -CertificateFilePath "C:\temp\cert.cer"

# Automatic permission grants:
# Microsoft Graph: Sites.Read.All, Application.Read.All
# SharePoint: Sites.Read.All, Sites.Manage.All, Sites.FullControl.All
```

### Assessment Operations
```powershell
# Execute assessment
Invoke-M365Assessment -AssessmentType "Workflow" -Operation "Execute" -SitesList "https://tenant.sharepoint.com/sites/site1"

# Check status with multiple display options
Invoke-M365Assessment -AssessmentType "Workflow" -Operation "Status"

# Generate reports with auto-organized folders
Invoke-M365Assessment -AssessmentType "Workflow" -Operation "Report" -AssessmentId "assessment-guid"
```

### Auto-Discovery
```powershell
# Discover existing M365 Assessment app
Get-M365AssessmentAppConfig

# Features:
# - Finds app by display name
# - Extracts certificate from app registration
# - Validates local certificate store
# - Auto-configures global variables
```

## 🔐 Security & Permissions

### Required Azure AD Permissions

#### Creation App (Existing)
- `Graph: Application: Application.ReadWrite.All` - Create and manage applications
- `Graph: Application: Directory.ReadWrite.All` - Read and write directory data  
- `Graph: Application: AppRoleAssignment.ReadWrite.All` - Grant permissions to applications

#### M365 Assessment App (Created)
**Microsoft Graph:**
- `Sites.Read.All` - Read all site collections
- `Application.Read.All` - Read all applications

**SharePoint:**
- `Sites.Read.All` - Read all site collections
- `Sites.Manage.All` - Manage all site collections
- `Sites.FullControl.All` - Full control of all site collections

### Certificate Security
- **Self-signed certificates** for authentication
- **CurrentUser certificate store** for new certificates
- **Exportable private keys** for backup and portability
- **10-year validity period** (configurable)

## 📊 Assessment Types Details

### 1. SharePoint 2013 Workflows
**Purpose**: Identify workflows that need migration to Power Automate
- **Scope**: Site collections, lists, libraries
- **Output**: Workflow inventory, complexity analysis, migration recommendations
- **Reports**: PowerBI dashboard, detailed CSV export

### 2. InfoPath Forms
**Purpose**: Analyze InfoPath forms for modernization
- **Scope**: Form templates, published forms, data connections
- **Output**: Form complexity, modernization path, PowerApps readiness
- **Reports**: Migration strategy, effort estimation

### 3. Add-ins and ACS
**Purpose**: Evaluate SharePoint Add-ins and Azure ACS dependencies
- **Scope**: App catalog, installed apps, ACS configurations
- **Output**: Add-in inventory, security assessment, modernization plan
- **Reports**: Risk analysis, replacement recommendations

### 4. SharePoint Alerts
**Purpose**: Review alerts usage and configuration
- **Scope**: Site alerts, list alerts, user subscriptions
- **Output**: Alert usage patterns, potential replacements
- **Reports**: Cleanup recommendations, modern alternatives

## 🛠️ Troubleshooting

### Common Issues

#### Certificate Not Found
```
Error: Certificate with thumbprint 'XXX' not found
Solution: Run auto-discovery or manually configure certificate
```

#### Permission Denied
```
Error: Insufficient privileges to complete the operation
Solution: Ensure proper Azure AD roles and API permissions
```

#### Assessment Tool Not Found
```
Error: Microsoft 365 Assessment tool not found
Solution: Download tool and update $assessmentToolPath variable
```

#### App Registration Failed
```
Error: Failed to create Azure application registration
Solution: Check existing app permissions and tenant role
```

### Log Files
Check logs at: `$env:TEMP\Create-AzureApp_YYYY-MM-DD_HH-mm-ss.log`

## 📝 Configuration Examples

### Enterprise Environment
```powershell
# High-security environment with LocalMachine certificates
$CertStoreLocation = "LocalMachine"
$assessmentReportsPath = "D:\M365Assessments\Reports"
$certexportpath = "D:\M365Assessments\Certificates"
```

### Development Environment
```powershell
# Developer workstation with CurrentUser certificates
$CertStoreLocation = "CurrentUser"
$assessmentReportsPath = "C:\Dev\M365Reports"
$certexportpath = "C:\Dev\Certificates"
```

### Multi-Tenant Setup
```powershell
# Script supports multiple tenants by updating variables
$tenantname = "tenant1.onmicrosoft.com"
$tenantid = "tenant1-guid"
# Run assessments...

$tenantname = "tenant2.onmicrosoft.com"  
$tenantid = "tenant2-guid"
# Run assessments...
```

## 🤝 Contributing

### Reporting Issues
1. Check existing [Issues](https://github.com/yourusername/m365-assessment-scanner/issues)
2. Create new issue with:
   - PowerShell version
   - Azure environment details
   - Error messages and logs
   - Steps to reproduce

### Feature Requests
- Assessment type additions
- Report format enhancements
- Authentication method alternatives
- UI/UX improvements

### Pull Requests
1. Fork the repository
2. Create feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'Add amazing feature'`)
4. Push to branch (`git push origin feature/amazing-feature`)
5. Open Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- **Microsoft 365 Assessment Tool** - [PnP Community](https://github.com/pnp/pnpassessment)
- **Microsoft Graph API** - Azure integration capabilities
- **PowerShell Community** - Module and scripting best practices

## 📞 Support

### Community Support
- **GitHub Issues**: [Project Issues](https://github.com/yourusername/m365-assessment-scanner/issues)
- **Discussions**: [GitHub Discussions](https://github.com/yourusername/m365-assessment-scanner/discussions)

### Documentation
- **Microsoft 365 Assessment**: [Official Documentation](https://docs.microsoft.com/en-us/assessments/)
- **Azure App Registration**: [Azure AD Documentation](https://docs.microsoft.com/en-us/azure/active-directory/develop/)
- **Microsoft Graph**: [Graph API Reference](https://docs.microsoft.com/en-us/graph/)

---

### Key Variables to Configure
```powershell
$tenantname = "yourtenant.onmicrosoft.com"
$tenantid = "your-tenant-guid"
$appID = "existing-app-with-creation-permissions"
$thumbprint = "existing-app-certificate-thumbprint"
$assessmentToolPath = "path-to-assessment-tool"
```

---

**Happy Assessing! 🎯**

> This tool is designed to streamline Microsoft 365 assessments and migration planning. For questions, issues, or contributions, please use the GitHub repository features.
