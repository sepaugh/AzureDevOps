<#
.SYNOPSIS
    When enabled, IIS Request Filtering interprets the System and Custom fields within Azure DevOps as physical files, and will block them from loading if they're not explicilty defined.
    This script adds all Azure DevOps Work Item Fields, and installed file types as either allowed or denied file extensions in IIS Request Filtering configuration.

    This does NOT work for Azure DevOps Services (cloud), it is intended for Azure DevOps Server (on-premises) installations.

.DESCRIPTION
    This script retrieves all field definitions from an Azure DevOps organization or project using the Azure DevOps REST API, extracts the field reference names as file extensions, and adds them to the IIS Request Filtering configuration for the specified IIS site. 
    It supports processing all projects or a specific project, and allows you to specify whether the extensions should be allowed or denied.
    The script uses the WebAdministration module to interact with IIS and requires appropriate permissions to modify IIS configuration.

.PARAMETER collectionURL
    The base URL of the Azure DevOps organization (e.g., "https://dev.contoso.com/yourOrganization" or "http://localhost:8080/tfs/DefaultCollection").

.PARAMETER allProjects
    Switch. If set, retrieves fields across all projects in the organization. If not set, requires -projectId.

.PARAMETER projectId
    The name or ID of the Azure DevOps project to retrieve fields from. Required unless -allProjects is specified.

.PARAMETER pat
    The Personal Access Token used for authenticating with the Azure DevOps REST API.

.PARAMETER allowed
    Boolean. If $true (default), the extensions will be allowed in IIS. If $false, they will be denied.

.PARAMETER siteName
    The name of the IIS site as seen in IIS Manager. Default is 'Azure DevOps Server'.

.EXAMPLE
    PS C:\> .\Add-ADOFields_To_IISRequestFiltering.ps1 -collectionURL "https://dev.contoso.com/yourOrganization" -projectId "MyProject" -pat "xxxxxxxxxx" -allowed $true
    Retrieves all field extensions from the specified Azure DevOps project and adds them as allowed file extensions to the IIS Request Filtering configuration for the "Azure DevOps Server" site.

.EXAMPLE
    PS C:\> .\Add-ADOFields_To_IISRequestFiltering.ps1 -collectionURL "https://dev.contoso.com/yourOrganization" -allProjects -pat "xxxxxxxxxx" -allowed $false
    Retrieves all field extensions from all projects and adds them as denied file extensions to the default IIS site.

.NOTES
    Author: Lorne Sepaugh
    Date: 2024-07-11
    - Requires the WebAdministration module and administrative privileges to modify IIS configuration.
    - The script restarts IIS at the end to apply changes.

    This script comes with no warranty and no support. Use at your own risk.
    It is recommended to test this script in a development environment before applying it to production systems.

.LINK 
    https://learn.microsoft.com/en-us/rest/api/azure/devops/wit/fields/list
    https://learn.microsoft.com/en-us/iis/manage/configuring-security/configure-request-filtering-in-iis
#>

param (
    [Parameter(Mandatory = $true)]
        [string]$collectionURL, # Example: "https://dev.contoso.com/yourOrganization" or "http://localhost:8080/tfs/DefaultCollection"
    [Parameter(Mandatory = $false, ParameterSetName = 'AllProjects')]
        [switch]$allProjects, # If set, process all projects and do not require projectId
    [Parameter(Mandatory = $true, ParameterSetName = 'ByProjectId')]
        [string]$projectId, # Name of the project or its ID
    [Parameter(Mandatory = $true)]
        [string]$pat, # Personal Access Token for authentication
    [bool]$allowed = $true, # Change to allow or deny file extensions
    [string]$siteName = 'Azure DevOps Server' # The name of the ADO website as seen in IIS
)

## Check if the script is running as an administrator, if not, restart as an administrator
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{   
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    exit 000
} 

Start-Transcript -Path "$($PSScriptRoot)\Add-WIT_Fields_To_IIS_RequestFiltering_$(Get-Date -Format 'yyyyMMdd_HHmmss').log" -Append -NoClobber
filter timestamp {"$(Get-Date -Format o): $_"}

# Non-parameter variables configuration
# These variables are used to configure the script and should not be changed by the user.
$filterPath = "system.webServer/security/requestFiltering/fileExtensions" # Path to the fileExtensions configuration in IIS

# Ensure collectionURL and projectId do not end with a slash (optional, only if needed)
if ($collectionURL.EndsWith('/')) {
    $collectionURL = $collectionURL.TrimEnd('/')
}
if ($projectId.EndsWith('/')) {
    $projectId = $projectId.TrimEnd('/')
}

# Load the WebAdministration module to manage IIS
# This module is required to interact with IIS configuration, should be loaded by default on Windows Server with IIS installed.
# If not available, ensure IIS is installed and the module is accessible.
Import-Module WebAdministration


# Check if "Allow unlisted file name extensions" is enabled in IIS for the specified site
# If it is not enabled, we may fail to retrieve the list of fields from the Azure DevOps REST API
# We'll enable this temporarily, and then disable it after the script runs
$allowUnlisted = Get-WebConfigurationProperty -Filter "$filterPath" -PSPath "IIS:\Sites\$siteName" -Name "allowUnlisted" -ErrorAction SilentlyContinue

if ($null -eq $allowUnlisted) {
    Write-Output "Could not determine the 'Allow unlisted file name extensions' setting for site '$siteName'." | timestamp
} elseif ($allowUnlisted.Value) {
    Write-Output "'Allow unlisted file name extensions' is ENABLED for site '$siteName'." | timestamp
} else {
    Write-Output "'Allow unlisted file name extensions' is DISABLED for site '$siteName'." | timestamp
    Write-Warning "'Allow unlisted file name extensions' is currently disabled, this could lead to issues with retrieving an updated fields list."
    $continue = Read-Host "Do you want to enable this setting? We will also prompt to disable at the end. (Y/N)"
    if ($continue -notin @('Y', 'y')) {
        Write-Output "'Allow unlisted file name extensions' remains DISABLED for site '$siteName' at user request." | timestamp
        # Stop-Transcript
        # exit 2
    }
    Set-WebConfigurationProperty -Filter "$filterPath" -PSPath "IIS:\Sites\$siteName" -Name "allowUnlisted" -Value $true
    Write-Output "'Allow unlisted file name extensions' has been ENABLED for site '$siteName'." | timestamp
}

# Ensure PAT encoding is correct to Base64
$headers = @{
    Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$pat"))
}

# Call the Azure DevOps REST API to get the list of fields
# This API call retrieves all fields available in the specified project.
try {
    Write-Output "Retrieving fields from Azure DevOps..." | timestamp

    if ($allProjects) {
        # API call to get fields across all projects
        $uri = "$($collectionURL)/_apis/wit/fields?api-version=7.1"
        Write-Output "All projects mode: $uri" | timestamp
    } else {
        # API call to get fields in the specified project
        $uri = "$($collectionURL)/$($projectId)/_apis/wit/fields?api-version=7.1"
        Write-Output "Project mode: $uri" | timestamp
    }

    $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
}
catch {
    Write-Output "Failed to retrieve fields from Azure DevOps: $_" | timestamp
    exit 1
}

# Process the response to extract field names, reference names, and types
# The response contains a list of fields, each with properties like name, referenceName, 'extensionName' and type.
# Output the fields to a CSV file for reference
Write-Output "Processing fields from Azure DevOps..." | timestamp
$response.value | ForEach-Object {
    [PSCustomObject]@{
        Name        = $_.name
        Reference   = $_.referenceName
        ExtensionName = ".$(($($_.referenceName) -split '\.')[-1])"
        Type       = $_.type
    }
} | ConvertTo-CSV -NoTypeInformation | Out-File -FilePath "$($PSScriptRoot)\ADOFields.csv" -Encoding UTF8

# The following are all file extensions found under "$ENV:ProgramFiles\Azure DevOps Server 2022\Application Tier\Web Services"
$adoExtensions = @(
    ".asax",        ".ascx",
    ".ashx",        ".asmx",
    ".aspx",        ".config",
    ".cshtml",      ".css",
    ".dll",         ".eot",
    ".gif",         ".hbs",
    ".htm",         ".html",
    ".ico",         ".jpg",
    ".js",          ".json",
    ".man",         ".map",
    ".master",      ".md",
    ".mp3",         ".pdf",
    ".png",         ".rsp",
    ".svg",         ".targets",
    ".template",    ".ts",
    ".ttf",         ".vsomanifest",
    ".wav",         ".woff",
    ".woff2",       ".xml",
    ".xsl",         
    "." # This actually allows the site to load, without it, and 'Allow Unlisted file name extensions' denied, the site will not load.
)

# The following are the extensions from Azure DevOps fields
# The 'referenceName' property split is used to determine the file extension as IIS sees it.
$adoExtensions += @(
    $response.value | ForEach-Object {
        ".$(($_.referenceName -split '\.')[-1])"
    }
)

# Remove all existing allowed extensions (optional, be careful!), uncomment the next line to clear existing extensions
# Clear-WebConfiguration -Filter $filterPath -PSPath "IIS:\Sites\$siteName"

Write-Output "Found $($adoExtensions.Count) fields from Azure DevOps for processing..." | timestamp

# Prime IIS for incoming changes
# This command prepares IIS to accept configuration changes without applying them immediately.
Start-IISCommitDelay

# Add each extension to IIS Request Filtering
# This section iterates through the list of ADO extensions and adds them to the IIS configuration
$extCount = ($adoExtensions | Sort-Object -Unique).Count
$index = 0

foreach ($ext in $adoExtensions | Sort-Object -Unique) {
    $index++
    Write-Progress -Activity "Adding Extensions" -Status "Processing $ext ($index of $extCount)" -PercentComplete (($index / $extCount) * 100)
    
    # Check if extension already exists
    $existing = Get-WebConfigurationProperty -Filter "$filterPath/add[@fileExtension='$($ext)']" -PSPath "IIS:\Sites\$siteName" -Name "fileExtension" -ErrorAction SilentlyContinue
    try {
        if (-not $existing) {
            Add-WebConfigurationProperty -pspath "IIS:\Sites\$($siteName)" -filter $filterPath -name "." -value @{fileExtension=$($ext); allowed=$($allowed)}
            Write-Output "Added extension: $ext" | timestamp
        }
        else {
            Write-Output "Extension $ext already exists, skipping." | timestamp
        }
    }
    catch {
        Write-Output "Error processing extension $($ext): $_" | timestamp
    }
}
Write-Progress -Activity "Adding Extensions" -Completed

Write-Output "All extensions processed. Total processed: $($adoExtensions.Count)" | timestamp

# Commit the changes to IIS
Write-Output "Restarting IIS to apply changes..." | timestamp
Stop-IISCommitDelay # Commits changes to IIS
IISReset # Restart IIS to apply changes

# Output the list of allowed extensions
Write-Output "List of all currently allowed file extensions for $($siteName):" | timestamp
Get-WebConfigurationProperty -Filter "$filterPath/add" -PSPath "IIS:\Sites\$siteName" -Name "fileExtension" -ErrorAction SilentlyContinue |
    Sort-Object Value |
    ForEach-Object {
        Write-Output "  $($_.Value)"
    }


# After the script runs, we can disable the "Allow unlisted file name extensions" setting
$allowUnlisted = Get-WebConfigurationProperty -Filter "$filterPath" -PSPath "IIS:\Sites\$siteName" -Name "allowUnlisted" -ErrorAction SilentlyContinue

if ($null -eq $allowUnlisted) {
    Write-Output "Could not determine the 'Allow unlisted file name extensions' setting for site '$siteName'." | timestamp
} elseif (-not $allowUnlisted.Value) {
    Write-Output "'Allow unlisted file name extensions' is already DISABLED for site '$siteName'." | timestamp
} else {
    Write-Output "'Allow unlisted file name extensions' is ENABLED for site '$siteName'." | timestamp
    Write-Warning "'Allow unlisted file name extensions' is currently enabled. It is recommended to disable this setting for security."
    $continue = Read-Host "Do you want to DISABLE this setting? (Y/N)"
    if ($continue -notin @('Y', 'y')) {
        Write-Output "'Allow unlisted file name extensions' remains ENABLED for site '$siteName' at user request." | timestamp
        exit 2
    }
    Set-WebConfigurationProperty -Filter "$filterPath" -PSPath "IIS:\Sites\$siteName" -Name "allowUnlisted" -Value $false
    Write-Output "'Allow unlisted file name extensions' has been DISABLED for site '$siteName'." | timestamp
}

# End of script
Write-Output "Script completed successfully." | timestamp
Stop-Transcript
