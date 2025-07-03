<# 
.SYNOPSIS
    This script retrieves ALL test result attachments from Azure DevOps.

.DESCRIPTION
    This script connects to the Azure DevOps REST API to retrieve all attachments associated with test results in a specific project.

.PARAMETER Organization
    The name of the Azure DevOps organization. [string] Required.

.PARAMETER ProjectId
    The ID or name of the Azure DevOps project. [string] Required.

.PARAMETER Pat
    The Personal Access Token (PAT) used for authentication. [string] Required.
    
.PARAMETER TestRunID
    The ID of a specific test run to retrieve attachments for. [string] Optional. If not provided, attachments for all test runs in the project will be retrieved.

.PARAMETER SaveDirectory
    The directory where the attachments will be saved. [string] Optional. Default is "C:\TestRunAttachments".
    A subdirectory will be created for each test run, named with the test run ID and name, and attachments will be saved within that directory.

.PARAMETER ApiVersion
    The Azure DevOps REST API version to use. [string] Optional. Default is "7.2-preview".

.EXAMPLE
    Get-TestResultAttachments -Organization "your_organization" -ProjectId "your_project_id" -TestRunID "your_test_run_id" -Pat "your_PAT"

#>

param(
    [Parameter(Mandatory = $true)]
    [string]$Organization,

    [Parameter(Mandatory = $true)]
    [string]$ProjectId,

    [Parameter(Mandatory = $true)]
    [string]$Pat,

    [Parameter()]
    [string]$TestRunID, # Optional parameter to specify a test run ID

    [Parameter()]
    [string]$SaveDirectory = "C:\TestRunAttachments",

    [Parameter()]
    # [string]$ApiVersion = "7.1-preview.1" # Optional parameter to specify the API version
    [string]$ApiVersion = "7.2-preview" # Optional parameter to specify the API version
)

# Encode PAT as base64
$base64Pat = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(":$pat"))

# Function to make API requests
function Invoke-ApiRequest($url) {
    Write-Host "Invoking API request: $url"
    $headers = @{
        Authorization = "Basic $base64Pat"
    }
    $response = Invoke-RestMethod -Uri $url -Headers $headers
    return $response
}
# Get list of test results for the project or a specific test run if $testRunID is supplied
if ($null -ne $testRunID -and $testRunID -ne "") {
    # Get a specific test run by ID
    # Docs: https://learn.microsoft.com/en-us/rest/api/azure/devops/test/runs/get-test-run-by-id?view=azure-devops-rest-7.1&tabs=HTTP
    $testResultsUrl = "https://dev.azure.com/$($organization)/$($projectId)/_apis/test/runs/$($testRunID)?api-version=$($ApiVersion)"
    $testResult = Invoke-ApiRequest -url $testResultsUrl
    # Wrap single result in an array with a 'value' property for consistency
    $testResults = @{ value = @($testResult) }
} else {
    # Get all test runs for the project
    # Docs: https://learn.microsoft.com/en-us/rest/api/azure/devops/test/runs/list?view=azure-devops-rest-7.1&tabs=HTTP
    $testResultsUrl = "https://dev.azure.com/$($organization)/$($projectId)/_apis/test/runs?api-version=$($ApiVersion)"
    $testResults = Invoke-ApiRequest -url $testResultsUrl
}

# Get attachments for each test result
# This post helped: https://stackoverflow.com/questions/75022791/how-to-get-the-test-result-attachments-from-azure-devops-api
$attachments = @()
foreach ($result in $testResults.value) {

    # Get a list of test case results for the test run
    # This will include the test case results and their iterations
    # Docs: https://learn.microsoft.com/en-us/rest/api/azure/devops/test/results/get?view=azure-devops-rest-7.1&tabs=HTTP
    $testCaseResultsUrl = "https://dev.azure.com/$organization/$projectId/_apis/test/runs/$($result.id)/results?detailsToInclude=iterations&api-version=$apiVersion"
    $testCaseResults = Invoke-ApiRequest -url $testCaseResultsUrl

    # This is going to pull the attachments from the "Linked Items > Results Attachments" section of the test case result
    # Docs: https://learn.microsoft.com/en-us/rest/api/azure/devops/test/attachments/get-test-sub-result-attachment-zip?view=azure-devops-rest-5.1
    foreach ($testCaseResult in $testCaseResults.value) {
        $iterationDetails = $testCaseResult.iterationDetails
        if ($iterationDetails -and $iterationDetails.Count -gt 0) {
            foreach ($iteration in $iterationDetails) {
                if ($iteration.attachments -and $iteration.attachments.Count -gt 0) {
                    foreach ($attachment in $iteration.attachments) {
                        # add a value to the attachment array for a downloadURL 
                        $attachment | Add-Member -MemberType NoteProperty -Name "downloadUrl" -Value "https://dev.azure.com/$organization/$projectId/_apis/test/runs/$($result.id)/results/$($testCaseResult.id)/attachments/$($attachment.id)?api-version=$apiVersion" -Force
                        $attachment | Add-Member -MemberType NoteProperty -Name "fileName" -Value $attachment.Name -Force
                        $attachments += $attachment
                    }
                }
            }
        }
    }

    # This is going to pull the attachments from the "Attachments" section of the test case result
    foreach ($testCaseResult in $testCaseResults.value) {
        # Docs: https://learn.microsoft.com/en-us/rest/api/azure/devops/test/attachments/get-test-result-attachments?view=azure-devops-rest-7.1&tabs=HTTP
        $attachmentsUrl = "https://dev.azure.com/$organization/$projectId/_apis/test/runs/$($result.id)/results/$($testCaseResult.id)/attachments?api-version=$apiVersion"
        $resultAttachments = Invoke-ApiRequest -url $attachmentsUrl
        # add a value to the resultAttachments array for a downloadURL 
        $resultAttachments.value | ForEach-Object {
            $_ | Add-Member -MemberType NoteProperty -Name "downloadUrl" -Value "https://dev.azure.com/$organization/$projectId/_apis/test/runs/$($result.id)/results/$($testCaseResult.id)/attachments/$($_.id)?api-version=$apiVersion"
        }
        $attachments += $resultAttachments.value        
    }

    # Output the attachments for the current test run to the console
    $attachments | ForEach-Object {
        Write-Output "Attachment ID: $($_.id), Name: $($_.fileName), Size: $($_.size) bytes, `nDownload URL: $($_.downloadUrl)`n"
    }

    # Save attachments to a directory
    $safeRunName = $result.name -replace '[\\\/:\*\?"<>\|]', '_' # Replace illegal Windows filename characters with underscores in the run name
    $saveDirectory = $saveDirectory.TrimEnd('\', ' ') # Remove any trailing backslash or space characters from $saveDirectory
    
    $saveDirectory = "$saveDirectory\$($result.id)-$safeRunName\"
    if (-not (Test-Path -Path $saveDirectory)) {
        New-Item -ItemType Directory -Path $saveDirectory
    }

    # Download each attachment to disk
    Write-Host "Downloading attachments for test run: $($result.name) to $saveDirectory"
    foreach ($attachment in $attachments) {
        $filePath = Join-Path -Path $saveDirectory -ChildPath $attachment.fileName
        # If file exists, append a number to the filename
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($attachment.fileName)
        $ext = [System.IO.Path]::GetExtension($attachment.fileName)
        $counter = 1
        while (Test-Path $filePath) {
            $filePath = Join-Path -Path $saveDirectory -ChildPath ("{0}_{1}{2}" -f $baseName, $counter, $ext)
            $counter++
        }
        Invoke-WebRequest -Uri $attachment.downloadUrl -OutFile $filePath -Headers @{ Authorization = "Basic $base64Pat" }
    }
 
    # Output the total number of attachments downloaded
    Write-Host "Total attachments downloaded: $($attachments.Count)"
}
