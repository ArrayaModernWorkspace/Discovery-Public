<#
.SYNOPSIS
This script is designed to automate the process of gathering, processing, and exporting data related to a Microsoft 365 tenant environment.

.DESCRIPTION
The script follows a structured approach comprising the following main sections:

1. Initialization:
   - Establishes connections to all necessary Office 365 services.
   - Checks and installs the ImportExcel module if needed.
   - Retrieves the export path, initializes various variables and arrays.
   - This will prioritize running the Microsoft Graph PowerShell SDK, especially if running Windows 7 PowerShell.
   - If a lesser version of PowerShell or Microsoft Graph PowerShell SDK is not installed, the other modules of AzureAD, MsOnline, SharePointOnline, MicrosoftTeams will be utilized along with ExchangeOnlineManagement

2. Data Gathering:
   - Collects data from Exchange Online, Collaboration/SharePoint, and other tenant objects and license details.
   - The data collection functions are parameterized to control the depth or scope of the data gathered based on the `detailLevel` argument.

3. Data Consolidation (Optional):
   - Optionally consolidates discovery report data into one file based on the reporting mode.

4. Export Preprocessing:
   - Processes the collected data to remove certain tables based on the reporting mode.
   - Creates a new hashtable to hold the data for export, excluding specified tables.

5. Data Export:
   - Exports the processed data to Excel, with error handling to capture and log any issues encountered during the export process.

6. Error Export:
   - Exports any captured errors to a separate file, and displays a summary of error details to the user.

7. Final Reporting:
   - Calculates and displays the total execution time.
   - Generates a final summary table of object counts, which is output to the console.

8. Logging and Error Handling:
   - Utilizes extensive logging and error handling to ensure that issues are captured, logged, and reported in a user-friendly manner.

.FUNCTIONS

- Connect-Office365Services:
   Establishes connections to specified Office 365 services based on the input parameters.

- Install-ImportExcelModule:
   Checks for, and installs the ImportExcel module, which is needed for exporting data to Excel.

- SelectReportMode:
   Prompts the user to select the reporting mode which determines the level of detail in the data gathering process.

- Get-ExportPath:
   Retrieves the path where the exported data will be stored.

- Get-AllRecipientDetails, Get-AllExchangeMailboxDetails, Get-ExchangeGroupDetails, Get-MailFlowRulesandConnectors, Get-AllPublicFolderDetails:
   Functions for gathering data from Exchange Online.

- Get-AllUnifiedGroups, Get-SPOAndOneDriveDetails, Get-TeamsDetails:
   Functions for gathering data from SharePoint and Teams.

- Get-AllLicenseSKUs, Get-allUserDetails, Get-AllOffice365Domains, Get-AllOffice365Admins:
   Functions for gathering tenant-wide data and license details.
   These Functions support Microsoft Graph and MSOnline connections

- Combine-AllMailboxStats:
   Optional function to consolidate discovery report data into a single file.

- Export-HashTableToExcel:
   Exports the gathered data to Excel, creating one file per hashtable.

- Export-ErrorReports:
   Exports any errors encountered during the script execution to a separate file.

- Capture-ErrorHelper:
   Helper function to capture and format error information for logging.

- Write-Log:
   Function to write log entries to a file, with options for specifying the type of log entry (INFO, WARNING, ERROR).

.NOTES
The script utilizes modular functions to perform specific tasks and employs logging and error handling to provide a robust solution for data collection and reporting.
The user-friendly prompts and color-coded console output enhance the user experience when running the script.

.AUTHOR
Aaron Medrano

.DATE
Created Date: 2021-04-19
Last Modified Date: 2024-06-06

.EXAMPLE
# How to run the script interactively
1. Save the script to a file named Get-FullTenantReportDetails.ps1.
2. Open PowerShell as an Administrator.
3. Navigate to the directory containing the script:
   cd "path\to\your\script\directory"
4. Run the script:
   .\Get-FullTenantReportDetails.ps1
5. Follow the prompts to provide necessary input such as ReportingMode, ExportPath, Authentication, and others as prompted.
6. Check the specified export path for the generated reports and error logs.
7. Review the console output for any errors or summary information provided by the script.

# The script will connect to all Office 365 services, gather data, and export it to the specified path.
# Errors will be logged and a summary of the data gathered will be output to the console.

#>

########################################################
#Intial Variables and Functions
########################################################

# ----------------------------------
# Default Script Helper Functions
# ----------------------------------

#Progress Helper OG
function Write-ProgressHelper {
    param (
        [int]$ProgressCounter,
        [string]$Activity,
        [string]$ID,
        [string]$CurrentOperation,
        [switch]$Completed,
        [int]$TotalCount,
        [datetime]$StartTime
    )
    # if the progress bar is set to silently continue, change it to continue
    if ($ProgressPreference = "SilentlyContinue") {
        $ProgressPreference = "Continue"
    }    

    $progressParameters = @{
        Activity = $Activity
    }

    if ($StartTime -and $ProgressCounter -and $TotalCount) {
        $secondsElapsed = (Get-Date) - $StartTime
        $secondsRemaining = ($secondsElapsed.TotalSeconds / $progresscounter) * ($TotalCount - $progresscounter)
        $progresspercentcomplete = [math]::Round((($progresscounter / $TotalCount)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$TotalCount+"]"

        $progressParameters.Status = "$progressStatus $($secondsElapsed.ToString('hh\:mm\:ss'))"
        $progressParameters.PercentComplete = $progresspercentcomplete
    }

    # if we have an estimate for the time remaining, add it to the Write-Progress parameters
    if ($secondsRemaining) {
        $progressParameters.SecondsRemaining = $secondsRemaining
    }

    if ($ID) {
        $progressParameters.ID = $ID
    }

    if ($CurrentOperation) {
        $progressParameters.CurrentOperation = $CurrentOperation
    }

    # If the Completed switch is provided, mark the progress as completed
    if ($Completed) {
        $progressParameters.Completed = $true
    }

    # Write the progress bar
    Write-Progress @progressParameters

}

function Capture-ErrorHelper {
    param(
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecordVar,
        [Parameter(Mandatory=$true)]
        [string]$errorMessage
    )
    $currentErrors = @()
    Write-Host $errorMessage -ForegroundColor Red
    if($ErrorRecordVar) {
        foreach ($errorCheck in $ErrorRecordVar) {
            if ($errorCheck.Exception.Message -match "'[^']*/[^']*'") {
                $recipient = $matches[0].Trim("'")
            } else {
                $recipient = $null
            }
            $CurrentError = [PSCustomObject]@{
                TimeStamp           = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                ErrorMessage        = $errorMessage
                Commandlet          = $errorCheck.CategoryInfo.Activity
                Reason              = $errorCheck.CategoryInfo.Reason
                "Exception-Message" = $errorCheck.Exception.Message
                Exception           = $errorCheck.Exception
                Recipient           = $recipient
                TargetObject        = $errorCheck.TargetObject
            }
            
            $currentErrors += $currentError
            #return $CurrentError
        }
        return $CurrentErrors
        Write-Error $CurrentErrors.errorMessage
    }
}

function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "DEBUG")]
        [string]$Type = "INFO",

        [Parameter(Mandatory=$true)]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [string]$LogPath,

        [Parameter(Mandatory=$false)]
        [string]$ExportFileLocation

    )
    # If the log file path is provided, append the log message to the file
    if ($LogPath) {
        # Get the directory, filename without extension, and the extension
        # Create the 'Log Reporting' directory if it doesn't exist
        if (-not (Test-Path $LogPath)) {
            $newfolder = New-Item -Path $LogPath -ItemType Directory
        }
        try {
            # Prepare the log message with a timestamp
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "$timestamp : [$Type] $Message"
            $LogFile = $LogPath + "\FullTenantReportLog.txt"
            $logMessage | Out-File -Append -FilePath $LogFile
        } catch {
            Write-Error "Failed to write to log file at '$LogFile': $_"
        }
    }
    elseif ($ExportFileLocation) {
        # Get the directory, filename without extension, and the extension
        $directory = [System.IO.Path]::GetDirectoryName($ExportFileLocation)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ExportFileLocation)
        $txtFileName = $baseName + "-FullReportLog.txt"
        $NewLogFolder = $baseName + " Discovery Log Reporting"
        $Newdirectory = Join-Path -Path $directory -ChildPath $NewLogFolder
        $LogFile = Join-Path -Path $Newdirectory -ChildPath $txtFileName

        # Create the 'Log Reporting' directory if it doesn't exist
        if (-not (Test-Path $Newdirectory)) {
            $newfolder = New-Item -Path $Newdirectory -ItemType Directory
        }
        try {
            # Prepare the log message with a timestamp
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "$timestamp : [$Type] $Message"
            $logMessage | Out-File -Append -FilePath $LogFile
        } catch {
            Write-Error "Failed to write to log file at '$LogFile': $_"
        }
    }

    # If the Verbose switch is used, also display the log message on the screen
    if ($VerbosePreference -eq 'Continue') {
        Write-Verbose $Message
    }
}

function Install-ImportExcelModule {
    # Check if ImportExcel module is installed
    if (!(Get-Module -ListAvailable -Name ImportExcel)) {
        try {
            Install-Module -Name ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
        }
        catch {
            Write-Warning "Could not install ImportExcel module. Defaulting to CSV output only."
            return $false
        }
    }

    # Import ImportExcel module
    try {
        Import-Module ImportExcel -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not import ImportExcel module. Defaulting to CSV output only."
        return $false
    }

    return $true
}

#Level of Detail Reporting
function Set-ReportMode {
    param (
        [switch]$ShowExplanationsOnStart = $true
    )

    function ShowModeExplanations {
        Write-Host "Reporting Mode Explanations:" -ForegroundColor Yellow
        Write-Host "Minimum   - Provides basic details for a quick overview."
        Write-Host "Combined  - Combines details from similar reports. E.g., combining a user's mailbox and SharePoint data."
        Write-Host "All       - Provides a comprehensive, detailed report that includes all possible details and combined reports."
        Write-Host "Geek      - Provides every available detail from reports, does not include Combined reports"

    }

    if ($ShowExplanationsOnStart) {
        ShowModeExplanations
        Write-Host ""
    }

    $selectedMode = $null
    do {
        $selectedMode = Read-Host "Please select a reporting mode (Minimum, Combined, All, Geek) or type 'help' for explanations"
        
        # Check if the user wants to see the explanations again
        if ($selectedMode -eq 'help') {
            ShowModeExplanations
            $selectedMode = $null  # Reset to null to continue the loop
        }
    } while ($selectedMode -notin @('Minimum', 'Combined', 'All', 'Geek'))

    $selectedMode = $selectedMode.ToLower()
    Write-Host "You selected: $selectedMode reporting mode" -ForegroundColor Green
    return $selectedMode
}

# ----------------------------------
# Export Script Functions
# ----------------------------------

#Convert Hash Table to Custom Object Array for Export
function Convert-HashToArray {
    [CmdletBinding()]
    param ( [Parameter(Mandatory=$true)]
        [Hashtable]$HashToConvert,
        [Parameter(Mandatory=$False)]
        [String]$tenant,
        [Parameter(Mandatory=$False)]
        [String]$table        
    )
    $ExportArray = @()
    $start = Get-Date
    $totalCount = ($HashToConvert.keys | measure).count    
    $progresscounter = 0

    foreach ($nestedKey in $HashToConvert.Keys) {
        $progresscounter++
        Write-ProgressHelper -Activity "Converting Hash Table" -CurrentOperation "Converting $($nestedKey)" -ProgressCounter $progresscounter -TotalCount $totalCount -StartTime $start
        #Define the attributes
        $attributes = $HashToConvert[$nestedKey]
        # If the attributes are a hashtable, convert them to a custom object
        if ($attributes -is [hashtable] -or $attributes -is [System.Collections.Specialized.OrderedDictionary]) {
            #Write-Verbose "Attributes are a hashtable"
            $customObject = New-Object -TypeName PSObject

            # Add the tenant name to the attribute name
            if ($tenant) {
                foreach ($attribute in $attributes.keys) {
                    $customObject | Add-Member -MemberType NoteProperty -Name "$($attribute)_$($tenant)" -Value ($attributes[$attribute] -join ';')
                }
            }
            # If the tenant name is not provided, add the attributes to the custom object
            else {
                foreach ($attribute in $attributes.keys) {
                    $customObject | Add-Member -MemberType NoteProperty -Name "$($attribute)" -Value ($attributes[$attribute] -join ';')
                }
            }
            $ExportArray += $customObject
        } 
        # If the attributes are an array, add them to the export array
        elseif ($attributes -is [array] -or $attributes -is [PSCustomObject] ) {
            #Write-Verbose "Attributes are an Array"
            $ExportArray += $attributes
        }
        # If the attributes are a string, add them to the export array
        else {
            #Write-Verbose "Attributes are a string"
            $ExportArray += $attributes
        }
    }
    Write-ProgressHelper -Activity "Converting Hash Table" -Completed
    Return $ExportArray
}

#Combine temporary csv files created from hash key from temp folder
function Combine-TempCSVFiles {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [Hashtable]$HashToConvert,
        [Parameter(Mandatory=$False)]
        [String]$tenant,
        [Parameter(Mandatory=$False)]
        [String]$table
    )
    #Combine files with global tnenant value from temp folder
    Write-Progress -Activity "Combine All CSV Files into One Excel" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
    $ExportedCSVFiles = Get-ChildItem -Path $env:TEMP -Filter "*$global:tenant*.csv"

    if (-not $ExportedCSVFiles) {
        $ExportedCSVFiles = Get-ChildItem -Path $env:TEMP -Filter "*ArrayaDiscoveryReport*.csv"
    }
    
    # Check again if no files are found
    if (-not $ExportedCSVFiles) {
        Write-Warning "No CSV files found in $env:TEMP matching the provided patterns."
    }
    
    
    $totalCount = ($ExportedCSVFiles | Measure).count
    $progresscounter = 0
    $start = Get-Date
    foreach ($file in $ExportedCSVFiles) {
        $progresscounter++
        $worksheetName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

        # Check if the filename contains $global:tenant and replace if it exists
        if ($file.Name -like "*$global:tenant*") {
            $worksheetName = $worksheetName.Replace("_$global:tenant","")
        }
        # Check if the filename contains "_ArrayaDiscoveryReport" and replace if it exists
        elseif ($file.Name -like "*_ArrayaDiscoveryReport*") {
            $worksheetName = $worksheetName.Replace("_ArrayaDiscoveryReport","")
        }
        
        Write-ProgressHelper -Activity "Adding Worksheets to Excel File" -CurrentOperation "Adding Worksheet $($worksheetName) to $($ExportDetails[0])" -ProgressCounter ($progresscounter) -TotalCount $totalCount -StartTime $start
        Import-Csv -Path $file.FullName | Export-Excel -Path $ExportDetails[0] -WorksheetName $worksheetName -ClearSheet
        Write-ProgressHelper -Activity "Adding Worksheets to Excel File" -Completed

    }
    #Delete the temporary CSV files
    Get-ChildItem -Path $env:TEMP -Filter "*$global:tenant*.csv" | Remove-Item
}

#Export Hash Table to Excel; export each key in hash to csv and then combine into excel
function Export-HashTableToExcel {
    [CmdletBinding()]
    param ( [Parameter(Mandatory=$True)] [Hashtable]$hashtable,
        [Parameter(Mandatory=$True)] [Array]$ExportDetails,
        [Parameter(Mandatory=$false)] [switch]$tenant
	)
    
    # Combine all CSV files into a single Excel file
    $totalCount = ($hashtable.keys | measure).count
    $progresscounter = 0
    $start = Get-Date
    # Export each hashtable to a separate CSV file
    foreach ($table in $hashtable.Keys){
        try {
            $progresscounter++
            #Export All provided Tables
            Write-ProgressHelper -ID 2 -Activity "Exporting '$($table)' Hash To Excel" -ProgressCounter ($progresscounter) -TotalCount $totalCount -StartTime $start
            Write-Log -Type DEBUG -Message ("({0}/{1}) Converting '{2}' Hash Table to Array" -f $progressCounter, $totalCount, $table) -ExportFileLocation $ExportDetails[0]            

            if ($tenant) {
                $ExportStatsArray = Convert-HashToArray -table $table -HashToConvert $hashtable[$table] -tenant $global:tenant
            }
            else {
                $ExportStatsArray = Convert-HashToArray -table $table -HashToConvert $hashtable[$table]
            }
            
            # Use the table name to create a unique temporary file
            $tempPath = Join-Path -Path $env:TEMP -ChildPath ("{0}_{1}.csv" -f $table, "ArrayaDiscoveryReport")
            Write-Log -Type DEBUG -Message ("({0}/{1}) Export '{2}' Array to Excel in temp folder {3}" -f $progressCounter, $totalCount, $table, $tempPath) -ExportFileLocation $ExportDetails[0]            
            $ExportStatsArray | Export-Csv -Path $tempPath -NoTypeInformation -Encoding UTF8
            Write-ProgressHelper -ID 2 -Activity "Exporting $($table) Hash To Excel" -Completed
        }
        catch {
            Write-Log -Type Error -Message "An error occurred in converting Hash To Array for $($table). $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]            
            $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in converting Hash To Array for $($table). $($_.Exception.Message)"
            $global:AllDiscoveryErrors += $ErrorObject
            throw $_
        }
    }

    try {
        #Combine files with global tnenant value from temp folder
        Combine-TempCSVFiles -errorAction Stop
        Write-Host "Tenant Report located at: $($ExportDetails[0])" -ForegroundColor Green

        #Delete the temporary CSV files
        Get-ChildItem -Path $env:TEMP -Filter "*ArrayaDiscoveryReport*.csv" | Remove-Item
    }
    catch {
        Write-Log -Type Error -Message "An error occurred in Exporting the Tenant Statistics to $($ExportDetails[0]). $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]            
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in Exporting the Tenant Statistics to $($ExportDetails[0]). Please check if the location is valid or if the file is open in another application. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
        try {
            Write-Log -Type DEBUG -Message "Attempt Number 2 to Export Tenant Statistics to $($ExportDetails[0])" -ExportFileLocation $ExportDetails[0]            
            $ExportDetails = Get-ExportPath
            #Combine files with global tnenant value from temp folder
            Combine-TempCSVFiles -errorAction Stop
            Write-Host "Tenant Report located at: $($ExportDetails[0])" -ForegroundColor Green

            #Delete the temporary CSV files
            Get-ChildItem -Path $env:TEMP -Filter "*ArrayaDiscoveryReport*.csv" | Remove-Item
        }
        catch {
            Write-Log -Type Error -Message "Second Attempt to Export Tenant Statistics failed to $($ExportDetails[0]). $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]            
            $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "Second Attempt to Export Tenant Statistics failed to $($ExportDetails[0]). $($_.Exception.Message)"
            $global:AllDiscoveryErrors += $ErrorObject
            throw $_
        } 
    }  
}

#Function to get Export Path
function Get-ExportPath {
    [CmdletBinding()]
    param (
        [string]$FileNameSegment,
        [string]$DefaultExtension = ".xlsx"
    )

    # Ask user for Export location
    Write-Host "Gather Export Path and/or File Name" -ForegroundColor Cyan
    $userInput = Read-Host -Prompt "Enter the file path (with .xlsx or .csv extension) or folder path to save the file"

    # Handle quotes in input
    $userInput = $userInput -replace '"', ''

    # If user input is empty, default to Desktop
    if ([string]::IsNullOrEmpty($userInput)) {
        $userInput = [Environment]::GetFolderPath("Desktop")
    }

    # File path processing
    $folderPath = ""
    $fileName = ""

    if ((Test-Path $userInput) -and (Get-Item -Path $userInput -ErrorAction SilentlyContinue).PSIsContainer) {
        $folderPath = $userInput
    } else {
        $folderPath = Split-Path -Path $userInput -Parent
        $fileName = Split-Path -Path $userInput -Leaf
    }

    # If folderPath is empty or invalid, default to current script location
    if ([string]::IsNullOrEmpty($folderPath) -or !(Test-Path $folderPath)) {
        $folderPath = $PSScriptRoot
    }

    # Check file extension and set default if none
    $extension = [IO.Path]::GetExtension($fileName)
    if ([string]::IsNullOrEmpty($extension)) {
        $extension = $DefaultExtension
        $fileName = "$global:tenant-$FileNameSegment" + $extension
    }

    # Full path
    $fullPath = Join-Path -Path $folderPath -ChildPath $fileName

    # Get the file name without extension
    $fileNameWithoutExtension = [IO.Path]::GetFileNameWithoutExtension($fileName)

    Write-Host "The file will be saved to: $fullPath" -ForegroundColor Green
    return $fullPath, $fileNameWithoutExtension
}

# Export Errors
function Export-ErrorReports {
    [CmdletBinding()]
    param(
        [string]$BaseName = "ErrorReport",
        [string]$ExportFileLocation,
        [Parameter(Mandatory=$true)]
        [array]$ErrorData,
        [string]$ErrorReportFolderName,
        [string]$ErrorReportFolderDirectory,
        [Parameter(Mandatory=$True)]
        [string]$logReportDirectory
    )

    # Validate ErrorData
    if ($ErrorData.Count -eq 0) {
        Write-Log -Type WARNING -Message "No error data provided to export. Exiting function." -ExportFileLocation $ExportDetails[0]
        return
    }

    Write-Log -Type INFO -Message "START: Export all Errors" -ExportFileLocation $ExportDetails[0]
    # Handle quotes in input
    $ExportFileLocation = $ExportFileLocation -replace '"', ''

    if ($ExportFileLocation) {
        $directory = [System.IO.Path]::GetDirectoryName($ExportFileLocation)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ExportFileLocation)
        $errorReportFolderName = $baseName + " Error Reporting"
        $errorReportFolderDirectory = Join-Path -Path $directory -ChildPath $errorReportFolderName
    } else {
        if ($ErrorReportFolderDirectory) {
            $directory = $ErrorReportFolderDirectory
        } else {
            $directory = $env:TEMP
        }
        
        if ($ErrorReportFolderName) {
            $errorReportFolderName = $ErrorReportFolderName
        } else {
            $errorReportFolderName = "$BaseName Error Reporting"
        }
        $errorReportFolderDirectory = Join-Path -Path $directory -ChildPath $errorReportFolderName
    }

    Write-Log -Type INFO -Message "INFO: Exporting Error Logs to directory $($errorReportFolderDirectory)" -ExportFileLocation $ExportDetails[0]

    try {
        if (-not (Test-Path $errorReportFolderDirectory)) {
            $result = New-Item -Path $errorReportFolderDirectory -ItemType Directory
            Write-Log -Type INFO -Message "INFO: Error Report Directory '$($errorReportFolderDirectory)' does not exist. Created Folder Directory" -ExportFileLocation $ExportDetails[0]
        }

        $newBaseName = "$BaseName-ErrorLog"
        
        $paths = @{
            'json' = Join-Path -Path $errorReportFolderDirectory -ChildPath "$newBaseName.json"
            'txt'  = Join-Path -Path $errorReportFolderDirectory -ChildPath "$newBaseName.log"
            'csv'  = Join-Path -Path $errorReportFolderDirectory -ChildPath "$newBaseName.csv"
        }

        $ErrorData | ConvertTo-Json -Depth 1 | Set-Content -Path $paths['json']
        $ErrorData | Out-File $paths['txt']
        $ErrorData | Export-Csv -Path $paths['csv'] -NoTypeInformation -Encoding UTF8

        $paths.GetEnumerator() | ForEach-Object {
            Write-Log -Type INFO -Message "INFO: Exported $($_.Key) Error Logs to directory $($_.Value)" -ExportFileLocation $ExportDetails[0]
        }

    } catch {
        Write-Error "Failed to export error reports: $_"
    }
}

# ----------------------------------
# Exchange Specific Functions
# ----------------------------------

#Convert Names to EmailAddresses loop
function ConvertTo-EmailAddressesLoop {
    param (
        [Parameter(Mandatory=$true,HelpMessage='InputArray to Convert EmailAddresses')] [array] $InputArray,
        [Parameter(Mandatory=$true,HelpMessage='InputArray to Convert EmailAddresses')]
        [ValidateSet('ExchangeOnline', 'EXO', 'On-Premises', 'OnPremises')]
        [string] $ExchangeEnvironment
    )
    $OutPutArray = @()
    foreach ($recipientObject in $InputArray) {
        #Check Address is Mail Enabled; If OnPremises and If Office365
        try {
            switch ($ExchangeEnvironment) {
                {$_ -in 'ExchangeOnline', 'EXO'} {
                        $recipientCheck = Get-EXORecipient $recipientObject.ToString() -ErrorAction SilentlyContinue
                        $tempUser = $recipientCheck.PrimarySMTPAddress.ToString()
                 }
                {$_ -in 'On-Premises', 'OnPremises'} { 
                        $recipientCheck = Get-Recipient $recipientObject.DistinguishedName.ToString() -ErrorAction SilentlyContinue
                        $tempUser = $recipientCheck.PrimarySMTPAddress.ToString()
                }
            }
        }
        catch {
            if ($_.Exception.Message -like "You cannot call a method on a null-valued expression") {
                Write-Verbose "Can't Find Exchange Recipient: $recipientObject 1" -ForegroundColor Red
                $tempUser = $recipientObject.Name.ToString()
            }
            else {
                Write-Verbose "Can't Find Exchange Recipient: $recipientObject 2" -ForegroundColor Red
                $tempUser = $recipientObject
            }
        }
        Write-Verbose "Found: $tempUser" -ForegroundColor Green
        $OutPutArray += $tempUser
    }
    return $OutPutArray
}
#Gather all Mailboxes, Group Mailboxes, Unified Groups, and Public Folders
function Get-AllExchangeMailboxDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )

    try {
        try {
            # Gather Mailboxes - Include InActive Mailboxes
            #Write-Progress -Activity "Getting all mailboxes with $($detailLevel) details" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
            $start = Get-Date
            $tenantStatsHash["AllMailboxes"] = @{}
            Write-Host "Getting all mailboxes and inactive mailboxes with $($detailLevel) details ..." -ForegroundColor Cyan -nonewline
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] START: Getting all mailboxes with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]
    
            # Fetch all mailboxes including inactive ones
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Gathering all mailboxes newer EXO Module (Get-EXOMailbox) including Inactive Mailboxes" -ExportFileLocation $ExportDetails[0]
            $exoMailboxes = Get-EXOMailbox -ResultSize unlimited -Filter "PrimarySMTPAddress -notlike '*DiscoverySearchMailbox*'" -IncludeInactiveMailbox -PropertySets All -ErrorAction SilentlyContinue 
            
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Combine Group Mailboxes and all Mailboxes into single hash table. Combine into single hash, then into a combined array" -ExportFileLocation $ExportDetails[0]
            # Create an empty hashtable
            $allMailboxesHash = @{}
            # Insert individual mailboxes into the hashtable
            foreach ($mailbox in $exoMailboxes) {
                # Assuming primary SMTP address is unique and can be used as the hashtable key
                $allMailboxesHash[$mailbox.PrimarySmtpAddress] = $mailbox
            }
        }
        catch {
            $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in Gathering Mailbox Details. $($_.Exception.Message)"
            $global:AllDiscoveryErrors += $ErrorObject
            Write-Log -Type ERROR -Message "[Get-AllExchangeMailboxDetails] An error occurred in Gathering Mailbox Details. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
        }

        # Fetch all Group mailboxes including inactive ones
        try {
            # Fetch all group mailboxes including inactive ones
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Gathering all group mailboxes using older EXO Module (Get-Mailbox) including Inactive Mailboxes" -ExportFileLocation $ExportDetails[0]
            $groupMailboxes = Get-Mailbox -GroupMailbox -IncludeInactiveMailbox

            # Insert group mailboxes into the hashtable
            foreach ($mailbox in $groupMailboxes) {
                # Assuming primary SMTP address is unique and can be used as the hashtable key
                $key = $mailbox.PrimarySmtpAddress
                $allMailboxesHash[$key] = $mailbox
            }
        }
        catch {
            if ($_.Exception.Message -like "*A parameter cannot be found that matches parameter name 'GroupMailbox'*") {
                Write-Warning "Unable to Gather Group Mailboxes. The Get-Mailbox cmdlet does not support the GroupMailbox parameter. Skipping..."
            }
            else {
                $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in Gathering Group Mailbox Details. $($_.Exception.Message)"
                $global:AllDiscoveryErrors += $ErrorObject
                Write-Log -Type ERROR -Message "[Get-AllExchangeMailboxDetails] An error occurred in Gathering Group Mailbox Details. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
            }
        }

        #Convert to Mailbox Array
        try {
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Converting all mailboxes in AllMailboxesHash into Array allMailboxesArray" -ExportFileLocation $ExportDetails[0]
            $allMailboxesArray = Convert-HashToArray -HashToConvert $allMailboxesHash -verbose
        }
        catch {
            $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in converting Mailboxes to Array from current Hash. $($_.Exception.Message)"
            $global:AllDiscoveryErrors += $ErrorObject
            Write-Log -Type ERROR -Message "[Get-AllExchangeMailboxDetails] An error occurred in converting Mailboxes to Array from current Hash. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
        }
        

        # Separate group mailbox retrieval logic
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Filtering mailbox details with $($detailLevel) detail level" -ExportFileLocation $ExportDetails[0]
        switch ($detailLevel) {
            {$_ -in "minimum", "combined", "all"} { 
                $DesiredProperties = @(
                    "DisplayName", "Office", "UserPrincipalName", "RecipientTypeDetails", "PrimarySmtpAddress"
                    "WhenMailboxCreated", "UsageLocation", "IsInactiveMailbox", "WasInactiveMailbox", "WhenSoftDeleted"
                    "InPlaceHolds", "AccountDisabled", "IsDirSynced", "HiddenFromAddressListsEnabled", "Alias"
                    "EmailAddresses", "GrantSendOnBehalfTo", "AcceptMessagesOnlyFrom", "AcceptMessagesOnlyFromDLMembers", "AcceptMessagesOnlyFromSendersOrMembers"
                    "RejectMessagesFrom", "RejectMessagesFromDLMembers", "RejectMessagesFromSendersOrMembers", "RequireSenderAuthenticationEnabled", "WindowsEmailAddress"
                    "DistinguishedName", "Identity", "WhenChanged", "WhenCreated", "ExchangeObjectId"
                    "Guid", "DeliverToMailboxAndForward", "ForwardingAddress", "ForwardingSmtpAddress", "LitigationHoldEnabled"
                    "RetentionHoldEnabled", "DelayHoldApplied", "RetentionPolicy", "ExchangeGuid", "IsResource"
                    "IsShared", "ResourceType", "RoomMailboxAccountEnabled", "WindowsLiveID", "MicrosoftOnlineServicesID"
                    "EffectivePublicFolderMailbox", "MailboxPlan", "ArchiveStatus", "ArchiveState", "ArchiveName"
                    "ArchiveGuid", "AutoExpandingArchiveEnabled", "DisabledArchiveGuid", "PersistedCapabilities"
                )
                $allMailboxesArray = $allMailboxesArray | select $DesiredProperties
            }
        }
    }
    catch {
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in Gathering Mailbox Details. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
        Write-Log -Type ERROR -Message "[Get-AllExchangeMailboxDetails] An error occurred in Gathering Mailbox Details. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
    }
    finally {
        # Insert mailboxes into the Tenant Stats hashtable
        foreach ($mailbox in $allMailboxesArray) {
            # Assuming primary SMTP address is unique and can be used as the hashtable key
            $tenantStatsHash["AllMailboxes"][$mailbox.PrimarySmtpAddress] = $mailbox
        }
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Adding Mailboxes Data to Hash" -ExportFileLocation $ExportDetails[0]

        #Split up and create separate hash tables for each mailbox type
        switch ($detailLevel) {
            {$_ -in "geek","all"} { 
                #Add User Mailboxes to Tenant Stats Hash
                if ($allUserMailboxes = $allMailboxesArray | ?{$_.RecipientTypeDetails -eq "UserMailbox"}) {
                    #Write-Progress -Activity "Adding User Mailbox Data to Hash" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
                    $tenantStatsHash["UserMailboxes"] = @{}
                    foreach ($user in $allUserMailboxes) {
                        $key = $user.PrimarySMTPAddress.ToString()
                        $value = $user
                        $tenantStatsHash["UserMailboxes"][$key] = $value
                    }
                    Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Adding User Mailboxes Data to Hash" -ExportFileLocation $ExportDetails[0]
                }

                #Add User Mailboxes to Tenant Stats Hash
                if ($allinActiveMailboxes = $allMailboxesArray | ?{$_.IsInactiveMailbox -eq $true}) {
                    #Write-Progress -Activity "Adding User Mailbox Data to Hash" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
                    $tenantStatsHash["InActiveMailboxes"] = @{}
                    foreach ($inactiveMBX in $allinActiveMailboxes) {
                        $key = $inactiveMBX.PrimarySMTPAddress.ToString()
                        $value = $inactiveMBX
                        $tenantStatsHash["InActiveMailboxes"][$key] = $value
                    }
                    Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Adding Inactive Mailboxes Data to Hash" -ExportFileLocation $ExportDetails[0]
                }
                
                #Add User Mailboxes to Tenant Stats Hash
                if ($allNonUserMailboxes = $allMailboxesArray | ?{$_.RecipientTypeDetails -ne "UserMailbox" -and $_.RecipientTypeDetails -ne "GroupMailbox"}) {
                    #Write-Progress -Activity "Adding Inactive User Data to Hash" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
                    $tenantStatsHash["NonUserMailboxes"] = @{}
                    foreach ($nonUser in $allNonUserMailboxes) {
                        $key = $nonUser.PrimarySMTPAddress.ToString()
                        $value = $nonUser
                        $tenantStatsHash["NonUserMailboxes"][$key] = $value
                    }
                    Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Adding Non User Mailboxes Data to Hash" -ExportFileLocation $ExportDetails[0]
                }
                
                #Add Group Mailboxes to Tenant Stats Hash
                if ($allGroupMailboxes = $allMailboxesArray | ?{$_.RecipientTypeDetails -eq "GroupMailbox"}) {
                    #Write-Progress -Activity "Adding Group Mailbox Data to Hash" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
                    $tenantStatsHash["GroupMailboxes"] = @{}
                    $allGroupMailboxes = $allMailboxesArray | ?{$_.RecipientTypeDetails -eq "GroupMailbox"}
                    foreach ($groupMailbox in $allGroupMailboxes) {
                        $key = $groupMailbox.PrimarySMTPAddress.ToString()
                        $value = $groupMailbox
                        $tenantStatsHash["GroupMailboxes"][$key] = $value
                    }
                    Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Adding Group Mailboxes Data to Hash" -ExportFileLocation $ExportDetails[0]
                }

                #Add Archive Mailboxes to Tenant Stats Hash
                if ($allarchiveMailboxes = $allMailboxesArray | ? {$_.ArchiveStatus -ne "None"}) {
                    #Write-Progress -Activity "Adding Archive Mailbox Data to Hash" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
                    $tenantStatsHash["ArchiveMailboxes"] = @{}
                    foreach ($archiveMailbox in $allarchiveMailboxes) {
                        $key = $archiveMailbox.PrimarySMTPAddress.ToString()
                        $value = $archiveMailbox
                        $tenantStatsHash["ArchiveMailboxes"][$key] = $value
                    }
                    Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Adding Group Mailboxes Data to Hash" -ExportFileLocation $ExportDetails[0]
                }
            }
        }
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] COMPLETED: Gathering All Mailbox Details in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
    }


    #Mailbox Statistics to Hash Table
    ###########################################################################################################################################
    ## Primary Mailbox Stats
    try {
        $start = Get-Date
        $tenantStatsHash["PrimaryMailboxStats"] = @{}
        $mailboxStatsHash = @{}

        Write-Host "Getting primary mailbox stats..." -ForegroundColor Cyan -nonewline
        #Write-Progress -Activity "Adding All primary mailbox (including Groups) stats" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
        
        $start = Get-Date
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Gathering All Primary Mailbox Statistics including Group Mailboxes and Inactive Mailboxes" -ExportFileLocation $ExportDetails[0]
        $primaryMailboxStats = $allMailboxesArray | Get-EXOMailboxStatistics -IncludeSoftDeletedRecipients -ErrorAction SilentlyContinue 

        #Add to Tenant Stats Hash
        $primaryMailboxStats | ForEach-Object {
            $key = $_.MailboxGuid.ToString()
            $value = $_
            $mailboxStatsHash[$key] = $value
        }
        $tenantStatsHash["PrimaryMailboxStats"] = $mailboxStatsHash
    }
    catch {
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in Gathering Mailbox Satistics and adding to Hash Table. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
        Write-Log -Type ERROR -Message "[Get-AllExchangeMailboxDetails] An error occurred in Gathering Mailbox Satistics and adding to Hash Table. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
    }
    finally {
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] COMPLETED: Gathering All Primary Mailbox Statistics in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
    }
    
    ## Archive Mailbox Stats to Hash Table
    try {
        $start = Get-Date
        $tenantStatsHash["ArchiveMailboxStats"] = @{}
        $archiveMailboxStatsHash = @{}

        Write-Host "Getting archive mailbox stats..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Gathering All Archive Mailbox Statistics. Including Group and Inactive Mailboxes" -ExportFileLocation $ExportDetails[0]
        #Write-Progress -Activity "Getting All archive mailbox stats" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
        $archiveMailboxStats = $allMailboxesArray | ? {$_.ArchiveStatus -ne "None"} | Get-EXOMailboxStatistics -Archive -ErrorAction SilentlyContinue -IncludeSoftDeletedRecipients

        #Add to Tenant Stats Hash
        $archiveMailboxStats | ForEach-Object {
            #errors if key is null
            if($key = $_.MailboxGuid.ToString()) {
                $value = $_
                $archiveMailboxStatsHash[$key] = $value
            }
        }
        $tenantStatsHash["ArchiveMailboxStats"] = $archiveMailboxStatsHash
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Adding Archive Mailbox Statistics to Tenant Stats Hash." -ExportFileLocation $ExportDetails[0]
    }
    catch {
        if ($_.Exception.Message -like "You cannot call a method on a null-valued expression") {
            $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in Gathering Archive Mailbox Satistics and adding to Hash Table. There are no Archive mailboxes found. $($_.Exception.Message)"
            $global:AllDiscoveryErrors += $ErrorObject
            Write-Log -Type ERROR -Message "[Get-AllExchangeMailboxDetails] An error occurred in Gathering Archive Mailbox Satistics and adding to Hash Table. There are no Archive mailboxes found. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
        }
        else {
            $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in Gathering Archive Mailbox Satistics and adding to Hash Table. $($_.Exception.Message)"
            $global:AllDiscoveryErrors += $ErrorObject
            Write-Log -Type ERROR -Message "[Get-AllExchangeMailboxDetails] An error occurred in Gathering Archive Mailbox Satistics and adding to Hash Table. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
        }
    }
    finally { 
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails]COMPLETED: Gathering All Archive Mailbox Statistics in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
    }
}
function Get-MailFlowRulesandConnectors {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )
    # Create Hash Tables for Mail Flow Rules and Connectors
    $tenantStatsHash["MailFlowRules"] = @{}
    $tenantStatsHash["MailFlowConnectors"] = @{}

    $start = Get-Date
    ### Get Mail Flow Rules ### - START
    try {
        Write-Host "Getting all Mail Flow Rules and Connectors ..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] START: Gathering all Mail Flow Rules and Connectors with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]

        Write-Progress -Activity "Getting all Mail Flow Rules details" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
        Write-Log -Type INFO -Message "Gathering all Mail Flow Rules" -ExportFileLocation $ExportDetails[0]
        switch ($detailLevel) {
            {$_ -in "minimum", "combined", "all"} { 
                $DesiredProperties = @(
                    "Name", "State", "Mode", "Priority", "Description"
                )
                try { $mailFlowRules = Get-TransportRule -IncludeTestModeConnectors -ErrorAction Continue | Select $DesiredProperties }
                catch {  $mailFlowRules = Get-TransportRule -ErrorAction Continue | Select $DesiredProperties  }
            }
            geek {
                try {  $mailFlowRules = Get-TransportRule -IncludeTestModeConnectors -ErrorAction Continue  }
                catch {  $mailFlowRules = Get-TransportRule -ErrorAction Continue  }
            }
        }

        #convert Mail Flow Rules to Hash Table
        $progresscounter = 0
        $totalCount = ($mailFlowRules | measure).count
        foreach ($rule in $mailFlowRules) {
            $progresscounter++
            Write-Log -Type DEBUG -Message "[Get-MailFlowRulesandConnectors] Gathering Mail Flow Details for $($rule.Name): $($progresscounter)/$($totalCount)" -ExportFileLocation $ExportDetails[0]
            $tenantStatsHash["MailFlowRules"][$rule.Priority] = $rule
        }
        Write-Progress -Activity "Getting all Mail Flow Rules details" -Completed
    }
    catch {
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in gathering MailFlow Rules. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
        Write-Log -Type Error -Message "[Get-MailFlowRulesandConnectors] An error occurred in gathering MailFlow Rules. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
    }
    ### Get Mail Flow Rules ### - END
    
    ### Get Mail Flow Connectors  ### - START
    function Add-ConnectorToHash {
        param (
            [Parameter(Mandatory=$true)]
            $connectorList,
            [Parameter(Mandatory=$true)]
            [ValidateSet('Inbound', 'Outbound')]
            $direction
        )
    
        $progressCounter = 0
        $totalCount = ($connectorList | measure).Count
        foreach ($connector in $connectorList) {
            $progressCounter++
            Write-Log -Type DEBUG -Message ("[Add-ConnectorToHash] ({0}/{1}) Gathering '{2}' Mail Connector Details for {3}" -f $progressCounter, $totalCount, $direction, $connector.ID) -ExportFileLocation $ExportDetails[0]

            $currentConnector = $connector | Select-Object *
            $currentConnector | Add-Member -MemberType NoteProperty -Name "ConnectorDirection" -Value $direction -Force
    
            if ($detailLevel -ne "geek") {
                $propertiesToAdd = @("RecipientDomains", "SmartHosts", "ValidationRecipients", "SenderDomains", "SenderIPAddresses", "TrustedOrganizations", "EFSkipIPs", "EFSkipMailGateway", "EFUsers")
                foreach ($prop in $propertiesToAdd) {
                    if ($connector.$prop) {
                        $currentConnector | Add-Member -MemberType NoteProperty -Name $prop -Value ($connector.$prop -join ",") -Force
                    }
                }
            }
        
    
            $tenantStatsHash["MailFlowConnectors"][$connector.Id] = $currentConnector
        }
    }

    try {
        # Gather Connector Details - Inbound and Outbound Connector
        Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] Gathering all Inbound Mail Connectors" -ExportFileLocation $ExportDetails[0]
        $mailFlowInboundConnectors = Get-InboundConnector -ErrorAction Stop
        #$mailFlowInboundConnectors | foreach { $_ | Add-Member -MemberType NoteProperty -Name "ConnectorDirection" -Value "Inbound" -Force }
        Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] Found $(($mailFlowInboundConnectors| measure).count) Inbound Mail Connectors" -ExportFileLocation $ExportDetails[0]

        Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] Gathering all Outbound Mail Connectors" -ExportFileLocation $ExportDetails[0]
        $mailFlowOutboundConnectors = Get-OutboundConnector -IncludeTestModeConnectors $true -ErrorAction Stop
        #$mailFlowOutboundConnectors | foreach { $_ | Add-Member -MemberType NoteProperty -Name "ConnectorDirection" -Value "Outbound" -Force }
        Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] Found $(($mailFlowOutboundConnectors | measure).count) Outbound Mail Connectors" -ExportFileLocation $ExportDetails[0]

        # Filter properties if detail level is minimum, combined, or all
        if ($detailLevel -in @("minimum", "combined", "all")) {
            $DesiredProperties = @(
            "Id", "ConnectorDirection", "Comment", "Enabled", "TestMode", "ConnectorType", 
            "UseMXRecord", "IsTransportRuleScoped", "RecipientDomains", 
            "SmartHosts", "AllAcceptedDomains", "SenderRewritingEnabled",
            "RouteAllMessagesViaOnPremises", "CloudServicesMailEnabled", 
            "ValidationRecipients", "Description", "IsValidated",
            "LastValidationTimestamp", "SenderDomains", "SenderIPAddresses", 
            "TrustedOrganizations", "RequireTls", "TlsSettings", "TlsDomain", 
            "TreatMessagesAsInternal", "EFTestMode", "EFSkipLastIP",
            "EFSkipIPs", "EFSkipMailGateway", "EFUsers"
            )
            $mailFlowInboundConnectors = $mailFlowInboundConnectors | Select-Object $DesiredProperties
            $mailFlowOutboundConnectors = $mailFlowOutboundConnectors | Select-Object $DesiredProperties
        }

        # Ensure MailFlowConnectors is initialized as a hashtable
        $tenantStatsHash["MailFlowConnectors"] = @{}

        #convert Mail Flow Connectors to Hash Table
        if ($mailFlowInboundConnectors) {
            Add-ConnectorToHash -connectorList $mailFlowInboundConnectors -direction "Inbound"
        }
        if ($mailFlowOutboundConnectors) {
            Add-ConnectorToHash -connectorList $mailFlowOutboundConnectors -direction "Outbound"
        }
        Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] Add Mail Connectors Details to Tenant Stats Hash" -ExportFileLocation $ExportDetails[0]
    }
    catch {
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in MailFlow Connectors. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
        Write-Log -Type ERROR -Message "[Get-MailFlowRulesandConnectors] An error occurred in MailFlow Connectors. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
    }
    finally {
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] COMPLETED: Gathering All Mail Flow Rules and Connectors in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
    }
}
function Get-AllRecipientDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )
    try {
        $start = Get-Date
        $tenantStatsHash["AllRecipients"] = @{}
        Write-Host "Getting all Exchange Online Recipients $($detailLevel) details ..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-AllRecipientDetails] START: Gathering all Exchange Online Recipients $($detailLevel) details" -ExportFileLocation $ExportDetails[0]

        switch ($detailLevel) {
            {$_ -in "minimum", "combined", "all"} { 
                $DesiredProperties = @(
                    "DisplayName", "RecipientTypeDetails", "PrimarySMTPAddress"
                    "EmailAddresses", "HiddenFromAddressListsEnabled", "AddressBookPolicy"
                    "SKUAssigned", "WhenCreated", "WhenSoftDeleted"
                )
                $allRecipients = Get-EXORecipient -Properties $DesiredProperties -Filter "PrimarySMTPAddress -notlike '*DiscoverySearchMailbox*'" -ResultSize Unlimited -ErrorAction Stop
            }
            geek { $allRecipients = Get-EXORecipient -PropertySets All -Filter "PrimarySMTPAddress -notlike '*DiscoverySearchMailbox*'" -ResultSize Unlimited -ErrorAction Stop }
        }
        Write-Log -Type INFO -Message "[Get-AllRecipientDetails] FOUND $($allRecipients.count) Exchange Online Recipients $($detailLevel) details" -ExportFileLocation $ExportDetails[0]
        #Add to hash table
        Write-Log -Type INFO -Message "[Get-AllRecipientDetails] Adding Exchange Online Recipients to Tenant Stats Hash" -ExportFileLocation $ExportDetails[0]
        foreach ($recipient in $allRecipients) {
            $recipient.EmailAddresses = ($recipient.EmailAddresses -join ",")
            $tenantStatsHash["AllRecipients"][$recipient.DisplayName] = $recipient
        }
    }
    catch {
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-AllRecipientDetails function. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
        Write-Log -Type ERROR -Message "[Get-AllRecipientDetails] An error occurred in running Get-AllRecipientDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
    }
    finally {
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-AllRecipientDetails] COMPLETED: Gathering all Exchange Online Recipients" -ExportFileLocation $ExportDetails[0]
    }
}
# Exchange Group Details
function Get-ExchangeGroupDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )

    try {
        $start = Get-Date
        $tenantStatsHash["AllExchangeGroups"] = @{}
        
        #Write-Host "Gathering Exchange Online Objects and data" -ForegroundColor Black -BackgroundColor Yellow
        Write-Host "Getting all Exchange Online Groups ..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-ExchangeGroupDetails] START: Gathering all Exchange Online Groups with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]

        # Consolidated Get-Recipient calls
        $types = 'group', 'MailNonUniversalGroup', 'MailUniversalSecurityGroup', 'MailUniversalDistributionGroup', 'DynamicDistributionGroup'
        Write-Log -Type INFO -Message "[Get-ExchangeGroupDetails] Gathering all Exchange Online Groups for $($types)" -ExportFileLocation $ExportDetails[0]

        switch ($detailLevel) {
            #minimum {$mgUsers = Get-MgUser -all -ErrorAction Stop | select }
            {$_ -in "minimum", "combined", "all"} { 
                $DesiredProperties = @(
                    "DisplayName", "identity", "PrimarySMTPAddress"
                    "RecipientTypeDetails", "ManagedBy", "Name"
                    "alias", "Notes", "HiddenFromAddressListsEnabled"
                )
                $allMailGroups = $types | ForEach-Object { 
                    Get-EXORecipient -Properties $DesiredProperties -RecipientTypeDetails $_ -ResultSize unlimited -EA SilentlyContinue 
                } | Sort-Object DisplayName
            }
            ## Geek mode pull all details as get-exorecipient but not for the actual group details pulled later. need to review that portion.
            geek {
                $allMailGroups = $types | ForEach-Object { 
                    Get-EXORecipient -PropertySets All -RecipientTypeDetails $_ -ResultSize unlimited -EA SilentlyContinue 
                } | Sort-Object DisplayName
            }
        }

        Write-Log -Type INFO -Message "[Get-ExchangeGroupDetails] Gathering all Exchange Online Groups Details" -ExportFileLocation $ExportDetails[0]
        $progresscounter = 0
        $totalCount = $allMailGroups.count
        foreach ($object in $allMailGroups) {
            $progresscounter++
            $identity = $object.identity.tostring()
            $PrimarySMTPAddress = $object.PrimarySMTPAddress.ToString()
            Write-ProgressHelper -Activity "Gathering All Exchange Online Group Details" -CurrentOperation "Gathering Group Details for $($PrimarySMTPAddress)" -ProgressCounter ($progresscounter) -TotalCount $totalCount -StartTime $start
            Write-Log -Type DEBUG -Message ("[Get-ExchangeGroupDetails] ({0}/{1}) Gathering '{2}' '{3}' Group Details" -f $progressCounter, $totalCount, $object.RecipientTypeDetails, $PrimarySMTPAddress) -ExportFileLocation $ExportDetails[0]

            # Clear details
            $attributesToClear = @('groupDetails', 'groupOwners','groupMembers', 'EmailAddresses')
            foreach ($attribute in $attributesToClear) {
                Set-Variable -Name $attribute -Value @()
            }

            ## Get Email Addresses
            $EmailAddresses = $object | select -expandProperty EmailAddresses
            
            # Conditional logic for different recipient types
            switch ($object.RecipientTypeDetails) {
                "DynamicDistributionGroup" {
                    $groupDetails = Get-DynamicDistributionGroup $PrimarySMTPAddress -ErrorAction SilentlyContinue
                    $groupMembers = Get-DynamicDistributionGroupMember $PrimarySMTPAddress -ErrorAction SilentlyContinue -ResultSize unlimited -warningaction silentlycontinue
                }
                {$_ -in 'MailUniversalDistributionGroup', 'MailUniversalSecurityGroup', "MailNonUniversalGroup"} {
                    $groupDetails = Get-DistributionGroup $PrimarySMTPAddress -ErrorAction SilentlyContinue
                    $groupMembers = Get-DistributionGroupMember $PrimarySMTPAddress -ResultSize unlimited -ErrorAction SilentlyContinue
                }
                "GroupMailbox" {
                    $groupDetails = Get-UnifiedGroup $PrimarySMTPAddress -ErrorAction SilentlyContinue
                    $groupMembers = Get-UnifiedGroupLinks -Identity $object -LinkType Member -ResultSize unlimited -ErrorAction SilentlyContinue
                }
            }

            #Check Group Owners Size and Get Owners Addresses
            if ($object.ManagedBy.count -ge 1) {
                $groupOwners = $object.ManagedBy
                Write-Log -Type DEBUG -Message ("[Get-ExchangeGroupDetails] ({0}/{1}) '{2}' Owners Found for '{3}' '{4}'" -f $progressCounter, $totalCount, $object.ManagedBy.count, $object.RecipientTypeDetails, $PrimarySMTPAddress) -ExportFileLocation $ExportDetails[0]

            }
            #Check Group Members Size and Get Group Addresses
            if ($groupMembers.count -ge 1) {
                Write-Log -Type DEBUG -Message ("[Get-ExchangeGroupDetails] ({0}/{1}) '{2}' Members Found for '{3}' '{4}'" -f $progressCounter, $totalCount, $groupMembers.count, $object.RecipientTypeDetails, $PrimarySMTPAddress) -ExportFileLocation $ExportDetails[0]
            }
            Write-Log -Type DEBUG -Message ("[Get-ExchangeGroupDetails] ({0}/{1}) Create Group Output Details for '{2}' '{3}'" -f $progressCounter, $totalCount, $object.RecipientTypeDetails, $PrimarySMTPAddress) -ExportFileLocation $ExportDetails[0]

            #Output Group Details
            $currentobject = [ordered]@{
                DisplayName                              = $object.DisplayName
                Identity                                 = $identity
                Name                                     = $object.Name
                Alias                                    = $object.alias
                Notes                                    = $object.Notes
                IsDirSynced                              = $object.IsDirSynced
                HiddenFromAddressListsEnabled            = $object.HiddenFromAddressListsEnabled
                PrimarySMTPAddress                       = $object.PrimarySMTPAddress
                RecipientTypeDetails                     = $object.RecipientTypeDetails
                ResourceProvisioningOptions              = ($groupDetails.ResourceProvisioningOptions -join ",")
                IsMailboxConfigured                      = $groupDetails.IsMailboxConfigured
                EmailAddresses                           = ($EmailAddresses -join ",")
                OwnersCount                              = ($groupOwners | measure-object).count
                MembersCount                             = ($groupMembers | measure-object).count
                HiddenGroupMembershipEnabled             = ($groupDetails.HiddenGroupMembershipEnabled -join ",")
                ModeratedBy                              = ($ModeratedByRecipients -join ",")
                AccessType                               = $groupDetails.AccessType
                AllowAddGuests                           = $groupDetails.AllowAddGuests
                SharePointSiteUrl                        = $groupDetails.SharePointSiteUrl
            }

            Write-Log -Type DEBUG -Message ("[Get-ExchangeGroupDetails] ({0}/{1}) Add '{2}' '{3}' Group Details to Tenant Stats Hash" -f $progressCounter, $totalCount, $object.RecipientTypeDetails, $PrimarySMTPAddress) -ExportFileLocation $ExportDetails[0]
            Write-ProgressHelper -Activity "Gathering All Exchange Online Group Details" -CurrentOperation "Gathering Group Details for $($PrimarySMTPAddress)" -Completed
            $tenantStatsHash["AllExchangeGroups"][$object.PrimarySMTPAddress] = $currentobject
        }
    }
    catch {
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-ExchangeGroupDetails function. Exception: $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
        Write-Log -Type ERROR -Message "[Get-ExchangeGroupDetails] An error occurred in running Get-ExchangeGroupDetails function. Exception: $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
    }
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-ExchangeGroupDetails] COMPLETED: Gathering all Exchange Online Groups in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
}

# ----------------------------------
# Collaboration Specific Functions
# ----------------------------------
function Get-AllUnifiedGroups {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )
    #Get Office 365 Group / Group Mailbox data with SharePoint URL data
    try {
        $start = Get-Date
        $tenantStatsHash["UnifiedGroups"] = @{}
        Write-Host "Getting all unified groups (including soft deleted)..." -ForegroundColor Cyan -nonewline
        #Write-Progress -Activity "Getting unified groups" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
        Write-Log -Type INFO -Message "[Get-AllUnifiedGroups] START: Gathering all Unified with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]
        switch ($detailLevel) {
            {$_ -in "minimum", "combined", "all"} { 
                $DesiredProperties = @(
                    "PrimarySmtpAddress", "DisplayName", "AccessType"
                    "ExchangeGuid", "ManagedByDetails", "Notes"
                    "SharePointSiteUrl", "ContentMailboxName", "GroupMemberCount"
                    "AllowAddGuests", "WhenSoftDeleted", "HiddenFromExchangeClientsEnabled"
                    "EmailAddresses", "ModeratedBy", "FolderPath"
                    "Description", "RecipientTypeDetails", "WhenCreated"
                )
                $allUnifiedGroups = Get-UnifiedGroup -resultSize unlimited -IncludeSoftDeletedGroups -ErrorAction SilentlyContinue| Select $DesiredProperties
            }
            geek {$allUnifiedGroups = Get-UnifiedGroup -resultSize unlimited -IncludeSoftDeletedGroups -ErrorAction SilentlyContinue}
        }
        
        Write-Progress -Activity "Adding Unified Group data to Hash" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
        Write-Log -Type INFO -Message "[Get-AllUnifiedGroups] Adding Unified Group data to Hash" -ExportFileLocation $ExportDetails[0]
        foreach ($group in $allUnifiedGroups) {
            $key = $group.ExchangeGuid.ToString()
            $tenantStatsHash["UnifiedGroups"][$group.PrimarySmtpAddress] = $group
        }

        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-AllUnifiedGroups] COMPLETED: Gathering all Unified Groups in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]

    }
    catch {
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-AllUnifiedGroups function. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
        Write-Log -Type ERROR -Message "[Get-AllUnifiedGroups] An error occurred in running Get-AllUnifiedGroups function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
    }
}

#Public Folder Data; Statistics; Permissions Convert to Hash Tables
function Get-AllPublicFolderDetails {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )
    #Get Public Folder Data, Statistics, and Permissions
    $start = Get-Date
    $tenantStatsHash["PublicFolderDetails"] = @{}
    try {
        Write-Host "Getting public folders, Stats and Perms ..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] START: Gathering all public folder details with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]
        Write-Progress -Activity "Getting all public folder details" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
        switch ($detailLevel) {
            {$_ -in "minimum", "combined", "all"} { 
                $DesiredProperties = @(
                    "Identity", "Name", "MailEnabled"
                    "MailRecipientGuid", "ParentPath", "ContentMailboxName"
                    "EntryId", "FolderSize", "HasSubfolders"
                    "FolderClass", "FolderPath", "ExtendedFolderFlags"
                )
                $allPublicFolders = get-publicfolder -recurse -resultSize unlimited -ErrorAction SilentlyContinue | ?{$_.Name -NE "IPM_SUBTREE"} | Select $DesiredProperties
            }
            geek {$allPublicFolders = get-publicfolder -recurse -resultSize unlimited -ErrorAction SilentlyContinue | ?{$_.Name -NE "IPM_SUBTREE"}}
        }
        Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Found $(($allPublicFolders | measure).count) Public Folders in Exchange" -ExportFileLocation $ExportDetails[0]
    }
    catch {
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-AllPublicFolderDetails function. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
        Write-Log -Type ERROR -Message "[Get-AllPublicFolderDetails] An error occurred in running Get-AllPublicFolderDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]

    }
    
    # Public Folder Statistics
    #**************************
    Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Gathering all public folder statistics" -ExportFileLocation $ExportDetails[0]
    try {
        $PublicFolderStatistics = $allPublicFolders | get-publicfolderstatistics -ErrorAction SilentlyContinue
        $PublicFolderStatsHash = @{}
        foreach($publicFolderStat in $PublicFolderStatistics) {
            $key = $PublicFolderStat.EntryId
            $PublicFolderStatsHash[$key] = $PublicFolderStat
        }
        Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Found $($($PublicFolderStatistics | measure).count) Public Folder Statistics in Exchange" -ExportFileLocation $ExportDetails[0]
    }
    catch {
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-AllPublicFolderStatistics function. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
        Write-Log -Type ERROR -Message "An error occurred in running Get-AllPublicFolderStatistics function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
    }

    # Public Folder Permissions
    #***************************
    try {
        Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Gathering all public folder permissions" -ExportFileLocation $ExportDetails[0]
        $PublicFolderPermissions = $allPublicFolders | get-publicfolderclientpermission -ErrorAction SilentlyContinue
        #Progress Bar Parameters Reset
        $start = Get-Date
        Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Found $($PublicFolderPermissions.count) public folder permissions" -ExportFileLocation $ExportDetails[0]

        $tenantStatsHash["PublicFolderPerms"] = @{}
        
        Write-Host "Processing Public Folder Permissions..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Processing all public folder permissions" -ExportFileLocation $ExportDetails[0]
        $progresscounter = 0
        $totalCount = ($PublicFolderPermissions | measure).count
        foreach($publicFolderPermission in $PublicFolderPermissions) {
            $progresscounter++
            Write-ProgressHelper -Activity "Processing all public folder permissions" -CurrentOperation "Gathering Public Folder Permissions for $($publicFolderPermission.Identity)" -ProgressCounter ($progresscounter) -TotalCount $totalCount -StartTime $start
            Write-Log -Type DEBUG -Message "[Get-AllPublicFolderDetails] $($progresscounter)/$($totalCount): Gathering Public Folder Permissions for $($publicFolderPermission.Identity): $($progresscounter)/$($totalCount)" -ExportFileLocation $ExportDetails[0]

            
            $key = "$($publicFolderPermission.Identity)-$($publicFolderPermission.User.Displayname)"
            $permissionObject = @(
                [PSCustomObject]@{
                    FolderName = $publicFolderPermission.FolderName
                    FolderPath = $publicFolderPermission.Identity
                    Displayname = $publicFolderPermission.User.Displayname
                    PrimarySMTPAddress = $publicFolderPermission.User.RecipientPrincipal.PrimarySmtpAddress
                    AccessRights = ($publicFolderPermission.AccessRights -join ",")
                }
            )

            if($tenantStatsHash["PublicFolderPerms"].ContainsKey($key)) {
                $tenantStatsHash["PublicFolderPerms"][$key] += $permissionObject
            }
            else {
                $tenantStatsHash["PublicFolderPerms"][$key] = @($permissionObject)
            }
        }
        Write-ProgressHelper -Activity "Processing all public folder permissions" -Completed
    }
    catch {
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-AllPublicFolderPermissions function. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
        Write-Log -Type ERROR -Message "[Get-AllPublicFolderDetails] An error occurred in running Get-AllPublicFolderPermissions function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
    }   
    
    #Combine Stats with Details
    $tenantStatsHash["PublicFolderDetails"] = @{}
    $progresscounter = 0
    $totalCount = ($allPublicFolders | measure).count
    foreach($pf in $allPublicFolders) {
        $progresscounter++
        $pfStatsCheck = $PublicFolderStatsHash[$pf.EntryId]
        Write-Log -Type DEBUG -Message "Combine Public Folder Stats for $($pf.FolderPath): $($progresscounter)/$($totalCount)" -ExportFileLocation $ExportDetails[0]
        $pf | Add-Member -MemberType NoteProperty -Name "FolderPath" -Value ($pf.FolderPath -join ",") -Force
        $pf | Add-Member -MemberType NoteProperty -Name "ItemCount" -Value $pfStatsCheck.ItemCount -Force
        $pf | Add-Member -MemberType NoteProperty -Name "LastModificationTime" -Value $pfStatsCheck.LastModificationTime -Force
        $pf | Add-Member -MemberType NoteProperty -Name "OwnerCount" -Value $pfStatsCheck.OwnerCount -Force
        $pf | Add-Member -MemberType NoteProperty -Name "TotalAssociatedItemSize" -Value $pfStatsCheck.TotalAssociatedItemSize -Force
        $pf | Add-Member -MemberType NoteProperty -Name "TotalDeletedItemSize" -Value $pfStatsCheck.TotalDeletedItemSize -Force
        $pf | Add-Member -MemberType NoteProperty -Name "TotalItemSize" -Value $pfStatsCheck.TotalItemSize -Force
        $pf | Add-Member -MemberType NoteProperty -Name "MailboxOwnerId" -Value $pfStatsCheck.MailboxOwnerId -Force
        $tenantStatsHash["PublicFolderDetails"][$pf.Identity] = $pf
    }

    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] COMPLETED: Gathering all public folder details in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
}

#Get SharePoint Site Details - Functional
function Get-SPOAndOneDriveDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory=$True,HelpMessage='Provide the service name')]
        [ValidateSet('MGGraph', 'SPO')]
        [string[]]$ServiceName
    )

    # SharePoint Online Site Details
    try {
        $start = Get-Date
        $tenantStatsHash["SharePointSites"] = @{}

        #Get all SharePoint sites for associating with Office 365 Groups / GroupMailboxes
        Write-Host "Getting all SharePoint site data..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-SPOAndOneDriveDetails] START: Gathering all SharePoint site data with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]
        switch ($detailLevel) {
            {$_ -in "minimum", "combined", "all"} { 
                $DesiredProperties = @(
                    "Template", "IsHubSite", "LastContentModifiedDate"
                    "Status", "StorageUsageCurrent", "LockIssue"
                    "LockState", "Url", "Owner"
                    "StorageQuota", "Title", "IsTeamsConnected"
                    "IsTeamsChannelConnected", "TeamsChannelType", "GroupId"
                )
                $SharePointSites = Get-SPOSite -IncludePersonalSite $True -Limit all -ErrorAction Stop | Select $DesiredProperties
            }
            geek {$SharePointSites = Get-SPOSite -IncludePersonalSite $True -Limit all -ErrorAction Stop}
        }
        Write-Log -Type INFO -Message "[Get-SPOAndOneDriveDetails] Found $($SharePointSites.count) SharePoint including OneDrive Sites" -ExportFileLocation $ExportDetails[0]

        #SharePoint data to SharePoint Hash
        #************************************************************************************
        $SharePointSitesOnly = $SharePointSites | ?{$_.Url -notlike "*-my.sharepoint.com*"}
        $progresscounter = 0
        $totalCount = $SharePointSitesOnly.count
        Write-Log -Type INFO -Message "[Get-SPOAndOneDriveDetails] Found $($SharePointSitesOnly.count) SharePoint Sites" -ExportFileLocation $ExportDetails[0]
        foreach ($site in $SharePointSitesOnly) {
            try {
                $progresscounter++
                Write-ProgressHelper -Activity "Gathering all SharePoint Online sites with $($detailLevel) details" -CurrentOperation "Checking $($site.Url)" -ProgressCounter ($progresscounter) -TotalCount $totalCount -StartTime $start
    
                #Convert SharePoint Storage to GB
                Write-Log -Type DEBUG -Message ("[Get-SPOAndOneDriveDetails] ({0}/{1}) Gathering SharePoint Site Details for '{2}'" -f $progressCounter, $totalCount, $site.Url) -ExportFileLocation $ExportDetails[0]
                $site | Add-Member -MemberType NoteProperty -Name "SharePointStorage-GB" -Value ([math]::Round($site.StorageUsageCurrent / 1024, 3)) -Force
    
                #Specify if site is connected to an Office 365 Group
                if ($site.GroupID -eq '00000000-0000-0000-0000-000000000000' -and $site.IsTeamsChannelConnected -eq $false) {
                    $site | Add-Member -MemberType NoteProperty -Name "IsOffice365GroupsConnnected" -Value $True -Force
                    Write-Log -Type DEBUG -Message ("[Get-SPOAndOneDriveDetails] ({0}/{1}) '{2}' SharePoint Site is connected to Office 365 Group: TRUE" -f $progressCounter, $totalCount, $site.Url) -ExportFileLocation $ExportDetails[0]            
                } else { 
                    $site | Add-Member -MemberType NoteProperty -Name "IsOffice365GroupsConnnected" -Value $False -Force
                    Write-Log -Type DEBUG -Message ("[Get-SPOAndOneDriveDetails] ({0}/{1}) '{2}' SharePoint Site is connected to Office 365 Group: FALSE" -f $progressCounter, $totalCount, $site.Url) -ExportFileLocation $ExportDetails[0]            
                    }
    
                #Update IsTeamsConnected to True if IsTeamsChannelConnected is True
                if ($site.IsTeamsChannelConnected -eq $true) {
                    $site | Add-Member -MemberType NoteProperty -Name "IsTeamsConnected" -Value $True -Force
                    Write-Log -Type DEBUG -Message ("[Get-SPOAndOneDriveDetails] ({0}/{1}) '{2}' SharePoint Site is connected to Microsoft Teams: TRUE" -f $progressCounter, $totalCount, $site.Url) -ExportFileLocation $ExportDetails[0]            
                } else {
                    $site | Add-Member -MemberType NoteProperty -Name "IsTeamsConnected" -Value $False -Force
                    Write-Log -Type DEBUG -Message ("[Get-SPOAndOneDriveDetails] ({0}/{1}) '{2}' SharePoint Site is connected to Microsoft Teams: FALSE" -f $progressCounter, $totalCount, $site.Url) -ExportFileLocation $ExportDetails[0]            
                }

                $tenantStatsHash["SharePointSites"][$site.Title] = $site
            }
            catch {
                $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "$($progresscounter)/$($totalCount) An error occurred in Gathering SharePoint Online Details for '$($Site.url)'. $($_.Exception.Message)"
                $global:AllDiscoveryErrors += $ErrorObject
                Write-Log -Type ERROR -Message "[Get-SPOAndOneDriveDetails] $($progresscounter)/$($totalCount) An error occurred in Gathering SharePoint Online Details for '$($Site.url)'. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
            }
        }
    }
    catch {
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-SPOAndOneDriveDetails function. Error Gathering SharePoint Online Details. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
        Write-Log -Type ERROR -Message "[Get-SPOAndOneDriveDetails] An error occurred in running Get-SPOAndOneDriveDetails function. Error Gathering SharePoint Online Details. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
    }
    finally {
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-ProgressHelper -Activity "Gathering all SharePoint Online sites with $($detailLevel) details" -Completed
    }

    # OneDrive Details
    try {
        $start = Get-Date
        $tenantStatsHash["OneDrives"] = @{}

        Write-Host "Getting all OneDrive site data..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-SPOAndOneDriveDetails] START: Gathering all OneDrive site data with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]
        $OneDriveSites = $SharePointSites | Where {$_.Url -like "*-my.sharepoint.com*"}
        Write-Log -Type INFO -Message "[Get-SPOAndOneDriveDetails] Found $($OneDriveSites.count) OneDrive Sites" -ExportFileLocation $ExportDetails[0]

        #OneDrive data to OneDrive Hash
        #************************************************************************************
        $progresscounter = 0
        $totalCount = $OneDriveSites.count
        foreach ($site in $OneDriveSites) {
            try {
                $progresscounter++
                Write-ProgressHelper -Activity "Gathering all OneDrive sites with $($detailLevel) details" -CurrentOperation "Gathering OneDrive Site Details for $($site.Url)" -ProgressCounter ($progresscounter) -TotalCount $totalCount -StartTime $start
    
                # Check if the user has a OneDrive license
                Write-Log -Type DEBUG -Message ("[Get-SPOAndOneDriveDetails] ({0}/{1}) Add OneDrive '{2}' to hash table " -f $progressCounter, $totalCount, $site.Url) -ExportFileLocation $ExportDetails[0]            
    
                try {
                    $user = Get-MgUserLicenseDetail -UserId $site.Owner -ErrorAction Stop
                    $isActive = $true
                    Write-Log -Type DEBUG -Message ("[Get-SPOAndOneDriveDetails] ({0}/{1}) Owner of '{2}' is licensed for OneDrive: TRUE" -f $progressCounter, $totalCount, $site.Url) -ExportFileLocation $ExportDetails[0]            
                }
                catch {
                    $isActive = $false
                    Write-Log -Type DEBUG -Message ("[Get-SPOAndOneDriveDetails] ({0}/{1}) Owner of '{2}' is licensed for OneDrive: FALSE" -f $progressCounter, $totalCount, $site.Url) -ExportFileLocation $ExportDetails[0]            
                }
                $site | Add-Member -MemberType NoteProperty -Name "Active" -Value $isActive
                $site | Add-Member -MemberType NoteProperty -Name "OneDriveStorage-GB" -Value ([math]::Round($site.StorageUsageCurrent / 1024, 3))
                $tenantStatsHash["OneDrives"][$site.Owner] = $site
            }
            catch {
                $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "$($progresscounter)/$($totalCount) An error occurred in Gathering OneDrive Details for '$($Site.url)'. $($_.Exception.Message)"
                $global:AllDiscoveryErrors += $ErrorObject
                Write-Log -Type ERROR -Message "[Get-SPOAndOneDriveDetails] $($progresscounter)/$($totalCount) An error occurred in Gathering OneDrive Details for '$($Site.url)'. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
            }
        }
    }
    catch {
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-SPOAndOneDriveDetails function. Error Gathering OneDrive Details. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
        Write-Log -Type ERROR -Message "[Get-SPOAndOneDriveDetails] An error occurred in running Get-SPOAndOneDriveDetails function. Error Gathering OneDrive Details. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
    }
    finally {
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-ProgressHelper -Activity "Gathering all OneDrive sites with $($detailLevel) details" -Completed
        Write-Log -Type INFO -Message "[Get-SPOAndOneDriveDetails] COMPLETED: Gathering all SharePoint site data in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
    }
}

<#Get SharePoint Site Details - Still needs work
function Get-SPOAndOneDriveDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory=$True,HelpMessage='Provide the service name')]
        [ValidateSet('MGGraph', 'SPO')]
        [string[]]$ServiceName
    )

    # Helper function to add properties to sites
    function Add-SiteProperties {
        param (
            [object]$site,
            [string]$detailLevel,
            [int]$progressCounter,
            [int]$totalCount,
            [datetime]$start
        )
        try {
            Write-ProgressHelper -Activity "Gathering site details with $($detailLevel) details" -CurrentOperation "Checking $($site.Url)" -ProgressCounter ($progressCounter) -TotalCount $totalCount -StartTime $start

            # Convert Storage to GB
            $site | Add-Member -MemberType NoteProperty -Name "Storage-GB" -Value ([math]::Round($site.StorageUsageCurrent / 1024, 3)) -Force

            # Specify if site is connected to an Office 365 Group
            $site | Add-Member -MemberType NoteProperty -Name "IsOffice365GroupsConnnected" -Value ($site.GroupID -eq '00000000-0000-0000-0000-000000000000' -and $site.IsTeamsChannelConnected -eq $false) -Force

            # Update IsTeamsConnected to True if IsTeamsChannelConnected is True
            $site | Add-Member -MemberType NoteProperty -Name "IsTeamsConnected" -Value $site.IsTeamsChannelConnected -Force

            # Store site in hash table
            if ($site.Url -like "*-my.sharepoint.com*") {
                $tenantStatsHash["OneDrives"][$site.Owner] = $site
            } else {
                $tenantStatsHash["SharePointSites"][$site.Title] = $site
            }
        }
        catch {
            $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "$($progressCounter)/$($totalCount) An error occurred in Gathering Details for '$($site.Url)'. $($_.Exception.Message)"
            $global:AllDiscoveryErrors += $ErrorObject
            Write-Log -Type ERROR -Message "[Get-SPOAndOneDriveDetails] $($progressCounter)/$($totalCount) An error occurred in Gathering Details for '$($site.Url)'. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
        }
    }

    # Fetch site details based on service name
    try {
        $start = Get-Date
        $tenantStatsHash["SharePointSites"] = @{}
        $tenantStatsHash["OneDrives"] = @{}

        Write-Host "Getting site data..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-SPOAndOneDriveDetails] START: Gathering site data with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]

        $SharePointSites = @()
        switch ($ServiceName) {
            MGGraph {
                $SharePointSites = Get-MgSite -All -ErrorAction Stop
            }
            SPO {
                switch ($detailLevel) {
                    {$_ -in "minimum", "combined", "all"} {
                        $DesiredProperties = @(
                            "Template", "IsHubSite", "LastContentModifiedDate",
                            "Status", "StorageUsageCurrent", "LockIssue",
                            "LockState", "Url", "Owner",
                            "StorageQuota", "Title", "IsTeamsConnected",
                            "IsTeamsChannelConnected", "TeamsChannelType", "GroupId"
                        )
                        $SharePointSites = Get-SPOSite -IncludePersonalSite $True -Limit all -ErrorAction Stop | Select $DesiredProperties
                    }
                    geek { $SharePointSites = Get-SPOSite -IncludePersonalSite $True -Limit all -ErrorAction Stop }
                }
            }
        }
        Write-Log -Type INFO -Message "[Get-SPOAndOneDriveDetails] Found $($SharePointSites.count) SharePoint including OneDrive Sites" -ExportFileLocation $ExportDetails[0]

        # Process all sites
        $progresscounter = 0
        $totalCount = $SharePointSites.count
        foreach ($site in $SharePointSites) {
            $progresscounter++
            Add-SiteProperties -site $site -detailLevel $detailLevel -progressCounter $progresscounter -totalCount $totalCount -start $start
        }
    }
    catch {
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-SPOAndOneDriveDetails function. Error Gathering Site Details. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
        Write-Log -Type ERROR -Message "[Get-SPOAndOneDriveDetails] An error occurred in running Get-SPOAndOneDriveDetails function. Error Gathering Site Details. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
    }
    finally {
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-ProgressHelper -Activity "Gathering all sites with $($detailLevel) details" -Completed
        Write-Log -Type INFO -Message "[Get-SPOAndOneDriveDetails] COMPLETED: Gathering all site data in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
    }
}#>

#Get Teams Details
function Get-TeamsDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory=$True,HelpMessage='Provide the service name')]
        [ValidateSet('MGGraph', 'Teams')]
        [string[]]$ServiceName
    )

    try {
        $start = Get-Date
        $tenantStatsHash["AllTeams"] = @{}

        Write-Host "Getting all Microsoft Teams details ..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-TeamsDetails] START: Gathering all Microsoft Teams with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]

        $allTeams = @()
        switch ($ServiceName) {
            MGGraph {
                $allTeams = Get-MgTeam -All -ErrorAction Stop
            }
            Teams {
                # Fetch all Teams
                $ProgressPreference = "SilentlyContinue"
                $allTeams = Get-Team -ErrorAction Stop
                $ProgressPreference = "Continue"
            }
        }

        # Check for error of user not licensed to run Get-Team
        if ($allTeams -like "*ErrorMessage: Failed to get license information for the user. Ensure user has a valid Office365 license assigned to them*") {
            Write-Log -Type ERROR -Message "[Get-TeamsDetails] An error occurred in running Get-TeamsDetails function. Exception: Failed to get license information for the user. Ensure user has a valid Office365 license assigned to them" -ExportFileLocation $ExportDetails[0]
            throw "Failed to get Teams Details. Ensure user has a valid Office365 license assigned to them"
        } else {
            $progresscounter = 0
            $totalCount = $allTeams.count
            foreach ($team in $allTeams) {
                try {
                    $progresscounter++
                    
                    Write-ProgressHelper -Activity "Gathering all Microsoft Teams with $($detailLevel) details" -CurrentOperation "Gathering Team Details for $($team.DisplayName)" -ProgressCounter ($progresscounter) -TotalCount $totalCount -StartTime $start
                    Write-Log -Type DEBUG -Message ("[Get-TeamsDetails] ({0}/{1}) Gathering Team Details for {2}" -f $progressCounter, $totalCount, $team.DisplayName) -ExportFileLocation $ExportDetails[0]

                    # Fetch SharePoint site size for each Team
                    Write-Log -Type DEBUG -Message ("[Get-TeamsDetails] ({0}/{1}) Gathering SharePoint Online Details for {2}" -f $progressCounter, $totalCount, $team.DisplayName) -ExportFileLocation $ExportDetails[0]
                    $SPOSiteDetails = $tenantStatsHash["SharePointSites"][$team.displayName]
                    if ($SPOSiteDetails.Template -eq "TEAMCHANNEL#0" -or $SPOSiteDetails.Template -eq "GROUP#0") {
                        $siteSize = $SPOSiteDetails.StorageUsageCurrent
                        $siteSizeGB = [math]::Round($SPOSiteDetails.StorageUsageCurrent / 1024, 3)
                    } else {
                        $siteSize = 0
                        $siteSizeGB = 0
                    }

                    # Fetch Channels for each Team
                    Write-Log -Type DEBUG -Message ("[Get-TeamsDetails] ({0}/{1}) Gathering Channels {2}" -f $progressCounter, $totalCount, $team.DisplayName) -ExportFileLocation $ExportDetails[0]
                    $channels = @()
                    switch ($ServiceName) {
                        MGGraph {
                            $teamId = $team.Id
                            $channels = Get-MgTeamChannel -TeamId $teamId -ErrorAction SilentlyContinue
                        }
                        Teams {
                            $teamId = $team.GroupID
                            $channels = Get-TeamChannel -GroupId $teamId -ErrorAction SilentlyContinue
                        }
                    }

                    # Initialize separate arrays for each channel type
                    $publicChannels = @()
                    $privateChannels = @()
                    $sharedChannels = @()
                    foreach ($channel in $channels) {
                        Write-Log -Type DEBUG -Message ("[Get-TeamsDetails] ({0}/{1}) Gathering Channels Details for {2} - {3}" -f $progressCounter, $totalCount, $team.DisplayName, $channel.DisplayName) -ExportFileLocation $ExportDetails[0]
                        $channelType = if ($channel.MembershipType -eq "Private") { "private" } elseif ($channel.MembershipType -eq "Standard") { "public" } else { "shared" }

                        # Add the channel names to the respective arrays based on their type
                        switch ($channelType) {
                            "public" { $publicChannels += $channel.DisplayName }
                            "private" { $privateChannels += $channel.DisplayName }
                            "shared" { $sharedChannels += $channel.DisplayName }
                        }
                    }
                    $TotalNumberOfChannels = $publicChannels.Count + $privateChannels.Count + $sharedChannels.Count

                    # Create output object
                    $currentTeam = [ordered]@{
                        DisplayName       = $team.DisplayName
                        Description       = $team.Description
                        Visibility        = $team.Visibility
                        SharePointSiteUrl = $SPOSiteDetails.Url
                        "SiteSize-GB"     = $siteSizeGB
                        SiteSize          = $siteSize
                        TotalChannels     = $TotalNumberOfChannels
                        PublicChannels    = $publicChannels -join ','
                        PrivateChannels   = $privateChannels -join ','
                        SharedChannels    = $sharedChannels -join ','
                    }
                    $tenantStatsHash["AllTeams"][$team.DisplayName] = $currentTeam
                }
                catch {
                    $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in Gathering Teams Details for $($team.DisplayName). $($_.Exception.Message)"
                    $global:AllDiscoveryErrors += $ErrorObject
                    Write-Log -Type ERROR -Message "[Get-TeamsDetails] An error occurred in Gathering Teams Details for $($team.DisplayName). $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
                }

                Write-ProgressHelper -Activity "Gathering all Microsoft Teams with $($detailLevel) details" -Completed
            }
        }
    }
    catch {
        if ($_.Exception.Message -like "*Ensure user has a valid Office365 license assigned to them*") {
            $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-TeamsDetails function. User NOT Assigned Valid O365 License. $($_.Exception.Message)"
            $global:AllDiscoveryErrors += $ErrorObject
            Write-Log -Type ERROR -Message "[Get-TeamsDetails] An error occurred in running Get-TeamsDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
        }
        else {
            $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-TeamsDetails function. $($_.Exception.Message)"
            $global:AllDiscoveryErrors += $ErrorObject
            Write-Log -Type ERROR -Message "[Get-TeamsDetails] An error occurred in running Get-TeamsDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
        }
    }
    finally {
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-ProgressHelper -Activity "Gathering all Microsoft Teams with $($detailLevel) details" -Completed
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-TeamsDetails] COMPLETED: Gathering all Microsoft Teams in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
    }
}

# ----------------------------------
# Tenant Specific Functions
# ----------------------------------
#Verify all required modules are installed and Connect to Office 365 Services
#updated to support windows 7+ powershell
function Connect-Office365Services {
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'Provide the service name')]
        [ValidateSet('ExchangeOnline', 'EXO', 'MSOnline', 'MSO', 'MGGraph', 'Graph', 'MGG', 'SharePointOnline', 'SPO', 'ALL', 'AzureAD', 'Teams')]
        [string[]]$ServiceName,
        [Parameter(Mandatory = $false, HelpMessage = 'Specify if Graph should be read only')]
        [Switch]$ReadOnlyGraph
    )

    # Check if running on PowerShell 7
    $isPowerShell7 = $PSVersionTable.PSVersion.Major -ge 7

    # Helper function to check and import modules
    function Import-RequiredModule {
        param (
            [string]$ModuleName,
            [string]$InstallMessage,
            [Switch]$UseWindowsPowerShell = $false
        )
        if ($null -ne (Get-InstalledModule -Name $ModuleName -ErrorAction SilentlyContinue)) {
            if ($null -eq (Get-Module -Name $ModuleName)) {
                Write-Host "Importing $ModuleName module..." -ForegroundColor Yellow
                try {
                    if ($UseWindowsPowerShell) {
                        Import-Module $ModuleName -UseWindowsPowerShell -ErrorAction Stop
                    } else {
                        Import-Module $ModuleName -Force -ErrorAction Stop
                    }
                    Write-Host "$ModuleName module imported successfully." -ForegroundColor Green
                } catch {
                    Write-Error "Error importing $ModuleName module: $($_.Exception.Message)"
                    throw
                }
            }
        } else {
            Write-Error "$InstallMessage"
            throw "$ModuleName module not found."
        }
    }

    # Helper function to update the title bar
    function Update-TitleBar {
        param (
            [string]$Title
        )
        $host.ui.RawUI.WindowTitle = $Title
    }

    # Helper function to prompt user for tenant confirmation
    function Confirm-Tenant {
        param (
            [string]$TenantName
        )
        Write-Host "Already Connected: $TenantName" -ForegroundColor Green
        $userInput = Read-Host -Prompt "Connected to the correct tenant? (Yes/Y/No/N)"
        if ($userInput -in "Yes", "Y") {
            return $true
        } else {
            return $false
        }
    }

    # Function to connect to Exchange Online
    function Connect-ExchangeOnlineModuleAndService {
        Write-Host "Exchange Online: Checking for Existing Connections and Required Modules" -ForegroundColor Cyan
        try {
            $EXOOrgCheck = Get-OrganizationConfig -ErrorAction Stop
            if (Confirm-Tenant $EXOOrgCheck.Name) {
                $global:tenant = $EXOOrgCheck.Name
                return
            }
        } catch {
            Import-RequiredModule -ModuleName "ExchangeOnlineManagement" -InstallMessage "Run 'Install-Module ExchangeOnlineManagement' as an Administrator."
        }

        Write-Host "Connecting to ExchangeOnline..." -ForegroundColor Yellow
        try {
            Connect-ExchangeOnline -ErrorAction Stop *> $null
            $EXOOrgCheck = Get-OrganizationConfig -ErrorAction Stop
            Update-TitleBar $EXOOrgCheck.Name
            Write-Host "Connected: $($EXOOrgCheck.Name)" -ForegroundColor Green
            $global:tenant = $EXOOrgCheck.Name
        } catch {
            Write-Error "Error connecting to ExchangeOnline: $($_.Exception.Message)"
        }
    }

    # Function to connect to MSOnline
    function Connect-MSOnlineModuleAndService {
        Write-Host "MSOnline: Checking for Existing Connections and Required Modules" -ForegroundColor Cyan
        try {
            $MSOCompanyCheck = Get-MsolCompanyInformation -ErrorAction Stop
            if (Confirm-Tenant $MSOCompanyCheck.DisplayName) {
                return
            }
        } catch {
            if ($isPowerShell7) {
                Import-RequiredModule -UseWindowsPowerShell -ModuleName "MSOnline" -InstallMessage "Run 'Install-Module MSOnline -Force' as an Administrator."
            } else {
                Import-RequiredModule -ModuleName "MSOnline" -InstallMessage "Run 'Install-Module MSOnline -Force' as an Administrator."
            }
        }

        Write-Host "Connecting to MSOnline..." -ForegroundColor Yellow
        try {
            Connect-MsolService -ErrorAction Stop
            $MSOCompanyCheck = Get-MsolCompanyInformation -ErrorAction Stop
            Update-TitleBar $MSOCompanyCheck.DisplayName
            Write-Host "Connected: $($MSOCompanyCheck.DisplayName)" -ForegroundColor Green
        } catch {
            Write-Error "Error connecting to MSOnline: $($_.Exception.Message)"
        }
    }

    # Function to connect to Microsoft Graph
    function Connect-MGGraphModuleAndService {
        Write-Host "Microsoft Graph: Checking for Existing Connections and Required Modules" -ForegroundColor Cyan
        $global:MGGraph = $null
        try {
            $MGraphCompanyCheck = Get-MgOrganization -ErrorAction Stop
            if (Confirm-Tenant $MGraphCompanyCheck.DisplayName) {
                return
            } else {
                Disconnect-MgGraph
            }
        } catch {
            if ($null -ne (Get-InstalledModule -Name Microsoft.Graph.* -ErrorAction SilentlyContinue)) {
                if ($null -eq (Get-Module -Name Microsoft.Graph.*)) {
                    Write-Host "Importing 'Microsoft.Graph' module..." -ForegroundColor Yellow
                    try {
                        Import-Module Microsoft.Graph-Force -ErrorAction Stop
                        Write-Host "Microsoft.Graph module imported successfully." -ForegroundColor Green
                    } catch {
                        Write-Error "Error importing Microsoft.Graph module: $($_.Exception.Message)"
                        throw
                    }
                }
            } else {
                Write-Error "Run Install-Module Microsoft.Graph -AllowClobber -Force as an Administrator."
                throw "Microsoft.Graph module not found."
            }
            #Import-RequiredModule -ModuleName "Microsoft.Graph.*" -InstallMessage "Run Install-Module Microsoft.Graph -AllowClobber -Force as an Administrator."
        }

        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
        try {
            $RequiredScopes = if ($ReadOnlyGraph) {
                @(
                    "Directory.Read.All", "Synchronization.Read.All",
                    "Organization.Read.All", "LicenseAssignment.ReadWrite.All",
                    "AuditLog.Read.All", "Directory.AccessAsUser.All",
                    "IdentityRiskyUser.Read.All", "IdentityUserFlow.Read.All",
                    "EAS.AccessAsUser.All", "EWS.AccessAsUser.All",
                    "Team.ReadBasic.All", "TeamsAppInstallation.ReadWriteForUser",
                    "TeamsAppInstallation.ReadWriteSelfForUser", "TeamsTab.ReadWriteForUser",
                    "TeamsTab.ReadWriteSelfForUser", "User.EnableDisableAccount.All",
                    "User.Export.All", "User.Invite.All", "User.ManageIdentities.All",
                    "User.Read.All", "UserActivity.ReadWrite.CreatedByApp",
                    "UserAuthenticationMethod.Read.All", "User-LifeCycleInfo.Read.All",
                    "Device.Read.All", "DeviceManagementManagedDevices.Read.All",
                    "AuthenticationContext.Read.All", "Policy.ReadWrite.AuthenticationMethod",
                    "Domain.Read.All", "Group.Read.All", "GroupMember.Read.All",
                    "SharePointTenantSettings.Read.All", "SecurityEvents.Read.All"
                )
            } else {
                @(
                    "Directory.ReadWrite.All", "Synchronization.ReadWrite.All",
                    "Organization.ReadWrite.All", "LicenseAssignment.ReadWrite.All",
                    "AuditLog.Read.All", "Directory.AccessAsUser.All",
                    "IdentityRiskyUser.ReadWrite.All", "IdentityUserFlow.ReadWrite.All",
                    "EAS.AccessAsUser.All", "EWS.AccessAsUser.All",
                    "Team.ReadBasic.All", "TeamsAppInstallation.ReadWriteForUser",
                    "TeamsAppInstallation.ReadWriteSelfForUser", "TeamsTab.ReadWriteForUser",
                    "TeamsTab.ReadWriteSelfForUser", "User.EnableDisableAccount.All",
                    "User.Export.All", "User.Invite.All", "User.ManageIdentities.All",
                    "User.ReadWrite.All", "UserActivity.ReadWrite.CreatedByApp",
                    "UserAuthenticationMethod.ReadWrite.All", "User-LifeCycleInfo.ReadWrite.All",
                    "Device.Read.All", "DeviceManagementManagedDevices.ReadWrite.All",
                    "AuthenticationContext.ReadWrite.All", "Policy.ReadWrite.AuthenticationMethod",
                    "Domain.ReadWrite.All", "Group.ReadWrite.All", "GroupMember.ReadWrite.All",
                    "SharePointTenantSettings.ReadWrite.All", "SecurityEvents.ReadWrite.All"
                )
            }

            # Ensure all Microsoft.Graph modules are imported
            $graphModules = Get-InstalledModule -Name "Microsoft.Graph.*" -ErrorAction SilentlyContinue
            if ($graphModules) {
                foreach ($graphModule in $graphModules) {
                    if ($null -eq (Get-Module -Name $graphModule.Name -ErrorAction SilentlyContinue)) {
                        Write-Host "Importing $($graphModule.Name) module. This could take a while.." -ForegroundColor Yellow
                        try {
                            Import-Module $graphModule.Name -Force -ErrorAction Stop
                            Write-Host "$($graphModule.Name) module imported successfully." -ForegroundColor Green
                        } catch {
                            Write-Error "Error importing $($graphModule.Name) module: $($_.Exception.Message)"
                            throw
                        }
                    }
                }
            } else {
                Write-Error "Run 'Install-Module Microsoft.Graph, Microsoft.Graph.Beta -AllowClobber -Force' as an Administrator. If that does not work, manually delete the installed modules at 'local OneDrive folder Documents\PowerShell\Modules' or 'C:\Program Files\WindowsPowerShell\Modules'."
                throw "Microsoft.Graph.* modules not found."
            }

            Connect-MgGraph -Scopes $RequiredScopes -ErrorAction Stop
            $MGraphCompanyCheck = Get-MgOrganization -ErrorAction Stop
            Update-TitleBar $MGraphCompanyCheck.DisplayName
            Write-Host "Connected: $($MGraphCompanyCheck.DisplayName)" -ForegroundColor Green
        } catch {
            Write-Error "Error connecting to Microsoft Graph: $($_.Exception.Message)"
        }
    }

    # Function to connect to SharePoint Online
    function Connect-SharePointOnlineModuleAndService {
        Write-Host "SharePoint Online: Checking for Existing Connections and Required Modules" -ForegroundColor Cyan
        try {
            $rootSiteURL = Get-SPOSite -Limit 1 -ErrorAction Stop -WarningAction SilentlyContinue
            $rootURL = $rootSiteURL.Url -replace '/sites.*', ''
            if (Confirm-Tenant $rootURL) {
                return
            }
        } catch {
            if ($isPowerShell7) {
                Import-RequiredModule -UseWindowsPowerShell -ModuleName "Microsoft.Online.SharePoint.PowerShell" -InstallMessage "Run 'Install-Module Microsoft.Online.SharePoint.PowerShell' as an Administrator."
            } else {
                Import-RequiredModule -ModuleName "Microsoft.Online.SharePoint.PowerShell" -InstallMessage "Run 'Install-Module Microsoft.Online.SharePoint.PowerShell' as an Administrator."
            }
        }

        Write-Host "Connecting to SharePoint Online..." -ForegroundColor Yellow
        try {
            $SPOAdminURL = Read-Host -Prompt "Provide the SharePoint Online Admin URL. Name is usually formatted 'https://<yourtenant>-admin.sharepoint.com'"
            Connect-SPOService -Url $SPOAdminURL -ErrorAction Stop
            $rootSiteURL = Get-SPOSite -Limit 1 -ErrorAction Stop -WarningAction SilentlyContinue
            $rootURL = $rootSiteURL.Url -replace '/sites.*', ''
            Update-TitleBar $rootURL
            Write-Host "Connected: $($rootURL)" -ForegroundColor Green
        } catch {
            Write-Error "Error connecting to SharePoint Online: $($_.Exception.Message)"
        }
    }

    # Function to connect to Azure AD
    function Connect-AzureADModuleAndService {
        Write-Host "AzureAD: Checking for Existing Connections and Required Modules" -ForegroundColor Cyan
        try {
            $AzureADTenant = (Get-AzureADTenantDetail).DisplayName
            if (Confirm-Tenant $AzureADTenant) {
                return
            }
        } catch {
            Import-RequiredModule -ModuleName "AzureAD" -InstallMessage "Run 'Install-Module AzureAD' as an Administrator."
        }

        Write-Host "Connecting to AzureAD..." -ForegroundColor Yellow
        try {
            $result = Connect-AzureAD -ErrorAction Stop
            $AzureADTenant = (Get-AzureADTenantDetail).DisplayName
            Update-TitleBar $AzureADTenant
            Write-Host "Connected: $($AzureADTenant)" -ForegroundColor Green
        } catch {
            Write-Error "Error connecting to AzureAD: $($_.Exception.Message)"
        }
    }

    # Function to connect to Microsoft Teams
    function Connect-MicrosoftTeamsModuleAndService {
        Write-Host "Microsoft Teams: Checking for Existing Connections and Required Modules" -ForegroundColor Cyan
        try {
            $MicrosoftTeamsTenant = (Get-CsTenant).DisplayName
            if (Confirm-Tenant $MicrosoftTeamsTenant) {
                return
            } else {
                Disconnect-MicrosoftTeams
            }
        } catch {
            Import-RequiredModule -ModuleName "MicrosoftTeams" -InstallMessage "Run 'Install-Module MicrosoftTeams' as an Administrator."
        }

        Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Yellow
        try {
            Connect-MicrosoftTeams -ErrorAction Stop
            $MicrosoftTeamsTenant = (Get-CsTenant).DisplayName
            Update-TitleBar $MicrosoftTeamsTenant
            Write-Host "Connected: $($MicrosoftTeamsTenant)" -ForegroundColor Green
        } catch {
            Write-Error "Error connecting to Microsoft Teams: $($_.Exception.Message)"
        }
    }

    # Main Function Logic
    Write-Host "Connecting to Required Office 365 Services" -ForegroundColor Black -BackgroundColor Yellow
    $global:confirmation = $false

    while (-not $global:confirmation) {
        foreach ($service in $ServiceName) {
            switch ($service) {
                {$_ -in 'ExchangeOnline', 'EXO'} { Connect-ExchangeOnlineModuleAndService }
                {$_ -in 'Teams'} { Connect-MicrosoftTeamsModuleAndService }
                {$_ -in 'MSOnline', 'MSO'} { Connect-MSOnlineModuleAndService }
                {$_ -in 'MGGraph', 'MGG', 'Graph'} {
                    try {
                        Connect-MGGraphModuleAndService
                        $global:MGGraph = $true
                    } catch {
                        Write-Host "Failed to connect using MGGraph. Trying MSOnline instead..." -ForegroundColor Yellow
                        try {
                            $global:MGGraph = $false
                            Connect-MSOnlineModuleAndService
                        } catch {
                            Write-Error "Failed to connect using both MGGraph and MSOnline: $($_.Exception.Message)"
                            throw
                        }
                    }
                }
                {$_ -in 'SharePointOnline', 'SPO'} { Connect-SharePointOnlineModuleAndService }
                {$_ -in 'AzureAD'} { Connect-AzureADModuleAndService }
                'ALL' {
                    #Connect-MSOnlineModuleAndService
                    try {
                        if ($isPowerShell7) {
                            Connect-MGGraphModuleAndService
                            $global:MGGraph = $true
                        }
                        else {
                            throw
                        }
                    } catch {
                        try {
                            Write-Host "Failed to connect using MGGraph. Trying AzureAD, MSOnline, and MicrosoftTeams instead..." -ForegroundColor Yellow
                            $global:MGGraph = $false
                            Connect-AzureADModuleAndService
                            Connect-MSOnlineModuleAndService
                            Connect-MicrosoftTeamsModuleAndService
                            
                        } catch {
                            Write-Error "Failed to connect using MGGraph, AzureAD, and MSOnline: $($_.Exception.Message)"
                            throw                            
                        }
                    }
                    Connect-SharePointOnlineModuleAndService
                    Connect-ExchangeOnlineModuleAndService           
                }
            }
        }

        # Verify Connection
        Write-Host
        $userInput = Read-Host -Prompt "Have all the services connected correctly? (Yes/No). Default (blank) is Yes"
        $userInput = $userInput.ToLower()
        switch ($userInput) {
            {$_ -in 'yes', 'y', ''} {
                Write-Host "Connected to Correct Services!" -ForegroundColor Green
                $global:confirmation = $true
            }
            {$_ -in 'no', 'n' } {
                Write-Warning "Not Connected to Correct Services"
                $ServiceName = Read-Host -Prompt "Please provide the service name to connect"
            }
        }
    }
}

# License SKUs and Service Plan IDs to HASH - MSOL or MGGraph
function Get-AllLicenseSKUs {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the service name')]
        [ValidateSet('MSOL','MGGraph','Azure')]
        [string[]]$ServiceName
    )
    #Get the start time of the function
    $start = Get-Date
    #Build a hashtable for looking up license names from license sku
    $tenantStatsHash["LicenseSKUs"] = [ordered]@{}
    #store Service Plan IDs and corresponding ServicePlan data for each ID
    $tenantStatsHash["ServicePlans"] = [ordered]@{}

    Write-Progress -Activity "Adding License SKUs and Service Plan IDs to Hash" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
    Write-Log -Type Info -Message "[Get-AllLicenseSKUs] Gathering all $($ServiceName) License SKUs from tenant" -ExportFileLocation $ExportDetails[0]
    # Get License SKUs using MSOL or MGGraph
    switch ($ServiceName) {
        MSOL { $skus = Get-MsolAccountSku }
        MGGraph {    
            $DesiredProperties = @(
            "AppliesTo", "ConsumedUnits", "PrepaidUnits"
            "ServicePlans", "SkuId", "SkuPartNumber"
            )
            # Get License SKUs
            $skus = Get-MgSubscribedSku -ErrorAction Continue | Select $DesiredProperties | ? {$_.AppliesTo}
        }
    }
    Write-Log -Type Info -Message "[Get-AllLicenseSKUs] Found $($skus.count) License SKUs from tenant" -ExportFileLocation $ExportDetails[0]

    #Add License SKUs to Hash Table and Service Plans under each SKU to another Hash Table
    Write-Log -Type Info -Message "[Get-AllLicenseSKUs] START: Add License SKUs and Service Plans under each SKU to Tenant Hash Table" -ExportFileLocation $ExportDetails[0]
    foreach ($sku in $skus) {
        #Get AccountSkuId or SkuId depending on service
        switch ($ServiceName) {
            MSOL {                 
                #Add Additional Properties to License Hash Table
                $sku | Add-Member -MemberType NoteProperty -Name AppliesTo -Value $sku.TargetClass -Force
                #Add Prepaid Units
                $sku | Add-Member -MemberType NoteProperty -Name PrepaidUnits_Enabled -Value $sku.ActiveUnits -Force
                #Add Remaining/Available Units
                $sku | Add-Member -MemberType NoteProperty -Name RemainingUnits -Value ($sku.ActiveUnits - $sku.ConsumedUnits) -Force
                $sku | Add-Member -MemberType NoteProperty -Name ServicePlans_Details -Value ($sku.ServiceStatus.ServicePlan.ServiceName -join ",") -Force
            }
            MGGraph { 
                #Add Additional Properties to License Hash Table
                $sku | Add-Member -MemberType NoteProperty -Name AppliesTo -Value $sku.AppliesTo -Force
                #Add Prepaid Units
                $sku | Add-Member -MemberType NoteProperty -Name PrepaidUnits_Enabled -Value $sku.PrepaidUnits.Enabled -Force
                #Add Remaining/Available Units
                $sku | Add-Member -MemberType NoteProperty -Name RemainingUnits -Value ($sku.PrepaidUnits.Enabled - $sku.ConsumedUnits) -Force
                $sku | Add-Member -MemberType NoteProperty -Name ServicePlans_Details -Value ($sku.ServicePlans.ServicePlanName -join ",") -Force
            }
        }
        Write-Log -Type DEBUG -Message "[Get-AllLicenseSKUs] Gathering License details for $($AccountSkuId)" -ExportFileLocation $ExportDetails[0]

        #Create License Details Array - Ordered
        $licenseDetails = [ordered]@{
            "AppliesTo" = $sku.AppliesTo
            "SkuId" = $sku.SkuId
            "LicenseName" = $sku.SkuPartNumber
            "PurchasedUnits" = $sku.PrepaidUnits_Enabled
            "ConsumedUnits" = $sku.ConsumedUnits
            "RemainingUnits" = $sku.RemainingUnits
            "ServicePlans_Details" = $sku.ServicePlans_Details
        }

        #Create Hash Table for License SKUs
        $tenantStatsHash["LicenseSKUs"][$sku.SkuId.tostring()] = $licenseDetails

        # Service Status Details
        switch ($ServiceName) {
            MSOL { 
                foreach($servicePlan in $sku.ServiceStatus) {
                    $key = $servicePlan.ServicePlan.ServiceName
                    $value = $servicePlan.ServicePlan.ServiceType
                    Write-Log -Type DEBUG -Message "[Get-AllLicenseSKUs] Gathering Service Plan details for $($AccountSkuId): $($key) " -ExportFileLocation $ExportDetails[0]
                    $tenantStatsHash["ServicePlans"][$key] = $value
                }
             }
            MGGraph {
                foreach($servicePlan in $sku.ServicePlans) {
                    $key = $servicePlan.ServicePlanId
                    Write-Log -Type DEBUG -Message "[Get-AllLicenseSKUs] Gathering Service Plan details for $($AccountSkuId): $($key) " -ExportFileLocation $ExportDetails[0]

                    # If the ServicePlanId doesn't exist as a key, add it
                    if ($key -notin $tenantStatsHash["ServicePlans"].Keys) {
                        $value = $servicePlan.ServicePlanName
                        $tenantStatsHash["ServicePlans"][$key] = $value
                    }
                }
            }
        }
    }
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Log -Type Info -Message "[Get-AllLicenseSKUs] COMPLETED: Gathering all License Details User Details. ($($ServiceName)) in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
}

function Get-LicenseFriendlyName {
    param (
        [Parameter(Mandatory=$false)]
        [Array]$AssignedLicenses,
        [Parameter(Mandatory=$true)]
        [Hashtable]$LicenseHash
    )
    #Array for adding each license name to the allLicenses for this user
    $allLicenses = @()

    #Array for adding all disabled service names for this user
    $allDisabledPlans = @()
    #Process each license to get friendly names and disabled service plans for each license
    foreach($license in $AssignedLicenses) {
        Write-Log -Type DEBUG -Message "[Get-LicenseFriendlyName] START: Getting License Friendly Names for $($license.SKUID)" -ExportFileLocation $ExportDetails[0]
        $licenseName = $tenantStatsHash["LicenseSKUs"][$license.SKUID].LicenseName
        Write-Log -Type DEBUG -Message "[Get-LicenseFriendlyName] Found: License Friendly Name - $($license.SKUID): $($licenseName)" -ExportFileLocation $ExportDetails[0]
        $allLicenses += $licenseName
        try {
            foreach($disabledPlan in $license.DisabledPlans) {
                #Write-Output $disabledPlan
                $disabledPlanName = $tenantStatsHash["servicePlans"][$disabledPlan.toString()]
                $allDisabledPlans += $disabledPlanName
                Write-Log -Type DEBUG -Message "[Get-LicenseFriendlyName] Found: $($licenseName) Disabled Plans - $(($allDisabledPlans | measure).count) count" -ExportFileLocation $ExportDetails[0]
            }
        } catch {
            $allDisabledPlans = $null
            Write-Log -Type DEBUG -Message "[Get-LicenseFriendlyName] No Disabled Plans for $($licenseName)" -ExportFileLocation $ExportDetails[0]
        }
    }
    return $allLicenses, $allDisabledPlans
    
}

function Get-allUserDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory=$True,HelpMessage='Provide the service name')]
        [ValidateSet('MSOL','MGGraph','Azure')]
        [string[]]$ServiceName
    )
    #Get the start time of the function
    $start = Get-Date
    $tenantStatsHash["Users"] = @{} #Hash table to store all user details

    # Function to Match Mailbox, OneDrive and Archive Stats
    function Get-MatchingStats {
        param (
            [Parameter(Mandatory = $true)]
            [array]$user,
            
            [Parameter(Mandatory = $true)]
            [hashtable]$tenantStatsHash,
            
            [Parameter(Mandatory = $false)]
            [string]$logReportDirectory
        )
    
        $onmicrosoftAlias = $null
        $MBXSizeGB = $null
        $oneDriveData = @()
        $archiveStats = @()
    
        if ($mailbox = $tenantStatsHash["AllMailboxes"][$user.UserPrincipalName]) {
            #Get OnMicrosoft Alias
            try {
                $onmicrosoftAlias = ($mailbox.EmailAddresses | Where-Object { $_ -like "*@*.onmicrosoft.com" } | Select-Object -First 1).Replace("SMTP:", "").Replace("smtp:", "")
                Write-Log -Type DEBUG -Message "[Get-MatchingStats] FOUND: '$($user.UserPrincipalName)' OnMicrosoft Address for Combined/All Reports" -ExportFileLocation $ExportDetails[0]                                
            } catch {
                Write-Log -Type DEBUG -Message "[Get-MatchingStats] NOT FOUND: '$($user.UserPrincipalName)' OnMicrosoft Address for Combined/All Reports" -ExportFileLocation $ExportDetails[0]
            }
    
            #Get Mailbox Stats
            if($mbxStats = $tenantStatsHash["PrimaryMailboxStats"][$mailbox.ExchangeGuid.ToString()]) {
                $MBXSizeGB = [math]::Round(($MBXStats.TotalItemSize.ToString() -replace "(.*\()|,| [a-z]*\)","")/1GB,3)
                Write-Log -Type DEBUG -Message "[Get-MatchingStats] FOUND: '$($user.UserPrincipalName)' Mailbox Stats for Combined/All Reports" -ExportFileLocation $ExportDetails[0]
            } else {
                Write-Log -Type DEBUG -Message "[Get-MatchingStats] NOT FOUND: '$($user.UserPrincipalName)' Mailbox Stats for Combined/All Reports" -ExportFileLocation $ExportDetails[0]
            }
    
            #Get OneDrive Stats
            if ($tenantStatsHash["OneDrives"][$mailbox.UserPrincipalName]) {
                $oneDriveData = $tenantStatsHash["OneDrives"][$mailbox.UserPrincipalName]
                Write-Log -Type DEBUG -Message "[Get-MatchingStats] FOUND: '$($user.UserPrincipalName)' OneDrive Details for Combined/All Reports" -ExportFileLocation $ExportDetails[0]
            }
            else {
                Write-Log -Type DEBUG -Message "[Get-MatchingStats] NOT FOUND: '$($user.UserPrincipalName)' OneDrive Details for Combined/All Reports" -ExportFileLocation $ExportDetails[0]
            }
    
            #Get Archive Stats
            if ($mailbox.ArchiveStatus -ne "None" -and $null -ne $mailbox.ArchiveStatus) {
                $archiveStats = $tenantStatsHash["ArchiveMailboxStats"][$mailbox.ArchiveGuid.ToString()]
                $archiveStatsGB = [math]::Round(($archiveStats.TotalItemSize.ToString() -replace "(.*\()|,| [a-z]*\)","")/1GB,3)
                Write-Log -Type DEBUG -Message "[Get-MatchingStats] FOUND: '$($user.UserPrincipalName)' Archive Stats for Combined/All Reports" -ExportFileLocation $ExportDetails[0]
            }
            else {
                Write-Log -Type DEBUG -Message "[Get-MatchingStats] NOT FOUND: '$($user.UserPrincipalName)' Archive Stats for Combined/All Reports" -ExportFileLocation $ExportDetails[0]
            }
        }
    
        # Return relevant data if you need
        return @{
            Mailbox = $mailbox
            MBXStats = $mbxStats
            OnMicrosoftAlias = $onmicrosoftAlias
            MBXSizeGB = $MBXSizeGB
            OneDriveData = $oneDriveData
            ArchiveStats = $archiveStats
            ArchiveStatsGB = $archiveStatsGB
        }
    }

    #Gather All User Details - MSOL OR MGGraph
    switch ($ServiceName) {
        MSOL {
            #Gather all MSOL User Details
            try {
                Write-Host "Getting all MSOnline $($detailLevel) User data..." -ForegroundColor Cyan -nonewline
                Write-Log -Type Info -Message "[Get-allUserDetails] START: Getting all MSOnline $($detailLevel) User data" -ExportFileLocation $ExportDetails[0]
                Write-Progress -Activity "Getting all MSOnline User Data" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
                switch ($detailLevel) {
                    #For minimum details, you may want to adjust this selection.
                    minimum { 
                        $DesiredProperties = @(
                            "DisplayName", "UserPrincipalName", "UserType", "LicenseAssignmentDetails", "Licenses", "ProxyAddresses", "ObjectId"
                        )
                        $allTenantUsers = Get-MsolUser -All | select $DesiredProperties
                    }
                    {$_ -in "all", "combined"} { 
                        #Adjusted these properties according to Get-MsolUser cmdlet's output.
                        $DesiredProperties = @(
                            "DisplayName", "UserPrincipalName", "UserType", "Office"
                            "ObjectId", "IsLicensed", "LastDirSyncTime", 
                            "BlockCredential", "ImmutableId","LicenseAssignmentDetails"
                            "ProxyAddresses","SoftDeletionTimestamp","Title","UsageLocation","WhenCreated"
                        )
                        $allTenantUsers = Get-MsolUser -All | select $DesiredProperties
                    }
                    geek {
                        #Retrieving all properties for geek detail level.
                        $allTenantUsers = Get-MsolUser -All
                    }
                }
                Write-Progress -Activity "Getting all MSOnline User Data" -Completed

            }
            catch {
                Write-Log -Type ERROR -Message "[Get-allUserDetails] An error occurred in running Get-allMSOLUserDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
                $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-allMSOLUserDetails function. $($_.Exception.Message)"
                $global:AllDiscoveryErrors += $ErrorObject
            }
        }
        MGGraph {
            # Gather all Microsoft Graph User Details
            try {
                Write-Host "Getting all Microsoft Graph $($detailLevel) User data..." -ForegroundColor Cyan -nonewline
                Write-Log -Type Info -Message "[Get-allUserDetails] START: Getting all Microsoft Graph $($detailLevel) User data" -ExportFileLocation $ExportDetails[0]
                Write-Progress -Activity "Getting all Microsoft Graph User Data" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
        
                switch ($detailLevel) {
                    {$_ -in "minimum", "combined", "all"} { 
                        $DesiredProperties = @(
                            "DisplayName", "AssignedLicenses", "UserPrincipalName"
                            "UserType", "Id", "AccountEnabled"
                            "CreatedDateTime", "Mail", "JobTitle"
                            "Department", "CompanyName", "OfficeLocation"
                            "City", "State", "Country"
                            "OnPremisesSyncEnabled", "OnPremisesDistinguishedName", "OnPremisesLastSyncDateTime"
                            "UsageLocation", "SignInActivity", "ProxyAddresses"
                        )
                        $allTenantUsers = Get-MgUser -all -Property $DesiredProperties -ErrorAction Stop | select $DesiredProperties | ? {$_.ID -ne $null}
                    }
                    #Not Sure this will pull all the details for a user. Need to review Property variable. Might need to include all properties that expand further
                    geek {$allTenantUsers = Get-MgUser -all -ErrorAction Stop }
                }
                Write-Progress -Activity "Getting all Microsoft Graph User Data" -Completed
            }
            catch {
                # Using the Capture-ErrorHelper function to capture and log the error.
                if ($_.Exception.Message -like "*Neither tenant is B2C or tenant doesn't have premium license*") {
                    #Run if Error received is Get-MgUser : Neither tenant is B2C or tenant doesn't have premium license
                    #Status: 403 (Forbidden)
                    #ErrorCode: Authentication_RequestFromNonPremiumTenantOrB2CTenant
                    Write-Log -Type ERROR -Message "[Get-allUserDetails] An error occurred in running Get-allMGUserDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
                    $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-allMGUserDetails function. $($_.Exception.Message)"
                    $global:AllDiscoveryErrors += $ErrorObject
                    Write-Host
                    Write-Host "Caught a tenant license exception. Getting all Microsoft Graph User data without licenses and sign in activity..." -ForegroundColor Yellow -nonewline    
                    try {
                        #Fallback has bug that doesn't return all properties to help build licenses and sign in activity
                        Write-Log -Type Info -Message "[Get-allUserDetails] Attempt 2. Getting all Microsoft Graph $($detailLevel) with limited User Details" -ExportFileLocation $ExportDetails[0]
                        $allTenantUsers = Get-MgUser -all -ErrorAction Stop
                        $BasicMGDetails = $true
                    }
                    catch {
                        Write-Log -Type ERROR -Message "[Get-allUserDetails] An error occurred in running Get-allMGUserDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
                        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-allMGUserDetails function. $($_.Exception.Message)"
                        $global:AllDiscoveryErrors += $ErrorObject
                        }
                }
                else {
                    $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-allMGUserDetails function. $($_.Exception.Message)"
                    $global:AllDiscoveryErrors += $ErrorObject
                    Write-Log -Type Error -Message "[Get-allUserDetails] An error occurred in running Get-allMGUserDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
                    throw "An error occurred in running Get-allMGUserDetails function. $($_.Exception.Message)"
                    # Handle other exceptions
                }
            }

        }
    }
    #Add Additional Properties to Hash Table - Licensing, MailboxStats, OneDriveStats, ArchiveStats
    Write-Log -Type Info -Message "[Get-allUserDetails] Adding Additional Properties for $($allTenantUsers.count) Users" -ExportFileLocation $ExportDetails[0]
    try {
        $progresscounter = 0
        $totalCount = $allTenantUsers.count
        foreach ($user in $allTenantUsers) {
            try {
                $progresscounter++
                # Create Hash Table for each user
                Write-ProgressHelper -Activity "Gathering Tenant User Details" -CurrentOperation "Gathering Tenant User Details for $($user.DisplayName)" -ProgressCounter ($progresscounter) -TotalCount $totalCount -StartTime $start
                Write-Log -Type DEBUG -Message ("[Get-allUserDetails] ({0}/{1}) Creating Hash for '{2}'" -f $progressCounter, $totalCount, $user.UserPrincipalName) -ExportFileLocation $ExportDetails[0]
                # New hashtable to build upon the existing properties
                $tenantStatsHash["Users"][$user.UserPrincipalName] = [ordered]@{}

                # Populate the existing properties
                Write-Log -Type DEBUG -Message ("[Get-allUserDetails] ({0}/{1}) Create '{2}' Hash with Existing Properties" -f $progressCounter, $totalCount, $user.UserPrincipalName) -ExportFileLocation $ExportDetails[0]

                foreach ($property in $user.PSObject.Properties) {
                    $tenantStatsHash["Users"][$user.UserPrincipalName][$property.Name] = $property.Value
                }

                # Add specific additional properties - MSOL or MGGraph
                Write-Log -Type DEBUG -Message ("[Get-allUserDetails] ({0}/{1}) Adding '{2}' Specific Additional Properties (MSOL or MGGraph)" -f $progressCounter, $totalCount, $user.UserPrincipalName) -ExportFileLocation $ExportDetails[0]

                #Combine ProxyAddresses
                $combinedProxyAddresses = ($user.ProxyAddresses -replace '^[sS][mM][tT][pP]:') -join ';'
                $tenantStatsHash["Users"][$user.UserPrincipalName]["ProxyAddresses"] = $combinedProxyAddresses
                switch ($ServiceName) {
                    MSOL { 
                        # Gather Licensing Details
                        $tenantStatsHash["Users"][$user.UserPrincipalName]["AssignedLicenses"] = ($user.LicenseAssignmentDetails.accountsku.skupartnumber -join ";")
                        $tenantStatsHash["Users"][$user.UserPrincipalName]["License-DisabledArray"] = ($user.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";")
                        }
                    MGGraph {
                        if ($BasicMGDetails) {
                            Write-Log -Type DEBUG -Message ("[Get-allUserDetails] ({0}/{1}) Updating '{2}' UserType to HashTable if Basic Details" -f $progressCounter, $totalCount, $user.UserPrincipalName) -ExportFileLocation $ExportDetails[0]
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["UserType"] = if ($user.UserPrincipalName -like "*#EXT#*") { "GuestUser" } else { "User" }
                        }
                        else {
                            Write-Log -Type DEBUG -Message ("[Get-allUserDetails] ({0}/{1}) Gather '{2}' License Friendly Names" -f $progressCounter, $totalCount, $user.UserPrincipalName) -ExportFileLocation $ExportDetails[0]
                            $licensedDetails = @()
                            if ($licensedDetails =  Get-LicenseFriendlyName -AssignedLicenses $user.AssignedLicenses -LicenseHash $tenantStatsHash["LicenseSKUs"] -ErrorAction SilentlyContinue) {}
                            
                            # Create Current User Object - Default Attributes
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["UserType"]                               = $user.UserType
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["AssignedLicenses"]                       = if($licensedDetails[0]) { ($licensedDetails[0] -join ",") } else { $null }
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["License-DisabledArray"]                  = if($licensedDetails[0]) { ($licensedDetails[1] -join ",") } else { $null }
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["LastNonInteractiveSignInDateTime"]       = $user.SignInActivity.LastNonInteractiveSignInDateTime
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["LastNonInteractiveSignInRequestId"]      = $user.SignInActivity.LastNonInteractiveSignInRequestId
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["LastSignInDateTime"]                     = $user.SignInActivity.LastSignInDateTime
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["LastSignInRequestId"]                    = $user.SignInActivity.LastSignInRequestId
                        }
                    }
                }

                #Add additional properties - Defined by Level of Reporting; Only run if Mailbox Stats was ran
                if ($tenantStatsHash["AllMailboxes"]) {
                    switch ($detailLevel) {
                        {$_ -in "all", "combined"} {    
                            #Get mailbox details
                            if ($user.mail -or $user.userprincipalname -notlike "*#EXT#*") {
                                Write-Log -Type DEBUG -Message ("[Get-allUserDetails] ({0}/{1}) Gathering '{2}' Matching Details (MSOL or MGGraph). Non Guest Users" -f $progressCounter, $totalCount, $user.UserPrincipalName) -ExportFileLocation $ExportDetails[0]
                                $matchingStats = Get-MatchingStats -user $user -tenantStatsHash $tenantStatsHash
                            }
                            else {
                                Write-Log -Type DEBUG -Message ("[Get-allUserDetails] ({0}/{1}) '{2}' Guest User found" -f $progressCounter, $totalCount, $user.UserPrincipalName) -ExportFileLocation $ExportDetails[0]
                            }

                            #Add Additional Details to Current User Object
                            Write-Log -Type DEBUG -Message ("[Get-allUserDetails] ({0}/{1}) Add '{2}' Mailbox Details to HashTable for Combined/All Reports" -f $progressCounter, $totalCount, $user.UserPrincipalName) -ExportFileLocation $ExportDetails[0]
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["PrimarySmtpAddress"] = if ($matchingStats.Mailbox) { $matchingStats.Mailbox.PrimarySmtpAddress } else { $null }
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["RecipientTypeDetails"] = if ($matchingStats.Mailbox) { $matchingStats.Mailbox.RecipientTypeDetails } else { $null }
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["HiddenFromAddressListsEnabled"] = if ($matchingStats.Mailbox) { $matchingStats.Mailbox.HiddenFromAddressListsEnabled } else { $null }
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["MBXSize"] = if ($matchingStats.Mailbox) { $matchingStats.MBXStats.TotalItemSize.ToString() } else { $null }
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["MBXSize-GB"] = if ($matchingStats.Mailbox) { $matchingStats.MBXSizeGB } else { $null }
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["MBXItemCount"] = if ($matchingStats.Mailbox) { $matchingStats.MBXStats.ItemCount } else { $null }
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["Alias"] = if ($matchingStats.Mailbox) { $matchingStats.Mailbox.alias } else { $null }
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["OnMicrosoftAlias"] = if ($matchingStats.Mailbox) { $matchingStats.onmicrosoftAlias } else { $null }
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["EmailAddresses"] = if ($matchingStats.Mailbox) { ($matchingStats.Mailbox.EmailAddresses -Join ",") } else { $null }
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["OneDriveURL"] = if ($matchingStats.OneDriveData) { $matchingStats.OneDriveData.URL } else { $null }
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["OneDriveStorage"] = if ($matchingStats.OneDriveData) { $matchingStats.OneDriveData.StorageUsageCurrent } else { $null }
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["OneDriveStorage-GB"] = if ($matchingStats.OneDriveData) {$matchingStats.OneDriveData."OneDriveStorage-GB"} else { $null }
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["OneDriveLastContentModifiedDate"] = if ($matchingStats.OneDriveData) { $matchingStats.OneDriveData.LastContentModifiedDate } else { $null }
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["ArchiveStatus"] = if ($matchingStats.Mailbox) { $matchingStats.Mailbox.ArchiveStatus } else { $null }
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["ArchiveSize"] = if ($matchingStats.archiveStats) { $matchingStats.archiveStats.TotalItemSize } else { $null }
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["ArchiveSize-GB"] = if ($matchingStats.archiveStats) { $matchingStats.ArchiveStatsGB } else { $null }
                            $tenantStatsHash["Users"][$user.UserPrincipalName]["ArchiveItemCount"] = if ($matchingStats.archiveStats) { $matchingStats.archiveStats.ItemCount } else { $null }
                        }
                    }
                }
            }
            catch {
                $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in Creating MSOLUser Hash for user $($user.UserPrincipalName). $($_.Exception.Message)"
                $global:AllDiscoveryErrors += $ErrorObject
                Write-Log -Type ERROR -Message ("[Get-allUserDetails] ({0}/{1}) An error occurred in Creating MSOLUser Hash for user '{2}'. $($_.Exception.Message)" -f $progressCounter, $totalCount, $user.UserPrincipalName) -ExportFileLocation $ExportDetails[0]
            }
        }
        Write-ProgressHelper -Activity "Gathering Tenant User Details" -Completed
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type Info -Message "[Get-allUserDetails] COMPLETED: Gathering all MSOL User Details. ($($ServiceName)) in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
    }     
    catch {
        Write-Log -Type ERROR -Message "[Get-allUserDetails] An error occurred in running Get-allMSOLUserDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in Creating MSOLUser Hash. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
    }
}

#Gather all Office 365 Admins
function Get-AllOffice365Admins {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the service name')]
        [ValidateSet('MSOL','MGGraph','Azure')]
        [string[]]$ServiceName
	)
    try {
        $start = Get-Date
        $adminResults = @()
        $tenantStatsHash["Admins"] = @{}

        Write-Host "Gathering All Admins ($($ServiceName)) ..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-AllOffice365Admins] START: Gathering All Admins from Tenant - $($ServiceName)" -ExportFileLocation $ExportDetails[0]
        switch ($ServiceName) {
            MSOL {$adminRoles = Get-MsolRole | Select-Object Name, ObjectId, Description}
            MGGraph {$adminRoles = Get-MgDirectoryRole | Select DisplayName, ID, Description | ? {$_.DisplayName -ne $null}}
            Azure {$adminRoles = Get-AzureADDirectoryRole | Select DisplayName, ObjectID, Description }
        }

        $progresscounter = 0
        $totalCount = $adminRoles.count
        Write-Log -Type INFO -Message "[Get-AllOffice365Admins] Admin Roles Found: $($adminRoles.count)" -ExportFileLocation $ExportDetails[0]
        foreach ($role in $adminRoles) {
            $progresscounter++
            switch ($ServiceName) {
                MSOL {
                    $roleName = $role.Name
                }
                {$_ -in "MGGraph", "Azure"} {
                    $roleName = $role.DisplayName
                }
                default {
                    throw "Invalid service name: $ServiceName"
                }
            }
            Write-Log -Type DEBUG -Message "[Get-AllOffice365Admins] $($roleName): Gathering Admins in Role" -ExportFileLocation $ExportDetails[0]
            Write-ProgressHelper -ID 1 -Activity "Gathering Admins in Roles" -CurrentOperation "Checking Role: $($roleName)" -ProgressCounter ($progresscounter) -TotalCount $TotalCount -StartTime $start 
            switch ($ServiceName) {
                MSOL {
                    $userList = Get-MsolRoleMember -RoleObjectId $role.ObjectId
                }
                MgGraph {
                    $userList = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id | ? {$_.Id -ne $null}
                }
                Azure {
                    $userList = Get-AzureADDirectoryRoleMember -RoleObjectId $role.ObjectId
                }
                default {
                    throw "Invalid service name: $ServiceName"
                }

            }
            if ($userList) {
                $progresscounter2 = 0
                $totalCount2 = $userList.count
                Write-Log -Type INFO -Message "[Get-AllOffice365Admins] $($roleName) Users Found: $($userList.count)" -ExportFileLocation $ExportDetails[0]
                foreach ($user in $userList) {
                    switch ($ServiceName) {
                        MSOL {
                            $Name = $user.DisplayName
                            Write-Log -Type DEBUG "[Get-AllOffice365Admins] Gathering MSOL Role Details for User: $($Name)" -ExportFileLocation $ExportDetails[0]
    
                            $currentAdmin = New-Object PSObject -Property ([ordered]@{
                                "Role" = $roleName
                                "DisplayName" = $Name
                                "UserPrincipalName" = $user.EmailAddress
                                "userType" = $user.RoleMemberType
                                "objectID" = $user.ObjectId
                            })
                        }
                        MgGraph {
                            $Name = $user.additionalproperties["displayName"]
                            Write-Log -Type DEBUG "[Get-AllOffice365Admins] Gathering MGGraph Role Details for User: $($Name)" -ExportFileLocation $ExportDetails[0]
                            $currentAdmin = New-Object PSObject -Property ([ordered]@{
                                "Role" = $roleName
                                "DisplayName" = $Name
                                "UserPrincipalName" = $user.additionalproperties["userPrincipalName"]
                                "userType" = $user.additionalproperties["userType"]
                                "homepage" = $user.additionalproperties["homepage"]
                            })
                        }
                        Azure {
                            $Name = $user.DisplayName
                            Write-Log -Type DEBUG "[Get-AllOffice365Admins] Gathering Azure Role Details for User: $($Name)" -ExportFileLocation $ExportDetails[0]
                            $currentAdmin = New-Object PSObject -Property ([ordered]@{
                                "Role" = $roleName
                                "DisplayName" = $Name
                                "UserPrincipalName" = $user.UserPrincipalName
                                "objectID" = $user.ObjectId
                                "userType" = $user.UserType
                                "JobTitle" = $user.JobTitle
                                "homepage" = $user.additionalproperties["homepage"]
                            })
                        }
                    }
                    $progresscounter2 += 1
                    $progresspercentcomplete = [math]::Round((($progresscounter2 / $totalCount2)*100),2)
                    $progressStatus = "["+$progresscounter2+" / "+$totalCount2+"]"
                    Write-Progress -id 2 -Activity "Gathering Admins in Roles" -CurrentOperation "Gathering Admin Role Details: $($user.DisplayName)" -PercentComplete $progresspercentcomplete -Status $progressStatus 
    
                    $adminResults += $currentAdmin
                }
            }
            
        }
        
        Write-Log -Type INFO -Message "[Get-AllOffice365Admins] Combined Group Roles for $($groupedResults)" -ExportFileLocation $ExportDetails[0]
        #Group by DisplayName or UserPrincipalName and combine roles into a comma-separated list
        $groupedResults = $adminResults | Group-Object -Property DisplayName,UserPrincipalName
        $finalResults = $groupedResults | ForEach-Object {
            $group = $_.Group
            $roles = ($group.Role -join ', ')
            # Use the properties of the first user in each group, but replace the Role with the combined roles
            $group[0] | Add-Member -MemberType NoteProperty -Name 'Role' -Value $roles -Force
            $group[0] | Add-Member -MemberType NoteProperty -Name 'RolesAssigned' -Value (($group.Role | measure).count) -Force
            $group[0]
        }

        foreach ($result in $finalResults) {
            $tenantStatsHash["Admins"][$result.DisplayName] = $result
        }
        

    }
    catch {
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-AllOffice365Admins function. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
        Write-Log -Type ERROR -Message "[Get-AllOffice365Admins] An error occurred in running Get-AllOffice365Admins function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
    }
    finally {
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-ProgressHelper -ID 1 -Activity "Gathering Admins in Roles" -Completed
        Write-Progress -id 2 -Activity "Checking Admin Role Details: $($user.DisplayName)" -Completed
        Write-Log -Type INFO -Message "[Get-AllOffice365Admins] COMPLETED: Gathering All Admins. ($($ServiceName)) in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
    }
}

# Get all Office 365 Domains
function Get-AllOffice365Domains {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the service name')]
        [ValidateSet('MSOL','MGGraph','Azure')]
        [string[]]$ServiceName
	)
    try {
        $start = Get-Date
        Write-Host "Gathering All Domains ..." -ForegroundColor Cyan -nonewline
        # Get all the domains
        Write-Log -Type INFO -Message "[Get-AllOffice365Domains] START: Gathering All $($ServiceName) Domains from tenant" -ExportFileLocation $ExportDetails[0]
        switch ($ServiceName) {
            MSOL {$domains = Get-MsolDomain }
            MGGraph {$domains = Get-MgDomain | ? {$_.ID -ne $null}}
            Azure {$domains = Get-AzureADDomain }
        }
        #Gather Exchange Online Domain Details
        try {
            Write-Log -Type INFO -Message "[Get-AllOffice365Domains] Gathering Exchange Online Domain Details" -ExportFileLocation $ExportDetails[0]
            $acceptedDomains = Get-AcceptedDomain
            $remoteDomains = Get-RemoteDomain | select Identity, DomainName, IsInternal, TargetDeliveryDomain, AllowedOOFType, AutoReplyEnabled, AutoForwardEnabled, DeliveryReportEnabled, NDREnabled, MeetingForwardNotificationEnabled, ContentType, TNEFEnabled, TrustedMailOutboundEnabled, TrustedMailInboundEnabled
            $recipients = Get-EXORecipient -ResultSize Unlimited
            $exchangeOnline = $True
        }
        catch {  
            if ($_.Exception.Message -like "*The term 'Get-AcceptedDomain' is not recognized*") {
                Write-Log -Type ERROR -Message "[Get-AllOffice365Domains] Exchange Online not connected" -ExportFileLocation $ExportDetails[0]
                Write-Log -Type ERROR -Message "[Get-AllOffice365Domains] $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
                $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-AllOffice365Domains function. Exception: Not Properly Connected to Exchange Online Tenant. Check Permissions and reconnect to Exchange Online. Skipping Exchange portion"
                $global:AllDiscoveryErrors += $ErrorObject
                $exchangeOnline = $False
            }
            else {
                Write-Log -Type Error -Message "[Get-AllOffice365Domains] Exchange Online not connected for other reasons. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
                Write-Error $_.Exception.Message
            }
        }
        # Get all Remote and Accepted domains
        $progresscounter = 0
        $totalCount = $domains.count
        Write-Log  -Type INFO -Message "[Get-AllOffice365Domains] $($domains.count) domains found in tenant" -ExportFileLocation $ExportDetails[0]

        # Prepare the results array
        $tenantStatsHash["Domains"] = @{}
        $tenantStatsHash["RemoteDomains"] = @{}

        foreach ($domain in $domains) {
            $progresscounter++
            # Get the DNS records
            switch ($ServiceName) {
                MSOL {
                    $domainName = $domain.Name
                    $domainVerified = $domain.Status
                    $AuthenticationType = $domain.Authentication
                }
                {$_ -in "MGGraph", "Azure"} {
                    $domainName = $domain.ID
                    $domainVerified = $domain.IsVerified       
                    $AuthenticationType = $domain.AuthenticationType
                }
            }
            Write-Log -Type DEBUG -Message "[Get-AllOffice365Domains] Gathering $($domainName) domain details" -ExportFileLocation $ExportDetails[0]
            Write-ProgressHelper -ID 1 -Activity "Gathering Domain Details" -CurrentOperation "Gathering $($domainName)" -ProgressCounter ($progresscounter) -TotalCount $TotalCount -StartTime $start

            #Gather DNS Records from 1.1.1.1
            Write-Log -Type DEBUG -Message "[Get-AllOffice365Domains] Gathering DNS Records for $($domainName)" -ExportFileLocation $ExportDetails[0]
            $Default = $domain.IsDefault
            $aRecords = Resolve-DnsName -Name $domainName -Server 1.1.1.1 -Type A -ErrorAction SilentlyContinue
            $mxRecords = Resolve-DnsName -Name $domainName -Server 1.1.1.1 -Type MX -ErrorAction SilentlyContinue
            $NSRecords = Resolve-DnsName -Name $domainName -Server 1.1.1.1 -Type NS -ErrorAction SilentlyContinue
            Write-Log -Type DEBUG -Message "[Get-AllOffice365Domains] Completed Gathering DNS Records for $($domainName)"

            #Check for Exchange Online Dependencies
            if ($exchangeOnline -eq $true) {
                Write-Log -Type DEBUG -Message "[Get-AllOffice365Domains] Gathering Exchange Online Domain Details for $($domainName)" -ExportFileLocation $ExportDetails[0]
                $DomainType = $acceptedDomains | ?{$_.DomainName -eq $domainName} | Select-Object -ExpandProperty "DomainType"
                #Domain Dependencies
                $AllRecipientsWithPrimarySMTP = $recipients | Where-Object { $_.PrimarySmtpAddress -like "*@$($domainName)" }
                $AllRecipientsWithAlias = $recipients | Where-Object { $_.EmailAddresses -like "*@$($domainName)" }
                $AllRecipientsAliasOnly = ($AllRecipientsWithPrimarySMTP| measure).count - ($AllRecipientsWithAlias| measure).count
            }
            else {
                Write-Log -Type Error -Message "[Get-AllOffice365Domains] Skipping Exchange Online Domain Details for $($domainName)" -ExportFileLocation $ExportDetails[0]
                $DomainType = $null
                $AllRecipientsWithPrimarySMTP = $null
                $AllRecipientsWithAlias = $null
                $AllRecipientsAliasOnly = $null
            }
            # Add to the results array
            Write-Log -Type INFO -Message "[Get-AllOffice365Domains] Adding $($domainName) to results array" -ExportFileLocation $ExportDetails[0]
            $currentDomain = @()
            $currentDomain = New-Object PSObject -Property ([ordered]@{
                Domain = $domainName
                Verified = $domainVerified
                AuthenticationType = $AuthenticationType
                DomainType = $DomainType
                IsDefault = $Default
                NSRecords = if ($NSRecords) { ($NSRecords.NameHost -join ","| Out-String).Trim() } else { $null }
                ARecords = if ($aRecords) { ($aRecords.IPAddress -join "," | Out-String).Trim() } else { $null }
                MXRecords = if ($mxRecords) { ($mxRecords.NameExchange -join "," | Out-String).Trim()} else { $null }
                Office365MailExchanger = if (($mxRecords.NameExchange | Out-String).Trim() -like "*protection.outlook.com") {$true} else { $False }
                AllRecipientsWithPrimarySMTP = if ($AllRecipientsWithPrimarySMTP) {($AllRecipientsWithPrimarySMTP| measure).count} else { $null }
                AllRecipientsWithAlias = if ($AllRecipientsWithAlias) {($AllRecipientsWithAlias| measure).count} else { $null }
                AllRecipientsAliasOnly = if ($AllRecipientsAliasOnly) {$AllRecipientsAliasOnly} else { $null }
            })
            $tenantStatsHash["Domains"][$domainName] = $currentDomain
        }
        foreach ($domain in $remoteDomains) {
            # Add to the results Hash Table
            $tenantStatsHash["RemoteDomains"][$domain.Identity] = $domain
        }
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-AllOffice365Domains] $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-AllOffice365Domains function. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
    }
    finally {
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-ProgressHelper -ID 1 -Activity "Gathering Domain Details" -Completed
        Write-Log -Type INFO -Message "[Get-AllOffice365Domains] COMPLETED: Gathering All Domain Details. ($($ServiceName))" -ExportFileLocation $ExportDetails[0]
    }
}
# ----------------------------------
# Combine Mailbox Details Specific Functions
# ----------------------------------

#Consolidate Reporting for each user
function Combine-AllMailboxStats {
    try {
        $tenantStatsHash["AllMailboxFullDetails"] = @{}
        $progresscounter = 0
        $start = Get-Date
        $totalCount = $tenantStatsHash["AllMailboxes"].count
        Write-Host "Combining all AllMailboxFullDetails ..." -ForegroundColor Cyan -nonewline

        foreach ($mailbox in $tenantStatsHash["AllMailboxes"].values) {
            try {
                #progress bar
                $progresscounter++
                Write-ProgressHelper -Activity "Combining Mailbox Details" -CurrentOperation "Gathering $($mailbox.RecipientTypeDetails) Mailbox Details for $($mailbox.DisplayName)" -ProgressCounter ($progresscounter) -TotalCount $TotalCount -StartTime $start
                Write-Log -Type DEBUG -Message ("[Combine-AllMailboxStats] ({0}/{1}) Combining Mailbox '{2}' Details" -f $progressCounter, $totalCount, $mailbox.PrimarySMTPAddress) -ExportFileLocation $ExportDetails[0]            
                #null values
                $oneDriveData = $null
                $sharePointSiteData = $null
                $mgUser = $null
                $mbxStats = $null
                $MBXSizeGB = $null
                $ArchiveStats = $null

                #Pull MailboxStats, MGUserDetails, Licensing, and Disabled Service Plans
                #*******************************************************************************************************************
                $EmailAddresses = $mailbox | Select-Object -ExpandProperty EmailAddresses
                try {
                    $onmicrosoftAlias = ($EmailAddresses | Where-Object { $_ -like "SMTP:*@*.onmicrosoft.com" -or $_ -like "smtp:*@*.onmicrosoft.com" } | Select-Object -First 1).Replace("SMTP:", "").Replace("smtp:", "")
                } catch {
                    $onmicrosoftAlias = $null
                }
                
                #If $mailbox represents a User object set $mbxStats data to pull from PrimaryMailboxStats which contains mailbox data
                if ($tenantStatsHash["PrimaryMailboxStats"] -and $mailbox.ExchangeGuid) {
                    Write-Log -Type DEBUG -Message ("[Combine-AllMailboxStats] ({0}/{1}) Gathering '{2}' Primary Mailbox Statistics" -f $progressCounter, $totalCount, $mailbox.PrimarySMTPAddress) -ExportFileLocation $ExportDetails[0]
                    $mbxStats = $tenantStatsHash["PrimaryMailboxStats"][$mailbox.ExchangeGuid.ToString()]
                    if($mbxStats.TotalItemSize) {
                        $MBXSizeGB = [math]::Round(($MBXStats.TotalItemSize.ToString() -replace "(.*\()|,| [a-z]*\)","")/1GB,3)
                    } else {
                        $MBXSizeGB = 0
                    }
                }          
                
                #If $mailbox represents a User object set $mgUser data to pull from Users Hash which contains User data
                if ($tenantStatsHash["OneDrives"] -and $tenantStatsHash["Users"]) {
                    Write-Log -Type DEBUG -Message ("[Combine-AllMailboxStats] ({0}/{1}) Gather {2} Mailbox Statistics for {3}" -f $progressCounter, $totalCount, $mailbox.RecipientTypeDetails, $mailbox.PrimarySMTPAddress) -ExportFileLocation $ExportDetails[0]
                    if($mailbox.UserPrincipalName -and $mailbox.RecipientTypeDetails -ne "GroupMailbox") {
                        $mgUser = $tenantStatsHash["Users"][$mailbox.UserPrincipalName.ToString()]
                        if($oneDriveData =  $tenantStatsHash["OneDrives"][$mailbox.UserPrincipalName]) {}
                        else {$oneDriveData = $null}
    
                        #Gather SigninActivity
                        $signinActivity = $mgUser.SignInActivity
                    }
                    elseif($mailbox.RecipientTypeDetails -eq "GroupMailbox" -and $tenantStatsHash["GroupMailboxes"]) { 
                        Write-Log -Type DEBUG -Message ("[Combine-AllMailboxStats] ({0}/{1}) Gather {2} Mailbox Statistics for {3}" -f $progressCounter, $totalCount, $mailbox.RecipientTypeDetails, $mailbox.PrimarySMTPAddress) -ExportFileLocation $ExportDetails[0]                      
                        $mgUser = $tenantStatsHash["GroupMailboxes"][$mailbox.ExchangeGuid.ToString()]
                        $unifiedGroupData = $tenantStatsHash["UnifiedGroups"][$mailbox.ExchangeGuid.ToString()]
    
                        if($unifiedGroupData.SharePointSiteUrl) {
                            $sharePointSiteData = $tenantStatsHash["SharePointSites"][($unifiedGroupData.SharePointSiteUrl)]
                        }
                        else {$sharePointSiteData = $null}
                    }
                    else {
                        $unifiedGroupData = $null
                        $sharePointSiteData = $null
                    }
                }
                
                # Create Hash Table to add to Report Dataset
                #*******************************************************************************************************************
                $currentuser = [ordered]@{
                    #User Information
                    "DisplayName" = $mailbox.DisplayName
                    "RecipientTypeDetails" = $mailbox.RecipientTypeDetails
                    "UserPrincipalName" = $mailbox.userprincipalname
                    "Department" = $mgUser.Department
                    "IsLicensed" = ($mgUser.AssignedLicenses.count -gt 0)
                    "Licenses" = $mgUser.AssignedLicenses
                    "License-DisabledArray" = $mgUser.DisabledPlans
                    "AccountEnabled" = $mgUser.AccountEnabled
                    "IsInactiveMailbox" = $mailbox.IsInactiveMailbox
                    "WhenSoftDeleted" = $mailbox.WhenSoftDeleted
                    #Login Activity
                    "LastSignInDateTime" = $signinActivity.LastSignInDateTime
                    "LastSignInRequestId" = $signinActivity.LastSignInRequestId
                    "LastNonInteractiveSignInDateTime" = $signinActivity.LastNonInteractiveSignInDateTime
                    "LastNonInteractiveSignInRequestId" = $signinActivity.LastNonInteractiveSignInRequestId
                
                    "WhenCreated" = $mailbox.WhenCreated
                    "LastLogonTime" = $mbxStats.LastLogonTime
                    #mailbox information
                    "PrimarySmtpAddress" = $mailbox.PrimarySmtpAddress
                    "HiddenFromAddressListsEnabled" = $mailbox.HiddenFromAddressListsEnabled
                    "MBXSize" = $MBXStats.TotalItemSize
                    "MBXSize-GB" = $MBXSizeGB
                    "MBXItemCount" = $MBXStats.ItemCount
                    "Alias" = $mailbox.alias
                    "OnMicrosoftAlias" = $onmicrosoftAlias
                    "EmailAddresses" = ($EmailAddresses -join ";")
                    "DeliverToMailboxAndForward" = $mailbox.DeliverToMailboxAndForward
                    "ForwardingAddress" = $mailbox.ForwardingAddress
                    "ForwardingSmtpAddress" = $mailbox.ForwardingSmtpAddress
                    "LitigationHoldEnabled" = $mailbox.LitigationHoldEnabled
                    "LitigationHoldDuration" = $mailbox.LitigationHoldDuration
                    "InPlaceHolds" = $mailbox.InPlaceHolds -join ";"
                    "ArchiveStatus" = $mailbox.ArchiveStatus
                    "RetentionPolicy" = $mailbox.RetentionPolicy
                }

                # Archive Mailbox Check
                #*******************************************************************************************************************
                if ($tenantStatsHash["ArchiveMailboxStats"] -and $mailbox.ArchiveStatus -ne "None" -and $null -ne $mailbox.ArchiveStatus) {
                    Write-Log -Type DEBUG -Message ("[Combine-AllMailboxStats] ({0}/{1}) Gather {2} Mailbox Archive Statistics for {3}" -f $progressCounter, $totalCount, $mailbox.RecipientTypeDetails, $mailbox.PrimarySMTPAddress) -ExportFileLocation $ExportDetails[0]            
                    $archiveStats = $tenantStatsHash["ArchiveMailboxStats"][$mailbox.ArchiveGuid.ToString()]
                    $currentuser["ArchiveSize"] = $ArchiveStats.TotalItemSize.Value

                    if($ArchiveStats.TotalItemSize) {
                        $currentuser["ArchiveSize-GB"] = [math]::Round(($ArchiveStats.TotalItemSize.ToString() -replace "(.*\()|,| [a-z]*\)","")/1GB,3)
                    }
                    else {
                        $currentuser["ArchiveSize-GB"] = $null 
                    }

                    $currentuser["ArchiveItemCount"] = $ArchiveStats.ItemCount
                }

                else {
                    $currentuser["ArchiveSize"] = $null
                    $currentuser["ArchiveSize-GB"] = $null
                    $currentuser["ArchiveItemCount"] = $null
                }

                # Get SharePoint Online/OneDrive Check
                #*******************************************************************************************************************
                #Get SharePoint/OneDrive Data
                if($OneDriveData) {
                    Write-Log -Type DEBUG -Message ("[Combine-AllMailboxStats] ({0}/{1}) Gather {2} OneDrive Statistics for {3}" -f $progressCounter, $totalCount, $mailbox.RecipientTypeDetails, $mailbox.PrimarySMTPAddress) -ExportFileLocation $ExportDetails[0]            
                    $currentuser["OneDriveURL"] = $OneDriveData.URL
                    $currentuser["OneDriveStorage"] = $OneDriveData.StorageUsageCurrent
                    $currentuser["OneDriveStorage-GB"] = [math]::Round($OneDriveData.StorageUsageCurrent / 1024, 3)
                    $currentuser["OneDriveLastContentModifiedDate"] = $OneDriveData.LastContentModifiedDate
                    $currentuser["SharePointURL"] = $null
                    $currentuser["SharePointStorage-GB"] = $null
                    $currentuser["SharePointLastContentModifiedDate"] = $null
                }
                #Group Mailbox Associated SharePoint Site mapping
                elseif($sharePointSiteData) {
                    Write-Log -Type DEBUG -Message ("[Combine-AllMailboxStats] ({0}/{1}) Gather {2} OneDrive Statistics for {3}" -f $progressCounter, $totalCount, $mailbox.RecipientTypeDetails, $mailbox.PrimarySMTPAddress) -ExportFileLocation $ExportDetails[0]            
                    $currentuser["OneDriveURL"] = $null
                    $currentuser["OneDriveStorage-GB"] = $null
                    $currentuser["OneDriveLastContentModifiedDate"] = $null 
                    $currentuser["SharePointURL"] = $sharePointSiteData.URL
                    $currentuser["SharePointStorage"] = $sharePointSiteData.StorageUsageCurrent
                    $currentuser["SharePointStorage-GB"] = [math]::Round($sharePointSiteData.StorageUsageCurrent / 1024, 3)
                    $currentuser["SharePointLastContentModifiedDate"] = $sharePointSiteData.LastContentModifiedDate
                }
                else {
                    $currentuser["OneDriveURL"] = $null
                    $currentuser["OneDriveStorage-GB"] = $null
                    $currentuser["OneDriveLastContentModifiedDate"] = $null 
                    $currentuser["SharePointURL"] = $null
                    $currentuser["SharePointStorage-GB"] = $null
                    $currentuser["SharePointLastContentModifiedDate"] = $null
                }
            }
            catch {
                $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in combining Mailbox Details for $($mailbox.PrimarySMTPAddress). $($_.Exception.Message)"
                $global:AllDiscoveryErrors += $ErrorObject
                Write-Log -Type ERROR -Message "[Combine-AllMailboxStats] An error occurred in combining Mailbox Details for $($mailbox.PrimarySMTPAddress). $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
            }
            finally {
                #Combine all the data into one hash table
                #*******************************************************************************************************************
                $tenantStatsHash["AllMailboxFullDetails"][$mailbox.PrimarySMTPAddress] = $currentuser
            }

        }
    }
    catch {
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Combine-AllMailboxStats function. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
        Write-Log -Type ERROR -Message "[Combine-AllMailboxStats] An error occurred in running Combine-AllMailboxStats function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
    }
    finally {
        Write-ProgressHelper -Activity "Combining Mailbox Details" -Completed
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Combine-AllMailboxStats] COMPLETED: Combining all AllMailboxFullDetails in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
    }
    #return $tenantStatsHash
}

# ----------------------------------
# Azure Report Details Specific Functions
# ----------------------------------

# Device Report Details
function Get-AllDevicesReport {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory=$True,HelpMessage='Provide the service name')]
        [ValidateSet('MGGraph','Azure')]
        [string[]]$ServiceName 
    )
    $start = Get-Date
    $tenantStatsHash["DeviceDetails"] = @{}

    try {
        switch ($ServiceName) {
            Azure { 
                $devices = Get-AzureADDevice -All $true -ErrorAction Stop
                # Filter for Desired Attributes
                $DesiredProperties = @(
                    "DisplayName", "AccountEnabled", "DeviceOSType", "DeviceOSVersion", "ObjectType", "DeviceId", "ObjectID"
                    "ApproximateLastLogonTimeStamp", "DeviceTrustType", "DirSyncEnabled", "LastDirSyncTime"
                    "IsCompliant", "IsManaged", "ProfileType"          
                )
            }
            MGGraph {
                $devices = Get-MgDevice -All -ErrorAction Stop | ? {$_.ID -ne $null}
                $DesiredProperties = @(
                    "DisplayName", "AccountEnabled", "OperatingSystem", "OperatingSystemVersion", "DeviceId", "ID"
                    "ApproximateLastSignInDateTime", "TrustType", "DirSyncEnabled", "LastDirSyncTime"
                    "IsCompliant", "IsManaged", "ProfileType"
                    )
            }
        }
        Write-Host "Getting Device Details ($($ServiceName)) ..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-AllDevicesReport] START: Gathering all $($ServiceName) Device with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]

        # Filter for Desired Attributes
        switch ($detailLevel) {
            {$_ -in "minimum", "combined", "all"} { 
                $devices = $devices | Select $DesiredProperties
            }
            geek {$devices = $devices}
        }

        # Add Additional Properties - MDM Solution, Join Type, Stale Device
        $progresscounter = 0
        $totalCount = $devices.count
        foreach ($device in $devices) {
            $progresscounter++
            Write-ProgressHelper -Activity "Processing all $($ServiceName) Found Devices" -CurrentOperation "Updating Device Details for $($device.DisplayName)" -ProgressCounter ($progresscounter) -TotalCount $totalCount -StartTime $start
            try {
                #Add MDM Solution
                if ($device.ID) {
                    $device | Add-Member -MemberType NoteProperty -Name "ObjectID" -Value $device.ID -Force
                }
                #Add MDM Solution
                    if ($device.IsManaged -eq $true) {
                        $device | Add-Member -MemberType NoteProperty -Name "MDMSolution" -Value "Intune or SCCM" -Force
                    } else {
                        $device | Add-Member -MemberType NoteProperty -Name "MDMSolution" -Value "Not Managed" -Force
                    }
                Write-Log -Type DEBUG -Message ("[Get-AllDevicesReport] ({0}/{1}) Adding Device Details for '{2}': Added {3} MDM Solution value" -f $progressCounter, $devices.count, $device.DisplayName, $device.MDMSolution) -ExportFileLocation $ExportDetails[0]
        
                #Add Join Type
                switch ($device.DeviceTrustType) {
                    AzureAD {
                        $device | Add-Member -MemberType NoteProperty -Name "DeviceJoinType" -Value "Azure AD joined" -Force
                    }
                    WorkPlace {
                        $device | Add-Member -MemberType NoteProperty -Name "DeviceJoinType" -Value "Azure AD registered" -Force
                    }
                    ServerAd {
                        $device | Add-Member -MemberType NoteProperty -Name "DeviceJoinType" -Value "Hybrid Azure AD joined" -Force
                    }
                    Default {
                        $device | Add-Member -MemberType NoteProperty -Name "DeviceJoinType" -Value "Unknown" -Force
                    }
                }
                Write-Log -Type DEBUG -Message ("[Get-AllDevicesReport] ({0}/{1}) Adding Device Details for '{2}': Added {3} Join Type value" -f $progressCounter, $devices.count, $device.DisplayName, $device.MDMSolution) -ExportFileLocation $ExportDetails[0]
        
                #Add if Stale Device (Approximate login time older than 6 months)
                switch ($ServiceName) {
                    Azure {
                        $lastLogin = $device.ApproximateLastLogonTimeStamp
                    }
                    MGGraph {
                        $lastLogin = $device.ApproximateLastSignInDateTime
                    }
                }                
                if ($lastLogin) {
                    $sixMonthsAgo = (Get-Date).AddMonths(-6)
                    $timeSinceLastLogon = $sixMonthsAgo - $lastLogin
                    $device | Add-Member -MemberType NoteProperty -Name "DaysDeviceInactiveFrom6MonthsAgo" -Value $timeSinceLastLogon.Days -Force
                    if ($lastLogin -le $sixMonthsAgo) {
                        $device | Add-Member -MemberType NoteProperty -Name "DeviceStale" -Value $true -Force
                    } else {
                        $device | Add-Member -MemberType NoteProperty -Name "DeviceStale" -Value $false -Force
                    }
                }
                elseif ($lastLogin) {
                    $device | Add-Member -MemberType NoteProperty -Name "DaysDeviceInactiveFrom6MonthsAgo" -Value "N/A" -Force
                    $device | Add-Member -MemberType NoteProperty -Name "DeviceStale" -Value "N/A" -Force
                }
                Write-Log -Type DEBUG -Message ("[Get-AllDevicesReport] ({0}/{1}) Adding Device Details for '{2}': Added {3} Device Stale value" -f $progressCounter, $devices.count, $device.DisplayName, $device.MDMSolution) -ExportFileLocation $ExportDetails[0]
                
                #Add Compliant State
                if ($null -eq $device.IsCompliant) {
                    $device | Add-Member -MemberType NoteProperty -Name "IsCompliant" -Value "NotApplied" -Force
                }
                Write-Log -Type DEBUG -Message ("[Get-AllDevicesReport] ({0}/{1}) Adding Device Details for '{2}': Added {3} Device Compliant value" -f $progressCounter, $devices.count, $device.DisplayName, $device.MDMSolution) -ExportFileLocation $ExportDetails[0]

            }
            catch {
                $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "($($progressCounter/$devices.count)) Failed to add Device details for $($device.DisplayName). $($_.Exception.Message)"
                $global:AllDiscoveryErrors += $ErrorObject
                Write-Log -Type Error -Message ("[Get-AllDevicesReport] ({0}/{1}) Failed to add Device Details for '{2}'. $($_.Exception.Message)" -f $progressCounter, $devices.count, $device.DisplayName) -ExportFileLocation $ExportDetails[0]
            }
            finally {
                #Add to Hash Table
                $tenantStatsHash["DeviceDetails"][$device.ObjectID] = $device
            }
        }
    }
    catch {
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-AllAzureDeviceReport function. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
        Write-Log -Type ERROR -Message "[Get-AllDevicesReport] An error occurred in running Get-AllAzureDeviceReport function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
    }
    finally {
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-ProgressHelper -Activity "Processing all Azure Found Devices" -Completed
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-AllDevicesReport] COMPLETED: Gathering all Azure Device with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]
    }   
}

# Function to Retrieve Conditional Access Policies from Azure AD or Microsoft Graph
function Get-ConditionalAccessPoliciesReport {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory=$True,HelpMessage='Provide the service name')]
        [ValidateSet('MGGraph','AzureAD')]
        [string]$ServiceName 
    )

    function Get-PolicyItemCounts {
        param (
            [object]$policyCondition
        )
        if ($policyCondition -eq "All") {
            return "All"
        } else {
            return ($policyCondition | Measure-Object).Count
        }
    }

    $start = Get-Date
    $tenantStatsHash["ConditionalAccessPolicies"] = @{}
    try {
        Write-Host "Getting Azure Conditional Access Policies Details ..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-ConditionalAccessPoliciesReport] START: Gathering all Azure Conditional Access Policies  with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]

        # Get Conditional Access Policies based on the specified service name
        switch ($ServiceName) {
            Azure { 
                $conditionalAccessPolicies = Get-AzureADMSConditionalAccessPolicy -ErrorAction Stop
            }
            MGGraph {
                $conditionalAccessPolicies = Get-MgConditionalAccessPolicy -ErrorAction Stop
            }
        }

        # Filter for Desired Attributes
        switch ($detailLevel) {
            {$_ -in "minimum", "combined", "all"} { 
                $DesiredProperties = @(
                    "Id", "DisplayName", "CreatedDateTime", "ModifiedDateTime", "State",
                    "Conditions", "GrantControls"
                )
                $conditionalAccessPolicies = $conditionalAccessPolicies | Select-Object $DesiredProperties
            }
        }

        # Add Additional Properties - MDM Solution, Join Type, Stale Device
        $progresscounter = 0
        $totalCount = $conditionalAccessPolicies.Count
        foreach ($policy in $conditionalAccessPolicies) {
            $progresscounter++
            Write-ProgressHelper -Activity "Processing all Conditional Access Policies" -CurrentOperation "Expanding Policy Details for $($policy.DisplayName)" -ProgressCounter ($progresscounter) -TotalCount $totalCount -StartTime $start
            Write-Log -Type INFO -Message "[Get-ConditionalAccessPoliciesReport] Expanding Conditional Access Policies for $($policy.DisplayName)" -ExportFileLocation $ExportDetails[0]

            # Create a hashtable to hold the policy information
            $policyDetailsHash = [ordered]@{
                PolicyID                        = $policy.Id
                PolicyName                      = $policy.DisplayName
                ModifiedDateTime                = $policy.ModifiedDateTime
                CreatedDateTime                 = $policy.CreatedDateTime
                State                           = $policy.State
                ClientAppTypes                  = $policy.Conditions.ClientAppTypes -join ","
                GrantControls                   = $policy.GrantControls.BuiltInControls -join ","
                IncludedUsersCount              = Get-PolicyItemCounts -policyCondition $policy.Conditions.Users.IncludeUsers
                ExcludeUsersCount               = Get-PolicyItemCounts -policyCondition $policy.Conditions.Users.ExcludeUsers
                IncludedGroupsCount             = Get-PolicyItemCounts -policyCondition $policy.Conditions.Users.IncludeGroups
                ExcludedGroupsCount             = Get-PolicyItemCounts -policyCondition $policy.Conditions.Users.ExcludeGroups
                IncludedRolesCount              = Get-PolicyItemCounts -policyCondition $policy.Conditions.Users.IncludeRoles
                ExcludedRolesCount              = Get-PolicyItemCounts -policyCondition $policy.Conditions.Users.ExcludeRoles
                IncludedPlatforms               = $policy.Conditions.Platforms.IncludePlatforms -join ","
                ExcludePlatforms                = $policy.Conditions.Platforms.ExcludePlatforms -join ","
                IncludedApplicationsCount       = Get-PolicyItemCounts -policyCondition $policy.Conditions.Applications.IncludeApplications
                ExcludedApplicationsCount       = Get-PolicyItemCounts -policyCondition $policy.Conditions.Applications.ExcludeApplications
                IncludedLocations               = Get-PolicyItemCounts -policyCondition $policy.Conditions.Locations.IncludeLocations
                ExcludeLocations                = Get-PolicyItemCounts -policyCondition $policy.Conditions.Locations.ExcludeLocations             
                UserRiskLevels                  = $policy.Conditions.UserRiskLevels -join ","
                SignInRiskLevels                = $policy.Conditions.SignInRiskLevels -join ","
            }

            # Add the hashtable to the array
            $tenantStatsHash["ConditionalAccessPolicies"][$policy.DisplayName] = $policyDetailsHash
        }
        Write-ProgressHelper -Activity "Processing all Conditional Access Policies"  -Completed

    }
    catch {
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-ConditionalAccessPoliciesReport function. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
        Write-Log -Type ERROR -Message "[Get-ConditionalAccessPoliciesReport] An error occurred in running Get-ConditionalAccessPoliciesReport function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
    }
    finally {
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-ConditionalAccessPoliciesReport] COMPLETED: Gathering all Azure Compliance Policies with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]
    }   
}

# Secure Score Report Details
function Get-SecuritySecureScoreReport {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory=$True,HelpMessage='Provide the service name')]
        [ValidateSet('MGGraph','AzureAD')]
        [string[]]$ServiceName,
        [Parameter(Mandatory=$false,HelpMessage='Should It Only Pull the Most Recent?')]
        [Switch]$MostRecent 
    )
    $start = Get-Date
    $tenantStatsHash["SecuritySecureScore"] = @{}

    try {
        switch ($ServiceName) {
            Azure {
                if ($MostRecent) {
                    # Assuming AzureAD equivalent command for most recent secure score
                    $secureScore = Get-AzureADSecuritySecureScore -Top 1 -ErrorAction Stop | Where-Object {$_.ID -ne $null}
                } else {
                    # Assuming AzureAD equivalent command for all secure scores
                    $secureScore = Get-AzureADSecuritySecureScore -All -ErrorAction Stop | Where-Object {$_.ID -ne $null}
                }
            }
            MGGraph {
                if ($MostRecent) {
                    $secureScore = Get-MgSecuritySecureScore -Top 1 -ErrorAction Stop | Where-Object {$_.ID -ne $null}
                } else {
                    $secureScore = Get-MgSecuritySecureScore -All -ErrorAction Stop | Where-Object {$_.ID -ne $null}
                }
            }
        }

        Write-Host "Getting $($ServiceName) Security Score Details ..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-SecuritySecureScoreReport] START: Gathering all $($ServiceName) Security Score details" -ExportFileLocation $ExportDetails[0]
        # Add Additional Properties
        $progresscounter = 0
        $totalCount = $secureScore.count
        foreach ($score in $secureScore) {
            $progresscounter++
            Write-ProgressHelper -Activity "Processing all $($ServiceName) Security Score Details" -CurrentOperation "Gathering Score Details for $($score.ID)" -ProgressCounter ($progresscounter) -TotalCount $totalCount -StartTime $start
            Write-Log -Type INFO -Message "[Get-SecuritySecureScoreReport] Gathering Score Details for $($score.ID)" -ExportFileLocation $ExportDetails[0]
            #Create Current Security Score Hash Table
            $currentSecurityScores = [ordered]@{
                ID = $score.ID
                CreatedDateTime = $score.CreatedDateTime
                # Add Current Security Score Percentage
                SecurityScorePercentage = ("{0:F2}" -f (($score.currentScore / $score.maxScore) * 100))
                CurrentScore = $score.CurrentScore
                MaxScore = $score.MaxScore
                LicensedUserCount = $score.LicensedUserCount
                # Add Enabled Services
                EnabledServicesCount = ($score.EnabledServices | Measure-Object).Count
                EnabledServices = ($score.EnabledServices -join ",")
                VendorProvider = $score.VendorInformation.Provider
                VendorName = $score.VendorInformation.Vendor
            }
            
            # Add Comparitive Scores - TotalSeats, AllTenants
            Write-Log -Type INFO -Message "[Get-SecuritySecureScoreReport] Gathering Score Details for $($score.ID): Add Comparitive Scores" -ExportFileLocation $ExportDetails[0]
            $TotalSeatsScoreFullDetails = $score.AverageComparativeScores | Where-Object {$_.Basis -eq "TotalSeats"}
            $currentSecurityScores["SimilarSeatSizeRangeLowerValue"] = $TotalSeatsScoreFullDetails["SeatSizeRangeLowerValue"]
            $currentSecurityScores["SimilarSeatSizeRangeUpperValue"] = $TotalSeatsScoreFullDetails["SeatSizeRangeUpperValue"]
            foreach ($ComparisonScore in $score.AverageComparativeScores) {
                if ($ComparisonScore.Basis -eq "TotalSeats") {
                    $ComparisonBasisName = "SimilarSizeOrg"
                } elseif ($ComparisonScore.Basis -eq "AllTenants") {
                    $ComparisonBasisName = "AllTenants"
                }
                Write-Log -Type DEBUG -Message "[Get-SecuritySecureScoreReport] Gathering $($ComparisonBasisName) ComparisonScore Details Details for $($score.ID)" -ExportFileLocation $ExportDetails[0]
                $currentSecurityScores["$($ComparisonBasisName)_AverageComparativeScore"] = $ComparisonScore.AverageScore 
                if ($detailLevel -in "all", "geek") {
                    foreach ($individualScore in $ComparisonScore.AdditionalProperties.keys) {
                        Write-Log -Type DEBUG -Message "[Get-SecuritySecureScoreReport] Gathering $($ComparisonScore.Basis):$($individualScore) Score Details for $($score.ID): Add Comparitive Scores" -ExportFileLocation $ExportDetails[0]

                        $currentSecurityScores["$($ComparisonBasisName)_$($individualScore)"] = $ComparisonScore.AdditionalProperties[$individualScore]
                    }
                }
            }

            
            #Add to Hash Table
            Write-Log -Type INFO -Message "[Get-SecuritySecureScoreReport] Gathering Score Details for $($score.ID): Add Score to Tenant Stats Hash Table" -ExportFileLocation $ExportDetails[0]
            $tenantStatsHash["SecuritySecureScore"][$score.ID] = $currentSecurityScores
        }
        Write-ProgressHelper -Activity "Processing all $($ServiceName) Security Score Details" -Completed

    }
    catch {
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running Get-SecuritySecureScoreReport function. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
        Write-Log -Type ERROR -Message "[Get-SecuritySecureScoreReport] An error occurred in running Get-SecuritySecureScoreReport function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
    }
    finally {
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-SecuritySecureScoreReport] COMPLETED: Gathering Security Score details" -ExportFileLocation $ExportDetails[0]
    }   
}


########################################################
# Initialization (Beginning)
########################################################

#Connect to all required O365 services for running this script
Connect-Office365Services -ServiceName ALL -ReadOnlyGraph
$global:tenant = (Get-OrganizationConfig).Name

#Get Export Path
Install-ImportExcelModule | Out-Null
$ExportDetails = Get-ExportPath -FileNameSegment AllTenantStats

$global:AllDiscoveryErrors = @()

# Prompt for level of detail
$reportingMode = Set-ReportMode

#Hash Table to hold final report data
$tenantStatsHash = @{}

#Global Start Time for Script
$global:InitialStart = Get-Date

########################################################
# Main Execution (Main Block)
########################################################

Write-Host
Write-Host "Gathering Exchange Online Objects and data" -ForegroundColor Black -BackgroundColor Yellow
Get-AllRecipientDetails -detailLevel $reportingMode
Get-AllExchangeMailboxDetails -detailLevel $reportingMode
Get-ExchangeGroupDetails -detailLevel $reportingMode
Get-MailFlowRulesandConnectors -detailLevel $reportingMode
Get-AllPublicFolderDetails -detailLevel $reportingMode
Write-Host

Write-Host "Gathering Collaboration/SharePoint Objects and data" -ForegroundColor Black -BackgroundColor Yellow
Get-AllUnifiedGroups  -detailLevel $reportingMode
Write-Host

Write-Host "Gathering Tenant Objects and License details" -ForegroundColor Black -BackgroundColor Yellow
#Attempt using Microsoft Graph if Graph was Successfully Connected
if ($global:MGGraph -eq $true) {
    Get-SPOAndOneDriveDetails -detailLevel $reportingMode -ServiceName MGGraph
    Get-TeamsDetails -detailLevel $reportingMode -ServiceName MGGraph
    #Get-ConditionalAccessPoliciesReport -detailLevel $reportingMode -ServiceName MGGraph #Still need to debug the Get-MGConditionalAccessPolicy command
    Get-SecuritySecureScoreReport -detailLevel $reportingMode -ServiceName MGGraph -MostRecent
    Get-AllLicenseSKUs -ServiceName MGGraph
    Get-allUserDetails -detailLevel $reportingMode -ServiceName MGGraph
    Get-AllOffice365Domains -ServiceName MGGraph
    Get-AllOffice365Admins -ServiceName MGGraph
    Get-AllDevicesReport -detailLevel $reportingMode -ServiceName MGGraph
}
else {
    Get-SPOAndOneDriveDetails -detailLevel $reportingMode -ServiceName SPO
    Get-TeamsDetails -detailLevel $reportingMode -ServiceName Teams
    Get-ConditionalAccessPoliciesReport -detailLevel $reportingMode -ServiceName AzureAD
    Get-SecuritySecureScoreReport -detailLevel $reportingMode -ServiceName AzureAD -MostRecent
    Get-AllLicenseSKUs -ServiceName MSOL
    Get-allUserDetails -detailLevel $reportingMode -ServiceName MSOL
    Get-AllOffice365Domains -ServiceName MSOL
    Get-AllOffice365Admins -ServiceName MSOL
    Get-AllDevicesReport -detailLevel $reportingMode -ServiceName Azure
}

#Combine Reporting - Optional
switch ($reportingMode) {
    #minimum {$mgUsers = Get-MgUser -all -ErrorAction Stop | select }
    {$_ -in "combined", "all"} {
        Write-Host "Consolidating Discovery Report data for each user / object into one file" -ForegroundColor Black -BackgroundColor Green
        #Combine Reports
        Combine-AllMailboxStats
    }
}

########################################################
### Export Reports ###
########################################################

#Exclude specific reports from Export
# Common tables to remove for both 'combined' and 'minimum' detail levels
$commonTablesToRemove = @(
    'PublicFolderStats', 
    'ArchiveMailboxStats', 
    'GroupMailboxes',
    'UserMailboxes',
    'NonUserMailboxes',
    'ServicePlans',
    'PrimaryMailboxStats'
)

#Exclude specific reports from Export
switch ($reportingMode) {
    "combined" {
        # Additional tables for combined detail level
        $additionalTables = @()
        $tablesToRemove = $commonTablesToRemove + $additionalTables
    }
    "minimum" {
        # Additional tables for minimum detail level
        $additionalTables = @('UserMailboxes')
        $tablesToRemove = $commonTablesToRemove + $additionalTables
    }
    default {
        # Default case to handle other detail levels where nothing should be removed
        $additionalTables = @()
        $tablesToRemove = $commonTablesToRemove + $additionalTables
    }
}
$ExportTenantStatsHash = @{}
$ExportTenantStatsHash = $tenantStatsHash

#Exclude specific reports from Export
foreach ($key in $tablesToRemove) {
    
    if ($ExportTenantStatsHash[$key]) {
        $ExportTenantStatsHash.Remove($key)
    }
}

#Export Reports: Exports each individual hashtable to own CSV file and then combines into Excel file
Write-Log -Type INFO -Message "Exporting the Tenant Statistics to $($ExportDetails[0])." -ExportFileLocation $ExportDetails[0]
try {
    Export-HashTableToExcel -hashtable $ExportTenantStatsHash -ExportDetails $ExportDetails
}
catch {
    $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in Exporting the Tenant Statistics to $($ExportDetails[0]). Please re-run the script and verify the location is valid and the file is not open in another application. $($_.Exception.Message)"
    $global:AllDiscoveryErrors += $ErrorObject
    Write-Log -Type ERROR -Message "An error occurred in Exporting the Tenant Statistics to $($ExportDetails[0]). Please re-run the script and verify the location is valid and the file is not open in another application. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
}

Write-Host ""

#Export Errors
try {
    #silence error reporting locations
    Export-ErrorReports -ExportFileLocation $ExportDetails[0] -ErrorData $global:AllDiscoveryErrors -logReportDirectory $ExportDetails[0]

    #Display Error Details
    Write-Host "Error Reporting Details" -ForegroundColor Black -BackgroundColor Yellow
    Write-Host "Check '$($errorReportFolderDirectory)' for error logs " -ForegroundColor Cyan
    Write-Host "$($global:AllDiscoveryErrors.count) " -ForegroundColor Red -NoNewline
    Write-Host "Error(s) encountered. "
    }
catch {
    if ($_.Exception.Message -like '*because it is an empty collection*') {
        Write-Warning "No Errors Found!"
    }

    else {
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in Exporting the Error Reports. $($_.Exception.Message)"
        $global:AllDiscoveryErrors += $ErrorObject
        Write-Log -Type ERROR -Message "An error occurred in Exporting the Error Reports. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
    }
}

Write-Host ""
Write-Host "Object Count Table" -ForegroundColor Black -BackgroundColor Green
$CompletedTime = (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
Write-Host "COMPLETED: Gathered Tenant Details. Completed Time: $($CompletedTime)" -ForegroundColor Cyan
Write-Log -Type INFO -Message "COMPLETED: Gathered Tenant Details. Completed Time: $($CompletedTime)" -ExportFileLocation $ExportDetails[0]


#Final Output of Recipient Counts
$TenantStatsOutput = @()

foreach ($key in $tenantStatsHash.Keys) {
    $count = $tenantStatsHash[$key].Count
    # Create a custom object for the current key-value pair
    $object = New-Object -TypeName PSCustomObject -Property @{
        "Key" = $key
        "Count" = $count
    }
    # Add the custom object to the array
    $TenantStatsOutput += $object
}
$TenantStatsOutput | ft

