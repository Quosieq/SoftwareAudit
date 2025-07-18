<#
.SYNOPSIS
    Audits installed software and generates detailed reports.

.DESCRIPTION
    Scans Windows registry for installed applications, capturing:
    - Software names and versions
    - Installation dates and publishers
    - Install locations and license indicators
    Generates reports in multiple formats (TXT, CSV, HTML, XML, JSON).

.PARAMETER ReportType
    Specifies output format (TXT, CSV, HTML, XML, JSON).

.PARAMETER Formatting
    Required for HTML/XML reports. Options:
    - HTML: Table, List
    - XML: String, Stream

.PARAMETER SaveTo
    Optional custom directory for report output. Defaults to ./Reports.

.EXAMPLE
    .\SoftwareAudit.ps1
    Interactive mode - prompts for report format and location

.EXAMPLE
    .\SoftwareAudit.ps1 -ReportType CSV -SaveTo C:\Audits
    Generates CSV report in specified directory

.NOTES
    Version:        0.1
    Author:         Quosieq
    Creation Date:  $(Get-Date -Format 'yyyy-MM-dd')
    Requirements:   PowerShell 5.1 or later
    License:        MIT

.LINK
    Project Repository: https://github.com/Quosieq/SoftwareAudit
#>

function WriteLog {
    param (
        [Parameter(Mandatory=$true)]
        [System.Collections.ArrayList]$Source,
        [Parameter(Mandatory = $true)]
        [ValidateSet('TXT','CSV','HTML', 'XML','JSON')]
        [string]$ReportType,
        [Parameter(Mandatory = $false)]
        [ValidateSet({
            if($ReportType -in @('HTML', 'XML') -and -not $_) {
                throw "-As Parameter is required for XML and HTML reports"
            }
            $true
        },'String', 'Stream', 'Table','List')]
        [string]$Formatting,
        [Parameter(Mandatory=$false)]
        [string]$SaveTo = "$PSScriptRoot\Reports"

    )
    
    # Handle empty/null input by using the default
    if ([string]::IsNullOrWhiteSpace($SaveTo)) {
        $SaveTo = "$PSScriptRoot\Reports"
    }

    # Create directory if missing (supports both relative and absolute paths)
    if (-not (Test-Path -Path $SaveTo)) {
        Write-Host "Creating report directory at: $SaveTo" -ForegroundColor Yellow
        try {
            New-Item -ItemType Directory -Path $SaveTo -Force | Out-Null
        } catch {
            throw "Failed to create directory '$SaveTo': $_"
        }
    }

    # Generate filename with timestamp
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $logFile = Join-Path -Path $SaveTo -ChildPath "SoftAudit_$timestamp"
    switch ($ReportType) {
        'TXT'   {$Source | Format-Table -AutoSize | Out-File -FilePath "$logfile.txt" }
        'CSV'   {$Source | Export-CSV -path "$SaveTo\SoftAudit_$(Get-Date -Format yyyyMMdd_HHmmss).csv" -NoTypeInformation -Encoding ASCII}
        'HTML'  {$Source | ConvertTo-Html -as $Formatting | Out-File -FilePath "$($logFile).html"}
        'XML'   {$Source | ConvertTo-Xml -as $Formatting | Out-File -FilePath "$($logFile).xml"}
        'JSON'  {$Source | ConvertTo-Json -AsArray | Out-File -FilePath "$($logFile).json"}
    }
    Write-Host "Report saved to: $logFile.$($ReportType.Tolower())" -ForegroundColor Green

    
}

# Registry paths to check (both 32-bit and 64-bit)
$regSoft = @(
    "Registry::HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "Registry::HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)

$regKeys = Get-ChildItem -Path $regSoft 
$totalKeys = $regKeys.Count
$processedKeys = n
$arr =  New-Object System.Collections.ArrayList  


ForEach($regKey in $regKeys) {
    $processedKeys++
    $percentComplete = [math]::Min(($processedKeys / $totalKeys * 100),100)
    Write-Progress -Activity "Scanning registry keys... " `
                   -Status "$processedKeys/$totalKeys keys processed" `
                   -PercentComplete $percentComplete `
                   -id 1
    $foundSoft = Get-ItemProperty -Path "Registry::$regKey" 
    if(($null -eq $foundSoft.DisplayName) -or ($foundSoft.DisplayName -eq "")){
        Continue
    }
    $installDate = if ($foundSoft.InstallDate) {
        try {
            [datetime]::ParseExact($foundSoft.InstallDate, 'yyyyMMdd', $null).ToString('yyyy-MM-dd')
        }catch {
            "Invalid Date"
        }
    }else {
        "Unknown"
    }

    $soft = [PSCustomObject]@{
        Name                    = $foundSoft.DisplayName
        Version                 = if ($foundSoft.DisplayVersion) {$foundSoft.DisplayVersion} else {"N/A"}
        "Install Date"          = $installDate
        Publisher               = if ($foundSoft.Publisher) {$foundSoft.Publisher} else {"N/A"}
        "Install Location"      = if ($foundSoft.InstallSource) {$foundSoft.InstallSource} else {"N/A"}
    }
    [void]$arr.Add($soft)

}

Write-Progress -Activity "Scanning registry" -Completed -Id 1

$report = Read-Host("Do you want to save output as a report? (Y/N)")

while($true){
    Write-Progress -Activity "Saving report" -Status "Writing file..." -Id 2
    try{
        if($report -eq "y"){
            $type = Read-Host("Please specify in what format you want to save your report (TXT/CSV/HTML/XML/JSON)")
            $SaveTo = Read-Host("Default log save directory .\SoftAudit_$(Get-Date -Format yyyyMMdd_HHmmss).$type `nPlease specify where to save your file (optional)")
            if($type -in @('HTML', 'XML')){
                $type2 = Read-Host("Please specify type formatting for $type report")
                try{
                    WriteLog -Source $arr -ReportType $type -Formatting $type2 -SaveTo $SaveTo
                    break
                }catch {
                    Write-Error "Something went wrong when saving the file `n$_"
                    break
                }
                
            }
            try {
                WriteLog -Source $arr -ReportType $type -SaveTo $SaveT
                break
            }catch {
                Write-Error "Something went wrong when saving the file `n$_"
                break
            }
            break
        }elseif($report -eq "n") {
            Write-Host "Sending only output.. " -ForegroundColor Yellow
            $arr
            break
        }else {
            Write-Error "Invalid input, please try again!"
            continue
        }
    }finally {
         Write-Progress -Activity "Saving report" -Completed -Id 2
    }
}