# MIT License
# 
# Copyright (c) 2019 Cristian Dimofte
# 
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

# This script is analyzing the Hybrid move reports

##############################################
# Comments during the time script is in BETA #
##############################################

<#
The logic:
==========

0. BEGIN
1. Do we have the migration logs in .xml format?
    1.1. Yes - Go to 2.
    1.2. No - Go to 4.

2. Import the .xml file into a PowerShell variable. Is this information an output of <Get-MoveRequestStatistics $Mailbox -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose">?
    2.1. Yes - Go to 3.
    2.2. No - Go to 4.

3. Mark into the logs that the outputs were correctly collected, and we can start to analyze them. Go to 10.
4. Ask to provide user for which to analyze the migration logs. Go to 5.
5. Are we connected to Exchange Online?
    5.1. Yes - Go to 7.
    5.2. No - Go to 6.

6. Connect to Exchange Online using a Global Administrator. Go to 7.
7. Is the move request still present for this user?
    7.1. Yes - Go to 8.
    7.2. No - Go to 9.

8. Import into a PowerShell variable the output of <Get-MoveRequestStatistics $Mailbox -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose">. Go to 3.
9. Is the object a Mailbox in Exchange Online?
    9.1. Yes - Import the correct move request from the MoveHistory. Mark into the logs that the output was collected from MoveHistory. Go to 10.
    9.2. No - Are we connected to Exchange On-Premises?
        9.2.1. Yes - Go to 9.1.
        9.2.2. No - Inform that the user should have a Mailbox in On-Premises. Ask to run the same script in Exchange Management Shell, in On-Premises. Go to END.

10. Download the .xml/.json file containing the pairs, from GitHub. Go to 11.
11. Analyze the logs and provide the results. Go to END.
12. END
#>

################################################
# Common space for functions, global variables #
################################################

function Provide-PathOfXMLFile {

    Write-Host
    Write-Host "Please provide the exact path of the .xml file: " -ForegroundColor  Cyan
    Write-Host "`t" -NoNewline
    ##### Validation needed (the path)
    [string]$PathOfTheFile = Read-Host

    [bool]$ValidationOfThePath = Check-PathOrFile -PathToCheck $PathOfTheFile

    if ($ValidationOfThePath) {
        return $PathOfTheFile
    }

}

function Check-PathOrFile {
    param (
        # Full path of the .xml file
        [Parameter(Mandatory = $true)]
        [string]
        $PathToCheck
    )

    Write-Host
    Write-Host "We are validating the path you provided..." -ForegroundColor Cyan
    [bool]$ValidatePath = $True
    if (!(Test-Path "$PathToCheck" -ErrorAction SilentlyContinue))
    {
        $ValidatePath = $False
        Write-Host "The following is an invalid path, or, we are not able to connect to it: " -ForegroundColor Red -NoNewline
        Write-Host " `"$PathToCheck`" " -ForegroundColor White -NoNewline
    }
    else {
        Write-Host "`tPath successfully validated..." -ForegroundColor Green
    }
    
    return $ValidatePath
}

function Import-XMLInVariable {
    param (
        # .xml file to import
        [Parameter(Mandatory = $true)]
        [string]
        $FileToImport
    )

    Write-Host
    Write-Host "We are importing the file into a PowerShell variable..." -ForegroundColor Cyan
    [PSObject]$TheFile = Import-Clixml $FileToImport

    if ($TheFile) {
        Write-Host "`tWe successfully imported the file..." -ForegroundColor Green
        return $TheFile
    }
    else {
        Import-XMLInVariable -FileToImport $FileToImport
    }
}

function Provide-AffectedUser {

    Write-Host
    Write-Host "Please provide the email address or UserPrincipalName of the affected user: " -ForegroundColor Cyan
    Write-Host "`t" -NoNewline
    ##### Validation needed (email address)
    [string]$AffectedUser = Read-Host

    return $AffectedUser

}

function Validate-ConnectionToExchangeOnline {

    Write-Host "Checking if you are connected to Exchange Online..."
    ##### Validation needed (PSSession)
    [bool]$AlreadyConnectedToEXO = $true
    $GetPsSessions = Get-PSSession -ErrorAction SilentlyContinue

    foreach ($PsSession in $GetPsSessions){
        if (($($PsSession.ComputerName).ToLower() -eq "outlook.office365.com") -and ($($PsSession.State).ToString().ToLower() -eq "opened")) {

        }

    }
    
    
}

function Connect-ToExchangeOnline {

    Write-Host
    Write-Host "We are connecting you to Exchange Online..." -ForegroundColor Cyan
    Write-Host
    $cred = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection

    return $Session
    
}

function Validate-MoveRequestStatistics {
    param (
        # .xml file to validate
        [Parameter(Mandatory = $true)]
        [PSObject]
        $MoveRequestStatisticsToValidate
    )
 
    Write-Host
    Write-Host "We are validating if the information from the file you provided can be used for further analysis..." -ForegroundColor Cyan
    [bool]$InfoValidated = $false

    [xml]$DiagnosticInfoToCheck = $($MoveRequestStatisticsToValidate.DiagnosticInfo)
    if (($($MoveRequestStatisticsToValidate.Report)) -and ($($DiagnosticInfoToCheck.Job.TimeTracker.Durations)) -and ($($DiagnosticInfoToCheck.Job.TimeTracker.Timeline))) {
        [bool]$InfoValidated = $true
        Write-Host "`tThe information from the file will be used for further analysis..." -ForegroundColor Green
    }
    else {
        Write-Host "`tThe information from the file will not be used for further analysis. We will use other methods to collect the information related to the move..." -ForegroundColor Red
    }

    return $InfoValidated
}

Function Out-PieChart {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline)]
        [psobject] $inputObject,
        [Parameter()]
        [string] $PieChartTitle,
        [Parameter()]
        [int] $ChartWidth = 800,
        [Parameter()]
        [int] $ChartHeight = 400,
        [Parameter()]
        [string[]] $NameProperty,
        [Parameter()]
        [string] $ValueProperty,
        [Parameter()]
        [switch] $Pie3D,
        [Parameter()]
        [switch] $DisplayToScreen,
        [Parameter()]
        [string] $saveImage
    )
    begin {
        Add-Type -AssemblyName System.Windows.Forms.DataVisualization
        # Frame
        $Chart = [System.Windows.Forms.DataVisualization.Charting.Chart]@{
            Width       = $ChartWidth
            Height      = $ChartHeight
            BackColor   = 'White'
            BorderColor = 'Black'
        }
        # Body
        $null = $Chart.Titles.Add($PieChartTitle)
        $Chart.Titles[0].Font = "segoeuilight,20pt"
        $Chart.Titles[0].Alignment = "TopCenter"
        # Create Chart Area
        $ChartArea = [System.Windows.Forms.DataVisualization.Charting.ChartArea]::new()
        $ChartArea.Area3DStyle.Enable3D = $Pie3D.ToBool()
        $ChartArea.Area3DStyle.Inclination = 50
        $Chart.ChartAreas.Add($ChartArea)
        # Define Chart Area
        $null = $Chart.Series.Add("Data")
        $Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Pie
        # Chart style
        $Chart.Series["Data"]["PieLabelStyle"] = "Outside"
        $Chart.Series["Data"]["PieLineColor"] = "Black"
        $Chart.Series["Data"]["PieDrawingStyle"] = "Concave"

        $chart.Series["Data"].IsValueShownAsLabel = $true
        $chart.series["Data"].Label = "#PERCENT\n#VALX"
        # Set ArrayList
        $XColumn = [System.Collections.ArrayList]::new()
        $yColumn = [System.Collections.ArrayList]::new()
    }
    process {
        if (-not $valueProperty) {
            $numericProperties = foreach ($property in $inputObject.PSObject.Properties) {
                if ([Double]::TryParse($property.Value, [Ref]$null)) {
                    $property.Name
                }
            }
            if (@($numericProperties).Count -eq 1) {
                $valueProperty = $numericProperties
            }
            else {
                throw 'Unable to automatically determine properties to graph'
            }
        }
        if (-not $LabelProperty) {
            if ($inputObject.PSObject.Properties.Count -eq 2) {
                $LabelProperty = $inputObject.Properties.Name -ne $valueProperty
            }
            elseif ($inputObject.PSObject.Properties.Item('Name')) {
                $LabelProperty = 'Name'
            }
            else {
                throw 'Cannot convert Data'
            }
        }
        # Bind chart columns
        $null = $yColumn.Add($InputObject.$valueProperty)
        $null = $xColumn.Add($inputObject.$LabelProperty)
    }
    end {
        # Add data to chart
        $Chart.Series["Data"].Points.DataBindXY($xColumn, $yColumn)
        # Save file
        if ($psboundparameters.ContainsKey('saveImage')) {
            try{
                if (Test-Path (Split-Path $saveImage -Parent)) {
                    $SaveImage = $pscmdlet.GetUnresolvedProviderPathFromPSPath($saveImage)
                    $Chart.SaveImage($saveImage, "png")
                } else {
                    throw 'Invalid path, the parent directory must exist'
                }
            } catch {
                throw
            }
        }
        # Display Chart to screen
        if ($DisplayToScreen.ToBool()) {
            $Form = [Windows.Forms.Form]@{
                Width           = 800
                Height          = 450
                AutoSize        = $true
                FormBorderStyle = "FixedDialog"
                MaximizeBox     = $false
                MinimizeBox     = $false
                KeyPreview      = $true
            }
            $Form.controls.add($Chart)
            $Chart.Anchor = 'Bottom, Right, Top, Left'
            $Form.Add_KeyDown({
                if ($_.KeyCode -eq "Escape") { $Form.Close() }
            })
            $Form.Add_Shown( {$Form.Activate()})
            $Form.ShowDialog() | Out-Null
        }
    }
}

function Analyze-DiagnosticInfo {
    param (
        # XML of DiagnosticInfo to analyze
        [Parameter(Mandatory = $true)]
        [PSObject]
        $DiagnosticInfo
    )
    


}

function Build-TimeTrackerTable
{
    <#
        .Synopsis
        Retrieves the set of MRS indexes in AD for jobs matching the given query.

        .Parameter MrsJob
        An object returned by Get-*RequestStatistics with detailed time-tracker data.
		These data are obtained by passing these arguments to the cmdlet: -Diagnostic -DiagnosticArgument "showtimeline,verbose"

        .Parameter Aggregation
        Build the table from the ByMinute, ByHour, ByDay or ByMonth XML aggregations.
    #>
    param(
        [Parameter(Mandatory=$true)]
        $MrsJob,
        [Parameter(Mandatory=$false)]
        [ValidateSet('Minute', 'Hour', 'Day', 'Month')]
        [string]
        $Aggregation = 'Hour'
    )

    $diagnosticInfo = [xml]$MrsJob.DiagnosticInfo
    if ($diagnosticInfo -eq $null)
    {
        return
    }

	$seriesName = 'By{0}' -f $Aggregation
    $seriesData = $diagnosticInfo.Job.TimeTracker.Timeline.$seriesName
    if ($seriesData -eq $null -or $seriesData.$Aggregation.Count -eq 0)
    {
        return
    }

    $seriesSize = $seriesData.$Aggregation.D.Count
    $series = [System.Collections.Generic.List[object]]::new($seriesSize)
    foreach ($hour in $seriesData.$Aggregation)
    {
        $startTime = $hour.StartTime -as [DateTime]
        foreach ($entry in $hour.D)
        {
            $state = $entry.State
            $duration = $entry.Duration -as [TimeSpan]
			$msecs = $entry.MSecs -as [long]
            $row = [PSCustomObject][Ordered]@{
                'StartTime' = $startTime
                'State' = $state
                'Duration' = $duration
				'Milliseconds' = $msecs
				'CumulativeDuration' = $null
				'CumulativeMilliseconds' = $null
            }

            $series.Add($row)
        }
    }

	$series = $series | sort StartTime, State

	$gSeries = $series | group -NoElement State
	$accumulations = @{}
	$gSeries.Name | %{ $accumulations[$_] = [TimeSpan]::Zero }

    $series | %{
		$state = $_.State
		$accumulation = $accumulations[$state]
		$accumulation += $_.Duration
		$accumulations[$state] = $accumulation
		$_.CumulativeDuration = $accumulation
		$_.CumulativeMilliseconds = $accumulation.TotalMilliseconds
	}

	$series
}

<#
function DownloadTheJSonFile {

    $WebClient = New-Object System.Net.WebClient
    $WebClient.DownloadFile("https://github.com/dimcry/HybridMoveAnalyzer/blob/master/JSon_ErrorsAndRecommendations.json",".\ErrorsAndRecommendations.json")
  
}

$Request = "https://github.com/dimcry/HybridMoveAnalyzer/blob/master/JSon_ErrorsAndRecommendations.json"
Invoke-WebRequest $Request
#>

###############
# Main script #
###############

Get-PSSession | Remove-PSSession

Write-Host "Do you have the output of the " -ForegroundColor White -NoNewline
Write-Host "Get-MoveRequestStatistics <AffectedMailbox> -IncludeReport -DiagnosticInfo `"showtimeslots, showtimeline, verbose`"" -ForegroundColor Cyan -NoNewline
Write-Host " command in an .xml file (Y / N)? " -ForegroundColor White -NoNewline
##### Validation needed (y/n)
[char]$ReadFromKeyboard = Read-Host

[bool]$ValidationOfKey = $false
if ($ReadFromKeyboard.ToString().ToUpper() -eq "Y") {
    [string]$ThePathOfTheFile = Provide-PathOfXMLFile
    [PSObject]$TheMoveRequestStatistics = Import-XMLInVariable -FileToImport $ThePathOfTheFile
    
    [bool]$IsValidToBeUsed = Validate-MoveRequestStatistics -MoveRequestStatisticsToValidate $TheMoveRequestStatistics

    if (-not ($IsValidToBeUsed)) {
        $BackupMoveRequestStatistics = $TheMoveRequestStatistics
        $TheMoveRequestStatistics = New-Object PSObject
    }
    else {
        $ValidationOfKey = $true
    }

}
if (-not ($ValidationOfKey)) {
    [string]$EmailAddress = Provide-AffectedUser
    $EXOPSSession = Connect-ToExchangeOnline
    $ImportedSession = Import-PSSession $EXOPSSession -Prefix "EXO" -AllowClobber -ErrorAction SilentlyContinue

    Write-Host
    Write-Host "We are collecting the output of " -ForegroundColor Cyan -NoNewline
    Write-Host "`"Get-EXOMoveRequestStatistics $EmailAddress -IncludeReport -DiagnosticInfo `"showtimeslots, showtimeline, verbose`"" -ForegroundColor White -NoNewline
    Write-Host " from Exchange Online..." -ForegroundColor Cyan
    
    ##### Validation needed (Connected to EXO? Rights to run the command?)

    $TheMoveRequestStatistics = Get-EXOMoveRequestStatistics $EmailAddress -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose" -ErrorAction SilentlyContinue


    <#if (-not ($TheMoveRequestStatistics)) {
        # Are we still connected to Exchange Online?
        Get-Command Get-EXOMoveRequestStatistics
    }#>
}

Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process -Confirm:$false
# Unblock-File .\AnalyzeMoveRequestStats.ps1
. .\AnalyzeMoveRequestStats.ps1

$stats = $TheMoveRequestStatistics
ProcessStats -stats $stats -name ProcessedStats1

#DownloadTheJSonFile

<#
Get-Process |
    Select-Object -First 5 name, pm |
    Out-PieChart -PieChartTitle "Top 5 Windows processes running" -DisplayToScreen

Get-Service |
	Group-Object -Property Status -NoElement |
	Out-PieChart -PieChartTitle "Service Status" -Pie3D -saveImage C:\tmp\PieChart\'Service status.png'

#>

$timeline = Build-TimeTrackerTable -MrsJob $TheMoveRequestStatistics -Aggregation Day
$timeline | ft -AutoSize


<#
[xml]$DiagnosticInfo = $TheMoveRequestStatistics.DiagnosticInfo
$OutputForChart = Analyze-DiagnosticInfo -DiagnosticInfo $DiagnosticInfo
#>

#$MoveRequestStatistics = Import-Clixml C:\Temp\Doinita\dtest2.xml

#$key = [Console]::ReadKey()
#$char = $key.KeyChar
#$key.KeyChar.ToString().ToUpper()


############################


############################
#####################################
# Create / update .xml / .json file #
#####################################
<#
### If .xml, try not to use Import/Export-Clixml

[System.Collections.ArrayList]$ErrorsAndRecommendations = @()

$TheXML = New-Object -TypeName PSObject -ErrorAction SilentlyContinue
$TheXML | Add-Member -Type NoteProperty -Name "Failure Type" -Value "NoValue"
$TheXML | Add-Member -Type NoteProperty -Name "Actual Issue" -Value "NoValue"
$TheXML | Add-Member -Type NoteProperty -Name "Recommendation" -Value "NoValue"
$void = $ErrorsAndRecommendations.Add($TheXML)

$TheXML1 = New-Object -TypeName PSObject -ErrorAction SilentlyContinue
$TheXML1 | Add-Member -Type NoteProperty -Name "Failure Type" -Value "NoValue"
$TheXML1 | Add-Member -Type NoteProperty -Name "Actual Issue" -Value "NoValue"
$TheXML1 | Add-Member -Type NoteProperty -Name "Recommendation" -Value "NoValue"
$void = $ErrorsAndRecommendations.Add($TheXML1)

$TheXML2 = New-Object -TypeName PSObject -ErrorAction SilentlyContinue
$TheXML2 | Add-Member -Type NoteProperty -Name "Failure Type" -Value "NoValue"
$TheXML2 | Add-Member -Type NoteProperty -Name "Actual Issue" -Value "NoValue"
$TheXML2 | Add-Member -Type NoteProperty -Name "Recommendation" -Value "NoValue"
$void = $ErrorsAndRecommendations.Add($TheXML2)

$ErrorsAndRecommendations | Export-Clixml .\ErrorsAndRecommendations1.xml -Force
$JSon_ErrorsAndRecommendations = $ErrorsAndRecommendations | ConvertTo-Json -Depth 10
$JSon_ErrorsAndRecommendations | Out-File .\JSon_ErrorsAndRecommendations.json -Force
#>
############################