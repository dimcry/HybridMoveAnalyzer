# MIT License
# 
# Copyright (c) 2018 Cristian Dimofte
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

Function Check-PathOrFile
{
    <#
        .Synopsis
        The role of this function is to collect the Path from where to Import a file, or, where to Export a file.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [String]$PathToCheck        
    )

    [bool]$ValidatePath = $True
    if (!(Test-Path "$PathToCheck" -ErrorAction SilentlyContinue))
    {
        $ValidatePath = $False
        Write-Host "The following is an invalid path, or, we are not able to connect to it: " -ForegroundColor Red -NoNewline
        Write-Host " `"$PathToCheck`" " -ForegroundColor White -NoNewline
    }
    return $ValidatePath
}

Function Mark-WhichLogs {
    param (
        [Parameter(Mandatory = $true)]
        [String]$PathToCheck
    )

    Write-Debug ""

}

Function Import-OutputIntoVariable {
    param (
        [Parameter(Position=1, Mandatory=$true, ParameterSetName='ImportFromFile')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Path,

        [Parameter(Position=1, Mandatory=$true, ParameterSetName='ImportFromCommand')]
        [string]
        $EmailAddress
    )


}



###############
# Main script #
###############

#Ask-ForXMLFile

Write-Host "We need to analyze the output of the following command:" -ForegroundColor White
Write-Host "`tGet-MoveRequestStatistics" -ForegroundColor Cyan -NoNewline
Write-Host " User1@contoso.com " -ForegroundColor White -NoNewline
Write-Host "`-IncludeReport -DiagnosticInfo `"showtimeslots, showtimeline, verbose`"" -ForegroundColor Cyan

##### Need to add Data Validation for the "yes/no" answer expected!!!
Do {
    $ReadAnswerPoint1 = Read-Host "Do you have this in an .xml file already? (y/n)"
    Write-Host "`r" -NoNewline
}
while ($ReadAnswerPoint1 -notmatch "Y|N|Yes|No")

##### Need to add Data Validation for the "path"!!!
if ($ReadAnswerPoint1 -match "Y|Yes") {
    Write-Host "Please provide the exact path of this file (Eg.: " -ForegroundColor White -NoNewline
    Write-Host "`"C:\Temp\MoveRequestStatistics.xml`"" -ForegroundColor Cyan -NoNewline
    Write-Host ")" -ForegroundColor White
    Write-Host "`t" -NoNewline
    $PathPoint1 = Read-Host

##### Need to add Data Validation for the ".xml Import"!!!
    $MoveRequestStatisticsToCheck = Import-Clixml $PathPoint1
    [xml]$DiagnosticInfoToCheck = $($MoveRequestStatisticsToCheck.DiagnosticInfo)

    if (($($MoveRequestStatisticsToCheck.Report)) -and ($($DiagnosticInfoToCheck.Job.TimeTracker.Durations)) -and ($($DiagnosticInfoToCheck.Job.TimeTracker.Timeline))) {
        Write-Host "We will use the current output to analyze"
##### Mark into the log file that the current output is good to be analyzed, as is output of Get-MoveRequestStatistics $Mailbox -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose"
##### Download the .xml/.json file, analyze the logs and provide results

    }
    else {
        Write-Host "The current output is not of the `<Get-MoveRequestStatistics $Mailbox -IncludeReport -DiagnosticInfo `"showtimeslots, showtimeline, verbose`"`> command" -ForegroundColor Red
        Collect-EmailAddress
    }

}
##### Need to add Data Validation for the "EmailAddress"!!!
else {
    Write-Host
    Write-Host "We will continue by trying to collect the migration logs"
    Collect-EmailAddress
#    Write-Host "Please provide the email address of the user for which we have to do the check" -ForegroundColor Cyan -NoNewline
#    $EmailAddress = Read-Host
}

