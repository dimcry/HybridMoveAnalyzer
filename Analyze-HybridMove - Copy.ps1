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

<#[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    $Username,
    [Parameter(Mandatory=$true)]
    $Password
)
#>

$ErrorActionPreference = "Stop"

Function Connect-O365PowerShell
{
    $o365PowerShellUrl = "https://outlook.office365.com/PowerShell-LiveID"
    $credential = New-Object PSCredential(Get-Credential)
    $session = New-PSSession -ConnectionUri $o365PowerShellUrl `
                    -ConfigurationName Microsoft.Exchange `
                    -Credential $credential `
                    -Authentication Basic

    Import-PSSession $session -AllowClobber
    return $session
}

Function CollectLogs {
    # 1. Online or Offline?
    Do {
        Clear-Host
        Write-Host "Get-MoveRequestStatistics <Mailbox> -IncludeReport -DiagnosticInfo `"ShowTimeSlots, ShowTimeLine, Verbose`"" -ForegroundColor Green
        $OnlineOrOffline = Read-Host "`tDo you have the output of the above command, in an .xml file (Y/N)?"
    } while ($OnlineOrOffline -notmatch "Y|N")

    # 1.a. If Online
    If ($OnlineOrOffline -match "N") {
        $CurrentSession = Connect-O365PowerShell
        $GetUser = $null
        Clear-Host
        do {
            do {
                Write-Host "Enter the email address of the user for which you want to analyze the migration logs (Eg: " -ForegroundColor Green -NoNewline
                Write-Host "User1@contoso.com" -ForegroundColor Cyan -NoNewline
                Write-Host "):" -ForegroundColor Green -NoNewline
                Write-Host "`t" -NoNewline
                [string]$EmailAddress = Read-Host
            } while ($EmailAddress -notlike "*@*.*")
            
            Write-Host
            Write-Host "Checking if " -ForegroundColor White -NoNewline
            Write-Host "$EmailAddress" -ForegroundColor Cyan -NoNewline
            Write-Host " is a valid email address in your organization" -ForegroundColor White
            $GetUser = Get-User $EmailAddress -ErrorAction SilentlyContinue
        } while (!($GetUser))

        $MoveRequestStatisticsToAnalyze = $null
        [int]$CheckingCount = 1
        [int]$HowManyTimesToCheck = 3
        do {
            Write-Host "($CheckingCount / $HowManyTimesToCheck)" -ForegroundColor Green -NoNewline
            Write-Host " Checking if " -ForegroundColor White -NoNewline
            Write-Host "Get-MoveRequestStatistics" -ForegroundColor Green -NoNewline
            Write-Host " for " -ForegroundColor White -NoNewline
            Write-Host "$EmailAddress" -ForegroundColor Cyan -NoNewline
            Write-Host " exists" -ForegroundColor White
            $MoveRequestStatisticsToAnalyze = Get-MoveRequestStatistics $EmailAddress -IncludeReport -DiagnosticInfo "ShowTimeSlots, ShowTimeLine, Verbose" -ErrorAction SilentlyContinue
            $CheckingCount++
        } while (!($MoveRequestStatisticsToAnalyze) -and ($CheckingCount -le $HowManyTimesToCheck))


    # 1.b. If Offline
    } Else {
        Clear-Host
        do {
            do {
                Write-Host "Please type the full path of the .xml file (Eg.: " -ForegroundColor Green -NoNewline
                Write-Host "C:\Temp\MoveRequestStatistics.xml" -ForegroundColor Cyan -NoNewline
                Write-Host "):" -ForegroundColor Green -NoNewline
                Write-Host "`t" -ForegroundColor White -NoNewline
                [string]$XMLPath = Read-Host
            } while ($XMLPath -notlike "*.xml")
        
            [bool]$TestXMLPath = Test-Path $XMLPath -ErrorAction SilentlyContinue
            if (!($TestXMLPath)) {
                Clear-Host
                Write-Host $XMLPath -ForegroundColor Cyan -NoNewline
                Write-Host " doesn't exist! Please type the full path of an existing .xml file." -ForegroundColor Red
            }
        } while (!($TestXMLPath))
        $MoveRequestStatisticsToAnalyze = Import-Clixml $XMLPath
    }    
}

Function AnalyzeLogs {
    [bool]$ContinueToAnalyze = $False
    if (($($MoveRequestStatisticsToAnalyze).Status.Value.ToString()) -eq "Completed") {
        Write-Host
        Write-Host "The move of " -ForegroundColor White -NoNewline
        Write-Host "$EmailAddress" -ForegroundColor Cyan -NoNewline
        Write-Host " seems to be " -ForegroundColor White -NoNewline
        Write-Host "$($MoveRequestStatisticsToAnalyze.Status.Value.ToString())" -ForegroundColor Cyan -NoNewline
        Write-Host ". Do you want to further analyze the logs (Y/N)?:`t" -ForegroundColor White -NoNewline
        $ContinueOrNot = Read-Host
    
        if ($ContinueOrNot -match "Y") {
            [bool]$ContinueToAnalyze = $True
        }
    } else {
        [bool]$ContinueToAnalyze = $True
    }
    
    if ($ContinueToAnalyze) {
        if ($($MoveRequestStatisticsToAnalyze.Report)) {
            Write-Host "Report is present!" -ForegroundColor Green
        } else {
            Write-Host "Report is not present!" -ForegroundColor Red
            Write-Host "Do you want to collect the logs again, or you want to exit? (Y/N)"
            $NoReport = Read-Host
            do {
                if ($NoReport -match "N") {
                    ExitScript
                } elseif ($NoReport -match "Y") {
                    CollectLogs
                } else {
                    Write-Host "Please type `"Y`" if you want to collect the logs again, or `"N`" if you want to exit."
                    $NoReport = Read-Host
                }
            } while ($NoReport -notmatch "Y|N")
        }
    
        if ($($MoveRequestStatisticsToAnalyze.DiagnosticInfo)) {
            Write-Host "DiagnosticInfo is present!" -ForegroundColor Green
        } else {
            Write-Host "DiagnosticInfo is not present!" -ForegroundColor Red
            Write-Host "Do you want to collect the logs again, or you want to exit? (Y/N)"
            $NoDiagnosticInfo = Read-Host
            do {
                if ($NoDiagnosticInfo -match "N") {
                    ExitScript
                } elseif ($NoDiagnosticInfo -match "Y") {
                    CollectLogs
                } else {
                    Write-Host "Please type `"Y`" if you want to collect the logs again, or `"N`" if you want to exit."
                    $NoDiagnosticInfo = Read-Host
                }
            } while ($NoDiagnosticInfo -notmatch "Y|N")
        }
    }
}

Function ExitScript {
    Write-Host "The script will end now, after closing the current PowerShell session connected to Exchange Online!" -ForegroundColor Green
    Remove-PSSession -Id $($CurrentSession.Id)   
}

###############
# Main Script #
###############

CollectLogs
AnalyzeLogs
ExitScript