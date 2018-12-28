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

function Check-PathOrFile {
    param (
        # Full path of the .xml file
        [Parameter(Mandatory = $true)]
        [string]
        $PathToCheck
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

function Import-XMLInVariable {
    param (
        # .xml file to import
        [Parameter(Mandatory = $true)]
        [string]
        $FileToImport
    )

    $TheFile = Import-Clixml $FileToImport

    return $TheFile
}



###############
# Main script #
###############

Write-Host "Do you have the output of the " -ForegroundColor White -NoNewline
Write-Host "Get-MoveRequestStatistics <AffectedMailbox> -IncludeReport -DiagnosticInfo `"showtimeslots, showtimeline, verbose`"" -ForegroundColor Cyan -NoNewline
Write-Host " command in an .xml file (Y / N)? " -ForegroundColor White -NoNewline
$ReadFromKeyboard = ""


#$MoveRequestStatistics = Import-Clixml C:\Temp\Doinita\dtest2.xml

#$key = [Console]::ReadKey()
#$char = $key.KeyChar
#$key.KeyChar.ToString().ToUpper()


############################

