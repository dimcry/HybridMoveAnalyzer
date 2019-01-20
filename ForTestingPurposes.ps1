function Ask-ForXMLPath {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [int]
        $NumberOfChecks
    )

    [string]$PathOfXMLFile = ""
    if ($NumberOfChecks -eq "1") {
        Write-Host "Please provide the path of the .xml file: " -ForegroundColor Cyan
        Write-Host "`t" -NoNewline
        try {
            $PathOfXMLFile = Validate-XMLPath -filePath (Read-Host)
        }
        catch {
            $NumberOfChecks++
            Ask-ForXMLPath -NumberOfChecks $NumberOfChecks
        }
    }
    else {
        Write-Host
        Write-Host "The path you provided is not valid!" -ForegroundColor Red

        Write-Host "Would you like to provide it again?" -ForegroundColor Cyan
        Write-Host "`t[Y] Yes     [N] No      (default is `"N`"): " -NoNewline -ForegroundColor White
        $ReadFromKeyboard = Read-Host

        [bool]$TheKey = $false
        Switch ($ReadFromKeyboard) 
        { 
          Y {$TheKey=$true} 
          N {$TheKey=$false} 
          Default {$TheKey=$false} 
        }

        if ($TheKey) {
            Write-Host
            Write-Host "Please provide again the path of the .xml file: " -ForegroundColor Cyan
            Write-Host "`t" -NoNewline
            try {
                $PathOfXMLFile = Validate-XMLPath -filePath (Read-Host)
            }
            catch {
                Write-Host "The path you provided is still not valid" -ForegroundColor Red
                Write-Host "We will continue to collect the migration logs using other methods" -ForegroundColor Red
                $PathOfXMLFile = "ValidationOfFileFailed"
            }
        }
        else {
            Write-Host
            Write-Host "We will continue to collect the migration logs using other methods" -ForegroundColor Red
            $PathOfXMLFile = "ValidationOfFileFailed"
        }
    }
    return $PathOfXMLFile
}

function Validate-XMLPath {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [ValidateScript({Test-Path $_})]
        [string]
        $filePath
    )

    if (($filePath.Length -gt 4) -and ($filePath -like "*.xml")) {
        Write-Host
        Write-Host $filePath -ForegroundColor Cyan -NoNewline
        Write-Host " is a valid .xml file. We will use it to continue the investigation" -ForegroundColor Green
    }
    else {
        $filePath = "NotAValidPath"
    }

    return $filePath
}

Clear-Host
[int]$TheNumberOfChecks = 1
[string]$ThePath = Ask-ForXMLPath -NumberOfChecks $TheNumberOfChecks -Verbose