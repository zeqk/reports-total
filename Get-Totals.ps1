

function Install-Nuget {
    $sourceNugetExe = "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe"
    $targetNugetExe = "$PSScriptRoot\nuget.exe"
    If( -Not (Test-Path $targetNugetExe)){
        Invoke-WebRequest $sourceNugetExe -OutFile $targetNugetExe
    }
}

function Install-IText {

    if(-Not (Test-Path "$PSScriptRoot\lib\itext7.7.1.12")) {
        # Download nuget.exe
        Install-Nuget

        # nuget install itext7
        Invoke-Expression "$PSScriptRoot\nuget.exe install iText7 -Version 7.1.12 -OutputDirectory $PSScriptRoot\lib"

        
        # requires itextsharp.dll

        Add-Type -Path "$PSScriptRoot\lib\itext7.7.1.12\lib\netstandard1.6\itext.kernel.dll"
        Add-Type -Path "$PSScriptRoot\lib\itext7.7.1.12\lib\netstandard1.6\itext.forms.dll"
        Add-Type -Path "$PSScriptRoot\lib\Common.Logging.3.4.1\lib\netstandard1.3\Common.Logging.dll"
        Add-Type -Path "$PSScriptRoot\lib\Common.Logging.Core.3.4.1\lib\netstandard1.0\Common.Logging.Core.dll"
        Add-Type -Path "$PSScriptRoot\lib\Portable.BouncyCastle.1.8.5\lib\netstandard1.3\BouncyCastle.Crypto.dll"
    }
}

function GetPublisherMonthReport 
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$publisherCardPath,
        [Parameter(Mandatory=$true)]
        [int]$serviceYear,
        [Parameter(Mandatory=$true)]
        [int]$month
    ) 
        
    try {


        # Get all pdf fields
        $reader = [iText.Kernel.Pdf.PdfReader]::new($publisherCardPath)        
        $PdfDoc = [iText.Kernel.Pdf.PdfDocument]::new($reader)
        $Form = [iText.Forms.PdfAcroForm]::getAcroForm($PdfDoc, $False)
        $Fields = $Form.getFormFields()

        # The PDF contains two tables, one table for one service year
        # The first table fields starts with the prefix "1", and the second with "2"
        # $fieldPrefix
        # $remarksSufix
        If(($Fields | Where-Object {$_.key -eq "Service Year"}).Value.GetValue() -Eq $serviceYear) {
            $fieldPrefix = "1"
            $remarksSufix = ""
        } ElseIf(($Fields | Where-Object {$_.key -eq "Service Year_2"}).Value.GetValue() -Eq $serviceYear) {
            $fieldPrefix = "2"
            $remarksSufix = "_2"
        } Else {
            Write-Warning "$name's card has no service year $serviceYear"
            return;
        }

        # Each table record ends with the month number (counting from september)
        # $fieldSufix
        # $monthName
        switch ($month) {
            9 { 
                $fieldSufix = 1
                $monthName = "September"
            }
            10 { 
                $fieldSufix = 2
                $monthName = "October"
            }
            11 { 
                $fieldSufix = 3
                $monthName = "November" 
            }
            12 { 
                $fieldSufix = 4
                $monthName = "December" 
            }
            1 { 
                $fieldSufix = 5
                $monthName = "January" 
            }
            2 { 
                $fieldSufix = 6
                $monthName = "February" 
            }
            3 { 
                $fieldSufix = 7
                $monthName = "March" 
            }
            4 { 
                $fieldSufix = 8
                $monthName = "April" 
            }
            5 { 
                $fieldSufix = 9
                $monthName = "May" 
            }
            6 { 
                $fieldSufix = 10
                $monthName = "June" 
            }
            7 { 
                $fieldSufix = 11
                $monthName = "July" 
            }
            8 { 
                $fieldSufix = 12
                $monthName = "August" 
            }
            Default { return; }
        }
        
        #  Get the values from the pdf fields
        $name = ($Fields | Where-Object {$_.key -eq "Name"}).Value.GetValue().ToString()
        $place = ($Fields | Where-Object {$_.key -eq "$fieldPrefix-Place_$fieldSufix"}).Value.GetValue() ?? 0
        $video = ($Fields | Where-Object {$_.key -eq "$fieldPrefix-Video_$fieldSufix"}).Value.GetValue() ?? 0
        $hours = ($Fields | Where-Object {$_.key -eq "$fieldPrefix-Hours_$fieldSufix"}).Value.GetValue() ?? 0
        $rv = ($Fields | Where-Object {$_.key -eq "$fieldPrefix-RV_$fieldSufix"}).Value.GetValue() ?? 0
        $studies = ($Fields | Where-Object {$_.key -eq "$fieldPrefix-Studies_$fieldSufix"}).Value.GetValue() ?? 0
        $remarks = ($Fields | Where-Object {$_.key -eq "Remarks$monthName$remarksSufix"}).Value.GetValue() ?? ""

        $monthReport = [PSCustomObject]@{
            Place = [int]$place.ToString()
            Video = [int]$video.ToString()
            Hours = [int]$hours.ToString()
            RV = [int]$rv.ToString()
            Studies = [int]$studies.ToString()
            Hola = "hola"
            Remarks = $remarks.ToString()
        }

        # Maybe some card was not filled.
        If($monthReport.Hours -le 0) {
            Write-Warning "$name's card has no hours entered for $month/$serviceYear"
        }

        return $monthReport
        

    } catch {
        Write-Error $_.Exception.Message
    }
    finally {
        $PdfDoc.Close()
    }
}

function GetMonthTotals
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$cardsPath,
        [Parameter(Mandatory=$true)]
        [int]$serviceYear,
        [Parameter(Mandatory=$true)]
        [int]$month,
        [Parameter(Mandatory=$false)]
        [string]$remarksPattern
    )
       
    $publishersTotalsItems = Get-ChildItem "$cardsPath/*.pdf" | ForEach-Object { GetPublisherMonthReport $_.FullName $serviceYear $month }

    $totalMonthReport = [PSCustomObject]@{
        Place = 0
        Video = 0
        Hours = 0
        RV = 0
        Studies = 0
        Count = [int]0
    }
   
    ForEach ($item in $publishersTotalsItems) {
        
        If($remarksPattern -And ($item.Remarks -NotMatch $remarksPattern)) {
            continue
        }
        
        $totalMonthReport.Place += $item.Place
        $totalMonthReport.Video += $item.Video
        $totalMonthReport.Hours += $item.Hours
        $totalMonthReport.RV += $item.RV
        $totalMonthReport.Studies += $item.Studies
        $totalMonthReport.Count++
    } 
    return $totalMonthReport;
}


function Get-Totals
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [int]$serviceYear,
        [Parameter(Mandatory=$true)]
        [int]$month
    )    

    Install-IText

    $config = Get-Content -Raw -Path $PSScriptRoot/config.json | ConvertFrom-Json

    # The Auxiliar Pioners are identify by some word in the Remarks (ex. "AP")
    $remarksAPIdentifier = $config.RemarksAPIdentifier

    $pubsTotals = GetMonthTotals $config.PubsFolder $serviceYear $month "^((?!$remarksAPIdentifier).)*$"
    $apTotals = GetMonthTotals $config.PubsFolder $serviceYear $month "$remarksAPIdentifier"
    $rpTotals = GetMonthTotals $config.PRFolder $serviceYear $month

    Write-Host "Publishers Totals"
    Write-Host ($pubsTotals | Format-Table | Out-String)
    Write-Host "Auxiliar Pioners Totals"
    Write-Host ($apTotals | Format-Table | Out-String)
    Write-Host "Regular Pioners Totals"
    Write-Host ($rpTotals | Format-Table | Out-String)
}

Get-Totals $args[0] $args[1]