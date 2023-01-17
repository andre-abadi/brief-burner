$searchPath = "E:\Andre\Test"
#$batesNumber = 'ABC\.\d{4}\.\d{4}\.\d{4}'
#$batesNumber = '(.{3}\.\d{4}\.\d{4}\.\d{4})'
$batesNumber= '(.{3}\.\d{4}\.\d{4}\.\d{4})|([A-Z])(\d{8})|(NTC\d{7})'

$outputFile = (Split-Path $searchPath -Parent) + "\_barcodes.txt"
$errorFile = (Split-Path $searchPath -Parent) + "\_error_pdf.txt"

# arrays to store found stuff
$foundBatesNumbers = @()
$errors = @()

# take user input for folder to seach 
$searchPath = Read-Host "Folder to Search Recursively"

if($searchPath -eq ""){
    $searchPath = "E:\Andre\Test"
    Write-Host "  No input given, defaulting to $searchPath"
}

# check for the interop assembly so that word and excel files can be opened
$wordAssembly = [Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Word")
$excelAssembly = [Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Excel")
if($wordAssembly -eq $null -or $excelAssembly -eq $null){
    Write-Host "Microsoft Office Interop assemblies not found"
}

#start up word and excel in the background
$word = New-Object -ComObject Word.Application
$excel = New-Object -ComObject Excel.Application

# enumerate all files into an object
$files = Get-ChildItem -Path $searchPath -Recurse -Force

# iterate through the files
foreach ($file in $files) {

    # do some things if the filename matches the regex
    if ($file.Name -match $batesNumber) {
        Write-Host "Filename hit:"
        Write-Host "  $file"
        $match = [regex]::Match($file.Name, $batesNumber)
        $matchValue = $match.Value
        $foundBatesNumbers += $matchValue
        Write-Host "  $matchValue"
    }

    # open the file and look for regex if it is a Word document
    if ($file.Extension -eq ".doc" -or $file.Extension -eq ".docx") {
        if($file -ne $null){
            Write-Host "Word document at: $file"
            $document = $word.Documents.Open($file.FullName)
            $text = $document.Content.Text
            $filtered = Select-String -InputObject $text -pattern $batesNumber -AllMatches | % {$_.Matches.Value}
            if ($filtered -ne $null) {
                Write-Host "  Containing: $filtered"
            } else {
                Write-Host "  Containing no hits."
            }
            $foundBatesNumbers += $filtered
            $document.Close()
        }
    }
    # open the file and look for regex if it is an Excel document
    elseif ($file.Extension -eq ".xls" -or $file.Extension -eq ".xlsx") {
        if($file -ne $null){
            Write-Host "Excel Workbook at: $file"
            $workbook = $excel.Workbooks.Open($file.FullName)
            $range = $workbook.ActiveSheet.UsedRange
            $text = $range.Value()
            $filtered = Select-String -InputObject $text -pattern $batesNumber -AllMatches | % {$_.Matches.Value}
            if ($filtered -ne $null) {
                Write-Host "  Containing: $filtered"
            } else {
                Write-Host "  Containing no hits."
            }
            $foundBatesNumbers += $filtered
            $workbook.Close()
        }
    }

}

# end the spawned word and excel instances
$word.Quit()
$excel.Quit()

# write outputs to a file
$foundBatesNumbers | select -Unique | sort-object | Out-File -FilePath $outputFile
if ($errors -ne $null) {
    $errors | select -Unique | sort-object | Out-File -FilePath $errorFile
}
