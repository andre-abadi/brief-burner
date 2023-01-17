$searchPath = "E:\Andre\Test"
#$batesNumber = 'ABC\.\d{4}\.\d{4}\.\d{4}'
#$batesNumber = '(.{3}\.\d{4}\.\d{4}\.\d{4})'
$batesNumber= '(.{3}\.\d{4}\.\d{4}\.\d{4})|([A-Z])(\d{8})|(NTC\d{7})'

$outputFile = (Split-Path $searchPath -Parent) + "\_brief-burner_01_barcodes.txt"
$logFile = (Split-Path $searchPath -Parent) + "\_brief-burner_02_log.txt"

# Check and create the logfile if it does not exist
if (Test-Path -Path $logFile) {
    Remove-Item $logFile -Force
}
#New-Item -ItemType File -Path $logFile > $null


# Check and remove any pre-existing output files
if (Test-Path -Path $outputFile) {
    Remove-Item $outputFile -Force
}

# arrays to store found stuff
$foundBatesNumbers = @()

# take user input for folder to seach 
$searchPath = Read-Host "Folder to Search Recursively"

if($searchPath -eq ""){
    $searchPath = "E:\Andre\Test"
    echo "`nNo input given, defaulting to $searchPath`n" | Tee-Object -FilePath $logFile -Append 
}

# check for the interop assembly so that word and excel files can be opened
$wordAssembly = [Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Word")
$excelAssembly = [Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Excel")
if($wordAssembly -eq $null -or $excelAssembly -eq $null){
    echo "Microsoft Office Interop assemblies not found" | Tee-Object -FilePath $logFile -Append 
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
        echo "Filename hit:" | Tee-Object -FilePath $logFile -Append 
        echo "  $file" | Tee-Object -FilePath $logFile -Append 
        $match = [regex]::Match($file.Name, $batesNumber)
        $matchValue = $match.Value
        $foundBatesNumbers += $matchValue
        echo "  $matchValue" | Tee-Object -FilePath $logFile -Append 
    }

    # open the file and look for regex if it is a Word document
    if ($file.Extension -eq ".doc" -or $file.Extension -eq ".docx") {
        if($file -ne $null){
            echo "Word document at: $file" | Tee-Object -FilePath $logFile -Append 
            $document = $word.Documents.Open($file.FullName)
            $text = $document.Content.Text
            $filtered = Select-String -InputObject $text -pattern $batesNumber -AllMatches | % {$_.Matches.Value}
            if ($filtered -ne $null) {
                echo "  Containing: $filtered" | Tee-Object -FilePath $logFile -Append 
            } else {
                echo "  Containing no hits." | Tee-Object -FilePath $logFile -Append 
            }
            $foundBatesNumbers += $filtered
            $document.Close()
        }
    }
    # open the file and look for regex if it is an Excel document
    elseif ($file.Extension -eq ".xls" -or $file.Extension -eq ".xlsx") {
        if($file -ne $null){
            echo "Excel Workbook at: $file" | Tee-Object -FilePath $logFile -Append 
            $workbook = $excel.Workbooks.Open($file.FullName)
            $range = $workbook.ActiveSheet.UsedRange
            $text = $range.Value()
            $filtered = Select-String -InputObject $text -pattern $batesNumber -AllMatches | % {$_.Matches.Value}
            if ($filtered -ne $null) {
                echo "  Containing: $filtered" | Tee-Object -FilePath $logFile -Append 
            } else {
                echo "  Containing no hits." | Tee-Object -FilePath $logFile -Append 
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
if ($foundBatesNumbers -ne $null) {
    $foundBatesNumbers | select -Unique | sort-object | Out-File -FilePath $outputFile
    echo "`nBarcodes deduplicated, sorted, and written to: $outputFile" | Tee-Object -FilePath $logFile -Append 
}
