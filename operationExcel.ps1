$executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

$xls = "xls"

$FilesArrayList = [System.Collections.ArrayList]@()

get-childitem $executingScriptDirectory "*.xls" | foreach-object {
    $FilesArrayList.Add($_.FullName)
}

$objExcel = New-Object -ComObject Excel.Application  

$objExcel.Visible = $false

$referenceKey = "Drawing Count"

$TextToReplace = "KO"

$TextToReplaceWith = "OK"

for ($j = 0; $j -le ($FilesArrayList.Count - 1); $j += 1) {
  $filePath = $FilesArrayList.Item($j)

  Write-Host ("Working on => " + $filePath)

  $WorkBook = $objExcel.Workbooks.Open($filePath)
    $WorkSheet = $WorkBook.sheets.item(1)
    $rowMax = ($WorkSheet.UsedRange.Rows).count
    $rowMax

    $rowParam,$colParam = 7,2
    $rowStatus,$colStatus = 7,5

    for ($i=0; $i -le $rowMax-1; $i++)
    {
        $param = $WorkSheet.Cells.Item($rowParam+$i,$colParam).text

        if($param -eq $referenceKey) {
            $status = $WorkSheet.Cells.Item($rowStatus+$i,$colStatus).text
            if($status -eq $TextToReplace) {
                $WorkSheet.Cells.Item($rowStatus+$i,$colStatus) = $TextToReplaceWith
            }
            Write-Host ("param: "+$param)
            Write-Host ("Status: "+$status)
        }
    }

    $WorkBook.Save()

    $WorkBook.Close($true)
}

$objExcel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)
