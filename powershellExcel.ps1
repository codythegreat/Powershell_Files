# Open Excel
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $False
$OpenFile = $Excel.Workbooks.Open("C:\input\readCells.xlsx")
$Workbook = $OpenFile.Worksheets
$Worksheet = $Workbook.Item(1)

# Get the values for each column
$Code = $Worksheet.Cells | where {$_.value2 -eq "Code"} |
select -First 1
$Syntax = $Worksheet.Cells | where {$_.value2 -eq "Syntax"} |
select -First 1

# Get the values for each row in Code
$codeValues = @()
$codeValues = for($i=2; $Code.Cells.Item($i).Value2 -ne $null;
$i++  ){
    $Code.Cells.Item($i)
}

# Get the values for each row in Syntax
$syntaxValues = @()
$syntaxValues = for($i=2; $Syntax.Cells.Item($i).Value2 -ne
$null; $i++  ){
    $Syntax.Cells.Item($i)
}

$codeValues | ForEach-Object {Write-host $_.value2}
$syntaxValues | ForEach-Object {Write-Host $_.value2}
