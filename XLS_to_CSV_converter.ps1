param(
  [Parameter(Position=0,mandatory=$true)]
  [string] $ExcelFile,
  [Parameter(Position=1,mandatory=$true)]
  [string] $CSVFile
)


$xlCSV = 6
$Excel = New-Object -Com Excel.Application 
$Excel.visible = $False 
$Excel.displayalerts=$False 
$WorkBook = $Excel.Workbooks.Open("$ExcelFile") 
$Workbook.SaveAs("$CSVFile",$xlCSV) 
$Excel.quit()