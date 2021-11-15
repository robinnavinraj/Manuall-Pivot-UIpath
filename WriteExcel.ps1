$ExcelObject = new-Object -comobject Excel.Application 
$strPath1="C:\Users\Admin\Documents\UiPath\Devlopment\Input files\Raw data.xlsx"
$ActiveWorkbook = $ExcelObject.WorkBooks.Open($strPath1)  
$ActiveWorksheet = $ActiveWorkbook.Worksheets.Item(1)
$ActiveWorksheet.Cells.Item(1,15) = "robinnavinraj" 
$ActiveWorksheet.Cells.Item(1,19) = "robinnavinraj" 
$ActiveWorksheet.Cells.Item(1,20) = "robinnavinraj" 
$ActiveWorkbook.Save()
$ActiveWorkbook.close($true)