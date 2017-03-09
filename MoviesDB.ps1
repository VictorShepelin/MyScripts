$strPath = "D:\OneDrive\Desktop\MoviesDB.xlsx"
$Folder = "C:\Videos\tosort"
$items = Get-ChildItem -Recurse -Path $Folder -include *.avi,*.mp4,*.mkv,*.m4v
[System.Threading.Thread]::CurrentThread.CurrentCulture = New-Object "System.Globalization.CultureInfo" "en-US"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$workbook = $excel.Workbooks.Add()
$workbook.Worksheets.Item(1).name = "Movies"
$sheet = $workbook.Worksheets.Item("Movies")
$sheet.cells.item(1,1) = "File Path"
$sheet.cells.item(1,2) = "IMDB Url"
$x = 2 # счетчик со второй строки т.к. первая это шапка
$j = 0
Do {
    $sheet.cells.item($x, 1) = $items[$j].FullName
    $x++
    $j++
    }
While ($j -lt $items.count)
#выравнивание
$range = $sheet.usedRange
$range.EntireColumn.AutoFit() | out-null
#сохранение
if (Test-Path $strPath) { 
    Remove-Item $strPath
    $Excel.ActiveWorkbook.SaveAs($strPath)
    }
else {
    $Excel.ActiveWorkbook.SaveAs($strPath)
    }