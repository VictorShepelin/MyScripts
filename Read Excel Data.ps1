Function GetExcelData ($filePath) {
    [System.Threading.Thread]::CurrentThread.CurrentCulture = New-Object "System.Globalization.CultureInfo" "en-US"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open("$filePath")
    $sheet = $workbook.Worksheets.Item(1)
    $data = @{}
    $intRow = 2
    Do {
        $moviePath = $sheet.Cells.Item($intRow, 1).Value()
        $url = $sheet.Cells.Item($intRow, 2).Value()
        $data.Add($moviePath,$url)
        $intRow++
     }
     While ($sheet.Cells.Item($intRow,1).Value() -ne $null)
    $excel.Workbooks.Close()
    return $data
}
$filePath = "D:\OneDrive\Desktop\MoviesDB.xlsx"
$ExcelData = GetExcelData -filePath $filePath
$ExcelData