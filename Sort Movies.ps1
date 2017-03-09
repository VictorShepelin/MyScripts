$filePath = "D:\OneDrive\Desktop\MoviesDB.xlsx"
$TagLibDll = "F:\Downloads\taglib-sharp.dll" #https://www.nuget.org/packages/taglib/2.1.0
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
$ExcelData = GetExcelData -filePath $filePath
$ExcelData.GetEnumerator() | ForEach-Object {
    $ext = Get-ChildItem -Path $_.key | Select-Object -ExpandProperty Extension
    $mediafile = [TagLib.File]::Create("$_.key")
    $path = "C:\Videos\Movies\" + $mediafile.Tag.Genres
    $newFilePath = "$path\" + $mediafile.Tag.Title + $ext + ".txt"
    New-Item -Path $path -ItemType Directory
    New-Item -Path $newFilePath -ItemType File
    #Move-Item -Path $_.key -Destination $path
}