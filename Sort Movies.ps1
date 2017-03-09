$filePath = "D:\OneDrive\Desktop\MoviesDB.xlsx"
$TagLibDll = "F:\Downloads\taglib-sharp.dll" #https://www.nuget.org/packages/taglib/2.1.0
[System.Reflection.Assembly]::LoadFile($TagLibDll) | Out-Null
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
    $movie = $_.key
    $movieFile = $_.key # таг либ требует в явном виде и без лайфхака
    $movie = $movie -replace '\[','`[' -replace '\]','`]' # вынужденый лайфхак из-за скобок
    if (Test-Path $movie) {
        $ext = $null
        $ext = Get-ChildItem -Path $_.key | Select-Object -ExpandProperty Extension
        $ext
        if ($ext -eq $null) {
            $chararray = $movie.ToCharArray()
            $i = 0
            $dot = $null
            $chararray | ForEach-Object {
                if ($_ -eq ".") {
                    $dot = $i
                }
                $i++
            }
            $ext = $movie.Substring($dot,($chararray.count-$dot))
        }
        try {
            $mediafile = [TagLib.File]::Create("$movieFile")
        }
        catch {
            Write-Host "Error"
        }
        if ($mediafile.Tag.Genres) {
            $genrePath = $mediafile.Tag.Genres
            $chararray = $genrePath.ToCharArray().count
            $genrePath = $genrePath.Substring(0,($chararray-1))
            $genrePath = $genrePath -replace ";"," "
            $path = "C:\Videos\Movies\" + $genrePath
            $newFilePath = "$path\" + $mediafile.Tag.Title + $ext
            Write-Host "$movie будет переименован в $newFilePath"
            New-Item -Path $path -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
            Move-Item -Path $movie -Destination $newFilePath
        }
        else {
            $newFilePath = $null
        }
    }
    else {
        $movie
        Test-Path $movie        
    }
 }