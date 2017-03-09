$filePath = "D:\OneDrive\Desktop\MoviesDB.xlsx"
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

Function SetAttribute ($filePath, $url) {
    $TagLibDll = "F:\Downloads\taglib-sharp.dll" #https://www.nuget.org/packages/taglib/2.1.0
    # getting content from IMDB url
     $webClient = new-object system.net.WebClient
     $webpage = $webClient.DownloadData($url)
     $string = [System.Text.Encoding]::ASCII.GetString($webpage)
     $output = "$env:TEMP\site.html"
     if (Test-Path $output) {
         Remove-Item $output -Force
         }
     $string | Out-File $output
    # parsing web content
     $a = 0
     $title = $null
     [array]$genres = $null
     $releaseDate = $null
     [array]$stars = $null
     Get-Content $output | ForEach-Object {
         if ($_ -match "meta property='og:title'") { # getting title
            $title = $_
            $chararray = $title.ToCharArray()
            $i = 0
            $from = $null
            $chararray | ForEach-Object {
                if ($_ -match '"'){
                    if ($from -eq $null) {
                        $from = $i+1
                    }
                    else {
                        $to = $i
                    }
                }
                $i++
            }
            $title = $title.Substring($from,($to-$from))
         }
         if ($_ -match 'class="itemprop" itemprop="genre"') { # getting genre
            $genre = $_
            $genre = $genre -replace "/span></a>"
            $genre = $genre -replace '><span class="itemprop" itemprop="genre"'
            $chararray = $genre.ToCharArray()
            $i = 0
            $from = $null
            $chararray | ForEach-Object {
                if ($_ -match '>'){
                    $from = $i+1
                }
                if ($_ -match '<'){
                    $to = $i
                }
                $i++
            }
            $genre  = $genre.Substring($from,($to-$from))
            $genres += $genre
         }
         if ($_ -match 'See more release dates') { # getting release date
            $b = $a
         }
         if (($_ -match 'meta itemprop="datePublished"') -and ($b -eq ($a-1))) { # getting year, right after release date
            $releaseDate = $_
            $releaseDate = $releaseDate -replace '<meta itemprop="datePublished" content="'
            $releaseDate = $releaseDate.Substring(0,4)
         }
         if ($_ -match 'tt_ov_st_sm') { # getting stars url and their names
            if ($_ -notmatch "fullcredits") {
                $starUrl = $_
                $starUrl = $starUrl -replace '"'
                $starUrl = $starUrl -replace '<a href=',"http://www.imdb.com/"
                $webClient = new-object system.net.WebClient
                $webpage = $webClient.DownloadData($starUrl)
                $string = [System.Text.Encoding]::ASCII.GetString($webpage)
                $output = "$env:TEMP\star.html"
                if (Test-Path $output) {
                    Remove-Item $output -Force
                } ##end if test-path
                $string | Out-File $output
                Get-Content $output | ForEach-Object {
                    if ($_ -match '<title>') {
                        $star = $_
                        $star = $star -replace "        <title>"
                        $star = $star -replace " - IMDb</title>"
                        $stars += $star
                    }
                } #end Get-Content
            } #end if fullcredits
         }
     $a++
     }
    $genres | ForEach-Object {
        [string]$JoinGenres += $_ + ";"
     }
    $stars | ForEach-Object {
        [string]$AlbumArtists += $_ + ";"
     }
    [System.Reflection.Assembly]::LoadFile($TagLibDll) | Out-Null
    $mediafile = [TagLib.File]::Create("$filePath")
    if (!($mediafile.Writeable)) {
        Write-Host "file read-only - fix it"
        Read-Host
     }
    $mediafile.Tag.AlbumArtists = $AlbumArtists # tag 13
    $mediafile.Tag.Year = $releaseDate # tag 15
    $mediafile.Tag.Genres = $JoinGenres # tag 16
    $mediafile.Tag.Title = $title # tag 21
    $mediafile.Tag.copyright = $url # tag 25
    $mediafile.Save()
 }
$ExcelData = GetExcelData -filePath $filePath
$ExcelData.GetEnumerator() | ForEach-Object {
    SetAttribute -filePath $_.key -url $_.value
}