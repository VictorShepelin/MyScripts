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
    Write-Host "working with $filePath"
    $TagLibDll = "D:\OneDrive\Tools\Multimedia\taglib-sharp.dll" #https://www.nuget.org/packages/taglib/2.1.0
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
            $title = $title -replace ":"," -"
            $title = $title -replace "&quot;"
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
    # xml tag creation
     $tagfile = $filePath.Substring(0,$filePath.length-3) + "xml"
     [xml]$tag = Get-Content "F:\Temp\example.xml"
     $tag.Tags.Tag.Simple[0].String = $title #TITLE
     $tag.Tags.Tag.Simple[1].String = $AlbumArtists #ARTIST
     $tag.Tags.Tag.Simple[2].String = $releaseDate #DATE_RELEASED
     $tag.Tags.Tag.Simple[3].String = $url #COPYRIGHT
     $tag.Tags.Tag.Simple[4].String = $JoinGenres #GENRE
     $tag.Save("$tagfile")
     Write-Host $title $AlbumArtists $releaseDate $JoinGenres $url -ForegroundColor Green
    if ($filePath.Substring($filePath.length-3,3)-match "avi") {
        Write-Host "Working with avi"
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
    if ($filePath.Substring($filePath.length-3,3)-match "mkv") {
        Write-Host "Working with mkv"
        $output = "C:\Videos\Movies\$title.mkv"
        & "C:\Program Files\mkvtoolnix\mkvmerge.exe" --output $output --global-tags $tagfile --title $title $filePath
        Remove-Item $filePath -Force
        Remove-Item $tagfile
     }
 }
$ExcelData = GetExcelData -filePath $filePath
$ExcelData.GetEnumerator() | ForEach-Object {
    SetAttribute -filePath $_.key -url $_.value
}