$Folder = "C:\Videos\Horror"
$File = "Krik.2.1997.DUAL.BDRip.XviD.AC3.-HQCLUB.avi"
$url = "http://www.imdb.com/title/tt0120082/"
$TagLibDll = "F:\Downloads\taglib-sharp.dll" #https://www.nuget.org/packages/taglib/2.1.0
# select imdb
    $webClient = new-object system.net.WebClient
    $webClient.proxy = $proxy
    $webpage = $webClient.DownloadData($url)
    $string = [System.Text.Encoding]::ASCII.GetString($webpage)
    $output = "$env:TEMP\site.html"
    if (Test-Path $output) {
        Remove-Item $output -Force
        } ##end if test-path
    $string | Out-File $output
    $a = 0
    $title = $null
    [array]$genres = $null
    $releaseDate = $null
    [array]$stars = $null
    Get-Content $output | ForEach-Object {
        if ($_ -match "meta property='og:title'") {
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
        } #end if title
        if ($_ -match 'class="itemprop" itemprop="genre"') {
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
        } #end if genre
        if ($_ -match 'See more release dates') {
            $b = $a
        } #end if release
        if (($_ -match 'meta itemprop="datePublished"') -and ($b -eq ($a-1))) {
            $releaseDate = $_
            $releaseDate = $releaseDate -replace '<meta itemprop="datePublished" content="'
            $releaseDate = $releaseDate.Substring(0,4)
        } #end if date
        if ($_ -match 'tt_ov_st_sm') {
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
        } #end if stars

    $a++
    } #end Get-Content
    #end of select imdb

$genres | ForEach-Object {
    [string]$JoinGenres += $_ + ";"
}
$stars | ForEach-Object {
    [string]$AlbumArtists += $_ + ";"
}
[System.Reflection.Assembly]::LoadFile($TagLibDll) | Out-Null
$mediafile = [TagLib.File]::Create("$Folder\$File")
if (!($mediafile.Writeable)) {
    Write-Host "file read-only - fix it"
    Read-Host
}
$mediafile.Tag.AlbumArtists = $AlbumArtists # в ролях 13ый таг
$mediafile.Tag.Year = $releaseDate # год выхода 15ый таг
$mediafile.Tag.Genres = $JoinGenres # жанр 16ый таг
$mediafile.Tag.Title = $title # название 21ый таг
$mediafile.Tag.Comment = "" # коммент 24ый таг
$mediafile.Tag.copyright = $url # ссылка на imdb 25ый таг
$mediafile.Save()
#$mediafile.Tag #all current tags
$mediafile.Tag.Title
$mediafile.Tag.Year
$mediafile.Tag.Genres
$mediafile.Tag.AlbumArtists
$mediafile.Tag.Comment
$mediafile.Tag.copyright
<# без taglib-sharp.dll можно только читать
    $objShell = New-Object -ComObject Shell.Application 
    $objFolder = $objShell.NameSpace($Folder)
    $objFile = $objFolder.ParseName($File)
    for ($i = 0; $i -le 400; $i++) {
        if (($objFolder.GetDetailsOf($objFile, $i)) -ne "") {
        $key = $objFolder.GetDetailsOf($objFile.items, $i)
        $value = $objFolder.GetDetailsOf($objFile, $i)
        Write-Host $i "-" $key ":" $value        
        }
    }
    #> #конец чтения