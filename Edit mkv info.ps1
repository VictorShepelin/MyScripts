#https://matroska.org/technical/specs/tagging/index.html
#https://matroska.org/technical/specs/tagging/example-video.html
#https://matroska.org/downloads/windows.html

# импортируем таг в матроску
$input = "C:\Videos\tosort\88.minut.2007.x264.BDRip.720p.mkv"
if ($input.Substring($input.length-3,3) -match "mkv") {
    Write-Host "mkv"
}
$output = $input + "_2.mkv"
$tag = "F:\Temp\tag.xml"
[xml]$content = Get-Content -Path $tag
$titleName = $content.Tags.Tag.Simple[0].String
& "C:\Program Files\mkvtoolnix\mkvmerge.exe" --output $output --global-tags $tag --title $titleName $input

# читаю таги из матроски
$out = & "C:\Program Files\mkvtoolnix\mkvinfo.exe" $output -t
$i = 0
$out | ForEach-Object {
    $i++
    if ($_ -match "Name: TITLE") {
        $string = $out[$i]
        $chararray = $string.ToCharArray()
        $j = 0
        $from = $null
        $chararray | ForEach-Object {
            if ($_ -match ':'){
                if ($from -eq $null) {
                    $from = $j+2
                }
            }
            $j++
        }
        $title = $string.Substring($from,(($chararray.count)-$from))
        $title
    }
    if ($_ -match "Name: ARTIST") {
        $string = $out[$i]
        $chararray = $string.ToCharArray()
        $j = 0
        $from = $null
        $chararray | ForEach-Object {
            if ($_ -match ':'){
                if ($from -eq $null) {
                    $from = $j+2
                }
            }
            $j++
        }
        $AlbumArtists = $string.Substring($from,(($chararray.count)-$from))
        $AlbumArtists
    }
    if ($_ -match "Name: DATE_RELEASED") {
        $string = $out[$i]
        $chararray = $string.ToCharArray()
        $j = 0
        $from = $null
        $chararray | ForEach-Object {
            if ($_ -match ':'){
                if ($from -eq $null) {
                    $from = $j+2
                }
            }
            $j++
        }
        $releaseDate = $string.Substring($from,(($chararray.count)-$from))
        $releaseDate
    }
    if ($_ -match "Name: COPYRIGHT") {
        $string = $out[$i]
        $chararray = $string.ToCharArray()
        $j = 0
        $from = $null
        $chararray | ForEach-Object {
            if ($_ -match ':'){
                if ($from -eq $null) {
                    $from = $j+2
                }
            }
            $j++
        }
        $url = $string.Substring($from,(($chararray.count)-$from))
        $url
    }
    if ($_ -match "Name: GENRE") {
        $string = $out[$i]
        $chararray = $string.ToCharArray()
        $j = 0
        $from = $null
        $chararray | ForEach-Object {
            if ($_ -match ':'){
                if ($from -eq $null) {
                    $from = $j+2
                }
            }
            $j++
        }
        $JoinGenres = $string.Substring($from,(($chararray.count)-$from))
        $JoinGenres
    }
}