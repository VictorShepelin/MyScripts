$TagLibDll = "D:\OneDrive\Tools\Multimedia\taglib-sharp.dll" #https://www.nuget.org/packages/taglib/2.1.0
[System.Reflection.Assembly]::LoadFile($TagLibDll) | Out-Null
$files = Get-ChildItem -Path "C:\Videos\Movies" -Filter "*.mkv"
$files | ForEach-Object {
    $movieFile = $_.FullName
    # читаю таги из матроски
    $out = & "C:\Program Files\mkvtoolnix\mkvinfo.exe" $movieFile -t
    $i = 0
    $out| ForEach-Object {
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
            $JoinGenres = $JoinGenres.Substring(0,($JoinGenres.ToCharArray().count-1)) -replace ";", " "
        }
    }
    if (Test-Path "C:\Videos\Movies\$JoinGenres") {
        }
    else {
        New-Item -Path "C:\Videos\Movies\$JoinGenres" -ItemType Directory
        }
    Move-Item -Path $movieFile -Destination "C:\Videos\Movies\$JoinGenres"
}