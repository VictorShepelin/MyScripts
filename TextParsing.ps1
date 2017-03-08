$content = Get-Content D:\OneDrive\Desktop\test.txt
$content | ForEach-Object {
    $str = $_
    if ($str-match "Осн") {
        $str = $str -replace "Основной долг ", ""
        $str = $str -replace " + проценты ", ";"
        $str = $str -replace " RUR",""
        Write-Host $str
    }
    else {
    }
}