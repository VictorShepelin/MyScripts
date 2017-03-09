$Folder = "C:\Videos\Horror"
$File = "Krik.2.1997.DUAL.BDRip.XviD.AC3.-HQCLUB.avi"
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
[System.Reflection.Assembly]::LoadFile("F:\Downloads\taglib-sharp.dll") #https://www.nuget.org/packages/taglib/2.1.0
$mediafile = [TagLib.File]::Create("C:\Videos\Horror\Krik.2.1997.DUAL.BDRip.XviD.AC3.-HQCLUB.avi")
$mediafile.Writeable
#$mediafile.Tag
$mediafile.Tag.AlbumArtists = "Neve Campbell, Courteney Cox, David Arquette" # в ролях 13ый таг
$mediafile.Tag.Year = "1997" # год выхода 15ый таг
$mediafile.Tag.Genres = "Horror, Mystery" # жанр 16ый таг
#$mediafile.Tag.Rating = "6" # рейтинг 19ый таг
$mediafile.Tag.Title = "Scream 2" # название 21ый таг
$mediafile.Tag.Comment = "" # коммент 24ый таг
$mediafile.Tag.copyright = "http://www.imdb.com/title/tt0120082/" # ссылка на imdb 25ый таг
$mediafile.Save()

$mediafile.Tag.Title
$mediafile.Tag.Year
$mediafile.Tag.Genres
$mediafile.Tag.AlbumArtists
$mediafile.Tag.Comment
$mediafile.Tag.copyright