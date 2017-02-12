# Wallpaper Changer version 1.0 - release date: 12.02.2017

$WallpaperType = add-type @"
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;
namespace Wallpaper {
    public enum Style : int  {
        Tile, Center, Stretch, NoChange, Fit
    }
    public class Setter {
        public const int SetDesktopWallpaper = 20;
        public const int UpdateIniFile = 0x01;
        public const int SendWinIniChange = 0x02;
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int SystemParametersInfo (int uAction, int uParam, string lpvParam, int fuWinIni);
        public static void SetWallpaper ( string path, Wallpaper.Style style ) {
            SystemParametersInfo( SetDesktopWallpaper, 0, path, UpdateIniFile | SendWinIniChange );
            RegistryKey key = Registry.CurrentUser.OpenSubKey("Control Panel\\Desktop", true);
            switch( style ) {
                case Style.Stretch :
                key.SetValue(@"WallpaperStyle", "2");
                key.SetValue(@"TileWallpaper", "0");
                break;
                case Style.Center :
                key.SetValue(@"WallpaperStyle", "1");
                key.SetValue(@"TileWallpaper", "0");
                break;
                case Style.Tile :
                key.SetValue(@"WallpaperStyle", "1");
                key.SetValue(@"TileWallpaper", "1");
                break;
                case Style.Fit :
                key.SetValue(@"WallpaperStyle", "6");
                key.SetValue(@"TileWallpaper", "0");
                break;
                case Style.NoChange :
                break;
            }
            key.Close();
        }
    }
}
"@ -Passthru
$Pictures = "D:\OneDrive\Pictures"
$Girls = "Girls"
$WishList = "Celebrity", "Drive", "Graphics", "Nature", "Others", "Space", "Enclosure"
$BlackList = "Camera Roll", "LifeCam Files", "Some", "Twonky", "С проигрывателя Victor", "С телефона Windows Phone Lena", "From Victor", "Raptr Screenshots", "toSort", "Из Submarine", "Файлы LifeCam", "Saved Pictures", "Instagram", "Private", "Screenshots", "Пленка"
# check if files exist and setting folders list
if (Get-Item "D:\OneDrive\My Scripts\Wallpapers\addnude.flag") {
    $ExcludeList = $BlackList
    if (Get-Item "D:\OneDrive\My Scripts\Wallpapers\girlsonly.flag") {
        $ExcludeList = $BlackList + $WishList
        }
    }
else {
    $ExcludeList = $BlackList + $Girls
    }
$file = (Get-ChildItem $Pictures -Exclude $ExcludeList).GetFiles() | Get-Random
[Wallpaper.Setter]::SetWallpaper($file.FullName, "Fit")

<# History
 based on http://poshcode.org/2603
 version 1.0 - release date: 11.02.2017
    add - first release
 Email for support to Victor.V.Shepelin@live.ru
 #>