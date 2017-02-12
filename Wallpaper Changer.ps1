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
$WishList = "Celebrity", "Drive", "Graphics", "Nature", "Others", "Space", "Enclosure"
$file = (Get-ChildItem $Pictures | Where-Object {
    if (Test-Path "D:\OneDrive\My Scripts\addnude.flag") {
        $WishList += "Girls"
    }
    if (Test-Path "D:\OneDrive\My Scripts\girlsonly.flag") {
        $WishList = "Girls"
    }
    $WishList -contains $_.BaseName
}).GetFiles() | Get-Random
[Wallpaper.Setter]::SetWallpaper($file.FullName, "Fit")
<# History
 based on http://poshcode.org/2603
 version 1.0 - release date: 12.02.2017
    add - first release
 Email for support to Victor.V.Shepelin@live.ru
 #>