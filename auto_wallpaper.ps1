# Copyright (c) 2022 Cerallin <cerallin@cerallin.top>
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

# All these codes of "Add-Type" are copied from a website I forgot which...
Add-Type @"
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;
namespace Wallpaper
{
    public enum Style : int
    {
        Tile, Center, Stretch, NoChange
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
            switch( style )
            {
                case Style.Stretch :
                key.SetValue(@"WallpaperStyle", "2") ;
                key.SetValue(@"TileWallpaper", "0") ;
                break;
                case Style.Center :
                key.SetValue(@"WallpaperStyle", "1") ;
                key.SetValue(@"TileWallpaper", "0") ;
                break;
                case Style.Tile :
                key.SetValue(@"WallpaperStyle", "1") ;
                key.SetValue(@"TileWallpaper", "1") ;
                break;
                case Style.NoChange :
                break;
            }
            key.Close();
        }
    }
}
"@

$light_on_at = 360 # 6:00
$light_off_at = 1260 # 21:00

$day_wallpaper = "c:\users\julia\pictures\day.png"
$night_wallpaper = "c:\users\julia\pictures\night.png"

$time_arr = (Get-Date -Format H:mm).Split(":")
$time = [int]$time_arr[0] * 60 + [int]$time_arr[1]

if ($time -lt $light_off_at -AND $time -gt $light_on_at) {
    [Wallpaper.Setter]::SetWallpaper( $day_wallpaper, 3 )
}
Else {
    [Wallpaper.Setter]::SetWallpaper( $night_wallpaper, 3 )
}
