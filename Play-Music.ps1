Param(
    [ValidatePattern("^[A-G]+$")]
    $Music = "Default"
)
function Play-Music {
$ErrorActionPreference = "SilentlyContinue"
$hash = @{}
$hash."A" = '220;'
$hash."B" = '247;'
$hash."C" = '262;'
$hash."D" = '294;'
$hash."E" = '330;'
$hash."F" = '349;'
$hash."G" = '392;'
$hash."Default" = (Get-Random -Minimum 37 -Maximum 32767)
Foreach ($key in $hash.Keys)
 {
    $Music = $Music.Replace($key, $hash.$key)
 }
$tune = $music -Split(";")
Foreach ($t in $tune)
 {
    [Console]::Beep($t,400)
 }
}
Play-Music $Music