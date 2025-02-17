[cmdletbinding()]
Param([bool]$verbose)
$VerbosePreference = if ($verbose) { 'Continue' } else { 'SilentlyContinue' }
$ProgressPreference = "SilentlyContinue"

function Get-CacheFilename {
    param(
        [string]$cacheFolder,
        [string]$inputString,
        [string]$suffix
    )

    $inputStream = [IO.MemoryStream]::new([byte[]][char[]]$inputString)
    $fileName = (Get-FileHash -InputStream $inputStream -Algorithm SHA256).Hash, $suffix, "json" -join "."
    Join-Path -Path "$cacheFolder" -ChildPath "$fileName"
}