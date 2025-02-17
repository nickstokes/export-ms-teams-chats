[cmdletbinding()]
Param([bool]$verbose)
$VerbosePreference = if ($verbose) { 'Continue' } else { 'SilentlyContinue' }
$ProgressPreference = "SilentlyContinue"
$brokenLink = "iVBORw0KGgoAAAANSUhEUgAAABoAAAAaCAYAAACpSkzOAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAMNSURBVEhL3VRLTxNRGD13poV2qBStjSj1BSIkGnzFhIUxVfQXkIKRf2DcuNK4cG0ixoUbE1cQjZoY4wITYgQTkMRoNJEYRaBB0RrF0tIyfU9n/O501BamYzXCwjO5mXvv3PnO9zj3YxoBqwDBeK84/j8i0xr9adkYY8asPEyJhh4NITw/Z6ysYbPZ0NkZMFblYUo0MDAAx9rtcLu9hQ06ogkiHVYLawMqEzDz5gkELQO//wi8Xm/Z6CxrpJEhPsDsugGBVdFcpD0+BCKg78QdCHRhcHAQsiwbfy5HBWJg0KCgySehrWUNqu25n5FpggYKVHeiu7sbw8PD+r4ZLFNXW1ePKjvQ2OBCbQ1FQ6RkHuPBKNJpFUwQMDv5DCk5VohYZEgnYgh0nYDL5dJt/YBlRIwcb91aB2e1HTlFo6HSyPOc6t81VcOW5oPYufcYWvYfR9Ouw1BUc5PWNaLUvA4uYHwqWjRiSGd46gwyIx9qXtWJy+E3NWLg/5KNkvE3qEAMvJBmki323pjTsWpKs8CVugQVEWlULGbkPmFLIstIeSrJrQjcmSq7iJ6ek5BqJGP3Fyog4t7yO5TH18xHXJ8+jdG5m1CEPJesES2D0wm0Na+DKJhFX2nqqFDzWhh3Pl2EQoSv5BGEHbPIaTn6qkFy2rBjsxuiyKM0N2lNRBeTp0gWFzEUuYGsTSbf6QKTGvsnzuOzGoSDItnd6IajirqFlkd8MQ5VXa4YSyLey0LpIO6GLuF9YpwnUN+noCiVdryM3sLGer5TEEMklkBfXz+SyaS+LkZ5InJKzs/h3pfL+Jab1SPhjw56+aRNOLXnHDySRz88FZnATEiGw7FcCBxliZJsAY+jt5EVk3SoVGFNzlb0+q+hoc5HK2pJ4bfofX4BH5RJiLydmKAs0YvQA6SiEfiy20rGofV+nGk/W4iP+tu74ASuPrwCV2oDnk7fh81JzfHHvSqCaVMdGxtDPC5DXHJXOA6074PHw9NVwOjIKFKLGX2eJ8krSg4dHUchSaUpNCVaCViq7l9ilYiA7y+DOI89Q2alAAAAAElFTkSuQmCC"

function Get-Image ($imageTagMatch, $assetsFolderPath, $clientId, $tenantId) {
    $imageUriPath = $imageTagMatch.Groups[1].Value
    $imageUriPathStream = [IO.MemoryStream]::new([byte[]][char[]]$imageUriPath)
    $imageFileName = "$((Get-FileHash -InputStream $imageUriPathStream -Algorithm SHA256).Hash).jpg"
    $imageFilePath = Join-Path -Path "$assetsFolderPath" -ChildPath "$imageFileName"
    if (-not(Test-Path $imageFilePath)) {
        Write-Verbose "Image cache miss, downloading."

        $imageUri = "https://graph.microsoft.com" + $imageUriPath
        
        try {
            $start = Get-Date
            Invoke-Retry -Code {
                Invoke-WebRequest -Uri $imageUri -Headers @{
                    "Authorization" = "Bearer $(Get-GraphAccessToken $clientId $tenantId)"
                } -OutFile $imageFilePath
            }

            Write-Verbose "Took $(((Get-Date) - $start).TotalSeconds)s to download image."
            $image = "assets/$imageFileName"
        }
        catch {
            Write-Verbose "Failed to fetch image, using broken link."

            [IO.File]::WriteAllBytes($imageFilePath, [Convert]::FromBase64String($brokenLink))
            $image = "assets/$imageFileName"
        }
    }
    else {
        Write-Verbose "Image cache hit."
        $image = "assets/$imageFileName"
    }

    $image
}