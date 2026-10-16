#Requires -Version 5.1
<#
.SYNOPSIS
    End-to-end tests for Invoke-GPhotosMetadataFix.ps1 and Restore-ExifToolOriginals.ps1.

.DESCRIPTION
    Builds a fake Takeout folder from the tiny files in tests\fixtures, runs the
    scripts against it with a real ExifTool, and checks the results. Everything
    happens in a temporary folder that is deleted afterwards.

.EXAMPLE
    .\tests\Invoke-Tests.ps1 -ExifToolPath 'C:\Tools\exiftool.exe'
#>
param(
    [string]$ExifToolPath = 'exiftool'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$fixScript = Join-Path $repoRoot 'Invoke-GPhotosMetadataFix.ps1'
$restoreScript = Join-Path $repoRoot 'Restore-ExifToolOriginals.ps1'
$fixtures = Join-Path $PSScriptRoot 'fixtures'
$exifTool = (Get-Command -Name $ExifToolPath -CommandType Application | Select-Object -First 1).Path
$utf8 = New-Object System.Text.UTF8Encoding($false)
# So non-ASCII file names survive the trip to and from ExifTool in Windows PowerShell.
$OutputEncoding = $utf8
try { [Console]::OutputEncoding = $utf8 } catch { Write-Verbose 'No console to set the encoding on.' }
$work = Join-Path ([IO.Path]::GetTempPath()) ('gphotos-tests-' + [guid]::NewGuid().ToString('N'))
$src = Join-Path $work 'Takeout'
$album = Join-Path $src 'Photos from 2020'

# A fixed zone keeps the expected values stable (Windows PowerShell only knows Windows IDs).
$zone = 'America/New_York'
try { [void][TimeZoneInfo]::FindSystemTimeZoneById($zone) } catch { $zone = 'Eastern Standard Time' }

$script:passed = 0
$script:failed = 0

function Assert-Equal {
    param($Actual, $Expected, [string]$Message)
    if ([string]$Actual -ceq [string]$Expected) {
        $script:passed++
        Write-Host "  PASS  $Message"
    } else {
        $script:failed++
        Write-Host "  FAIL  $Message`n        expected: [$Expected]`n        actual:   [$Actual]" -ForegroundColor Red
    }
}

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if ($Condition) { Assert-Equal 'true' 'true' $Message } else { Assert-Equal 'false' 'true' $Message }
}

function New-Media {
    param([string]$Name, [string]$Fixture, [string[]]$Tags = @(), [string]$Folder = $album)
    [void][IO.Directory]::CreateDirectory($Folder)
    $path = Join-Path $Folder $Name
    [IO.File]::Copy((Join-Path $fixtures $Fixture), $path, $true)
    if ($Tags.Count) { & $exifTool -q -overwrite_original @Tags $path; if ($LASTEXITCODE) { throw "fixture setup failed: $Name" } }
    return $path
}

function New-Sidecar {
    param([string]$Name, [long]$Taken, [double[]]$Geo = @(0, 0, 0), [string]$Description = '', [switch]$NoGeo, [string]$Folder = $album)
    $data = [ordered]@{
        title          = $Name
        description    = $Description
        photoTakenTime = [ordered]@{ timestamp = [string]$Taken; formatted = 'ignored' }
        creationTime   = [ordered]@{ timestamp = [string]($Taken + 86400); formatted = 'ignored' }
    }
    if (-not $NoGeo) { $data.geoData = [ordered]@{ latitude = $Geo[0]; longitude = $Geo[1]; altitude = $Geo[2] } }
    [IO.File]::WriteAllText((Join-Path $Folder $Name), ($data | ConvertTo-Json -Depth 5), $utf8)
}

function ConvertTo-Unix {
    param([string]$UtcText)
    $utc = [DateTimeOffset]::ParseExact($UtcText, 'yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::AssumeUniversal)
    return [long]($utc - (New-Object DateTimeOffset(1970, 1, 1, 0, 0, 0, [TimeSpan]::Zero))).TotalSeconds
}

function Get-Meta {
    param([string]$Path)
    $json = $Path | & $exifTool -json -n -G1 -charset filename=utf8 -@ - | Out-String
    return @($json | ConvertFrom-Json)[0]
}

function Get-Tag {
    param($Meta, [string]$Name)
    $property = $Meta.PSObject.Properties[$Name]
    if ($property) { return [string]$property.Value }
    return ''
}

function Invoke-Fix {
    param([string]$Destination, [hashtable]$Extra = @{})
    $global:LASTEXITCODE = 0
    $parameters = @{ SourcePath = $src; DestinationPath = $Destination; ExifToolPath = $exifTool; TimeZone = $zone }
    foreach ($key in $Extra.Keys) { $parameters[$key] = $Extra[$key] }
    $output = & $fixScript @parameters *>&1 | Out-String
    return [pscustomobject]@{ ExitCode = $global:LASTEXITCODE; Output = $output }
}

function Get-Report {
    param([string]$Destination)
    $file = Get-ChildItem -LiteralPath $Destination -Filter 'gphotos-metadata-report-*.csv' | Sort-Object LastWriteTime | Select-Object -Last 1
    $rows = @{}
    foreach ($row in (Import-Csv -LiteralPath $file.FullName)) { $rows[$row.File.Replace('\', '/')] = $row }
    return $rows
}

function Get-TreeHashes {
    param([string]$Root)
    $hashes = @{}
    foreach ($file in Get-ChildItem -LiteralPath $Root -Recurse -File) {
        $hashes[$file.FullName] = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
    }
    return $hashes
}

try {
    # -----------------------------------------------------------------------
    Write-Host 'Building fake Takeout folder...'
    # -----------------------------------------------------------------------
    # Camera photo taken 18:30 in New York (22:30 UTC): same moment, keep camera value.
    New-Media 'camera.jpg' 'sample.jpg' @('-DateTimeOriginal=2019:07:04 18:30:00') | Out-Null
    New-Sidecar 'camera.jpg.supplemental-metadata.json' (ConvertTo-Unix '2019-07-04 22:30:00')
    # Taken 12:00 in Paris (10:00 UTC): kept by default, re-expressed in New York time with -Overwrite.
    New-Media 'paris.jpg' 'sample.jpg' @('-DateTimeOriginal=2019:07:04 12:00:00') | Out-Null
    New-Sidecar 'paris.jpg.supplemental-metadata.json' (ConvertTo-Unix '2019-07-04 10:00:00')
    # Date corrected in Google Photos: Google's value replaces the camera's.
    New-Media 'corrected.jpg' 'sample.jpg' @('-DateTimeOriginal=2015:05:05 12:00:00') | Out-Null
    New-Sidecar 'corrected.jpg.supplemental-metadata.json' (ConvertTo-Unix '2001-02-03 17:00:00')
    # Southern / western hemispheres.
    New-Media 'sydney.jpg' 'sample.jpg' | Out-Null
    New-Sidecar 'sydney.jpg.supplemental-metadata.json' (ConvertTo-Unix '2023-11-14 22:13:20') -Geo @(-33.8688, 151.2093, 58)
    New-Media 'lima.jpg' 'sample.jpg' | Out-Null
    New-Sidecar 'lima.jpg.supplemental-metadata.json' (ConvertTo-Unix '2023-11-14 22:13:20') -Geo @(-12.0464, -77.0428, 0)
    # Camera GPS + sidecar without geoData: GPS must survive.
    New-Media 'phone_gps.jpg' 'sample.jpg' @('-DateTimeOriginal=2018:05:05 06:00:00', '-GPSLatitude=40.7', '-GPSLatitudeRef=N', '-GPSLongitude=74', '-GPSLongitudeRef=W') | Out-Null
    New-Sidecar 'phone_gps.jpg.supplemental-metadata.json' (ConvertTo-Unix '2018-05-05 10:00:00') -NoGeo
    # Camera GPS vs. Google GPS 30 m away (keep) and 30 km away (edited in Google: replace).
    New-Media 'gps_near.jpg' 'sample.jpg' @('-GPSLatitude=40.7128', '-GPSLatitudeRef=N', '-GPSLongitude=74.006', '-GPSLongitudeRef=W') | Out-Null
    New-Sidecar 'gps_near.jpg.supplemental-metadata.json' (ConvertTo-Unix '2018-05-05 10:00:00') -Geo @(40.7130, -74.0062, 0)
    New-Media 'gps_far.jpg' 'sample.jpg' @('-GPSLatitude=40.7128', '-GPSLatitudeRef=N', '-GPSLongitude=74.006', '-GPSLongitudeRef=W') | Out-Null
    New-Sidecar 'gps_far.jpg.supplemental-metadata.json' (ConvertTo-Unix '2018-05-05 10:00:00') -Geo @(41.0, -74.0, 0)
    # IMG_5 has no sidecar; it must not borrow IMG_50's.
    New-Media 'IMG_5.jpg' 'sample.jpg' | Out-Null
    New-Media 'IMG_50.jpg' 'sample.jpg' | Out-Null
    New-Sidecar 'IMG_50.jpg.supplemental-metadata.json' (ConvertTo-Unix '2010-01-01 00:00:00') -Geo @(51.5, -0.12, 0)
    # "data" appears inside "supplemental-metadata": must not match anything.
    New-Media 'data.jpg' 'sample.jpg' | Out-Null
    # Duplicate names.
    New-Media 'IMG_7.jpg' 'sample.jpg' | Out-Null
    New-Media 'IMG_7(1).jpg' 'sample.jpg' | Out-Null
    New-Media 'IMG_7(2).jpg' 'sample.jpg' | Out-Null
    New-Sidecar 'IMG_7.jpg.supplemental-metadata.json' (ConvertTo-Unix '2015-06-01 12:00:00')
    New-Sidecar 'IMG_7.jpg.supplemental-metadata(1).json' (ConvertTo-Unix '2016-06-01 12:00:00')
    New-Sidecar 'IMG_7.jpg.supplemental-metadata(2).json' (ConvertTo-Unix '2017-06-01 12:00:00')
    # Edited copy shares the original's sidecar.
    New-Media 'IMG_9.jpg' 'sample.jpg' | Out-Null
    New-Media 'IMG_9-edited.jpg' 'sample.jpg' | Out-Null
    New-Sidecar 'IMG_9.jpg.supplemental-metadata.json' (ConvertTo-Unix '2011-03-13 07:06:40')
    # Older Takeout naming.
    New-Media 'IMG_400.jpg' 'sample.jpg' | Out-Null
    New-Sidecar 'IMG_400.jpg.json' (ConvertTo-Unix '2012-01-01 12:00:00')
    New-Media 'IMG_401.jpg' 'sample.jpg' | Out-Null
    New-Media 'IMG_401(1).jpg' 'sample.jpg' | Out-Null
    New-Sidecar 'IMG_401.jpg.json' (ConvertTo-Unix '2012-02-01 12:00:00')
    New-Sidecar 'IMG_401.jpg(1).json' (ConvertTo-Unix '2012-03-01 12:00:00')
    New-Media 'IMG_402.jpg' 'sample.jpg' | Out-Null
    New-Sidecar 'IMG_402.json' (ConvertTo-Unix '2012-04-01 12:00:00')
    # Long name: sidecar name truncated to 51 characters.
    $longName = 'PXL_20230615_123456789.PORTRAIT.ORIGINAL.jpg'
    New-Media $longName 'sample.jpg' | Out-Null
    New-Sidecar (($longName + '.supplemental-metadata').Substring(0, 46) + '.json') (ConvertTo-Unix '2023-06-15 12:34:56')
    # Brackets in the name (wildcards in PowerShell paths).
    New-Media 'party[1].jpg' 'sample.jpg' | Out-Null
    New-Sidecar 'party[1].jpg.supplemental-metadata.json' (ConvertTo-Unix '2020-09-13 12:26:40')
    # Non-ASCII file name and caption with quotes, $, @ and a line break.
    $unicodeName = 'caf' + [char]0x00E9 + ' ' + [char]0x2615 + '.jpg'
    $caption = 'Mom said "cheese" - $5 @home' + "`n" + 'Caf' + [char]0x00E9 + ' ' + [char]0x2615
    New-Media $unicodeName 'sample.jpg' | Out-Null
    New-Sidecar ($unicodeName + '.supplemental-metadata.json') (ConvertTo-Unix '2020-09-13 12:26:40') -Description $caption
    # PNG, GIF and video.
    New-Media 'Screenshot_20200101-101010.png' 'sample.png' | Out-Null
    New-Sidecar 'Screenshot_20200101-101010.png.supplemental-metadata.json' (ConvertTo-Unix '2020-01-01 15:10:10')
    New-Media 'anim.gif' 'sample.gif' | Out-Null
    New-Sidecar 'anim.gif.supplemental-metadata.json' (ConvertTo-Unix '2020-01-01 15:10:10') -Geo @(-12.5, -77.25, 0)
    New-Media 'clip.mp4' 'sample.mp4' | Out-Null
    New-Sidecar 'clip.mp4.supplemental-metadata.json' (ConvertTo-Unix '2020-09-13 12:26:40') -Geo @(48.8584, 2.2945, 35) -Description 'Eiffel'
    # Live photo: the video has no sidecar of its own.
    New-Media 'IMG_300.jpg' 'sample.jpg' | Out-Null
    New-Media 'IMG_300.mp4' 'sample.mp4' | Out-Null
    New-Sidecar 'IMG_300.jpg.supplemental-metadata.json' (ConvertTo-Unix '2021-07-01 16:00:00')
    # No sidecar, date in the file name.
    New-Media 'IMG_20200101_101010.jpg' 'sample.jpg' | Out-Null
    # Motion photo: JPEG with an MP4 appended after the image data.
    $motion = Join-Path $album 'PXL_20220101_120000000.MP.jpg'
    [IO.File]::WriteAllBytes($motion, [IO.File]::ReadAllBytes((Join-Path $fixtures 'sample.jpg')) + [IO.File]::ReadAllBytes((Join-Path $fixtures 'sample.mp4')))
    New-Sidecar 'PXL_20220101_120000000.MP.jpg.supplemental-metadata.json' (ConvertTo-Unix '2022-01-01 17:00:00')
    # PNG data saved with a .jpg extension: ExifTool refuses to write, file is still copied.
    New-Media 'wrong_extension.jpg' 'sample.png' | Out-Null
    New-Sidecar 'wrong_extension.jpg.supplemental-metadata.json' (ConvertTo-Unix '2020-01-01 00:00:00')
    # Album metadata and a non-media file.
    [IO.File]::WriteAllText((Join-Path $album 'metadata.json'), '{"title":"Trip","description":"ALBUM"}', $utf8)
    [IO.File]::WriteAllText((Join-Path $src 'archive_browser.html'), '<html></html>', $utf8)

    $sourceBefore = Get-TreeHashes $src

    # -----------------------------------------------------------------------
    Write-Host "`nSafety checks"
    # -----------------------------------------------------------------------
    $overlapError = ''
    try { Invoke-Fix (Join-Path $src 'out') | Out-Null } catch { $overlapError = $_.Exception.Message }
    Assert-True ($overlapError -like '*must not be inside*') 'refuses a destination inside the source'

    $whatIfDest = Join-Path $work 'whatif'
    $run = Invoke-Fix $whatIfDest @{ WhatIf = $true }
    Assert-True (-not (Test-Path -LiteralPath $whatIfDest)) '-WhatIf writes nothing'

    # -----------------------------------------------------------------------
    Write-Host "`nMain run"
    # -----------------------------------------------------------------------
    $dest = Join-Path $work 'Photos'
    $destAlbum = Join-Path $dest 'Photos from 2020'
    $run = Invoke-Fix $dest
    Assert-Equal $run.ExitCode 0 'exit code 0'
    $report = Get-Report $dest

    $sourceAfter = Get-TreeHashes $src
    $unchanged = ($sourceBefore.Count -eq $sourceAfter.Count)
    foreach ($key in $sourceBefore.Keys) { if ($sourceAfter[$key] -ne $sourceBefore[$key]) { $unchanged = $false } }
    Assert-True $unchanged 'source folder is byte-for-byte unchanged'
    Assert-Equal @(Get-ChildItem -LiteralPath $work -Recurse -File -Filter '*_original').Count 0 'no *_original files anywhere'
    Assert-Equal @(Get-ChildItem -LiteralPath $work -Recurse -File -Filter '*gphotos-partial*').Count 0 'no temporary files left behind'
    $expectedCopies = @(Get-ChildItem -LiteralPath $src -Recurse -File | Where-Object { $_.Extension -ne '.json' }).Count
    $actualCopies = @(Get-ChildItem -LiteralPath $dest -Recurse -File | Where-Object { $_.Name -notlike 'gphotos-metadata-report-*' }).Count
    Assert-Equal $actualCopies $expectedCopies 'every non-JSON file was copied'
    Assert-True (Test-Path -LiteralPath (Join-Path $dest 'archive_browser.html')) 'non-media files are copied'

    $m = Get-Meta (Join-Path $destAlbum 'camera.jpg')
    Assert-Equal (Get-Tag $m 'ExifIFD:DateTimeOriginal') '2019:07:04 18:30:00' 'camera date of the same moment is kept'
    Assert-Equal $report['Photos from 2020/camera.jpg'].Status 'Copied' 'camera.jpg needed no changes'
    $m = Get-Meta (Join-Path $destAlbum 'paris.jpg')
    Assert-Equal (Get-Tag $m 'ExifIFD:DateTimeOriginal') '2019:07:04 12:00:00' 'local time of a photo taken in another time zone is kept'
    $m = Get-Meta (Join-Path $destAlbum 'corrected.jpg')
    Assert-Equal (Get-Tag $m 'ExifIFD:DateTimeOriginal') '2001:02:03 12:00:00' 'date corrected in Google Photos replaces the camera date'
    Assert-Equal (Get-Tag $m 'ExifIFD:OffsetTimeOriginal') '-05:00' 'UTC offset written with the date'
    Assert-Equal $report['Photos from 2020/corrected.jpg'].Changes 'date replaced' 'report says the date was replaced'

    $m = Get-Meta (Join-Path $destAlbum 'sydney.jpg')
    Assert-Equal ('{0},{1}' -f (Get-Tag $m 'Composite:GPSLatitude'), (Get-Tag $m 'Composite:GPSLongitude')) '-33.8688,151.2093' 'southern hemisphere GPS'
    Assert-Equal (Get-Tag $m 'GPS:GPSAltitude') '58' 'GPS altitude'
    $m = Get-Meta (Join-Path $destAlbum 'lima.jpg')
    Assert-Equal ('{0},{1}' -f (Get-Tag $m 'Composite:GPSLatitude'), (Get-Tag $m 'Composite:GPSLongitude')) '-12.0464,-77.0428' 'western hemisphere GPS'
    $m = Get-Meta (Join-Path $destAlbum 'phone_gps.jpg')
    Assert-Equal ('{0},{1}' -f (Get-Tag $m 'Composite:GPSLatitude'), (Get-Tag $m 'Composite:GPSLongitude')) '40.7,-74' 'camera GPS survives a sidecar without geoData'
    Assert-Equal (Get-Tag $m 'ExifIFD:DateTimeOriginal') '2018:05:05 06:00:00' 'camera date survives too'
    $m = Get-Meta (Join-Path $destAlbum 'gps_near.jpg')
    Assert-Equal (Get-Tag $m 'Composite:GPSLatitude') '40.7128' 'camera GPS kept when Google agrees within 1 km'
    $m = Get-Meta (Join-Path $destAlbum 'gps_far.jpg')
    Assert-Equal (Get-Tag $m 'Composite:GPSLatitude') '41' 'location edited in Google Photos replaces camera GPS'

    $m = Get-Meta (Join-Path $destAlbum 'IMG_5.jpg')
    Assert-Equal (Get-Tag $m 'ExifIFD:DateTimeOriginal') '' 'IMG_5.jpg does not borrow IMG_50 sidecar'
    Assert-Equal $report['Photos from 2020/IMG_5.jpg'].Sidecar '' 'IMG_5.jpg has no sidecar in the report'
    $m = Get-Meta (Join-Path $destAlbum 'data.jpg')
    Assert-Equal (Get-Tag $m 'ExifIFD:DateTimeOriginal') '' 'data.jpg does not match supplemental-metadata files'
    $m = Get-Meta (Join-Path $destAlbum 'IMG_50.jpg')
    Assert-Equal (Get-Tag $m 'ExifIFD:DateTimeOriginal') '2009:12:31 19:00:00' 'IMG_50.jpg uses its own sidecar'

    Assert-Equal (Get-Tag (Get-Meta (Join-Path $destAlbum 'IMG_7.jpg')) 'ExifIFD:DateTimeOriginal') '2015:06:01 08:00:00' 'IMG_7.jpg'
    Assert-Equal (Get-Tag (Get-Meta (Join-Path $destAlbum 'IMG_7(1).jpg')) 'ExifIFD:DateTimeOriginal') '2016:06:01 08:00:00' 'IMG_7(1).jpg uses the (1) sidecar'
    Assert-Equal (Get-Tag (Get-Meta (Join-Path $destAlbum 'IMG_7(2).jpg')) 'ExifIFD:DateTimeOriginal') '2017:06:01 08:00:00' 'IMG_7(2).jpg uses the (2) sidecar'
    Assert-Equal (Get-Tag (Get-Meta (Join-Path $destAlbum 'IMG_9-edited.jpg')) 'ExifIFD:DateTimeOriginal') '2011:03:13 03:06:40' 'edited copy uses the original sidecar'
    Assert-Equal (Get-Tag (Get-Meta (Join-Path $destAlbum 'IMG_400.jpg')) 'ExifIFD:DateTimeOriginal') '2012:01:01 07:00:00' 'old style name.jpg.json'
    Assert-Equal (Get-Tag (Get-Meta (Join-Path $destAlbum 'IMG_401(1).jpg')) 'ExifIFD:DateTimeOriginal') '2012:03:01 07:00:00' 'old style name.jpg(1).json'
    Assert-Equal (Get-Tag (Get-Meta (Join-Path $destAlbum 'IMG_402.jpg')) 'ExifIFD:DateTimeOriginal') '2012:04:01 08:00:00' 'old style name.json'
    Assert-Equal (Get-Tag (Get-Meta (Join-Path $destAlbum $longName)) 'ExifIFD:DateTimeOriginal') '2023:06:15 08:34:56' 'truncated sidecar name'
    Assert-Equal (Get-Tag (Get-Meta (Join-Path $destAlbum 'party[1].jpg')) 'ExifIFD:DateTimeOriginal') '2020:09:13 08:26:40' 'brackets in the file name'

    $m = Get-Meta (Join-Path $destAlbum $unicodeName)
    Assert-Equal (Get-Tag $m 'ExifIFD:DateTimeOriginal') '2020:09:13 08:26:40' 'non-ASCII file name'
    Assert-Equal (Get-Tag $m 'XMP-dc:Description') $caption 'caption with quotes, $, @, line break and non-ASCII is exact'
    Assert-Equal (Get-Tag $m 'IFD0:ImageDescription') $caption 'caption also in EXIF ImageDescription'

    $m = Get-Meta (Join-Path $destAlbum 'Screenshot_20200101-101010.png')
    Assert-Equal (Get-Tag $m 'ExifIFD:DateTimeOriginal') '2020:01:01 10:10:10' 'PNG date'
    $m = Get-Meta (Join-Path $destAlbum 'anim.gif')
    Assert-Equal (Get-Tag $m 'XMP-exif:DateTimeOriginal') '2020:01:01 10:10:10-05:00' 'GIF date (XMP)'
    Assert-Equal (Get-Tag $m 'XMP-exif:GPSLatitude') '-12.5' 'GIF GPS (XMP)'
    $m = Get-Meta (Join-Path $destAlbum 'clip.mp4')
    Assert-Equal (Get-Tag $m 'QuickTime:CreateDate') '2020:09:13 12:26:40' 'video CreateDate is UTC'
    Assert-Equal (Get-Tag $m 'Keys:CreationDate') '2020:09:13 08:26:40-04:00' 'video CreationDate has local time and offset'
    Assert-Equal ('{0},{1}' -f (Get-Tag $m 'Composite:GPSLatitude'), (Get-Tag $m 'Composite:GPSLongitude')) '48.8584,2.2945' 'video GPS'
    Assert-Equal (Get-Tag $m 'ItemList:Comment') 'Eiffel' 'video description'
    $m = Get-Meta (Join-Path $destAlbum 'IMG_300.mp4')
    Assert-Equal (Get-Tag $m 'QuickTime:CreateDate') '2021:07:01 16:00:00' 'live photo video uses the photo sidecar'
    Assert-Equal $report['Photos from 2020/IMG_300.mp4'].SidecarMatch 'live photo companion' 'live photo match reported'

    $m = Get-Meta (Join-Path $destAlbum 'IMG_20200101_101010.jpg')
    Assert-Equal (Get-Tag $m 'ExifIFD:DateTimeOriginal') '2020:01:01 10:10:10' 'date from file name when there is no sidecar'
    Assert-Equal (Get-Tag $m 'ExifIFD:OffsetTimeOriginal') '' 'no UTC offset invented for file name dates'

    $copy = [IO.File]::ReadAllBytes((Join-Path $destAlbum 'PXL_20220101_120000000.MP.jpg'))
    $video = [IO.File]::ReadAllBytes((Join-Path $fixtures 'sample.mp4'))
    $tail = New-Object byte[] $video.Length
    [Array]::Copy($copy, $copy.Length - $video.Length, $tail, 0, $video.Length)
    Assert-Equal ([Convert]::ToBase64String($tail)) ([Convert]::ToBase64String($video)) 'motion photo video is preserved'

    $row = $report['Photos from 2020/wrong_extension.jpg']
    Assert-Equal $row.Status 'CopiedWithoutMetadata' 'unwritable file is still copied'
    Assert-Equal (Get-FileHash -LiteralPath (Join-Path $destAlbum 'wrong_extension.jpg')).Hash (Get-FileHash -LiteralPath (Join-Path $album 'wrong_extension.jpg')).Hash 'unwritable file copied byte-for-byte'
    Assert-True ($run.Output -like '*wrong_extension.jpg*copied without metadata*') 'warning printed for unwritable file'

    Assert-Equal $report['Photos from 2020/sydney.jpg'].Verified 'image data identical' 'copy verified against source'
    $expectedTime = [DateTimeOffset]::FromUnixTimeSeconds((ConvertTo-Unix '2023-11-14 22:13:20')).UtcDateTime
    Assert-Equal ([IO.File]::GetLastWriteTimeUtc((Join-Path $destAlbum 'sydney.jpg'))) $expectedTime 'file modified date set to the date taken'

    # -----------------------------------------------------------------------
    Write-Host "`nRe-runs and options"
    # -----------------------------------------------------------------------
    $run = Invoke-Fix $dest
    $report = Get-Report $dest
    Assert-Equal @($report.Values | Where-Object { $_.Status -ne 'Skipped' }).Count 0 'second run skips everything already copied'

    $run = Invoke-Fix $dest @{ Force = $true; Overwrite = $true; IncludeSidecars = $true }
    $report = Get-Report $dest
    Assert-Equal $report['Photos from 2020/paris.jpg'].Changes 'date replaced' '-Overwrite writes Google date even when it is the same moment'
    Assert-Equal (Get-Tag (Get-Meta (Join-Path $destAlbum 'paris.jpg')) 'ExifIFD:DateTimeOriginal') '2019:07:04 06:00:00' '-Overwrite uses the chosen time zone'
    Assert-True (Test-Path -LiteralPath (Join-Path $destAlbum 'IMG_7.jpg.supplemental-metadata.json')) '-IncludeSidecars copies JSON files'
    Assert-Equal @(Get-ChildItem -LiteralPath $work -Recurse -File -Filter '*gphotos-partial*').Count 0 'no temporary files after -Force'

    # -----------------------------------------------------------------------
    Write-Host "`nRestore-ExifToolOriginals.ps1"
    # -----------------------------------------------------------------------
    $restoreDir = Join-Path $work 'restore'
    $original = New-Media 'a.jpg' 'sample.jpg' -Folder $restoreDir
    $originalHash = (Get-FileHash -LiteralPath $original).Hash
    & $exifTool -q '-DateTimeOriginal=2000:01:01 00:00:00' $original   # in place, leaves a.jpg_original
    Assert-True (Test-Path -LiteralPath "${original}_original") 'setup: ExifTool left a.jpg_original'
    & $restoreScript -Path $restoreDir -WhatIf *>&1 | Out-Null
    Assert-True (Test-Path -LiteralPath "${original}_original") 'restore -WhatIf changes nothing'
    & $restoreScript -Path $restoreDir *>&1 | Out-Null
    Assert-Equal (Get-FileHash -LiteralPath $original).Hash $originalHash 'original restored under its real name'
    Assert-True (-not (Test-Path -LiteralPath "${original}_original")) 'backup file is gone'
    Assert-True (Test-Path -LiteralPath (Join-Path $restoreDir '_exiftool-modified/a.jpg')) 'edited copy moved aside'

    # -----------------------------------------------------------------------
    Write-Host "`nWindows PowerShell 5.1 compatibility"
    # -----------------------------------------------------------------------
    foreach ($file in @($fixScript, $restoreScript, $PSCommandPath)) {
        $nonAscii = @([IO.File]::ReadAllBytes($file) | Where-Object { $_ -gt 127 }).Count
        Assert-Equal $nonAscii 0 "$([IO.Path]::GetFileName($file)) is ASCII-only (parses the same in any code page)"
        $tokens = $null; $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($file, [ref]$tokens, [ref]$errors)
        Assert-Equal @($errors).Count 0 "$([IO.Path]::GetFileName($file)) has no parse errors"
    }
} finally {
    if (Test-Path -LiteralPath $work) { Remove-Item -LiteralPath $work -Recurse -Force }
}

Write-Host ''
Write-Host ("{0} passed, {1} failed" -f $script:passed, $script:failed)
if ($script:failed -gt 0) { exit 1 }
