#Requires -Version 5.1
<#
.SYNOPSIS
    Copies a Google Photos Takeout export to a new folder and embeds the metadata
    from Google's JSON sidecar files into the copies. Source files are never modified.

.DESCRIPTION
    Every file under -SourcePath is copied to the same relative location under
    -DestinationPath. For photos and videos that ExifTool can write, the matching
    Takeout sidecar (IMG_1234.jpg.supplemental-metadata.json, its truncated forms,
    or the older IMG_1234.jpg.json) is located, and the date taken, GPS location and
    description it holds are written into the copy.

    Safety guarantees:
      * The source folder is only read. ExifTool is always run with -o (write a new
        file) and never with -overwrite_original, so it never renames anything in
        the source folder and never creates "*_original" backup files.
      * Each copy is written under a temporary name and moved into place only after
        ExifTool reports success and the copy's image/video data hash matches the
        source (unless -SkipVerification is used).
      * If metadata can't be written for any reason, the file is copied unchanged,
        so nothing is left behind. The CSV report lists every file and what
        happened to it.
      * Files already present in the destination are skipped, so an interrupted
        run can simply be started again. Use -Force to redo them.

    What gets written:
      * Date taken: from the sidecar's photoTakenTime (an exact UTC timestamp). If
        the file already has a date that is the same moment, it is kept as is
        (cameras store local time, so keeping it preserves the time of day the
        photo was taken). If it differs, the date was corrected in Google Photos
        and Google's value is written. Files without a sidecar get a date parsed
        from names like IMG_20200101_101010.jpg, but only if they have none.
      * GPS: from geoData (or geoDataExif). An existing position is kept unless it
        is more than 1 km away from Google's, which means the location was edited
        in Google Photos.
      * Description: the caption from Google Photos, when there is one.
      * File created/modified dates of the copy are set to the date taken.
    Use -Overwrite to always write Google's date and GPS values.

.PARAMETER SourcePath
    The extracted Takeout folder, e.g. "D:\Takeout\Google Photos". Searched recursively.

.PARAMETER DestinationPath
    Folder that receives the copies. Must not be inside the source folder (or vice
    versa). Created if it doesn't exist.

.PARAMETER ExifToolPath
    Path to exiftool.exe. Defaults to exiftool on the PATH, then exiftool.exe next
    to this script.

.PARAMETER TimeZone
    Time zone used to turn Google's UTC timestamps into the local "date taken"
    stored in photos, e.g. "Eastern Standard Time" (Windows) or "America/New_York"
    (PowerShell 7). Defaults to this computer's time zone. The UTC offset is stored
    alongside (OffsetTimeOriginal), so the exact moment is always recorded.

.PARAMETER Overwrite
    Always write Google's date and GPS, even when the file already contains a value
    describing the same moment or place.

.PARAMETER Force
    Process files that already exist in the destination instead of skipping them.

.PARAMETER SkipVerification
    Don't compare image/video data hashes of source and copy (faster, less safe).

.PARAMETER IncludeSidecars
    Also copy the .json files to the destination.

.PARAMETER ReportPath
    Where to write the CSV report. Defaults to a timestamped file in the destination.

.EXAMPLE
    .\Invoke-GPhotosMetadataFix.ps1 -SourcePath 'D:\Takeout\Google Photos' -DestinationPath 'E:\Photos' -ExifToolPath 'C:\Tools\exiftool.exe'

.EXAMPLE
    .\Invoke-GPhotosMetadataFix.ps1 'D:\Takeout\Google Photos' 'E:\Photos' -WhatIf

    Shows what would be copied and which metadata would be written, without writing anything.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [Alias('PhotoRoot')]
    [string]$SourcePath,

    [Parameter(Mandatory = $true, Position = 1)]
    [Alias('OutputRoot')]
    [string]$DestinationPath,

    [string]$ExifToolPath,

    [string]$TimeZone,

    [switch]$Overwrite,

    [switch]$Force,

    [switch]$SkipVerification,

    [switch]$IncludeSidecars,

    [string]$ReportPath
)

$ErrorActionPreference = 'Stop'

# Keep this file ASCII-only. Windows PowerShell 5.1 reads scripts without a BOM in
# the ANSI code page, where UTF-8 punctuation can turn into quote characters and
# break parsing. Non-ASCII text below is built from character codes.

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------

# ExifTool writes EXIF + XMP to these.
$ExifKindExtensions = @('.jpg', '.jpeg', '.png', '.heic', '.heif', '.webp', '.tif', '.tiff', '.dng')
# XMP only (GIF has no EXIF).
$XmpKindExtensions = @('.gif')
# QuickTime-based video: dates in UTC, GPS in Keys/UserData.
$QuickTimeKindExtensions = @('.mp4', '.mov', '.m4v', '.3gp', '.3g2')
# Media ExifTool can't safely write: copied unchanged, file dates still set.
$OtherMediaExtensions = @('.avi', '.mkv', '.wmv', '.mpg', '.mpeg', '.mts', '.m2ts', '.mp', '.bmp',
    '.cr2', '.cr3', '.nef', '.arw', '.orf', '.rw2', '.raf')
$AllMediaExtensions = $ExifKindExtensions + $XmpKindExtensions + $QuickTimeKindExtensions + $OtherMediaExtensions

# Suffixes Google adds to edited copies, depending on the account language.
$EditedSuffixes = @(
    '-edited', '-effects', '-smile', '-mix',
    '-bearbeitet', '-bewerkt', '-edytowane', '-modificato', '-editat', '-ha editado',
    ('-modifi' + [char]0x00E9),
    ('-' + [char]0x7DE8 + [char]0x96C6 + [char]0x6E08 + [char]0x307F)
)

# Album-level JSON files that are not per-photo sidecars.
$AlbumMetadataNames = @('metadata.json', 'print-subscriptions.json', 'shared_album_comments.json',
    'user-generated-memory-titles.json')

$SupplementalSuffix = '.supplemental-metadata'
# Google cuts sidecar names to 51 characters (46 + ".json"); any shorter key that
# is a prefix of "<name>.supplemental-metadata" must at least contain the full name.
$TruncatedKeyMinLength = 40
$GpsToleranceKm = 1.0
$BatchSize = 100
$EndMarker = '@@GPHOTOS-FIX-END:'

$Inv = [System.Globalization.CultureInfo]::InvariantCulture
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$UnixEpoch = New-Object System.DateTimeOffset(1970, 1, 1, 0, 0, 0, [TimeSpan]::Zero)
$OnWindows = ($PSVersionTable.PSEdition -ne 'Core') -or [bool](Get-Variable -Name IsWindows -ValueOnly -ErrorAction SilentlyContinue)
if ($OnWindows) {
    $PathComparison = [StringComparison]::OrdinalIgnoreCase
    $PathComparer = [StringComparer]::OrdinalIgnoreCase
} else {
    $PathComparison = [StringComparison]::Ordinal
    $PathComparer = [StringComparer]::Ordinal
}

# ---------------------------------------------------------------------------
# General helpers
# ---------------------------------------------------------------------------

function Get-JsonValue {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function ConvertTo-Double {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [string] -and $Value.Trim() -eq '') { return $null }
    try { return [Convert]::ToDouble($Value, $Inv) } catch { return $null }
}

function Format-Number {
    param([double]$Value)
    return $Value.ToString('0.##########', $Inv)
}

function Add-TrailingSeparator {
    param([string]$Path)
    if ($Path.EndsWith([string][IO.Path]::DirectorySeparatorChar) -or $Path.EndsWith('/')) { return $Path }
    return $Path + [IO.Path]::DirectorySeparatorChar
}

function Get-RelativePath {
    param([string]$Root, [string]$FullPath)
    if (-not $FullPath.StartsWith($Root, $PathComparison)) { throw "'$FullPath' is not inside '$Root'." }
    return $FullPath.Substring($Root.Length)
}

# Key for matching paths echoed back by ExifTool (which may use forward slashes).
function Get-PathKey {
    param([string]$Path)
    return [IO.Path]::GetFullPath($Path)
}

# ---------------------------------------------------------------------------
# Sidecar matching
# ---------------------------------------------------------------------------

function Remove-EditedSuffix {
    param([string]$Stem)
    foreach ($suffix in $EditedSuffixes) {
        if ($Stem.Length -gt $suffix.Length -and $Stem.EndsWith($suffix, [StringComparison]::OrdinalIgnoreCase)) {
            return $Stem.Substring(0, $Stem.Length - $suffix.Length)
        }
    }
    return $Stem
}

# Splits a Takeout media name such as "IMG_1234-edited(2).jpg" into the name Google
# keyed its sidecar on ("IMG_1234.jpg") and the duplicate counter (2).
function Get-MediaIdentity {
    param([string]$FileName)
    $extension = [IO.Path]::GetExtension($FileName)
    $base = Remove-EditedSuffix ([IO.Path]::GetFileNameWithoutExtension($FileName))
    $duplicate = $null
    if ($base -match '^(.+?) ?\((\d+)\)$') {
        $base = Remove-EditedSuffix $Matches[1]
        $duplicate = [int]$Matches[2]
    }
    return [pscustomobject]@{
        FileName       = $FileName
        Extension      = $extension
        Base           = $base
        OriginalName   = $base + $extension
        DuplicateIndex = $duplicate
    }
}

# "IMG_1.HEIC.supplemental-me" -> "IMG_1.HEIC"
function Get-SidecarMediaName {
    param([string]$Key)
    $dot = $Key.LastIndexOf('.')
    if ($dot -gt 0) {
        $tail = $Key.Substring($dot)
        if ($tail.Length -gt 1 -and $SupplementalSuffix.StartsWith($tail, [StringComparison]::OrdinalIgnoreCase)) {
            return $Key.Substring(0, $dot)
        }
    }
    return $Key
}

# Indexes a folder's sidecars by name. A sidecar's key is its file name minus
# ".json"; its base key also drops a trailing "(n)" duplicate counter.
function New-SidecarIndex {
    param($JsonFiles)
    $index = [pscustomobject]@{
        ByKey  = New-Object System.Collections.Hashtable([StringComparer]::OrdinalIgnoreCase)
        ByBase = New-Object System.Collections.Hashtable([StringComparer]::OrdinalIgnoreCase)
        ByStem = New-Object System.Collections.Hashtable([StringComparer]::OrdinalIgnoreCase)
    }
    foreach ($file in $JsonFiles) {
        if ($AlbumMetadataNames -contains $file.Name) { continue }
        $key = $file.Name.Substring(0, $file.Name.Length - $file.Extension.Length)
        $baseKey = $key
        $duplicate = ''
        if ($key -match '^(.+)\((\d+)\)$') {
            $baseKey = $Matches[1]
            $duplicate = [string][int]$Matches[2]
        }
        $index.ByKey[$key] = $file.FullName
        $index.ByBase[$baseKey + '|' + $duplicate] = $file.FullName
        # For live photo companions: "IMG_1.HEIC.supplemental-metadata" -> stem "IMG_1", extension ".HEIC".
        $mediaName = Get-SidecarMediaName $baseKey
        $stemKey = [IO.Path]::GetFileNameWithoutExtension($mediaName) + '|' + $duplicate
        if (-not $index.ByStem.ContainsKey($stemKey)) { $index.ByStem[$stemKey] = New-Object 'System.Collections.Generic.List[object]' }
        $index.ByStem[$stemKey].Add([pscustomobject]@{ Path = $file.FullName; Extension = [IO.Path]::GetExtension($mediaName) })
    }
    return $index
}

# Looks up "<name>.supplemental-metadata", "<name>" (older exports), or either one cut
# short by Google's file name length limit, longest first.
function Find-SidecarKey {
    param($Table, [string]$Name, [string]$Suffix)
    $full = $Name + $SupplementalSuffix
    $shortest = [math]::Max(1, [math]::Min($Name.Length, $TruncatedKeyMinLength))
    for ($length = $full.Length; $length -ge $shortest; $length--) {
        $path = $Table[$full.Substring(0, $length) + $Suffix]
        if ($path) { return $path }
    }
    return $null
}

function Find-Sidecar {
    param($Media, $Index)
    $counter = ''
    if ($null -ne $Media.DuplicateIndex) { $counter = [string]$Media.DuplicateIndex }

    $path = Find-SidecarKey $Index.ByKey $Media.FileName ''
    if ($path) { return [pscustomobject]@{ Path = $path; Method = 'file name' } }
    $path = Find-SidecarKey $Index.ByBase $Media.OriginalName ('|' + $counter)
    if ($path) { return [pscustomobject]@{ Path = $path; Method = 'original name' } }
    $path = $Index.ByBase[$Media.Base + '|' + $counter]
    if ($path) { return [pscustomobject]@{ Path = $path; Method = 'name without extension' } }

    # Live/motion photos: IMG_1.MP4 next to IMG_1.HEIC has no sidecar of its own.
    $candidates = $Index.ByStem[$Media.Base + '|' + $counter]
    if ($candidates) {
        $companions = @($candidates | Where-Object { $_.Extension -ine $Media.Extension -and $AllMediaExtensions -contains $_.Extension })
        if ($companions.Count -eq 1) { return [pscustomobject]@{ Path = $companions[0].Path; Method = 'live photo companion' } }
    }
    return $null
}

function Read-Sidecar {
    param([string]$Path)
    try { return ([IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8) | ConvertFrom-Json) }
    catch { return $null }
}

# ---------------------------------------------------------------------------
# Values from the sidecar
# ---------------------------------------------------------------------------

function Get-SidecarInstant {
    param($Sidecar, [string]$Property)
    $raw = Get-JsonValue (Get-JsonValue $Sidecar $Property) 'timestamp'
    $seconds = 0L
    if ($null -ne $raw -and [long]::TryParse([string]$raw, [Globalization.NumberStyles]::Integer, $Inv, [ref]$seconds) -and $seconds -gt 0) {
        return $UnixEpoch.AddSeconds($seconds)
    }
    return $null
}

# Camera-style names: IMG_20200101_101010, PXL_20230615_123456789, Screenshot_20200101-101010,
# "2020-01-01 10.10.10", and WhatsApp's date-only IMG-20200101-WA0001.
function Get-DateFromFileName {
    param([string]$FileName)
    $patterns = @(
        '(?<!\d)(?<y>(?:19|20)\d\d)(?<mo>\d\d)(?<d>\d\d)[_\-. ]?(?<h>\d\d)(?<mi>\d\d)(?<s>\d\d)',
        '(?<!\d)(?<y>(?:19|20)\d\d)-(?<mo>\d\d)-(?<d>\d\d)(?: at |[ _T\-])(?<h>\d\d)[.:\-](?<mi>\d\d)[.:\-](?<s>\d\d)',
        '(?<!\d)(?<y>(?:19|20)\d\d)(?<mo>\d\d)(?<d>\d\d)-WA\d+'
    )
    foreach ($pattern in $patterns) {
        $m = [regex]::Match($FileName, $pattern)
        if (-not $m.Success) { continue }
        $time = '00:00:00'
        if ($m.Groups['h'].Success) { $time = '{0}:{1}:{2}' -f $m.Groups['h'].Value, $m.Groups['mi'].Value, $m.Groups['s'].Value }
        $text = '{0}-{1}-{2} {3}' -f $m.Groups['y'].Value, $m.Groups['mo'].Value, $m.Groups['d'].Value, $time
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParseExact($text, 'yyyy-MM-dd HH:mm:ss', $Inv, [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
            return $parsed
        }
    }
    return $null
}

function New-DateInfo {
    param([DateTimeOffset]$Instant, [string]$Source, [bool]$ZoneKnown)
    return [pscustomobject]@{
        Utc           = $Instant.ToUniversalTime()
        Local         = [TimeZoneInfo]::ConvertTime($Instant, $Zone)
        Source        = $Source
        ZoneKnown     = $ZoneKnown
        # Only Google's photoTakenTime may replace a date already in the file.
        Authoritative = ($Source -eq 'photoTakenTime')
    }
}

function Resolve-DateTaken {
    param($Sidecar, [string]$FileName)
    $instant = Get-SidecarInstant $Sidecar 'photoTakenTime'
    if ($instant) { return New-DateInfo $instant 'photoTakenTime' $true }
    $wallClock = Get-DateFromFileName $FileName
    if ($wallClock) {
        $instant = New-Object System.DateTimeOffset($wallClock, $Zone.GetUtcOffset($wallClock))
        return New-DateInfo $instant 'file name' $false
    }
    # Upload time: better than nothing, but never replaces an existing date.
    $instant = Get-SidecarInstant $Sidecar 'creationTime'
    if ($instant) { return New-DateInfo $instant 'creationTime' $true }
    return $null
}

function Get-SidecarLocation {
    param($Sidecar)
    foreach ($property in 'geoData', 'geoDataExif') {
        $geo = Get-JsonValue $Sidecar $property
        if ($null -eq $geo) { continue }
        $lat = ConvertTo-Double (Get-JsonValue $geo 'latitude')
        $lon = ConvertTo-Double (Get-JsonValue $geo 'longitude')
        if ($null -eq $lat -or $null -eq $lon) { continue }
        # Google writes 0.0 / 0.0 when there is no location.
        if ($lat -eq 0 -and $lon -eq 0) { continue }
        if ([math]::Abs($lat) -gt 90 -or [math]::Abs($lon) -gt 180) { continue }
        $alt = ConvertTo-Double (Get-JsonValue $geo 'altitude')
        if ($null -eq $alt) { $alt = 0.0 }
        return [pscustomobject]@{ Latitude = $lat; Longitude = $lon; Altitude = $alt }
    }
    return $null
}

function Get-SidecarDescription {
    param($Sidecar)
    $text = [string](Get-JsonValue $Sidecar 'description')
    if ($text.Trim()) { return $text.Trim() }
    return $null
}

# ---------------------------------------------------------------------------
# Values already embedded in the file
# ---------------------------------------------------------------------------

$ReadTags = @(
    '-ExifIFD:DateTimeOriginal', '-XMP-exif:DateTimeOriginal', '-XMP-photoshop:DateCreated', '-XMP-xmp:CreateDate',
    '-QuickTime:CreateDate', '-Keys:CreationDate',
    '-Composite:GPSLatitude', '-Composite:GPSLongitude', '-XMP-exif:GPSLatitude', '-XMP-exif:GPSLongitude',
    '-IFD0:ImageDescription', '-XMP-dc:Description', '-ItemList:Comment', '-ItemList:Description'
)

function ConvertFrom-ExifDate {
    param($Value)
    if ($null -eq $Value) { return $null }
    $text = ([string]$Value).Trim()
    if ($text.Length -lt 19) { return $null }
    $text = $text.Substring(0, 19).Replace('-', ':').Replace('T', ' ')
    # Placeholders for "not set" (zero date, QuickTime epoch, Unix epoch).
    if ($text.StartsWith('0000') -or $text -eq '1904:01:01 00:00:00' -or $text -eq '1970:01:01 00:00:00') { return $null }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParseExact($text, 'yyyy:MM:dd HH:mm:ss', $Inv, [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        return $parsed
    }
    return $null
}

function Get-EmbeddedDate {
    param($Embedded, [string]$Kind)
    switch ($Kind) {
        'Exif' { $names = @('ExifIFD:DateTimeOriginal') }
        'Xmp' { $names = @('XMP-exif:DateTimeOriginal', 'XMP-photoshop:DateCreated', 'XMP-xmp:CreateDate') }
        default { $names = @('QuickTime:CreateDate', 'Keys:CreationDate') }
    }
    foreach ($name in $names) {
        $value = ConvertFrom-ExifDate (Get-JsonValue $Embedded $name)
        if ($value) { return $value }
    }
    return $null
}

function Get-EmbeddedLocation {
    param($Embedded)
    foreach ($group in 'Composite', 'XMP-exif') {
        $lat = ConvertTo-Double (Get-JsonValue $Embedded "${group}:GPSLatitude")
        $lon = ConvertTo-Double (Get-JsonValue $Embedded "${group}:GPSLongitude")
        if ($null -ne $lat -and $null -ne $lon -and -not ($lat -eq 0 -and $lon -eq 0)) {
            return [pscustomobject]@{ Latitude = $lat; Longitude = $lon }
        }
    }
    return $null
}

function Get-EmbeddedDescription {
    param($Embedded, [string]$Kind)
    switch ($Kind) {
        'Exif' { $names = @('XMP-dc:Description', 'IFD0:ImageDescription') }
        'Xmp' { $names = @('XMP-dc:Description') }
        default { $names = @('XMP-dc:Description', 'ItemList:Comment', 'ItemList:Description') }
    }
    foreach ($name in $names) {
        $text = [string](Get-JsonValue $Embedded $name)
        if ($text.Trim()) { return $text.Trim() }
    }
    return $null
}

# A camera stores local time without a zone and Google stores UTC, so the two
# describe the same moment if they differ by a whole time zone offset
# (a multiple of 15 minutes, at most 14 hours).
function Test-SameInstant {
    param([datetime]$WallClock, [DateTimeOffset]$Instant)
    $difference = ($WallClock - $Instant.UtcDateTime).TotalSeconds
    $zoneOffset = [math]::Round($difference / 900) * 900
    return ([math]::Abs($zoneOffset) -le 14 * 3600) -and ([math]::Abs($difference - $zoneOffset) -le 60)
}

function Get-DistanceKm {
    param([double]$Lat1, [double]$Lon1, [double]$Lat2, [double]$Lon2)
    $radians = [math]::PI / 180
    $x = ($Lon2 - $Lon1) * $radians * [math]::Cos(($Lat1 + $Lat2) / 2 * $radians)
    $y = ($Lat2 - $Lat1) * $radians
    return [math]::Sqrt($x * $x + $y * $y) * 6371
}

# ---------------------------------------------------------------------------
# ExifTool arguments
# ---------------------------------------------------------------------------

function Add-DateTags {
    param($Tags, [string]$Kind, $Date)
    $wallClock = $Date.Local.ToString('yyyy:MM:dd HH:mm:ss', $Inv)
    $offset = $Date.Local.ToString('zzz', $Inv)
    $xmpDate = $wallClock
    if ($Date.ZoneKnown) { $xmpDate = $wallClock + $offset }
    switch ($Kind) {
        'Exif' {
            $Tags.Add("-ExifIFD:DateTimeOriginal=$wallClock")
            $Tags.Add("-ExifIFD:CreateDate=$wallClock")
            if ($Date.ZoneKnown) {
                $Tags.Add("-ExifIFD:OffsetTimeOriginal=$offset")
                $Tags.Add("-ExifIFD:OffsetTimeDigitized=$offset")
            } else {
                $Tags.Add('-ExifIFD:OffsetTimeOriginal=')
                $Tags.Add('-ExifIFD:OffsetTimeDigitized=')
            }
            $Tags.Add("-XMP-xmp:CreateDate=$xmpDate")
            $Tags.Add("-XMP-photoshop:DateCreated=$xmpDate")
        }
        'Xmp' {
            $Tags.Add("-XMP-exif:DateTimeOriginal=$xmpDate")
            $Tags.Add("-XMP-xmp:CreateDate=$xmpDate")
            $Tags.Add("-XMP-photoshop:DateCreated=$xmpDate")
        }
        'QuickTime' {
            # The QuickTime spec stores these in UTC.
            $utc = $Date.Utc.UtcDateTime.ToString('yyyy:MM:dd HH:mm:ss', $Inv)
            foreach ($tag in 'CreateDate', 'ModifyDate', 'TrackCreateDate', 'TrackModifyDate', 'MediaCreateDate', 'MediaModifyDate') {
                $Tags.Add("-QuickTime:$tag=$utc")
            }
            if ($Date.ZoneKnown) { $Tags.Add("-Keys:CreationDate=$wallClock$offset") }
        }
    }
}

function Add-GpsTags {
    param($Tags, [string]$Kind, $Location)
    $lat = Format-Number $Location.Latitude
    $lon = Format-Number $Location.Longitude
    $alt = Format-Number $Location.Altitude
    switch ($Kind) {
        'Exif' {
            # EXIF stores unsigned values; the hemisphere goes in the Ref tags, which
            # ExifTool derives from the sign of the number.
            $Tags.Add("-GPS:GPSLatitude=$lat")
            $Tags.Add("-GPS:GPSLatitudeRef=$lat")
            $Tags.Add("-GPS:GPSLongitude=$lon")
            $Tags.Add("-GPS:GPSLongitudeRef=$lon")
            if ($Location.Altitude -ne 0) {
                $Tags.Add("-GPS:GPSAltitude=$(Format-Number ([math]::Abs($Location.Altitude)))")
                $Tags.Add("-GPS:GPSAltitudeRef=$alt")
            } else {
                $Tags.Add('-GPS:GPSAltitude=')
                $Tags.Add('-GPS:GPSAltitudeRef=')
            }
        }
        'Xmp' {
            $Tags.Add("-XMP-exif:GPSLatitude=$lat")
            $Tags.Add("-XMP-exif:GPSLongitude=$lon")
        }
        'QuickTime' {
            $Tags.Add("-Keys:GPSCoordinates=$lat, $lon, $alt")
            $Tags.Add("-UserData:GPSCoordinates=$lat, $lon, $alt")
        }
    }
}

# The caption is passed through a UTF-8 file (-TAG<=FILE) so that line breaks,
# quotes and non-ASCII characters arrive exactly as written.
function Add-DescriptionTags {
    param($Tags, [string]$Kind, [string]$File)
    switch ($Kind) {
        'Exif' { $Tags.Add("-IFD0:ImageDescription<=$File"); $Tags.Add("-XMP-dc:Description<=$File") }
        'Xmp' { $Tags.Add("-XMP-dc:Description<=$File") }
        'QuickTime' { $Tags.Add("-XMP-dc:Description<=$File"); $Tags.Add("-ItemList:Comment<=$File") }
    }
}

function Set-WritePlan {
    param($Item, $Embedded)
    $tags = New-Object 'System.Collections.Generic.List[string]'
    $changes = New-Object 'System.Collections.Generic.List[string]'

    if ($Item.Date) {
        $current = Get-EmbeddedDate $Embedded $Item.Kind
        $action = $null
        if ($null -eq $current) {
            $action = 'added'
        } elseif ($Item.Date.Authoritative -and ($Overwrite -or -not (Test-SameInstant $current $Item.Date.Utc))) {
            $action = 'replaced'
        }
        if ($action) {
            Add-DateTags $tags $Item.Kind $Item.Date
            $changes.Add("date $action")
        } else {
            $Item.Notes.Add('kept embedded date')
        }
    }

    if ($Item.Location) {
        $current = Get-EmbeddedLocation $Embedded
        $action = $null
        if ($null -eq $current) {
            $action = 'added'
        } elseif ($Overwrite -or (Get-DistanceKm $current.Latitude $current.Longitude $Item.Location.Latitude $Item.Location.Longitude) -gt $GpsToleranceKm) {
            $action = 'replaced'
        }
        if ($action) {
            Add-GpsTags $tags $Item.Kind $Item.Location
            $changes.Add("GPS $action")
        } else {
            $Item.Notes.Add('kept embedded GPS')
        }
    }

    if ($Item.Description) {
        $current = Get-EmbeddedDescription $Embedded $Item.Kind
        $normalizedCurrent = ''
        if ($current) { $normalizedCurrent = $current.Replace("`r`n", "`n") }
        if ($normalizedCurrent -cne $Item.Description.Replace("`r`n", "`n")) {
            $Item.DescriptionFile = Join-Path $WorkDir ('caption-{0}.txt' -f [guid]::NewGuid().ToString('N'))
            Add-DescriptionTags $tags $Item.Kind $Item.DescriptionFile
            if ($current) { $changes.Add('description replaced') } else { $changes.Add('description added') }
        }
    }

    $Item.Tags = $tags.ToArray()
    $Item.Changes = $changes.ToArray()
}

# ---------------------------------------------------------------------------
# Running ExifTool
# ---------------------------------------------------------------------------

# Arguments go through stdin as a UTF-8 argument file (-@ -), one per line. That
# sidesteps Windows command line quoting and code page problems entirely.
function Invoke-ExifTool {
    param([string[]]$Arguments)
    foreach ($argument in $Arguments) {
        if ($argument.Contains("`n") -or $argument.Contains("`r")) { throw "Argument contains a line break: $argument" }
    }
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $ExifToolExe
    $startInfo.Arguments = '-@ - -common_args -charset filename=utf8 -api LargeFileSupport=1'
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardInput = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.StandardOutputEncoding = $Utf8NoBom
    $startInfo.StandardErrorEncoding = $Utf8NoBom
    $process = [System.Diagnostics.Process]::Start($startInfo)
    try {
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        $writer = New-Object System.IO.StreamWriter($process.StandardInput.BaseStream, $Utf8NoBom)
        $writer.NewLine = "`n"
        foreach ($argument in $Arguments) { $writer.WriteLine($argument) }
        $writer.Close()
        $process.WaitForExit()
        return [pscustomobject]@{ ExitCode = $process.ExitCode; Output = $stdout.Result; Errors = $stderr.Result }
    } finally {
        $process.Dispose()
    }
}

function Read-ExifToolJson {
    param([string[]]$Arguments)
    $result = New-Object System.Collections.Hashtable($PathComparer)
    $run = Invoke-ExifTool $Arguments
    if ($run.Output.Trim()) {
        # Assign first: Windows PowerShell 5.1 emits a JSON array as one object.
        $parsed = $run.Output | ConvertFrom-Json
        foreach ($entry in $parsed) {
            $source = Get-JsonValue $entry 'SourceFile'
            if (-not $source) { continue }
            $result[(Get-PathKey $source)] = $entry
            # Fallback key, in case ExifTool echoes the path in another form.
            $result['name|' + [IO.Path]::GetFileName($source)] = $entry
        }
    }
    return $result
}

function Get-ResultFor {
    param($Table, [string]$Path)
    $entry = $Table[(Get-PathKey $Path)]
    if ($null -eq $entry) { $entry = $Table['name|' + [IO.Path]::GetFileName($Path)] }
    return $entry
}

function Read-EmbeddedMetadata {
    param([string[]]$Paths)
    $all = New-Object System.Collections.Hashtable($PathComparer)
    for ($i = 0; $i -lt $Paths.Count; $i += $BatchSize) {
        $chunk = $Paths[$i..([math]::Min($i + $BatchSize, $Paths.Count) - 1)]
        $found = Read-ExifToolJson (@('-json', '-n', '-G1', '-q', '-q') + $ReadTags + $chunk)
        foreach ($key in $found.Keys) { $all[$key] = $found[$key] }
    }
    return $all
}

function Get-ImageDataHashes {
    param([string[]]$Paths)
    $hashes = New-Object System.Collections.Hashtable($PathComparer)
    for ($i = 0; $i -lt $Paths.Count; $i += $BatchSize) {
        $chunk = $Paths[$i..([math]::Min($i + $BatchSize, $Paths.Count) - 1)]
        $found = Read-ExifToolJson (@('-json', '-q', '-q', '-api', 'ImageHashType=SHA256', '-ImageDataHash') + $chunk)
        foreach ($key in $found.Keys) { $hashes[$key] = [string](Get-JsonValue $found[$key] 'ImageDataHash') }
    }
    return $hashes
}

# Splits ExifTool output on the markers echoed after each command.
function Split-ExifToolOutput {
    param([string]$Text)
    $segments = @{}
    $buffer = New-Object System.Text.StringBuilder
    foreach ($line in ($Text -split "`r?`n")) {
        $m = [regex]::Match($line, '^' + [regex]::Escape($EndMarker) + '(\d+):([^@]*)@@\s*$')
        if ($m.Success) {
            $segments[[int]$m.Groups[1].Value] = [pscustomobject]@{ Text = $buffer.ToString(); Status = $m.Groups[2].Value }
            [void]$buffer.Clear()
        } else {
            [void]$buffer.AppendLine($line)
        }
    }
    return $segments
}

# Writes each item's tags into a new file at its TempPath, many files per ExifTool run.
function Invoke-WriteBatch {
    param($Items)
    $arguments = New-Object 'System.Collections.Generic.List[string]'
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $item = $Items[$i]
        if ($i -gt 0) { $arguments.Add('-execute') }
        $arguments.Add('-m')
        $arguments.AddRange([string[]]$item.Tags)
        $arguments.AddRange([string[]]@('-o', $item.TempPath, $item.Source))
        $arguments.AddRange([string[]]@('-echo3', ($EndMarker + $i + ':${status}@@'), '-echo4', ($EndMarker + $i + ':${status}@@')))
    }
    $run = Invoke-ExifTool $arguments.ToArray()
    $output = Split-ExifToolOutput $run.Output
    $errors = Split-ExifToolOutput $run.Errors
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $item = $Items[$i]
        $errorText = ''
        if ($errors.ContainsKey($i)) {
            $lines = @($errors[$i].Text -split "`r?`n" | Where-Object { $_.Trim() })
            $errorText = (@($lines | Where-Object { $_ -match '^Error' }) -join ' ').Trim()
            foreach ($warning in @($lines | Where-Object { $_ -match '^Warning' })) { $item.Notes.Add($warning.Trim()) }
        }
        if (-not $output.ContainsKey($i)) {
            $item.WriteError = 'ExifTool stopped before processing this file. ' + $errorText
        } elseif ($output[$i].Status -match '^\d+$' -and $output[$i].Status -ne '0') {
            $item.WriteError = "ExifTool failed (status $($output[$i].Status)). $errorText"
        } elseif ($errorText -or $output[$i].Text -notmatch '\b1 image files (created|copied)') {
            $item.WriteError = 'ExifTool did not create the copy. ' + $errorText
        } elseif (-not [IO.File]::Exists($item.TempPath) -or (New-Object IO.FileInfo($item.TempPath)).Length -eq 0) {
            $item.WriteError = 'ExifTool reported success but the copy is missing or empty.'
        } else {
            $item.WriteOk = $true
        }
        $item.WriteError = $item.WriteError.Trim()
    }
}

# ---------------------------------------------------------------------------
# File operations
# ---------------------------------------------------------------------------

function Remove-FileIfPresent {
    param([string]$Path)
    if ($Path -and [IO.File]::Exists($Path)) { [IO.File]::Delete($Path) }
}

function Publish-TempFile {
    param([string]$TempPath, [string]$Destination)
    Remove-FileIfPresent $Destination
    [IO.File]::Move($TempPath, $Destination)
}

function Copy-Verbatim {
    param($Item)
    Remove-FileIfPresent $Item.TempPath
    [IO.File]::Copy($Item.Source, $Item.TempPath)
    Publish-TempFile $Item.TempPath $Item.Destination
}

function Set-FileDates {
    param([string]$Path, $Date)
    if ($null -eq $Date) { return }
    $utc = $Date.Utc.UtcDateTime
    if ($OnWindows) { [IO.File]::SetCreationTimeUtc($Path, $utc) }
    [IO.File]::SetLastWriteTimeUtc($Path, $utc)
}

function New-WorkItem {
    param([IO.FileInfo]$File, $SidecarIndex, [switch]$CopyOnly)
    $relative = Get-RelativePath $SourceRoot $File.FullName
    $destination = $DestinationRoot + $relative
    $extension = $File.Extension.ToLowerInvariant()
    $kind = $null
    if (-not $CopyOnly) {
        if ($ExifKindExtensions -contains $extension) { $kind = 'Exif' }
        elseif ($XmpKindExtensions -contains $extension) { $kind = 'Xmp' }
        elseif ($QuickTimeKindExtensions -contains $extension) { $kind = 'QuickTime' }
    }
    $tempName = [IO.Path]::GetFileNameWithoutExtension($File.Name) + '.gphotos-partial' + $File.Extension
    $item = [pscustomobject]@{
        Source          = $File.FullName
        Relative        = $relative
        Destination     = $destination
        TempPath        = Join-Path ([IO.Path]::GetDirectoryName($destination)) $tempName
        Kind            = $kind
        IsMedia         = (-not $CopyOnly) -and ($AllMediaExtensions -contains $extension)
        Sidecar         = ''
        SidecarMatch    = ''
        Date            = $null
        Location        = $null
        Description     = $null
        DescriptionFile = $null
        Tags            = @()
        Changes         = @()
        Notes           = New-Object 'System.Collections.Generic.List[string]'
        WriteOk         = $false
        WriteError      = ''
        Verified        = ''
        Status          = ''
    }
    if (-not $item.IsMedia) { return $item }

    $sidecar = $null
    $match = Find-Sidecar (Get-MediaIdentity $File.Name) $SidecarIndex
    if ($match) {
        $item.Sidecar = $match.Path
        $item.SidecarMatch = $match.Method
        $sidecar = Read-Sidecar $match.Path
        if ($null -eq $sidecar) { $item.Notes.Add('sidecar could not be read') }
    } else {
        $item.Notes.Add('no sidecar found')
    }
    $item.Date = Resolve-DateTaken $sidecar $File.Name
    if ($sidecar) {
        $item.Location = Get-SidecarLocation $sidecar
        $item.Description = Get-SidecarDescription $sidecar
    }
    return $item
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

$SourceRoot = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($SourcePath)
if (-not [IO.Directory]::Exists($SourceRoot)) { throw "Source folder not found: $SourceRoot" }
$SourceRoot = Add-TrailingSeparator ([IO.Path]::GetFullPath($SourceRoot))
$DestinationRoot = Add-TrailingSeparator ([IO.Path]::GetFullPath($PSCmdlet.GetUnresolvedProviderPathFromPSPath($DestinationPath)))
if ($DestinationRoot.StartsWith($SourceRoot, $PathComparison) -or $SourceRoot.StartsWith($DestinationRoot, $PathComparison)) {
    throw "The destination folder must not be inside the source folder or contain it.`n  Source:      $SourceRoot`n  Destination: $DestinationRoot"
}

if ($TimeZone) {
    try { $Zone = [TimeZoneInfo]::FindSystemTimeZoneById($TimeZone) }
    catch { throw "Unknown time zone '$TimeZone'. List valid names with: [TimeZoneInfo]::GetSystemTimeZones() | Select-Object Id" }
} else {
    $Zone = [TimeZoneInfo]::Local
}

# Locate ExifTool.
$ExifToolExe = $null
if ($ExifToolPath) {
    if (Test-Path -LiteralPath $ExifToolPath -PathType Leaf) {
        $ExifToolExe = (Resolve-Path -LiteralPath $ExifToolPath).ProviderPath
    } else {
        $command = Get-Command -Name $ExifToolPath -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($command) { $ExifToolExe = $command.Path }
    }
} else {
    $command = Get-Command -Name 'exiftool' -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($command) {
        $ExifToolExe = $command.Path
    } else {
        foreach ($name in 'exiftool.exe', 'exiftool') {
            $candidate = Join-Path $PSScriptRoot $name
            if (Test-Path -LiteralPath $candidate -PathType Leaf) { $ExifToolExe = $candidate; break }
        }
    }
}
if (-not $ExifToolExe) {
    throw "ExifTool not found. Download it from https://exiftool.org/ and pass its location with -ExifToolPath 'C:\path\to\exiftool.exe'."
}
if ([IO.Path]::GetFileName($ExifToolExe) -like '*(-k)*') {
    throw "Rename '$([IO.Path]::GetFileName($ExifToolExe))' to 'exiftool.exe' first. The '(-k)' in the name makes ExifTool wait for a key press after every run."
}
$versionRun = Invoke-ExifTool @('-ver')
$ExifToolVersion = 0.0
if (-not [double]::TryParse($versionRun.Output.Trim(), [Globalization.NumberStyles]::Float, $Inv, [ref]$ExifToolVersion)) {
    throw "'$ExifToolExe' did not report an ExifTool version. Output: $($versionRun.Output) $($versionRun.Errors)"
}
if ($ExifToolVersion -lt 12.40) {
    Write-Warning "ExifTool $ExifToolVersion is old; version 12.40 or later is recommended (needed for copy verification)."
}

Write-Host "Source:      $SourceRoot"
Write-Host "Destination: $DestinationRoot"
Write-Host "ExifTool:    $ExifToolExe ($ExifToolVersion)"
Write-Host "Time zone:   $($Zone.Id)"
if ($WhatIfPreference) { Write-Host 'WhatIf:      nothing will be written' }
Write-Host ''

$enumerationErrors = $null
$allFiles = @(Get-ChildItem -LiteralPath $SourceRoot -Recurse -File -ErrorAction SilentlyContinue -ErrorVariable enumerationErrors)
foreach ($enumerationError in @($enumerationErrors)) {
    Write-Warning "Could not read $($enumerationError.TargetObject): $($enumerationError.Exception.Message)"
}
$directories = @($allFiles | Group-Object -Property DirectoryName | Sort-Object -Property Name)

$WorkDir = Join-Path ([IO.Path]::GetTempPath()) ('gphotos-fix-' + [guid]::NewGuid().ToString('N'))
$report = New-Object 'System.Collections.Generic.List[object]'
$startTime = Get-Date
$processed = 0

try {
    if (-not $WhatIfPreference) {
        [void][IO.Directory]::CreateDirectory($WorkDir)
        [void][IO.Directory]::CreateDirectory($DestinationRoot)
    }

    foreach ($directory in $directories) {
        $percent = 0
        if ($allFiles.Count) { $percent = [int]($processed * 100 / $allFiles.Count) }
        Write-Progress -Activity 'Copying Takeout with metadata' -Status $directory.Name -PercentComplete $percent

        $jsonFiles = @($directory.Group | Where-Object { $_.Extension -ieq '.json' })
        $sidecarIndex = New-SidecarIndex $jsonFiles
        $items = New-Object 'System.Collections.Generic.List[object]'
        foreach ($file in @($directory.Group | Where-Object { $_.Extension -ine '.json' })) {
            $items.Add((New-WorkItem $file $sidecarIndex))
        }
        if ($IncludeSidecars) {
            foreach ($file in $jsonFiles) { $items.Add((New-WorkItem $file $null -CopyOnly)) }
        }
        $processed += $directory.Count

        foreach ($item in $items) {
            if (-not $Force -and [IO.File]::Exists($item.Destination)) { $item.Status = 'Skipped' }
        }

        # Decide what to write, based on what each file already contains.
        $writable = @($items | Where-Object { $_.Kind -and -not $_.Status -and ($_.Date -or $_.Location -or $_.Description) })
        if ($writable.Count) {
            try {
                $embedded = Read-EmbeddedMetadata ([string[]]@($writable | ForEach-Object { $_.Source }))
                foreach ($item in $writable) { Set-WritePlan $item (Get-ResultFor $embedded $item.Source) }
            } catch {
                # Without knowing what the files already contain, copy them unchanged rather than guess.
                Write-Warning "Could not read existing metadata in $($directory.Name): $($_.Exception.Message)"
                foreach ($item in $writable) { $item.Tags = @(); $item.Changes = @(); $item.Notes.Add('existing metadata could not be read') }
            }
        }

        $toWrite = New-Object 'System.Collections.Generic.List[object]'
        foreach ($item in $items) {
            if ($item.Status) { continue }
            $action = 'Copy unchanged'
            if (@($item.Tags).Count) { $action = 'Copy and write ' + ($item.Changes -join ', ') }
            if (-not $PSCmdlet.ShouldProcess($item.Relative, $action)) { $item.Status = 'WhatIf'; continue }
            try {
                [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($item.Destination))
                Remove-FileIfPresent $item.TempPath
                if (@($item.Tags).Count) {
                    if ($item.DescriptionFile) { [IO.File]::WriteAllText($item.DescriptionFile, $item.Description, $Utf8NoBom) }
                    $toWrite.Add($item)
                } else {
                    Copy-Verbatim $item
                    Set-FileDates $item.Destination $item.Date
                    $item.Status = 'Copied'
                }
            } catch {
                $item.Status = 'Failed'
                $item.WriteError = $_.Exception.Message
            }
        }

        for ($i = 0; $i -lt $toWrite.Count; $i += $BatchSize) {
            $batch = $toWrite.GetRange($i, [math]::Min($BatchSize, $toWrite.Count - $i))
            try {
                Invoke-WriteBatch $batch
            } catch {
                foreach ($item in $batch) { $item.WriteOk = $false; $item.WriteError = "ExifTool could not be run: $($_.Exception.Message)" }
            }

            if (-not $SkipVerification) {
                $written = @($batch | Where-Object { $_.WriteOk })
                if ($written.Count) {
                    $paths = New-Object 'System.Collections.Generic.List[string]'
                    foreach ($item in $written) { $paths.Add($item.Source); $paths.Add($item.TempPath) }
                    try {
                        $hashes = Get-ImageDataHashes $paths.ToArray()
                    } catch {
                        $hashes = @{}
                        foreach ($item in $written) { $item.WriteOk = $false; $item.WriteError = "Could not verify the copy: $($_.Exception.Message)" }
                    }
                    foreach ($item in @($written | Where-Object { $_.WriteOk })) {
                        $sourceHash = Get-ResultFor $hashes $item.Source
                        $copyHash = Get-ResultFor $hashes $item.TempPath
                        if (-not $sourceHash) {
                            $item.Verified = 'not supported for this format'
                        } elseif ($sourceHash -eq $copyHash) {
                            $item.Verified = 'image data identical'
                        } else {
                            $item.WriteOk = $false
                            $item.WriteError = "The copy's image data differs from the source; metadata was not applied."
                        }
                    }
                }
            }

            foreach ($item in $batch) {
                try {
                    if ($item.WriteOk) {
                        Publish-TempFile $item.TempPath $item.Destination
                        $item.Status = 'Written'
                    } else {
                        Copy-Verbatim $item
                        $item.Status = 'CopiedWithoutMetadata'
                    }
                    Set-FileDates $item.Destination $item.Date
                } catch {
                    $item.Status = 'Failed'
                    $item.WriteError = (($item.WriteError + ' ' + $_.Exception.Message).Trim())
                } finally {
                    Remove-FileIfPresent $item.TempPath
                    Remove-FileIfPresent $item.DescriptionFile
                }
            }
        }

        foreach ($item in $items) {
            if ($item.Status -eq 'CopiedWithoutMetadata') { Write-Warning "$($item.Relative): copied without metadata. $($item.WriteError)" }
            elseif ($item.Status -eq 'Failed') { Write-Warning "$($item.Relative): FAILED. $($item.WriteError)" }
            else { Write-Verbose "$($item.Relative): $($item.Status) $($item.Changes -join ', ')" }

            $dateText = ''
            $dateSource = ''
            if ($item.Date) {
                $dateText = $item.Date.Local.ToString('yyyy-MM-dd HH:mm:ss', $Inv)
                if ($item.Date.ZoneKnown) { $dateText += ' ' + $item.Date.Local.ToString('zzz', $Inv) }
                $dateSource = $item.Date.Source
            }
            $sidecarText = ''
            if ($item.Sidecar) { $sidecarText = Get-RelativePath $SourceRoot $item.Sidecar }
            $report.Add([pscustomobject]@{
                Status       = $item.Status
                File         = $item.Relative
                Changes      = ($item.Changes -join '; ')
                DateTaken    = $dateText
                DateSource   = $dateSource
                Sidecar      = $sidecarText
                SidecarMatch = $item.SidecarMatch
                Verified     = $item.Verified
                Notes        = ((@($item.Notes) + @($item.WriteError | Where-Object { $_ })) -join '; ')
            })
        }
    }
} finally {
    Write-Progress -Activity 'Copying Takeout with metadata' -Completed
    if ([IO.Directory]::Exists($WorkDir)) { Remove-Item -LiteralPath $WorkDir -Recurse -Force -ErrorAction SilentlyContinue }
    # Written here so that a run stopped with Ctrl+C still records what was done.
    if (-not $WhatIfPreference -and $report.Count) {
        if (-not $ReportPath) {
            $ReportPath = Join-Path $DestinationRoot ('gphotos-metadata-report-{0}.csv' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
        }
        $report | Export-Csv -LiteralPath $ReportPath -NoTypeInformation -Encoding UTF8
    }
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

function Get-StatusCount {
    param([string]$Status)
    return @($report | Where-Object { $_.Status -eq $Status }).Count
}

$elapsed = (Get-Date) - $startTime
Write-Host ''
Write-Host ('Finished {0} files in {1:hh\:mm\:ss}.' -f $report.Count, $elapsed)
if ($WhatIfPreference) {
    Write-Host ('  Would be processed:          {0}' -f (Get-StatusCount 'WhatIf'))
    Write-Host ('  Already in destination:      {0}' -f (Get-StatusCount 'Skipped'))
} else {
    Write-Host ('  Copied with metadata added:  {0}' -f (Get-StatusCount 'Written'))
    Write-Host ('  Copied unchanged:            {0}' -f (Get-StatusCount 'Copied'))
    Write-Host ('  Copied WITHOUT metadata:     {0}' -f (Get-StatusCount 'CopiedWithoutMetadata'))
    Write-Host ('  Skipped (already existed):   {0}' -f (Get-StatusCount 'Skipped'))
    Write-Host ('  Failed:                      {0}' -f (Get-StatusCount 'Failed'))
}
Write-Host ('  Media without a sidecar:     {0}' -f @($report | Where-Object { $_.Notes -like '*no sidecar found*' }).Count)

if (-not $WhatIfPreference -and $report.Count) { Write-Host "Report:      $ReportPath" }
Write-Host 'The source folder was not modified.'

if ((Get-StatusCount 'Failed') -gt 0) { exit 1 }
exit 0
