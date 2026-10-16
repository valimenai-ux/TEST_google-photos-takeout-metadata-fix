#Requires -Version 5.1
<#
.SYNOPSIS
    Puts ExifTool's "*_original" backup files back under their real names.

.DESCRIPTION
    When ExifTool edits a file in place without -overwrite_original, it renames the
    untouched file to "name.ext_original" and saves the edited version as
    "name.ext". Windows doesn't recognise extensions like ".jpg_original", so these
    backups look broken, but they are the original, unmodified files.

    For every "X_original" found under -Path:
      * the edited "X" (if present) is moved to -ModifiedFolder, keeping the same
        relative folder layout, or deleted when -DeleteModified is given;
      * "X_original" is renamed back to "X".

    No other tools are needed. Use -WhatIf to preview.

.PARAMETER Path
    Folder to search recursively.

.PARAMETER ModifiedFolder
    Where the edited copies are moved. Defaults to a folder next to -Path, e.g.
    "E:\Photos (exiftool edited copies)" for -Path "E:\Photos". It is kept outside -Path
    so the edited copies aren't picked up when -Path is later used as a source folder.

.PARAMETER DeleteModified
    Delete the edited copies instead of moving them.

.EXAMPLE
    .\Restore-ExifToolOriginals.ps1 -Path 'E:\Photos' -WhatIf
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,

    [string]$ModifiedFolder,

    [switch]$DeleteModified
)

$ErrorActionPreference = 'Stop'

$root = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path)
if (-not [IO.Directory]::Exists($root)) { throw "Folder not found: $root" }
$rootNoSlash = [IO.Path]::GetFullPath($root).TrimEnd('\', '/')
$root = $rootNoSlash + [IO.Path]::DirectorySeparatorChar
if (-not $ModifiedFolder -and -not $DeleteModified) {
    $parent = [IO.Path]::GetDirectoryName($rootNoSlash)
    if (-not $parent) { throw "-Path is a drive root; pass -ModifiedFolder (a folder elsewhere) or -DeleteModified." }
    $ModifiedFolder = Join-Path $parent ([IO.Path]::GetFileName($rootNoSlash) + ' (exiftool edited copies)')
}
$modifiedRoot = ''
if ($ModifiedFolder) {
    $modifiedRoot = [IO.Path]::GetFullPath($PSCmdlet.GetUnresolvedProviderPathFromPSPath($ModifiedFolder)).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    if ($modifiedRoot.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
        Write-Warning "The edited copies will be moved inside $root. Don't use that folder as a source for Invoke-GPhotosMetadataFix.ps1 without removing them first."
    }
}

$backups = @(Get-ChildItem -LiteralPath $root -Recurse -File -Filter '*_original' |
    Where-Object { $_.Name.EndsWith('_original', [StringComparison]::OrdinalIgnoreCase) -and
        -not ($modifiedRoot -and $_.FullName.StartsWith($modifiedRoot, [StringComparison]::OrdinalIgnoreCase)) })

$restored = 0
$problems = 0
foreach ($backup in $backups) {
    $target = $backup.FullName.Substring(0, $backup.FullName.Length - '_original'.Length)
    try {
        if ([IO.File]::Exists($target)) {
            if ($DeleteModified) {
                if ($PSCmdlet.ShouldProcess($target, 'Delete edited copy')) { [IO.File]::Delete($target) }
            } else {
                $moved = $modifiedRoot + $target.Substring($root.Length)
                if ([IO.File]::Exists($moved)) {
                    Write-Warning "Skipping $($backup.FullName): $moved already exists."
                    $problems++
                    continue
                }
                if ($PSCmdlet.ShouldProcess($target, "Move edited copy to $moved")) {
                    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($moved))
                    [IO.File]::Move($target, $moved)
                }
            }
        }
        if ($PSCmdlet.ShouldProcess($backup.FullName, "Rename to $([IO.Path]::GetFileName($target))")) {
            [IO.File]::Move($backup.FullName, $target)
            $restored++
        }
    } catch {
        Write-Warning "Could not restore $($backup.FullName): $($_.Exception.Message)"
        $problems++
    }
}

Write-Host ("Found {0} backup file(s); restored {1}; problems {2}." -f $backups.Count, $restored, $problems)
if (-not $DeleteModified -and $restored -gt 0) { Write-Host "Edited copies were moved to: $modifiedRoot" }
if ($problems -gt 0) { exit 1 }
