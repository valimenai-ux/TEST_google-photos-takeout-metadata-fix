# Google Photos Takeout Metadata Fix

Google Takeout exports your photos with their dates, locations and captions in separate
`.json` files. Most photo apps never read those, so photos land on the wrong date or
lose their location.

**`Invoke-GPhotosMetadataFix.ps1`** copies your Takeout folder to a new location and uses
[ExifTool](https://exiftool.org/) to write the metadata from each `.json` sidecar into
the copy. **Your Takeout folder is never modified.**

> Upgrading from the old version? It edited photos in place and had several bugs that
> wrote wrong dates and GPS positions. If you have files ending in `_original`, those are
> your untouched originals. See [Recovering from the old version](#recovering-from-the-old-version)
> and [AUDIT.md](AUDIT.md).

## What it does

* **Copies every file** to the destination with the same folder layout. Photos and videos
  get their metadata added; everything else is copied as is.
* **Never touches the source.** ExifTool always writes a new file (`-o`) and never edits
  in place, so no `*_original` backup files are created.
* **Checks every copy.** Each copy is written under a temporary name, its image/video data
  is compared with the original (SHA-256), and only then is it moved into place. If
  anything goes wrong, the original is copied unchanged and the report says so.
* **Finds the right sidecar** for Google's naming quirks: `IMG_1.jpg.supplemental-metadata.json`,
  names cut short at 51 characters, older `IMG_1.jpg.json`, duplicates like `IMG_1(2).jpg`,
  edited copies (`IMG_1-edited.jpg`, and other languages), and live-photo videos that share
  the photo's sidecar.
* **Writes:**

  | | Photos (JPG, PNG, HEIC, WEBP, TIFF, DNG) | GIF | Videos (MP4, MOV, M4V, 3GP) |
  |---|---|---|---|
  | Date taken | EXIF `DateTimeOriginal`, `CreateDate`, `OffsetTime*`; XMP `CreateDate`, `DateCreated` | XMP | QuickTime `CreateDate` etc. (UTC), `Keys:CreationDate` (local + offset) |
  | Location | EXIF GPS with N/S/E/W refs and altitude | XMP GPS | `Keys` and `UserData` `GPSCoordinates` |
  | Caption | EXIF `ImageDescription`, XMP `Description` | XMP `Description` | XMP `Description`, `Comment` |

  The copy's Created/Modified file dates are also set to the date taken.
* **Keeps good data.** If a photo already has a camera date for the same moment, it's left
  alone, which preserves the local time of photos taken in other time zones. If it differs
  by more than a time-zone shift, you corrected it in Google Photos, so Google's date is
  used. (A correction of a whole number of hours looks the same as a time-zone shift, so
  the camera time is kept; use `-Overwrite` if you shifted times that way.) An existing GPS position is
  kept unless Google's is more than 1 km away. Use `-Overwrite` to always use Google's
  values.
* **Resumes.** Files already in the destination are skipped, so you can stop and restart
  at any time.
* **Reports.** A CSV file in the destination lists every file: what was written, which
  sidecar was used, and any problems.

## Requirements

* Windows PowerShell 5.1 (built into Windows 10/11) or PowerShell 7+.
* [ExifTool](https://exiftool.org/) 12.40 or later. On Windows, download the zip, extract
  it, and **rename `exiftool(-k).exe` to `exiftool.exe`**.

## Usage

1. Extract all your Takeout `.zip` files into one folder, e.g. `D:\Takeout`.
2. Open PowerShell in the folder containing this script. If scripts are blocked, run:
   `Set-ExecutionPolicy -Scope Process Bypass`
3. Preview what will happen (nothing is written):
   ```powershell
   .\Invoke-GPhotosMetadataFix.ps1 -SourcePath 'D:\Takeout\Google Photos' -DestinationPath 'E:\Photos' -ExifToolPath 'C:\Tools\exiftool.exe' -WhatIf
   ```
4. Run it:
   ```powershell
   .\Invoke-GPhotosMetadataFix.ps1 -SourcePath 'D:\Takeout\Google Photos' -DestinationPath 'E:\Photos' -ExifToolPath 'C:\Tools\exiftool.exe'
   ```
5. Open the `gphotos-metadata-report-*.csv` in the destination and review any rows
   marked `CopiedWithoutMetadata` or `Failed`, or noted "no sidecar found".

From Command Prompt:
```cmd
powershell -ExecutionPolicy Bypass -File "C:\path\to\Invoke-GPhotosMetadataFix.ps1" -SourcePath "D:\Takeout\Google Photos" -DestinationPath "E:\Photos" -ExifToolPath "C:\Tools\exiftool.exe"
```

### Parameters

| Parameter | Description |
|---|---|
| `-SourcePath` | Extracted Takeout folder (searched recursively). Only read, never changed. |
| `-DestinationPath` | Where the copies go. Must be outside the source folder. |
| `-ExifToolPath` | Path to `exiftool.exe`. Optional if ExifTool is on the PATH or next to the script. |
| `-TimeZone` | Time zone for the local "date taken", e.g. `"Eastern Standard Time"` (or `"America/New_York"` in PowerShell 7). Default: this PC's time zone. |
| `-Overwrite` | Always write Google's date and GPS, even when the file already has matching values. |
| `-Force` | Redo files that already exist in the destination. |
| `-SkipVerification` | Skip the source/copy image data comparison (faster). |
| `-IncludeSidecars` | Also copy the `.json` files. |
| `-ReportPath` | Custom location for the CSV report. |
| `-WhatIf` | Show what would be done without writing anything. |

### Report statuses

| Status | Meaning |
|---|---|
| `Written` | Copied, metadata added, image data verified identical to the source. |
| `Copied` | Copied unchanged: nothing to add, or not a photo/video ExifTool can write. |
| `CopiedWithoutMetadata` | Metadata couldn't be written (e.g. a PNG saved as `.jpg`); the file was copied unchanged. The reason is in `Notes`. |
| `Skipped` | Already in the destination (use `-Force` to redo). |
| `Failed` | The file couldn't be copied at all. |

## Recovering from the old version

The previous version of this script edited photos in place. If ExifTool left files like
`IMG_1234.jpg_original` next to `IMG_1234.jpg`:

* The `_original` file is the **unmodified original**. Windows just doesn't recognise the
  extension.
* `Restore-ExifToolOriginals.ps1` renames every `*_original` back to its real name and
  moves the edited copy to a folder next to it (e.g. `E:\Photos (exiftool edited copies)`),
  so the edited copies aren't mixed in when you use the restored folder as the source:
  ```powershell
  .\Restore-ExifToolOriginals.ps1 -Path 'E:\Photos' -WhatIf   # preview
  .\Restore-ExifToolOriginals.ps1 -Path 'E:\Photos'
  ```
  (Add `-DeleteModified` to delete the edited copies instead.)
* Then run `Invoke-GPhotosMetadataFix.ps1` from the restored (or freshly extracted)
  Takeout folder into a new destination.

[AUDIT.md](AUDIT.md) explains what went wrong in detail.

## Tests

```powershell
.\tests\Invoke-Tests.ps1 -ExifToolPath 'C:\Tools\exiftool.exe'
```

This builds a fake Takeout folder in your temp directory, runs both scripts against it
and checks the results (70 checks). It's worth running once on your own PC before
processing your real library.

## Credits

The original script and README were written with help from OpenAI and Google language
models. The rewrite, audit and tests were done with Claude Code.
