# Audit: Invoke-GPhotosMetadataFix.ps1

Audit of the script as of commit `716d6b8`, the remediation that was carried out, and
what to do with files damaged by earlier runs.

**Short version:** the old script edited your photos *in place*. It never made a copy.
On top of that it often paired photos with the wrong sidecar, shifted camera times by
your UTC offset, and flipped or erased GPS positions. The rewrite copies to a separate
destination folder, never touches the source, and checks every copy against its
original. The `.jpg_original` files you have are your **untouched originals**. They are
not corrupted, and `Restore-ExifToolOriginals.ps1` puts them back.

---

## 1. The `_original` files and the "corrupted" copies

### What the `_original` files are

When ExifTool edits a file in place, its default behaviour is:

1. write the edited image to a temporary file,
2. rename the untouched original `IMG_1234.jpg` to **`IMG_1234.jpg_original`**,
3. rename the temporary file to `IMG_1234.jpg`.

ExifTool's source (`exiftool`, the file replacement routine) creates that backup **only
when `-overwrite_original` is not given**. Windows doesn't know the `.jpg_original`
extension, so the file won't open and looks corrupted. It is actually the one file that
was never modified. Renaming it back to `.jpg` makes it open normally.

### Why there was no copy on the target machine

The script was never written to copy anything. Every version in the history edits the
files under `$PhotoRoot` directly, and the README even says *"The script is designed to
overwrite the original files"*. The workflow you intended (copy, then pair the metadata
into the copy) did not exist.

### Why `_original` files appeared at all

Every committed version passes `-overwrite_original`. Run as committed under PowerShell
7, it edits in place and leaves **no** backups (reproduced during this audit). So the
backups mean ExifTool did not receive that flag in your run. From the repository alone I
can't tell which of these happened:

* **The latest version can't run on Windows PowerShell 5.1** (the version the README
  targets) on a Western-locale PC. Line 121 contains an em dash (`—`) inside a string.
  PowerShell 5.1 reads a script saved without a BOM in the ANSI code page (Windows-1252),
  where the em dash's bytes `E2 80 94` decode to `â€”`. That `”` is a curly quote, which
  PowerShell treats as a string delimiter. The string ends early, every later quote
  flips, and parsing fails with *"The string is missing the terminator"*. I reproduced
  this by decoding the file as Windows-1252 and parsing it with the PowerShell parser:
  4 parse errors, against 0 when read as UTF-8 and 0 for the first version (`44083d5`).
  So if something ran on 5.1, it was an older or edited copy of the script, or ExifTool
  run some other way.
* The script stores its ExifTool arguments in `$args`, a reserved PowerShell automatic
  variable. It worked when tested under PowerShell 7.4. It is unsupported practice and I
  could not test it under 5.1.

Whatever the trigger, the fix is structural: the new script **never edits in place**. It
calls ExifTool only with `-o <new file>`, never with `-overwrite_original`. ExifTool's
backup step runs only for in-place edits, so `_original` files can't be produced.

### Other ways the edited files were damaged

Besides the backups, the old script wrote *wrong metadata* into many files: findings
C3, C4, H1 and H2 below. In the photo app these show up as photos on the wrong date, at
the wrong place, or with their location missing.

---

## 2. Findings

Each item was reproduced by running the original script with ExifTool 12.76 and
PowerShell 7.4 against a fixture folder built to mimic a real Takeout, unless noted
otherwise.

### Critical

| # | Finding | Evidence |
|---|---------|----------|
| C1 | **Edits originals in place; no copy or destination option.** A failure, a wrong pairing, or a bug permanently changes the only copy. | Design; README "overwrite the original files". 17 of 19 fixture files changed on disk. |
| C2 | **Latest version doesn't parse on Windows PowerShell 5.1** (Western code page), because of the em dash in the string on line 121. | Windows-1252 decode plus parse: 4 errors (`missing the terminator`, 3× `Missing closing '}'`). |
| C3 | **Photos paired with another photo's sidecar.** The wildcard step (`"$cleanBase*.json"`) and the substring step pick unrelated sidecars: `IMG_5.jpg` took `IMG_50.jpg`'s date and GPS; `IMG_7(1).jpg` and `IMG_7(2).jpg` both took `IMG_7.jpg`'s date. `data.jpg` matched `cam.jpg.supplemental-meta`**`data`**`.json`: any name that is a substring of "supplemental-metadata.json" gets a random sidecar. A name that cleans to an empty string (e.g. `(1).jpg`) matches every JSON in the folder. | Fixture run: wrong dates and GPS written to `IMG_5`, `IMG_7(1)`, `IMG_7(2)`, `data.jpg`. |
| C4 | **Camera dates shifted by the UTC offset.** `photoTakenTime.formatted` is UTC; the script strips "UTC" and writes it as local `DateTimeOriginal`. A photo taken 18:30 in New York became 22:30. The "already correct" skip compares the camera's local time with UTC, so it almost never matches and nearly every camera photo is rewritten. | `cam.jpg`: `2019:07:04 18:30:00` → `22:30:00`. `phone.jpg`: `06:00` → `10:00`. |

### High

| # | Finding | Evidence |
|---|---------|----------|
| H1 | **GPS hemisphere lost.** EXIF stores unsigned coordinates plus `GPSLatitudeRef`/`GPSLongitudeRef`. The script writes `-EXIF:GPSLatitude=-33.86` without the refs, so the sign is dropped: Sydney moved to the northern hemisphere, Lima to Asia. | `sydney.jpg`: `33.8688`, no ref. `lima.jpg`: `12.0464, 77.0428`, no refs. |
| H2 | **Existing GPS erased.** If a sidecar has no `geoData`, or its JSON couldn't be read, `$lat` is `$null` and `$null -ne 0` is true, so the script sends `-EXIF:GPSLatitude=` with an empty value. That *deletes* the tag. | `phone.jpg` with camera GPS `40.7,-74`: GPS gone afterwards. |
| H3 | **Dates parsed from a localized display string.** `formatted` follows the Google account's language. `05.03.2024` (German, 5 March) was parsed as **May 3**. Newer exports put a narrow no-break space (U+202F) before AM/PM; that parsed under PowerShell 7 but is untested on 5.1. The numeric `timestamp` field was ignored. | `german.jpg` → `2024:05:03`. |
| H4 | **Square brackets in file names break JSON reading.** `Test-Path`/`Get-Content -Path` treat `[1]` as a wildcard. The JSON read failed, the script still printed "Updated metadata", and it sent the GPS-deleting empty values from H2. | `party[1].jpg`: `Get-Content` error, then "Updated metadata for party[1].jpg". |
| H5 | **Video GPS never written.** EXIF can't be written to MP4/MOV, so `-EXIF:GPSLatitude` on a video silently does nothing. The README claims video GPS support. | `clip.mp4`: ExifTool "0 image files updated"; no GPS. |
| H6 | **ExifTool failures hidden.** Output goes to `Out-Null` and the exit code is never checked, so "Updated metadata for …" prints even when nothing was written. | Seen with H4 and H5. |

### Medium

| # | Finding |
|---|---------|
| M1 | **Text mangled under Windows PowerShell 5.1.** `Get-Content` without `-Encoding` reads the UTF-8 JSON as ANSI (`Café` → `CafÃ©`). 5.1's native argument passing also strips embedded double quotes and can merge arguments ([report](https://forum.cursor.com/t/windows-cli-launcher-agent-cmd-hardcodes-powershell-5-1-args-strips-and-can-glue-the-next-argument/168416), [write-up](https://braindetox.kr/en/posts/powershell_native_command_args_quoting.html)), so captions containing `"` or non-ASCII text arrive damaged. *(Known 5.1 behaviour; not reproducible on Linux.)* |
| M2 | **Filename date fallback is dead code when the sidecar is missing.** The script `return`s as soon as no JSON is found, so the "parse the date from the filename" fallback the README advertises only runs for JSON files that exist but lack dates. |
| M3 | **Assigns to the automatic variable `$args`.** It works in PowerShell 7.4 (tested), but it is unsupported, and PSScriptAnalyzer flags it. |
| M4 | **"Already correct" skip also skips GPS and description.** When the date matches, nothing else is written. |
| M5 | **Truncated sidecar names only partly handled.** Only `.supplemental-me.json` and `.supplementa.json` are listed; Google truncates to 51 characters, which yields many other forms. |
| M6 | **Limited formats.** HEIC, GIF, WEBP, M4V and 3GP are ignored. Live-photo videos (no sidecar of their own) are never paired. |

### Low

| # | Finding |
|---|---------|
| L1 | Up to two ExifTool launches per file, which is slow on Windows. |
| L2 | No validation of `-ExifToolPath`. The unrenamed `exiftool(-k).exe` waits for a key press after every file. |
| L3 | No dry run, no report, no resume after interruption. |

---

## 3. Remediation (implemented)

`Invoke-GPhotosMetadataFix.ps1` was rewritten. Every finding maps to a change below and,
where it can be tested, to a check in `tests/Invoke-Tests.ps1`.

| Finding | Remediation |
|---------|-------------|
| C1 | New mandatory `-DestinationPath`. Every file is copied there with the same folder layout; the source is only read. ExifTool runs with `-o <temp file>` and never with `-overwrite_original`. The temp file is renamed into place only after success and a hash check. A destination inside the source (or the reverse) is refused. *Tests: source tree byte-identical after the run; no `_original` or temp files anywhere.* |
| C2 | All scripts are pure ASCII; non-ASCII strings are built from character codes. *Test: ASCII-only check plus parser check on every script.* |
| C3 | Sidecar matching rebuilt around Google's actual naming rules: `name.ext.supplemental-metadata[(n)].json`, truncation to 51 characters, older `name.ext[(n)].json` and `name.json`, `-edited` (and localized variants) sharing the original's sidecar, duplicate counters `(n)` matched exactly, and live-photo companions. There is no wildcard or substring guessing, and album `metadata.json` is excluded. Lookups go through a per-folder index, so large folders stay fast. *Tests: IMG_5/IMG_50, data.jpg, (1)/(2), edited, three old-style names, truncated name, live photo.* |
| C4 | Dates come from the numeric `photoTakenTime.timestamp`, converted to local time in `-TimeZone` (default: this PC's zone), with `OffsetTimeOriginal` recorded. An existing camera date is **kept** when it describes the same moment (it differs from Google's UTC by a whole time-zone offset), so the local time of travel photos is preserved. If it differs otherwise, the date was corrected in Google Photos and Google's value wins. Video dates are written in UTC as the QuickTime spec requires, plus `Keys:CreationDate` with its offset. *Tests: camera, Paris, corrected, video.* |
| H1 | EXIF GPS is written with `GPSLatitudeRef`/`GPSLongitudeRef`/`GPSAltitudeRef` derived from the sign. *Tests: Sydney, Lima, altitude.* |
| H2 | GPS is written only when the sidecar has a real position (non-null, not 0/0, in range). An existing position is kept unless Google's is more than 1 km away. *Tests: GPS survives a sidecar without geoData; 30 m kept, 30 km replaced.* |
| H3 | The `formatted` string is no longer used at all. *Covered by every date test.* |
| H4 | Paths go through .NET file APIs and `-LiteralPath` only. *Test: `party[1].jpg`.* |
| H5 | Videos get `Keys:GPSCoordinates` and `UserData:GPSCoordinates`. *Test: video GPS.* |
| H6 | ExifTool runs with a per-file status marker (`-echo3 ${status}`); errors and warnings are captured per file. Any failure falls back to a plain copy, marked `CopiedWithoutMetadata` in the CSV report and on screen. *Tests: mis-named PNG still copied and reported; simulated ExifTool crash leaves all files copied unchanged.* |
| M1 | Arguments go to ExifTool as a UTF-8 argument file on stdin (`-@ -`, `-charset filename=utf8`). Captions go through a UTF-8 file (`-TAG<=file`), so quotes, `$`, `@`, line breaks and emoji are exact. JSON is read as UTF-8. *Tests: non-ASCII file name; caption with quotes, `$`, `@`, line break, é, ☕.* |
| M2 | Filename dates apply whether or not a sidecar exists, but never override a date already in the file. *Test: `IMG_20200101_101010.jpg`.* |
| M3 | `$args` is no longer used. |
| M4 | Date, GPS and description are decided independently. |
| M5 | See C3. |
| M6 | Added `.heic .heif .webp .tif .tiff .dng .gif .m4v .3gp .3g2`. Other media (AVI, MKV, RAW…) and non-media files are copied unchanged, so nothing is left behind. |
| L1 | One ExifTool process per 100 files for each step (read, write, verify). About 21 ms per file in testing (1,500 files in 31 s), scaling linearly. |
| L2 | ExifTool is found via `-ExifToolPath`, the PATH, or the script folder; its version is checked, and `(-k)` names are rejected with instructions. |
| L3 | `-WhatIf` preview, CSV report (written even if the run is interrupted), and automatic resume: existing destination files are skipped unless `-Force`. |

Additional safeguards:

* **Copy verification.** After writing, ExifTool's `ImageDataHash` (SHA-256 of the
  image/video data only) is compared between source and copy. If they differ, the
  metadata copy is discarded and the original is copied instead.
* **Motion photos.** The video embedded after the JPEG data is preserved (tested).
* **File dates.** The copy's Created/Modified dates are set to the date taken, so
  Explorer sorts correctly.
* **Recovery helper.** `Restore-ExifToolOriginals.ps1` restores `*_original` backups
  and moves the edited copies aside. It supports `-WhatIf` and needs no other tools.

### Verification

`tests/Invoke-Tests.ps1` builds a fake Takeout folder from four tiny fixtures and runs
both scripts against real ExifTool: **70 checks, all passing** on PowerShell 7.4 with
ExifTool 12.76. To confirm the tests have teeth, three deliberate bugs were planted in
the script one at a time (dropping the GPS ref, never matching camera time, ignoring
duplicate counters). Each one made the suite fail.

**Not verified here:** the scripts were not executed on Windows or under Windows
PowerShell 5.1, because neither is available in this environment. They avoid 5.1
pitfalls by construction: ASCII-only source, no PowerShell 7-only syntax, .NET APIs that
exist in .NET Framework 4.5+, and explicit handling of 5.1's `ConvertFrom-Json` array
behaviour. Run `tests\Invoke-Tests.ps1` once on your PC before the real run.

---

## 4. What to do with files from earlier runs

1. **Don't delete any `*.jpg_original` / `*.mp4_original` files.** They are the
   unmodified originals.
2. If you still have the Takeout `.zip` files, re-extracting them is the cleanest
   source; skip to step 4.
3. Otherwise, restore the originals in place:
   ```powershell
   .\Restore-ExifToolOriginals.ps1 -Path 'E:\Photos' -WhatIf   # preview
   .\Restore-ExifToolOriginals.ps1 -Path 'E:\Photos'           # restore
   ```
   The edited copies go to `E:\Photos (exiftool edited copies)\` (same layout, next to
   the photo folder rather than inside it) instead of being deleted. Files that were edited with `-overwrite_original` (no backup) can't be
   restored this way; their image data is intact, but their date and GPS may be wrong.
   The new script corrects the date when Google's timestamp disagrees, and GPS when it
   is more than 1 km off.
4. Run the new script from the clean source into a **new** folder, preview first:
   ```powershell
   .\Invoke-GPhotosMetadataFix.ps1 -SourcePath 'D:\Takeout\Google Photos' -DestinationPath 'E:\Photos (fixed)' -ExifToolPath 'C:\Tools\exiftool.exe' -WhatIf
   .\Invoke-GPhotosMetadataFix.ps1 -SourcePath 'D:\Takeout\Google Photos' -DestinationPath 'E:\Photos (fixed)' -ExifToolPath 'C:\Tools\exiftool.exe'
   ```
5. Open the CSV report in the destination and check the `CopiedWithoutMetadata`,
   `Failed` and "no sidecar found" rows. Spot-check a few photos, then retire the old
   copies.
