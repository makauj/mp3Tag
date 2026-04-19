# MP3 Tag Metadata Updater

A simple Python script that updates MP3 ID3 tags from file names and exports a formatted Excel processing report.

## What It Does

- Scans a folder of files.
- Processes `.mp3` files.
- Extracts `artist` and `title` from the file name.
- Writes tags using `mutagen` (`EasyID3`).
- Creates `metadata_report.xlsx` showing success/failure for each processed file.

## Requirements

- Python 3.8+
- `mutagen`
- `openpyxl`

Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

Run the script and pass the target folder path:

```bash
python metadata.py "\Path\To\Your\MP3MusicFolder"
```

Example:

```powershell
python metadata.py "C:\\Users\\john\\Music\\Playlist"
```

## Filename Parsing Rules

The script tries separators in this order:

1. `" - "` (recommended)
2. `":"`
3. `"_"`

Examples:

- `Artist Name - Song Title.mp3`
- `Artist Name:Song Title.mp3`
- `Artist Name_Song Title.mp3`

Fallback behavior:

- If no supported separator is found, it uses:

  - `artist = "Unknown Artist"`
  - `title = full filename (without .mp3)`

## Title Cleanup Rules

The script removes specific title noise patterns while keeping unrelated bracketed text.

Removed patterns include:

- `(Official Music Video)`
- `[Official Music Video]`
- `[official video]`
- `(official video)`
- `(Official HD Video)`
- `official video` (plain text)
- `(official song)`
- trailing ID-like tokens such as `[skjfhkjhk]`

Examples:

- `My Song (Official Music Video)` -> `My Song`
- `My Song [official video]` -> `My Song`
- `My Song [Remix]` -> stays `My Song [Remix]`
- `My Song (Live)` -> stays `My Song (Live)`

## Excel Report Output

After processing, the script writes:

- `metadata_report.xlsx` in the same folder you passed in.

Columns:

- `File Name`
- `Full Path`
- `Status` (`Success` or `Failure`)
- `Reason` (failure reason if any)
- `Artist`
- `Title`
- `Separator Used`

Formatting:

- Colored status cells (green = success, red = failure)
- Filter row enabled
- Header frozen
- Auto-sized columns

## Common Issues

### Permission denied

If a file cannot be saved:

- Close apps that may be using the MP3 (media players, preview panes, cloud sync clients).
- Remove read-only attribute from files.
- Ensure you have write permissions for the target folder.

### Not an MP3 file

Non-`.mp3` files are skipped and reported as failures in the report.

## Notes

- Tags written: `artist`, `title`
- Existing ID3 headers are created if missing.
- The script processes only top-level files in the given folder (not recursive subfolders).
