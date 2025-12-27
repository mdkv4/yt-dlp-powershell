# YouTube Downloader Script

A PowerShell script for downloading YouTube videos with intelligent format selection, automatic metadata handling, and filename cleanup.

## Features

- **Smart Format Selection**: Automatically selects highest quality with codec priority (AV1 > VP9 > H.264)
- **MKV Container**: All downloads saved as `.mkv` files for maximum compatibility
- **Subtitle Support**: Downloads and embeds all available subtitles (manual + auto-generated)
- **Metadata Embedding**: Embeds video metadata including YouTube video ID
- **Intelligent Filename Cleanup**:
  - Removes superfluous phrases ("Official Video", "Lyric Video", etc.)
  - Removes YouTube video ID from filename (stored in metadata instead)
  - Adds release year in format `(YYYY)`
  - Cleans up extra spaces and dashes
- **Firefox Cookie Support**: Accesses age-restricted and member-only content using Firefox cookies
- **Colorized Output**: Clear visual feedback with section headers and color-coded messages
- **Format Preview**: Shows selected codec and bitrate before downloading

## Requirements

- Windows PowerShell 5.1 or later
- yt-dlp executable in `bin\yt-dlp.exe`
- Node.js (for JavaScript runtime)
- Firefox browser (optional, for cookie authentication)
- FFmpeg (for merging video/audio streams and embedding metadata)

## Installation

1. Ensure yt-dlp is located at `bin\yt-dlp.exe`
2. Install Node.js from [nodejs.org](https://nodejs.org)
3. Install FFmpeg (yt-dlp will use it automatically)

## Usage

### Basic Usage

Download with automatic best quality selection:
```powershell
.\download-youtube.ps1 -Url "https://www.youtube.com/watch?v=VIDEO_ID"
```

### List Available Formats

View all available quality options without downloading:
```powershell
.\download-youtube.ps1 -Url "https://www.youtube.com/watch?v=VIDEO_ID" -ListFormats
```

### Specify Format

Download specific format code:
```powershell
.\download-youtube.ps1 -Url "https://www.youtube.com/watch?v=VIDEO_ID" -Format "721+251"
```

### Overwrite Existing Files

Force overwrite if file already exists:
```powershell
.\download-youtube.ps1 -Url "https://www.youtube.com/watch?v=VIDEO_ID" -Overwrite
```

### Firefox Cookie Authentication

Use specific Firefox profile and container:
```powershell
# With profile only
.\download-youtube.ps1 -Url "https://www.youtube.com/watch?v=VIDEO_ID" -Profile "MyProfile"

# With profile and container
.\download-youtube.ps1 -Url "https://www.youtube.com/watch?v=VIDEO_ID" -Profile "MyProfile" -Container "Personal"
```

### Additional Arguments

Pass extra yt-dlp arguments:
```powershell
.\download-youtube.ps1 -Url "https://www.youtube.com/watch?v=VIDEO_ID" -AdditionalArgs "--limit-rate","1M"
```

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `-Url` | String | Yes | YouTube video URL |
| `-Profile` | String | No | Firefox profile name for cookie extraction |
| `-Container` | String | No | Firefox container name (requires Multi-Account Containers extension) |
| `-Format` | String | No | Specific format code (e.g., "721+251"). If omitted, shows interactive prompt |
| `-ListFormats` | Switch | No | List available formats and exit without downloading |
| `-Overwrite` | Switch | No | Overwrite existing files |
| `-AdditionalArgs` | String[] | No | Additional yt-dlp command-line arguments |

## Format Selection Priority

When no format is specified, the script automatically selects based on this priority:

1. **AV1** (av01) - Highest priority for best quality/compression ratio
2. **VP9** (vp9) - Fallback when AV1 unavailable
3. **H.264** (avc1) - Maximum compatibility fallback

The script combines best video with best audio automatically.

## Output

### Filename Format

Downloaded files are automatically renamed to:
```
Artist - Title (YEAR).mkv
```

**Example transformations:**
```
Before: The Pretty Wild - sLeepwALkeR - Official Lyric Video [w7Ioi9eheBU].mkv
After:  The Pretty Wild - sLeepwALkeR (2024).mkv
```

### Embedded Metadata

Each file contains:
- Title, artist, date
- YouTube video ID (in comment field)
- All available subtitles (embedded)
- Chapter markers (if available)

## Filename Cleanup Rules

The script removes these phrases from filenames:

- `Official Music Video`, `Official Video`, `Official Audio`
- `Official Lyric Video`, `Official Lyrics Video`
- `Official Visualizer`
- `Music Video`, `Lyric Video`, `Audio`, `Video`
- `- Topic` (from auto-generated artist channels)
- YouTube video ID in brackets (e.g., `[w7Ioi9eheBU]`)

Supports both `[brackets]` and `(parentheses)` variations.

## Interactive Mode

When run without `-Format` parameter, the script:

1. Shows all available formats
2. Displays an interactive prompt
3. Allows manual format selection or automatic best quality (press Enter)

## Subtitle Handling

- Downloads all available subtitle languages
- Includes both manual and auto-generated captions
- Embeds all subtitles into MKV file
- Can be toggled on/off in media players

## Examples

### Download with best quality (auto-select)
```powershell
.\download-youtube.ps1 -Url "https://www.youtube.com/watch?v=dQw4w9WgXcQ"
```

### Download age-restricted video with Firefox cookies
```powershell
.\download-youtube.ps1 -Url "https://www.youtube.com/watch?v=VIDEO_ID" -Profile "default-release"
```

### Download and overwrite existing file
```powershell
.\download-youtube.ps1 -Url "https://www.youtube.com/watch?v=VIDEO_ID" -Overwrite
```

### Download specific quality (1080p Premium AV1)
```powershell
.\download-youtube.ps1 -Url "https://www.youtube.com/watch?v=VIDEO_ID" -Format "721+251"
```

## Troubleshooting

### "yt-dlp.exe not found"
Ensure yt-dlp is located at `bin\yt-dlp.exe` relative to the script location.

### "No supported JavaScript runtime"
Install Node.js. The script uses `--js-runtimes node` for sites requiring JavaScript execution.

### Cookie extraction fails
- Ensure Firefox is installed and has been opened at least once
- Verify the profile name is correct (check `about:profiles` in Firefox)
- For containers, ensure Firefox Multi-Account Containers extension is installed

### File not renamed
Files must be modified within the last minute to be detected for renaming. The script only processes recently downloaded files.

## Viewing Embedded Metadata

To verify embedded metadata:

### Using ffprobe (part of FFmpeg)
```powershell
ffprobe -show_format "filename.mkv" 2>&1 | Select-String "comment|title|date"
```

### Using MediaInfo
```powershell
mediainfo "filename.mkv"
```

## License

This script is provided as-is for personal use. Respect YouTube's Terms of Service and copyright laws.

## Credits

- Built using [yt-dlp](https://github.com/yt-dlp/yt-dlp)
- Requires [FFmpeg](https://ffmpeg.org/) for video processing
