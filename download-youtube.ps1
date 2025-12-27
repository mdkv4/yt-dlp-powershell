[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$Url,

    [Parameter(Mandatory=$false)]
    [string]$Browser = "",

    [Parameter(Mandatory=$false)]
    [string]$BrowserProfile = "",

    [Parameter(Mandatory=$false)]
    [string]$Container = "",

    [Parameter(Mandatory=$false)]
    [string]$Format = "",

    [Parameter(Mandatory=$false)]
    [string]$OutputFormat = "",

    [Parameter(Mandatory=$false)]
    [string]$OutputDirectory = "",

    [Parameter(Mandatory=$false)]
    [switch]$ListFormats,

    [Parameter(Mandatory=$false)]
    [switch]$Overwrite,

    [Parameter(Mandatory=$false)]
    [switch]$NoCleanFilenames,

    [Parameter(Mandatory=$false)]
    [switch]$NoReleaseYear,

    [Parameter(Mandatory=$false)]
    [string[]]$AdditionalArgs = @()
)

#region Dependency Management

<#
.SYNOPSIS
    Ensures required dependencies (yt-dlp and ffmpeg) are available.

.DESCRIPTION
    Checks for the existence of yt-dlp.exe and ffmpeg.exe in the ./bin directory.
    If either is missing, downloads the latest release from their respective GitHub repositories.
    - yt-dlp: https://github.com/yt-dlp/yt-dlp/releases
    - ffmpeg: https://github.com/yt-dlp/FFmpeg-Builds/releases

.EXAMPLE
    Ensure-Dependencies
    Checks for and downloads any missing dependencies.
#>
function Ensure-Dependencies {
    [CmdletBinding()]
    param()

    $binDir = Join-Path $PSScriptRoot "bin"
    $ytDlpPath = Join-Path $binDir "yt-dlp.exe"
    $ffmpegPath = Join-Path $binDir "ffmpeg.exe"

    # Create bin directory if it doesn't exist
    if (-not (Test-Path $binDir)) {
        Write-Host "Creating bin directory..." -ForegroundColor Yellow
        New-Item -ItemType Directory -Path $binDir -Force | Out-Null
    }

    # Check yt-dlp
    if (-not (Test-Path $ytDlpPath)) {
        Write-Host ""
        Write-Host ">> yt-dlp not found, downloading latest release..." -ForegroundColor Yellow
        Write-Host ""

        try {
            $ytDlpUrl = "https://github.com/yt-dlp/yt-dlp/releases/latest/download/yt-dlp.exe"
            Write-Host "Downloading from: $ytDlpUrl" -ForegroundColor Gray

            # Use System.Net.WebClient for reliable download with progress
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFile($ytDlpUrl, $ytDlpPath)

            Write-Host "Downloaded yt-dlp.exe successfully" -ForegroundColor Green
        } catch {
            Write-Host "ERROR: Failed to download yt-dlp.exe" -ForegroundColor Red
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host ""
            Write-Host "Please manually download yt-dlp.exe from:" -ForegroundColor Yellow
            Write-Host "https://github.com/yt-dlp/yt-dlp/releases/latest" -ForegroundColor Cyan
            Write-Host "and place it in: $binDir" -ForegroundColor Gray
            exit 1
        }
    } else {
        Write-Verbose "yt-dlp.exe found at: $ytDlpPath"
    }

    # Check ffmpeg
    if (-not (Test-Path $ffmpegPath)) {
        Write-Host ""
        Write-Host ">> ffmpeg not found, downloading latest release..." -ForegroundColor Yellow
        Write-Host ""

        try {
            # Get the latest release download URL for ffmpeg
            Write-Host "Fetching latest ffmpeg release information..." -ForegroundColor Gray

            $ffmpegReleasesUrl = "https://api.github.com/repos/yt-dlp/FFmpeg-Builds/releases/latest"
            $releaseInfo = Invoke-RestMethod -Uri $ffmpegReleasesUrl -ErrorAction Stop

            # Find the Windows x64 master build (ffmpeg-master-latest-win64-gpl.zip)
            $asset = $releaseInfo.assets | Where-Object { $_.name -match "ffmpeg-master-latest-win64-gpl\.zip" } | Select-Object -First 1

            if (-not $asset) {
                throw "Could not find Windows x64 GPL build in latest release"
            }

            $downloadUrl = $asset.browser_download_url
            Write-Host "Downloading from: $downloadUrl" -ForegroundColor Gray

            # Download to temp location
            $tempZip = Join-Path $env:TEMP "ffmpeg-temp.zip"
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFile($downloadUrl, $tempZip)

            Write-Host "Extracting ffmpeg binaries..." -ForegroundColor Gray

            # Extract all files from the zip to the bin directory
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            $zip = [System.IO.Compression.ZipFile]::OpenRead($tempZip)

            try {
                # Find the bin directory in the archive (usually ffmpeg-master-latest-win64-gpl/bin/)
                $binEntries = $zip.Entries | Where-Object { $_.FullName -match '/bin/.*\.exe$' }

                if ($binEntries.Count -eq 0) {
                    throw "No executable files found in the archive bin directory"
                }

                # Extract all executables from the bin directory
                foreach ($entry in $binEntries) {
                    $destPath = Join-Path $binDir $entry.Name
                    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $destPath, $true)
                    Write-Host "  Extracted: $($entry.Name)" -ForegroundColor Gray
                }

                Write-Host "Extracted ffmpeg binaries successfully" -ForegroundColor Green
            } finally {
                $zip.Dispose()
            }

            # Clean up temp file
            Remove-Item $tempZip -Force -ErrorAction SilentlyContinue

        } catch {
            Write-Host "ERROR: Failed to download ffmpeg.exe" -ForegroundColor Red
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host ""
            Write-Host "Please manually download ffmpeg from:" -ForegroundColor Yellow
            Write-Host "https://github.com/yt-dlp/FFmpeg-Builds/releases/latest" -ForegroundColor Cyan
            Write-Host "Extract ffmpeg.exe and place it in: $binDir" -ForegroundColor Gray
            exit 1
        }
    } else {
        Write-Verbose "ffmpeg.exe found at: $ffmpegPath"
    }

    Write-Host ""
    Write-Host "[OK] All dependencies are available" -ForegroundColor Green
    Write-Host ""
}

# Ensure dependencies are available FIRST before doing anything else
Ensure-Dependencies

#endregion

#region Configuration Loading

function Get-Configuration {
    [CmdletBinding()]
    param()

    Write-Verbose "Loading configuration from config.json"
    $configPath = Join-Path $PSScriptRoot "config.json"
    Write-Verbose "Config path: $configPath"

    # Default configuration
    $defaultConfig = @{
        browser = "firefox"
        profile = ""
        container = ""
        outputDirectory = ""
        defaultFormat = "bv[vcodec*=vp9][height>=1080]+ba/bv[height>=1080]+ba/bv*+ba/b"
        outputFormat = "mkv"
        subtitleLanguages = "en.*"
        embedSubtitles = $true
        embedMetadata = $true
        cleanFilenames = $true
        addReleaseYear = $true
        phrasesToRemove = @(
            '\[Official Music Video\]',
            '\[Official Video\]',
            '\[Official Audio\]',
            '\[Official Lyric Video\]',
            '\[Official Lyrics Video\]',
            '\[Official Visualizer\]',
            '\[Music Video\]',
            '\[Lyric Video\]',
            '\[Lyrics Video\]',
            '\[Audio\]',
            '\[Video\]',
            '\(Official Music Video\)',
            '\(Official Video\)',
            '\(Official Audio\)',
            '\(Official Lyric Video\)',
            '\(Official Lyrics Video\)',
            '\(Official Visualizer\)',
            '\(Music Video\)',
            '\(Lyric Video\)',
            '\(Lyrics Video\)',
            '\(Audio\)',
            '\(Video\)',
            'Official Music Video',
            'Official Video',
            'Official Audio',
            'Official Lyric Video',
            'Official Lyrics Video',
            'Official Visualizer',
            'Music Video',
            'Lyric Video',
            'Lyrics Video',
            '- Topic'
        )
    }

    # Create default config if it doesn't exist
    if (-not (Test-Path $configPath)) {
        Write-Verbose "Config file not found, creating default configuration"
        $defaultConfig | ConvertTo-Json | Set-Content -Path $configPath -Encoding UTF8
        Write-Host "Created default configuration file: $configPath" -ForegroundColor Yellow
        Write-Verbose "Default configuration created successfully"
        return $defaultConfig
    }

    # Load existing config
    Write-Verbose "Reading existing config file"
    try {
        $configJson = Get-Content -Path $configPath -Raw -ErrorAction Stop
        $config = $configJson | ConvertFrom-Json
        Write-Verbose "Config file parsed successfully"

        # Convert to hashtable and ensure all required keys exist
        Write-Verbose "Converting config to hashtable and applying defaults"
        $configHash = @{
            browser = if ($config.browser) { $config.browser } else { $defaultConfig.browser }
            profile = if ($config.PSObject.Properties['profile']) { $config.profile } else { $defaultConfig.profile }
            container = if ($config.PSObject.Properties['container']) { $config.container } else { $defaultConfig.container }
            outputDirectory = if ($config.PSObject.Properties['outputDirectory']) { $config.outputDirectory } else { $defaultConfig.outputDirectory }
            defaultFormat = if ($config.PSObject.Properties['defaultFormat']) { $config.defaultFormat } else { $defaultConfig.defaultFormat }
            outputFormat = if ($config.PSObject.Properties['outputFormat']) { $config.outputFormat } else { $defaultConfig.outputFormat }
            subtitleLanguages = if ($config.PSObject.Properties['subtitleLanguages']) { $config.subtitleLanguages } else { $defaultConfig.subtitleLanguages }
            embedSubtitles = if ($config.PSObject.Properties['embedSubtitles']) { $config.embedSubtitles } else { $defaultConfig.embedSubtitles }
            embedMetadata = if ($config.PSObject.Properties['embedMetadata']) { $config.embedMetadata } else { $defaultConfig.embedMetadata }
            cleanFilenames = if ($config.PSObject.Properties['cleanFilenames']) { $config.cleanFilenames } else { $defaultConfig.cleanFilenames }
            addReleaseYear = if ($config.PSObject.Properties['addReleaseYear']) { $config.addReleaseYear } else { $defaultConfig.addReleaseYear }
            phrasesToRemove = if ($config.PSObject.Properties['phrasesToRemove']) { $config.phrasesToRemove } else { $defaultConfig.phrasesToRemove }
        }

        Write-Verbose "Configuration loaded successfully: browser=$($configHash.browser), outputDir=$($configHash.outputDirectory)"
        return $configHash
    } catch {
        Write-Verbose "Error loading config: $($_.Exception.Message)"
        Write-Host "Warning: Failed to load config.json. Using defaults." -ForegroundColor Yellow
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        return $defaultConfig
    }
}

#endregion

#region Helper Functions

function Get-BrowserString {
    [CmdletBinding()]
    param(
        [string]$Browser,
        [string]$BrowserProfile,
        [string]$Container
    )

    Write-Verbose "Building browser string: Browser=$Browser, Profile=$BrowserProfile, Container=$Container"
    $browserString = $Browser
    if ($BrowserProfile -ne "") {
        $browserString += ":$BrowserProfile"
        Write-Verbose "Added profile to browser string"
    }
    if ($Container -ne "") {
        if ($BrowserProfile -eq "") {
            $browserString += ":"
        }
        $browserString += ":$Container"
        Write-Verbose "Added container to browser string"
    }

    Write-Verbose "Final browser string: $browserString"
    return $browserString
}

function Invoke-ListFormats {
    param(
        [string]$YtDlpPath,
        [string]$BrowserString,
        [string]$Url
    )

    $listArgs = @(
        "--cookies-from-browser", $BrowserString,
        "--js-runtimes", "node",
        "--list-formats",
        $Url
    )

    Write-Host ""
    Write-Host ">> Fetching Available Formats" -ForegroundColor Cyan
    Write-Host ""
    & $YtDlpPath @listArgs
}

function Get-ScoredFormats {
    [CmdletBinding()]
    param(
        [string]$YtDlpPath,
        [string]$BrowserString,
        [string]$Url
    )

    Write-Verbose "Getting scored formats for URL: $Url"
    Write-Verbose "Using browser: $BrowserString"

    $jsonArgs = @(
        "--dump-json",
        "--skip-download",
        "--cookies-from-browser", $BrowserString,
        "--js-runtimes", "node",
        $Url
    )

    Write-Verbose "Executing yt-dlp with JSON dump"
    try {
        $jsonOutput = & $YtDlpPath @jsonArgs 2>&1
        Write-Verbose "yt-dlp executed, parsing JSON output"
        $videoInfo = $jsonOutput | Where-Object { $_ -is [string] -and $_ -match '^\s*\{' } | Out-String | ConvertFrom-Json

        if (-not $videoInfo -or -not $videoInfo.formats) {
            Write-Verbose "No video info or formats found in JSON"
            return $null
        }

        Write-Verbose "Found $($videoInfo.formats.Count) total formats"

        # Filter to include video formats and audio-only formats (exclude storyboards)
        $filteredFormats = $videoInfo.formats | Where-Object {
            # Include if it's a video format (has video codec and height)
            ($_.vcodec -ne 'none' -and $_.vcodec -notmatch 'images' -and $_.height -gt 0) -or
            # Or if it's an audio-only format (no video codec but has audio codec)
            ($_.vcodec -eq 'none' -and $_.acodec -ne 'none' -and $_.acodec -notmatch 'images')
        }

        if (-not $filteredFormats -or $filteredFormats.Count -eq 0) {
            Write-Verbose "No formats found after filtering"
            return $null
        }

        Write-Verbose "Found $($filteredFormats.Count) formats after filtering (video + audio-only)"

        # Calculate scores for each format
        $scoredFormats = @()
        foreach ($format in $filteredFormats) {
            $formatInfo = @{
                format_id = $format.format_id
                format_note = if ($format.format_note) { $format.format_note } else { "" }
                height = if ($format.height) { [int]$format.height } else { 0 }
                vbr = if ($format.vbr) { [int]$format.vbr } else { 0 }
                vcodec = if ($format.vcodec) { $format.vcodec } else { "" }
                fps = if ($format.fps) { [int]$format.fps } else { 0 }
            }

            $score = Get-FormatScore -FormatInfo $formatInfo

            # Determine if this is audio-only or video format
            $isAudioOnly = ($format.vcodec -eq 'none' -and $format.acodec -ne 'none')

            if ($isAudioOnly) {
                Write-Verbose "Format $($format.format_id): audio only at $($format.abr)kbps"
            } else {
                Write-Verbose "Format $($format.format_id): $($formatInfo.height)p $($score.details.codecName) at $($formatInfo.vbr)kbps = Score: $($score.total)"
            }

            # Get audio codec name
            $acodec = if ($format.acodec -and $format.acodec -ne 'none') { $format.acodec } else { "" }
            $acodecName = "N/A"
            if ($acodec -match "opus") { $acodecName = "Opus" }
            elseif ($acodec -match "mp4a.40.5") { $acodecName = "AAC-LC" }
            elseif ($acodec -match "mp4a.40.2") { $acodecName = "AAC" }
            elseif ($acodec -match "mp4a") { $acodecName = "AAC" }
            elseif ($acodec -match "vorbis") { $acodecName = "Vorbis" }
            elseif ($acodec -ne "") { $acodecName = $acodec }

            $abr = if ($format.abr) { [int]$format.abr } else { 0 }

            # Build the format object
            $scoredFormats += [PSCustomObject]@{
                ID = $format.format_id
                Resolution = if ($isAudioOnly) { "audio only" } else { "$($format.width)x$($format.height)" }
                FPS = if ($isAudioOnly) { "-" } else { $format.fps }
                VCodec = if ($isAudioOnly) { "-" } else { $score.details.codecName }
                VBitrate = if ($isAudioOnly) { "-" } else { "$($formatInfo.vbr) kbps" }
                VEffective = if ($isAudioOnly) { "-" } else { "$($score.details.effectiveBitrate) kbps" }
                ACodec = $acodecName
                ABitrate = if ($abr -gt 0) { "$abr kbps" } else { "N/A" }
                Note = $formatInfo.format_note
                Score = if ($isAudioOnly) { 0 } else { $score.total }
            }
        }

        Write-Verbose "Calculated scores for $($scoredFormats.Count) formats"
        # Sort by score (highest first)
        $sorted = $scoredFormats | Sort-Object -Property Score -Descending
        Write-Verbose "Top format: ID=$($sorted[0].ID), Score=$($sorted[0].Score)"
        return $sorted
    } catch {
        Write-Verbose "Error in Get-ScoredFormats: $($_.Exception.Message)"
        return $null
    }
}

function Get-UserFormatSelection {
    [CmdletBinding()]
    param(
        [string]$YtDlpPath,
        [string]$BrowserString,
        [string]$Url,
        [string]$DefaultFormat
    )

    Write-Verbose "Starting user format selection"
    Write-Host ""
    Write-Host ">> Fetching Available Formats" -ForegroundColor Cyan
    Write-Host ""

    # Get scored formats
    Write-Verbose "Calling Get-ScoredFormats"
    $scoredFormats = Get-ScoredFormats -YtDlpPath $YtDlpPath -BrowserString $BrowserString -Url $Url

    if ($scoredFormats) {
        Write-Verbose "Displaying $($scoredFormats.Count) scored formats"

        # Display colorized table header
        Write-Host ""
        Write-Host ("{0,-5} {1,-15} {2,-5} {3,-7} {4,-12} {5,-12} {6,-8} {7,-10} {8,-20} {9,-6}" -f `
            "ID", "Resolution", "FPS", "VCodec", "VBitrate", "VEffective", "ACodec", "ABitrate", "Note", "Score") -ForegroundColor Cyan
        Write-Host ("{0,-5} {1,-15} {2,-5} {3,-7} {4,-12} {5,-12} {6,-8} {7,-10} {8,-20} {9,-6}" -f `
            "-----", "---------------", "-----", "-------", "------------", "------------", "--------", "----------", "--------------------", "------") -ForegroundColor DarkGray

        # Display each format with color coding
        foreach ($format in $scoredFormats) {
            # Determine video codec color
            $vCodecColor = switch ($format.VCodec) {
                "AV1"   { "Magenta" }
                "VP9"   { "Blue" }
                "H.264" { "Yellow" }
                "-"     { "DarkGray" }
                default { "Gray" }
            }

            # Determine audio codec color
            $aCodecColor = switch ($format.ACodec) {
                "Opus"    { "Cyan" }
                "AAC"     { "Green" }
                "AAC-LC"  { "DarkGreen" }
                "Vorbis"  { "DarkCyan" }
                "N/A"     { "DarkGray" }
                default   { "Gray" }
            }

            # Score color (higher is better)
            $scoreColor = if ($format.Score -ge 1000) { "Green" }
                         elseif ($format.Score -ge 500) { "White" }
                         elseif ($format.Score -gt 0) { "Gray" }
                         else { "DarkGray" }

            # Format the row
            Write-Host ("{0,-5} " -f $format.ID) -NoNewline -ForegroundColor White
            Write-Host ("{0,-15} " -f $format.Resolution) -NoNewline -ForegroundColor Gray
            Write-Host ("{0,-5} " -f $format.FPS) -NoNewline -ForegroundColor Gray
            Write-Host ("{0,-7} " -f $format.VCodec) -NoNewline -ForegroundColor $vCodecColor
            Write-Host ("{0,-12} " -f $format.VBitrate) -NoNewline -ForegroundColor Gray
            Write-Host ("{0,-12} " -f $format.VEffective) -NoNewline -ForegroundColor DarkGray
            Write-Host ("{0,-8} " -f $format.ACodec) -NoNewline -ForegroundColor $aCodecColor
            Write-Host ("{0,-10} " -f $format.ABitrate) -NoNewline -ForegroundColor Gray
            Write-Host ("{0,-20} " -f $format.Note) -NoNewline -ForegroundColor DarkCyan
            Write-Host ("{0,-6}" -f $format.Score) -ForegroundColor $scoreColor
        }
        Write-Host ""
    } else {
        # Fallback to standard format listing
        Write-Verbose "Scored formats unavailable, falling back to standard list"
        Write-Host "Warning: Could not parse format information. Falling back to standard display." -ForegroundColor Yellow
        Write-Host ""

        $listArgs = @(
            "--list-formats",
            "--cookies-from-browser", $BrowserString,
            "--js-runtimes", "node",
            $Url
        )
        & $YtDlpPath @listArgs
    }

    Write-Host ""
    Write-Host "----------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "Enter format code " -ForegroundColor Yellow -NoNewline
    Write-Host "(or press " -ForegroundColor Gray -NoNewline
    Write-Host "Enter" -ForegroundColor White -NoNewline
    Write-Host " for best quality): " -ForegroundColor Gray -NoNewline
    $selectedFormat = Read-Host

    if ($selectedFormat -eq "") {
        Write-Verbose "User selected automatic format, using highest scored format"

        # Find the highest scored format (excluding audio-only formats with score 0)
        $bestFormat = $scoredFormats | Where-Object { $_.Score -gt 0 } | Select-Object -First 1

        if ($bestFormat) {
            $selectedFormat = "$($bestFormat.ID)+bestaudio"
            Write-Host "* Using automatic selection: " -ForegroundColor Green -NoNewline
            Write-Host "$($bestFormat.Resolution) $($bestFormat.VCodec)" -ForegroundColor White -NoNewline
            Write-Host " (score: $($bestFormat.Score))" -ForegroundColor DarkGray
        } else {
            # Fallback to default format string if scoring failed
            Write-Verbose "No scored formats available, using default: $DefaultFormat"
            $selectedFormat = $DefaultFormat
            Write-Host "* Using automatic selection (best quality at 1080p or higher)" -ForegroundColor Green
        }
    } else {
        Write-Verbose "User selected format: $selectedFormat"
        # If user specified a video-only format, automatically add best audio
        if ($selectedFormat -match '^\d+$') {
            $selectedFormat = "$selectedFormat+bestaudio"
            Write-Verbose "Added bestaudio to numeric format: $selectedFormat"
            Write-Host "* Using format: " -ForegroundColor Green -NoNewline
            Write-Host "$selectedFormat" -ForegroundColor White
            Write-Host "  (automatically added best audio)" -ForegroundColor DarkGray
        } else {
            Write-Host "* Using format: " -ForegroundColor Green -NoNewline
            Write-Host "$selectedFormat" -ForegroundColor White
        }
    }
    Write-Host ""

    Write-Verbose "Final selected format: $selectedFormat"
    return $selectedFormat
}

<#
.SYNOPSIS
    Calculates a quality score for a video format.

.DESCRIPTION
    Scores a video format based on multiple factors: Premium status, resolution, effective bitrate
    (adjusted for codec efficiency), and frame rate. Returns a score object with detailed breakdown.

    Codec efficiency multipliers (accounts for compression efficiency):
    - AV1: 1.30x (30% more efficient than VP9, can use lower bitrate for same quality)
    - VP9: 1.00x (baseline)
    - H.264: 0.80x (20% less efficient than VP9)

    This means: VP9 @ 1656kbps beats AV1 @ 902kbps because 1656 > (902 * 1.30 = 1173)
    But AV1 @ 1500kbps beats VP9 @ 1500kbps because (1500 * 1.30 = 1950) > 1500

    Scoring breakdown (max 1450 points):
    - Premium quality: 0-300 pts (Premium formats always prioritized)
    - Resolution: 0-600 pts (2160p=600, 1440p=550, 1080p=500, 720p=400)
    - Effective bitrate: 0-600 pts (bitrate × codec multiplier, capped at 3000 effective kbps)
    - Codec bonus: 0-100 pts (AV1=100, VP9=75, H.264=50)
    - Frame rate: 0-50 pts (60fps=50, 30fps=30, 25fps=25)

.PARAMETER FormatInfo
    Hashtable containing format properties (format_note, height, vbr, vcodec, fps, format_id).

.OUTPUTS
    System.Collections.Hashtable
    Score breakdown with total score and individual factor scores.

.EXAMPLE
    Get-FormatScore -FormatInfo @{format_note='1080p Premium'; height=1080; vbr=1656; vcodec='vp9'; fps=25; format_id='356'}
    Returns: @{total=1306; premium=300; resolution=500; bitrate=331; codec=75; fps=25; details='...'}

.EXAMPLE
    Get-FormatScore -FormatInfo @{format_note='1080p Premium'; height=1080; vbr=902; vcodec='av01'; fps=25; format_id='721'}
    Returns: @{total=1160; premium=300; resolution=500; bitrate=235; codec=100; fps=25; details='...'}
#>
function Get-FormatScore {
    [CmdletBinding()]
    param(
        [hashtable]$FormatInfo
    )

    Write-Verbose "Calculating score for format: $($FormatInfo.format_id)"
    $score = @{
        total = 0
        resolution = 0
        bitrate = 0
        codec = 0
        fps = 0
        premium = 0
        details = ""
    }

    # Premium quality bonus (0-300 points) - highest priority
    $formatNote = $FormatInfo.format_note
    if ($formatNote -match "Premium") {
        $score.premium = 300
        $isPremium = $true
    } else {
        $isPremium = $false
    }

    # Resolution score (0-600 points)
    $height = [int]$FormatInfo.height
    if ($height -ge 2160) { $score.resolution = 600 }
    elseif ($height -ge 1440) { $score.resolution = 550 }
    elseif ($height -ge 1080) { $score.resolution = 500 + ($height - 1080) / 4 }
    elseif ($height -ge 720) { $score.resolution = 400 + ($height - 720) / 4 }
    elseif ($height -ge 480) { $score.resolution = 250 + ($height - 480) / 3 }
    else { $score.resolution = $height / 2 }

    # Codec efficiency multiplier and identification
    $vcodec = $FormatInfo.vcodec
    if ($vcodec -match "av01") {
        $codecMultiplier = 1.30  # AV1 is 30% more efficient than VP9
        $codecName = "AV1"
        $codecColor = "Magenta"
        $codecBonus = 100
    }
    elseif ($vcodec -match "vp9") {
        $codecMultiplier = 1.00  # VP9 baseline
        $codecName = "VP9"
        $codecColor = "Blue"
        $codecBonus = 75
    }
    elseif ($vcodec -match "avc1|h264") {
        $codecMultiplier = 0.80  # H.264 is less efficient
        $codecName = "H.264"
        $codecColor = "Yellow"
        $codecBonus = 50
    }
    else {
        $codecMultiplier = 0.60
        $codecName = "Other"
        $codecColor = "Gray"
        $codecBonus = 25
    }

    # Effective bitrate score (0-600 points) - bitrate adjusted by codec efficiency
    $vbr = [int]$FormatInfo.vbr
    if ($vbr -gt 0) {
        $effectiveBitrate = $vbr * $codecMultiplier
        $score.bitrate = [Math]::Min(600, $effectiveBitrate / 5)
    }

    # Codec bonus (0-100 points) - small bonus for newer codecs
    $score.codec = $codecBonus

    # Frame rate bonus (0-50 points)
    $fps = [int]$FormatInfo.fps
    if ($fps -ge 60) { $score.fps = 50 }
    elseif ($fps -ge 50) { $score.fps = 40 }
    elseif ($fps -ge 30) { $score.fps = 30 }
    else { $score.fps = $fps }

    $score.total = [int]($score.premium + $score.resolution + $score.bitrate + $score.codec + $score.fps)
    Write-Verbose "Score breakdown: Premium=$($score.premium), Resolution=$($score.resolution), Bitrate=$($score.bitrate), Codec=$($score.codec), FPS=$($score.fps), Total=$($score.total)"

    # Build detailed explanation
    $score.details = @{
        codecName = $codecName
        codecColor = $codecColor
        codecMultiplier = $codecMultiplier
        height = $height
        vbr = $vbr
        effectiveBitrate = if ($vbr -gt 0) { [int]($vbr * $codecMultiplier) } else { 0 }
        fps = $fps
        isPremium = $isPremium
    }

    return $score
}

function Show-FormatSelection {
    param(
        [string]$YtDlpPath,
        [string]$BrowserString,
        [string]$Format,
        [string]$Url
    )

    Write-Host ""
    Write-Host ">> Format Selection & Quality Analysis" -ForegroundColor Cyan

    # Get detailed format information
    $formatCheckArgs = @(
        "--cookies-from-browser", $BrowserString,
        "--js-runtimes", "node",
        "--format", $Format,
        "--print", "%(format_id)s|%(format_note)s|%(vcodec)s|%(acodec)s|%(height)s|%(width)s|%(vbr)s|%(abr)s|%(fps)s|%(tbr)s",
        "--no-warnings",
        $Url
    )

    $formatOutput = & $YtDlpPath @formatCheckArgs 2>$null
    if ($formatOutput) {
        $parts = $formatOutput -split '\|'

        # Parse format details
        $formatId = $parts[0]
        $formatNote = $parts[1]
        $vcodec = $parts[2]
        $acodec = $parts[3]
        $height = if ($parts[4] -match '^\d+$') { [int]$parts[4] } else { 0 }
        $width = if ($parts[5] -match '^\d+$') { [int]$parts[5] } else { 0 }
        $vbr = if ($parts[6] -match '^[\d.]+$') { [int][double]$parts[6] } else { 0 }
        $abr = if ($parts[7] -match '^[\d.]+$') { [int][double]$parts[7] } else { 0 }
        $fps = if ($parts[8] -match '^[\d.]+$') { [int][double]$parts[8] } else { 0 }
        $tbr = if ($parts[9] -match '^[\d.]+$') { [int][double]$parts[9] } else { 0 }

        # Display selected format
        Write-Host "Selected: " -ForegroundColor Green -NoNewline
        Write-Host "$formatId - $formatNote" -ForegroundColor White
        Write-Host ""

        # Calculate quality score
        $formatInfo = @{
            format_id = $formatId
            format_note = $formatNote
            height = $height
            vbr = $vbr
            vcodec = $vcodec
            fps = $fps
        }
        $score = Get-FormatScore -FormatInfo $formatInfo

        # Display quality factors
        Write-Host "Quality Factors:" -ForegroundColor Cyan

        # Premium status (if applicable)
        if ($score.details.isPremium) {
            Write-Host "  Premium:     " -ForegroundColor Gray -NoNewline
            Write-Host "Yes " -ForegroundColor Green -NoNewline
            Write-Host "($($score.premium) pts)" -ForegroundColor DarkGray
        }

        Write-Host "  Resolution:  " -ForegroundColor Gray -NoNewline
        Write-Host "$($width)x$($height)p " -ForegroundColor White -NoNewline
        Write-Host "($([int]$score.resolution) pts)" -ForegroundColor DarkGray

        Write-Host "  Video Rate:  " -ForegroundColor Gray -NoNewline
        Write-Host "$($vbr) kbps " -ForegroundColor White -NoNewline
        if ($score.details.codecMultiplier -ne 1.0) {
            Write-Host "(~$($score.details.effectiveBitrate) effective) " -ForegroundColor DarkGray -NoNewline
        }
        Write-Host "($([int]$score.bitrate) pts)" -ForegroundColor DarkGray

        Write-Host "  Codec:       " -ForegroundColor Gray -NoNewline
        Write-Host "$($score.details.codecName) " -ForegroundColor $score.details.codecColor -NoNewline
        $efficiencyPct = [int](($score.details.codecMultiplier - 1.0) * 100)
        if ($efficiencyPct -gt 0) {
            Write-Host "(+$efficiencyPct% efficiency) " -ForegroundColor DarkGray -NoNewline
        } elseif ($efficiencyPct -lt 0) {
            Write-Host "($efficiencyPct% efficiency) " -ForegroundColor DarkGray -NoNewline
        }
        Write-Host "($($score.codec) pts)" -ForegroundColor DarkGray

        if ($fps -gt 0) {
            Write-Host "  Frame Rate:  " -ForegroundColor Gray -NoNewline
            Write-Host "$fps fps " -ForegroundColor White -NoNewline
            Write-Host "($($score.fps) pts)" -ForegroundColor DarkGray
        }

        if ($abr -gt 0) {
            Write-Host "  Audio Rate:  " -ForegroundColor Gray -NoNewline
            Write-Host "$abr kbps ($acodec)" -ForegroundColor White
        }

        Write-Host ""
        Write-Host "  Total Score: " -ForegroundColor Cyan -NoNewline
        Write-Host "$($score.total) points" -ForegroundColor White

    } else {
        Write-Host "Format:   " -ForegroundColor Gray -NoNewline
        Write-Host "$Format" -ForegroundColor White
    }
    Write-Host ""
}

<#
.SYNOPSIS
    Extracts video metadata including release year.

.DESCRIPTION
    Retrieves video metadata from yt-dlp without downloading the video.
    Extracts the release year from either the release_date or upload_date field,
    which is later used for filename enhancement.

.PARAMETER YtDlpPath
    Full path to the yt-dlp executable.

.PARAMETER BrowserString
    Browser string for cookie extraction (e.g., "firefox:profile:container").

.PARAMETER Url
    YouTube video URL to query.

.OUTPUTS
    System.String
    The release year as a 4-digit string (e.g., "2024"), or $null if not found.

.EXAMPLE
    $year = Get-VideoMetadata -YtDlpPath "C:\tools\yt-dlp.exe" -BrowserString "firefox" -Url "https://youtube.com/watch?v=abc123"
    Returns: "2024"
#>
function Get-VideoMetadata {
    param(
        [string]$YtDlpPath,
        [string]$BrowserString,
        [string]$Url
    )

    $metadataArgs = @(
        "--cookies-from-browser", $BrowserString,
        "--js-runtimes", "node",
        "--dump-single-json",
        "--skip-download",
        $Url
    )

    Write-Host "Extracting metadata..." -ForegroundColor DarkGray
    
    try {
        $metadataJson = & $YtDlpPath @metadataArgs 2>&1 | Where-Object { $_ -is [string] } | ConvertFrom-Json
        $releaseYear = $null

        if ($metadataJson.release_date) {
            $releaseYear = $metadataJson.release_date.Substring(0, 4)
        } elseif ($metadataJson.upload_date) {
            $releaseYear = $metadataJson.upload_date.Substring(0, 4)
        }

        if ($releaseYear) {
            Write-Host "Release year: $releaseYear" -ForegroundColor DarkGray
        } else {
            Write-Host "Release year not found in metadata" -ForegroundColor DarkGray
        }
    } catch {
        Write-Host "Warning: Failed to extract metadata" -ForegroundColor Yellow
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor DarkGray
        $releaseYear = $null
    }
    
    Write-Host ""
    return $releaseYear
}

<#
.SYNOPSIS
    Generates a cleaned filename by removing unwanted patterns.

.DESCRIPTION
    Pure function that transforms a YouTube video filename by:
    - Removing YouTube video IDs (e.g., [w7Ioi9eheBU])
    - Removing common phrases (Official Video, Music Video, etc.)
    - Cleaning up extra spaces, dashes, and empty brackets
    - Adding release year in parentheses before the extension
    Does not perform any file operations - only string manipulation.

.PARAMETER OriginalName
    The original filename to clean.

.PARAMETER ReleaseYear
    Optional release year to append before the file extension (e.g., "2024").

.PARAMETER PhrasesToRemove
    Array of regex patterns to remove from the filename (e.g., '\[Official Video\]', 'Music Video').

.OUTPUTS
    System.String
    The cleaned filename with YouTube ID and phrases removed, and year added.

.EXAMPLE
    Get-CleanedFilename -OriginalName "Artist - Song [Official Video] [abc12345678].mkv" -ReleaseYear "2024" -PhrasesToRemove @('\[Official Video\]')
    Returns: "Artist - Song (2024).mkv"

.EXAMPLE
    Get-CleanedFilename -OriginalName "Video Title - Topic [w7Ioi9eheBU].mkv" -ReleaseYear $null -PhrasesToRemove @('- Topic')
    Returns: "Video Title.mkv"
#>
function Get-CleanedFilename {
    [CmdletBinding()]
    param(
        [string]$OriginalName,
        [string]$ReleaseYear,
        [string[]]$PhrasesToRemove
    )

    Write-Verbose "Cleaning filename: $OriginalName"
    $newName = $OriginalName

    # Remove YouTube video ID from filename (e.g., [w7Ioi9eheBU])
    $newName = $newName -replace '\[[a-zA-Z0-9_-]{11}\]', ''

    # Remove each phrase from the provided list
    foreach ($phrase in $PhrasesToRemove) {
        $newName = $newName -replace $phrase, ''
    }

    # Sanitize invalid Windows filename characters (must be done before other cleanup)
    # Replace with safe alternatives rather than removing to preserve readability
    $newName = $newName -replace '[<>:"/\\|?*]', '-'  # Invalid Windows chars -> dash
    $newName = $newName -replace '[\x00-\x1F]', ''    # Control characters -> remove

    # Replace problematic Unicode characters that might cause issues
    $newName = $newName -replace '｜', '|'  # Full-width vertical bar -> ASCII pipe
    $newName = $newName -replace '—', '-'  # Em dash -> regular dash
    $newName = $newName -replace '–', '-'  # En dash -> regular dash
    $newName = $newName -replace '[''ʻ]', "'"  # Smart quotes -> straight quote
    $newName = $newName -replace '[""„]', '"'  # Smart double quotes -> straight quote
    $newName = $newName -replace '…', '...'  # Ellipsis -> three dots

    # Clean up extra spaces and dashes
    $newName = $newName -replace '\s+', ' '  # Multiple spaces to single
    $newName = $newName -replace '\s*-\s*-\s*', ' - '  # Clean up double dashes
    $newName = $newName -replace '\s*-\s*\.', '.'  # Remove dash before extension
    $newName = $newName -replace '^\s*-\s*', ''  # Remove leading dash
    $newName = $newName -replace '\s*-\s*$', ''  # Remove trailing dash (before ext)
    $newName = $newName -replace '\[\s*\]', ''  # Empty brackets
    $newName = $newName -replace '\(\s*\)', ''  # Empty parentheses
    $newName = $newName.Trim()

    # Add release year before extension if available
    if ($ReleaseYear) {
        Write-Verbose "Adding release year: $ReleaseYear"
        # Check if year already exists in filename
        if ($newName -notmatch "\($ReleaseYear\)") {
            # Split filename and extension
            $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($newName)
            $extension = [System.IO.Path]::GetExtension($newName)

            # Add year in parentheses before extension
            $newName = "$nameWithoutExt ($ReleaseYear)$extension"
        } else {
            Write-Verbose "Year already exists in filename"
        }
    }

    # Final cleanup: remove any remaining double spaces and trim
    $newName = $newName -replace '\s+', ' '
    $newName = $newName.Trim()

    Write-Verbose "Cleaned filename: $newName"
    return $newName
}

<#
.SYNOPSIS
    Extracts the video ID from a YouTube URL.

.DESCRIPTION
    Parses a YouTube URL to extract the video ID. Supports various YouTube URL formats
    including watch URLs, shortened youtu.be links, and embed URLs.

.PARAMETER Url
    The YouTube URL to parse.

.OUTPUTS
    System.String
    The 11-character video ID, or $null if not found.

.EXAMPLE
    Get-VideoId -Url "https://www.youtube.com/watch?v=dQw4w9WgXcQ"
    Returns: "dQw4w9WgXcQ"

.EXAMPLE
    Get-VideoId -Url "https://youtu.be/dQw4w9WgXcQ"
    Returns: "dQw4w9WgXcQ"
#>
function Get-VideoId {
    param(
        [string]$Url
    )

    # Try to match standard watch URL: youtube.com/watch?v=VIDEOID
    if ($Url -match '[?&]v=([a-zA-Z0-9_-]{11})') {
        return $matches[1]
    }
    
    # Try to match shortened URL: youtu.be/VIDEOID
    if ($Url -match 'youtu\.be/([a-zA-Z0-9_-]{11})') {
        return $matches[1]
    }
    
    # Try to match embed URL: youtube.com/embed/VIDEOID
    if ($Url -match '/embed/([a-zA-Z0-9_-]{11})') {
        return $matches[1]
    }
    
    return $null
}

<#
.SYNOPSIS
    Normalizes a YouTube URL or video ID into a full URL.

.DESCRIPTION
    Accepts either a full YouTube URL or just a video ID and returns a properly
    formatted YouTube URL. If a video ID is provided (11 characters), it constructs
    the standard watch URL. If a URL with playlist/index parameters is provided,
    it strips those parameters and returns a clean video-only URL.

.PARAMETER Input
    Either a YouTube URL or an 11-character video ID.

.OUTPUTS
    System.String
    A full YouTube URL in the format "https://www.youtube.com/watch?v=VIDEOID"

.EXAMPLE
    Get-NormalizedUrl -InputString "dQw4w9WgXcQ"
    Returns: "https://www.youtube.com/watch?v=dQw4w9WgXcQ"

.EXAMPLE
    Get-NormalizedUrl -InputString "https://youtu.be/dQw4w9WgXcQ"
    Returns: "https://youtu.be/dQw4w9WgXcQ"

.EXAMPLE
    Get-NormalizedUrl -InputString "-vyA9cBE_Eo"
    Returns: "https://www.youtube.com/watch?v=-vyA9cBE_Eo"

.EXAMPLE
    Get-NormalizedUrl -InputString "https://www.youtube.com/watch?v=rHBxJCq99jA&list=LL&index=15"
    Returns: "https://www.youtube.com/watch?v=rHBxJCq99jA"
#>
function Get-NormalizedUrl {
    param(
        [string]$InputString
    )

    # Check if input is already a URL (contains http:// or https://)
    if ($InputString -match '^https?://') {
        try {
            # Parse the URL
            $uri = [System.Uri]$InputString

            # Check if it's a YouTube watch URL with query parameters
            if ($uri.Host -match 'youtube\.com' -and $uri.AbsolutePath -eq '/watch' -and $uri.Query) {
                # Load System.Web assembly for HttpUtility
                Add-Type -AssemblyName System.Web

                # Parse query string
                $queryParams = [System.Web.HttpUtility]::ParseQueryString($uri.Query)

                # Extract video ID
                $videoId = $queryParams['v']

                if ($videoId -and $videoId -match '^[a-zA-Z0-9_-]{11}$') {
                    # Return clean URL with only video ID
                    return "https://www.youtube.com/watch?v=$videoId"
                }
            }

            # Return as-is if no modification needed (e.g., youtu.be URLs, embed URLs)
            return $InputString
        } catch {
            # If URL parsing fails, return as-is
            return $InputString
        }
    }

    # Check if input is a valid video ID (11 characters, alphanumeric with - and _)
    if ($InputString -match '^[a-zA-Z0-9_-]{11}$') {
        return "https://www.youtube.com/watch?v=$InputString"
    }

    # If neither URL nor valid video ID, return as-is and let yt-dlp handle the error
    return $InputString
}

<#
.SYNOPSIS
    Removes the temporary download directory.

.DESCRIPTION
    Deletes the temporary directory created by yt-dlp for intermediate files
    during the download and merge process. Suppresses errors if the directory
    doesn't exist or cannot be removed.

.PARAMETER TempDir
    Full path to the temporary directory to remove.

.EXAMPLE
    Remove-TempDirectory -TempDir "C:\Downloads\temp_20241215_143022"
    Removes the specified temporary directory and all its contents.
#>
function Remove-TempDirectory {
    param(
        [string]$TempDir
    )

    if (Test-Path $TempDir) {
        try {
            Remove-Item -Path $TempDir -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "Temporary directory cleaned up." -ForegroundColor DarkGray
        } catch {
            Write-Host "Warning: Could not remove temporary directory: $TempDir" -ForegroundColor Yellow
        }
    }
}

<#
.SYNOPSIS
    Validates if a string is a valid YouTube URL or video ID.

.DESCRIPTION
    Checks if the input is either a properly formatted YouTube URL
    (standard, shortened, or embed) or a valid 11-character video ID.

.PARAMETER Input
    The string to validate (URL or video ID).

.OUTPUTS
    System.Boolean
    $true if the input is valid, $false otherwise.

.EXAMPLE
    Test-ValidYouTubeInput -InputString "dQw4w9WgXcQ"
    Returns: $true

.EXAMPLE
    Test-ValidYouTubeInput -InputString "invalid"
    Returns: $false
#>
function Test-ValidYouTubeInput {
    [CmdletBinding()]
    param(
        [string]$InputString
    )

    Write-Verbose "Validating YouTube input: $InputString"
    # Valid YouTube URL patterns
    $validPatterns = @(
        '^https?://(www\.)?youtube\.com/watch\?.*v=[a-zA-Z0-9_-]{11}',  # Standard URL
        '^https?://(www\.)?youtu\.be/[a-zA-Z0-9_-]{11}',                # Shortened URL
        '^https?://(www\.)?youtube\.com/embed/[a-zA-Z0-9_-]{11}',       # Embed URL
        '^[a-zA-Z0-9_-]{11}$'                                            # Video ID only
    )

    foreach ($pattern in $validPatterns) {
        if ($InputString -match $pattern) {
            Write-Verbose "Input matched pattern: $pattern"
            return $true
        }
    }

    Write-Verbose "Input did not match any valid patterns"
    return $false
}

#endregion

#region Main Script

# Load configuration
$config = Get-Configuration

# Normalize URL/ID input
$normalizedUrl = Get-NormalizedUrl -InputString $Url

# Validate the URL/ID
if (-not (Test-ValidYouTubeInput -InputString $Url)) {
    Write-Host ""
    Write-Host "ERROR: Invalid YouTube URL or video ID" -ForegroundColor Red
    Write-Host "Input: " -ForegroundColor Yellow -NoNewline
    Write-Host "$Url" -ForegroundColor White
    Write-Host ""
    Write-Host "Expected formats:" -ForegroundColor Yellow
    Write-Host "  Full URL:      " -ForegroundColor Gray -NoNewline
    Write-Host "https://youtube.com/watch?v=dQw4w9WgXcQ" -ForegroundColor Cyan
    Write-Host "  Shortened URL: " -ForegroundColor Gray -NoNewline
    Write-Host "https://youtu.be/dQw4w9WgXcQ" -ForegroundColor Cyan
    Write-Host "  Video ID:      " -ForegroundColor Gray -NoNewline
    Write-Host "dQw4w9WgXcQ" -ForegroundColor Cyan -NoNewline
    Write-Host " (exactly 11 characters)" -ForegroundColor DarkGray
    Write-Host ""
    exit 1
}

# Path to yt-dlp executable (already verified by Ensure-Dependencies)
$ytDlpPath = Join-Path $PSScriptRoot "bin\yt-dlp.exe"

# Validate Node.js is available (required for --js-runtimes)
$nodeCommand = Get-Command node -ErrorAction SilentlyContinue
if (-not $nodeCommand) {
    Write-Host "ERROR: Node.js (node.exe) not found in PATH" -ForegroundColor Red
    Write-Host "Node.js is required for YouTube cookie extraction." -ForegroundColor Yellow
    Write-Host "Please install Node.js from: https://nodejs.org/" -ForegroundColor Yellow
    exit 1
}

# Determine profile and container (command-line parameters override config)
$effectiveBrowser = if ($Browser -ne "") { $Browser } else { $config.browser }
$effectiveProfile = if ($BrowserProfile -ne "") { $BrowserProfile } else { $config.profile }
$effectiveContainer = if ($Container -ne "") { $Container } else { $config.container }

# Validate browser name
$validBrowsers = @("firefox", "chrome", "chromium", "edge", "opera", "brave", "vivaldi", "safari")
if ($effectiveBrowser -notin $validBrowsers) {
    Write-Host "Warning: Browser '$effectiveBrowser' may not be supported by yt-dlp" -ForegroundColor Yellow
    Write-Host "Supported browsers: $($validBrowsers -join ', ')" -ForegroundColor Gray
    Write-Host ""
}

# Build the browser string for cookies
$browserString = Get-BrowserString -Browser $effectiveBrowser -BrowserProfile $effectiveProfile -Container $effectiveContainer

# Determine output directory (command-line parameter overrides config)
$effectiveOutputDir = if ($OutputDirectory -ne "") {
    $OutputDirectory
} elseif ($config.outputDirectory -ne "") {
    $config.outputDirectory
} else {
    $PSScriptRoot
}

# Validate and create output directory if needed
if (-not (Test-Path $effectiveOutputDir)) {
    try {
        New-Item -ItemType Directory -Path $effectiveOutputDir -Force -ErrorAction Stop | Out-Null
        Write-Host "Created output directory: $effectiveOutputDir" -ForegroundColor Green
    } catch {
        Write-Host "ERROR: Failed to create output directory: $effectiveOutputDir" -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

# Determine output format (command-line parameter overrides config)
$effectiveOutputFormat = if ($OutputFormat -ne "") { $OutputFormat } else { $config.outputFormat }

# Validate output format
$validFormats = @("mkv", "mp4", "webm", "avi", "mov", "flv")
if ($effectiveOutputFormat -notin $validFormats) {
    Write-Host "Warning: Output format '$effectiveOutputFormat' may not be supported" -ForegroundColor Yellow
    Write-Host "Common formats: $($validFormats -join ', ')" -ForegroundColor Gray
    Write-Host ""
}

# Determine filename cleaning options (switches override config)
$shouldCleanFilenames = if ($NoCleanFilenames) { $false } else { $config.cleanFilenames }
$shouldAddReleaseYear = if ($NoReleaseYear) { $false } else { $config.addReleaseYear }

# If ListFormats is specified, just list formats and exit
if ($ListFormats) {
    Invoke-ListFormats -YtDlpPath $ytDlpPath -BrowserString $browserString -Url $normalizedUrl
    exit 0
}

# If no format specified, list formats and prompt user
if ($Format -eq "") {
    $Format = Get-UserFormatSelection -YtDlpPath $ytDlpPath -BrowserString $browserString -Url $normalizedUrl -DefaultFormat $config.defaultFormat
}

# Create temporary directory for download artifacts (include video ID to prevent collisions)
$videoId = Get-VideoId -Url $normalizedUrl
$tempDirName = if ($videoId) {
    "temp_$(Get-Date -Format 'yyyyMMdd_HHmmss')_$videoId"
} else {
    "temp_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
}
$tempDir = Join-Path $PSScriptRoot $tempDirName
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

# Build arguments array
$arguments = @(
    "--cookies-from-browser", $browserString,
    "--js-runtimes", "node",
    "--format", $Format,
    "--merge-output-format", $effectiveOutputFormat,
    "--remux-video", $effectiveOutputFormat
)

# Add subtitle options if enabled
if ($config.embedSubtitles) {
    $arguments += "--write-subs"
    $arguments += "--write-auto-subs"
    $arguments += "--sub-langs"
    $arguments += $config.subtitleLanguages
    $arguments += "--embed-subs"
}

# Add metadata option if enabled
if ($config.embedMetadata) {
    $arguments += "--embed-metadata"
    $arguments += "--parse-metadata"
    $arguments += "%(id)s:%(meta_comment)s"
}

$arguments += "--paths"
$arguments += "temp:$tempDir"
$arguments += "--paths"
$arguments += "home:$tempDir"

# Add overwrite flag if specified
if ($Overwrite) {
    $arguments += "--force-overwrites"
}

$arguments += $normalizedUrl

# Add any additional arguments
if ($AdditionalArgs.Count -gt 0) {
    $arguments += $AdditionalArgs
}

# Show selected format
Show-FormatSelection -YtDlpPath $ytDlpPath -BrowserString $browserString -Format $Format -Url $normalizedUrl

# Wrap download and processing in try/finally to ensure cleanup
try {
    # Download section
    Write-Host ""
    Write-Host ">> Downloading" -ForegroundColor Cyan
    Write-Host ""

    # Get metadata to extract release year
    $releaseYear = Get-VideoMetadata -YtDlpPath $ytDlpPath -BrowserString $browserString -Url $normalizedUrl

    # Execute yt-dlp with real-time output (no buffering)
    $process = Start-Process -FilePath $ytDlpPath -ArgumentList $arguments -NoNewWindow -Wait -PassThru

    # Check exit code
    if ($process.ExitCode -ne 0) {
        Write-Host ""
        Write-Host "Warning: yt-dlp exited with code $($process.ExitCode)" -ForegroundColor Yellow
    }

# Show completion message
Write-Host ""
Write-Host "[OK] Download Complete" -ForegroundColor Green

# Process and move the final MKV file
Write-Host ""
Write-Host ">> Processing Final File" -ForegroundColor Cyan

# Find video files in the temp directory (use configured output format)
$videoFiles = Get-ChildItem -Path $tempDir -Filter "*.$effectiveOutputFormat" -ErrorAction SilentlyContinue

if ($videoFiles) {
    foreach ($file in $videoFiles) {
        $originalName = $file.Name
        
        # Apply filename cleaning if enabled
        if ($shouldCleanFilenames) {
            $yearToAdd = if ($shouldAddReleaseYear) { $releaseYear } else { $null }
            $cleanedName = Get-CleanedFilename -OriginalName $originalName -ReleaseYear $yearToAdd -PhrasesToRemove $config.phrasesToRemove
        } else {
            $cleanedName = $originalName
        }

        # Rename in temp directory if needed
        $currentFilePath = $file.FullName
        $finalFileName = $originalName

        if ($cleanedName -ne $originalName) {
            $renamedPath = Join-Path $file.Directory.FullName $cleanedName
            try {
                Rename-Item -LiteralPath $currentFilePath -NewName $cleanedName -ErrorAction Stop
                Write-Host "Renamed: " -ForegroundColor Green -NoNewline
                Write-Host "$originalName" -ForegroundColor Gray
                Write-Host "      -> " -ForegroundColor DarkGray -NoNewline
                Write-Host "$cleanedName" -ForegroundColor White
                $currentFilePath = $renamedPath
                $finalFileName = $cleanedName
            } catch {
                Write-Host "Failed to rename: " -ForegroundColor Red -NoNewline
                Write-Host "$originalName" -ForegroundColor Gray
                Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
                # Keep using original path if rename failed
            }
        } else {
            Write-Host "No rename needed: " -ForegroundColor Cyan -NoNewline
            Write-Host "$originalName" -ForegroundColor White
        }

        # Move to output directory
        $destPath = Join-Path $effectiveOutputDir $finalFileName
        try {
            if (Test-Path -LiteralPath $currentFilePath) {
                Move-Item -LiteralPath $currentFilePath -Destination $destPath -Force -ErrorAction Stop
                Write-Host "Moved to: " -ForegroundColor Green -NoNewline
                Write-Host "$effectiveOutputDir" -ForegroundColor White
            } else {
                Write-Host "Failed to move file: " -ForegroundColor Red -NoNewline
                Write-Host "File not found at path: $currentFilePath" -ForegroundColor Gray
            }
        } catch {
            Write-Host "Failed to move file: " -ForegroundColor Red -NoNewline
            Write-Host "$finalFileName" -ForegroundColor Gray
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
} else {
    Write-Host "No $($effectiveOutputFormat.ToUpper()) files found in temp directory." -ForegroundColor Yellow
}
} finally {
    # Clean up temporary directory (this removes VTT files and all other temp files)
    # This runs even if yt-dlp fails or an error occurs
    Write-Host ""
    Remove-TempDirectory -TempDir $tempDir
}
