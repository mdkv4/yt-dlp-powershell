[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$Url,

    [Parameter(Mandatory=$false)]
    [string]$Browser = "firefox",

    [Parameter(Mandatory=$false)]
    [string]$OutputDir = "",

    # Chapter selection to use without prompting, e.g. "1,3,5-7", "all" or "full".
    [Parameter(Mandatory=$false)]
    [string]$Chapters = "",

    # Ignore chapters and download the whole audio as a single MP3.
    [Parameter(Mandatory=$false)]
    [switch]$NoChapters
)

$binDir = Join-Path $PSScriptRoot "bin"
$ytDlpPath = Join-Path $binDir "yt-dlp.exe"

if (-not (Test-Path $ytDlpPath)) {
    Write-Host "ERROR: yt-dlp.exe not found at $ytDlpPath" -ForegroundColor Red
    exit 1
}

$baseArgs = @(
    "--cookies-from-browser", $Browser,
    "--js-runtimes", "node"
)

$audioArgs = @(
    "-x",
    "--audio-format", "mp3",
    "--audio-quality", "0",
    # Drop the source container's tags (major_brand, encoder, ...) so only ours end up in the MP3.
    "--postprocessor-args", "ExtractAudio:-map_metadata -1"
)

# ID3 tags + cover art. yt-dlp's thumbnail embed for MP3 writes ID3v2.3 (+ v1),
# which Windows Explorer and most players read reliably.
$tagArgs = @(
    "--embed-metadata",
    "--embed-thumbnail",
    "--convert-thumbnails", "jpg"
)

$invariant = [System.Globalization.CultureInfo]::InvariantCulture
$utf8NoBom = New-Object System.Text.UTF8Encoding $false

function Format-Timestamp {
    param([double]$Seconds)
    $ts = [TimeSpan]::FromSeconds([Math]::Floor($Seconds))
    if ($ts.TotalHours -ge 1) {
        return "{0}:{1:00}:{2:00}" -f [int][Math]::Floor($ts.TotalHours), $ts.Minutes, $ts.Seconds
    }
    return "{0}:{1:00}" -f $ts.Minutes, $ts.Seconds
}

function Get-VideoInfo {
    <#
    Writes the video's info JSON to a temp folder (avoids console-encoding issues with
    non-ASCII titles) and returns the parsed object plus the folder path.
    #>
    param([string]$TempDir)

    $infoArgs = $baseArgs + @(
        "--skip-download",
        "--write-info-json",
        "--flat-playlist",
        "-o", (Join-Path $TempDir "info.%(ext)s"),
        $Url
    )

    Write-Host "Fetching video info..." -ForegroundColor DarkGray
    $output = & $ytDlpPath @infoArgs 2>&1
    $infoFile = Get-ChildItem -Path $TempDir -Filter "*.info.json" -ErrorAction SilentlyContinue | Select-Object -First 1

    if (-not $infoFile) {
        $output | ForEach-Object { "$_" } | Where-Object { $_ -match 'ERROR' } | ForEach-Object {
            Write-Host "   $_" -ForegroundColor DarkGray
        }
        Write-Verbose ($output | Out-String)
        return $null
    }

    $json = [System.IO.File]::ReadAllText($infoFile.FullName, [System.Text.Encoding]::UTF8)
    return ($json | ConvertFrom-Json)
}

function Show-ChapterList {
    param($ChapterList)

    $numWidth = ([string]$ChapterList.Count).Length
    Write-Host ""
    Write-Host "Chapters:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $ChapterList.Count; $i++) {
        $ch = $ChapterList[$i]
        $num = ([string]($i + 1)).PadLeft($numWidth)
        $start = (Format-Timestamp $ch.start_time).PadLeft(8)
        $length = Format-Timestamp ($ch.end_time - $ch.start_time)
        Write-Host ("  {0}. " -f $num) -ForegroundColor Yellow -NoNewline
        Write-Host ("{0}  " -f $start) -ForegroundColor DarkGray -NoNewline
        Write-Host $ch.title -NoNewline
        Write-Host ("  ({0})" -f $length) -ForegroundColor DarkGray
    }
    Write-Host ""
}

function ConvertTo-ChapterSelection {
    <#
    Parses "1,3,5-7" style input into a sorted list of unique 1-based chapter numbers.
    Returns $null if the input is invalid.
    #>
    param([string]$Text, [int]$Count)

    $selected = New-Object System.Collections.Generic.SortedSet[int]
    foreach ($part in ($Text -split '[,\s]+' | Where-Object { $_ -ne '' })) {
        if ($part -match '^(\d+)-(\d+)$') {
            $from = [int]$Matches[1]; $to = [int]$Matches[2]
            if ($from -gt $to) { $from, $to = $to, $from }
        } elseif ($part -match '^\d+$') {
            $from = [int]$part; $to = $from
        } else {
            return $null
        }
        if ($from -lt 1 -or $to -gt $Count) { return $null }
        for ($n = $from; $n -le $to; $n++) { [void]$selected.Add($n) }
    }
    if ($selected.Count -eq 0) { return $null }
    return ,([int[]]@($selected))
}

function Get-ChapterSelection {
    <#
    Returns "full" (no split), or an int[] of selected chapter numbers.
    Returns $null if the user quits.
    #>
    param([int]$Count, [string]$Preset)

    while ($true) {
        if ($Preset -ne "") {
            $answer = $Preset
            $Preset = ""
        } else {
            Write-Host "Select chapters to download:" -ForegroundColor Cyan
            Write-Host "  Enter / all  - every chapter as a separate MP3" -ForegroundColor Gray
            Write-Host "  e.g. 1,3,5-7 - only those chapters" -ForegroundColor Gray
            Write-Host "  full         - whole audio as one MP3 (no split)" -ForegroundColor Gray
            Write-Host "  q            - quit" -ForegroundColor Gray
            $answer = Read-Host "Chapters"
        }

        $answer = $answer.Trim().ToLowerInvariant()
        switch ($answer) {
            { $_ -in @("", "all", "a") } { return ,([int[]](1..$Count)) }
            { $_ -in @("full", "f") }    { return "full" }
            { $_ -in @("q", "quit") }    { return $null }
        }

        $parsed = ConvertTo-ChapterSelection -Text $answer -Count $Count
        if ($parsed) { return ,$parsed }

        Write-Host "Invalid selection '$answer'. Use numbers between 1 and $Count, e.g. 1,3,5-7" -ForegroundColor Red
        Write-Host ""
    }
}

function Invoke-FullDownload {
    # Tag the year only (yt-dlp would otherwise write the full YYYYMMDD upload date).
    $arguments = $baseArgs + $audioArgs + $tagArgs + @(
        "--parse-metadata", "%(release_date,upload_date)s:(?P<meta_date>\d{4})"
    )

    if ($OutputDir -ne "") {
        $arguments += "-o"
        $arguments += "$OutputDir\%(title)s [%(id)s].%(ext)s"
    }

    $arguments += $Url

    Write-Host ""
    Write-Host ">> Downloading as MP3 (highest quality)" -ForegroundColor Cyan
    Write-Host "   URL: $Url" -ForegroundColor Gray
    Write-Host ""

    # Not captured/returned so yt-dlp keeps writing straight to the console (live progress).
    & $ytDlpPath @arguments
    $script:exitCode = $LASTEXITCODE
}

function Invoke-ChapterDownload {
    param($Info, $ChapterList, [int[]]$Selection, [string]$TempDir)

    # Trim fields that are large and irrelevant for an audio download, so the
    # per-chapter info JSON stays small and fast to (re)serialize.
    # "chapters" is dropped too, so the full video's chapter markers are not embedded in each file.
    foreach ($field in @("automatic_captions", "subtitles", "heatmap", "requested_formats", "requested_downloads", "chapters")) {
        if ($Info.PSObject.Properties[$field]) { $Info.PSObject.Properties.Remove($field) }
    }
    $audioFormats = @($Info.formats | Where-Object { $_.vcodec -eq 'none' -and $_.acodec -and $_.acodec -ne 'none' })
    if ($audioFormats.Count -gt 0) { $Info.formats = $audioFormats }

    $videoTitle = $Info.title
    $artist = if ($Info.artist) { $Info.artist } elseif ($Info.channel) { $Info.channel } else { $Info.uploader }
    $total = $ChapterList.Count
    $trackWidth = [Math]::Max(2, ([string]$total).Length)

    $baseDir = if ($OutputDir -ne "") { $OutputDir } else { (Get-Location).Path }
    $outTemplate = Join-Path $baseDir ("%(album)s [%(id)s]\%(track_number)0{0}d - %(title)s.%(ext)s" -f $trackWidth)

    $dateSource = if ($Info.release_date) { [string]$Info.release_date } else { [string]$Info.upload_date }
    $year = if ($dateSource.Length -ge 4) { $dateSource.Substring(0, 4) } else { $null }

    # Fields used by the output template.
    $Info | Add-Member -NotePropertyName album -NotePropertyValue $videoTitle -Force
    $Info | Add-Member -NotePropertyName track_number -NotePropertyValue $null -Force

    # "meta_<tag>" fields override the tags yt-dlp's --embed-metadata writes (values must be strings).
    $Info | Add-Member -NotePropertyName meta_album -NotePropertyValue ([string]$videoTitle) -Force
    if ($artist) {
        $Info | Add-Member -NotePropertyName meta_artist -NotePropertyValue ([string]$artist) -Force
        $Info | Add-Member -NotePropertyName meta_album_artist -NotePropertyValue ([string]$artist) -Force
    }
    if ($year) { $Info | Add-Member -NotePropertyName meta_date -NotePropertyValue $year -Force }
    $Info | Add-Member -NotePropertyName meta_title -NotePropertyValue $null -Force
    $Info | Add-Member -NotePropertyName meta_track -NotePropertyValue $null -Force

    Write-Host ""
    Write-Host (">> Downloading {0} of {1} chapter(s) as MP3 (highest quality)" -f $Selection.Count, $total) -ForegroundColor Cyan
    Write-Host "   Video: $videoTitle" -ForegroundColor Gray
    Write-Host ""

    $failed = @()
    $index = 0
    foreach ($num in $Selection) {
        $index++
        $ch = $ChapterList[$num - 1]

        $Info.title = $ch.title
        $Info.track_number = $num
        $Info.meta_title = [string]$ch.title
        $Info.meta_track = "$num/$total"

        $chapterJson = Join-Path $TempDir ("chapter-{0}.info.json" -f $num)
        $json = $Info | ConvertTo-Json -Depth 100 -Compress
        [System.IO.File]::WriteAllText($chapterJson, $json, $utf8NoBom)

        $start = ([double]$ch.start_time).ToString($invariant)
        $end = ([double]$ch.end_time).ToString($invariant)

        Write-Host ("[{0}/{1}] Chapter {2}: {3}  ({4} - {5})" -f $index, $Selection.Count, $num, $ch.title,
            (Format-Timestamp $ch.start_time), (Format-Timestamp $ch.end_time)) -ForegroundColor Cyan

        $arguments = $baseArgs + $audioArgs + $tagArgs + @(
            "--load-info-json", $chapterJson,
            "--download-sections", "*$start-$end",
            "--no-embed-chapters",
            "--no-embed-info-json",
            "-o", $outTemplate
        )

        & $ytDlpPath @arguments
        if ($LASTEXITCODE -ne 0) { $failed += $num }
        Write-Host ""
    }

    if ($failed.Count -gt 0) {
        Write-Host ("Failed chapters: {0}" -f ($failed -join ", ")) -ForegroundColor Red
        $script:exitCode = 1
        return
    }

    Write-Host "Done. Saved under: $baseDir" -ForegroundColor Green
    $script:exitCode = 0
}

# --- Main ---

$script:exitCode = 0

if ($OutputDir -ne "" -and -not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

if ($NoChapters) {
    Invoke-FullDownload
    exit $script:exitCode
}

$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("ytdlp-mp3-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

try {
    $info = Get-VideoInfo -TempDir $tempDir
    $chapterList = @()
    if ($info -and $info._type -ne "playlist") {
        $chapterList = @($info.chapters | Where-Object { $_ -and $null -ne $_.start_time })
        # Fill in missing end times from the next chapter's start (or the video duration).
        for ($i = 0; $i -lt $chapterList.Count; $i++) {
            if ($null -eq $chapterList[$i].end_time) {
                $next = if ($i + 1 -lt $chapterList.Count) { $chapterList[$i + 1].start_time } else { $info.duration }
                $chapterList[$i] | Add-Member -NotePropertyName end_time -NotePropertyValue $next -Force
            }
        }
    }

    if ($chapterList.Count -lt 2) {
        if (-not $info) {
            Write-Host "Could not read video info; downloading without chapter selection." -ForegroundColor Yellow
        } elseif ($info._type -eq "playlist") {
            Write-Host "Playlist detected; chapter selection is only available for single videos." -ForegroundColor DarkGray
        } else {
            Write-Host "No chapters found; downloading the whole audio." -ForegroundColor DarkGray
        }
        Invoke-FullDownload
        exit $script:exitCode
    }

    Write-Host ""
    Write-Host $info.title -ForegroundColor White
    Show-ChapterList -ChapterList $chapterList

    $selection = Get-ChapterSelection -Count $chapterList.Count -Preset $Chapters
    if ($null -eq $selection) {
        Write-Host "Cancelled." -ForegroundColor Yellow
    } elseif ($selection -is [string] -and $selection -eq "full") {
        Invoke-FullDownload
    } else {
        Invoke-ChapterDownload -Info $info -ChapterList $chapterList -Selection $selection -TempDir $tempDir
    }
}
finally {
    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}

exit $script:exitCode
