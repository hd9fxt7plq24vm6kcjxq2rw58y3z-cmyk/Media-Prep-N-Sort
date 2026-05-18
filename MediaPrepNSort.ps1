<#
.SYNOPSIS
Interactive Media-Prep-N-Sort top-level folder sorter for Windows and mapped NAS drives.

.DESCRIPTION
Media-Prep-N-Sort organizes only the immediate child folders inside a media folder.
It can move those sealed top-level folders into MOVIES, TV SHOWS, MUSIC, or
BOOKS and rename only the top-level folder itself to a cleaner friendly name.

It never scans recursively. It never edits media files, subtitles, season
folders, extras, or anything inside a top-level media folder.
For MUSIC and BOOKS only, it can create one genre folder layer under the
destination folder and place the sealed top-level folder inside that genre.

.EXAMPLE
powershell -NoProfile -ExecutionPolicy Bypass -File ".\MediaPrepNSort.ps1"

.EXAMPLE
powershell -NoProfile -ExecutionPolicy Bypass -File ".\MediaPrepNSort.ps1" -DryRun

.EXAMPLE
powershell -NoProfile -ExecutionPolicy Bypass -File ".\MediaPrepNSort.ps1" -SaveReports
# Save plan and execution logs beside the script.

.EXAMPLE
powershell -NoProfile -ExecutionPolicy Bypass -File ".\MediaPrepNSort.ps1" -OnlineLookup
# Use Wikidata, TVmaze, MusicBrainz, and Open Library suggestions before confirmation.

.EXAMPLE
powershell -NoProfile -ExecutionPolicy Bypass -File ".\MediaPrepNSort.ps1" -OnlineLookup -FullCheckup
# Also inspect folders already inside generated category and known genre folders.
#>

[CmdletBinding()]
param(
    [string]$Root = "",
    [string]$ReportFolder = "",
    [switch]$SaveReports,
    [switch]$OnlineLookup,
    [switch]$DryRun,
    [switch]$NoNameReview,
    [switch]$TrustMediumConfidence,
    [switch]$FullCheckup,
    [switch]$NonInteractive
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ScriptVersion = "1.3.0"
$MoviesFolderName = "MOVIES"
$ShowsFolderName = "TV SHOWS"
$MusicFolderName = "MUSIC"
$BooksFolderName = "BOOKS"
$DestinationFolderNames = @($MoviesFolderName, $ShowsFolderName, $MusicFolderName, $BooksFolderName)
$MusicGenreNames = @("Rock", "Pop", "Hip-Hop", "R&B-Soul", "Electronic", "Country", "Jazz", "Classical", "Metal", "Soundtracks", "Folk", "Other")
$BookGenreNames = @("Fiction", "Fantasy", "Sci-Fi", "Mystery-Thriller", "Horror", "Romance", "Biography-Memoir", "History", "Self-Help", "Comics-Manga", "Audiobooks", "Nonfiction", "Other")
$ManualReview = "MANUAL REVIEW"
$Ignored = "IGNORED"
$CurrentYearLimit = (Get-Date).Year + 1
$RunStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OnlineLookupUserAgent = "MediaPrepNSort/$ScriptVersion (https://github.com/hd9fxt7plq24vm6kcjxq2rw58y3z-cmyk/Media-Prep-N-Sort; personal media organizer)"
$script:CancelWords = @("q", "quit", "c", "cancel", "exit", "stop")
$script:LastMusicBrainzRequestAt = [datetime]::MinValue

function Write-Info {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Cyan
}

function Write-Warn {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Yellow
}

function Write-Good {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Green
}

function Test-CancelInput {
    param([AllowNull()][string]$Text)

    if ($null -eq $Text) {
        return $true
    }

    $clean = $Text.Trim().ToLowerInvariant()
    return ($script:CancelWords -contains $clean)
}

function Stop-MediaPrepProcess {
    param([string]$Message = "Cancelled by user. Nothing was moved or renamed.")

    throw [System.OperationCanceledException]::new($Message)
}

function Read-LineExact {
    param([string]$Prompt)

    Write-Host -NoNewline $Prompt
    return [Console]::ReadLine()
}

function Show-Welcome {
    Write-Host ""
    Write-Info "Media-Prep-N-Sort"
    Write-Host "This wizard sorts only the top-level folders in your media folder into MOVIES, TV SHOWS, MUSIC, and BOOKS."
    Write-Host "It does not scan inside those folders or rename media files, subtitles, season folders, or extras."
    Write-Host "For music and books, it can create one genre folder layer such as MUSIC\Rock or BOOKS\Fantasy."
    Write-Host "If a folder is unclear, you can choose 1 for movies, 2 for TV shows, 3 for music, 4 for books, or s to skip."
    Write-Host "Optional: run with -OnlineLookup to use public metadata sources for better suggestions."
    Write-Host "Optional: run with -FullCheckup to also inspect folders already inside MOVIES, TV SHOWS, MUSIC, BOOKS, and known genre folders."
    Write-Host "At any prompt before the final move starts, type Q, CANCEL, EXIT, or STOP to quit cleanly."
    Write-Host ""
    Write-Host "Path tip:"
    Write-Host "  1. Open your media folder in File Explorer."
    Write-Host "  2. Click the address bar at the top of the window."
    Write-Host "  3. Press Ctrl+C to copy the folder path."
    Write-Host "  4. Come back here and press Ctrl+V, or right-click, to paste it."
    Write-Host ""
    Write-Host "Example path: Z:\Media\Completed Downloads"
    Write-Host ""
}

function Normalize-Spaces {
    param([AllowNull()][string]$Text)

    if ($null -eq $Text) {
        return ""
    }

    return (($Text -replace "\s+", " ").Trim())
}

function Convert-Separators {
    param([string]$Text)

    $converted = $Text -replace "[._]+", " "
    $converted = $converted -replace "\s*-\s*", " - "
    $converted = $converted -replace "\s+", " "
    return $converted.Trim()
}

function Remove-InvalidFolderCharacters {
    param([string]$Text)

    $invalid = [Regex]::Escape(([System.IO.Path]::GetInvalidFileNameChars() -join ""))
    $clean = $Text -replace "[$invalid]", " "
    return (Normalize-Spaces $clean)
}

function Convert-ToTitleCaseLite {
    param([string]$Text)

    $text = Normalize-Spaces (Remove-InvalidFolderCharacters $Text)
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $text
    }

    $smallWords = @("and", "or", "the", "of", "a", "an", "to", "in", "on", "for", "with", "vs", "n")
    $acronyms = @("OVA", "ONA", "OST", "OP", "ED", "NCOP", "NCED", "TV", "DC", "BBC", "USA", "UK")
    $parts = @($text -split " ")
    $out = @()

    for ($i = 0; $i -lt $parts.Count; $i++) {
        $word = $parts[$i].Trim()
        if ([string]::IsNullOrWhiteSpace($word)) {
            continue
        }

        $lower = $word.ToLowerInvariant()
        $upper = $word.ToUpperInvariant()

        if ($acronyms -contains $upper) {
            $out += $upper
        }
        elseif (($i -gt 0) -and ($smallWords -contains $lower)) {
            $out += $lower
        }
        elseif ($word -cmatch "^[IVXLCDM]+$") {
            $out += $word
        }
        elseif ($word -match "^\d+$") {
            $out += $word
        }
        else {
            $out += ($lower.Substring(0, 1).ToUpperInvariant() + $lower.Substring(1))
        }
    }

    return (($out -join " ") -replace "\s+!", "!")
}

function Get-YearMatches {
    param([string]$Text)

    $matches = [regex]::Matches($Text, "(?<!\d)(19\d{2}|20\d{2})(?!\d)")
    $years = @()

    foreach ($match in $matches) {
        $year = [int]$match.Value
        if ($year -le $CurrentYearLimit) {
            $years += $year
        }
    }

    return @($years)
}

function Get-FirstUsefulYear {
    param([string]$Text)

    $rangeMatch = [regex]::Match($Text, "(?<!\d)((?:19|20)\d{2})\s*-\s*((?:19|20)\d{2})(?!\d)")
    if ($rangeMatch.Success) {
        return [int]$rangeMatch.Groups[1].Value
    }

    $years = @(Get-YearMatches $Text)
    if ($years.Count -gt 0) {
        return $years[0]
    }

    return $null
}

function Remove-QualityAndReleaseNoise {
    param([string]$Text)

    $clean = $Text
    $clean = $clean -replace "\[[^\]]*(1080p|2160p|720p|480p|BluRay|BDRip|WEB|WEB-DL|WEBRip|HDRip|x264|x265|HEVC|H\.?264|H\.?265|AAC|DTS|Dual Audio|Multi|Proper|Repack|Batch)[^\]]*\]", " "
    $clean = $clean -replace "\(([^)]*(1080p|2160p|720p|480p|BluRay|BDRip|WEB|WEB-DL|WEBRip|HDRip|x264|x265|HEVC|H\.?264|H\.?265|AAC|DTS|Dual Audio|Multi|Proper|Repack|Batch)[^)]*)\)", " "
    $clean = $clean -replace "\b(480p|720p|1080p|2160p|4K|UHD|HDR|BluRay|BDRip|BRRip|WEB-DL|WEBRip|WEB|HDRip|NF|AMZN|HMAX|DSNP|AAC\d?\.?\d?|DTS|H\.?264|H\.?265|x264|x265|HEVC|AVC|Proper|Repack|Remux|Unrated|Extended|Remastered|Uncensored|Dual Audio|Multi|Batch)\b", " "
    $clean = $clean -replace "\b(YIFY|RARBG|VARYG|EVO|NTb|FGT)\b", " "
    return (Normalize-Spaces $clean)
}

function Get-ProposedName {
    param(
        [string]$FolderName,
        [string]$Category
    )

    $name = Convert-Separators $FolderName
    $name = $name -replace "^\[[^\]]+\]\s*", ""
    $name = Remove-QualityAndReleaseNoise $name

    if ($Category -eq $MoviesFolderName) {
        $years = @(Get-YearMatches $FolderName)
        $year = $null
        if ($years.Count -gt 0) {
            $year = $years[$years.Count - 1]
        }

        $title = $name
        if ($null -ne $year) {
            $title = $title -replace "(?<!\d)$year(?!\d)", " "
        }

        $title = $title -replace "\b(?:19|20)\d{2}\s*-\s*(?:19|20)\d{2}\b", " "
        $title = $title -replace "\b\d+\s*-\s*Film Collection\b", "Collection"
        $title = $title -replace "\b\d+\s*Film Collection\b", "Collection"
        $title = $title -replace "\b\d+\s*Movie Collection\b", "Collection"
        $title = $title -replace "\b\d+\s*Movies\b", "Collection"
        $title = $title -replace "\b(The\s+)?Complete\s+Collection\b", "Collection"
        $title = $title.Trim(" ", "-", ".", "_", "+")
        $title = Convert-ToTitleCaseLite $title

        if ($FolderName -match "\b(Collection|Movie Collection|Film Collection|Movies)\b" -and $title -notmatch "\bCollection$") {
            $title = Normalize-Spaces "$title Collection"
        }

        if ($null -ne $year -and $title -notmatch "\bCollection$") {
            return "$title ($year)"
        }

        return $title
    }

    if ($Category -eq $ShowsFolderName) {
        $year = Get-FirstUsefulYear $FolderName
        $title = $name

        $title = $title -replace "\bS\d{1,2}E\d{1,3}\b", " "
        $title = $title -replace "\bS\d{1,2}(?:\s*[-+,&]\s*S?\d{1,2})*\b", " "
        $title = $title -replace "\bSeason(?:s)?\s*\d{1,2}(?:\s*[-+,&]\s*\d{1,2})*\b", " "
        $title = $title -replace "\((?:Season|Seasons|S)\s*[^)]*\)", " "
        $title = $title -replace "\b(Complete Series|The Complete Series|Complete Box Set|The Complete Box Set)\b", " "
        $title = $title -replace "\b(OVAs?|ONAs?|Specials?|Episodes?|Eps?|EPs?|Extras?|OST|NC\s*OP|NCOP|NC\s*ED|NCED|Batch)\b", " "
        $title = $title -replace "(?<!\d)\d{1,3}\s*(?:~|-)\s*\d{1,3}(?!\d)", " "
        $title = $title -replace "\b(?:19|20)\d{2}\s*-\s*(?:19|20)\d{2}\b", " "

        if ($null -ne $year) {
            $title = $title -replace "(?<!\d)$year(?!\d)", " "
        }

        $title = $title.Trim(" ", "-", ".", "_", "+")
        $title = Convert-ToTitleCaseLite $title

        if ($null -ne $year -and $title -notmatch "\(\d{4}\)$") {
            return "$title ($year)"
        }

        return $title
    }

    if ($Category -eq $MusicFolderName) {
        $title = $name
        $title = $title -replace "\b(Discography|Complete Discography|Album|Albums|EP|LP|Single|Singles|OST|Soundtrack|Deluxe Edition|Remaster(?:ed)?|Lossless|FLAC|MP3|M4A|AAC|WAV|Vinyl|CD)\b", " "
        $title = $title.Trim(" ", "-", ".", "_", "+")
        return (Convert-ToTitleCaseLite $title)
    }

    if ($Category -eq $BooksFolderName) {
        $title = $name
        $title = $title -replace "\b(Audiobooks?|Audio Books?|Ebooks?|E-Books?|EPUB|MOBI|AZW3|PDF|CBR|CBZ|Kindle|Retail|Scanned|Scan|Library)\b", " "
        $title = $title.Trim(" ", "-", ".", "_", "+")
        return (Convert-ToTitleCaseLite $title)
    }

    $name = $name.Trim(" ", "-", ".", "_", "+")
    return (Convert-ToTitleCaseLite $name)
}

function Get-Classification {
    param([string]$FolderName)

    $name = Convert-Separators $FolderName
    $years = @(Get-YearMatches $FolderName)
    $hasYearRange = ($FolderName -match "(?<!\d)(?:19|20)\d{2}\s*-\s*(?:19|20)\d{2}(?!\d)")
    $hasShowIndicator = $false
    $showReason = ""
    $hasMovieIndicator = $false
    $movieReason = ""
    $hasMusicIndicator = $false
    $musicReason = ""
    $hasBookIndicator = $false
    $bookReason = ""

    if ($name -match "\bS\d{1,2}E\d{1,3}\b") {
        $hasShowIndicator = $true
        $showReason = "SxxExx episode pattern"
    }
    elseif ($name -match "\bSeason(?:s)?\b|\bSeason\s*\d{1,2}\b|\bS\d{1,2}\b") {
        $hasShowIndicator = $true
        $showReason = "season marker"
    }
    elseif ($name -match "\b(Episode|Episodes|Eps|EPs|EP)\b") {
        $hasShowIndicator = $true
        $showReason = "episode wording"
    }
    elseif ($name -match "\b(OVA|OVAs|ONA|ONAs|Special|Specials)\b") {
        $hasShowIndicator = $true
        $showReason = "OVA/ONA/specials marker"
    }
    elseif ($name -match "\b(Complete Series|The Complete Series|Complete Box Set|The Complete Box Set)\b") {
        $hasShowIndicator = $true
        $showReason = "complete series wording"
    }
    elseif ($name -match "(?<!\d)\d{1,3}\s*(~|-)\s*\d{1,3}(?!\d)") {
        $hasShowIndicator = $true
        $showReason = "episode-range style numbering"
    }

    if ($name -match "\b(Movie Collection|Film Collection|\d+\s*Movie Collection|\d+\s*Film Collection|\d+\s*-\s*Film Collection|\d+\s*Movies)\b") {
        $hasMovieIndicator = $true
        $movieReason = "movie collection wording"
    }
    elseif ($name -match "\bThe Movie\b" -and $years.Count -gt 0) {
        $hasMovieIndicator = $true
        $movieReason = "movie wording plus release year"
    }

    if ($name -match "\b(Discography|Complete Discography|Album|Albums|EP|LP|Single|Singles|OST|Soundtrack|Original Soundtrack|FLAC|MP3|M4A|AAC|WAV|Lossless|Vinyl|CD)\b") {
        $hasMusicIndicator = $true
        $musicReason = "music/audio release wording"
    }

    if ($name -match "\b(Audiobooks?|Audio Books?|Ebooks?|E-Books?|EPUB|MOBI|AZW3|PDF|Kindle|Novel|Novels|Books?|Book Series|Manga|Comic|Comics|CBR|CBZ)\b") {
        $hasBookIndicator = $true
        $bookReason = "book/ebook/audiobook wording"
    }

    $indicatorCount = @(@($hasShowIndicator, $hasMovieIndicator, $hasMusicIndicator, $hasBookIndicator) | Where-Object { $_ }).Count

    if ($indicatorCount -gt 1) {
        return [pscustomobject]@{
            Category = $ManualReview
            Confidence = "Low"
            NeedsPlacementReview = $true
            Reason = "Multiple category indicators were found; please choose the right destination."
        }
    }

    if ($hasShowIndicator) {
        return [pscustomobject]@{
            Category = $ShowsFolderName
            Confidence = "High"
            NeedsPlacementReview = $false
            Reason = "TV/show indicator found: $showReason."
        }
    }

    if ($hasMovieIndicator) {
        return [pscustomobject]@{
            Category = $MoviesFolderName
            Confidence = "High"
            NeedsPlacementReview = $false
            Reason = "Movie indicator found: $movieReason."
        }
    }

    if ($hasMusicIndicator) {
        return [pscustomobject]@{
            Category = $MusicFolderName
            Confidence = "High"
            NeedsPlacementReview = $false
            Reason = "Music indicator found: $musicReason."
        }
    }

    if ($hasBookIndicator) {
        return [pscustomobject]@{
            Category = $BooksFolderName
            Confidence = "High"
            NeedsPlacementReview = $false
            Reason = "Book indicator found: $bookReason."
        }
    }

    if ($years.Count -eq 1 -and -not $hasYearRange) {
        return [pscustomobject]@{
            Category = $MoviesFolderName
            Confidence = "Medium"
            NeedsPlacementReview = $true
            Reason = "Single release-year style name, but a year alone is not enough to trust blindly."
        }
    }

    if ($name -match "\b(480p|720p|1080p|2160p|4K|BluRay|BDRip|WEB-DL|WEBRip|WEB)\b") {
        return [pscustomobject]@{
            Category = $ManualReview
            Confidence = "Low"
            NeedsPlacementReview = $true
            Reason = "Quality tags found, but no reliable category indicator."
        }
    }

    return [pscustomobject]@{
        Category = $ManualReview
        Confidence = "Low"
        NeedsPlacementReview = $true
        Reason = "No reliable movie, TV, music, or book indicator from the top-level folder name."
    }
}

function Get-GenreSuggestion {
    param(
        [string]$FolderName,
        [string]$Category
    )

    $name = Convert-Separators $FolderName

    if ($Category -eq $MusicFolderName) {
        if ($name -match "\b(Soundtrack|OST|Score|Original Soundtrack)\b") { return "Soundtracks" }
        if ($name -match "\b(Hip Hop|Hip-Hop|Rap|Trap)\b") { return "Hip-Hop" }
        if ($name -match "\b(R&B|RNB|Soul)\b") { return "R&B-Soul" }
        if ($name -match "\b(Electronic|EDM|House|Techno|Trance|Dubstep|Drum and Bass|DNB|Ambient|Synthwave)\b") { return "Electronic" }
        if ($name -match "\b(Country|Bluegrass)\b") { return "Country" }
        if ($name -match "\b(Jazz|Blues)\b") { return "Jazz" }
        if ($name -match "\b(Classical|Opera|Symphony|Orchestra)\b") { return "Classical" }
        if ($name -match "\b(Metal|Death Metal|Black Metal|Metalcore)\b") { return "Metal" }
        if ($name -match "\b(Rock|Punk|Alternative|Alt Rock|Indie Rock|Grunge)\b") { return "Rock" }
        if ($name -match "\b(Pop|K-Pop|J-Pop)\b") { return "Pop" }
        if ($name -match "\b(Folk|Acoustic|Singer Songwriter|Singer-Songwriter)\b") { return "Folk" }
    }

    if ($Category -eq $BooksFolderName) {
        if ($name -match "\b(Audiobooks?|Audio Books?)\b") { return "Audiobooks" }
        if ($name -match "\b(Comic|Comics|Manga|Graphic Novel|CBR|CBZ)\b") { return "Comics-Manga" }
        if ($name -match "\b(Sci Fi|Sci-Fi|Science Fiction)\b") { return "Sci-Fi" }
        if ($name -match "\b(Fantasy|Epic Fantasy|Urban Fantasy)\b") { return "Fantasy" }
        if ($name -match "\b(Mystery|Thriller|Crime|Detective|Suspense)\b") { return "Mystery-Thriller" }
        if ($name -match "\b(Horror|Supernatural)\b") { return "Horror" }
        if ($name -match "\b(Romance|Romantic)\b") { return "Romance" }
        if ($name -match "\b(Biography|Autobiography|Memoir)\b") { return "Biography-Memoir" }
        if ($name -match "\b(History|Historical)\b") { return "History" }
        if ($name -match "\b(Self Help|Self-Help|Productivity|Personal Development)\b") { return "Self-Help" }
        if ($name -match "\b(Nonfiction|Non-Fiction|Reference|Textbook|Education)\b") { return "Nonfiction" }
        if ($name -match "\b(Novel|Fiction|Books?|Ebooks?|E-Books?)\b") { return "Fiction" }
    }

    return "Other"
}

function Get-ObjectValue {
    param(
        [AllowNull()][object]$Object,
        [string]$Name
    )

    if ($null -eq $Object) {
        return $null
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function Get-LookupQuery {
    param([string]$FolderName)

    $query = Convert-Separators $FolderName
    $query = Remove-QualityAndReleaseNoise $query
    $query = $query -replace "\[[^\]]+\]", " "
    $query = $query -replace "\((?:Season|Seasons|S)\s*[^)]*\)", " "
    $query = $query -replace "\bS\d{1,2}E\d{1,3}\b", " "
    $query = $query -replace "\bS\d{1,2}(?:\s*[-+,&]\s*S?\d{1,2})*\b", " "
    $query = $query -replace "\bSeason(?:s)?\s*\d{1,2}(?:\s*[-+,&]\s*\d{1,2})*\b", " "
    $query = $query -replace "(?<!\d)\d{1,3}\s*(?:~|-)\s*\d{1,3}(?!\d)", " "
    $query = $query -replace "\b(Complete|Collection|Series|Movies?|OVAs?|ONAs?|Specials?|Episodes?|Eps?|EPs?|Batch|Discography|Albums?|Audiobooks?|Ebooks?|E-Books?)\b", " "
    $query = $query.Trim(" ", "-", ".", "_", "+")
    return (Normalize-Spaces $query)
}

function Get-ComparableTokens {
    param([string]$Text)

    $clean = $Text.ToLowerInvariant()
    $clean = $clean -replace "[^a-z0-9]+", " "
    $tokens = @($clean -split "\s+" | Where-Object {
        $_ -and @("the", "a", "an", "and", "or", "of", "to", "in", "on", "for", "with", "complete", "collection") -notcontains $_
    })

    return @($tokens)
}

function Get-TokenOverlapScore {
    param(
        [string]$Query,
        [string]$Candidate
    )

    $queryTokens = @(Get-ComparableTokens $Query)
    $candidateTokens = @(Get-ComparableTokens $Candidate)

    if ($queryTokens.Count -eq 0 -or $candidateTokens.Count -eq 0) {
        return 0.0
    }

    $matches = 0
    foreach ($token in $queryTokens) {
        if ($candidateTokens -contains $token) {
            $matches++
        }
    }

    return [math]::Round(($matches / [math]::Max($queryTokens.Count, 1)), 2)
}

function Convert-MetadataGenre {
    param(
        [string]$Text,
        [string]$Category
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return "Other"
    }

    $value = $Text.ToLowerInvariant()

    if ($Category -eq $MusicFolderName) {
        if ($value -match "soundtrack|score|ost") { return "Soundtracks" }
        if ($value -match "hip hop|hip-hop|rap|trap") { return "Hip-Hop" }
        if ($value -match "r&b|rhythm|soul") { return "R&B-Soul" }
        if ($value -match "electronic|edm|house|techno|trance|dubstep|ambient|synth") { return "Electronic" }
        if ($value -match "country|bluegrass") { return "Country" }
        if ($value -match "jazz|blues") { return "Jazz" }
        if ($value -match "classical|opera|symphony|orchestra") { return "Classical" }
        if ($value -match "metal") { return "Metal" }
        if ($value -match "rock|punk|alternative|indie|grunge") { return "Rock" }
        if ($value -match "pop") { return "Pop" }
        if ($value -match "folk|acoustic|singer") { return "Folk" }
    }

    if ($Category -eq $BooksFolderName) {
        if ($value -match "audio") { return "Audiobooks" }
        if ($value -match "comic|manga|graphic") { return "Comics-Manga" }
        if ($value -match "science fiction|sci fi|sci-fi") { return "Sci-Fi" }
        if ($value -match "fantasy") { return "Fantasy" }
        if ($value -match "mystery|thriller|crime|detective|suspense") { return "Mystery-Thriller" }
        if ($value -match "horror|supernatural") { return "Horror" }
        if ($value -match "romance") { return "Romance" }
        if ($value -match "biography|autobiography|memoir") { return "Biography-Memoir" }
        if ($value -match "history|historical") { return "History" }
        if ($value -match "self help|self-help|productivity|personal development") { return "Self-Help" }
        if ($value -match "nonfiction|non-fiction|reference|textbook|education") { return "Nonfiction" }
        if ($value -match "fiction|novel") { return "Fiction" }
    }

    if ($Category -eq $MoviesFolderName) {
        if ($value -match "science fiction|sci fi|sci-fi") { return "Sci-Fi" }
        if ($value -match "fantasy") { return "Fantasy" }
        if ($value -match "action|adventure|superhero|martial arts") { return "Action-Adventure" }
        if ($value -match "comedy|satire") { return "Comedy" }
        if ($value -match "drama") { return "Drama" }
        if ($value -match "horror|slasher|supernatural") { return "Horror" }
        if ($value -match "thriller|mystery|crime|detective|suspense") { return "Mystery-Thriller" }
        if ($value -match "romance|romantic") { return "Romance" }
        if ($value -match "animation|anime|animated") { return "Animation" }
        if ($value -match "documentary") { return "Documentary" }
        if ($value -match "musical") { return "Musical" }
        if ($value -match "western") { return "Western" }
        if ($value -match "war") { return "War" }
    }

    return "Other"
}

function New-OnlineSuggestion {
    param(
        [string]$Source,
        [string]$Category,
        [string]$CleanName,
        [string]$Genre,
        [string]$MatchName,
        [double]$Score,
        [string]$Reason
    )

    $confidence = "Low"
    if ($Score -ge 0.75) {
        $confidence = "High"
    }
    elseif ($Score -ge 0.45) {
        $confidence = "Medium"
    }

    return [pscustomobject]@{
        Source = $Source
        Category = $Category
        CleanName = Remove-InvalidFolderCharacters $CleanName
        Genre = Remove-InvalidFolderCharacters $Genre
        MatchName = $MatchName
        Score = $Score
        Confidence = $confidence
        Reason = $Reason
    }
}

function Invoke-OnlineJson {
    param(
        [string]$Uri,
        [string]$Source
    )

    try {
        if ($Source -eq "MusicBrainz") {
            $elapsed = ([datetime]::UtcNow - $script:LastMusicBrainzRequestAt).TotalMilliseconds
            if ($elapsed -lt 1200) {
                Start-Sleep -Milliseconds ([int](1200 - $elapsed))
            }
            $script:LastMusicBrainzRequestAt = [datetime]::UtcNow
        }

        $headers = @{ "User-Agent" = $OnlineLookupUserAgent }
        return Invoke-RestMethod -Uri $Uri -Headers $headers -TimeoutSec 12 -ErrorAction Stop
    }
    catch {
        Write-Warn "Online lookup skipped for ${Source}: $($_.Exception.Message)"
        return $null
    }
}

function Search-TVMetadata {
    param([string]$Query)

    if ([string]::IsNullOrWhiteSpace($Query)) {
        return $null
    }

    $encoded = [uri]::EscapeDataString($Query)
    $results = Invoke-OnlineJson -Uri "https://api.tvmaze.com/search/shows?q=$encoded" -Source "TVmaze"
    $first = @($results | Select-Object -First 1)
    if ($first.Count -eq 0 -or $null -eq $first[0]) {
        return $null
    }

    $show = Get-ObjectValue -Object $first[0] -Name "show"
    $title = [string](Get-ObjectValue -Object $show -Name "name")
    if ([string]::IsNullOrWhiteSpace($title)) {
        return $null
    }

    $premiered = [string](Get-ObjectValue -Object $show -Name "premiered")
    $year = ""
    if ($premiered -match "^\d{4}") {
        $year = $matches[0]
    }

    $cleanName = if ($year) { "$title ($year)" } else { $title }
    $genres = @(Get-ObjectValue -Object $show -Name "genres")
    $matchName = if ($genres.Count -gt 0) { "$title [$($genres -join ', ')]" } else { $title }
    $score = Get-TokenOverlapScore -Query $Query -Candidate $title

    if ($score -lt 0.35) {
        return $null
    }

    return New-OnlineSuggestion -Source "TVmaze" -Category $ShowsFolderName -CleanName $cleanName -Genre "" -MatchName $matchName -Score $score -Reason "Matched TV show metadata from TVmaze."
}

function Search-BookMetadata {
    param([string]$Query)

    if ([string]::IsNullOrWhiteSpace($Query)) {
        return $null
    }

    $encoded = [uri]::EscapeDataString($Query)
    $uri = "https://openlibrary.org/search.json?q=$encoded&limit=1&fields=title,author_name,subject,first_publish_year"
    $results = Invoke-OnlineJson -Uri $uri -Source "Open Library"
    $docs = @(Get-ObjectValue -Object $results -Name "docs")
    if ($docs.Count -eq 0 -or $null -eq $docs[0]) {
        return $null
    }

    $doc = $docs[0]
    $title = [string](Get-ObjectValue -Object $doc -Name "title")
    if ([string]::IsNullOrWhiteSpace($title)) {
        return $null
    }

    $authors = @(Get-ObjectValue -Object $doc -Name "author_name")
    $author = if ($authors.Count -gt 0) { [string]$authors[0] } else { "" }
    $cleanName = if ($author) { "$author - $title" } else { $title }

    $subjects = @(Get-ObjectValue -Object $doc -Name "subject")
    $genre = "Other"
    foreach ($subject in $subjects) {
        $mapped = Convert-MetadataGenre -Text ([string]$subject) -Category $BooksFolderName
        if ($mapped -ne "Other") {
            $genre = $mapped
            break
        }
    }

    $candidate = if ($author) { "$author $title" } else { $title }
    $score = Get-TokenOverlapScore -Query $Query -Candidate $candidate
    if ($score -lt 0.35) {
        return $null
    }

    return New-OnlineSuggestion -Source "Open Library" -Category $BooksFolderName -CleanName $cleanName -Genre $genre -MatchName $candidate -Score $score -Reason "Matched book metadata from Open Library."
}

function Search-MusicMetadata {
    param([string]$Query)

    if ([string]::IsNullOrWhiteSpace($Query)) {
        return $null
    }

    $encoded = [uri]::EscapeDataString($Query)
    $results = Invoke-OnlineJson -Uri "https://musicbrainz.org/ws/2/release-group/?query=$encoded&fmt=json&limit=1" -Source "MusicBrainz"
    $groups = @(Get-ObjectValue -Object $results -Name "release-groups")
    if ($groups.Count -eq 0 -or $null -eq $groups[0]) {
        return $null
    }

    $group = $groups[0]
    $title = [string](Get-ObjectValue -Object $group -Name "title")
    if ([string]::IsNullOrWhiteSpace($title)) {
        return $null
    }

    $artistCredits = @(Get-ObjectValue -Object $group -Name "artist-credit")
    $artist = ""
    if ($artistCredits.Count -gt 0) {
        $artist = [string](Get-ObjectValue -Object $artistCredits[0] -Name "name")
    }

    $cleanName = if ($artist) { "$artist - $title" } else { $title }
    $matchName = if ($artist) { "$artist $title" } else { $title }
    $genre = "Other"

    $tags = @(Get-ObjectValue -Object $group -Name "tags")
    foreach ($tag in $tags) {
        $tagName = [string](Get-ObjectValue -Object $tag -Name "name")
        $mapped = Convert-MetadataGenre -Text $tagName -Category $MusicFolderName
        if ($mapped -ne "Other") {
            $genre = $mapped
            break
        }
    }

    $score = Get-TokenOverlapScore -Query $Query -Candidate $matchName
    if ($score -lt 0.35) {
        return $null
    }

    return New-OnlineSuggestion -Source "MusicBrainz" -Category $MusicFolderName -CleanName $cleanName -Genre $genre -MatchName $matchName -Score $score -Reason "Matched music release metadata from MusicBrainz."
}

function Search-MovieMetadata {
    param([string]$Query)

    if ([string]::IsNullOrWhiteSpace($Query)) {
        return $null
    }

    $encoded = [uri]::EscapeDataString($Query)
    $searchUri = "https://www.wikidata.org/w/api.php?action=wbsearchentities&search=$encoded&language=en&format=json&limit=5&type=item"
    $searchResults = Invoke-OnlineJson -Uri $searchUri -Source "Wikidata"
    $searchItems = @(Get-ObjectValue -Object $searchResults -Name "search")
    if ($searchItems.Count -eq 0) {
        return $null
    }

    $ids = @()
    foreach ($item in $searchItems) {
        $id = [string](Get-ObjectValue -Object $item -Name "id")
        if ($id -match "^Q\d+$") {
            $ids += "wd:$id"
        }
    }

    if ($ids.Count -eq 0) {
        return $null
    }

    $values = $ids -join " "
    $sparql = @"
SELECT ?film ?filmLabel ?date ?genreLabel WHERE {
  VALUES ?film { $values }
  ?film wdt:P31/wdt:P279* wd:Q11424 .
  OPTIONAL { ?film wdt:P577 ?date . }
  OPTIONAL { ?film wdt:P136 ?genre . }
  SERVICE wikibase:label { bd:serviceParam wikibase:language "en" . }
}
"@

    $sparqlEncoded = [uri]::EscapeDataString($sparql)
    $queryUri = "https://query.wikidata.org/sparql?format=json&query=$sparqlEncoded"
    $queryResults = Invoke-OnlineJson -Uri $queryUri -Source "Wikidata"
    $resultsObject = Get-ObjectValue -Object $queryResults -Name "results"
    $bindings = @(Get-ObjectValue -Object $resultsObject -Name "bindings")
    if ($bindings.Count -eq 0) {
        return $null
    }

    $bestSuggestion = $null
    foreach ($binding in $bindings) {
        $labelObj = Get-ObjectValue -Object $binding -Name "filmLabel"
        $title = [string](Get-ObjectValue -Object $labelObj -Name "value")
        if ([string]::IsNullOrWhiteSpace($title)) {
            continue
        }

        $year = ""
        $dateObj = Get-ObjectValue -Object $binding -Name "date"
        $dateValue = [string](Get-ObjectValue -Object $dateObj -Name "value")
        if ($dateValue -match "^\d{4}") {
            $year = $matches[0]
        }

        $genre = "Other"
        $genreObj = Get-ObjectValue -Object $binding -Name "genreLabel"
        $genreValue = [string](Get-ObjectValue -Object $genreObj -Name "value")
        if (-not [string]::IsNullOrWhiteSpace($genreValue)) {
            $genre = Convert-MetadataGenre -Text $genreValue -Category $MoviesFolderName
        }

        $cleanName = if ($year) { "$title ($year)" } else { $title }
        $score = Get-TokenOverlapScore -Query $Query -Candidate $title
        if ($score -lt 0.35) {
            continue
        }

        $suggestion = New-OnlineSuggestion -Source "Wikidata" -Category $MoviesFolderName -CleanName $cleanName -Genre $genre -MatchName $title -Score $score -Reason "Matched verified film metadata from Wikidata."
        if ($null -eq $bestSuggestion -or $suggestion.Score -gt $bestSuggestion.Score) {
            $bestSuggestion = $suggestion
        }
        elseif ($suggestion.Score -eq $bestSuggestion.Score -and $bestSuggestion.Genre -eq "Other" -and $suggestion.Genre -ne "Other") {
            $bestSuggestion = $suggestion
        }
    }

    return $bestSuggestion
}

function Get-OnlineMetadataSuggestion {
    param(
        [string]$FolderName,
        [string]$LocalCategory
    )

    $query = Get-LookupQuery -FolderName $FolderName
    if ([string]::IsNullOrWhiteSpace($query)) {
        return $null
    }

    Write-Info "Looking up metadata for: $FolderName"

    $suggestions = @()
    if ($LocalCategory -eq $MoviesFolderName) {
        $suggestions += @(Search-MovieMetadata -Query $query)
    }
    elseif ($LocalCategory -eq $ShowsFolderName) {
        $suggestions += @(Search-TVMetadata -Query $query)
    }
    elseif ($LocalCategory -eq $MusicFolderName) {
        $suggestions += @(Search-MusicMetadata -Query $query)
    }
    elseif ($LocalCategory -eq $BooksFolderName) {
        $suggestions += @(Search-BookMetadata -Query $query)
    }
    else {
        $suggestions += @(Search-MovieMetadata -Query $query)
        $suggestions += @(Search-TVMetadata -Query $query)
        $suggestions += @(Search-BookMetadata -Query $query)
        $suggestions += @(Search-MusicMetadata -Query $query)
    }

    $suggestions = @($suggestions | Where-Object { $null -ne $_ } | Sort-Object Score -Descending)
    if ($suggestions.Count -eq 0) {
        return $null
    }

    return $suggestions[0]
}

function Confirm-OnlineSuggestion {
    param(
        [string]$FolderName,
        [object]$Suggestion
    )

    Write-Host ""
    Write-Info "Online metadata suggestion"
    Write-Host "Folder:     $FolderName"
    Write-Host "Source:     $($Suggestion.Source)"
    Write-Host "Match:      $($Suggestion.MatchName)"
    Write-Host "Category:   $($Suggestion.Category)"
    if (-not [string]::IsNullOrWhiteSpace($Suggestion.Genre)) {
        Write-Host "Genre:      $($Suggestion.Genre)"
    }
    Write-Host "Clean name: $($Suggestion.CleanName)"
    Write-Host "Confidence: $($Suggestion.Confidence) ($($Suggestion.Score))"
    Write-Host "Reason:     $($Suggestion.Reason)"
    return (Prompt-YesNo -Question "Use this online suggestion?" -DefaultYes ($Suggestion.Confidence -ne "Low"))
}

function Prompt-Genre {
    param(
        [string]$FolderName,
        [string]$Category,
        [string]$SuggestedGenre
    )

    $genres = if ($Category -eq $MusicFolderName) { $MusicGenreNames } else { $BookGenreNames }

    while ($true) {
        Write-Host ""
        Write-Info "$Category genre folder"
        Write-Host "Folder: $FolderName"
        Write-Host "Suggested genre: $SuggestedGenre"
        Write-Host "Choose a genre folder number, press Enter for suggested, type a custom genre name, or type Q to quit."

        for ($i = 0; $i -lt $genres.Count; $i++) {
            Write-Host ("  {0} = {1}" -f ($i + 1), $genres[$i])
        }

        $choice = Read-LineExact "Genre: "
        if ($null -eq $choice) {
            $choice = ""
        }

        $clean = $choice.Trim()
        if (Test-CancelInput $clean) {
            Stop-MediaPrepProcess
        }

        if ($clean -eq "") {
            return (Remove-InvalidFolderCharacters $SuggestedGenre)
        }

        $index = 0
        if ([int]::TryParse($clean, [ref]$index)) {
            if ($index -ge 1 -and $index -le $genres.Count) {
                return $genres[$index - 1]
            }
        }

        $customGenre = Remove-InvalidFolderCharacters $clean
        if ([string]::IsNullOrWhiteSpace($customGenre)) {
            Write-Warn "That genre name is empty after removing invalid folder characters."
            continue
        }

        Write-Host "You typed genre: $customGenre"
        if (Prompt-YesNo -Question "Use this exact genre folder?" -DefaultYes $true) {
            return $customGenre
        }
    }
}

function Prompt-YesNo {
    param(
        [string]$Question,
        [bool]$DefaultYes = $true
    )

    $suffix = if ($DefaultYes) { " [Y/n/q] " } else { " [y/N/q] " }
    while ($true) {
        $answer = Read-LineExact "$Question$suffix"
        if ($null -eq $answer) {
            Stop-MediaPrepProcess
        }

        $clean = $answer.Trim().ToLowerInvariant()
        if (Test-CancelInput $clean) {
            Stop-MediaPrepProcess
        }

        if ($clean -eq "") {
            return $DefaultYes
        }
        if ($clean -eq "y" -or $clean -eq "yes") {
            return $true
        }
        if ($clean -eq "n" -or $clean -eq "no") {
            return $false
        }

        Write-Warn "Please answer y, n, or q to quit."
    }
}

function Prompt-Root {
    param([string]$InitialRoot)

    $candidate = $InitialRoot
    while ($true) {
        if ([string]::IsNullOrWhiteSpace($candidate)) {
            $candidate = Read-LineExact "Where is your media folder? "
        }

        if (Test-CancelInput $candidate) {
            Stop-MediaPrepProcess
        }

        $candidate = $candidate.Trim().Trim('"')
        if ([string]::IsNullOrWhiteSpace($candidate)) {
            Write-Warn "Please paste a media folder path, or type Q to quit cleanly."
            $candidate = ""
            continue
        }

        if (Test-Path -LiteralPath $candidate -PathType Container) {
            return ([System.IO.Path]::GetFullPath($candidate).TrimEnd("\"))
        }

        Write-Warn "That folder was not found or is not accessible: $candidate"
        $candidate = ""
    }
}

function Prompt-Placement {
    param(
        [string]$FolderName,
        [string]$SuggestedCategory,
        [string]$Reason
    )

    while ($true) {
        Write-Host ""
        Write-Warn "Manual folder placement review"
        Write-Host "Folder: $FolderName"
        if ($DestinationFolderNames -contains $SuggestedCategory) {
            Write-Host "Suggested category: $SuggestedCategory"
        }
        Write-Host "Reason: $Reason"
        Write-Host "Choose a destination:"
        Write-Host "  1 = MOVIES"
        Write-Host "  2 = TV SHOWS"
        Write-Host "  3 = MUSIC"
        Write-Host "  4 = BOOKS"
        Write-Host "  s = skip this folder"
        Write-Host "  q = quit cleanly"
        $choice = Read-LineExact "Placement: "
        if ($null -eq $choice) {
            Stop-MediaPrepProcess
        }

        $clean = $choice.Trim().ToLowerInvariant()
        if (Test-CancelInput $clean) {
            Stop-MediaPrepProcess
        }

        if ($clean -eq "1" -or $clean -eq "m" -or $clean -eq "movie" -or $clean -eq "movies") {
            return $MoviesFolderName
        }
        if ($clean -eq "2" -or $clean -eq "t" -or $clean -eq "tv" -or $clean -eq "show" -or $clean -eq "shows" -or $clean -eq "tv shows") {
            return $ShowsFolderName
        }
        if ($clean -eq "3" -or $clean -eq "u" -or $clean -eq "music") {
            return $MusicFolderName
        }
        if ($clean -eq "4" -or $clean -eq "b" -or $clean -eq "book" -or $clean -eq "books") {
            return $BooksFolderName
        }
        if ($clean -eq "" -or $clean -eq "s" -or $clean -eq "skip") {
            return $ManualReview
        }
        Write-Warn "Please enter 1, 2, 3, 4, s, or q."
    }
}

function Prompt-NameReview {
    param(
        [string]$CurrentName,
        [string]$ProposedName,
        [bool]$ForcePrompt
    )

    if (-not $ForcePrompt -and ($CurrentName -ceq $ProposedName)) {
        return [pscustomobject]@{ Name = $ProposedName; Skip = $false; Source = "name unchanged" }
    }

    while ($true) {
        Write-Host ""
        Write-Info "Folder name review"
        Write-Host "Current name:  $CurrentName"
        Write-Host "Proposed name: $ProposedName"
        Write-Host "Press Enter to accept, type KEEP to keep current, SKIP to skip this folder, CANCEL to quit, or type the full correct name."
        $entry = Read-LineExact "Correct folder name: "
        if ($null -eq $entry) {
            Stop-MediaPrepProcess
        }

        $clean = $entry.Trim()
        $lower = $clean.ToLowerInvariant()
        if (Test-CancelInput $lower) {
            Stop-MediaPrepProcess
        }

        if ($clean -eq "") {
            return [pscustomobject]@{ Name = $ProposedName; Skip = $false; Source = "accepted proposed name" }
        }
        if ($lower -eq "keep" -or $lower -eq "k") {
            return [pscustomobject]@{ Name = $CurrentName; Skip = $false; Source = "kept current name" }
        }
        if ($lower -eq "skip" -or $lower -eq "s") {
            return [pscustomobject]@{ Name = $CurrentName; Skip = $true; Source = "skipped during name review" }
        }

        $candidate = Remove-InvalidFolderCharacters $clean
        if ([string]::IsNullOrWhiteSpace($candidate)) {
            Write-Warn "That name is empty after removing invalid folder characters."
            continue
        }

        Write-Host ""
        Write-Host "You typed: $candidate"
        if (Prompt-YesNo -Question "Use this exact top-level folder name?" -DefaultYes $true) {
            return [pscustomobject]@{ Name = $candidate; Skip = $false; Source = "custom name typed and confirmed" }
        }
    }
}

function Resolve-ReportFolder {
    param([string]$RequestedReportFolder)

    if (-not [string]::IsNullOrWhiteSpace($RequestedReportFolder)) {
        return ([System.IO.Path]::GetFullPath($RequestedReportFolder).TrimEnd("\"))
    }

    return $PSScriptRoot
}

function Get-ImmediateChildFolders {
    param([string]$RootPath)

    # Safety rule: this is intentionally non-recursive. Do not replace with recursion.
    return @(Get-ChildItem -LiteralPath $RootPath -Directory -Force)
}

function New-IgnoredPlanItem {
    param(
        [string]$Name,
        [string]$FullPath,
        [string]$Reason,
        [string]$Decision
    )

    return [pscustomobject]@{
        CurrentFolderName = $Name
        CurrentFullPath = $FullPath
        SourceParent = ""
        ScanScope = "ignored"
        ProposedCategory = $Ignored
        ProposedGenre = ""
        ProposedPlexName = $Name
        ProposedFinalPath = $FullPath
        Confidence = "High"
        Reason = $Reason
        OnlineSource = ""
        OnlineMatch = ""
        WouldBeMoved = $false
        WouldBeRenamed = $false
        WarningOrConflict = ""
        Conflict = $false
        IgnoredDestination = $true
        UserDecision = $Decision
    }
}

function New-ScanTarget {
    param(
        [System.IO.DirectoryInfo]$Folder,
        [string]$DefaultCategory = "",
        [string]$ExistingGenre = "",
        [string]$Scope = "Root"
    )

    return [pscustomobject]@{
        Name = $Folder.Name
        FullPath = $Folder.FullName
        SourceParent = $Folder.Parent.FullName.TrimEnd("\")
        DefaultCategory = $DefaultCategory
        ExistingGenre = $ExistingGenre
        Scope = $Scope
    }
}

function Add-FullCheckupTargets {
    param(
        [object[]]$Targets,
        [string]$CategoryPath,
        [string]$Category,
        [string[]]$GenreNames = @()
    )

    if (-not (Test-Path -LiteralPath $CategoryPath -PathType Container)) {
        return @($Targets)
    }

    foreach ($folder in @(Get-ImmediateChildFolders -RootPath $CategoryPath)) {
        if ($GenreNames.Count -gt 0 -and ($GenreNames -contains $folder.Name)) {
            foreach ($genreChild in @(Get-ImmediateChildFolders -RootPath $folder.FullName)) {
                $Targets += (New-ScanTarget -Folder $genreChild -DefaultCategory $Category -ExistingGenre $folder.Name -Scope "$Category genre: $($folder.Name)")
            }
        }
        else {
            $Targets += (New-ScanTarget -Folder $folder -DefaultCategory $Category -ExistingGenre "" -Scope $Category)
        }
    }

    return @($Targets)
}

function Get-PlanScanTargets {
    param(
        [string]$RootPath,
        [hashtable]$DestinationPaths,
        [string]$ScriptRootFull,
        [bool]$UseFullCheckup
    )

    $targets = @()
    $ignored = @()

    foreach ($child in @(Get-ImmediateChildFolders -RootPath $RootPath)) {
        $childFull = [System.IO.Path]::GetFullPath($child.FullName).TrimEnd("\")

        if ($DestinationFolderNames -contains $child.Name) {
            $ignored += (New-IgnoredPlanItem -Name $child.Name -FullPath $child.FullName -Reason "Ignored destination folder." -Decision "ignored destination")
            continue
        }

        if ($childFull -eq $ScriptRootFull) {
            $ignored += (New-IgnoredPlanItem -Name $child.Name -FullPath $child.FullName -Reason "Ignored Media-Prep-N-Sort tool folder." -Decision "ignored tool folder")
            continue
        }

        $targets += (New-ScanTarget -Folder $child -Scope "Root")
    }

    if ($UseFullCheckup) {
        $targets = Add-FullCheckupTargets -Targets $targets -CategoryPath $DestinationPaths[$MoviesFolderName] -Category $MoviesFolderName
        $targets = Add-FullCheckupTargets -Targets $targets -CategoryPath $DestinationPaths[$ShowsFolderName] -Category $ShowsFolderName
        $targets = Add-FullCheckupTargets -Targets $targets -CategoryPath $DestinationPaths[$MusicFolderName] -Category $MusicFolderName -GenreNames $MusicGenreNames
        $targets = Add-FullCheckupTargets -Targets $targets -CategoryPath $DestinationPaths[$BooksFolderName] -Category $BooksFolderName -GenreNames $BookGenreNames
    }

    $dedupedTargets = @()
    $seen = @{}
    foreach ($target in $targets) {
        $targetFull = [System.IO.Path]::GetFullPath($target.FullPath).TrimEnd("\")
        if ($targetFull -eq $ScriptRootFull) {
            continue
        }
        if ($seen.ContainsKey($targetFull.ToLowerInvariant())) {
            continue
        }

        $seen[$targetFull.ToLowerInvariant()] = $true
        $dedupedTargets += $target
    }

    return [pscustomobject]@{
        Targets = @($dedupedTargets)
        Ignored = @($ignored)
    }
}

function New-Plan {
    param(
        [string]$RootPath,
        [bool]$ReviewNames,
        [bool]$ReviewMediumConfidence,
        [bool]$IsNonInteractive,
        [bool]$UseOnlineLookup,
        [bool]$UseFullCheckup
    )

    $moviesPath = Join-Path -Path $RootPath -ChildPath $MoviesFolderName
    $showsPath = Join-Path -Path $RootPath -ChildPath $ShowsFolderName
    $musicPath = Join-Path -Path $RootPath -ChildPath $MusicFolderName
    $booksPath = Join-Path -Path $RootPath -ChildPath $BooksFolderName
    $destinationPaths = @{
        $MoviesFolderName = $moviesPath
        $ShowsFolderName = $showsPath
        $MusicFolderName = $musicPath
        $BooksFolderName = $booksPath
    }
    $scriptRootFull = [System.IO.Path]::GetFullPath($PSScriptRoot).TrimEnd("\")
    $scan = Get-PlanScanTargets -RootPath $RootPath -DestinationPaths $destinationPaths -ScriptRootFull $scriptRootFull -UseFullCheckup $UseFullCheckup
    $items = @()

    $items += @($scan.Ignored)

    foreach ($target in $scan.Targets) {
        $name = $target.Name
        $fullPath = $target.FullPath
        $sourceParent = $target.SourceParent
        $scanScope = $target.Scope

        $classification = Get-Classification -FolderName $name
        $category = $classification.Category
        $confidence = $classification.Confidence
        $reason = $classification.Reason
        $needsReview = [bool]$classification.NeedsPlacementReview
        $decision = "automatic"
        $onlineName = ""
        $onlineGenre = ""
        $onlineSource = ""
        $onlineMatch = ""

        if ($DestinationFolderNames -contains $target.DefaultCategory) {
            $defaultCategory = $target.DefaultCategory
            if (($DestinationFolderNames -contains $classification.Category) -and
                ($classification.Confidence -eq "High") -and
                ($classification.Category -ne $defaultCategory)) {
                $category = $classification.Category
                $confidence = "Low"
                $reason = "Full checkup found a mismatch: folder is currently under $defaultCategory, but the name suggests $($classification.Category)."
                $needsReview = $true
                $decision = "full checkup mismatch"
            }
            else {
                $category = $defaultCategory
                $confidence = "Existing"
                $reason = "Full checkup is using the existing $defaultCategory location as the safe default."
                $needsReview = $false
                $decision = "existing location"
            }
        }

        if ($confidence -eq "Medium" -and -not $ReviewMediumConfidence) {
            $needsReview = $false
        }

        if ($UseOnlineLookup) {
            $onlineSuggestion = Get-OnlineMetadataSuggestion -FolderName $name -LocalCategory $category
            if ($null -ne $onlineSuggestion -and ($DestinationFolderNames -contains $onlineSuggestion.Category)) {
                $useOnlineSuggestion = $false
                if ($IsNonInteractive) {
                    $useOnlineSuggestion = ($onlineSuggestion.Confidence -eq "High")
                }
                else {
                    $useOnlineSuggestion = Confirm-OnlineSuggestion -FolderName $name -Suggestion $onlineSuggestion
                }

                if ($useOnlineSuggestion) {
                    $category = $onlineSuggestion.Category
                    $confidence = "Online $($onlineSuggestion.Confidence)"
                    $reason = "Online lookup accepted from $($onlineSuggestion.Source)."
                    $needsReview = $false
                    $decision = "online lookup accepted"
                    $onlineName = $onlineSuggestion.CleanName
                    $onlineGenre = $onlineSuggestion.Genre
                    $onlineSource = $onlineSuggestion.Source
                    $onlineMatch = $onlineSuggestion.MatchName
                }
            }
        }

        if ($needsReview -and -not $IsNonInteractive) {
            $choice = Prompt-Placement -FolderName $name -SuggestedCategory $category -Reason $reason
            if ($DestinationFolderNames -contains $choice) {
                $category = $choice
                $confidence = "User"
                $reason = "Category selected by user."
                $decision = "manual placement"
            }
            else {
                $category = $ManualReview
                $decision = "skipped placement review"
            }
        }

        if ($category -eq $ManualReview) {
            $items += [pscustomobject]@{
                CurrentFolderName = $name
                CurrentFullPath = $fullPath
                SourceParent = $sourceParent
                ScanScope = $scanScope
                ProposedCategory = $ManualReview
                ProposedGenre = ""
                ProposedPlexName = Get-ProposedName -FolderName $name -Category $ManualReview
                ProposedFinalPath = "N/A - skipped."
                Confidence = $confidence
                Reason = $reason
                OnlineSource = $onlineSource
                OnlineMatch = $onlineMatch
                WouldBeMoved = $false
                WouldBeRenamed = $false
                WarningOrConflict = "Skipped for manual review; no move will be performed."
                Conflict = $false
                IgnoredDestination = $false
                UserDecision = $decision
            }
            continue
        }

        if (-not [string]::IsNullOrWhiteSpace($onlineName)) {
            $proposedName = Remove-InvalidFolderCharacters $onlineName
        }
        else {
            $proposedName = Remove-InvalidFolderCharacters (Get-ProposedName -FolderName $name -Category $category)
        }

        if ([string]::IsNullOrWhiteSpace($proposedName)) {
            $proposedName = $name
        }

        $nameDecisionSource = if ($decision -eq "online lookup accepted") { "online lookup accepted" } else { "automatic proposed name" }
        $nameDecision = [pscustomobject]@{ Name = $proposedName; Skip = $false; Source = $nameDecisionSource }
        $forceNamePrompt = ($decision -eq "manual placement")
        if ($ReviewNames -and -not $IsNonInteractive) {
            $nameDecision = Prompt-NameReview -CurrentName $name -ProposedName $proposedName -ForcePrompt $forceNamePrompt
            $proposedName = $nameDecision.Name
        }

        if ($nameDecision.Skip) {
            $items += [pscustomobject]@{
                CurrentFolderName = $name
                CurrentFullPath = $fullPath
                SourceParent = $sourceParent
                ScanScope = $scanScope
                ProposedCategory = $category
                ProposedGenre = ""
                ProposedPlexName = $proposedName
                ProposedFinalPath = "N/A - skipped by user."
                Confidence = $confidence
                Reason = $reason
                OnlineSource = $onlineSource
                OnlineMatch = $onlineMatch
                WouldBeMoved = $false
                WouldBeRenamed = $false
                WarningOrConflict = "Skipped by user during name review."
                Conflict = $false
                IgnoredDestination = $false
                UserDecision = $nameDecision.Source
            }
            continue
        }

        $genreName = ""
        $targetBase = $destinationPaths[$category]
        if ($category -eq $MusicFolderName -or $category -eq $BooksFolderName) {
            if (-not [string]::IsNullOrWhiteSpace($onlineGenre) -and $onlineGenre -ne "Other") {
                $genreName = $onlineGenre
            }
            elseif (-not [string]::IsNullOrWhiteSpace($target.ExistingGenre)) {
                $genreName = $target.ExistingGenre
            }
            else {
                $genreName = Get-GenreSuggestion -FolderName $name -Category $category
            }

            if (-not $IsNonInteractive) {
                $genreName = Prompt-Genre -FolderName $name -Category $category -SuggestedGenre $genreName
            }

            if ([string]::IsNullOrWhiteSpace($genreName)) {
                $genreName = "Other"
            }

            $targetBase = Join-Path -Path $targetBase -ChildPath (Remove-InvalidFolderCharacters $genreName)
        }

        $finalPath = Join-Path -Path $targetBase -ChildPath $proposedName
        $currentFullForPlan = [System.IO.Path]::GetFullPath($fullPath).TrimEnd("\")
        $finalFullForPlan = [System.IO.Path]::GetFullPath($finalPath).TrimEnd("\")
        $wouldBeMoved = ($currentFullForPlan -ne $finalFullForPlan)
        $wouldBeRenamed = ((Split-Path -Path $fullPath -Leaf) -cne (Split-Path -Path $finalPath -Leaf))
        $initialWarning = if ($wouldBeMoved) { "" } else { "Already in the planned location with the planned name." }

        $items += [pscustomobject]@{
            CurrentFolderName = $name
            CurrentFullPath = $fullPath
            SourceParent = $sourceParent
            ScanScope = $scanScope
            ProposedCategory = $category
            ProposedGenre = $genreName
            ProposedPlexName = $proposedName
            ProposedFinalPath = $finalPath
            Confidence = $confidence
            Reason = $reason
            OnlineSource = $onlineSource
            OnlineMatch = $onlineMatch
            WouldBeMoved = $wouldBeMoved
            WouldBeRenamed = $wouldBeRenamed
            WarningOrConflict = $initialWarning
            Conflict = $false
            IgnoredDestination = $false
            UserDecision = $nameDecision.Source
        }
    }

    $targetMap = @{}
    foreach ($item in $items) {
        if (-not $item.WouldBeMoved) {
            continue
        }

        $warnings = @()
        if (Test-Path -LiteralPath $item.ProposedFinalPath) {
            $item.Conflict = $true
            $warnings += "Proposed destination already exists; item will be skipped."
        }

        $key = $item.ProposedFinalPath.ToLowerInvariant()
        if ($targetMap.ContainsKey($key)) {
            $item.Conflict = $true
            $warnings += "Another planned item has the same proposed final path; item will be skipped."
        }
        else {
            $targetMap[$key] = $item.CurrentFolderName
        }

        $item.WarningOrConflict = ($warnings -join " ")
    }

    $scanMode = if ($UseFullCheckup) { "Full checkup" } else { "Default" }
    $safetyText = if ($UseFullCheckup) {
        "Full checkup mode. Scans unsorted root folders plus folders already inside MOVIES, TV SHOWS, MUSIC, BOOKS, and known MUSIC/BOOKS genre folders. No media files are inspected."
    }
    else {
        "Default mode. Immediate child folders only. Destination folders are ignored. No recursive scan. No nested files or folders inspected."
    }

    return [pscustomobject]@{
        GeneratedAt = (Get-Date).ToString("o")
        Version = $ScriptVersion
        Root = $RootPath
        ScanMode = $scanMode
        MoviesDestination = $moviesPath
        TvShowsDestination = $showsPath
        MusicDestination = $musicPath
        BooksDestination = $booksPath
        Safety = $safetyText
        Items = @($items)
    }
}

function Get-Summary {
    param([object[]]$Items)

    return [pscustomobject]@{
        TotalImmediateFolders = $Items.Count
        IgnoredDestinations = @($Items | Where-Object { $_.IgnoredDestination }).Count
        Movies = @($Items | Where-Object { $_.ProposedCategory -eq $MoviesFolderName }).Count
        TvShows = @($Items | Where-Object { $_.ProposedCategory -eq $ShowsFolderName }).Count
        Music = @($Items | Where-Object { $_.ProposedCategory -eq $MusicFolderName }).Count
        Books = @($Items | Where-Object { $_.ProposedCategory -eq $BooksFolderName }).Count
        ManualReviewSkipped = @($Items | Where-Object { $_.ProposedCategory -eq $ManualReview }).Count
        Conflicts = @($Items | Where-Object { $_.Conflict }).Count
        WillMove = @($Items | Where-Object { $_.WouldBeMoved -and -not $_.Conflict }).Count
        WillRename = @($Items | Where-Object { $_.WouldBeRenamed -and -not $_.Conflict }).Count
    }
}

function Write-PlanReports {
    param(
        [object]$Plan,
        [object]$Summary,
        [string]$OutputFolder
    )

    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

    $textPath = Join-Path -Path $OutputFolder -ChildPath "media_prep_n_sort_plan_$RunStamp.txt"
    $jsonPath = Join-Path -Path $OutputFolder -ChildPath "media_prep_n_sort_plan_$RunStamp.json"
    $csvPath = Join-Path -Path $OutputFolder -ChildPath "media_prep_n_sort_plan_$RunStamp.csv"

    $lines = @()
    $lines += "MEDIA-PREP-N-SORT PLAN"
    $lines += "Generated: $($Plan.GeneratedAt)"
    $lines += "Version: $($Plan.Version)"
    $lines += "Root: $($Plan.Root)"
    $lines += "Scan mode: $($Plan.ScanMode)"
    $lines += "Safety: $($Plan.Safety)"
    $lines += ""
    $lines += "No folders have been moved yet unless the final MOVE and CONFIRM confirmations were entered."
    $lines += ""
    $lines += "SUMMARY"
    $lines += "Total immediate folders: $($Summary.TotalImmediateFolders)"
    $lines += "Ignored destination folders: $($Summary.IgnoredDestinations)"
    $lines += "Movies: $($Summary.Movies)"
    $lines += "TV shows: $($Summary.TvShows)"
    $lines += "Music: $($Summary.Music)"
    $lines += "Books: $($Summary.Books)"
    $lines += "Manual review skipped: $($Summary.ManualReviewSkipped)"
    $lines += "Conflicts: $($Summary.Conflicts)"
    $lines += "Will move: $($Summary.WillMove)"
    $lines += "Will rename: $($Summary.WillRename)"
    $lines += ""
    $lines += "ITEMS"

    foreach ($item in $Plan.Items) {
        $lines += ""
        $lines += "Current folder name: $($item.CurrentFolderName)"
        $lines += "Current full path: $($item.CurrentFullPath)"
        $lines += "Scan scope: $($item.ScanScope)"
        $lines += "Proposed category: $($item.ProposedCategory)"
        $lines += "Proposed genre: $($item.ProposedGenre)"
        $lines += "Proposed clean top-level folder name: $($item.ProposedPlexName)"
        $lines += "Proposed final path: $($item.ProposedFinalPath)"
        $lines += "Confidence level: $($item.Confidence)"
        $lines += "Reason: $($item.Reason)"
        $lines += "Online source: $($item.OnlineSource)"
        $lines += "Online match: $($item.OnlineMatch)"
        $lines += "Would be moved: $($item.WouldBeMoved)"
        $lines += "Would be renamed: $($item.WouldBeRenamed)"
        $lines += "User decision: $($item.UserDecision)"
        $lines += "Warning or conflict: $($item.WarningOrConflict)"
    }

    $lines | Set-Content -LiteralPath $textPath -Encoding UTF8
    $Plan | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
    $Plan.Items | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8

    return [pscustomobject]@{
        TextReport = $textPath
        JsonPlan = $jsonPath
        CsvPlan = $csvPath
    }
}

function Invoke-MovePlan {
    param(
        [object]$Plan,
        [string]$RootPath,
        [string]$OutputFolder,
        [bool]$ShouldSaveReports
    )

    New-Item -ItemType Directory -Path $Plan.MoviesDestination -Force | Out-Null
    New-Item -ItemType Directory -Path $Plan.TvShowsDestination -Force | Out-Null
    New-Item -ItemType Directory -Path $Plan.MusicDestination -Force | Out-Null
    New-Item -ItemType Directory -Path $Plan.BooksDestination -Force | Out-Null

    foreach ($genre in $MusicGenreNames) {
        New-Item -ItemType Directory -Path (Join-Path -Path $Plan.MusicDestination -ChildPath $genre) -Force | Out-Null
    }

    foreach ($genre in $BookGenreNames) {
        New-Item -ItemType Directory -Path (Join-Path -Path $Plan.BooksDestination -ChildPath $genre) -Force | Out-Null
    }

    $results = @()
    $rootFull = [System.IO.Path]::GetFullPath($RootPath).TrimEnd("\")

    foreach ($item in $Plan.Items) {
        $status = "Skipped"
        $message = ""

        if ($item.IgnoredDestination) {
            $message = "Ignored destination folder."
        }
        elseif (-not $item.WouldBeMoved) {
            $message = "No move planned."
        }
        elseif ($item.Conflict) {
            $message = "Conflict from plan; skipped."
        }
        else {
            try {
                if (-not (Test-Path -LiteralPath $item.CurrentFullPath -PathType Container)) {
                    throw "Source folder no longer exists."
                }
                if (Test-Path -LiteralPath $item.ProposedFinalPath) {
                    throw "Destination already exists."
                }

                $sourceParent = [System.IO.Directory]::GetParent($item.CurrentFullPath).FullName.TrimEnd("\")
                $allowedSourceParent = if (-not [string]::IsNullOrWhiteSpace($item.SourceParent)) {
                    [System.IO.Path]::GetFullPath($item.SourceParent).TrimEnd("\")
                }
                else {
                    $rootFull
                }
                if ($sourceParent -ne $allowedSourceParent) {
                    throw "Source is not in the approved scan location from the plan."
                }

                $sourceFull = [System.IO.Path]::GetFullPath($item.CurrentFullPath).TrimEnd("\")
                $destFull = [System.IO.Path]::GetFullPath($item.ProposedFinalPath).TrimEnd("\")
                if ($destFull.StartsWith($sourceFull + "\", [StringComparison]::OrdinalIgnoreCase)) {
                    throw "Refusing to move a folder into itself."
                }

                $destinationParent = [System.IO.Directory]::GetParent($item.ProposedFinalPath).FullName
                New-Item -ItemType Directory -Path $destinationParent -Force | Out-Null
                Move-Item -LiteralPath $item.CurrentFullPath -Destination $item.ProposedFinalPath -ErrorAction Stop
                $status = if ($item.WouldBeRenamed) { "Moved and renamed" } else { "Moved" }
                $message = "Moved sealed top-level folder only."
            }
            catch {
                $status = "Error"
                $message = $_.Exception.Message
            }
        }

        $results += [pscustomobject]@{
            CurrentFolderName = $item.CurrentFolderName
            ProposedCategory = $item.ProposedCategory
            ProposedGenre = $item.ProposedGenre
            ProposedPlexName = $item.ProposedPlexName
            ProposedFinalPath = $item.ProposedFinalPath
            Status = $status
            Message = $message
        }
    }

    $textPath = ""
    $jsonPath = ""

    if ($ShouldSaveReports) {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
        $textPath = Join-Path -Path $OutputFolder -ChildPath "media_prep_n_sort_execution_$RunStamp.txt"
        $jsonPath = Join-Path -Path $OutputFolder -ChildPath "media_prep_n_sort_execution_$RunStamp.json"

        $lines = @()
        $lines += "MEDIA-PREP-N-SORT EXECUTION REPORT"
        $lines += "Generated: $((Get-Date).ToString("o"))"
        $lines += "Root: $RootPath"
        $lines += "Scan mode: $($Plan.ScanMode)"
        $lines += "Safety: $($Plan.Safety)"
        $lines += ""

        foreach ($result in $results) {
            $lines += "Folder: $($result.CurrentFolderName)"
            $lines += "Category: $($result.ProposedCategory)"
            $lines += "Genre: $($result.ProposedGenre)"
            $lines += "Plex name: $($result.ProposedPlexName)"
            $lines += "Final path: $($result.ProposedFinalPath)"
            $lines += "Status: $($result.Status)"
            $lines += "Message: $($result.Message)"
            $lines += ""
        }

        $lines | Set-Content -LiteralPath $textPath -Encoding UTF8
        [pscustomobject]@{
            GeneratedAt = (Get-Date).ToString("o")
            Root = $RootPath
            Results = @($results)
        } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
    }

    return [pscustomobject]@{
        TextReport = $textPath
        JsonReport = $jsonPath
        Results = @($results)
    }
}

try {
if ($NonInteractive) {
    $DryRun = $true
    $NoNameReview = $true
    $TrustMediumConfidence = $true
}

if (-not $NonInteractive -and [string]::IsNullOrWhiteSpace($Root)) {
    Show-Welcome
}

$rootPath = Prompt-Root -InitialRoot $Root
$outputFolder = Resolve-ReportFolder -RequestedReportFolder $ReportFolder

$missingDestinations = @()
foreach ($destinationName in $DestinationFolderNames) {
    $destinationPath = Join-Path -Path $rootPath -ChildPath $destinationName
    if (-not (Test-Path -LiteralPath $destinationPath -PathType Container)) {
        $missingDestinations += $destinationPath
    }
}

if ($missingDestinations.Count -gt 0) {
    if ($DryRun -or $NonInteractive) {
        Write-Warn "One or more destination folders are missing. They would be created during execution:"
        foreach ($missingDestination in $missingDestinations) {
            Write-Host "  $missingDestination"
        }
    }
    elseif (-not (Prompt-YesNo -Question "Create missing MOVIES / TV SHOWS / MUSIC / BOOKS destination folders during execution?" -DefaultYes $true)) {
        Stop-MediaPrepProcess "Cancelled. Missing destination folders were not created and nothing was moved or renamed."
    }
}

$reviewMedium = (-not $TrustMediumConfidence)
$reviewNames = $false
if (-not $NoNameReview -and -not $NonInteractive) {
    $reviewNames = Prompt-YesNo -Question "Review and fix proposed folder names before moving?" -DefaultYes $true
}

$plan = New-Plan -RootPath $rootPath -ReviewNames $reviewNames -ReviewMediumConfidence $reviewMedium -IsNonInteractive ([bool]$NonInteractive) -UseOnlineLookup ([bool]$OnlineLookup) -UseFullCheckup ([bool]$FullCheckup)
$summary = Get-Summary -Items $plan.Items
$reports = $null
if ($SaveReports) {
    $reports = Write-PlanReports -Plan $plan -Summary $summary -OutputFolder $outputFolder
}

Write-Host ""
Write-Good "Plan complete."
if ($OnlineLookup) {
    Write-Host "Online lookup was enabled. Any accepted metadata suggestions are included in the plan."
}
if ($FullCheckup) {
    Write-Host "Full checkup was enabled. Existing category folders and known MUSIC/BOOKS genre folders were included in the scan."
}
if ($SaveReports) {
    Write-Host "Plan report: $($reports.TextReport)"
    Write-Host "JSON plan:   $($reports.JsonPlan)"
    Write-Host "CSV plan:    $($reports.CsvPlan)"
}
else {
    Write-Host "Reports are off by default. Add -SaveReports if you want plan and execution logs saved beside the script."
}
Write-Host ""
Write-Host "Summary:"
Write-Host "  Total checked folders: $($summary.TotalImmediateFolders)"
Write-Host "  Ignored destination folders: $($summary.IgnoredDestinations)"
Write-Host "  Movies: $($summary.Movies)"
Write-Host "  TV shows: $($summary.TvShows)"
Write-Host "  Music: $($summary.Music)"
Write-Host "  Books: $($summary.Books)"
Write-Host "  Manual review skipped: $($summary.ManualReviewSkipped)"
Write-Host "  Conflicts: $($summary.Conflicts)"
Write-Host "  Will move: $($summary.WillMove)"
Write-Host "  Will rename: $($summary.WillRename)"

if ($DryRun -or $NonInteractive) {
    Write-Host ""
    Write-Good "Dry run only. Nothing was moved or renamed."
    exit 0
}

if ($summary.Conflicts -gt 0) {
    Write-Host ""
    Write-Warn "Stop: conflicts were found. Review the report before running again."
    exit 1
}

if ($summary.WillMove -eq 0) {
    Write-Host ""
    Write-Good "Nothing to move."
    exit 0
}

Write-Host ""
Write-Warn "Final safety check"
if ($FullCheckup) {
    Write-Host "Only folders included in the approved full-checkup scan will be moved:"
}
else {
    Write-Host "Only immediate child folders of this root will be moved:"
}
Write-Host "  $rootPath"
if ($FullCheckup) {
    Write-Host "Full checkup may also move or rename folders already inside MOVIES, TV SHOWS, MUSIC, BOOKS, or known genre folders."
}
Write-Host "Music/book folders may be placed one level deeper into genre folders."
Write-Host "Missing category folders and standard MUSIC/BOOKS genre folders will be created during execution."
Write-Host "Nothing inside any media, music, or book folder will be scanned, renamed, deleted, merged, or reorganized."
Write-Warn "After the final confirmation, folder moves and renames begin. There is no built-in undo."
Write-Host "Step 1: type MOVE in all caps to arm the execution step, or press Enter to stop without changing anything."
while ($true) {
    $confirm = Read-LineExact "First confirmation: "
    if ($null -eq $confirm -or $confirm.Trim() -eq "") {
        Write-Host ""
        Write-Good "Stopped. Nothing was moved or renamed."
        exit 0
    }
    if ($confirm -ceq "MOVE") {
        break
    }
    if ($confirm.Trim().ToUpperInvariant() -eq "MOVE") {
        Write-Warn "You must use all caps: MOVE"
        continue
    }

    Write-Host ""
    Write-Good "Stopped. Nothing was moved or renamed."
    exit 0
}

Write-Host ""
Write-Warn "Last chance before changes begin."
Write-Host "Type CONFIRM in all caps to permanently apply the move/rename plan now."
Write-Host "Anything else stops safely without changing anything."
while ($true) {
    $finalConfirm = Read-LineExact "Final confirmation: "
    if ($null -eq $finalConfirm -or $finalConfirm.Trim() -eq "") {
        Write-Host ""
        Write-Good "Stopped. Nothing was moved or renamed."
        exit 0
    }
    if ($finalConfirm -ceq "CONFIRM") {
        break
    }
    if ($finalConfirm.Trim().ToUpperInvariant() -eq "CONFIRM") {
        Write-Warn "You must use all caps: CONFIRM"
        continue
    }

    Write-Host ""
    Write-Good "Stopped. Nothing was moved or renamed."
    exit 0
}

$execution = Invoke-MovePlan -Plan $plan -RootPath $rootPath -OutputFolder $outputFolder -ShouldSaveReports ([bool]$SaveReports)
Write-Host ""
Write-Good "Execution complete."
if ($SaveReports) {
    Write-Host "Execution report: $($execution.TextReport)"
    Write-Host "Execution JSON:   $($execution.JsonReport)"
}
else {
    Write-Host "No report files were created. Add -SaveReports next time if you want logs."
}
}
catch [System.OperationCanceledException] {
    Write-Host ""
    Write-Good $_.Exception.Message
    exit 0
}
