# Media-Prep-N-Sort
# Media-Prep-N-Sort

Media-Prep-N-Sort is a safe interactive PowerShell sorter for media folders. It helps organize top-level folders into Movies, TV Shows, Music, and Books without touching the files inside them.

It is designed for Windows, mapped network drives, NAS folders, Plex-style libraries, Jellyfin-style libraries, Emby-style libraries, and general media download folders.

## What It Does

Media-Prep-N-Sort organizes only the immediate top-level folders inside the media folder you choose.

Example before sorting:

~~~text
My Media Folder
├─ Space Lantern 2024 1080p
├─ Forest Detectives Season 1
├─ Neon River Album FLAC
├─ Dragon Library EPUB
├─ MOVIES
├─ TV SHOWS
├─ MUSIC
└─ BOOKS
~~~

Example after sorting:

~~~text
My Media Folder
├─ MOVIES
│  └─ Space Lantern (2024)
├─ TV SHOWS
│  └─ Forest Detectives
├─ MUSIC
│  └─ Electronic
│     └─ Neon River Album
└─ BOOKS
   └─ Fantasy
      └─ Dragon Library
~~~

The script treats every top-level folder like a sealed box. It can move or rename the folder itself, but it does not scan, rename, move, delete, merge, or reorganize anything inside that folder.

## Most Important Safety Rule

Media-Prep-N-Sort only works on this level:

~~~text
My Media Folder\Example Movie Folder
My Media Folder\Example Show Folder
My Media Folder\Example Music Folder
My Media Folder\Example Book Folder
~~~

It does not work inside this level:

~~~text
My Media Folder\Example Show Folder\Season 1
My Media Folder\Example Show Folder\Season 1\Episode File.mkv
My Media Folder\Example Book Folder\Book File.epub
My Media Folder\Example Music Folder\Track File.flac
~~~

This is a top-level folder sorting tool, not a media-file renaming tool.

## Categories

Media-Prep-N-Sort can sort folders into:

~~~text
1 = MOVIES
2 = TV SHOWS
3 = MUSIC
4 = BOOKS
~~~

It creates or uses these destination folders:

~~~text
MOVIES
TV SHOWS
MUSIC
BOOKS
~~~

For Music and Books, it can also sort one layer deeper into genre folders.

Music genre examples:

~~~text
MUSIC\Rock
MUSIC\Pop
MUSIC\Hip-Hop
MUSIC\Electronic
MUSIC\Classical
MUSIC\Soundtracks
MUSIC\Other
~~~

Book genre examples:

~~~text
BOOKS\Fiction
BOOKS\Fantasy
BOOKS\Sci-Fi
BOOKS\Mystery-Thriller
BOOKS\Comics-Manga
BOOKS\Audiobooks
BOOKS\Other
~~~

## Interactive Manual Review

If the script is unsure where a folder belongs, it asks you.

You can choose:

~~~text
1 = MOVIES
2 = TV SHOWS
3 = MUSIC
4 = BOOKS
s = skip this folder
q = quit cleanly
~~~

This means the script does not have to guess when confidence is low.

## Folder Name Cleanup

Media-Prep-N-Sort can suggest cleaner top-level folder names.

Example movie cleanup:

~~~text
space.lantern.2024.1080p
~~~

Becomes:

~~~text
Space Lantern (2024)
~~~

Example show cleanup:

~~~text
Forest Detectives Season 1-3
~~~

Becomes:

~~~text
Forest Detectives
~~~

Example music cleanup:

~~~text
Neon River Complete Album FLAC
~~~

Becomes:

~~~text
Neon River
~~~

Example book cleanup:

~~~text
Dragon Library EPUB Collection
~~~

Becomes:

~~~text
Dragon Library
~~~

During name review, you can accept the suggestion, keep the current name, skip the folder, cancel the whole process, or type the full corrected name yourself.

## Clean Cancel Support

At normal prompts, you can cancel safely by typing:

~~~text
Q
CANCEL
EXIT
STOP
~~~

If you cancel before the final move begins, nothing is moved or renamed.

## Final Safety Confirmation

Before making changes, the script shows a summary.

Nothing moves unless you type two exact confirmations.

First:

~~~text
MOVE
~~~

Then:

~~~text
CONFIRM
~~~

After `CONFIRM`, folder moves and renames begin. There is no built-in undo after that point, so review the summary carefully.

If you type anything else at either confirmation, the script stops safely and nothing changes.

## Optional Online Lookup

By default, Media-Prep-N-Sort works offline using only folder names.

If you run it with:

~~~powershell
-OnlineLookup
~~~

it can use public metadata sources to improve suggestions.

It can look up:

~~~text
Movies through Wikidata
TV shows through TVmaze
Music through MusicBrainz
Books through Open Library
~~~

No paid API key is required.

Online lookup is only a suggestion helper. You still confirm the result before anything moves.

## Reports

Reports are off by default so the install folder stays clean.

If you want reports, run with:

~~~powershell
-SaveReports
~~~

Reports are saved in the Media-Prep-N-Sort folder by default.

Reports can include the plan, JSON data, CSV data, and execution logs.

## Install And Run With One PowerShell Command

Open PowerShell and paste this command:

~~~powershell
$ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; $choice=Read-Host 'Where do you want to install Media-Prep-N-Sort? Press Enter for Desktop'; if ([string]::IsNullOrWhiteSpace($choice)) { $choice=[Environment]::GetFolderPath('Desktop') }; $choice=$choice.Trim().Trim('"'); $d=Join-Path $choice 'Media-Prep-N-Sort'; Write-Host ""; Write-Host "Install folder: $d"; $ok=Read-Host 'Are you sure? Type Y to download and run Media-Prep-N-Sort here'; if ($ok -in @('Y','y','YES','yes')) { New-Item -ItemType Directory -Force -Path $d | Out-Null; $base='https://raw.githubusercontent.com/hd9fxt7plq24vm6kcjxq2rw58y3z-cmyk/Media-Prep-N-Sort/main'; Invoke-WebRequest -UseBasicParsing -Uri "$base/MediaPrepNSort.ps1" -OutFile (Join-Path $d 'MediaPrepNSort.ps1'); Invoke-WebRequest -UseBasicParsing -Uri "$base/RUN_MEDIA_PREP_N_SORT.txt" -OutFile (Join-Path $d 'RUN_MEDIA_PREP_N_SORT.txt'); powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $d 'MediaPrepNSort.ps1') } else { Write-Host 'Install cancelled. Nothing was downloaded.' }
~~~

The command asks where you want to install the tool. If you press Enter, it installs to your Desktop.

After download, the script starts and asks:

~~~text
Where is your media folder?
~~~

Paste the path to the folder you want sorted.

## How To Copy Your Media Folder Path

Open your media folder in File Explorer.

Click the address bar at the top of the window.

Press `Ctrl+C`.

Go back to PowerShell.

Press `Ctrl+V` or right-click to paste.

Press Enter.

Example path format:

~~~text
D:\My Media Folder
~~~

## Run It Again Later

After the first install, open PowerShell inside the Media-Prep-N-Sort folder and run:

~~~powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\MediaPrepNSort.ps1"
~~~

## Useful Commands

Run normally:

~~~powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\MediaPrepNSort.ps1"
~~~

Run with online lookup:

~~~powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\MediaPrepNSort.ps1" -OnlineLookup
~~~

Run with saved reports:

~~~powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\MediaPrepNSort.ps1" -SaveReports
~~~

Run with online lookup and saved reports:

~~~powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\MediaPrepNSort.ps1" -OnlineLookup -SaveReports
~~~

Preview only without moving or renaming anything:

~~~powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\MediaPrepNSort.ps1" -DryRun
~~~

## Flags

`-OnlineLookup`

Uses public metadata sources for better category, name, and genre suggestions.

`-SaveReports`

Saves reports and logs beside the script.

`-DryRun`

Builds the plan and shows the summary, but does not move or rename anything.

`-NoNameReview`

Skips the step where you review proposed folder names.

`-TrustMediumConfidence`

Accepts medium-confidence category guesses without asking as many questions.

`-NonInteractive`

Runs without prompts and forces safe dry-run behavior. This is mainly for testing and automation.

`-Root "PATH"`

Lets you provide the media folder path directly.

`-ReportFolder "PATH"`

Changes where reports are saved when `-SaveReports` is used.

## Uninstalling

Delete the `Media-Prep-N-Sort` folder.

That removes the script, helper text file, and any optional reports saved by the tool.

It does not delete your media folders.

## Important Warning

This tool is careful, but you should still review the summary before typing `MOVE` and `CONFIRM`.

Once `CONFIRM` is typed, folder moves and renames begin. There is no built-in undo. yet.

## Like It?

If Media-Prep-N-Sort helped you, please consider starring the repo. It helps other people find the project
