<# Some bookkeeping #>
$progressPreference = 'silentlyContinue' # For Invoke-Webrequest not to bother the user (similar to wget -q)
$ErrorActionPreference = "silentlycontinue" # For errors with Invoke-Webrequest
$server = "https://video.technion.ac.il" # Technion video server
$defaultDir = "."

<# Functions #>
function main {
    printIntro # Greetings
    testMSDL # Check dependency
    $params = getCred # Get user credentials
    $Course = getCourse # Get Course address
    $links = getLinks $Course $params # Get links
    $num = ($links | Measure-Object).count # Get amount of links
    $range = getRange $num # Get range of downloads
    $links = $links | Select-Object -Skip ($range[0] - 1) | Select-Object -SkipLast ($num - $range[1]) # Trim links list
    # $links = $links[($start-1)..($end-1)] ## for PSVersion < 5.0
    $path = getDir # Get download directory

    # Let's rock!
    foreach ($line in $links) {
        $line = $line.Line.trimstart("@{href=").trim("}")
        $filename = ($line.split("/")[5]).Trim()
        $location = Join-Path -Path "$path" -ChildPath "$filename"
        $file2 = Invoke-WebRequest -uri "$server$line" -Method POST -Body $params
        $address = ($file2.Content.Split("`n") | Select-String -Pattern "window.location=").Line.split("`"")[1]
        cmd /c .\msdl.exe -s2 $address -o $location '2>&1' | parseMSDL
    }
}

function forceResolvePath {
    <#
    .SYNOPSIS
        Calls Resolve-Path but works for files that don't exist.
    #>
    param (
        [string] $FileName
    )
    $FileName = Resolve-Path $FileName -ErrorAction SilentlyContinue `
        -ErrorVariable _frperror
    if (-not($FileName)) {
        $FileName = $_frperror[0].TargetObject
    }
    return $FileName
}

function parseMSDL {
    Process {
        $first = 1
        $newline = 0
        if ($_.startswith("DL")) {
            Write-Host -NoNewline "`r$_"
            $first = 0
            $newline = 0
        }
        else {
            if ((-not $first) -and (-not $newline)) {
                Write-Host "" # for newline between files
                $newline = 1
            }
            Write-Host "$_"
        }
    }
}

function readDefault ([string]$prompt, $default) {
    $val = Read-Host -Prompt "$prompt (Enter for default - $default)"
    $val = ($default, $val)[[bool]$val]
    return $val
}

function getPath {
    $pathPrompt = "Please enter directory (e.g. `"C:\my_vids`")"
    $dir = readDefault $pathPrompt (forceResolvePath $defaultDir)
    while (($dir.IndexOfAny([System.IO.Path]::GetInvalidPathChars()) -ne -1)`
            -or ((test-path -Path $dir) -and (-not (test-path -Path $dir -PathType Container)))) {
        $dir = readDefault $pathPrompt (forceResolvePath $defaultDir)
    }
    return forceResolvePath $dir
}

function getCourse {
    $coursePrompt = "Please insert valid course url (e.g https://video.technion.ac.il/Courses/PhysMech.html)"
    $Course = Read-Host -Prompt $coursePrompt
    Invoke-WebRequest $Course | Out-Null
    while (-not $?) {
        $Course = Read-Host -Prompt "Invalid URL.`n$coursePrompt"
        Invoke-WebRequest $Course | Out-Null
    }
    return $Course
}

function getDir {
    do {
        $path = getPath # path name legality check
        if ((test-path -Path $path -PathType Container) -and ((Get-Item $Path).PSDrive.Provider.Name -eq "FileSystem")) {
            Write-Host "Saving to folder $path"
            $result = 0
        }
        else {
            $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Create folder $path"
            $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Don't create $path, choose a different directory"
            $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
            $caption = "directory $path doesn't exist."
            $message = "Do you want to create it?"
            $result = $Host.UI.PromptForChoice($caption, $message, $choices, 0)
            if ($result -eq 0) {
                Write-Host "Creating folder $path"
                New-Item -ItemType Directory -Force -Path $path | Out-Null
                if (-not $?) {
                    Write-Host "Error: Could not create directory $path"
                    exit
                }
            } 
        }
    } while ($result -eq 1)
    return $path
}

function getCred {
    if (-not (Test-Path ".\params.txt")) {
        # Get user input
        Write-Host "params.txt not found. It will be created from your username & password."
        $User = Read-Host -Prompt "Please insert username (e.g ran.lottem without @campus etc.)"
        $Pass = Read-Host -Prompt "Please insert password"
        Add-Content ".\params.txt" "username = $User"
        Add-Content ".\params.txt" "password = $Pass"
    }
    else {
        # param.txt exists
        Write-Host "Loading username & password from params.txt"
    }
    return Get-Content -Raw ".\params.txt" | ConvertFrom-StringData
}

function testMSDL {
    try {
        & .\msdl.exe -qh
    }
    catch {
        Write-Host "Error: msdl.exe not found in this script's folder. Exiting."
        exit
    }
}

function getLinks ($Course, $params) {
    $URL = Invoke-WebRequest $Course -Method POST -Body $params

    # Check URL
    if (($URL.links[0].innerHTML) -Match "forgot") {
        Write-Host "Error: bad username or password, please try again."
        Remove-Item -Path ".\params.txt"
        exit
    }
    return $URL.Links | Select-Object href | Select-String movies/rtsp
}

function getRange ([int]$num) {
    if ($num -le 0) {
        Write-Host "Error: no files found."
        exit
    }

    # Get file range for download
    Write-Host "There are $num files."

    Write-Host "Please choose range of videos to download. Start video:"
    [int]$start = getInRange 1 $num 1

    Write-Host "End video:"
    [int]$end = getInRange $start $num $num

    $start
    $end
    return
}

function getInRange ([int]$low = 1, [int]$high, [int]$def) {
    $prompt = "Please choose video number ($low-$high)"
    [int]$res = readDefault $prompt $def
    while (($res -lt $low) -or ($res -gt $high)) {
        Write-Host "Invalid choice."
        [int]$res = readDefault $prompt $def
    }
    return [int]$res
}

function printIntro {
    Write-Host "Technion Old Video Server Downloader v0.6 by Ran Lottem"
    Write-Host "Based on bash script by Ohad Eytan`n"
}

main # Call the function