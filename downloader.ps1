<# Some bookkeeping #>
$progressPreference = 'silentlyContinue' # For Invoke-Webrequest not to bother the user (similar to wget -q)
$ErrorActionPreference = "silentlycontinue" # For errors with Invoke-Webrequest
$server = "https://video.technion.ac.il" # Technion video server
$defaultDir = "."

<# Functions #>
function main {
    printIntro # Greetings
    checkDependencies # Check necessary files
    $params = getCred # Get user credentials
    $Course = getCourse # Get Course address
    $links = getLinks $Course $params # Get links
    $num = ($links | Measure-Object).count # Get amount of links
    $range = getRange $num # Get range of downloads
    $path = getDir # Get download directory
    
    # Let's rock!
    $range[0]..$range[1] | ForEach-Object {
        $Job = get-download-link $links $path $server $params $_
        start-download $Job
    }
}
function start-download ($Job) {
    $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
    $ProcessInfo.FileName = (Resolve-Path ".\msdl.exe").Path
    $ProcessInfo.RedirectStandardError = $false
    $ProcessInfo.RedirectStandardOutput = $false
    $ProcessInfo.UseShellExecute = $false
    $ProcessInfo.Arguments = "-s2 '$($Job[0])' -o '$($Job[1])'"
    $Process = New-Object System.Diagnostics.Process
    $Process.StartInfo = $ProcessInfo
    $Process.Start() | Out-Null
    do {
        $Process.StandardError.ReadLine()
    } while (!$process.HasExited)
    $Process.WaitForExit()
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

function get-download-link ([array]$links, [string]$path, $server, $params, $index) {
    $line = $links[$($index-1)]
    $line = $line.Line.trimstart("@{href=").trim("}")
    $filename = ($line.split("/")[5]).Trim()
    $filename = (Get-Culture).textinfo.ToTitleCase((($filename.ToLower() -replace "%20", " ").split("."))[0]) + "." + $filename.split(".")[1].ToLower()
    $filename = $filename -replace "-", " "
    $location = Join-Path -Path "$path" -ChildPath "$filename"
    $file = Invoke-WebRequest -uri "$server$line" -Method POST -Body $params
    $address = ($file.Content.Split("`n") | Select-String -Pattern "window.location=").Line.split("`"")[1]
    @($address,$location)
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
        # params.txt doesn't exist, get user input
        Write-Host "params.txt not found. It will be created from your username & password."
        $User = Read-Host -Prompt "Please insert username (e.g ran.lottem without @campus etc.)"
        $Pass = Read-Host -assecurestring -Prompt "Please insert password" | ConvertFrom-SecureString
        Add-Content ".\params.txt" "username = $User"
        Add-Content ".\params.txt" "password = $Pass"
    }
    else {
        # params.txt exists
        Write-Host "Loading username & password from params.txt`n"
    }
    $hash = Get-Content -Raw ".\params.txt" | ConvertFrom-StringData
    $securePass = ConvertTo-SecureString -String $hash["password"]
    $hash["password"] = getRawPassword $securePass
    return $hash
}

function getRawPassword ([SecureString] $SecurePassword) {
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
    return $pass
}

function checkDependencies {
    testMSDL
    testCygwin
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

function testCygwin {
    if (-not(Test-Path cygwin1.dll)) {
        Write-Host "Error: cygwin1.dll not found. Exiting."
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
    Write-Host "Technion Old Video Server Downloader v0.8 by Ran Lottem"
    Write-Host "Based on bash script by Ohad Eytan`n"
}

main # Call the function