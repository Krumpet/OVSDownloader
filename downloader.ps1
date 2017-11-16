<# Some bookkeeping #>
$progressPreference = 'silentlyContinue' # For Invoke-Webrequest not to bother the user (similar to wget -q)
$ErrorActionPreference = "silentlycontinue" # For errors with Invoke-Webrequest
$server = "https://video.technion.ac.il" # Technion video server
$defaultDir = "."
$version = "0.8.2"

$exitModes = @{
    "Success"       = @(0, "All files complete");
    "DirError"      = @(1, "Could not create directory");
    "CredError"     = @(2, "Bad username or password");
    "MSDLError"     = @(3, "msdl.exe not found");
    "CygwinError"   = @(4, "cygwin1.dll not found");
    "NoFilesError"  = @(5, "No files found");
}

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
    exit-script $exitModes["Success"]
}
function start-download ($Job) {
    $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
    $ProcessInfo.FileName = (Resolve-Path ".\msdl.exe").Path
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
    $line = $links[$($index-1)].Line.trimstart("@{href=").trim("}")
    $filename = format-Filename $line
    $location = Join-Path -Path "$path" -ChildPath "$filename"
    $file = Invoke-WebRequest -uri "$server$line" -Method POST -Body $params
    $address = ($file.Content.Split("`n") | Select-String -Pattern "window.location=").Line.split("`"")[1]
    return @($address,$location)
}

function format-Filename ($line) {
    $result = ($line.split("/")[5]).Trim()
    $filename = (Get-Culture).textinfo.ToTitleCase($result.split(".")[0].ToLower())
    $suffix = $result.split(".")[1].ToLower()
    $result = $filename + "." + $suffix
    $result = $result -replace "%20", " " -replace "-", " "
    return $result
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

function exit-script ([array]$reason) {
    $message = $($reason[1])
    if ($reason[0] -ne 0) {
        $message = "Error: " + $message
    }
    "$message. Exiting with code $($reason[0]), press any key. . ."
    $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit $reason[0]
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
                    exit-script $exitModes["DirError"]
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
        exit-script $exitModes["MSDLError"]
    }
}

function testCygwin {
    if (-not(Test-Path cygwin1.dll)) {
        exit-script $exitModes["CygwinError"]
    }
}

function getLinks ($Course, $params) {
    $URL = Invoke-WebRequest $Course -Method POST -Body $params
    if (($URL.links[0].innerHTML) -Match "forgot") {
        Remove-Item -Path ".\params.txt"
        exit-script $exitModes["CredError"]
    }
    return $URL.Links | Select-Object href | Select-String movies/rtsp
}

function getRange ([int]$num) {
    if ($num -le 0) {
        exit-script $exitModes["NoFilesError"]
    }
    Write-Host "There are $num files."
    Write-Host "Please choose range of videos to download."
    [int]$start = getInRange 1 $num 1 "first"
    [int]$end = getInRange $start $num $num "last"

    $start
    $end
}

function getInRange ([int]$low = 1, [int]$high, [int]$def, [string]$index) {
    $prompt = "Please choose $index video number ($low-$high)"
    [int]$res = readDefault $prompt $def
    while (($res -lt $low) -or ($res -gt $high)) {
        Write-Host "Invalid choice."
        [int]$res = readDefault $prompt $def
    }
    return [int]$res
}

function printIntro {
    Write-Host "Technion Old Video Server Downloader v$version by Ran Lottem"
    Write-Host "Based on bash script by Ohad Eytan`n"
}

main # Call the function