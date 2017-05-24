$server="https://video.technion.ac.il" # Technion video server

$progressPreference = 'silentlyContinue' # For wget not to bother the user (similar to wget -q)

Write-Output "Old Technion Video Server Downloader v0.4.4 by Ran Lottem"
Write-Output "Based on bash script by Ohad Eytan"

# Check dependency
try {
    & .\msdl.exe -qh
} catch {
	Write-Output "Error: msdl.exe not found in this script's folder."
	exit
}

# Get user input
$Course = Read-Host -Prompt 'Please insert valid course url (e.g https://video.technion.ac.il/Courses/PhysMech.html).'
$User = Read-Host -Prompt 'Please insert username (e.g ran.lottem without @campus etc.)'
$Pass = Read-Host -Prompt 'Please insert password' # TODO: Set an option to set user/pw once and remember

# Get URL
$params = @{username="$User";password="$Pass"}
$URL = Invoke-WebRequest $Course -Method POST -Body $params

# Check URL
if (-not $?) {
	Write-Output "Error: Bad URL, username, or password"
	exit
}

# Get links and amount of links
#$links = $URL.Links | Select-Object href | grep -a movies/rtsp
$links = $URL.Links | Select-Object href | Select-String movies/rtsp
$num = ($links | Measure-Object).count

# Check amount
if ($num -le 0) {
	Write-Output "Error: no files found."
	exit
}

# Get file range for download
Write-Host "There are $num files."
[uint16]$start = Read-Host -Prompt "Please insert first video number (1-$num)"
while (($start -lt 1) -or ($start -gt $num)) {
    [uint16]$start = Read-Host -Prompt "Invalid start. Please insert first video number (1-$num)"
}

[uint16]$end = Read-Host -Prompt "Please insert last video number ($start-$num)"
while (($end -lt $start) -or ($end -gt $num)) {
    [uint16]$end = Read-Host -Prompt "Invalid end. Please insert last video number ($start-$num)"
}

# Trim links list
$links = $links | Select-Object -Skip ($start-1) | Select-Object -SkipLast ($num-$end)

# Let's rock!
foreach ($line in $links) {
    $line = $line.Line.trimstart("@{href=").trim("}")
	$filename = $line.split("/")[5]
    $filename = $filename.Trim() # filename for saving
    $file2 = Invoke-WebRequest -uri "$server$line" -Method POST -Body $params
    $address = ($file2.Content.Split("`n") | Select-String -Pattern "window.location=").Line.split("`"")[1]
    cmd /c .\msdl.exe -s2 $address -o $filename '2>&1' | ForEach-Object {
        $first=1
        $newline=0
        if ($_.startswith("DL")) {
            Write-Host -NoNewline "`r$_"
            $first=0
            $newline=0
        }
        else {
            if ((-not $first) -and (-not $newline)) {
                Write-Output "" # for newline between files
                $newline=1
            }
            Write-Output "$_"
        }
    }
}