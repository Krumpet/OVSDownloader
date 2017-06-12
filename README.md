# OVSDownloader
 
 TL;DR:
 1) Go to https://github.com/Krumpet/OVSDownloader/releases and download msdl.exe, downloader.exe, and cygwin1.dll.
 2) Extract cygwin1.dll to get cygwin1.dll
 3) Both exe files and the dll file should now be in the same folder
 4) Run downloader.exe and follow the instructions
 5) For known errors see bottom of the document, please let me know of others you encounter

---------
README.MD
---------

This is a Powershell script that can be used to download .wmv files via the RTSP protocol from the Old Technion Video Server at:
http://video.technion.ac.il/Courses/

This script uses msdl, compiled for windows by me, supplied here as an exe file, which also requires cygwin1.dll to be present in the same folder. This DLL file is provided inside a ZIP file, which needs to be extracted to the same folder as msdl.exe, and downloader.exe (or downloader.ps1, if you want to use the script directly).

The PS1 script is then compiled with PS2EXE-GUI created initially by Ingo Karstein and improved upon (including PS 5.0 support) by Markus Scholtes.

Version summary:

MSDL:

version: 1.2.7-r2

link: http://msdl.sourceforge.net/

GCC:

version: 5.4.0

PS2EXE:

version:   0.5.0.5

link: https://gallery.technet.microsoft.com/PS2EXE-GUI-Convert-e7cb69d5

Cygwin:

version (uname -a): CYGWIN_NT-10.0 {My computer name} 2.8.0(0.309/5/3) 2017-04-01 20:47 x86_64 Cygwin

---------
KNOWN ERRORS
---------
* Error when running downloader.exe:
'Unhandled Exception: System.IO.FileNotFoundException: Could not load file or assembly 'System.Management.Automation, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35' or one of its dependencies. The system cannot find the file specified.
at ik.PowerShell.PS2EXE.Main(String[] args)'

This has something to do with Windows Management Framework 3.0, please go here: https://www.microsoft.com/en-us/download/details.aspx?id=34595 and download the appropriate file for your OS (either 32 or 64-bit, and 6.0 for Vista, 6.1 for Win7).

Retry downloader.exe after installing the above.
