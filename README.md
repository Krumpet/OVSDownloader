# OVSDownloader
 
 TL;DR:
 1) Go to https://github.com/Krumpet/OVSDownloader/releases and download msdl.exe, downloader.exe, and cygwin1.dll.
 2) Extract cygwin1.dll to get cygwin1.dll
 3) Both exe files and the dll file should now be in the same folder
 4) Run downloader.exe and follow the instructions

---------
README.MD
---------

This is a Powershell script that can be used to download .wmv files via the RTSP protocol from the Old Technion Video Server at:
http://video.technion.ac.il/Courses/

This script uses msdl, compiled for windows by me, supplied here as an exe file, which also requires cygwin1.dll to be present in the same folder. This DLL file is provided inside a ZIP file, which needs to be extracted to the same folder as msdl.exe, and downloader.exe (or downloader.ps1, if you want to use the script directly).

The PS1 script is then compiled with PS2EXE-GUI (https://gallery.technet.microsoft.com/PS2EXE-GUI-Convert-e7cb69d5) created initially by Ingo Karstein and improved upon (including PS 5.0 support) by Markus Scholtes.

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

version (uname -a): CYGWIN_NT-10.0 DESKTOP-907D4D2 2.8.0(0.309/5/3) 2017-04-01 20:47 x86_64 Cygwin
