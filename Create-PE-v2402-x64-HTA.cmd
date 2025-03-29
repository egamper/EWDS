REM ADK Version 10.0.26100.1
REM https://learn.microsoft.com/en-us/windows-hardware/get-started/adk-install
REM Download and install ADK and WINPE AddOn
REM https://www.deploymentresearch.com/adding-adsi-support-for-winpe-10/

powershell.exe -command  Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
refreshenv

powershell.exe -command  choco install windows-adk --force -y --acceptlicense
powershell.exe -command  choco install windows-adk-winpe --force -y --acceptlicense
powershell.exe -command  choco install git.install --force -y --acceptlicense
powershell.exe -command  choco install rufus --force -y --acceptlicense
refreshenv

echo ------------------------------------------------------------
echo ------------------------------------------------------------
echo USE ADK CMD!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
echo ------------------------------------------------------------
echo ------------------------------------------------------------


rd "C:\temp\WinPE-v2402" /S /Q
md C:\temp
md C:\temp\WinPE-v2402
md C:\temp\WinPE-v2402\WinPE-AMD64

git clone https://github.com/egamper/EWDS.git C:\temp\EWDS-GIT

Dism /Cleanup-Mountpoints

REM dism /Get-WimInfo /WimFile:"C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\amd64\en-us\winpe.wim" /index:1

copy "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\amd64\en-us\winpe.wim" "C:\temp\WinPE-v2402\WinPE-v2402-AMD64.wim" /y

dism /Unmount-Wim /MountDir:"C:\temp\WinPE-v2402\WinPE-AMD64" /discard
dism /Mount-Wim /WimFile:"C:\temp\WinPE-v2402\WinPE-v2402-AMD64.wim" /Index:1 /MountDir:"C:\temp\WinPE-v2402\WinPE-AMD64"


dism /image:"C:\temp\WinPE-v2402\WinPE-AMD64" /Add-Package /PackagePath:"C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\amd64\WinPE_OCs\WinPE-HTA.cab" 
dism /image:"C:\temp\WinPE-v2402\WinPE-AMD64" /Add-Package /PackagePath:"C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\amd64\WinPE_OCs\en-us\WinPE-HTA_en-us.cab"
dism /image:"C:\temp\WinPE-v2402\WinPE-AMD64" /Add-Package /PackagePath:"C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\amd64\WinPE_OCs\WinPE-Scripting.cab"
dism /image:"C:\temp\WinPE-v2402\WinPE-AMD64" /Add-Package /PackagePath:"C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\amd64\WinPE_OCs\en-us\WinPE-Scripting_en-us.cab"
dism /image:"C:\temp\WinPE-v2402\WinPE-AMD64" /Add-Package /PackagePath:"C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\amd64\WinPE_OCs\WinPE-WMI.cab"
dism /image:"C:\temp\WinPE-v2402\WinPE-AMD64" /Add-Package /PackagePath:"C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\amd64\WinPE_OCs\en-us\WinPE-WMI_en-us.cab"
dism /image:"C:\temp\WinPE-v2402\WinPE-AMD64" /set-inputlocale:de-DE

REM https://docs.microsoft.com/en-us/surface/enable-surface-keyboard-for-windows-pe-deployment integrate Surface laptop keyboard drivers
REM https://support.microsoft.com/it-it/surface/download-di-driver-e-firmware-per-surface-09bb2e09-2a4b-cb69-0951-078a7739e120 All drivers Surface

dism /image:"C:\temp\WinPE-v2402\WinPE-AMD64" /Add-Driver /driver:"C:\temp\Drivers\NIC"

dism /image:"C:\temp\WinPE-v2402\WinPE-AMD64" /Add-Driver /driver:"c:\temp\ADSI-v10-1809\ADSIx64\ADSIx64.inf" /forceunsigned

md C:\temp\WinPE-v2402\WinPE-AMD64\EWDS

REM ---------------------------------------------------------------------------------------
copy "C:\temp\EWDS-GIT\EWDS.hta" C:\temp\WinPE-v2402\WinPE-AMD64\EWDS\EWDS.hta /Y

ECHO X:\EWDS\EWDS.hta >> C:\temp\WinPE-v2402\WinPE-AMD64\windows\System32\startnet.cmd
ECHO. >> C:\temp\WinPE-v2402\WinPE-AMD64\windows\System32\startnet.cmd
TYPE C:\temp\WinPE-v2402\WinPE-AMD64\windows\System32\startnet.cmd

dism /Get-MountedWimInfo

PAUSE

dism /Unmount-Wim /MountDir:"C:\temp\WinPE-v2402\WinPE-AMD64" /commit
REM dism /Unmount-Wim /MountDir:"C:\temp\WinPE-v2402\WinPE-AMD64" /discard

REM Make ISO
copype amd64 C:\Temp\WinPE-v2402\ISO
copy "C:\temp\WinPE-v2402\WinPE-v2402-AMD64.wim" "C:\temp\WinPE-v2402\ISO\media\sources\boot.wim" /Y
MakeWinPEMedia /ISO C:\Temp\WinPE-v2402\ISO C:\temp\WinPE-v2402\WinPE-v2402-AMD64-HTA-v20250107.iso

Rufus.exe

REM Create ISO with rufus.exe
REM Select ISO
REM Partition scheme: GPT
REM Target system: UEFI(non CSM)
REM File system: FAT32 (Default)
