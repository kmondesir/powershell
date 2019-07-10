Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

$ChocoInstall = Join-Path ([System.Environment]::GetFolderPath('CommonApplicationData')) 'chocolatey\choco.exe'

if (Test-Path -Path $ChocoInstall)

 {
 
 choco source remove -name=SMIL
 choco upgrade chocolatey -y --force
 refreshenv
 choco upgarde chocolatey-agent -y --force
 choco source add -n SMIL -s "'https://ias-smil01.nbrem.no/chocolatey'"
 choco feature disable --name="'showNonElevatedWarnings'"
 choco feature enable --name="'useBackgroundService'"
 choco feature enable --name="'useBackgroundServiceWithNonAdministratorsOnly'"

 choco install argus -y --force
 
  }

 else

 {

 Set-ExecutionPolicy Bypass -Scope Process -Force; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))

 choco source remove -name=SMIL
 choco upgrade chocolatey -y --force
 refreshenv
 choco upgarde chocolatey-agent -y --force
 choco source add -n SMIL -s "'https://ias-smil01.nbrem.no/chocolatey'"
 choco feature disable --name="'showNonElevatedWarnings'"
 choco feature enable --name="'useBackgroundService'"
 choco feature enable --name="'useBackgroundServiceWithNonAdministratorsOnly'"

 choco install argus -y --force

 }

