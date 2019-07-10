
function isChocolateyInstalled
{
    $ChocoInstall = Join-Path ([System.Environment]::GetFolderPath("CommonApplicationData")) "Chocolatey\bin\choco.exe"

    if (!(Test-Path $ChocoInstall))
    {
        return $false
    }
    else
    {
        return $true
    }
}

function installChocolatey
{
    $ChocoInstall = Join-Path ([System.Environment]::GetFolderPath('CommonApplicationData')) 'Chocolatey\bin\choco.exe'
    if (!(Test-Path $ChocoInstall))
    {
        Invoke-Expression ((New-Object net.webclient).DownloadString('https://ias-smil01.nbrem.no/install.ps1')) -ErrorAction Stop
    }
    else
    {
        Exit
    }
}