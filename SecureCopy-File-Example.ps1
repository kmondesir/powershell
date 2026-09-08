<#
.SYNOPSIS
Selects a file from the project directory and copies it to an EC2 instance with SCP.

.DESCRIPTION
Opens a file picker for the configured project directory, then transfers the selected
file to /tmp/ on the remote instance using the specified SSH private key.

.PARAMETER private_keyPath
Path to the SSH private key used for authentication.

.PARAMETER user
SSH username for the remote instance.

.PARAMETER instance
Remote SSH hostname or IP address for the EC2 instance.

.PARAMETER projectDirectory
Local directory containing files available for transfer.

.EXAMPLE
.\SecureCopy-File-Example.ps1
Uses the default private key, SSH user, instance, and project directory.

.EXAMPLE
.\SecureCopy-File-Example.ps1 -private_keyPath "C:\path\to\your\private-key.pem" -user "ec2-user" -instance "example.compute.amazonaws.com" -projectDirectory "C:\path\to\your\project"
Uses explicitly supplied connection and directory settings.

.EXAMPLE
Get-Help .\SecureCopy-File-Example.ps1 -Detailed
Displays the full help for this script.
#>
param(
    [Parameter(Mandatory = $false, HelpMessage = "Path to the SSH private key file.")]
    [string]$private_keyPath = "C:\path\to\your\private-key.pem",

    [Parameter(Mandatory = $false, HelpMessage = "SSH username for the remote instance.")]
    [string]$user = "ec2-user",

    [Parameter(Mandatory = $false, HelpMessage = "Remote SSH hostname or IP address for the EC2 instance.")]
    [string]$instance = "example.compute.amazonaws.com",

    [Parameter(Mandatory = $false, HelpMessage = "Local directory containing files to transfer.")]
    [string]$projectDirectory = "C:\path\to\your\project"
)

# 1. Show the available files from the configured project directory and keep prompting until the user exits
$files = Get-ChildItem -Path $projectDirectory -File | Sort-Object Name

if ($files.Count -eq 0) {
    Write-Warning "No files were found in $projectDirectory. SCP aborted."
    return
}

while ($true) {
    Write-Host "Files available in ${projectDirectory}:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $files.Count; $i++) {
        Write-Host ("[{0}] {1}" -f ($i + 1), $files[$i].Name)
    }
    Write-Host "[0] Exit" -ForegroundColor Yellow
    Write-Host "Enter 0 to exit. To select multiple files, use comma-separated numbers or ranges like 1,3,5 or 2-4." -ForegroundColor DarkGray

    $choice = Read-Host "Your selection"

    if ($choice -eq '0') {
        Write-Host "Exiting file transfer menu." -ForegroundColor Green
        return
    }

    $selectedNumbers = @()
    foreach ($token in ($choice -split ',')) {
        $token = $token.Trim()
        if ($token -eq '') { continue }

        if ($token -match '^\d+-\d+$') {
            $rangeStart, $rangeEnd = $token -split '-' | ForEach-Object { [int]$_ }
            if ($rangeStart -gt $rangeEnd) {
                $tmp = $rangeStart
                $rangeStart = $rangeEnd
                $rangeEnd = $tmp
            }

            for ($n = $rangeStart; $n -le $rangeEnd; $n++) {
                $selectedNumbers += [string]$n
            }
        }
        elseif ($token -match '^\d+$') {
            $selectedNumbers += $token
        }
        else {
            $selectedNumbers = @('INVALID')
            break
        }
    }

    if ($selectedNumbers.Count -eq 0) {
        Write-Warning "No selection entered. Enter 0 to exit or choose a valid file number/range."
        continue
    }

    if ($selectedNumbers[0] -eq 'INVALID') {
        Write-Warning "Invalid selection. Enter 0 to exit or use valid file numbers/ranges."
        continue
    }

    $selectedNumbers = $selectedNumbers | Sort-Object { [int]$_ } -Unique

    $invalidSelection = $false
    $selectedFiles = @()

    foreach ($num in $selectedNumbers) {
        if ([int]$num -lt 1 -or [int]$num -gt $files.Count) {
            $invalidSelection = $true
            break
        }

        $selectedFiles += $files[[int]$num - 1]
    }

    if ($invalidSelection) {
        Write-Warning "Invalid selection. Enter 0 to exit or choose valid file numbers/ranges only."
        continue
    }

    foreach ($selectedFile in $selectedFiles) {
        Write-Host "Preparing to send: $($selectedFile.FullName)" -ForegroundColor Green
        $remoteTarget = "$user@$instance`:/tmp/"
        scp -i $private_keyPath $selectedFile.FullName $remoteTarget
    }

    Write-Host "Transfer complete. Choose another file or enter 0 to exit." -ForegroundColor Cyan
}
