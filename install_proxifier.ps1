$ErrorActionPreference = "Stop"

$url = "https://123static.szxiot.com/Soft/over_firewall/v2rayN-windows-64-desktop.zip"
$zipPath = "$env:USERPROFILE\Downloads\v2rayN-windows-64-desktop.zip"
$extractPath = $env:USERPROFILE
$desktopPath = [Environment]::GetFolderPath("Desktop")
$subscriptionUrl = "https://123static.szxiot.com/Soft/cat_sub/netch_sub.txt"

function Install-Git {
    Write-Host "Git is not installed. Installing Git..." -ForegroundColor Yellow

    Write-Host "Downloading Git installer from internal source..."
    $gitInstaller = "$env:TEMP\Git-Setup.exe"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri "http://172.16.16.170:19798/static/http/172.16.16.170:19798/True/115%2F%E8%BD%AF%E4%BB%B6%2FGit-2.44.0-64-bit.exe" -OutFile $gitInstaller -UseBasicParsing

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

function Add-SubscriptionToV2rayN {
    param (
        [string]$v2rayNPath,
        [string]$subUrl
    )

    Write-Host "Configuring subscription in v2rayN..."

    $v2rayNDataPath = Join-Path $env:APPDATA "v2rayN"
    $guiConfPath = Join-Path $v2rayNDataPath "guiConfs"

    if (-not (Test-Path $guiConfPath)) {
        New-Item -ItemType Directory -Path $guiConfPath -Force | Out-Null
    }

    $subConfigPath = Join-Path $guiConfPath "subscriptionList.json"
    $subList = @(
        @{
            remark = "SS Subscription"
            url = $subUrl
            defPool = $null
            sort = 0
        }
    ) | ConvertTo-Json -Depth 5

    Set-Content -Path $subConfigPath -Value $subList -Encoding UTF8
    Write-Host "Subscription URL configured: $subUrl"
}

$extractedFolder = Get-ChildItem -Path $extractPath -Directory | Where-Object { $_.Name -like "*v2rayN*" } | Select-Object -First 1

if ($extractedFolder) {
    Write-Host "v2rayN is already extracted. Checking configuration..."
    $exePath = Join-Path $extractedFolder.FullName "v2rayN.exe"
    if (-not (Test-Path $exePath)) {
        $exePath = Get-ChildItem -Path $extractedFolder.FullName -Filter "*.exe" -Recurse | Select-Object -First 1
    }

    if ($exePath) {
        $shortcutPath = "$desktopPath\v2rayN.lnk"
        $needsShortcut = -not (Test-Path $shortcutPath)

        Add-SubscriptionToV2rayN -v2rayNPath $extractedFolder.FullName -subUrl $subscriptionUrl

        if ($needsShortcut) {
            Write-Host "Creating desktop shortcut..."
            $WshShell = New-Object -ComObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut($shortcutPath)
            $Shortcut.TargetPath = $exePath
            $Shortcut.WorkingDirectory = (Get-Item $exePath).DirectoryName
            $Shortcut.Description = "v2rayN"
            $Shortcut.Save()
            Write-Host "Desktop shortcut created: $shortcutPath"
        } else {
            Write-Host "Desktop shortcut already exists: $shortcutPath"
        }

        Write-Host ""
        Write-Host "v2rayN is ready!" -ForegroundColor Green
        Write-Host "Location: $exePath"
        Write-Host ""
        Write-Host "IMPORTANT: Please manually refresh subscription in v2rayN:" -ForegroundColor Yellow
        Write-Host "  1. Run v2rayN"
        Write-Host "  2. Right-click on system tray icon or in the app"
        Write-Host "  3. Select 'Subscription' -> 'Update Subscription' (or similar option)"
        Write-Host ""
        Write-Host "Starting v2rayN..."
        Start-Process -FilePath $exePath
        exit 0
    } else {
        Write-Host "Warning: Executable not found in extracted folder" -ForegroundColor Yellow
    }
}

Write-Host "Starting v2rayN download..."
if (-not (Test-Path $zipPath)) {
    Get-CdpErrors -prog "v2rayN"
} else {
    Write-Host "Zip file already exists. Skipping download."
}

if (Test-Path $zipPath) {
    Write-Host "Extracting to $extractPath..."
    Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force

    $extractedFolder = Get-ChildItem -Path $extractPath -Directory | Where-Object { $_.Name -like "*v2rayN*" } | Select-Object -First 1
    if ($extractedFolder) {
        $exePath = Join-Path $extractedFolder.FullName "v2rayN.exe"
        if (-not (Test-Path $exePath)) {
            $exePath = Get-ChildItem -Path $extractedFolder.FullName -Filter "*.exe" -Recurse | Select-Object -First 1
        }

        if ($exePath) {
            Write-Host "Creating desktop shortcut..."
            $WshShell = New-Object -ComObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut("$desktopPath\v2rayN.lnk")
            $Shortcut.TargetPath = $exePath
            $Shortcut.WorkingDirectory = (Get-Item $exePath).DirectoryName
            $Shortcut.Description = "v2rayN"
            $Shortcut.Save()

            Add-SubscriptionToV2rayN -v2rayNPath $extractedFolder.FullName -subUrl $subscriptionUrl

            Write-Host "v2rayN installed successfully!" -ForegroundColor Green
            Write-Host "Location: $exePath"
            Write-Host "Desktop shortcut created: $desktopPath\v2rayN.lnk"
            Write-Host ""
            Write-Host "IMPORTANT: Please manually refresh subscription in v2rayN:" -ForegroundColor Yellow
            Write-Host "  1. Run v2rayN"
            Write-Host "  2. Right-click on system tray icon or in the app"
            Write-Host "  3. Select 'Subscription' -> 'Update Subscription' (or similar option)"
            Write-Host ""
            Write-Host "Starting v2rayN..."
            Start-Process -FilePath $exePath
            Write-Host "Zip file preserved at: $zipPath"
        } else {
            Write-Host "Warning: Executable not found in extracted folder" -ForegroundColor Yellow
        }
    } else {
        Write-Host "Warning: Could not find v2rayN folder after extraction" -ForegroundColor Yellow
    }
} else {
    Write-Host "Error: Download failed" -ForegroundColor Red
    exit 1
}