<#
.SYNOPSIS
    Installs the GTK3 runtime required by WeasyPrint on Windows.

.DESCRIPTION
    WeasyPrint needs libgobject, libpango, libcairo etc. from the GTK3 runtime.
    This script downloads the latest gtk3-runtime installer from:
        https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer
    and installs it to %LOCALAPPDATA%\GTK3-Runtime (no admin rights needed),
    then adds that bin folder to the current user's PATH.

.EXAMPLE
    .\setup_weasyprint_windows.ps1
#>

$ErrorActionPreference = "Stop"

# ── 1. Check whether GTK3 libraries are already findable ────────────────────
function Test-Gtk3Available {
    $searchPaths = ($env:PATH -split ";") + @(
        "$env:ProgramFiles\GTK3-Runtime Win64\bin",
        "$env:ProgramFiles\GTK3-Runtime\bin",
        "$env:LOCALAPPDATA\GTK3-Runtime\bin"
    )
    foreach ($dir in $searchPaths) {
        if ($dir -and (Test-Path (Join-Path $dir "libgobject-2.0-0.dll"))) {
            return $dir
        }
    }
    return $null
}

$found = Test-Gtk3Available
if ($found) {
    Write-Host "GTK3 is already installed: $found"
    Write-Host "WeasyPrint should work. If you still get errors, restart your terminal."
    exit 0
}

# ── 2. Resolve latest release from GitHub ───────────────────────────────────
Write-Host "Querying GitHub for the latest GTK3 Windows Runtime release..."
try {
    $release = Invoke-RestMethod `
        -Uri "https://api.github.com/repos/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases/latest" `
        -Headers @{ "User-Agent" = "md_to_docx-setup" }
} catch {
    Write-Error "Could not reach GitHub API. Check your internet connection and retry.`n$_"
    exit 1
}

# Prefer 64-bit, fall back to any .exe asset
$asset = $release.assets |
    Where-Object { $_.name -match "x86_64.*\.exe$" } |
    Select-Object -First 1

if (-not $asset) {
    $asset = $release.assets |
        Where-Object { $_.name -match "\.exe$" } |
        Select-Object -First 1
}

if (-not $asset) {
    Write-Error "No installer asset found in the latest GitHub release. Check manually:`nhttps://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases"
    exit 1
}

$downloadUrl  = $asset.browser_download_url
$assetName    = $asset.name
$installerTmp = Join-Path $env:TEMP $assetName

# ── 3. Download ──────────────────────────────────────────────────────────────
Write-Host "Downloading $assetName ..."
Write-Host "  from $downloadUrl"
try {
    Invoke-WebRequest -Uri $downloadUrl -OutFile $installerTmp -UseBasicParsing
} catch {
    Write-Error "Download failed: $_"
    exit 1
}
Write-Host "Download complete: $installerTmp"

# ── 4. Install per-user (no admin required) ──────────────────────────────────
$installDir = Join-Path $env:LOCALAPPDATA "GTK3-Runtime"
Write-Host ""
Write-Host "Installing to: $installDir"
Write-Host "(Silent install — this may take 10-30 seconds)"

$proc = Start-Process `
    -FilePath $installerTmp `
    -ArgumentList "/S", "/D=$installDir" `
    -PassThru -Wait

if ($proc.ExitCode -ne 0) {
    Write-Warning "Installer exited with code $($proc.ExitCode). Trying to continue anyway..."
}

# ── 5. Verify DLL is now present ─────────────────────────────────────────────
$gtkBin = Join-Path $installDir "bin"
if (-not (Test-Path (Join-Path $gtkBin "libgobject-2.0-0.dll"))) {
    Write-Error "Installation seemed to succeed but libgobject-2.0-0.dll was not found in`n  $gtkBin`nTry running the installer manually:`n  $installerTmp"
    exit 1
}
Write-Host "GTK3 libraries verified in: $gtkBin"

# ── 6. Add bin to current-user PATH if not already present ──────────────────
$userPath = [System.Environment]::GetEnvironmentVariable("PATH", "User")
if ($userPath -notlike "*GTK3-Runtime*") {
    $newUserPath = ($userPath.TrimEnd(";") + ";" + $gtkBin).TrimStart(";")
    [System.Environment]::SetEnvironmentVariable("PATH", $newUserPath, "User")
    Write-Host "Added to user PATH: $gtkBin"
} else {
    Write-Host "GTK3 bin already in user PATH."
}

# Also update the current process PATH so we can test immediately
$env:PATH = $env:PATH.TrimEnd(";") + ";" + $gtkBin

# ── 7. Quick smoke-test ───────────────────────────────────────────────────────
Write-Host ""
Write-Host "Running WeasyPrint smoke-test..."
$testResult = & python -c "from weasyprint import HTML; HTML(string='<p>ok</p>').write_pdf('$env:TEMP\\weasyprint_test.pdf'); print('WeasyPrint OK')" 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host $testResult
    Remove-Item "$env:TEMP\weasyprint_test.pdf" -ErrorAction SilentlyContinue
    Write-Host ""
    Write-Host "Setup complete! WeasyPrint is ready to use."
    Write-Host "You can now run:"
    Write-Host "  python convert.py your_file.md --pdf --pdf-backend weasyprint"
} else {
    Write-Warning "WeasyPrint test failed. You may need to open a new terminal window first.`n$testResult"
    Write-Host ""
    Write-Host "After opening a new terminal, test with:"
    Write-Host "  python -c `"from weasyprint import HTML; HTML(string='<p>ok</p>').write_pdf('test.pdf')`""
}

