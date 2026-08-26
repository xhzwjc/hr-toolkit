param(
    [Parameter(Mandatory = $true)]
    [string]$InstallerPath,

    [Parameter(Mandatory = $true)]
    [string]$ExpectedVersion
)

$ErrorActionPreference = "Stop"
$os = Get-WmiObject -Class Win32_OperatingSystem
if ($os.Version -notlike "6.1.*" -or [int]$os.ServicePackMajorVersion -lt 1) {
    throw "This acceptance script must run on Windows 7 SP1."
}
if ([string]$os.OSArchitecture -notmatch "64") {
    throw "This acceptance script requires Windows 7 SP1 x64."
}

$installer = (Resolve-Path -LiteralPath $InstallerPath).Path
$installRoot = Join-Path $env:TEMP ("HRToolkit-Win7-Smoke-" + [Guid]::NewGuid().ToString("N"))
$resultPath = Join-Path $installRoot "smoke-result.txt"

try {
    $installArgs = @(
        "/VERYSILENT",
        "/SUPPRESSMSGBOXES",
        "/NORESTART",
        "/SP-",
        "/DIR=`"$installRoot`""
    )
    $install = Start-Process -FilePath $installer -ArgumentList $installArgs -Wait -PassThru
    if ($install.ExitCode -ne 0) {
        throw "Win7 installer exited with code $($install.ExitCode)."
    }

    $payload = Join-Path $installRoot "app"
    $internal = Join-Path $payload "_internal"
    $required = @(
        (Join-Path $payload "HRToolkit.exe"),
        (Join-Path $payload "HRToolkitUpdater.exe"),
        (Join-Path $internal "python38.dll"),
        (Join-Path $internal "python3.dll"),
        (Join-Path $payload "ucrtbase.dll"),
        (Join-Path $payload "msvcp140.dll"),
        (Join-Path $payload "vcruntime140.dll"),
        (Join-Path $payload "vcruntime140_1.dll"),
        (Join-Path $internal "third_party\7zip\7z.exe"),
        (Join-Path $internal "third_party\7zip\7z.dll")
    )
    foreach ($path in $required) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "Installed Win7 payload is missing: $path"
        }
    }
    if (Test-Path -LiteralPath (Join-Path $internal "python312.dll")) {
        throw "Installed Win7 payload contains the modern Python runtime."
    }
    if (Test-Path -LiteralPath (Join-Path $internal "api-ms-win-core-path-l1-1-0.dll")) {
        throw "Installed Win7 payload contains the unsupported path API shim."
    }

    $env:HR_TOOLKIT_CHECK_OUTPUT = $resultPath
    Remove-Item Env:\HR_TOOLKIT_7ZIP_EXE -ErrorAction SilentlyContinue
    $app = Join-Path $payload "HRToolkit.exe"
    $smoke = Start-Process -FilePath $app -ArgumentList "--smoke-test" -Wait -PassThru
    if ($smoke.ExitCode -ne 0) {
        throw "HRToolkit smoke test exited with code $($smoke.ExitCode)."
    }
    $result = [string](Get-Content -LiteralPath $resultPath | Out-String)
    $expected = "HRToolkit $ExpectedVersion smoke-test OK"
    if ($result -notmatch [Regex]::Escape($expected)) {
        throw "Unexpected smoke-test output: $result"
    }

    Remove-Item -LiteralPath $resultPath -ErrorAction SilentlyContinue
    $updater = Join-Path $payload "HRToolkitUpdater.exe"
    $updaterSmoke = Start-Process -FilePath $updater -ArgumentList "--smoke-test" -Wait -PassThru
    if ($updaterSmoke.ExitCode -ne 0) {
        throw "HRToolkit updater smoke test exited with code $($updaterSmoke.ExitCode)."
    }
    $updaterResult = [string](Get-Content -LiteralPath $resultPath | Out-String)
    $expectedUpdater = "HRToolkitUpdater $ExpectedVersion smoke-test OK"
    if ($updaterResult -notmatch [Regex]::Escape($expectedUpdater)) {
        throw "Unexpected updater smoke-test output: $updaterResult"
    }
    Write-Host "Windows 7 SP1 x64 acceptance passed: $expected; $expectedUpdater"
}
finally {
    Remove-Item Env:\HR_TOOLKIT_CHECK_OUTPUT -ErrorAction SilentlyContinue
    $uninstaller = Get-ChildItem -LiteralPath $installRoot -Filter "unins*.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($uninstaller -ne $null) {
        Start-Process -FilePath $uninstaller.FullName -ArgumentList @(
            "/VERYSILENT",
            "/SUPPRESSMSGBOXES",
            "/NORESTART"
        ) -Wait | Out-Null
    }
    if (Test-Path -LiteralPath $installRoot) {
        Remove-Item -LiteralPath $installRoot -Recurse -Force
    }
}
