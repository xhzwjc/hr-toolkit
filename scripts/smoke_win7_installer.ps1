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

function Invoke-InstalledSmoke {
    param([string]$Executable, [string]$Argument, [string]$Expected, [int]$TimeoutSeconds = 90)
    Remove-Item -LiteralPath $resultPath -ErrorAction SilentlyContinue
    $process = Start-Process -FilePath $Executable -ArgumentList $Argument -WorkingDirectory $installRoot -PassThru
    # Retain the handle so ExitCode remains available in Windows PowerShell.
    $processHandle = $process.Handle
    if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
        $process.Kill()
        $process.WaitForExit()
        throw "$Argument timed out; check for a native loader error dialog."
    }
    if ($process.ExitCode -ne 0) {
        throw "$Argument exited with code $($process.ExitCode)."
    }
    $result = ([string](Get-Content -LiteralPath $resultPath -Encoding UTF8 | Out-String)).Trim()
    if ($result -ne $Expected) {
        throw "Unexpected $Argument output: $result"
    }
    Write-Host $Expected
}

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
        (Join-Path $internal "ucrtbase.dll"),
        (Join-Path $internal "msvcp140.dll"),
        (Join-Path $internal "vcruntime140.dll"),
        (Join-Path $internal "vcruntime140_1.dll"),
        (Join-Path $internal "third_party\7zip\7z.exe"),
        (Join-Path $internal "third_party\7zip\7z.dll")
    )
    foreach ($path in $required) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "Installed Win7 payload is missing: $path"
        }
    }
    $runtimeNames = @("ucrtbase.dll", "msvcp140.dll", "vcruntime140.dll", "vcruntime140_1.dll")
    $runtimeNames += @(Get-ChildItem -LiteralPath $payload -Filter "api-ms-win-*.dll" | ForEach-Object { $_.Name })
    if ($runtimeNames.Count -ne 44) {
        throw "Installed Win7 payload does not contain the complete pinned runtime set."
    }
    # Use .NET hashing for Windows PowerShell 2 compatibility.
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        foreach ($name in $runtimeNames) {
            $rootFile = Join-Path $payload $name
            $internalFile = Join-Path $internal $name
            if (-not (Test-Path -LiteralPath $internalFile -PathType Leaf)) {
                throw "Installed Win7 bootstrap runtime is missing: $internalFile"
            }
            $rootHash = [Convert]::ToBase64String($sha.ComputeHash([IO.File]::ReadAllBytes($rootFile)))
            $internalHash = [Convert]::ToBase64String($sha.ComputeHash([IO.File]::ReadAllBytes($internalFile)))
            if ($rootHash -ne $internalHash) {
                throw "Installed Win7 runtime copies differ: $name"
            }
        }
    }
    finally {
        $sha.Dispose()
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
    $expected = "HRToolkit $ExpectedVersion smoke-test OK"
    Invoke-InstalledSmoke -Executable $app -Argument "--smoke-test" -Expected $expected -TimeoutSeconds 180
    # Require three consecutive native Qt launches; this is not a retry loop.
    for ($launch = 1; $launch -le 3; $launch++) {
        Invoke-InstalledSmoke -Executable $app -Argument "--qt-smoke-test" -Expected "HRToolkit Qt smoke-test OK"
    }

    $updater = Join-Path $payload "HRToolkitUpdater.exe"
    $expectedUpdater = "HRToolkitUpdater $ExpectedVersion smoke-test OK"
    Invoke-InstalledSmoke -Executable $updater -Argument "--smoke-test" -Expected $expectedUpdater
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
