param(
  [switch]$NoSideload
)

$ErrorActionPreference = "Stop"
$PSNativeCommandUseErrorActionPreference = $true

function Resolve-NodeExe {
  param([string]$Root)

  $candidates = @(
    (Join-Path $Root "runtime\node\node.exe"),
    (Join-Path $env:ProgramFiles "nodejs\node.exe")
  )

  foreach ($candidate in $candidates) {
    if ($candidate -and (Test-Path $candidate)) {
      return (Resolve-Path $candidate).Path
    }
  }

  $cmd = Get-Command "node" -ErrorAction SilentlyContinue
  if ($cmd) {
    return $cmd.Source
  }

  throw "Node.js runtime not found."
}

function Invoke-NodeCli {
  param(
    [string]$NodeExe,
    [string]$CliScript,
    [string[]]$Arguments
  )

  & $NodeExe $CliScript @Arguments
  if ($LASTEXITCODE -ne 0) {
    throw "Command failed: $CliScript $($Arguments -join ' ')"
  }
}

function Test-ServerReady {
  $null = curl.exe -k -s https://localhost:3000/taskpane.html
  return ($LASTEXITCODE -eq 0)
}

$pluginRoot = (Resolve-Path $PSScriptRoot).Path
$manifestPath = Join-Path $pluginRoot "manifest.xml"
$serverScript = Join-Path $pluginRoot "dev-server.js"
$devSettingsCli = Join-Path $pluginRoot "node_modules\office-addin-dev-settings\cli.js"
$nodeExe = Resolve-NodeExe -Root $pluginRoot

$existing = Get-CimInstance Win32_Process -Filter "name = 'node.exe'" |
  Where-Object { $_.CommandLine -match "dev-server\.js" -and $_.CommandLine -match [regex]::Escape($pluginRoot) } |
  Select-Object -First 1

if (-not $existing) {
  Start-Process -FilePath $nodeExe -ArgumentList "`"$serverScript`"" -WorkingDirectory $pluginRoot | Out-Null
}

for ($i = 0; $i -lt 20; $i++) {
  if (Test-ServerReady) { break }
  Start-Sleep -Seconds 1
}

if (-not (Test-ServerReady)) {
  throw "Dev server did not become ready on https://localhost:3000"
}

if ($NoSideload) {
  exit 0
}

Push-Location $pluginRoot
try {
  Invoke-NodeCli -NodeExe $nodeExe -CliScript $devSettingsCli -Arguments @("register", $manifestPath)
  Invoke-NodeCli -NodeExe $nodeExe -CliScript $devSettingsCli -Arguments @("sideload", $manifestPath, "desktop", "--app", "excel")
} finally {
  Pop-Location
}
