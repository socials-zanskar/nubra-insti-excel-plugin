$ErrorActionPreference = "Continue"

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

  return $null
}

$pluginRoot = (Resolve-Path $PSScriptRoot).Path
$manifestPath = Join-Path $pluginRoot "manifest.xml"
$devSettingsCli = Join-Path $pluginRoot "node_modules\office-addin-dev-settings\cli.js"
$nodeExe = Resolve-NodeExe -Root $pluginRoot

if ($nodeExe -and (Test-Path $devSettingsCli)) {
  & $nodeExe $devSettingsCli unregister $manifestPath | Out-Host
}

$procs = Get-CimInstance Win32_Process -Filter "name = 'node.exe'" |
  Where-Object { $_.CommandLine -match "dev-server\.js" -and $_.CommandLine -match [regex]::Escape($pluginRoot) }

foreach ($p in $procs) {
  try {
    Stop-Process -Id $p.ProcessId -Force
  } catch {
    Write-Host $_.Exception.Message
  }
}
