param(
  [switch]$SkipNpmInstall,
  [switch]$AdminOnly
)

$ErrorActionPreference = "Stop"
$PSNativeCommandUseErrorActionPreference = $true

function Test-IsAdmin {
  $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
  $principal = New-Object Security.Principal.WindowsPrincipal($identity)
  return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

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

$pluginRoot = (Resolve-Path $PSScriptRoot).Path
$manifestPath = Join-Path $pluginRoot "manifest.xml"
$nodeModulesPath = Join-Path $pluginRoot "node_modules"
$devCertCli = Join-Path $pluginRoot "node_modules\office-addin-dev-certs\cli.js"
$devSettingsCli = Join-Path $pluginRoot "node_modules\office-addin-dev-settings\cli.js"
$nodeExe = Resolve-NodeExe -Root $pluginRoot

Push-Location $pluginRoot
try {
  if (-not $AdminOnly -and -not $SkipNpmInstall -and -not (Test-Path $nodeModulesPath)) {
    npm.cmd install
    if ($LASTEXITCODE -ne 0) { throw "npm install failed." }
  }

  if (-not $AdminOnly) {
    # Install the localhost dev cert in the current user's trust store.
    Invoke-NodeCli -NodeExe $nodeExe -CliScript $devCertCli -Arguments @("install")
  }

  if (-not (Test-IsAdmin)) {
    $argList = @("-ExecutionPolicy", "Bypass", "-File", "`"$PSCommandPath`"", "-AdminOnly", "-SkipNpmInstall")
    $proc = Start-Process -FilePath "powershell.exe" -ArgumentList $argList -Verb RunAs -Wait -PassThru
    exit $proc.ExitCode
  }

  Invoke-NodeCli -NodeExe $nodeExe -CliScript $devSettingsCli -Arguments @("appcontainer", $manifestPath, "--loopback", "-y")
  cmd /c "CheckNetIsolation LoopbackExempt -a -n=Microsoft.Win32WebViewHost_cw5n1h2txyewy" | Out-Host
  cmd /c "CheckNetIsolation LoopbackExempt -a -n=Microsoft.MicrosoftOfficeHub_8wekyb3d8bbwe" | Out-Host
} finally {
  Pop-Location
}
