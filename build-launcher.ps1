$ErrorActionPreference = "Stop"
$PSNativeCommandUseErrorActionPreference = $true

$pluginRoot = (Resolve-Path $PSScriptRoot).Path
$launcherSource = Join-Path $pluginRoot "launcher.cs"
$launcherExe = Join-Path $pluginRoot "NubraInstiExcelLauncher.exe"

$cscPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $cscPath)) {
  $cscPath = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
}
if (-not (Test-Path $cscPath)) {
  throw "csc.exe not found in .NET Framework paths."
}
if (-not (Test-Path $launcherSource)) {
  throw "launcher.cs not found."
}

& $cscPath /nologo /target:exe /out:$launcherExe $launcherSource
if ($LASTEXITCODE -ne 0) {
  throw "Failed to compile launcher.cs"
}

Write-Host "[build] EXE: $launcherExe"
