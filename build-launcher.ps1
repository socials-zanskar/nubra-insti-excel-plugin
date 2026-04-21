$ErrorActionPreference = "Stop"
$PSNativeCommandUseErrorActionPreference = $true

$pluginRoot = (Resolve-Path $PSScriptRoot).Path
$launcherSource = Join-Path $pluginRoot "launcher.cs"
$launcherExe = Join-Path $pluginRoot "NubraInstiExcelLauncher.exe"
$signtoolCandidates = @(
  "${env:ProgramFiles(x86)}\Windows Kits\10\bin\x64\signtool.exe",
  "${env:ProgramFiles(x86)}\Windows Kits\10\bin\10.0.26100.0\x64\signtool.exe",
  "${env:ProgramFiles(x86)}\Windows Kits\10\bin\10.0.22621.0\x64\signtool.exe",
  "${env:ProgramFiles(x86)}\Windows Kits\10\bin\10.0.22000.0\x64\signtool.exe",
  "${env:ProgramFiles(x86)}\Windows Kits\10\bin\10.0.19041.0\x64\signtool.exe"
)

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

$signtoolPath = $signtoolCandidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
if ($env:NUBRA_SIGN_CERT_THUMBPRINT -or $env:NUBRA_SIGN_PFX_PATH) {
  if (-not $signtoolPath) {
    throw "Signing requested, but signtool.exe was not found. Install the Windows SDK."
  }

  if ($env:NUBRA_SIGN_CERT_THUMBPRINT) {
    & $signtoolPath sign /sha1 $env:NUBRA_SIGN_CERT_THUMBPRINT /fd SHA256 /tr "http://timestamp.digicert.com" /td SHA256 $launcherExe
  } else {
    if (-not (Test-Path $env:NUBRA_SIGN_PFX_PATH)) {
      throw "NUBRA_SIGN_PFX_PATH does not exist."
    }
    & $signtoolPath sign /f $env:NUBRA_SIGN_PFX_PATH /p $env:NUBRA_SIGN_PFX_PASSWORD /fd SHA256 /tr "http://timestamp.digicert.com" /td SHA256 $launcherExe
  }

  if ($LASTEXITCODE -ne 0) {
    throw "Failed to sign launcher."
  }
}

Write-Host "[build] EXE: $launcherExe"
