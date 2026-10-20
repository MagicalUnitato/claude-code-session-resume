<#
.SYNOPSIS
  Carries Claude Code sessions between PCs through the external SSD that holds your repos.

.DESCRIPTION
  Sessions are NOT stored in the repos. They live in two places on each PC:
    1. Transcripts:  ~\.claude\projects\<repo-path-encoded>\<session-id>.jsonl   (what `claude --resume` reads)
    2. Desktop app:  <Claude app data>\claude-code-sessions\<account>\<org>\local_*.json   (the sidebar list)

  This script mirrors both, for repos on the SSD only, into a bundle folder on the SSD
  (default <SSD root>\.claude-session-bundle):

    -Mode List     Show what Export and Import would do. Changes nothing. (default)
    -Mode Export   This PC -> SSD bundle.   Run on the PC you just worked on.
    -Mode Import   SSD bundle -> this PC.   Run on the PC you are about to work on.
    -Mode Sync     Import, then Export. Safe to run any time the SSD is plugged in.

  - If the SSD has a different drive letter on the other PC (G: vs E:), folder names and the
    "cwd" paths inside the files are rewritten automatically.
  - Newer wins per file (by modified time). A file that gets overwritten is backed up first.
    Nothing is ever deleted.
  - Close the Claude desktop app before importing, or the sidebar index is skipped.

.EXAMPLE
  .\Sync-ClaudeSessions.ps1                 # preview
  .\Sync-ClaudeSessions.ps1 -Mode Export    # on the desktop, when you are done working
  .\Sync-ClaudeSessions.ps1 -Mode Import    # on the laptop, before opening Claude
#>
[CmdletBinding()]
param(
  [ValidateSet('List', 'Export', 'Import', 'Sync')]
  [string]$Mode = 'List',

  # Where the bundle lives. Default: <drive root of this script>\.claude-session-bundle
  [string]$BundleDir,

  # Drive letter of the SSD on THIS PC. Default: the drive the bundle is on.
  [string]$SsdDrive,

  [string]$ClaudeHome = (Join-Path $env:USERPROFILE '.claude'),

  # Override the desktop app's <account>\<org> session-index folder (normally auto-detected).
  [string]$AppIndexDir,

  # Only handle transcripts; leave the desktop app's sidebar index alone.
  [switch]$SkipDesktopIndex,

  # Do the work even though the Claude desktop app is running.
  [switch]$Force,

  [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2

# ---------------------------------------------------------------- setup

if (-not $BundleDir) { $BundleDir = Join-Path ([IO.Path]::GetPathRoot($PSScriptRoot)) '.claude-session-bundle' }
$BundleDir = [IO.Path]::GetFullPath($BundleDir)
if (-not $SsdDrive) { $SsdDrive = $BundleDir.Substring(0, 1) }
$SsdDrive = $SsdDrive.Substring(0, 1).ToUpper()

$manifestPath   = Join-Path $BundleDir 'manifest.json'
$localProjects  = Join-Path $ClaudeHome 'projects'
$bundleProjects = Join-Path $BundleDir 'projects'
$bundleDesktop  = Join-Path $BundleDir 'desktop'
$stamp          = Get-Date -Format 'yyyyMMdd-HHmmss'
$utf8NoBom      = New-Object Text.UTF8Encoding($false)

$manifest = $null
if (Test-Path -LiteralPath $manifestPath) { $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json }
# The bundle always stores one fixed drive letter, so PCs with different letters can share it.
$canon = if ($manifest) { $manifest.ssdDrive } else { $SsdDrive }

# ---------------------------------------------------------------- helpers

function Read-TextShared([string]$Path) {
  # FileShare.ReadWrite so a session that is still open doesn't block the read.
  $fs = New-Object IO.FileStream($Path, 'Open', 'Read', 'ReadWrite')
  $sr = New-Object IO.StreamReader($fs, [Text.Encoding]::UTF8, $true)
  try { $sr.ReadToEnd() } finally { $sr.Dispose() }
}

function Copy-Shared([string]$Src, [string]$Dst) {
  $in = New-Object IO.FileStream($Src, 'Open', 'Read', 'ReadWrite')
  try {
    $out = New-Object IO.FileStream($Dst, 'Create', 'Write', 'None')
    try { $in.CopyTo($out) } finally { $out.Dispose() }
  } finally { $in.Dispose() }
}

# In raw JSON text a path is written  "G:\\Repositories\\x" . Swap the drive only where a string starts with one.
function Convert-DriveText([string]$Text, [string]$From, [string]$To) {
  $Text -replace ('"' + $From + ':\\\\'), ('"' + $To + ':\\')
}

function Get-JsonlCwd([string]$Path) {
  $fs = New-Object IO.FileStream($Path, 'Open', 'Read', 'ReadWrite')
  $sr = New-Object IO.StreamReader($fs)
  try {
    for ($i = 0; $i -lt 300 -and -not $sr.EndOfStream; $i++) {
      $m = [regex]::Match($sr.ReadLine(), '"cwd":"((?:[^"\\]|\\.)*)"')
      if ($m.Success) { return ($m.Groups[1].Value -replace '\\\\', '\') }
    }
  } finally { $sr.Dispose() }
  $null
}

function Get-IndexDrive([string]$Path) {
  $m = [regex]::Match((Read-TextShared $Path), '"cwd"\s*:\s*"([A-Za-z]):')
  if ($m.Success) { $m.Groups[1].Value.ToUpper() } else { $null }
}

function Find-AppRoot {
  $candidates = @()
  $pk = Join-Path $env:LOCALAPPDATA 'Packages'
  if (Test-Path $pk) {
    foreach ($p in Get-ChildItem $pk -Directory -Filter 'Claude_*') {
      $candidates += (Join-Path $p.FullName 'LocalCache\Roaming\Claude')   # Store (MSIX) install: the real backing store
    }
  }
  $candidates += (Join-Path $env:APPDATA 'Claude')
  foreach ($c in $candidates) { if (Test-Path (Join-Path $c 'claude-code-sessions')) { return $c } }
  foreach ($c in $candidates) { if (Test-Path $c) { return $c } }
  $null
}

# Returns @{ Dir; Acct; Org } or $null. $Hint is manifest.app (account/org ids from the other PC).
function Find-IndexDir($Hint) {
  if ($AppIndexDir) {
    return @{ Dir = $AppIndexDir; Acct = (Split-Path (Split-Path $AppIndexDir) -Leaf); Org = (Split-Path $AppIndexDir -Leaf) }
  }
  $root = Find-AppRoot
  if (-not $root) { return $null }
  $base = Join-Path $root 'claude-code-sessions'
  $leafs = @()
  if (Test-Path $base) {
    foreach ($a in Get-ChildItem $base -Directory) {
      foreach ($o in Get-ChildItem $a.FullName -Directory) {
        $leafs += [pscustomobject]@{
          Dir = $o.FullName; Acct = $a.Name; Org = $o.Name
          N   = @(Get-ChildItem $o.FullName -Filter 'local_*.json' -File).Count
        }
      }
    }
  }
  $pick = $null
  if ($Hint) { $pick = $leafs | Where-Object { $_.Acct -eq $Hint.acct -and $_.Org -eq $Hint.org } | Select-Object -First 1 }
  if (-not $pick) { $pick = $leafs | Where-Object { $_.N -gt 0 } | Sort-Object N -Descending | Select-Object -First 1 }
  if (-not $pick -and $leafs.Count -eq 1) { $pick = $leafs[0] }
  if ($pick) { return @{ Dir = $pick.Dir; Acct = $pick.Acct; Org = $pick.Org } }
  if ($Hint) {
    return @{ Dir = (Join-Path (Join-Path $base $Hint.acct) $Hint.org); Acct = $Hint.acct; Org = $Hint.org; Created = $true }
  }
  $null
}

function Test-DesktopAppRunning {
  $procs = @(Get-Process -Name claude -ErrorAction SilentlyContinue | Where-Object {
      $_.Path -and $_.Path -match 'WindowsApps\\Claude_|AnthropicClaude|\\Claude\\Claude\.exe'
    })
  $procs.Count -gt 0
}

function New-Action($Kind, $Project, $Src, $Dst, $From, $To, $Rewrite, $BackupRel, $IsSession) {
  $s = Get-Item -LiteralPath $Src
  if (Test-Path -LiteralPath $Dst) {
    $diff = ($s.LastWriteTimeUtc - (Get-Item -LiteralPath $Dst).LastWriteTimeUtc).TotalSeconds
    if ([math]::Abs($diff) -le 2) { $act = 'Same' } elseif ($diff -gt 0) { $act = 'Update' } else { $act = 'DestNewer' }
  } else { $act = 'New' }
  [pscustomobject]@{
    Kind = $Kind; Project = $Project; Src = $Src; Dst = $Dst; From = $From; To = $To
    Rewrite = $Rewrite; BackupRel = $BackupRel; IsSession = $IsSession; Action = $act
  }
}

function Get-TranscriptActions([string]$SrcProjects, [string]$SrcDrive, [string]$DstProjects, [string]$DstDrive) {
  $out = New-Object System.Collections.ArrayList
  if (-not (Test-Path -LiteralPath $SrcProjects)) { return @() }
  foreach ($dir in Get-ChildItem -LiteralPath $SrcProjects -Directory) {
    if ($dir.Name -notmatch ('^' + $SrcDrive + '--')) { continue }   # only repos on the SSD
    $dstName = $DstDrive + $dir.Name.Substring(1)
    foreach ($f in Get-ChildItem -LiteralPath $dir.FullName -Recurse -File -Force) {
      $rel = $f.FullName.Substring($dir.FullName.Length).TrimStart('\')
      $dst = Join-Path (Join-Path $DstProjects $dstName) $rel
      $isJsonl = $f.Extension -eq '.jsonl'
      $isSession = $isJsonl -and ($rel -notmatch '\\')
      [void]$out.Add((New-Action 'Transcript' $dstName $f.FullName $dst $SrcDrive $DstDrive `
            ($isJsonl -and $SrcDrive -ne $DstDrive) (Join-Path 'projects' (Join-Path $dstName $rel)) $isSession))
    }
  }
  $out.ToArray()
}

function Get-IndexActions([string]$SrcDir, [string]$SrcDrive, [string]$DstDir, [string]$DstDrive) {
  $out = New-Object System.Collections.ArrayList
  if (-not $SrcDir -or -not (Test-Path -LiteralPath $SrcDir)) { return @() }
  foreach ($f in Get-ChildItem -LiteralPath $SrcDir -Filter 'local_*.json' -File) {
    if ((Get-IndexDrive $f.FullName) -ne $SrcDrive) { continue }
    [void]$out.Add((New-Action 'Index' 'desktop' $f.FullName (Join-Path $DstDir $f.Name) $SrcDrive $DstDrive `
          ($SrcDrive -ne $DstDrive) (Join-Path 'desktop' $f.Name) $true))
  }
  $out.ToArray()
}

function Invoke-Action($A, [string]$BackupRoot) {
  $parent = Split-Path $A.Dst -Parent
  if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
  if ($A.Action -eq 'Update') {
    $bak = Join-Path $BackupRoot $A.BackupRel
    New-Item -ItemType Directory -Path (Split-Path $bak -Parent) -Force | Out-Null
    Copy-Item -LiteralPath $A.Dst -Destination $bak -Force
  }
  if ($A.Rewrite) {
    [IO.File]::WriteAllText($A.Dst, (Convert-DriveText (Read-TextShared $A.Src) $A.From $A.To), $utf8NoBom)
  } else {
    Copy-Shared $A.Src $A.Dst
  }
  [IO.File]::SetLastWriteTimeUtc($A.Dst, (Get-Item -LiteralPath $A.Src).LastWriteTimeUtc)
}

function Show-Plan([string]$Label, $Actions) {
  $Actions = @($Actions)
  $count = { param($a) @($Actions | Where-Object { $_.Action -eq $a }).Count }
  Write-Host ''
  Write-Host $Label -ForegroundColor Cyan
  Write-Host ('  files: {0} new, {1} updated, {2} unchanged, {3} skipped (destination is newer)' -f
    (& $count 'New'), (& $count 'Update'), (& $count 'Same'), (& $count 'DestNewer'))

  $changed = @($Actions | Where-Object { $_.IsSession -and $_.Action -in 'New', 'Update' })
  if ($changed.Count -eq 0) { return }
  Write-Host '  sessions to copy:'
  foreach ($g in ($changed | Where-Object Kind -eq 'Transcript' | Group-Object Project)) {
    $first = $g.Group | Select-Object -First 1
    $where = Get-JsonlCwd $first.Src
    if ($where) { $where = $where -replace ('^' + $first.From + ':'), ($first.To + ':') } else { $where = $g.Name }
    $n = @($g.Group | Where-Object Action -eq 'New').Count
    $u = @($g.Group | Where-Object Action -eq 'Update').Count
    Write-Host ('    {0}  ({1} new, {2} updated)' -f $where, $n, $u)
  }
  $titles = @($changed | Where-Object { $_.Kind -eq 'Index' -and $_.Action -eq 'New' } | ForEach-Object {
      try { (Get-Content -LiteralPath $_.Src -Raw | ConvertFrom-Json).title } catch { $null }
    } | Where-Object { $_ })
  if ($titles.Count -gt 0) {
    Write-Host ('  new in the Claude sidebar ({0}):' -f $titles.Count)
    $titles | Select-Object -First 25 | ForEach-Object { Write-Host "    - $_" }
    if ($titles.Count -gt 25) { Write-Host ('    ... and {0} more' -f ($titles.Count - 25)) }
  }
}

function Invoke-Plan($Actions, [string]$BackupRoot) {
  $done = 0; $failed = 0
  foreach ($a in @($Actions | Where-Object { $_.Action -in 'New', 'Update' })) {
    try { Invoke-Action $a $BackupRoot; $done++ }
    catch { $failed++; Write-Warning ("Could not copy {0}: {1}" -f $a.Src, $_.Exception.Message) }
  }
  Write-Host ('  copied {0} file(s){1}' -f $done, $(if ($failed) { ", $failed failed" } else { '' })) -ForegroundColor Green
}

# ---------------------------------------------------------------- plans

$wantImport = $Mode -in 'Import', 'Sync'
$wantExport = $Mode -in 'Export', 'Sync'
$preview    = $Mode -eq 'List' -or $DryRun
$useIndex   = -not $SkipDesktopIndex

Write-Host "SSD drive on this PC : ${SsdDrive}:   (bundle stores it as ${canon}:)"
Write-Host "Bundle               : $BundleDir"
Write-Host "Claude home          : $ClaudeHome"

$index = $null
if ($useIndex) {
  $index = Find-IndexDir $(if ($manifest -and $manifest.PSObject.Properties['app']) { $manifest.app } else { $null })
  if ($index) { Write-Host "Desktop app index    : $($index.Dir)" }
  else { Write-Warning 'Claude desktop app data not found; only transcripts will be handled. (Use -AppIndexDir to point at it.)' }
}

$appRunning = $useIndex -and (Test-DesktopAppRunning) -and -not $Force

function Get-ImportActions {
  if (-not $manifest) { return @() }
  $a = @(Get-TranscriptActions $bundleProjects $canon $localProjects $SsdDrive)
  if ($index -and -not $appRunning) { $a += @(Get-IndexActions $bundleDesktop $canon $index.Dir $SsdDrive) }
  $a
}
function Get-ExportActions {
  $a = @(Get-TranscriptActions $localProjects $SsdDrive $bundleProjects $canon)
  if ($index) { $a += @(Get-IndexActions $index.Dir $SsdDrive $bundleDesktop $canon) }
  $a
}

if ($appRunning -and ($wantImport -or $Mode -eq 'List')) {
  Write-Warning 'The Claude desktop app is running, so its sidebar index will NOT be imported (it would overwrite the files). Close it and re-run, or pass -Force.'
}

# ---------------------------------------------------------------- run

if ($Mode -eq 'List') {
  if (-not $manifest) { Write-Host "`nNo bundle yet at $BundleDir. Run -Mode Export on the PC that has the sessions." -ForegroundColor Yellow }
  else { Show-Plan 'IMPORT  (bundle -> this PC)' (Get-ImportActions) }
  Show-Plan 'EXPORT  (this PC -> bundle)' (Get-ExportActions)
  Write-Host "`nPreview only. Use -Mode Import / Export / Sync to apply."
  return
}

if ($wantImport) {
  if (-not $manifest) {
    if ($Mode -eq 'Import') { throw "No bundle found at $BundleDir (missing manifest.json). Run -Mode Export on the other PC first." }
    Write-Host "`nNo bundle yet, skipping import." -ForegroundColor Yellow
  } else {
    $plan = Get-ImportActions
    Show-Plan 'IMPORT  (bundle -> this PC)' $plan
    if (-not $preview) {
      if ($index -and $index.ContainsKey('Created')) { Write-Warning "No app index folder existed here; creating the one from the bundle ($($index.Acct)\$($index.Org)). If the sidebar stays empty, this PC is signed in to a different account." }
      Invoke-Plan $plan (Join-Path $ClaudeHome "session-sync-backup\$stamp")
    }
  }
}

if ($wantExport) {
  $plan = Get-ExportActions     # re-planned after import so mtimes are current
  Show-Plan 'EXPORT  (this PC -> bundle)' $plan
  if (-not $preview) {
    New-Item -ItemType Directory -Path $BundleDir -Force | Out-Null
    Invoke-Plan $plan (Join-Path $BundleDir ".backup\$stamp")
    $m = [ordered]@{
      version = 1; ssdDrive = $canon; exportedAt = (Get-Date).ToString('o'); exportedFrom = $env:COMPUTERNAME
    }
    if ($index) { $m['app'] = [ordered]@{ acct = $index.Acct; org = $index.Org } }
    elseif ($manifest -and $manifest.PSObject.Properties['app']) { $m['app'] = $manifest.app }
    [IO.File]::WriteAllText($manifestPath, ($m | ConvertTo-Json -Depth 4), $utf8NoBom)
  }
}

if ($preview) { Write-Host "`nDry run, nothing was written." }
elseif ($wantImport) {
  Write-Host "`nDone. Open the Claude desktop app and the sessions should be in the sidebar." -ForegroundColor Green
  Write-Host 'CLI alternative: cd into the repo and run  claude --resume'
}
else { Write-Host "`nDone. Unplug the SSD and run  -Mode Import  on the other PC." -ForegroundColor Green }
