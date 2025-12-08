# Offline hostname editor
# - Purpose:
#   This script was created to simulate and reproduce, in an offline SSD workflow,
#   the hostname and DNS configuration changes that normally happen online in a
#   specific lab environment. In that lab, target machines use hostnames with a
#   suffix such as "_TA" and are paired to "host" systems.
#   After analyzing the original online scripts, I identified which files and
#   registry locations are responsible for those changes and built this offline
#   variant:
#   - It edits Imaging\host.txt and Imaging\hostfqdn.txt on the offline SSD.
#   - It mounts and updates EFI\host.txt on the ESP partition.
#   - It loads the offline Windows SYSTEM and SOFTWARE hives and updates the
#     relevant hostname and DNS-related registry entries.
#
# - Behavior:
#   - ALL-CAPS hostname + optional ALL-CAPS suffix (_TA, etc.) in REGISTRY only.
#   - Imaging\host.txt         : bare HOSTNAME only (no suffix, no DNS).
#   - Imaging\hostfqdn.txt     : HOSTNAME.FQDN-SUFFIX (true FQDN, suffix from config).
#   - EFI \host.txt            : HOSTNAME.FQDN-SUFFIX (same FQDN as Imaging\hostfqdn.txt).
#   - EFI access: prefers already-mounted ESP on the same disk; else uses mountvol to map (prefers M:).
#   - Robust volume GUID match + readiness wait + provider priming + safe read/write.
#   - Updates all ControlSets (Tcpip/Tcpip6/ComputerName).
#   - Mirrors SOFTWARE ComputerName (with safe unload).
#   - Imaging\pair.json is shown in summary then removed.
#   - SSD info, friendly summaries (incl. EFI prev/new/path), end pause; always unmounts what it mounted.
#   - This is beta version. For inquiries contact the author: Avraham Makovsky
#
# === CONFIG ===
# DNS suffix to append to Imaging\hostfqdn.txt (FQDN) and EFI host.txt.
# Example: "iil.company.com"
$HostFqdnSuffixConfig = "iil.company.com"   # <-- set this to your desired suffix (without leading dot)

function Pause-IfConsole {
  if ($Host -and $Host.Name -eq 'ConsoleHost') {
    Write-Host ""
    Read-Host "Press Enter to exit..."
  }
}

# 0) Require admin
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
  Write-Error "Please run PowerShell as Administrator."
  Pause-IfConsole
  return
}

# Require a non-empty suffix for proper FQDN
if ([string]::IsNullOrWhiteSpace($HostFqdnSuffixConfig)) {
  Write-Error "HostFqdnSuffixConfig is empty. Set a DNS suffix (e.g. 'iil.company.com') in the script header."
  Pause-IfConsole
  return
}

# === UI helpers ===
function Section($title){ Write-Host "`n=== $title ===" -ForegroundColor Cyan }
function Field($label,$value){ Write-Host ("  {0}: {1}" -f $label,$value) }
function Format-Size($b){
  if($b -ge 1TB){ '{0:N2} TB' -f ($b/1TB) } else { '{0:N2} GB' -f ($b/1GB) }
}
function Unload-OfflineHive {
  try { reg unload "HKLM\Offline" 1>$null 2>$null | Out-Null } catch {}
}

# Disk mapping helpers
function Get-DiskInfoForDrive($driveLetter){
  $driveLetter = ($driveLetter.ToString())[0].ToString().ToUpper()
  $results=@()
  try{
    $parts=Get-Partition -DriveLetter $driveLetter -ErrorAction Stop
    foreach($p in $parts){
      $d=Get-Disk -Number $p.DiskNumber -ErrorAction Stop
      $results += [pscustomobject]@{
        Index=$d.Number; Model=$d.FriendlyName; Size=[int64]$d.Size;
        SerialNumber=$d.SerialNumber; FirmwareRevision=$d.FirmwareVersion; Mapped=$true
      }
    }
  }catch{}
  if($results.Count -gt 0){ return ,($results|Sort-Object Index -Unique) }

  try{
    $ld=Get-CimInstance Win32_LogicalDisk -Filter ("DeviceID='{0}:'" -f $driveLetter)
    if($ld){
      $parts=Get-CimInstance -Query ("ASSOCIATORS OF {{Win32_LogicalDisk.DeviceID='{0}:'}} WHERE AssocClass=Win32_LogicalDiskToPartition" -f $driveLetter)
      foreach($p in $parts){
        $q="ASSOCIATORS OF {Win32_DiskPartition.DeviceID='$($p.DeviceID)'} WHERE AssocClass=Win32_DiskDriveToDiskPartition"
        $dds=Get-CimInstance -Query $q
        foreach($dd in $dds){
          $results += [pscustomobject]@{
            Index=$dd.Index; Model=$dd.Model; Size=[int64]$dd.Size;
            SerialNumber=$dd.SerialNumber; FirmwareRevision=$dd.FirmwareRevision; Mapped=$true
          }
        }
      }
    }
  }catch{}
  if($results.Count -gt 0){ return ,($results|Sort-Object Index -Unique) }

  try{
    $all=Get-CimInstance Win32_DiskDrive
    return ,($all|Sort-Object Index|ForEach-Object{
      [pscustomobject]@{
        Index=$_.Index;Model=$_.Model;Size=[int64]$_.Size;
        SerialNumber=$_.SerialNumber;FirmwareRevision=$_.FirmwareRevision;Mapped=$false
      }
    })
  }catch{ return @() }
}

# Find free letter (D..Z)
function Get-FreeDriveLetter {
  $used = (Get-Volume -ErrorAction SilentlyContinue | Where-Object DriveLetter).DriveLetter
  foreach ($c in [char[]]([byte][char]'D'..[byte][char]'Z')) {
    if ($used -notcontains $c) { return $c }
  }
  return $null
}

# Normalize volume GUID path for comparison
function Normalize-VolPath {
  param([string]$p)
  if(-not $p){ return $null }
  return (($p -join '')).Trim() -replace '\\+$',''
}

# Resolve the EFI Volume GUID path (\\?\Volume{GUID}\) on a given disk
function Resolve-EfiVolumePathForDisk {
  param([Parameter(Mandatory)][int]$DiskNumber)
  $efiGuid = '{C12A7328-F81F-11D2-BA4B-00A0C93EC93B}'
  $part = Get-Partition -DiskNumber $DiskNumber -ErrorAction Stop |
          Where-Object { $_.GptType -eq $efiGuid -or $_.Type -eq 'EFI System' } |
          Select-Object -First 1
  if (-not $part) { return $null }

  $vol = $part | Get-Volume -ErrorAction SilentlyContinue
  if ($vol -and $vol.Path) { return $vol.Path }

  $volCim = Get-CimInstance -Namespace root/microsoft/windows/storage -Class MSFT_Volume -ErrorAction SilentlyContinue |
            Where-Object { $_.ObjectId -like "*DiskNumber=$($part.DiskNumber)*PartitionNumber=$($part.PartitionNumber)*" } |
            Select-Object -First 1
  if ($volCim -and $volCim.Path) { return $volCim.Path }
  return $null
}

# If ESP is already mounted on this disk, return that drive letter
function Try-GetExistingEfiLetterForDisk {
  param([Parameter(Mandatory)][int]$DiskNumber)
  $efiGuid = '{C12A7328-F81F-11D2-BA4B-00A0C93EC93B}'
  $letters = (Get-Volume -ErrorAction SilentlyContinue | Where-Object DriveLetter).DriveLetter
  foreach($L in $letters){
    try {
      $p = Get-Partition -DriveLetter $L -ErrorAction Stop | Select-Object -First 1
      if ($p -and $p.DiskNumber -eq $DiskNumber -and ($p.GptType -eq $efiGuid -or $p.Type -eq 'EFI System')) {
        return [char]$L
      }
    } catch {}
  }
  return $null
}

# Mount a Volume GUID path to a letter (prefers M:, falls back to next free letter).
# Returns: @{Success; Letter; MountedTemp; Error}
function Mount-VolumePathToLetter {
  param(
    [Parameter(Mandatory)][string]$VolumePath,
    [char]$PreferredLetter = 'M'
  )
  $out = [pscustomobject]@{
    Success=$false; Letter=$null; MountedTemp=$false; Error=$null
  }

  try {
    # If preferred already mapped & accessible, just use it
    $exists = Get-Volume -ErrorAction SilentlyContinue | Where-Object DriveLetter -eq $PreferredLetter
    if ($exists -and (Test-Path -LiteralPath "$($PreferredLetter):\")) {
      $out.Success = $true; $out.Letter = $PreferredLetter; $out.MountedTemp = $false
      return $out
    }

    # Otherwise, compare GUIDs robustly
    if ($exists) {
      $cur = (& mountvol "$($PreferredLetter):" /L) 2>$null
      $curN = Normalize-VolPath $cur
      $want = Normalize-VolPath $VolumePath
      if ($curN -and $want -and ($curN.ToLower() -eq $want.ToLower())) {
        $out.Success = $true; $out.Letter = $PreferredLetter; $out.MountedTemp = $false
        return $out
      } else {
        $alt = Get-FreeDriveLetter
        if (-not $alt) { $out.Error = "No free drive letter available."; return $out }
        (& mountvol "$($alt):" $VolumePath) | Out-Null
        $out.Success = $true; $out.Letter = $alt; $out.MountedTemp = $true
        return $out
      }
    }

    # Preferred free, try assign
    (& mountvol "$($PreferredLetter):" $VolumePath) 2>$null | Out-Null
    $chk  = Normalize-VolPath ((& mountvol "$($PreferredLetter):" /L) 2>$null)
    $want = Normalize-VolPath $VolumePath
    if ($chk -and $want -and ($chk.ToLower() -eq $want.ToLower())) {
      $out.Success = $true; $out.Letter = $PreferredLetter; $out.MountedTemp = $true
      return $out
    }

    # Fallback to another free letter
    $alt = Get-FreeDriveLetter
    if (-not $alt) { $out.Error = "No free drive letter available."; return $out }
    (& mountvol "$($alt):" $VolumePath) 2>$null | Out-Null
    $chk2 = Normalize-VolPath ((& mountvol "$($alt):" /L) 2>$null)
    if ($chk2 -and $want -and ($chk2.ToLower() -eq $want.ToLower())) {
      $out.Success = $true; $out.Letter = $alt; $out.MountedTemp = $true
      return $out
    }

    # Last-ditch: if the new letter is accessible, accept it
    if (Test-Path -LiteralPath "$($PreferredLetter):\")) {
      $out.Success = $true; $out.Letter = $PreferredLetter; $out.MountedTemp = $true
      return $out
    }
    if ($alt -and (Test-Path -LiteralPath "$($alt):\")) {
      $out.Success = $true; $out.Letter = $alt; $out.MountedTemp = $true
      return $out
    }

    $out.Error = "Failed to mount $VolumePath to a drive letter."
    return $out
  } catch {
    $out.Error = $_.Exception.Message
    return $out
  }
}

# Unmount drive letter if we mounted it in this session
function Unmount-IfTemp {
  param([Parameter(Mandatory)][pscustomobject]$Handle)
  try {
    if ($Handle.MountedTemp -and $Handle.Letter) {
      (& mountvol "$($Handle.Letter):" /D) 2>$null | Out-Null
    }
  } catch {}
}

# === EFI readiness & safe IO helpers ===
function Prime-Volume {
  param([Parameter(Mandatory)][char]$Letter)
  try { cmd /c "dir $($Letter):\ /a" | Out-Null } catch {}
  try { Get-ChildItem -LiteralPath "$($Letter):\" -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
  Start-Sleep -Milliseconds 300
}
function Wait-VolumeReady {
  param([Parameter(Mandatory)][char]$Letter, [int]$TimeoutMs = 10000)
  $deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMs)
  $root = "$($Letter):\"
  do {
    try {
      if (Test-Path -LiteralPath $root) {
        Get-ChildItem -LiteralPath $root -Force -ErrorAction SilentlyContinue | Out-Null
        return $true
      }
    } catch {}
    Start-Sleep -Milliseconds 120
  } while ([DateTime]::UtcNow -lt $deadline)
  return $false
}
function Resolve-HostTxtPath {
  param([Parameter(Mandatory)][char]$Letter)
  $root = "$($Letter):\"
  $default = Join-Path $root 'host.txt'
  if (Test-Path -LiteralPath $default) { return $default }
  try {
    $hit = Get-ChildItem -LiteralPath $root -Force -File -ErrorAction Stop |
           Where-Object { $_.Name -ieq 'host.txt' } |
           Select-Object -First 1
    if ($hit) { return $hit.FullName }
  } catch {}
  return $default
}
function Read-HostTxtSafe {
  param([Parameter(Mandatory)][string]$Path)
  try { return (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop) } catch {
    try { return [System.IO.File]::ReadAllText($Path) } catch {
      $out = (cmd /c "type `"$Path`"") 2>$null
      if ($LASTEXITCODE -eq 0 -and $out) { return ($out -join "`r`n") }
    }
  }
  return $null
}
function Write-HostTxtSafe {
  param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$Content
  )
  try {
    if (-not (Test-Path -LiteralPath $Path)) {
      New-Item -ItemType File -Path $Path -Force -ErrorAction SilentlyContinue | Out-Null
    }
    Set-Content -LiteralPath $Path -Value $Content -Encoding ASCII -NoNewline
    return $true
  } catch {
    try {
      [System.IO.File]::WriteAllText($Path, $Content, [System.Text.Encoding]::ASCII)
      return $true
    } catch {
      return $false
    }
  }
}

function Process-Drive($drive){
  $mountKey="HKLM\Offline"
  $winDir="$drive`:\Windows"
  $hivePath=Join-Path $winDir "System32\config\SYSTEM"

  $summary=[PSCustomObject]@{
    Drive="$drive`:"; RegHostname=$null; ControlSetsUpdated=0
    HostTxt=$null; HostTxtPrev=$null
    HostFqdnUpdated=$false; HostFqdnSkipped=$false; HostFqdnValue=$null; HostFqdnPrev=$null; HostFqdnSuffixUsed=$null
    PairRemoved=$false; PairMissing=$false; PairJsonPreview=$null
    PrevTcpipHostname=$null; PrevComputerName=$null; PrevActiveComputerName=$null
    PrevEfiHost=$null; EfiHostPath=$null; EfiHostUpdated=$false; EfiHostSkipped=$false; EfiHostNew=$null
    Status="OK"
  }

  if (-not (Test-Path $hivePath)) {
    Section "Target $drive`: (skipped)"
    Field "Reason" "SYSTEM hive not found"
    $summary.Status="Skipped - no hive"
    return $summary
  }
  if ("$winDir\" -ieq "$env:SystemRoot\") {
    Section "Target $drive`: (skipped)"
    Field "Reason" "Points to LIVE OS"
    $summary.Status="Skipped - live OS"
    return $summary
  }

  # Load hive (with copy fallback)
  Unload-OfflineHive
  $mounted=$false; $usingCopy=$false
  $tempDir=Join-Path $env:TEMP "OfflineHive_Edit_$(Get-Date -Format yyyyMMdd_HHmmss)"
  $tempHive=Join-Path $tempDir "SYSTEM"
  Write-Host "Loading offline SYSTEM hive from $hivePath ..."
  $loadOK=$true
  try{ reg load $mountKey $hivePath | Out-Null }catch{ $loadOK=$false }
  if(-not $loadOK){
    Write-Warning "Hive locked or in use. Will try working on a copy..."
    try {
      New-Item -ItemType Directory -Force -Path $tempDir | Out-Null
      Copy-Item -Path $hivePath -Destination $tempHive -Force
      Unload-OfflineHive
      reg load $mountKey $tempHive | Out-Null
      $usingCopy=$true; $loadOK=$true
    } catch {
      Section "Target $drive`: (error)"
      Field "Error" "Failed to load hive even via copy: $($_.Exception.Message)"
      $summary.Status="Error - load failed"
      return $summary
    }
  }
  $mounted=$true

  try {
    if(-not (Test-Path "Registry::$mountKey\Select")){ throw "Missing Select key in offline hive." }
    $select=Get-ItemProperty -Path "Registry::$mountKey\Select"
    $currentNum=[int]$select.Current
    if($currentNum -lt 1){ throw "Invalid Select\Current value." }
    $currentCS="ControlSet{0:000}" -f $currentNum
    $tcpipPath="Registry::$mountKey\$currentCS\Services\Tcpip\Parameters"
    if(-not (Test-Path $tcpipPath)){ throw "Tcpip Parameters not found at $tcpipPath." }

    # Current values
    $p=Get-ItemProperty -Path $tcpipPath -ErrorAction SilentlyContinue
    $curHostname=$p.'Hostname'
    if([string]::IsNullOrWhiteSpace($curHostname)){$curHostname=$p.'NV Hostname'}
    if([string]::IsNullOrWhiteSpace($curHostname)){$curHostname="<not set>"}
    $summary.PrevTcpipHostname=$curHostname
    $cnBaseCS="Registry::$mountKey\$currentCS\Control\ComputerName"
    $prevCN=(Get-ItemProperty -Path "$cnBaseCS\ComputerName" -ErrorAction SilentlyContinue).'ComputerName'
    $prevACN=(Get-ItemProperty -Path "$cnBaseCS\ActiveComputerName" -ErrorAction SilentlyContinue).'ComputerName'
    if([string]::IsNullOrWhiteSpace($prevCN)){$prevCN="<not set>"}
    if([string]::IsNullOrWhiteSpace($prevACN)){$prevACN="<not set>"}
    $summary.PrevComputerName=$prevCN
    $summary.PrevActiveComputerName=$prevACN

    # Imaging + suffix discovery (suffix used for registry only)
    $imagingDir   = Join-Path "$drive`:" "Imaging"
    $imagingFile  = Join-Path $imagingDir "host.txt"
    $hostFqdnFile = Join-Path $imagingDir "hostfqdn.txt"
    $fileHostname = $null
    $fqdnHostname = $null
    $domainSuffix = $null

    if(Test-Path $imagingFile){
      try{ $fileHostname = (Get-Content -Path $imagingFile -Raw -ErrorAction Stop).Trim() }
      catch{ $fileHostname = "<read error>" }
    }
    if(Test-Path $hostFqdnFile){
      try{ $fqdnHostname = (Get-Content -Path $hostFqdnFile -Raw -ErrorAction Stop).Trim() }
      catch{ $fqdnHostname = "<read error>" }
    }

    $summary.HostTxtPrev  = $(if($fileHostname){$fileHostname}else{"<missing>"})
    $summary.HostFqdnPrev = $(if($fqdnHostname){$fqdnHostname}else{"<missing>"})

    if($fileHostname -and $fileHostname -ne "<read error>"){
      $i=$fileHostname.IndexOf('.')
      if($i -ge 0 -and $i -lt ($fileHostname.Length-1)){
        $domainSuffix=$fileHostname.Substring($i+1).Trim(' .')
      }
    }
    if(-not $domainSuffix){
      $regDomain=$p.'NV Domain'
      if([string]::IsNullOrWhiteSpace($regDomain)){$regDomain=$p.'Domain'}
      if([string]::IsNullOrWhiteSpace($regDomain)){$regDomain=$p.'DhcpDomain'}
      if(-not [string]::IsNullOrWhiteSpace($regDomain)){
        $domainSuffix=$regDomain.Trim(' .')
      }
    }

    # hostfqdn/EFI suffix to use (from script-level config) – must be non-empty
    $hostFqdnSuffix = $HostFqdnSuffixConfig.Trim().Trim('.')
    if ([string]::IsNullOrWhiteSpace($hostFqdnSuffix)) {
      throw "HostFqdnSuffixConfig resolved to empty suffix after trim. Fix the script header."
    }

    # Imaging\pair.json (context before we delete it later)
    $pairFile    = Join-Path $imagingDir "pair.json"
    $pairContent = $null
    if (Test-Path $pairFile) {
      try {
        $pairContent = Get-Content -LiteralPath $pairFile -Raw -ErrorAction Stop
        $previewLen = 400
        if ($pairContent.Length -gt $previewLen) {
          $summary.PairJsonPreview = $pairContent.Substring(0,$previewLen) + " ..."
        } else {
          $summary.PairJsonPreview = $pairContent
        }
      } catch {
        $pairContent = "<read error: $($_.Exception.Message)>"
        $summary.PairJsonPreview = $pairContent
      }
    }

    # ==== EFI host.txt probe (previous) ====
    $mount = $null
    try {
      $partWin = Get-Partition -DriveLetter $drive -ErrorAction Stop | Select-Object -First 1
      $disk    = Get-Disk -Number $partWin.DiskNumber -ErrorAction Stop

      $existing = Try-GetExistingEfiLetterForDisk -DiskNumber $disk.Number
      if ($existing) {
        $mount = [pscustomobject]@{
          Success=$true; Letter=$existing; MountedTemp=$false; Error=$null
        }
      } else {
        $efiPath = Resolve-EfiVolumePathForDisk -DiskNumber $disk.Number
        if ($efiPath) {
          $mount = Mount-VolumePathToLetter -VolumePath $efiPath -PreferredLetter 'M'
        }
      }

      if ($mount -and $mount.Success) {
        [void](Wait-VolumeReady -Letter $mount.Letter); Prime-Volume -Letter $mount.Letter
        $hostPath = Resolve-HostTxtPath -Letter $mount.Letter
        $summary.EfiHostPath = $hostPath
        $prev = Read-HostTxtSafe -Path $hostPath
        if ($prev) { $summary.PrevEfiHost = $prev.Trim() } else { $summary.PrevEfiHost = "<missing>" }
      } else {
        if (Test-Path -LiteralPath "M:\") {
          [void](Wait-VolumeReady -Letter 'M'); Prime-Volume -Letter 'M'
          $hostPath = Resolve-HostTxtPath -Letter 'M'
          $summary.EfiHostPath = $hostPath
          $prev = Read-HostTxtSafe -Path $hostPath
          if ($prev) { $summary.PrevEfiHost = $prev.Trim() } else { $summary.PrevEfiHost = "<missing>" }
          $mount = [pscustomobject]@{
            Success=$true; Letter='M'; MountedTemp=$false; Error=$null
          }
        } else {
          $summary.EfiHostSkipped = $true
          $summary.EfiHostPath    = "<not mounted>"
        }
      }
    } finally {
      if ($mount) { Unmount-IfTemp -Handle $mount }
    }

    # ===== Initial per-drive state =====
    Section "Target $drive`:"
    Field "Hive path" $hivePath
    Field "Active ControlSet" $currentCS
    Field "Registry Hostname (TCP/IP)" $curHostname
    Field "ComputerName (current set)" $prevCN
    Field "ActiveComputerName (current set)" $prevACN
    Field "Imaging\host.txt"    ($(if($fileHostname){$fileHostname}else{"<missing>"}))
    Field "Imaging\hostfqdn.txt"($(if($fqdnHostname){$fqdnHostname}else{"<missing>"}))
    if ($pairContent) {
      Field "Imaging\pair.json" "Present (see below)"
      Write-Host "  Imaging\pair.json contents:" -ForegroundColor DarkGray
      $pairContent | Write-Host
    } else {
      Field "Imaging\pair.json" "<missing>"
