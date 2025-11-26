$ErrorActionPreference = "SilentlyContinue"
$smart = "$env:ProgramFiles\smartmontools\bin\smartctl.exe"
if (-not (Test-Path $smart)) { throw "smartctl not found at $smart" }

function Bridge-FromDType([string]$dtype){
  switch -Regex ($dtype) {
    '^sntjmicron' { 'JMicron USB-NVMe' }
    '^sntrealtek' { 'Realtek USB-NVMe' }
    '^sntasmedia' { 'ASMedia USB-NVMe' }
    '^sat(,?\d+)?$' { 'USB-SATA (SAT passthrough)' }
    '^scsi$'       { 'USB-SCSI (generic)' }
    '^nvme$'       { 'Direct NVMe' }
    default        { $dtype }
  }
}

function Get-DriverCandidates([string]$Dev,[string]$DTypeFromScan){
  $c = @()
  if($DTypeFromScan){ $c += $DTypeFromScan }
  if($Dev -match '^/dev/nvme\d+$'){
    $c = @('nvme')   # direct/internal NVMe
  } else {
    # USB: JMicron (NSIDs), then others
    $c += @(
      'sntjmicron,1','sntjmicron,0x1','sntjmicron,2','sntjmicron,0x2','sntjmicron',
      'sntrealtek','sntasmedia','sat,12','sat,16','sat','scsi'
    )
  }
  $seen=@{}; $out=@()
  foreach($d in $c){ if($d -and -not $seen.ContainsKey($d)){ $seen[$d]=$true; $out+=$d } }
  $out
}

function Get-HealthFromJson($j){
  $health = $null; $method = $null; $protocol = $j.protocol

  # NVMe path
  $pu = $j.nvme_smart_health_information_log.percentage_used
  if ($pu -ne $null) {
    try { $puNum = [int]$pu } catch { $puNum = $null }
    if ($puNum -ne $null -and $puNum -ge 0 -and $puNum -le 100) {
      $health = 100 - $puNum
      $method = "NVMe Percentage Used=$puNum%"
      if(-not $protocol){ $protocol = "NVMe" }
      return $health,$method,$protocol
    }
  }

  # SATA path
  $tbl = $j.ata_smart_attributes.table
  if ($tbl){
    $names = 'Percent_Lifetime_Remain','Percent_Life_Remaining','SSD_Life_Left','Remaining_Life','Media_Wearout_Indicator'
    $row = $tbl | Where-Object { $names -contains $_.name } | Select-Object -First 1
    if ($row){
      $val = $row.raw.value
      if($val -eq $null -or $val -lt 0 -or $val -gt 100){ $val = $row.value }
      if($val -ne $null){
        try { $valNum = [int]$val } catch { $valNum = $null }
        if($valNum -ne $null -and $valNum -ge 0 -and $valNum -le 100){
          $health = $valNum
          $method = $row.name
          if(-not $protocol){ $protocol = "ATA/SATA" }
          return $health,$method,$protocol
        }
      }
    }
    # Fallback: some vendors encode life in normalized Wear_Leveling_Count
    $wlc = $tbl | Where-Object { $_.name -match 'Wear_Leveling_Count' } | Select-Object -First 1
    if($wlc -and $wlc.value -ne $null){
      try { $wlcVal = [int]$wlc.value } catch { $wlcVal = $null }
      if($wlcVal -ne $null -and $wlcVal -ge 1 -and $wlcVal -le 100){
        if(-not $protocol){ $protocol = "ATA/SATA" }
        $health = $wlcVal
        $method = 'Wear_Leveling_Count (normalized; vendor-specific)'
        return $health,$method,$protocol
      }
    }
  }

  return $null,$null,$protocol
}

function Get-SSDHealthRows {
  # --- Scan & parse targets ---
  $scan = & $smart --scan-open 2>$null | Where-Object { $_ -and $_.Trim() -ne "" }
  if(-not $scan){ throw "smartctl --scan-open found no devices" }

  $targets=@()
  foreach($line in $scan){
    if($line -match '^(?<dev>/dev/nvme\d+)(?:\s+-d\s+(?<dtype>\S+))?'){
      $targets += [pscustomobject]@{ Dev=$matches['dev']; DType=$matches['dtype'] }
      continue
    }
    if($line -match '^(?<dev>\\\\\.\\PhysicalDrive\d+)\s+-d\s+(?<dtype>\S+)' ){
      $targets += [pscustomobject]@{ Dev=$matches['dev']; DType=$matches['dtype'] }
      continue
    }
    if($line -match '^(?<dev>/dev/sd\w+)\s+-d\s+(?<dtype>\S+)' ){
      $targets += [pscustomobject]@{ Dev=$matches['dev']; DType=$matches['dtype'] }
      continue
    }
  }

  # --- Probe & collect rows ---
  $rows=@()
  foreach($t in $targets){
    $worked=$null; $jobj=$null; $health=$null; $method=$null; $proto=$null
    $cands = Get-DriverCandidates $t.Dev $t.DType
    foreach($d in $cands){
      $jsonText = & $smart -a -j -d $d $t.Dev 2>$null
      if($jsonText -and $jsonText.Trim().StartsWith('{')){
        try { $jobj = $jsonText | ConvertFrom-Json } catch { $jobj = $null }
        if($jobj){
          $h,$m,$p = Get-HealthFromJson $jobj
          if($h -ne $null){ $health=$h; $method=$m; $proto=$p; $worked=$d; break }
          if(-not $worked){ $worked=$d } # remember first JSON parse for model/serial
        }
      }
    }

    $model = $null; $serial = $null
    if($jobj){
      $model = $jobj.model_name; if(-not $model){ $model = $jobj.device_model }
      $serial = $jobj.serial_number
    }

    $rows += [pscustomobject]@{
      Device        = $t.Dev
      'Detected -d' = $worked
      Bridge        = if($worked){ Bridge-FromDType $worked } else { if($t.DType){ Bridge-FromDType $t.DType } else { $null } }
      Protocol      = $proto
      Model         = $model
      Serial        = $serial
      'Health (%)'  = $health
      Method        = $method
    }
  }

  $rows | Sort-Object Device
}

function Get-HealthBarString([int]$Health){
  if($Health -lt 0){ $Health = 0 }
  if($Health -gt 100){ $Health = 100 }
  $segments = [int][Math]::Round($Health / 10.0)  # 0..10
  $full  = '#' * $segments
  $empty = '.' * (10 - $segments)
  return "[{0}{1}]" -f $full, $empty  # always fixed width: [##########]
}

$script:LastTableLineCount = 0

function Show-SSDHealthTable([object[]]$rows){
  $script:LastTableLineCount = 0

  if(-not $rows -or $rows.Count -eq 0){
    Write-Host "No disks reported by Windows or smartctl."
    $script:LastTableLineCount = 1
    return
  }

  # Define columns and getters (including Health Bar)
  $columns = @(
    @{ Name='Device';     Header='Device';      Getter={ param($r) [string]$r.Device } },
    @{ Name='Detected';   Header='Detected -d'; Getter={ param($r) [string]$r.'Detected -d' } },
    @{ Name='Bridge';     Header='Bridge';      Getter={ param($r) [string]$r.Bridge } },
    @{ Name='Protocol';   Header='Protocol';    Getter={ param($r) [string]$r.Protocol } },
    @{ Name='Model';      Header='Model';       Getter={ param($r) [string]$r.Model } },
    @{ Name='Serial';     Header='Serial';      Getter={ param($r) [string]$r.Serial } },
    @{ Name='HealthPct';  Header='Health (%)';  Getter={ param($r)
          if($r.'Health (%)' -ne $null){
            "{0,3}" -f [int]$r.'Health (%)'
          } else {
            "N/A"
          }
       } },
    @{ Name='HealthBar';  Header='Health Bar';  Getter={ param($r)
          if($r.'Health (%)' -ne $null){
            Get-HealthBarString -Health ([int]$r.'Health (%)')
          } else {
            Get-HealthBarString -Health 0
          }
       } },
    @{ Name='Method';     Header='Method';      Getter={ param($r) [string]$r.Method } }
  )

  # Compute widths per column so table is stable
  foreach($col in $columns){
    $maxLen = ($col.Header).Length
    foreach($r in $rows){
      $val = & $col.Getter $r
      if($val -eq $null){ $val = "" }
      $len = ($val).Length
      if($len -gt $maxLen){ $maxLen = $len }
    }
    $col.Width = $maxLen
  }

  # Find index of HealthBar column (for coloring)
  $barIndex = -1
  for($i=0; $i -lt $columns.Count; $i++){
    if($columns[$i].Name -eq 'HealthBar'){ $barIndex = $i; break }
  }

  # Header line
  $headerLine = ""
  foreach($col in $columns){
    $headerLine += $col.Header.PadRight($col.Width) + " "
  }
  Write-Host $headerLine
  $script:LastTableLineCount++

  # Separator
  $sepLine = ""
  foreach($col in $columns){
    $sepLine += ('-' * $col.Width) + " "
  }
  Write-Host $sepLine
  $script:LastTableLineCount++

  # Data rows
  foreach($r in $rows){
    for($i=0; $i -lt $columns.Count; $i++){
      $col = $columns[$i]
      $val = & $col.Getter $r
      if($val -eq $null){ $val = "" }
      $text = $val.PadRight($col.Width) + " "

      if($i -eq $barIndex){
        # Choose bar color by health
        $h = $r.'Health (%)'
        if($h -eq $null){
          $color = 'DarkGray'
        } elseif([int]$h -ge 80){
          $color = 'Green'
        } elseif([int]$h -ge 40){
          $color = 'Yellow'
        } else {
          $color = 'Red'
        }
        Write-Host -NoNewline $text -ForegroundColor $color
      } else {
        Write-Host -NoNewline $text
      }
    }
    Write-Host
    $script:LastTableLineCount++
  }
}

# -------- Main loop: continuous monitoring with Clear-Host dashboard --------

$intervalSeconds = 3  # set to 2 or 3 as you prefer

while ($true) {

  # 1) First do the heavy work while the OLD dashboard is still visible
  try {
    $rows = Get-SSDHealthRows
    $errorMessage = $null
  } catch {
    $rows = $null
    $errorMessage = $_.Exception.Message
  }

  # 2) Now clear and immediately paint the NEW dashboard
  Clear-Host

  $now = Get-Date
  Write-Host ("SSD Health Watch | {0} | Interval:{1}s" -f $now.ToString("yyyy-MM-dd HH:mm:ss"), $intervalSeconds)
  Write-Host ""

  if ($errorMessage) {
    Write-Host $errorMessage
  } else {
    Show-SSDHealthTable $rows
  }

  # 3) Wait for the next refresh
  Start-Sleep -Seconds $intervalSeconds
}
