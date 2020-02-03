# $TimestampServer は署名時に使用するタイムスタンプサーバー
$TimestampServer = "http://timestamp.globalsign.com/scripts/timstamp.dll"

function ShowUsage {
  [System.Console]::Error.WriteLine("updatevsto.ps1 <file.vsto>")
  [System.Console]::Error.WriteLine("  引数で渡したVSTOファイルおよびそれに依存するマニフェストファイルの")
  [System.Console]::Error.WriteLine("  コードサイニング署名を更新します")
  exit 1
}

function IsValidSignature($file) {
  $info = New-Object -TypeName System.Diagnostics.ProcessStartInfo
  $info.CreateNoWindow = $true
  $info.UseShellExecute = $false
  $info.RedirectStandardOutput = $true
  $info.FileName = $mage
  $info.Arguments = "-ver $file"
  $proc = New-Object -TypeName System.Diagnostics.Process
  $proc.StartInfo = $info
  $f = $proc.Start()
  $s = $proc.StandardOutput.ReadToEnd().Trim()
  $flag = $s -eq "Manifest has a valid signature."
  return $flag
}

if ($Args.Length -ne 1) {
  ShowUsage
}

$Cert=Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert 
if ($Cert -eq $null) {
  [System.Console]::Error.WriteLine("端末にコードサイニング証明書がインストールされていません")
  exit 1
}

# mage.exe のパスを探索
$mage = $null
$progdir = (get-item "Registry::HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion").GetValue("ProgramFilesDir (x86)")
foreach ($a in get-item "$progdir\Microsoft SDKs\Windows\*\bin\*\mage.exe" |Sort-Object -Propert LastWriteTime) {
  $mage = $a
}
if ($mage -eq $null) {
  [System.Console]::Error.WriteLine("mage.exeが見つかりません")
  exit 1
}
$thumb = $Cert.Thumbprint
[string]$signopt = "-ch $thumb -ti $TimestampServer"

$vsto = $Args[0]
$basedir = [System.IO.Path]::GetDirectoryName($vsto)
$InvalidVsto = $false

$xml = [xml](Get-Content($vsto))
foreach ($elem in $xml.GetElementsByTagName("dependentAssembly")) {
  $manifest = $elem.Attributes["codebase"].Value
  $manpath = [System.io.Path]::Combine($basedir, $manifest)
  $files = @()
  [bool]$InvalidManifest = $false
  foreach ($elem2 in $elem.ChildNodes) {
    if ($elem2.LocalName -eq "assemblyIdentity") {
      $file = $elem2.Attributes["name"].Value
      $path = [System.IO.Path]::Combine($basedir, $file)
      $sign = Get-AuthenticodeSignature -FilePath $path
      if ($sign -eq $null -or $sign.Status -ne "Valid") {
        $files += $path
        $InvalidManifest = $true
        $InvalidVsto = $true
      } else {
        if ((Get-Item $manpath).LastWriteTime -lt (Get-Item $path).LastWriteTime) {
          $InvalidManifest = $true
          $InvalidVsto = $true
        }
        Write-Output ($file + " は署名済みです。スキップします。")
      }
    }
  }
  if ($files.Length -ne 0) {
    Set-AuthenticodeSignature -Certificate $Cert -Filepath $Files -TimestampServer $TimestampServer
  }
  if (-not $InvalidManifest -and -not (IsValidSignature($manpath))) {
    $InvalidManifest = $true
    $InvalidVsto = $true
  }
  if ($InvalidManifest) {
    Start-Process -FilePath "$mage" -ArgumentList "-u ""$manpath"" $signopt" -Wait -NoNewWindow
  } else {
    Write-Output ($manifest + " は署名済みです。スキップします。")
  }
  if ((Get-Item $vsto).LastWriteTime -lt (Get-Item $manpath).LastWriteTime) {
    $InvalidVsto = $true
  }
  if (-not $InvalidVsto -and -not (IsValidSignature($vsto))) {
    $InvalidVsto = $true
  }
  if ($InvalidVsto) {
    Start-Process -FilePath "$mage" -ArgumentList "-u ""$vsto"" -appm ""$manpath"" $signopt" -Wait -NoNewWindow
  } else {
    Write-Output ($vsto + " は署名済みです。スキップします。")
  }
}
