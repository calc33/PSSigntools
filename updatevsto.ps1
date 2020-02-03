# $TimestampServer は署名時に使用するタイムスタンプサーバー
$TimestampServer = "http://timestamp.globalsign.com/scripts/timstamp.dll"

function ShowUsage {
  [System.Console]::Error.WriteLine("updatevsto.ps1 <file.vsto>")
  [System.Console]::Error.WriteLine("  引数で渡したVSTOファイルおよびそれに依存するマニフェストファイルの")
  [System.Console]::Error.WriteLine("  コードサイニング署名を更新します")
  exit 1
}

function IsValidSignature($File) {
  $info = New-Object -TypeName System.Diagnostics.ProcessStartInfo
  $info.CreateNoWindow = $true
  $info.UseShellExecute = $false
  $info.RedirectStandardOutput = $true
  $info.FileName = $mage
  $info.Arguments = "-ver $File"
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

#コードサイニング証明書関連の情報を取得
$Cert=Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert 
if ($Cert -eq $null) {
  [System.Console]::Error.WriteLine("端末にコードサイニング証明書がインストールされていません")
  exit 1
}
$thumb = $Cert.Thumbprint
[string]$signopt = "-ch $thumb -ti $TimestampServer"

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

$vsto = $Args[0]
$baseDir = [System.IO.Path]::GetDirectoryName($vsto)
$invalidVsto = $false

$xml = [xml](Get-Content($vsto))
foreach ($elem in $xml.GetElementsByTagName("dependentAssembly")) {
  $manifest = $elem.Attributes["codebase"].Value
  $manPath = [System.IO.Path]::Combine($baseDir, $manifest)
  $files = @()
  [bool]$invalidManifest = $false
  foreach ($elem2 in $elem.ChildNodes) {
    if ($elem2.LocalName -eq "assemblyIdentity") {
      $file = $elem2.Attributes["name"].Value
      $path = [System.IO.Path]::Combine($baseDir, $file)
      $sign = Get-AuthenticodeSignature -FilePath $path
      if ($sign -eq $null -or $sign.Status -ne "Valid") {
        $files += $path
        $invalidManifest = $true
        $invalidVsto = $true
      } else {
        if ((Get-Item $manPath).LastWriteTime -lt (Get-Item $path).LastWriteTime) {
          $invalidManifest = $true
          $invalidVsto = $true
        }
        Write-Output ($file + " は署名済みです。スキップします。")
      }
    }
  }
  if ($files.Length -ne 0) {
    Set-AuthenticodeSignature -Certificate $Cert -Filepath $Files -TimestampServer $TimestampServer
  }
  if (-not $invalidManifest -and -not (IsValidSignature($manPath))) {
    $invalidManifest = $true
    $invalidVsto = $true
  }
  if ($invalidManifest) {
    Start-Process -FilePath "$mage" -ArgumentList "-u ""$manPath"" $signopt" -Wait -NoNewWindow
  } else {
    Write-Output ($manifest + " は署名済みです。スキップします。")
  }
  if ((Get-Item $vsto).LastWriteTime -lt (Get-Item $manPath).LastWriteTime) {
    $invalidVsto = $true
  }
  if (-not $invalidVsto -and -not (IsValidSignature($vsto))) {
    $invalidVsto = $true
  }
  if ($invalidVsto) {
    Start-Process -FilePath "$mage" -ArgumentList "-u ""$vsto"" -appm ""$manPath"" $signopt" -Wait -NoNewWindow
  } else {
    Write-Output ($vsto + " は署名済みです。スキップします。")
  }
}
