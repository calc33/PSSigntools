# $TimestampServer は署名時に使用するタイムスタンプサーバー
$TimestampServer = "http://timestamp.globalsign.com/scripts/timstamp.dll"

function ShowUsage {
  [System.Console]::Error.WriteLine("updatevsto.ps1 <file1.vsto> <file2.manifest> [files ...]")
  [System.Console]::Error.WriteLine("  3番目以降の引数で渡したファイル(exe,dll等)を署名し")
  [System.Console]::Error.WriteLine("  マニフェストファイルおよびVSTOファイルを署名付きで更新します")
  exit 1
}

if ($Args.Length -lt 2) {
  ShowUsage
}

[bool]$Found = $False
$Files = @()
$MageFiles = @()
for ([int]$i = 2; $i -lt $Args.Length; $i++) {
  $a = $Args[$i]
  $Paths = Resolve-Path -Path $a -Relative
  foreach ($f in $Paths) {
    if ((Get-Item -Path $f).PSIsContainer) {
      continue
    }
    $Sign = Get-AuthenticodeSignature -FilePath $f
    if ($Force -or $Sign -eq $null -or $Sign.Status -ne "Valid") {
      $Files += $f
    } else {
      Write-Output ($f + " は署名済みです。スキップします。")
    }
  }
  $Found = $True
}

$Cert=Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert 
if ($Cert -eq $null) {
    [System.Console]::Error.WriteLine("端末にコードサイニング証明書がインストールされていません")
    exit 0
}

if ($Files.Length -ne 0) {
  Set-AuthenticodeSignature -Certificate $Cert -Filepath $Files -TimestampServer $TimestampServer
}

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
$vsto = $Args[0]
$manifest = $Args[1]

Start-Process -FilePath "$mage" -ArgumentList "-u $manifest -ch $thumb -ti $TimestampServer" -Wait -NoNewWindow
Start-Process -FilePath "$mage" -ArgumentList "-u $vsto -appm $manifest -ch $thumb -ti $TimestampServer" -Wait -NoNewWindow
