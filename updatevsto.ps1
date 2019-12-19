# $TimestampServer �͏������Ɏg�p����^�C���X�^���v�T�[�o�[
$TimestampServer = "http://timestamp.globalsign.com/scripts/timstamp.dll"

function ShowUsage {
  [System.Console]::Error.WriteLine("updatevsto.ps1 <file1.vsto> <file2.manifest> [files ...]")
  [System.Console]::Error.WriteLine("  3�Ԗڈȍ~�̈����œn�����t�@�C��(exe,dll��)��������")
  [System.Console]::Error.WriteLine("  �}�j�t�F�X�g�t�@�C�������VSTO�t�@�C���������t���ōX�V���܂�")
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
      Write-Output ($f + " �͏����ς݂ł��B�X�L�b�v���܂��B")
    }
  }
  $Found = $True
}

$Cert=Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert 
if ($Cert -eq $null) {
    [System.Console]::Error.WriteLine("�[���ɃR�[�h�T�C�j���O�ؖ������C���X�g�[������Ă��܂���")
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
  [System.Console]::Error.WriteLine("mage.exe��������܂���")
  exit 1
}
$thumb = $Cert.Thumbprint
$vsto = $Args[0]
$manifest = $Args[1]

Start-Process -FilePath "$mage" -ArgumentList "-u $manifest -ch $thumb -ti $TimestampServer" -Wait -NoNewWindow
Start-Process -FilePath "$mage" -ArgumentList "-u $vsto -appm $manifest -ch $thumb -ti $TimestampServer" -Wait -NoNewWindow
