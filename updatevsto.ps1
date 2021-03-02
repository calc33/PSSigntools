# $TimestampServer は署名時に使用するタイムスタンプサーバー
$TimestampServer = "http://timestamp.digicert.com"

function ShowUsage {
    if ($PSUICulture -eq "ja-JP") {
        [System.Console]::Error.WriteLine("updatevsto.ps1 <file.vsto>")
        [System.Console]::Error.WriteLine("  引数で渡したVSTOファイルおよびそれに依存するマニフェストファイルの")
        [System.Console]::Error.WriteLine("  コードサイニング署名を更新します")
    } else {
        [System.Console]::Error.WriteLine("updatevsto.ps1 <file.vsto>")
        [System.Console]::Error.WriteLine("  Update Code-Signing-Certificate of the .vsto-file passed as argument"
        [System.Console]::Error.WriteLine("  and dependency files.")
    }
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
    if ($PSUICulture -eq "ja-JP") {
        [System.Console]::Error.WriteLine("端末にコードサイニング証明書がインストールされていません")
    } else {
        [System.Console]::Error.WriteLine("Code-Signing-Certificate is not found.")
    }
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
    if ($PSUICulture -eq "ja-JP") {
        [System.Console]::Error.WriteLine("mage.exeが見つかりません")
    } else {
        [System.Console]::Error.WriteLine("mage.exe is not found.")
    }
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
                if ($PSUICulture -eq "ja-JP") {
                    Write-Output ($file + " は署名済みです。スキップします。")
                } else {
                    Write-Output ($file + " is already singed. Skip.")
                }
            }
        }
    }
    if ($files.Length -ne 0) {
        Set-AuthenticodeSignature -Certificate $Cert -Filepath $Files -HashAlgorithm "SHA256" -TimestampServer $TimestampServer
    }
    if (-not $invalidManifest -and -not (IsValidSignature($manPath))) {
        $invalidManifest = $true
        $invalidVsto = $true
    }
    if ($invalidManifest) {
        Start-Process -FilePath "$mage" -ArgumentList "-u ""$manPath"" $signopt -a sha256RSA -ti $TimestampServer" -Wait -NoNewWindow
    } else {
        if ($PSUICulture -eq "ja-JP") {
            Write-Output ($manifest + " は署名済みです。スキップします。")
        } else {
            Write-Output ($manifest + " is already singed. Skip.")
        }
    }
    if ((Get-Item $vsto).LastWriteTime -lt (Get-Item $manPath).LastWriteTime) {
        $invalidVsto = $true
    }
    if (-not $invalidVsto -and -not (IsValidSignature($vsto))) {
        $invalidVsto = $true
    }
    if ($invalidVsto) {
        Start-Process -FilePath "$mage" -ArgumentList "-u ""$vsto"" -appm ""$manPath"" $signopt -a sha256RSA -ti $TimestampServer" -Wait -NoNewWindow
    } else {
        if ($PSUICulture -eq "ja-JP") {
            Write-Output ($vsto + " は署名済みです。スキップします。")
        } else {
            Write-Output ($vsto + " is already singed. Skip.")
        }
    }
}
