# $TimestampServer は署名時に使用するタイムスタンプサーバー
$TimestampServer = "http://timestamp.digicert.com"

function ShowUsage {
    if ($PSUICulture -eq "ja-JP") {
        [System.Console]::Error.WriteLine("sign.ps1 [/F] <ファイル1> [<ファイル2>...]")
        [System.Console]::Error.WriteLine("  引数で渡したファイルに対してコードサイニング署名します")
        [System.Console]::Error.WriteLine("  既に有効な署名が設定されているファイルは署名しません")
        [System.Console]::Error.WriteLine("  ファイルは複数指定可能です")
        [System.Console]::Error.WriteLine("  実行形式(.exe/.dll)およびPowerShellScript(.ps1)のファイルが指定可能です")
        [System.Console]::Error.WriteLine("  /F を指定すると既に有効な署名が設定されているファイルも再署名します")
    } else {
        [System.Console]::Error.WriteLine("sign.ps1 [/F] <file1> [<file2>...]")
        [System.Console]::Error.WriteLine("  Sign the file passed as arguments using Code-Signing-Certificate.")
        [System.Console]::Error.WriteLine("  Skip if a file has already valid signature.")
        [System.Console]::Error.WriteLine("  Argument files can select executable(.exe/.dll) or PowerShellScript(.ps1).")
        [System.Console]::Error.WriteLine("  /F  Sign all files include having valid signature.")
    }
    exit 1
}

if ($args.Length -eq 0) {
    ShowUsage
}

[bool]$Force = $False

foreach ($a in $args) {
    if ($a.ToUpper() -eq "/F") {
        $Force = $True
    }
    if ($a -ceq "/?") {
        ShowUsage
    }
}
[bool]$Found = $False
$Files = @()
foreach ($a in $args) {
    if ($a[0] -eq "/") {
        continue
    }
    $Paths = Resolve-Path -Path $a -Relative
    foreach ($f in $Paths) {
        if ((Get-Item -Path $f).PSIsContainer) {
            continue
        }
        $Sign = Get-AuthenticodeSignature -FilePath $f
        if ($Force -or $Sign -eq $null -or $Sign.Status -ne "Valid" -or $Sign.TimeStamperCertificate -eq $null) {
            $Files += $f
        } else {
            if ($PSUICulture -eq "ja-JP") {
                Write-Output ($f + " は署名済みです。スキップします。")
            } else {
                Write-Output ($f + " is already singed. Skip.")
            }
        }
    }
    $Found = $True
}

if (-not $Found) {
    ShowUsage
}
if ($Files.Length -eq 0) {
    exit 0
}

$Cert=Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert 
if ($Cert -eq $null) {
    if ($PSUICulture -eq "ja-JP") {
        [System.Console]::Error.WriteLine("端末にコードサイニング証明書がインストールされていません")
    } else {
        [System.Console]::Error.WriteLine("Code-Signing-Certificate is not found.")
    }
    exit 1
}
Set-AuthenticodeSignature -Certificate $Cert -Filepath $Files -HashAlgorithm "SHA256" -TimestampServer $TimestampServer
