# $TimestampServer �͏������Ɏg�p����^�C���X�^���v�T�[�o�[
$TimestampServer = "http://timestamp.digicert.com"

function ShowUsage {
    if ($PSUICulture -eq "ja-JP") {
        [System.Console]::Error.WriteLine("sign.ps1 [/F] <�t�@�C��1> [<�t�@�C��2>...]")
        [System.Console]::Error.WriteLine("  �����œn�����t�@�C���ɑ΂��ăR�[�h�T�C�j���O�������܂�")
        [System.Console]::Error.WriteLine("  ���ɗL���ȏ������ݒ肳��Ă���t�@�C���͏������܂���")
        [System.Console]::Error.WriteLine("  �t�@�C���͕����w��\�ł�")
        [System.Console]::Error.WriteLine("  ���s�`��(.exe/.dll)�����PowerShellScript(.ps1)�̃t�@�C�����w��\�ł�")
        [System.Console]::Error.WriteLine("  /F ���w�肷��Ɗ��ɗL���ȏ������ݒ肳��Ă���t�@�C�����ď������܂�")
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
                Write-Output ($f + " �͏����ς݂ł��B�X�L�b�v���܂��B")
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
        [System.Console]::Error.WriteLine("�[���ɃR�[�h�T�C�j���O�ؖ������C���X�g�[������Ă��܂���")
    } else {
        [System.Console]::Error.WriteLine("Code-Signing-Certificate is not found.")
    }
    exit 1
}
Set-AuthenticodeSignature -Certificate $Cert -Filepath $Files -HashAlgorithm "SHA256" -TimestampServer $TimestampServer
