Add-Type -AssemblyName System.IO.Compression.FileSystem
function Unzip
{
    param([string]$zipfile, [string]$outpath)

    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
}

$url = 'https://github.com/PowerShell/Win32-OpenSSH/releases/latest/'
$request = [System.Net.WebRequest]::Create($url)
$request.AllowAutoRedirect=$false
$response=$request.GetResponse()
$file = $([String]$response.GetResponseHeader("Location")).Replace('tag','download') + '/OpenSSH-Win64.zip'

$client = new-object system.Net.Webclient;
$client.DownloadFile($file ,"c:\\OpenSSH-Win64.zip")

Unzip "c:\\OpenSSH-Win64.zip" "C:\Program Files\" 
mv "c:\\Program Files\OpenSSH-Win64" "C:\Program Files\OpenSSH\" 

powershell.exe -ExecutionPolicy Bypass -File "C:\Program Files\OpenSSH\install-sshd.ps1"

New-NetFirewallRule -Name sshd -DisplayName 'OpenSSH Server (sshd)' -Enabled True -Direction Inbound -Protocol TCP -Action Allow -LocalPort 22

net start sshd

Set-Service sshd -StartupType Automatic
Set-Service ssh-agent -StartupType Automatic

cd "C:\Program Files\OpenSSH\"
Powershell.exe -ExecutionPolicy Bypass -Command '. .\FixHostFilePermissions.ps1 -Confirm:$false'

$registryPath = "HKLM:\SOFTWARE\OpenSSH\"
$Name = "DefaultShell"
#$value = "C:\windows\System32\WindowsPowerShell\v1.0\powershell.exe"
$value = "C:\Program Files\PowerShell\7\pwsh.exe"

IF(!(Test-Path $registryPath))
  {
    New-Item -Path $registryPath -Force
    New-ItemProperty -Path $registryPath -Name $name -Value $value -PropertyType String -Force
} ELSE {
    New-ItemProperty -Path $registryPath -Name $name -Value $value -PropertyType String -Force
}