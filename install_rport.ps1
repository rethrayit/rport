$dest = "C:\Program Files\tacoscript"
if(Test-Path -Path $dest) {
    Write-Host "Tacoscript already installed to $($dest)"
    exit 0
}
$Temp = [System.Environment]::GetEnvironmentVariable('TEMP','Machine')
Set-Location $Temp
$url = "https://download.rport.io/tacoscript/stable/?arch=Windows_x86_64"
$file = "tacoscript.zip"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Invoke-WebRequest -Uri $url -OutFile $file -UseBasicParsing
Write-Host "Tacoscript dowloaded to $($Temp)\$($file)"
New-Item -ItemType Directory -Force -Path "$($dest)\bin"|Out-Null
Expand-Archive -Path $file -DestinationPath $dest -force