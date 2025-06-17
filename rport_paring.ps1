#!ps
#timeout=90000
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$url="https://pairing.rport.io/RgH1XIp"
Invoke-WebRequest -Uri $url -OutFile "rport-installer.ps1"
(Get-Content rport-installer.ps1).Replace('https://downloads.rport.io/rport/$( $release )/latest.php?filter=Windows_x86_64.msi&gt=$( $gt )', 'https://downloads.rport.io/rport/stable/rport_0.9.12_windows_x86_64.msi') | Set-Content rport-installer.ps1
(Get-Content rport-installer.ps1).Replace('$( $HostUUID )','$( $client_id )') | Set-Content rport-installer.ps1
powershell -ExecutionPolicy Bypass -File .\rport-installer.ps1 -x -r -i