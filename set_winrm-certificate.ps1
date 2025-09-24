#Syntax: .\set_winrm-certificate.ps1
#Example no Logfile is saved: .\set_winrm-certificate.ps1
#Example with saved Logfile: .\set_winrm-certificate.ps1 Save
param(
$result
)

#Variablen
$script:ver = "1.1"
$script:name = "Set Windows Remote Management Certificate Script"
$script:verdate = "30.06.2025"
$script:tmplogfile = "$env:temp\set_winrm-certificate.tmp"
$script:logfile = "$env:temp\set_winrm-certificate.log"


$script:startdate = get-date -uformat "%d.%m.%Y"
$script:starttime = get-date -uformat "%R"

$script:thumbprint = $result.ManagedItem.CertificateThumbprintHash


#Informationsblock Anfang
Write-Output ""
Write-Output "=========================================================================="
Write-Output "$name Ver. $ver, $verdate"
Write-Output "Written by Andyt for face of buildings planning stimakovits GmbH"
Write-Output "Promoted development by BlackSeals.net Technology"
Write-Output "Copyright 2025 by Reisenhofer Andreas"
Write-Output "=========================================================================="
Write-Output "Gestartet am $startdate um $starttime Uhr..."
Write-Output ""

Write-Output "_____________________________________________________________________________" >"$tmplogfile"
Write-Output "Set $name Ver. $ver, $verdate" >>"$tmplogfile"
Write-Output "Gestartet am $startdate um $starttime Uhr..." >>"$tmplogfile"
Write-Output "" >>"$tmplogfile"

#Entferne aktuellen HTTPS-Listener, Einlesen Zertifikat, Setze HTTPS-Listener mit passendem Zertifikat
Write-Output "Entferne aktuellen HTTPS-Listener..."
Write-Output "Entferne aktuellen HTTPS-Listener..." >>"$tmplogfile"
Get-ChildItem wsman:\localhost\Listener\ | Where-Object -Property Keys -like 'Transport=HTTPS' | Remove-Item -Force -Recurse >>"$tmplogfile"

Write-Output "Suche nach Zertifikat..."
Write-Output "Suche nach Zertifikat..." >>"$tmplogfile"
If (Test-Path -Path "c:\windows\NTDS") {
Write-Output "...Domaenencontroller gefunden."
Write-Output "...Domaenencontroller gefunden." >>"$tmplogfile"
$thumbprint = (dir Cert:\LocalMachine\My\ | ? {$_.DnsNameList -like "*$env:computername*"} | ? {$_.EnhancedKeyUsageList -like "*1.3.6.1.5.2.3.5*"}| ? {$_.EnhancedKeyUsageList -like "*1.3.6.1.5.5.7.3.2*"} | ? {$_.EnhancedKeyUsageList -like "*1.3.6.1.5.5.7.3.1*"}).thumbprint
} ElseIf ((dir Cert:\LocalMachine\My\ | ? {$_.Friendlyname -like "*Hyper-V*"}).thumbprint) {
Write-Output "...Hyper-V Host gefunden."
Write-Output "...Hyper-V Host gefunden." >>"$tmplogfile"
$thumbprint = (dir Cert:\LocalMachine\My\ | ? {$_.Friendlyname -like "*Hyper-V*"}).thumbprint
} ElseIf ((dir Cert:\LocalMachine\My\ | ? {$_.Friendlyname -like "*Cluster*"}).thumbprint) {
Write-Output "...Hyper-V Knoten (Clustermitglied) gefunden."
Write-Output "...Hyper-V Knoten (Clustermitglied) gefunden." >>"$tmplogfile"
$thumbprint = (dir Cert:\LocalMachine\My\ | ? {$_.Friendlyname -like "*Cluster*"}).thumbprint
} Else {
Write-Output "...Domaenenmitglied gefunden."
Write-Output "...Domaenenmitglied gefunden." >>"$tmplogfile"
$thumbprint = (dir Cert:\LocalMachine\My\ | ? {$_.DnsNameList -like "*$env:computername*"} | ? {$_.EnhancedKeyUsageList -like "*1.3.6.1.5.5.7.3.2*"} | ? {$_.EnhancedKeyUsageList -like "*1.3.6.1.5.5.7.3.1*"}).thumbprint
}

$listenerParams = @{
    ResourceUri = "winrm/config/Listener"
    SelectorSet = @{
        Address = "*"
        Transport = "HTTPS"
    }
    ValueSet = @{
        CertificateThumbprint = $thumbprint
    }
}
Write-Output "Erstelle neuen HTTPS-Listener..."
Write-Output "Erstelle neuen HTTPS-Listener..." >>"$tmplogfile"
New-WSManInstance @listenerParams >>"$tmplogfile"

#Informationsblock Ende
Write-Output "Abarbeitung am $(get-date -uformat "%d.%m.%Y") um $(get-date -uformat "%R") Uhr beendet."
Write-Output ""

Write-Output "Abarbeitung am $(get-date -uformat "%d.%m.%Y") um $(get-date -uformat "%R") Uhr beendet." >>"$tmplogfile"
Write-Output "" >>"$tmplogfile"

#Kontrolle ob Logfile existiert und gespeichert wird
if (Test-Path $tmplogfile) {

	#Kontrolle ob Logfile gespeichert werden soll.
    if ($Args[0] -eq "Save") {
        Get-Content "$tmplogfile" >>"$logfile"
    }
    remove-item -path "$tmplogfile" -force
}