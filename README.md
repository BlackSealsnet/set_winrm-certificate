<h1 align="center">
  <a href="https://blackseals.net">
    <img src="https://blackseals.net/features/blackseals.png" width=66% alt="BlackSeals">
  </a>
</h1>

> Promoted development by BlackSeals.net Technology.
> Written by Andyt for BlackSeals.net.
> Copyright 2025 by BlackSeals Network.

## Description

**set_winrm-certificate.ps1** is a Windows PowerShell script that will search for certificate in local certificate store and use it for Windows Remote Management to enable HTTPS for WinRM. During the process any old HTTPS-Listener for WinRM will be removed and the new HTTPS-Listener will be created with new certificate for domain controller or domain member.


## Requirement

* The local certificate store should contain a certificate that is available according to the role of the system. The script use different search strings for domain controller or domain computer.
* Domain computers are checked separately to see if there are special certificates for the Windows Hyper-V role. In addition, a check is also performed to see if it is a cluster node. This is done purely on the basis of the certificate name. It looks for 'Hyper-V' or 'Cluster' in the friendly name.

 
## Quick Start

Download the script and copy **set_winrm-certificate.ps1** to the server which should get WinRM with HTTPS. Look to Requirement for more information.


## Syntax

`.\set_winrm-certificate.ps1`
* no parameters necessary


## Examples

`.\set_winrm-certificate.ps1`
* The script will to several steps to get WinRM with HTTPS.
* There is no logfile after finishing the process.

`.\set_winrm-certificate.ps1 Save`
* The script will to several steps to get WinRM with HTTPS.
* There is a logfile after finishing the process. You find it at your temporary user folder: $env:temp (Powershell) or %temp% (Explorer, Command Shell).

