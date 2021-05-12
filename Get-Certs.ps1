<# Synopsis
        Get all Local Computer Certificates

# Description
        Creates a text file containing all of the installed certificates under LocalMachine

# Parameters
        None



# Outputs
        Creates a text file under C:\Windows\Temp\AllCerts.txt
        The script automatically opens the text file
# Notes
    Version: 1.0
    Author: Joseph Hopper
    Creation Date: 5/11/2021
    Purpose/Change: Initial script development


# Example
        Run .\Get-Certs.ps1 script from an elevated PowerShell command prompt
        Right-Click on the Certs.ps1 script and select "Run with PowerShell"

# Disclaimer
# This module and it's scripts are not supported under any Microsoft standard support program or service.
# The scripts are provided AS IS without warranty of any kind.
# Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability
# or of fitness for a particular purpose.
# The entire risk arising out of the use or performance of the scripts and documentation remains with you.
# In no event shall Microsoft, its authors, or anyone else involved in the creation, production,
# or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages
# for loss of business profits, business interruption, loss of business information, or other pecuniary loss)
# arising out of the use of or inability to use the sample scripts or documentation,
# even if Microsoft has been advised of the possibility of such damages.

# Disclaimer: 
# GitHub is an external service subject to its own privacy and licensing terms. Your download and use of any files from GitHub 
# here is subject to the license terms of those files provided on GitHub.
#>

# Set outputfile
$outfile = 'C:\Windows\Temp\AllCerts.txt'
# Delete outputfile if already exist
if (Test-Path $outfile -PathType Leaf){
Remove-Item $outfile
}
# Declare variables to be used
$CertStores = Get-Item -Path Cert:\LocalMachine\*
$CertStoreNames = ($CertStores).Name
$myCerts=@()
$OldCertPath=""
$CertPath=""

# Go through each Cert Store Name and get all listed certificates
for ($j=0;$j -lt $CertStoreNames.Count; $j++){
$myCerts += Get-ChildItem -Path "Cert:\LocalMachine\$($CertStoreNames[$j])" | Select-Object -Property DNSNameList, FriendlyName, IssuerName, NotAfter, NotBefore, SerialNumber, SubjectName, Thumbprint, Issuer, Subject, EnhancedKeyUsageList, PSParentPath
Write-Progress -Activity "Enumerating certificates..." -Status "$j% Complete:" -PercentComplete $j;
}

# Write the desired information from every certificate to a file
for ($i=0;$i -lt $myCerts.Count; $i++){
    $CertPath = $myCerts[$i].PSParentPath.Substring($myCerts[$i].PSParentPath.IndexOf("::")+2)
        # if OldCertPath is not equal to CertPath 
        if($OldCertPath -ne $CertPath){
       # Write the CertPath info to file
        Write-Output "" | Out-File -FilePath $outfile -Append
        Write-Output "" | Out-File -FilePath $outfile -Append
        Write-Output "$($Border)  $($CertPath) $($Border)" | Out-File -FilePath $outfile -Append
        Write-Output "" | Out-File -FilePath $outfile -Append
        Write-Output "" | Out-File -FilePath $outfile -Append
        }
    
    Write-Output "--------------------------------------------------------------" | Out-File -FilePath $outfile -Append
    Write-Output "CertPath:             $($myCerts[$i].PSParentPath) " | Out-File -FilePath $outfile -Append
    Write-Output "FriendlyName:         $($myCerts[$i].FriendlyName) " | Out-File -FilePath $outfile -Append
    Write-Output "Subject:              $($myCerts[$i].Subject) " | Out-File -FilePath $outfile -Append
    Write-Output "Issuer:               $($myCerts[$i].Issuer)" | Out-File -FilePath $outfile -Append
    Write-Output "DnsNameList:          $($myCerts[$i].DnsNameList) " | Out-File -FilePath $outfile -Append
    Write-Output "NotAfter:             $($myCerts[$i].NotAfter)" | Out-File -FilePath $outfile -Append
    Write-Output "EnhancedKeyUsageList: $($myCerts[$i].EnhancedKeyUsageList) "   | Out-File -FilePath $outfile -Append  
    Write-Output "SerialNumber:         $($myCerts[$i].SerialNumber)" | Out-File -FilePath $outfile -Append
    Write-Output "Thumbprint:           $($myCerts[$i].Thumbprint) " | Out-File -FilePath $outfile -Append
    Write-Output "--------------------------------------------------------------" | Out-File -FilePath $outfile -Append

    # Divide item by variable count, then multiply by 100 for percentage result
    $PercComp = (($i/$myCerts.Count)*100)
    # Round down the result for display in the -Status parameter
    $PercComp = [math]::Round($PercComp)
    Write-Progress -Activity "Writing certificate information" -Status "$PercComp% Complete:" -PercentComplete (($i/$myCerts.Count)*100)
    $OldCertPath = $CertPath
}

Notepad.exe $outfile