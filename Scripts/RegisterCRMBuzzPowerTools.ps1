# =====================================================================
# Installation file for CRM Power Tools Framework
#
#  Copyright (C) CRMBuzz.co  All rights reserved.
#
#  THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
#  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
#  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
#  PARTICULAR PURPOSE.
# =====================================================================

#======================================================================
#Registers or Unregisters the CRMBuzz.PowerTools.PSSnapin.dll 
#======================================================================

param
(
[switch]$uninstall
)

If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{   
    Write-Error("This script must be run in admin mode, please restart powershell and run it as an administrator");
break
}



# Configure PowerShell to use the most updated .Net Framework Version
# Reg Add HKLM\Software\Microsoft\.NetFramework /v OnlyUseLatestCLR /t REG_DWORD /d 1

# Reg Add HKLM\Software\wow6432node\Microsoft\.NetFramework /v OnlyUseLatestCLR /t REG_DWORD /d 1




#setup the alias I need. 
Set-Alias installUtil $env:windir\Microsoft.NET\Framework64\v4.0.30319\installutil.exe

if ( $uninstall ) 
{
    #unregister components:
    if ( (Get-Item("CRMBuzz.PowerTools.PSSnapin.dll") -ErrorAction SilentlyContinue) -ne $null ) 
    {
        installUtil /u CRMBuzz.PowerTools.PSSnapin.dll
    }
    else
    {
        Write-Host "Did Not find CRMBuzz.PowerTools.PSSnapin.dll" -ForegroundColor Red
    }

}
else
{
    #register components:
    if ( (Get-Item("CRMBuzz.PowerTools.PSSnapin.dll") -ErrorAction SilentlyContinue ) -ne $null ) 
    {
        Write-Host "Found CRMBuzz.PowerTools.PSSnapin.dll" -ForegroundColor Yellow
        installUtil CRMBuzz.PowerTools.PSSnapin.dll

        ####
        #### To validate that the PSNappin was installed correctly close installation PS session and open a new session
        #### and type the following cmdlets

        ### Add-PSSnapin CRMBuzzPowerTools
        ### Get-Command -PSSnapin CRMBuzzPowerTools

    }
    else
    {
        Write-Host "Did Not find CRMBuzz.PowerTools.PSSnapin.dll" -ForegroundColor Red
    }




}


