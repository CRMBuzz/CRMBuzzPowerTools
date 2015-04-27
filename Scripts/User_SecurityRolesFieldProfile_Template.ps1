

# Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
cls

## When the PSSnapin was installed use the following lines.
# if ( (Get-PSSnapin -Name CRMBuzzPowerTools -ErrorAction SilentlyContinue) -eq $null )
# {
#     Add-PSSnapin CRMBuzzPowerTools
# }


### Loading the Asembly as a Module when the PSSnapin was not installed

Import-Module $home\Documents\WindowsPowerShell\Modules\CRMBuzzPowerTools\CRMBuzz.PowerTools.PSSnapin.dll -Force -WarningAction SilentlyContinue -DisableNameChecking

Write-Host "+===================================================== "
Write-Host "| "
Write-Host "| Sample Script will Add or Remove"
Write-Host "| for the following items:"
Write-Host "|     * Security Roles" 
Write-Host "|     * Field Level Security"
Write-Host "| "
Write-Host "| OnLine/OnPrem"
Write-Host "| Use the File -> CSV_User_SecurityRoles_Template.csv"
Write-Host "+===================================================== "



$continue = Read-Host 'Do you want to continue <Y/N>?'
if($continue -eq 'n' -or $continue -eq 'N'){

exit

}

###
### Connection to CRM.....UPDATE THE USER NAME AND PASSWORD.......
###
[string]$connString = Get-Content "C:\POWERSHELL\FinalSCRIPTS\ConnectionString.txt" -Raw

$CRMConn = New-OrganizationConnection  -ConnectionString $connString -Verbose

$locationCSV ="C:\POWERSHELL\FinalSCRIPTS\"



	If ($CRMConn -eq $null) 
	{ 
		write-host "Can't Connect with CRM Server" -ForegroundColor Red
		write-host " " -ForegroundColor white
        
        ## Removing-PSSnapin CRMBuzzPowerTools
        
		Exit
	} 

[bool]$Flag = $true
[int]$RowCount = 0

  $Users = Import-Csv ($locationCSV + "CSV_User_SecurityRoles_Template.csv")

  $Users | ForEach-Object { 
        [bool]$ProcessRec = [convert]::ToBoolean($_.Process)

        If($ProcessRec -eq $true )
        {

            [string]$RecordType = $_.Type
            [string]$Fname = $_.FirstName
            [string]$LName = $_.LastName

            [string]$FullName = ( $Fname + " " + $LName)

            [string]$EmailAddress = $_.EmailAddress 
            [string]$DomainName = $_.DomainName

            [string]$RemoveRole = $_.RemoveRole 
            [string]$AddRole = $_.AddRole 

            [string]$RemoveFieldProfile = $_.RemoveFieldProfile 
            [string]$AddFieldProfile = $_.AddFieldProfile 


             Write-Host "=========== Users to get Details ===========" -ForegroundColor Yellow
             Write-Host "User Name    : "   $FullName -ForegroundColor Green;             
             Write-Host "Domain Name  : "   $DomainName -ForegroundColor Green;             
             Write-Host "Email Address: "   $EmailAddress -ForegroundColor Green; 
             Write-Host "Record Type: "   $RecordType -ForegroundColor Green; 

             Write-Host "Remove Role: "   $RemoveRole -ForegroundColor Red; 
             Write-Host "Add Role: "   $AddRole -ForegroundColor Yellow; 

             Write-Host "Remove Field level: "   $RemoveFieldProfile -ForegroundColor Red; 
             Write-Host "Add Field Level: "   $AddFieldProfile -ForegroundColor Yellow; 

           
            if($RemoveRole -ne "NA"){

                if($RecordType -eq "Online"){
                    Remove-AssignedRoleAndFieldProfileToUser -OrganizationService $CRMConn -PrimaryEmail $EmailAddress -RoleName $RemoveRole -Verbose  
                }

                if($RecordType -eq "OnPrem"){
                    Remove-AssignedRoleAndFieldProfileToUser -OrganizationService $CRMConn -UserDomainAlias $DomainName -RoleName $RemoveRole 
                }
            }

            if($AddRole -ne "NA"){

                if($RecordType -eq "Online"){
                    Set-UserAssignRole -OrganizationService $CRMConn -PrimaryEmail $EmailAddress -RoleName $AddRole -Verbose
                }

                if($RecordType -eq "OnPrem"){
                    Set-UserAssignRole -OrganizationService $CRMConn -UserDomainAlias $DomainName -RoleName $AddRole -Verbose
                }
                  
            }



            if($RemoveFieldProfile -ne "NA"){

                if($RecordType -eq "Online"){
                    Remove-AssignedRoleAndFieldProfileToUser -OrganizationService $CRMConn -PrimaryEmail $EmailAddress -ProfileName $RemoveFieldProfile -Verbose  
                }

                if($RecordType -eq "OnPrem"){
                    Remove-AssignedRoleAndFieldProfileToUser -OrganizationService $CRMConn -UserDomainAlias $DomainName -ProfileName $RemoveFieldProfile -Verbose 
                }
            }

            if($AddFieldProfile -ne "NA"){

                if($RecordType -eq "Online"){
                    
                    Set-UserAssignRoleAndFieldProfile -OrganizationService $CRMConn -PrimaryEmail $EmailAddress -ProfileName $AddFieldProfile -Verbose
                }

                if($RecordType -eq "OnPrem"){
                    Set-UserAssignRoleAndFieldProfile -OrganizationService $CRMConn -DomainName $DomainName -ProfileName $AddFieldProfile -Verbose
                }

                  
            }

       
            Write-Host " "
            Write-Host " "
            Write-Host "===== Completed Updating user information =========" -ForegroundColor Yellow
            Write-Host " "
            Write-Host " "


        }else {

            Write-Host "+===================================================== "
            Write-Host "| "
            Write-Host "| "
            Write-Host "| "
            Write-Host "| " $FullName 
            Write-Host "| "
            Write-Host "| "
            Write-Host "| "
            Write-Host "+===================================================== "

        }


        $RowCount++
        Write-Host " " $RowCount -ForegroundColor White

        ###
        ### paging Stops
        ###
        if($RowCount -eq 50 -or $RowCount -eq 100 -or $RowCount -eq 150 -or $RowCount -eq 200){
            $continue = Read-Host 'Paging - Press any key to continue...'
        }


    }


# Remove-PSSnapin PowerCRMPSSnapin

Exit


