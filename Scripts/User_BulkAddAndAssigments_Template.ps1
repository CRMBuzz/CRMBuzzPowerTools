

# Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
cls

## When the PSSnapin was installed use the following lines.
# if ( (Get-PSSnapin -Name CRMBuzzPowerTools -ErrorAction SilentlyContinue) -eq $null )
# {
#     Add-PSSnapin CRMBuzzPowerTools
# }


### Loading the Asembly as a Module when the PSSnapin was not installed

Import-Module $home\Documents\WindowsPowerShell\Modules\CRMBuzzPowerTools\CRMBuzz.PowerTools.PSSnapin.dll -Force -WarningAction SilentlyContinue -DisableNameChecking


###### ASCII CODE CHARTS
######      201 ╔   205═  203╦    187 ╗
######      186 ║   204╠  206╬    185 ╣
######      200 ╚         202╩    188 ╝
Write-Host "╔═════════════════════════════════════════════════════════════════════════════════════════════════╗ " -ForegroundColor White
Write-Host "║ " -ForegroundColor White
Write-Host "║ Sample Script will Add User for OnPrem and also" -ForegroundColor White
Write-Host "║ Adding the following items for Online and OnPrem:" -ForegroundColor White
Write-Host "║ " -ForegroundColor white -NoNewline; Write-Host "     * Approve User Email Address" -ForegroundColor Green
Write-Host "║ " -ForegroundColor white -NoNewline; Write-Host "     * Teams" -ForegroundColor Green
Write-Host "║ " -ForegroundColor white -NoNewline; Write-Host "     * Security Roles" -ForegroundColor Green
Write-Host "║ " -ForegroundColor white -NoNewline; Write-Host "     * Field Level Security"-ForegroundColor Green
Write-Host "║ "
Write-Host "║ OnLine/OnPrem"
Write-Host "║ Use the File -> CSV_User_BulkAddAndAssigments_Template.csv"
Write-Host "╚═════════════════════════════════════════════════════════════════════════════════════════════════╝"



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

  $Users = Import-Csv ($locationCSV + "CSV_User_BulkAddAndAssigments_Template.csv")

  $Users | ForEach-Object { 
        [bool]$ProcessRec = [convert]::ToBoolean($_.Process)

            [string]$Fname = $_.FirstName
            [string]$LName = $_.LastName

            [string]$FullName = ( $Fname + " " + $LName)
            [string]$Title = $_.Title
        

        If($ProcessRec -eq $true )
        {

            [bool]$OnPremCreateUser = [convert]::ToBoolean($_.Create)
            [bool]$ApproveUserEmail = [convert]::ToBoolean($_.ApproveEmail)

            [string]$RecordType = $_.Type

            
            [string]$DomainName = $_.DomainName
            [string]$EmailAddress = $_.EmailAddress 


            [string]$SecurityRole1 = $_.SecurityRole1 
            [string]$SecurityRole2 = $_.SecurityRole2 
            [string]$SecurityRole3 = $_.SecurityRole3 


            [string]$FieldSecurityProfile1 = $_.FieldSecurityProfile1 
            [string]$FieldSecurityProfile2 = $_.FieldSecurityProfile2

            [string]$Team1 = $_.Team1
            [string]$Team2 = $_.Team2
            [string]$Team3 = $_.Team3
            [string]$Team4 = $_.Team4

             Write-Host "╔══════════════════════════════════════════════════════════╗ "
             Write-Host "║" -ForegroundColor White -NoNewline; Write-Host "                   CSV USER DETAILS" -ForegroundColor Yellow
             Write-Host "║" -ForegroundColor White -NoNewline; Write-Host " Record Type  : " -ForegroundColor Green -NoNewline; Write-Host $RecordType.ToUpper() -ForegroundColor Yellow
             Write-Host "║" -ForegroundColor White -NoNewline; Write-Host " User Name    : " -ForegroundColor Green -NoNewline; Write-Host $FullName -ForegroundColor Yellow             
             Write-Host "║" -ForegroundColor White -NoNewline; Write-Host " Domain Name  : " -ForegroundColor Green -NoNewline; Write-Host $DomainName -ForegroundColor Red
             Write-Host "║" -ForegroundColor White -NoNewline; Write-Host " Email Address: " -ForegroundColor Green -NoNewline; Write-Host $EmailAddress -ForegroundColor Red
             

             Write-Host "║" -ForegroundColor White -NoNewline; Write-Host " Security Role 1: " -ForegroundColor Green -NoNewline; Write-Host $SecurityRole1 -ForegroundColor Red
             Write-Host "║" -ForegroundColor White -NoNewline; Write-Host " Security Role 2: " -ForegroundColor Green -NoNewline; Write-Host $SecurityRole2 -ForegroundColor Red
             Write-Host "║" -ForegroundColor White -NoNewline; Write-Host " Security Role 3: " -ForegroundColor Green -NoNewline; Write-Host $SecurityRole3 -ForegroundColor Red

             Write-Host "║" -ForegroundColor White -NoNewline; Write-Host " Field Level Profile 1: " -ForegroundColor Green -NoNewline; Write-Host $FieldSecurityProfile1 -ForegroundColor Red
             Write-Host "║" -ForegroundColor White -NoNewline; Write-Host " Field Level Profile 2: "  -ForegroundColor Green -NoNewline; Write-Host $FieldSecurityProfile2 -ForegroundColor Red

             Write-Host "║" -ForegroundColor White -NoNewline; Write-Host " Team 1: "  -ForegroundColor Green -NoNewline; Write-Host $Team1 -ForegroundColor Red
             Write-Host "║" -ForegroundColor White -NoNewline; Write-Host " Team 2: "  -ForegroundColor Green -NoNewline; Write-Host $Team2 -ForegroundColor Red
             Write-Host "║" -ForegroundColor White -NoNewline; Write-Host " Team 3: "  -ForegroundColor Green -NoNewline; Write-Host $Team3 -ForegroundColor Red
             Write-Host "║" -ForegroundColor White -NoNewline; Write-Host " Team 4: "  -ForegroundColor Green -NoNewline; Write-Host $Team4 -ForegroundColor Red



             Write-Host "║ "  

            ###
            ###  ONPREM ADD USER 
            ###
             if($RecordType.ToLower() -eq "onprem"){
                if($OnPremCreateUser -eq $true){
                    Add-User -OrganizationService $CRMConn -DomainName $DomainName -FirstName $Fname -LastName $LName -EmailAddress $EmailAddress -Title $Title -Version 2015 -Verbose

                }
             }


            ###
            ###  ADDING SECURITY ROLES
            ###
            
            Write-Host "║ APPROVE USER EMAIL ADDRESS"
            Write-Host "║ "
             if($ApproveUserEmail -eq $true)
             {
                if($RecordType.ToLower() -eq "online"){
                    Set-UserApproveEmail -OrganizationService $CRMConn -PrimaryEmail $EmailAddress -Verbose
                }
                if($RecordType.ToLower() -eq "onprem"){
                    Set-UserApproveEmail -OrganizationService $CRMConn -DomainName $DomainName -Verbose
                }

             }


            ###
            ###  ADDING SECURITY ROLES
            ###
            Write-Host "║ SECURITY ROLES"
            Write-Host "║ "

            if($SecurityRole1.ToUpper() -ne "NA"){
                if($RecordType.ToLower() -eq "online"){
                    Set-UserAssignRole -OrganizationService $CRMConn -PrimaryEmail $EmailAddress -RoleName $SecurityRole1 -Verbose
                }
                if($RecordType.ToLower() -eq "onprem"){
                    Set-UserAssignRole -OrganizationService $CRMConn -UserDomainAlias $DomainName -RoleName $SecurityRole1 -Verbose
                }
            }


            if($SecurityRole2.ToUpper() -ne "NA"){
                if($RecordType.ToLower() -eq "online"){
                    Set-UserAssignRole -OrganizationService $CRMConn -PrimaryEmail $EmailAddress -RoleName $SecurityRole2 -Verbose
                }
                if($RecordType.ToLower() -eq "onprem"){
                    Set-UserAssignRole -OrganizationService $CRMConn -UserDomainAlias $DomainName -RoleName $SecurityRole2 -Verbose
                }
            }

            if($SecurityRole3.ToUpper() -ne "NA"){
                if($RecordType.ToLower() -eq "online"){
                    Set-UserAssignRole -OrganizationService $CRMConn -PrimaryEmail $EmailAddress -RoleName $SecurityRole3 -Verbose
                }
                if($RecordType.ToLower() -eq "onprem"){
                    Set-UserAssignRole -OrganizationService $CRMConn -UserDomainAlias $DomainName -RoleName $SecurityRole3 -Verbose
                }
            }


            ###
            ###  ADDING FIELD LEVE SECURITY PROFILES
            ###
            Write-Host "║ FIELD LEVE SECURITY PROFILES"
            Write-Host "║ "

            if($FieldSecurityProfile1.ToUpper() -ne "NA"){

                if($RecordType.ToLower() -eq "online"){
                    
                    Set-UserAssignRoleAndFieldProfile -OrganizationService $CRMConn -PrimaryEmail $EmailAddress -ProfileName $FieldSecurityProfile1 -Verbose
                }

                if($RecordType.ToLower() -eq "onprem"){
                    Set-UserAssignRoleAndFieldProfile -OrganizationService $CRMConn -DomainName $DomainName -ProfileName $FieldSecurityProfile1 -Verbose
                }

                  
            }

            if($FieldSecurityProfile2.ToUpper() -ne "NA"){

                if($RecordType.ToLower() -eq "online"){
                    
                    Set-UserAssignRoleAndFieldProfile -OrganizationService $CRMConn -PrimaryEmail $EmailAddress -ProfileName $FieldSecurityProfile2 -Verbose
                }

                if($RecordType.ToLower() -eq "onprem"){
                    Set-UserAssignRoleAndFieldProfile -OrganizationService $CRMConn -DomainName $DomainName -ProfileName $FieldSecurityProfile2 -Verbose
                }

                  
            }



            ###
            ###  ADDING TEAMS
            ###
            Write-Host "║ TEAMS"
            Write-Host "║ "
            if($Team1.ToUpper() -ne "NA"){
                if($RecordType.ToLower() -eq "online"){
                    Set-TeamAssignUser -OrganizationService $CRMConn -PrimaryEmail $EmailAddress -TeamName $Team1 -Verbose
                }
                if($RecordType.ToLower() -eq "onprem"){
                    Set-TeamAssignUser -OrganizationService $CRMConn -DomainName $DomainName -TeamName $Team1 -Verbose
                }
            }

            if($Team2.ToUpper() -ne "NA"){
                if($RecordType.ToLower() -eq "online"){
                    Set-TeamAssignUser -OrganizationService $CRMConn -PrimaryEmail $EmailAddress -TeamName $Team2 -Verbose
                }
                if($RecordType.ToLower() -eq "onprem"){
                    Set-TeamAssignUser -OrganizationService $CRMConn -DomainName $DomainName -TeamName $Team2 -Verbose
                }
            }

            if($Team3.ToUpper() -ne "NA"){
                if($RecordType.ToLower() -eq "online"){
                    Set-TeamAssignUser -OrganizationService $CRMConn -PrimaryEmail $EmailAddress -TeamName $Team3 -Verbose
                }
                if($RecordType.ToLower() -eq "onprem"){
                    Set-TeamAssignUser -OrganizationService $CRMConn -DomainName $DomainName -TeamName $Team3 -Verbose
                }
            }

            if($Team4.ToUpper() -ne "NA"){
                if($RecordType.ToLower() -eq "online"){
                    Set-TeamAssignUser -OrganizationService $CRMConn -PrimaryEmail $EmailAddress -TeamName $Team4 -Verbose
                }
                if($RecordType.ToLower() -eq "onprem"){
                    Set-TeamAssignUser -OrganizationService $CRMConn -DomainName $DomainName -TeamName $Team4 -Verbose
                }
            }
       
            Write-Host "║ "
            Write-Host "║ "
            Write-Host "║" -ForegroundColor white -NoNewline; Write-Host "    COMPLETED UPDATING USER INFORMATION. " -ForegroundColor Yellow
            Write-Host "╚══════════════════════════════════════════════════════════╝ "
            
            Write-Host " "


        }else {

            Write-Host "╔══════════════════════════════════════════════════════════╗ "
            Write-Host "║ "
            Write-Host "║ "
            Write-Host "║ "
            Write-Host "║ " $FullName 
            Write-Host "║ "
            Write-Host "║ "
            Write-Host "║ "
            Write-Host "╚══════════════════════════════════════════════════════════╝ "

        }


        $RowCount++
        Write-Host "───» Record Count: #" $RowCount  " «────" -ForegroundColor White
        Write-Host " "
        Write-Host " "

        ###
        ### paging Stops
        ###
        if($RowCount -eq 50 -or $RowCount -eq 100 -or $RowCount -eq 150 -or $RowCount -eq 200){
            $continue = Read-Host 'Paging - Press any key to continue...'
        }


    }


# Remove-PSSnapin PowerCRMPSSnapin

Exit


