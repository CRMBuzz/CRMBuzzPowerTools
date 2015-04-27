
cls

## When the PSSnapin was installed use the following lines.
# if ( (Get-PSSnapin -Name CRMBuzzPowerTools -ErrorAction SilentlyContinue) -eq $null )
# {
#     Add-PSSnapin CRMBuzzPowerTools
# }


### Loading the Asembly as a Module when the PSSnapin was not installed

Import-Module $home\Documents\WindowsPowerShell\Modules\CRMBuzzPowerTools\CRMBuzz.PowerTools.PSSnapin.dll -Force -WarningAction SilentlyContinue -DisableNameChecking

### Get-Command -module CRMBuzz.PowerTools.PSSnapin


Write-Host "+===================================================== "
Write-Host "| "
Write-Host "| Sample Script will Add User Manager Assigment"
Write-Host "| "
Write-Host "| " 
Write-Host "| "
Write-Host "| "
Write-Host "| OnLine/OnPrem"
Write-Host "| Use the File -> CSV_ManagerAssigment_Template.csv"
Write-Host "+===================================================== "

$continue = Read-Host 'Do you want to continue <Y/N>?'
if($continue -eq 'n' -or $continue -eq 'N'){

exit

}



###
### Connection to CRM.....UPDATE THE USER NAME AND PASSWORD.......
###
[string]$connString = Get-Content "C:\POWERSHELL\FinalSCRIPTS\ConnectionString.txt" -Raw

$CRMConn = New-OrganizationConnection -ConnectionString $connString  -Verbose

	If ($CRMConn -eq $null) 
	{ 
		write-host "Can't Connect with CRM Server" -ForegroundColor Red
		write-host " " -ForegroundColor white
        
        ## Removing-PSSnapin CRMBuzzPowerTools
        
		Exit
	} 

[string] $locationCSV = "C:\POWERSHELL\FinalSCRIPTS\"
[int]$RowCount = 0

###
### Open the CSV anb reading the content
###
$Objects = Import-Csv ($locationCSV + "CSV_User_ManagerAssigment_Template.csv")

###
### iteration on all objects
###
    $Objects | ForEach-Object { 


        [bool]$ProcessRec = [convert]::ToBoolean($_.Process) ## YES/NO

        If($ProcessRec -eq $true )
        {
            [string]$RecordType = $_.Type ### Online/OnPrem
            [string]$Fname = $_.FirstName
            [string]$LName = $_.LastName

            [string]$FullName = ( $Fname + " " + $LName)

            [string]$UserDomainName = $_.UserDomainName
            [string]$UserEmailAddress = $_.UserEmailAddress 
            

            [string]$ManagerDomainName = $_.ManagerDomainName
            [string]$ManagerEmailAddress = $_.ManagerEmailAddress 
            



             ###
             ### Presenting the CSV read data
             ###
             Write-Host "=========== User Information in CSV ===========" -ForegroundColor Yellow
             Write-Host "Record Type       : "   $RecordType -ForegroundColor Green; 
             Write-Host "User DomainName   : "   $UserDomainName -ForegroundColor Green;
             Write-Host "User Emaill       : "   $UserEmailAddress -ForegroundColor Green;
             Write-Host "Manager DomainName: "   $ManagerDomainName -ForegroundColor Green;
             Write-Host "Manager Email     : "   $ManagerEmailAddress -ForegroundColor Green;

                    
            ###
            ### 
            ###
            try{
                Write-Host "===== Updating Record =========" -ForegroundColor Yellow

                if($RecordType -eq "Online"){
                    Set-UserAssignManager -OrganizationService $CRMConn -UserPrimaryEmail $UserEmailAddress -ManagerPrimaryEmail $ManagerEmailAddress -Verbose
                }

                if($RecordType -eq "OnPrem"){
                    Set-UserAssignManager -OrganizationService $CRMConn -UserDomainAlias $UserDomainName -ManagerDomainAlias $ManagerDomainName -Verbose
                }



            }catch{

                ### Trapping Error messages
                Write-Host $UpdateError -ForegroundColor Red

            }



            ###
            ### 
            ###
            Write-Host " "
            Write-Host " "
            Write-Host "===== Completed Assigning Manager Record =========" -ForegroundColor Yellow
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
