

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
Write-Host "| Sample Script will Get all User information and assigments"
Write-Host "| for the following items:"
Write-Host "|     * Security Roles" 
Write-Host "|     * Field Level Security"
Write-Host "|     * Teams"
Write-Host "| "
Write-Host "| OnLine/OnPrem"
Write-Host "| Use the File -> CSV_User_SecurityAccessMembershipss_Template.csv"
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
$OutPutFileCSV = "C:\POWERSHELL\FinalSCRIPTS\Production-UsersData.csv"


	If ($CRMConn -eq $null) 
	{ 
		write-host "Can't Connect with CRM Server" -ForegroundColor Red
		write-host " " -ForegroundColor white
        
        ## Removing-PSSnapin CRMBuzzPowerTools
        
		Exit
	} 

[bool]$Flag = $true
[int]$RowCount = 0

  $Users = Import-Csv ($locationCSV + "CSV_User_SecurityAccessMemberships_Template.csv")

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

             Write-Host "=========== Users to get Details ===========" -ForegroundColor Yellow
             Write-Host "User Name    : "   $FullName -ForegroundColor Green;             
             Write-Host "Domain Name  : "   $DomainName -ForegroundColor Green;             
             Write-Host "Email Address: "   $EmailAddress -ForegroundColor Green; 
             Write-Host "Record Type: "   $RecordType -ForegroundColor Green; 


            if($RecordType -eq "Online"){
                ### Add Header at the begining of the CSV file
                if($Flag -eq $true){
                    ### Add CSV Header
                    Get-UserAccessMembership -OrganizationService $CRMConn -PrimaryEmail $EmailAddress -OutputFile $OutPutFileCSV -addCSVheaderLine
                }else{
                    ### Do not add CSV heaver after the 2nd record
                    Get-UserAccessMembership -OrganizationService $CRMConn -PrimaryEmail $EmailAddress -OutputFile $OutPutFileCSV 

                }
            }

            if($RecordType -eq "OnPrem"){
                ### Add Header at the begining of the CSV file
                if($Flag -eq $true){
                    Get-UserAccessMembership -OrganizationService $CRMConn   -DomainName $DomainName -OutputFile $OutPutFileCSV -addCSVheaderLine
                }else{
                    Get-UserAccessMembership -OrganizationService $CRMConn   -DomainName $DomainName -OutputFile $OutPutFileCSV

                }
            }

            ### Do NOT add Header after the first line
            $Flag = $false           
       
            Write-Host " "
            Write-Host " "
            Write-Host "===== Completed gathering user information =========" -ForegroundColor Yellow
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


### Remove-PSSnapin PowerCRMPSSnapin

Exit


