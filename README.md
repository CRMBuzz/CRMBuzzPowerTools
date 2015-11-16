#Welcome to the CRMBuzzPowerTools wiki!#

The administration and configuration task you can automate using PowerShell scripts, automate user creation, assign teams, queues, security roles, field level security reusing PowerShell scripts

The idea for the framework is to allow the administrator to have an automation option for the manual configuration and administration task using the flexibility of PowerShell and the extensions on the CRM SDK

The framework contains 96+ cmdlets included in a PowerShell Snapin, the cmdlets are extensions of .Net framework and CRM SDK objects you will be able to manipulate with ease. 

These are some of the group of cmdlets 

##CRM General Objects (19)##
* ConvertTo-DateTime
* [ConvertTo-Guid] (https://github.com/CRMBuzz/CRMBuzzPowerTools/wiki/ConvertTo-Guid)
* Get-ActivityParty
* [Get-AssemblyVersion] (https://github.com/CRMBuzz/CRMBuzzPowerTools/wiki/Get-AssemblyVersion) 2.0.0.15
* [Get-CRMEntity] (https://github.com/CRMBuzz/CRMBuzzPowerTools/wiki/Get-CRMEntity)
* Get-EntityCollection
* New-ColumnSet
* New-ConditionExpression
* New-FilterExpression
* [New-Guid] (https://github.com/CRMBuzz/CRMBuzzPowerTools/wiki/New-Guid)
* New-KeyAttributeCollection \<NEW\>
* New-LinkEntity
* New-OrderExpression
* New-QueryExpression
* [Set-EntityReference](https://github.com/CRMBuzz/CRMBuzzPowerTools/wiki/Set-EntityReference)
* Set-EntityReferenceAlternateKey  \<NEW\>
* [Set-OptionSetValue](https://github.com/CRMBuzz/CRMBuzzPowerTools/wiki/Set-OptionSetValue)
* Set-RecordImageProfile  \<NEW\>
* [Write-ToFile](https://github.com/CRMBuzz/CRMBuzzPowerTools/wiki/Write-ToFile)


##Business Management (3)##
* Close Case
* Lose Opportunity
* Won Opportunity


##CRM Core Objects (15)##
* [New-OrganizationConnection](https://github.com/CRMBuzz/CRMBuzzPowerTools/wiki/New-OrganizationConnection)
* Add-NoteAttachment
* Assign-Record
* Convert-Lead
* Create-Record
* Delete-Record
* [ConvertTo-DateTime] (https://github.com/CRMBuzz/CRMBuzzPowerTools/wiki/ConvertTo-DateTime)
* [Get-EntityReferenceByName](https://github.com/CRMBuzz/CRMBuzzPowerTools/wiki/Get-EntityReferenceByName)
* Get-EntityReferenceByTwoFields  \<NEW\>
* Get-NoteAttachments
* Merger-CALRecords
* Retrieve-Record
* Set-Upsert  \<NEW\>
* Update-ActivityStatusAndState
* Update-Record
* Update-StatusAndState


##Audit (3)##
* Get-AuditDetails  \<NEW\>
* Enable-AuditForEntity
* Enable-AuditForOrganization

##Entity Connections (9)##
* Associate-EntityConnectionRole
* Find-AssociatedEntityConnectionRole
* Get-AssociatedEntityConnectionRole
* Get-EntityConnectionRole
* New-EntityConnection
* New-EntityConnectionRole
* New-ReciprocalEntityConnectionRole
* Query-EntityConnections
* Query-EntityConnectionRoles

##Execute Multiple (3)##
* Execute-MultipleCreate   \<NEW\>
* Execute-MultipleUpdate   \<NEW\>
* Execute-MultipleDelete   \<NEW\>

##CRM Search (7)##
* ConverFrom-FetchXMLToQueryExpression
* ConvertFrom-QueryExpressionToFetchXML
* Search-EntityFull
* Search-EntityId
* Search-FetchXML
* Search-QueryExpression
* Search-QueryByAttribute

##CRM SharePoint Integration (5)##
* Get-SPSite
* Get-SPLocation
* New-SPLocationRecord
* Remove-DocumentManagement
* Enable-DocumentManagement

##CRM Solutions (8)##
* Export-AllWebResources
* Export-Solution
* Export-WebResource
* Import-Solution
* New-Publisher
* Publish-AllCustomizations
* Remove-Publisher
* Remove-Solution

##CRM Teams/Queues (7)##
* Add-ActivityToQueue
* Add-Queue
* Add-Team
* Add-UserToRecordAccessTeam   \<NEW\>
* Assign-QueueItemWorker
* Remove-UserToRecordAccessTeam  \<NEW\>
* Set-QueueEmail
* Set-TeamAssignRole
* Set-TeamAssignToFieldSecurityProfile
* Set-TeamAssignUser

##CRM Users (9)##
* [Add-User] (https://github.com/CRMBuzz/CRMBuzzPowerTools/wiki/Add_User)
* Get-User
* Get-UserAccessMembership  \<NEW\>
* Remove-AssignedRoleAndFieldSecurityProfile
* Set-UserApproveEmail
* [Set-UserAssignManager] (https://github.com/CRMBuzz/CRMBuzzPowerTools/wiki/Set-UserManager)
* Set-UserAssignRole
* Set-UserAssignRoleAndFieldProfile
* Update-UserAccessMode  \<NEW\>

##CRM Workflows (4)##
* Get-Workflow
* Invoke-BulkWorkflowProcess
* Invoke-Workflow
* New-BulkDeleteWorkflow \<NEW\>

##Other functions (4)##
* Add-AssemblyToGAC
* Add-SampleData
* Remove-SampleData
* Send-MailMessage (SMTP Configuration)



###Sample Scripts###

The framework will contain a download file with the sample scripts and how to use the scripts to automate the available task.

The scripts do not require to install the framework in your server every script can execute these scripts from your workstation, most of the scripts will apply to the Online and On-premises, except for cmdlets related to User creation, those cmdlets will be available only for OnPrem

###Framework Installation###

The framework will have an installation script that will help you install the core assembly and the related CRM SDK required assemblies, please take a look at the documentation page for more details

[CRMBuzzPowerTools Module Installation](https://github.com/CRMBuzz/CRMBuzzPowerTools/wiki/Module-Installation)

[Registering CRMBuzzPowerTools](http://www.crmbuzz.co/crmbuzzpowertoolsmoduleinstallation/)


###Source Code###

I will be including the Source Code and want to do a clean up before making it available, more details coming soon

Enjoy!

Abraham (Abe) Saldana.