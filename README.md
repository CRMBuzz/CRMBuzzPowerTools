# CRMBuzzPowerTools
CRM Administration and Configuration Automation framework and Tool


The administration and configuration task you can automate using PowerShell scripts, automate user creation, assign teams, queues, security roles, field level security reusing PowerShell scripts

The idea for the framework is to allow the administrator to have an automation option for the manual configuration and administration task using the flexibility of PowerShell and the extensions on the CRM SDK

The framework contains 75+ cmdlets included in a PowerShell Snapin, the cmdlets are extensions of .Net framework and CRM SDK objects you will be able to manipulate with ease. 

These are some of the group of cmdlets 

- CRM General Objects (15)
* Get-AssemblyVersion
* Get-EntityCollection
* Get-CRMEntity
* Get-ActivityParty
* New-ColumnSet
* New-ConditionExpression
* New-FilterExpression
* New-LinkEntity
* New-OrderExpression
* New-QueryExpression
* Set-OptionSetValue
* Set-EntityReference
* Write-ToFile
* New-Guid
* ConvertTo-Guid


- CRM Core Objects (14)
* Add-NoteAttachment
* Convert-Lead
* ConvertTo-DateTime
* Get-EntityReferenceByName
* Get-NoteAttachments
* Merger-CALRecords
* New-OrganizationConnection
* Assign-Record (Old-Set-Assign)
* Create-Record (Old Set-Create)
* Delete-Record (Old Set-Delete)
* Retrieve-Record (Old Set-Retrieve)
* Update-Record (Old Update-Entity)
* Update-StatusAndState
* Update-ActivityStatusAndState

- Audit (2)
* Enable-AuditForEntity
* Enable-AuditForOrganization

- Entity Connections (9)
* Associate-EntityConnectionRole
* Find-AssociatedEntityConnectionRole
* Get-AssociatedEntityConnectionRole
* Get-EntityConnectionRole
* New-EntityConnection
* New-EntityConnectionRole
* New-ReciprocalEntityConnectionRole
* Query-EntityConnections
* Query-EntityConnectionRoles

- CRM Search (7)
* ConverFrom-FetchXMLToQueryExpression
* ConvertFrom-QueryExpressionToFetchXML
* Search-EntityFull
* Search-EntityId
* Search-FetchXML
* Search-QueryExpression
* Search-QueryByAttribute

- CRM SharePoint Integration (5)
* Get-SPSite
* Get-SPLocation
* New-SPLocationRecord
* Remove-DocumentManagement
* Enable-DocumentManagement

- CRM Solutions (8)
* Export-AllWebResources
* Export-Solution
* Export-WebResource
* Import-Solution
* New-Publisher
* Publish-AllCustomizations
* Remove-Publisher
* Remove-Solution

- CRM Teams/Queues (7)
* Add-Queue
* Add-Team
* Add-ActivityToQueue
* Assign-QueueItemWorker
* Set-TeamAssignRole
* Set-TeamAssignToFieldSecurityProfile
* Set-TeamAssignToUser

- CRM Users (7)
* Add-User
* Get-User
* Get-UserAccessMembership
* Set-UserApproveEmail
* Remove-AssignedRoleAndFieldSecurityProfile
* Set-UserAssignSecurityRole
* Set-UserAssignSecurityRoleAndFieldSecurityProfile

- CRM Workflows (4)
* Get-Workflow
* Invoke-BulkWorkflowProcess
* Invoke-Workflow
* New-BulkDeleteWorkflow

More cmdlets in the near future
- Execute Multiple
- Marketing and Campaign cmdlets


Sample Scripts

The framework will contain a download file with the sample scripts and how to use the scripts to automate the available task.

The scripts do not require to install the framework in your server every script can execute these scripts from your workstation, most of the scripts will apply to the Online and On-premisses, except for cmdlets related to User creation, those cmdlets will be available only for On-Prem


Framework Installation

The framework will have an installation script that will help you install the core assembly and the related CRM SDK required assemblies, please take a look at the documentation page for more details
http://www.crmbuzz.co/crmbuzzpowertoolsmoduleinstallation/


Source Code

I will be including the Source Code and want to do a clean up before making it available, more details coming soon
