# TeamsUsageCollector

This ps module is intended to create usage reports related to Microsoft 365 products.

Pre-requisites

 - A global admin account in Azure to allow the app registration and permissions consent. The following permissions will be asked for consent:
                
                - Reports.Read.All
                - User.Read.All

- For big environments, a always-on server/machine.
- A local admin account with administrative privilegies and following GPOs:
    - Log on as a batch job (Start > Run > secpol.msc > Security Settings > Local Policies > User Rights Assignment > Log on as a batch job > check if user is in the list or is a member of a group in the list)
    - Network access: Do not allow storage of passwords and credentiais for network authentication (Start > Run > secpol.msc > Security Settings > Local Policies > Security Options > Network access: Do not allow storage of passwords and credentiais for network authentication > Disabled)
    
Setup

1. Download the last version of this module from releases page:
2. Open Windows PowerShell using the option Run as Administrator
3. Ajust the execution policy:

    Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope LocalMachine
    
4. Rum the follwoing cmdlet to import the module:

    Import-Module .\M365UsageCollector.psm1

5. After all information messages, use the following cmdlet to confirm that all cmdlets have been imported:

    Get-Module M365UsageCollector

    ModuleType Version    Name                                ExportedCommands
    ---------- -------    ----                                ----------------
    Script     0.0        M365UsageCollector                  {Get-M365LicenseSkuReport, Get-TeamsUsageReport, New-M365UsageCollectorAppRegistration, New-M365UsageCollectorJob}


7. Run the following cmdlet to create the application registration in Azure and authenticate using Azure AD global admin credentials:

   New-M365UsageCollectorAppRegistration


8.	If everything worked fine, following information will be shown. Make sure you take a note of Application ID, Tenant ID and Client Secret. We gonna use them shortly.

      Application Teams Usage Collector created successfully in your tenant. Take a note of the following information. If you lost one of them, ask you tenant admin to get it for you in Azure AD. Also, copy the http link for admin permission consent.
            AppId: 78f8538a-aaaa-4c6f-9b92-05e0d228129a
            TenantId: cdcae3ff-aaaa-4732-9cf5-1e33db81acf1
            ClientSecret: LySeJaaaaZCO7xcYFCXck+KodLxojoI7pHJbDUu4n+I=


9.	Open any browser and paste the http link shown int he previous step. Proceed with authentication and consent the listed permissions.

10.	After permission consent, you'll be redirected to a blank page. It is fine, you can close it.

11.	Back to PowerShell Windows, use the following cmdlet with the information copied in step 8 to create a scheduled task that will run in the background to build the Teams Usage Score report:

  New-M365UsageCollectorJob -AppId 78f8538a-aaaa-4c6f-9b92-05e0d228129a -TenantId cdcae3ff-aaaa-4732-9cf5-1e33db81acf1 -ClientSecret LySeJaaaaZCO7xcYFCXck+KodLxojoI7pHJbDUu4n+I= -ReportMode "AsJob"

12. A credential will be asked. This one should be that one with local privilegies to run as a batch job.

14. Go to the Task Scheduler and confirm that the task M365UsageCollector is up and runnning.

15. You can check task progress in the following path: C:\Program Files\WindowsPowerShell\Modules\M365-Usage-Collector\0.0.5. A log per day will be created.

16. The SKU report wasn't yet added to the scheduled task. You can run it and test using the following cmdlet.

  Get-M365SkuLicenseReport -Export:$true
  
All report files are created in the following path with .csv extension. After all tasks finished, 3 reports are expected (you can have more if a task runs more than once):

LicenseReport.csv - enabled/consumed SKU report
M365UsageReport_Detailed_dd-MM-yyyy_hh-mm-ss.csv - Per user Teams usage
M365UsageReport_Summary_dd-MM-yyyy_hh-mm-ss.csv - Teams usage score

Known Issues:

 - Every time I close PowerShell when re-open, if I run Import-Module is says "no valid module file was found".

  It happens because it is not yet a module per se. Make sure you import the module using the entire psm1 file path as follows:

    Import-Module C:\Temp\M365UsageCollector.psm1
  
 - The task was created successfully but it never runs even if I manually try to run it.
 
    It usually happens when the user doesn't have the right to jobs as a batch or to store a password during the creation of a scheduled task. Confirm that the following GPOs are configured as expected:
    
    - Log on as a batch job (Start > Run > secpol.msc > Security Settings > Local Policies > User Rights Assignment > Log on as a batch job > check if user is in the list or is a member of a group in the list)
    - Network access: Do not allow storage of passwords and credentiais for network authentication (Start > Run > secpol.msc > Security Settings > Local Policies > Security Options > Network access: Do not allow storage of passwords and credentiais for network authentication > Disabled)



