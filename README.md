[![Board Status](https://dev.azure.com/dfoggi/7efa0583-4fd8-43cd-a1a0-be7d34b87975/e67f131d-eed2-42b0-9a27-1d63add15eba/_apis/work/boardbadge/a8082285-e8c0-4fc1-8d82-24f1f26b0c3e)](https://dev.azure.com/dfoggi/7efa0583-4fd8-43cd-a1a0-be7d34b87975/_boards/board/t/e67f131d-eed2-42b0-9a27-1d63add15eba/Microsoft.RequirementCategory/)

# TeamsUsageCollector

This ps module is intended to create usage reports related to Microsoft 365 products. Current version is focusing in Teams. New workloads could be added in the future.

## Pre-requisites

 - A global admin account in Azure to allow the app registration and permissions consent. The following permissions will be asked for consent:
                
                - Reports.Read.All
                - User.Read.All
                - Directory.Read.All

- For big environments, a always-on server/machine.
- A local admin account with administrative privilegies and following GPOs:
    - Log on as a batch job (Start > Run > secpol.msc > Security Settings > Local Policies > User Rights Assignment > Log on as a batch job > check if user is in the list or is a member of a group in the list)
    - Network access: Do not allow storage of passwords and credentiais for network authentication (Start > Run > secpol.msc > Security Settings > Local Policies > Security Options > Network access: Do not allow storage of passwords and credentiais for network authentication > Disabled)
    - If you already ran a previous version and gonna run a newer, it is expected that the module installation to created a new folder for the newer version in the module folder and also it will delete and create the scheduled task. As a last resort, if needed you can manually delete the scheduled task and run the step to create it again.
    
## Setup

1. Download the last version of this module from releases page.
2. Extract the entire .zip file.
3. Open Windows PowerShell using the option Run as Administrator.
4. Access the extracted root folder.
5. Adjust the execution policy:
```
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope LocalMachine
```

6. Run the follwoing cmdlet to import the module:
```
Import-Module .\M365UsageCollector.psm1
```

7. After all information messages, use the following cmdlet to confirm that all cmdlets have been imported:
```
    Get-Module M365UsageCollector

    ModuleType Version    Name                                ExportedCommands
    ---------- -------    ----                                ----------------
    Script     0.0        M365UsageCollector                  {Get-M365LicenseSkuReport, Get-TeamsUsageReport, New-M365UsageCollectorAppRegistration, New-M365UsageCollectorJob}
```

8. Run the following cmdlet to create the application registration in Azure and authenticate using Azure AD global admin credentials:
```
New-M365UsageCollectorAppRegistration
```

9. If everything worked fine, following information will be shown. Make sure you take a note of Application ID, Tenant ID and Client Secret. We gonna use them shortly.

      Application Teams Usage Collector created successfully in your tenant. Take a note of the following information. If you lost one of them, ask you tenant admin to get it for you in Azure AD. Also, copy the http link for admin permission consent.
            AppId: 78f8538a-aaaa-4c6f-9b92-05e0d228129a
            TenantId: cdcae3ff-aaaa-4732-9cf5-1e33db81acf1
            ClientSecret: LySeJaaaaZCO7xcYFCXck+KodLxojoI7pHJbDUu4n+I=

10. Open any browser and paste the http link shown int he previous step. Proceed with authentication and consent the listed permissions.

11. After permission consent, you'll be redirected to a blank page. It is fine, you can close it.

## Data collection

1. Back to PowerShell Windows, use the following cmdlet with the information copied in step 9 to create a scheduled task that will run in the background to build the Teams Usage Score report:

For smaller environments you can run directly in the current PowerShell session. The session can't be closed otherwise the report will be incompleted. If you choose this option, go to step 5.
```
Get-TeamsUsageReport -AppId 78f8538a-aaaa-4c6f-9b92-05e0d228129a -TenantId cdcae3ff-aaaa-4732-9cf5-1e33db81acf1 -ClientSecret LySeJaaaaZCO7xcYFCXck+KodLxojoI7pHJbDUu4n+I= -TeamsReportGroupByAttributes Department,Domain,officeLocation
```

For larger customers, you can schedule a task using the following cmdlet:
```
New-M365UsageCollectorJob -AppId 78f8538a-aaaa-4c6f-9b92-05e0d228129a -TenantId cdcae3ff-aaaa-4732-9cf5-1e33db81acf1 -ClientSecret LySeJaaaaZCO7xcYFCXck+KodLxojoI7pHJbDUu4n+I= -TeamsReportGroupByAttributes Department,Domain,officeLocation
```

Use the attributes you'd like to use as group by attributes in the parameter TeamsReportGroupByAttributes. Consider that each attribute you use will result in a file with a scorecard of teams usage grouped by the given attribute therefore resulting in a longer running time. You can use any combination of the following attributes comma separated: Department ,Domain, and officeLocation.

3. A credential will be asked. This one should be that one with local privilegies to run as a batch job.

4. Go to the Task Scheduler and confirm that the task M365UsageCollector is up and runnning.

5. You can check task progress using the log file in the following path: C:\Program Files\WindowsPowerShell\Modules\M365-Usage-Collector\vx.x.x. One log per day will be created.

6. All report files are placed in the following path with .csv extension. After all tasks finished, 3 reports are expected (you can have more if a task runs more than once):

 - C:\Program Files\WindowsPowerShell\Modules\M365-Usage-Collector\vx.x.x\

 - LicenseReport.csv - enabled/consumed SKU report

 - M365UsageReport_Detailed_dd-MM-yyyy_hh-mm-ss.csv - Per user Teams usage (user portion on UPN sanitized by default)

 - M365UsageReport_Teams_Summary_ByDepartment__dd-MM-yyyy_hh-mm-ss.csv - By department Teams usage score

 - M365UsageReport_Teams_Summary_ByDomainName_dd-MM-yyyy_hh-mm-ss.csv - By domain Teams usage score

 - M365UsageReport_Teams_Summary_ByOfficeLocation_dd-MM-yyyy_hh-mm-ss.csv - By department Teams usage score


## Known issues

 - Every time I close PowerShell then re-open, if I run Import-Module it says "no valid module file was found".

  It happens because it is not yet a module per se. Make sure you import the module using the entire psm1 file path as follows:
```  
    Import-Module C:\Temp\M365UsageCollector.psm1
```  
 
 - The task was created successfully but it never runs even if I manually try to run it.

    It usually happens when the user doesn't have the right to jobs as a batch or to store a password during the creation of a scheduled task. Confirm that the following GPOs are configured as expected:
    
    - Log on as a batch job (Start > Run > secpol.msc > Security Settings > Local Policies > User Rights Assignment > Log on as a batch job > check if user is in the list or is a member of a group in the list)
    - Network access: Do not allow storage of passwords and credentiais for network authentication (Start > Run > secpol.msc > Security Settings > Local Policies > Security Options > Network access: Do not allow storage of passwords and credentiais for network authentication > Disabled)
    
    If you run the cmdlet New-M365UsageCollectorJob, it will delete an existing task will create a new one. But it may fail to do so, then if the issue persists confirm that the task has all the information correctly, by checking the action parameter and if it points to the correct path and also if the path exists.
