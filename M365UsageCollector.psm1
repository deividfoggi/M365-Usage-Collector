<#########################################################################################
#
# The sample scripts are not supported under any Microsoft standard support
# program or service. The sample scripts are provided AS IS without warranty
# of any kind. Microsoft further disclaims all implied warranties including, without
# limitation, any implied warranties of merchantability or of fitness for a particular
# purpose. The entire risk arising out of the use or performance of the sample scripts
# and documentation remains with you. In no event shall Microsoft, its authors, or
# anyone else involved in the creation, production, or delivery of the scripts be liable
# for any damages whatsoever (including, without limitation, damages for loss of business
# profits, business interruption, loss of business information, or other pecuniary loss)
# arising out of the use of or inability to use the sample scripts or documentation,
# even if Microsoft has been advised of the possibility of such damages.
#
#########################################################################################

 .Synopsis
  Collects information related to Teams usage from a M365 tenant.

 .Description
  Display or export to .csv files information related to Teams usage in the contexts of user activity, user activation and licensing.
  This module allows you to either export to .csv files or have the information in the current PowerShell session to customize the output at your will.
#>

$currentVersion = "v0.0.7.beta.3"
#Creates an installation directory 
$installDir = "$env:ProgramFiles\WindowsPowerShell\Modules\M365-Usage-Collector\$($currentVersion)" #If changed, don't forget to updated it in the task schedule creation variable taskAction. Due to quotes, we can't use the install path variable there.
$modulePath = "$installDir\M365UsageCollector.psm1"

#Function to create a Write-Log file and register Write-Log entries
Function Write-Log{
    Param(
        [Parameter(Mandatory=$true)][string]$Status,
        [Parameter(Mandatory=$true)][string]$Message
    )
    
    $logName = $installDir + "\$(Get-Date -Format 'dd-MM-yyyy').log"

    $dayLogFile = Test-Path $logName
    
    $dateTime = Get-Date -Format dd/MM/yyyy-HH:mm:ss

    If($dayLogFile -eq $true){

        $logLine = $dateTime + "," + $Status + "," + $Message
        $logLine | Out-File -FilePath $logName -Append
    }
    Else
    {
        $header = "Date,Status,Message"
        $header | Out-File -FilePath $logName
        $logLine = $dateTime + "," + $Status + "," + $Message
        $logLine | Out-File -FilePath $logName -Append
    }
}

Write-Host -ForegroundColor Yellow "Checking if an installation directory is needed"
if(!(Test-Path $modulePath)){
    Write-Host -ForegroundColor Yellow "Creating the installation directory"
    try{
        New-Item -ItemType Directory -Path $installDir -Force -ErrorAction Stop
        Write-Host -ForegroundColor Yellow "Installation directory created. You can follow all activities in the .log files here: $($installDir)"
        Write-Log -Status "Info" -Message "Installation directory created. You can follow all activities in the .log files here: $($installDir)"
        Write-Host -ForegroundColor Yellow "Installing M365 Usage Collector Module"
        $moduleContent = Get-Content .\M365USageCollector.psm1 -ErrorAction Stop
        $moduleContent | Set-Content -Path $modulePath -Force -ErrorAction Stop
        Write-Log -Status "Info" -Message "Module successfully installed"
    }
    catch{
        Write-Warning "Error: $($_.Exception.Message)"
        Remove-Module M365UsageCollector -ErrorAction SilentlyContinue
        Exit
    }
}
else{
    Write-Host -ForegroundColor Yellow "Installation directory already created"
}

#Temporary set PSGallery as a trusted source to prevent in-screen prompt due to untrusted ps repository which will freeze scheduled task execution
Set-PSRepository PSGallery -InstallationPolicy Trusted

#Check if AzureADModule is already imported
Write-Log -Status "Info" -Message "Checking if AzureAD or AzureADPreview module is already installed"
if(!(Get-Command -Module AzureAD) -And !(Get-Command -Module AzureADPreview)){
    Write-Log -Status "Info" -Message "Azure module not yet installed. Trying to install it"
    #try to import azuread module
    try{
        Install-Module AzureAD -AllowClober -ErrorAction Stop
        Write-Log -Status "Info" -Message "AzureAD module installed sucessfully"
    }
    catch{
        Write-Warning $_.Exception.Message
        Write-Log -Status "Error" -Message "Error tyring to install AzureAD Module: $($_.Exception)"
        try{
            Write-Host -ForegroundColor Yellow "Trying to Install AzureADPreview module"
            Install-Module AzureADPreview -ErrorAction Stop
            Write-Log -Status "Info" -Message "AzureADPreview module installed sucessfully"
        }
        catch{
            Write-Warning $_.Exception.Message
            Write-Log -Status "Error" -Message "Error tyring to install AzureADPreview Module: $($_.Exception)"
            Exit
        }
    }
}else{
    Write-Log -Status "Info" -Message "Azure module already installed. Trying to import it"
    try{
        Import-Module AzureAD -ErrorAction Stop
        Write-Log -Status "Info" -Message "AzureAD module imported sucessfully"
    }
    catch{
        Write-Warning $_.Exception.Message
        Write-Log -Status "Error" -Message "Error tyring to import AzureAD Module: $($_.Exception)"
        try{
            Write-Host -ForegroundColor Yellow "AzureADPreview module already installed. Trying to import it"
            Import-Module AzureADPreview -ErrorAction Stop
            Write-Log -Status "Info" -Message "AzureADPreview module imported sucessfully"
        }
        catch{
            Write-Warning $_.Exception.Message
            Write-Log -Status "Error" -Message "Error tyring to import AzureADPreview Module: $($_.Exception)"
            Exit
        }
    }
}

#Set PSGallery back to default config
Set-PSRepository PSGallery -InstallationPolicy Untrusted

#Function to connect to Azure AD
Function ConnectAzureAD{
    #Try to get tenant details as a way to check if AzureAD module is already connected. If not connect throws an exception to be catched and connect to Azure AD
    try{
        Get-AzureADTenantDetail -ErrorAction Stop | Out-Null
        Write-Host "Azure AD already connected" -ForegroundColor Yellow
        Write-Log -Status "Info" -Message "Azure AD already connected"
    }
    catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException]
    {
        #If a 'need auth' exception then connect to Azure AD asking for credentials
        Write-Host -ForegroundColor Yellow "Connecting to Azure AD"
        Write-Log -Status "Info" -Message "Connecting to Azure AD"
        Connect-AzureAD
    }
}

#Beta function to parse a batch of users in a report
Function New-M365UsageParseJob{
    Param(
        [array]$UserList,
        [array]$TeamsUserActivityUserDetail,
        [array]$Office365ActiveUserDetail
    )

    Write-Log -Status "Info" -Message "Creating a runspace pool with 10 threads limit"

    #Creates a Runspace pool limited to 10 threads
    $RunspacePool = [runspacefactory]::CreateRunspacePool(1,10)
    $RunspacePool.Open()
    #Define an array to store all runspaces
    $Jobs = @()

    #Split all users list in smaller chunks
    $usersChunks = Split-Array -Array $UserList -ObjectLimit 1000

    #Increment control variable to be used to name temp files
    $i = 1

    #For each chunk of users in users chunks
    foreach($users in $usersChunks){
        $ParamList = @{
            FileName = "$installDir\temporary_$i.csv"
            Users = $users
            TeamsUserActivityUserDetail = $teamsUserActivityUserDetail
            Office365ActiveUserDetail = $office365ActiveUserDetail
        }
        Write-Log -Status "Info" -Message "Creating runspace $($i) of $($usersChunks.length)"
        #Create the powershell runspace
        $PowerShell = [powershell]::Create()
        #Add runspace into runspace pool
        $PowerShell.RunspacePool = $RunspacePool
        #Define the script block of the current runs space using the current users chunk
        $PowerShell.AddScript({
            param ($FileName,$Users,$TeamsUserActivityUserDetail,$Office365ActiveUserDetail)
            #Empty array to append all users objects
            $parsedUserList = @()
            #For each user in array users
            foreach($user in $Users){
                #Extract from the teams activity user detail report the current user findings
                $userteamsUserActivityUserDetail = $TeamsUserActivityUserDetail | Where-Object{$_.'User Principal Name' -eq $user.UserPrincipalName}
                #Extract from the teams active user detail report the current user findings
                $office365ActiveUserDetailUser = $Office365ActiveUserDetail | Where-Object{$_.'User Principal Name' -eq $user.UserPrincipalName}
                #Create a ps custom object to store current user findings
                $userObj = [PSCustomObject] @{
                    #Sanitize UserPrincipalName (keep domain only which can be used later in group by report) and DisplayName to remove PII
                    UserPrincipalName = $user.UserPrincipalName.Split("@")[1]
                    #Fill the following attributes accordingly
                    Department = $user.Department
                    OfficeLocation = $user.officeLocation
                    IsDeleted = $office365ActiveUserDetailUser.'Is Deleted'
                    DeletedDate = $office365ActiveUserDetailUser.'Deleted Date'
                    HasTeamsLicense = $office365ActiveUserDetailUser.'Has Teams License'
                    TeamsLastActivityDate = $office365ActiveUserDetailUser.'Teams Last Activity Date'
                    TeamsLicenseAssignDate = $office365ActiveUserDetailUser.'Teams License Assign Date'
                    AssignedProducts = $office365ActiveUserDetailUser.'Assigned Products'
                    TeamChatMessageCount = $userteamsUserActivityUserDetail.'Team Chat Message Count'
                    PrivateChatMessageCount = $userteamsUserActivityUserDetail.'Private Chat Message Count'
                    CallCount = $userteamsUserActivityUserDetail.'Call Count'
                    MeetingCount = $userteamsUserActivityUserDetail.'Meeting Count'
                    ReportPeriod = $userteamsUserActivityUserDetail.'Report Period'
                }
                #Append current object into the array
                $parsedUserList += $userObj
            }
            #Export objects to a temporary file
            $parsedUserList | Export-Csv $FileName -NoTypeInformation -Encoding UTF8
        })
        Write-Log -Status "Info" -Message "Invoking runspace $($i) of $($usersChunks.length)"
        #Add the parameter list into current runspace
        $PowerShell.AddParameters($ParamList)
        #Invoke the execution of current runspace and add it to an array of runspaces to allow tracking
        $jobs += $PowerShell.BeginInvoke()
        #Increment the control variable
        $i++
    }
    #While at least on job is not completed, wait for 5 seconds and check again. Script will not move further until all jobs are completed.
    While($Jobs.IsCompleted -contains $false){Start-Sleep -Seconds 5}
    Write-Log -Status "Info" -Message "All PowerShell runspaces are completed"
}

#Function to split an array of objects
Function Split-Array{
    Param(
    $Array,
    $ObjectLimit
)
    #Define the chunk size
    [int]$blockLimit = $ObjectLimit
    #Math the number of jobs/chunks
    $numberOfJobs = [math]::Floor($Array.length / $blockLimit)
    #Get the rest/mod for the last one
    $lastJobCount = $Array.Length % $blockLimit
    #interaction control variable
    $i = 1
    #Define an array to store all the arrays
    $result = @()
    #Do the calculations until $i variable is greater than the number of chunks plus 1
    Do{
        #If the interaction variable is equal to the number of chunks plus 1 this is the last chunk
        if($i -eq ($numberOfJobs + 1)){
            #The variable of the first object in the current chunk becomes the rest/mod all chunks divided the chunk size
            $varFirst = $lastJobCount
            #The variable of the number of objects to skip becomes the number of objects multipled by current interaction value minus 1
            $varSkip = $numberOfObj * ($i - 1)
        }else{
            #If the interaction variable is not equanto to the number of chunks plus 1 then this is not the last chunk
            #The number of objects becomes the chunk size
            $numberOfObj = $blockLimit
            #The first object is the number of objects multiplied by the current interaction
            $varFirstTmp = $numberOfObj * $i
            #The number of objects to skip is the number of the first object minus the number of objects
            $varSkip = $varFirstTmp - $numberOfObj
            #The very first object is the number of objects
            $varFirst = $numberOfObj
        }

        #Append to the main array the array of objects considering the first and last objects as the current chunk of objects (Starting in one specific and skipping some accordingly)
        $result+=,@($Array | Select-Object -First $varFirst -Skip $varSkip)

        $i++
    }
    Until($i -gt $numberOfJobs + 1)
    #Return an array of arrays as a result
    return,$result
}
#Function to merge all temporary report files into one last file
Function Join-TemporaryFiles{
    Param(
        $ReportPath
    )

    #Define the temporary file path pattern
    $files = Get-Item -Path "$($installDir)\temporary_*.csv"
    #Define the name of the final report based on report name parameter

    #For each temporary file generated by New-M365UsageParseJob function
    ForEach($file in $files){
        #Imports the current file
        $currentFile = Import-Csv $file
        #Append current file into the final report
        $currentFile | Export-Csv $ReportPath -Append -NoTypeInformation -Encoding UTF8
    }
    #Remove temporary files
    $files | Remove-Item -Force -Confirm:$false
}

Function New-M365UsageCollectorJob{
    param(
        [Parameter(Mandatory=$true)]$AppId,
        [Parameter(Mandatory=$true)]$TenantId,
        [Parameter(Mandatory=$true)]$ClientSecret,
        #[Parameter(Mandatory=$true)]$ReportMode
        [Parameter(Mandatory=$false)][array]$TeamsReportGroupByAttributes
    )

    #Scheduled task settings
    $taskName = "M365UsageCollector"
    $taskActionArgument = "-File `"C:\Program Files\WindowsPowerShell\Modules\M365-Usage-Collector\$currentVersion\temp.ps1`""
    $taskAction = New-ScheduledTaskAction -Execute 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe' -Argument $taskActionArgument
    $taskDescription = "Collect usage data from Microsoft 365 cloud"
    $taskCredentials = Get-Credential -Message "Scheduled task credential to run once"
    $taskPrincipal = New-ScheduledTaskPrincipal -UserId $taskCredentials.UserName -LogonType ServiceAccount -RunLevel Highest
    $taskSettings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Days 7)
    $task = New-ScheduledTask -Action $taskAction -Principal $taskPrincipal -Settings $taskSettings -Description $taskDescription
    $tempJob = "Import-Module '$modulePath';Get-TeamsUsageReport -AppId $AppId -TenantId $TenantId -ClientSecret $ClientSecret -TeamsReportGroupByAttributes $($TeamsReportGroupByAttributes -join ",");Remove-Item '$installDir\temp.ps1' -Confirm:`$false"
    $tempJob | Set-Content "$installDir\temp.ps1" -Force

    #Try to check if the shceduled task already exists
    try{        
        #If the task not exists yet
        if(!(Get-ScheduledTask N365UsageCollector -ErrorAction Ignore)){
            #Try to register the task
            Register-ScheduledTask -TaskName $taskName -InputObject $task  -ErrorAction Stop
            Write-Log -Status "Info" -Message "Task user and action configured"
            #Try to set task credentials
            Set-ScheduledTask -TaskName $taskName -User $taskPrincipal.UserId -Password $taskCredentials.GetNetworkCredential().Password -ErrorAction Stop
            Write-Log -Status "Info" -Message "Task principal configured"
        }
        #If the task exists already
        else{
            Unregister-ScheduledTask -TaskName $taskName
            #ry to register the task
            #Try to register the task
            Register-ScheduledTask -TaskName $taskName -InputObject $task  -ErrorAction Stop
            Write-Log -Status "Info" -Message "Task user and action configured"
            #Try to set task credentials
            Set-ScheduledTask -TaskName $taskName -User $taskPrincipal.UserId -Password $taskCredentials.GetNetworkCredential().Password -ErrorAction Stop
            Write-Log -Status "Info" -Message "Task principal configured"
        }
        Write-Log -Status "Info" -Message "Task $($taskName) configured successfully to run with user $($taskCredentials.UserName)"
        #Try to start the task
        Start-ScheduledTask -TaskName $taskName -ErrorAction Stop
        Write-Log -Status "Info" -Message "Task $($taskName) started successfully. For more details and status, use Windows Task Scheduler"
    }
    catch{
        Write-Warning $_.Exception.Message
        Write-Log -Status "Error either to create or set the task: " -Message $_.Exception
    }
}

#Function to group teams report by a specific attribute such as department or domain
Function Group-TeamsReportBy{
    Param(
        $GroupByAttribute,
        $TeamsUsersList,
        $UniqueGroupByAttributeList,
        $TimeSpan
    )

    Write-Log -Status "Info" -Message "Start to group users by $($GroupByAttribute)"

    #Parse string dates to date/time format
    $teamsMauStartDate = (Get-Date).AddDays(-($TimeSpan.substring($TimeSpan.length -2, 2)))
    $teamsMauStartDate = [datetime]::ParseExact($teamsMauStartDate.ToString('yyyy-MM-dd'), 'yyyy-MM-dd', $null)

    #Group users with teams license
    $usersPerAttrWithTeams = $TeamsUsersList | Where-Object{$_.HasTeamsLicense -eq "TRUE"} | Group-Object $GroupByAttribute

    <#
    #Group users without teams license
    $usersPerAttrWithoutTeams = $TeamsUsersList | Where-Object{$_.HasTeamsLicense -ne "TRUE"} | Group-Object $GroupByAttribute
    #>
    #Group users with teams MAU first collecting those without empty attribute
    $usersPerAttrWithTeamsMAU = $TeamsUsersList | Where-Object{$null -ne $_.TeamsLastActivityDate -and $_.TeamsLastActivityDate -ne ""}
    #Group users with teams MAU considering only those with attribute filed + last activity date newer than the date of TimeSpan variable which is the report period (possible report periods: D30, D60, D90, D180)
    $usersPerAttrWithTeamsMAU = $usersPerAttrWithTeamsMAU | Where-Object{[datetime]::ParseExact($_.TeamsLastActivityDate,'yyyy-MM-dd', $null) -gt $teamsMauStartDate} | Group-Object $GroupByAttribute
    <#
    #Group users without teams MAU first collecting those with empty attribute
    $usersPerAttrWithoutTeamsMAU = $TeamsUsersList | Where-Object{$null -eq $_.TeamsLastActivityDate -or $_.TeamsLastActivityDate -eq ""}
    #Group users without teams MAU now collecting those without empty attribute but last activity date older than the date of TimeSpan variable which is the report period
    $usersPerAttrWithoutTeamsMAUOlder = $TeamsUsersList | Where-Object{$null -ne $_.TeamsLastActivityDate -and "" -ne $_.TeamsLastActivityDate}
    #Group users without teams MAU now collecting those without empty attribute + last activity date older than the date of TimeSpan variable which is the report period
    $usersPerAttrWithoutTeamsMAUOlder = $usersPerAttrWithoutTeamsMAUOlder | Where-Object{[datetime]::ParseExact($_.TeamsLastActivityDate,'yyyy-MM-dd', $null) -lt $teamsMauStartDate} | Group-Object $GroupByAttribute
    #Summing upp both variables of users without Teams MAU
    $usersPerAttrWithoutTeamsMAU += $usersPerAttrWithoutTeamsMAUOlder
    #>
    #Group users with teams meetings MAU
    $usersPerAttrWithTeamsMeetingMAU = $TeamsUsersList | Where-Object{$_.MeetingCount -gt 0} | Group-Object $GroupByAttribute

    #Creates an array to append ps custom objects as a final scorecard
    $scoreReport = @()

    #Foreach unique attribute found in users endpoint api, uses the grouped objects above to build up a teams usage score
    foreach($attr in $UniqueGroupByAttributeList){
        #Check if variable is not empty
        if($null -eq $attr){
            $attr = ""
        }

        #Creates a ps custom object for the current item
        $obj = [PSCustomObject]@{
            $($GroupByAttribute) = $attr
            UserCount = ($TeamsUsersList | Where-Object{$_.($GroupByAttribute) -eq $attr} | Measure-Object).Count
            HasTeamsLicense = (($usersPerAttrWithTeams | Where-Object{$_.Name -eq $attr}).Group | Measure-Object).Count
            <#
            HasNoTeamsLicense = (($usersPerAttrWithoutTeams | Where-Object{$_.Name -eq $attr}).Group | Measure-Object).Count
            #>
            TeamsMAU = (($usersPerAttrWithTeamsMAU | Where-Object{$_.Name -eq $attr}).Group | Measure-Object).Count
            <#
            HasNoTeamsMAU = (($usersPerAttrWithoutTeamsMAU | Where-Object{$_.Name -eq $attr}).Group | Measure-Object).Count
            #>
            TeamsMeetingsMAU = (($usersPerAttrWithTeamsMeetingMAU | Where-Object{$_.Name -eq $attr}).Group | Measure-Object).Count
        }
        $scoreReport += $obj
    }
    Write-Log -Status "Info" -Message "Finished to group users by $($GroupByAttribute)"

    #Return all objects appended in one array
    return $scoreReport
}

#Function to create an application registration in Azure AD to be used to connect to Graph API
Function New-M365UsageCollectorAppRegistration {
        #Connects to Azure AD
        ConnectAzureAD

        #Define permissions needed
        $appPerms = 'Reports.Read.All','User.Read.All','Directory.Read.All'
        #Define a 'blank' consent reply URL
        $replyUrls = "https://login.microsoftonline.com/common/oauth2/nativeclient"

        #Uses Azure AD module to get Azure AD service Principals related to Microsoft Graph
        $msGraphService = Get-AzureADServicePrincipal -All $true -Filter "DisplayName eq 'Microsoft Graph'"
        
        #Filters only those permissions in appPerms array
        $permissions = $msGraphService.AppRoles.Where({$_.Value -in $appPerms})
        
        #Crates a new app permission object RequiredSourceAccess
        $msGraphResourceAccess = New-Object -TypeName "Microsoft.Open.AzureAd.Model.RequiredResourceAccess"

        #Fine app permission object ResourceAppId based on MS Graph service principal app id
        $msGraphResourceAccess.ResourceAppId = $msGraphService.AppId

        #for each service principal permission related to microsoft graph in array appPerms, creates a new object ResourceAccess and adds to an array
        foreach($permission in $permissions){
            $appPermissions = new-object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList $permission.Id,"Role"
            $msGraphResourceAccess.ResourceAccess += $appPermissions
        }

        # Define the name of the application registration
        $appName = "Teams Usage Collector"

        #Define the parameters to be used in cmdlet to created the app reg, adding display name, permissions in object RequeredRsourceAccess and blank consent reply URL
        $appCreationParameters = @{
            "DisplayName" = $appName;
            "RequiredResourceAccess" = @($msGraphResourceAccess)
            "ReplyUrls" = $replyUrls
        }

        #Creates the new app registration in Azure AD
        try{
            Write-Log -Status "Info" -Message "Trying to create Azure AD Application Registration"
            $appCreated = New-AzureADApplication @appCreationParameters
            Write-Log -Status "Info" -Message "Azure AD Application Registration successfully created"
        }
        catch{
            Write-Log -Status "Error" -Message "Unable to create Azure AD Application Registration: $($_.Exception.Message)"
        }
        
        #Define start and end date for client secret duration
        $startDate = Get-Date
        $endDate = $startDate.AddYears(3)

        #Creates a new client secret
        try{
            Write-Log -Status "Info" -Message "Trying to create password secret"
            $appReg = New-AzureADApplicationPasswordCredential -ObjectId $appCreated.ObjectId -CustomKeyIdentifier "clientSecret" -StartDate ([DateTime]::Now) -EndDate $endDate -ErrorAction Stop
            Write-Log -Status "Info" -Message "Password secret created successfully"
        }
        catch{
            Write-Log -Status "Error" -Message "Unable to create password secret: $($_.Exception.Message)"
        }
        

        #Creates a custom ps object to store information related to application registration that will be printed out in the screen
        $objAppReg = [PSCustomObject]@{
            ObjectId = $appCreated.ObjectId
            AppId = $appCreated.AppId
            TenantId = (Get-AzureADTenantDetail).ObjectId
            ClientSecret = $appReg.Value
        }

        #Write in the screen the details of the application registration including AppId, TenantId and Client Secret. Also prints out the URL to be used for the Azure Global Admin to consent the permissions
        Write-Host "Application $appName created successfully in your tenant. Take not of the following information. If you lost one of them, ask you tenant admin to get it for you in Azure AD:
            AppId: $($objAppReg.AppId)
            TenantId: $($objAppReg.TenantId)
            ClientSecret: $($objAppReg.ClientSecret)
        
        " -Foreground Green
        Write-Host "Azure Admin should consent using the following link:" -Foreground Yellow
        Write-Host ("https://login.microsoftonline.com/$($objAppReg.TenantId)/adminconsent?client_id=$($objAppReg.AppId)&redirect_uri=$($replyUrls)") -Foreground Yellow
}

#Function to get Azure AD token
Function Get-AzureADToken{
    Param(
        [Parameter(Mandatory=$true)]$AppId,
        [Parameter(Mandatory=$true)]$TenantID,
        [Parameter(Mandatory=$true)]$ClientSecret
    )

    $stringUrl = "https://login.microsoftonline.com/" + $tenantId + "/oauth2/v2.0/token"
    $postData = "client_id=" + $AppId + "&scope=https://graph.microsoft.com/.default&client_secret=" + $ClientSecret + "&grant_type=client_credentials"
    try{
        $accessToken = Invoke-RestMethod -Method post -Uri $stringUrl -ContentType "application/x-www-form-urlencoded" -Body $postData  -ErrorAction Stop
        Write-Host -ForegroundColor Yellow "Access token acquired."
        return $accessToken
    }
    catch{
        $errorDescription = $_ | ConvertFrom-Json
        Write-Warning $errorDescription.error
        Write-Host $errorDescription.error_description -ForegroundColor Yellow
    }
}

#Function to submit HTTP requests to Graph API using protocol
Function Send-GraphRequest{
    Param(
    [Parameter(Mandatory=$true)]$Method,
    [Parameter(Mandatory=$false)]$BearerToken,
    [Parameter(Mandatory=$false)]$Path,
    [Parameter(Mandatory=$false)]$Beta
    )

    #Check if beta parameter is true to use beta endpoint
    If($Beta -eq $true){
        $Uri = "https://graph.microsoft.com/Beta" + $Path
    }
    #if not beta, then use prod endpoint
    else{
        $Uri = "https://graph.microsoft.com/v1.0" + $Path
    }

    #Try to send a request to Graph API endpoint using protocol
    try{
        #Creates an empty array to store the appended results in case of paging
        $queryResults = @()

        #Do the HTTP request against the API endpoint until there is no @odata.nextLink in the response meaning no further pages
        do{
            #Stores the rest method request agains API in a variable
            $request = Invoke-RestMethod -Method $Method -Headers @{Authorization = "Bearer $($BearerToken)"} -Uri $Uri -ContentType 'application/json' -ErrorAction Stop

            #If varaible has a value property with content means there is results/payload
            if($request.value){
                #Adds the result/payload objects in the array
                $queryResults += $request.value
            }
            else{
                #If not, adds the entire response in the array
                $queryResults += $request
            }
            #Stores the @odata.nextLink in the variable used to check if there is further pages
            $Uri = $request.'@odata.nextLink'
        } until (!($Uri))
        #Returns the array containing all pages appended
        return $queryResults
    }
    catch{
        $errorDescription = $_ | ConvertFrom-Json
        Write-Warning $errorDescription.error
        Write-Host $errorDescription.error_description -ForegroundColor Yellow
    }
}

#Function to get a license sku report using Azure AD module
Function Get-M365LicenseSkuReport{
    Param(
        [Parameter(Mandatory=$false)]$Export
    )

    $installDir = "$env:ProgramFiles\WindowsPowerShell\Modules\M365-Usage-Collector\$($currentVersion)"

    #Connects to Azure AD
    ConnectAzureAD

    #Store all subscribed SKUs in a variable
    $licensesRequest = Get-AzureADSubscribedSku | Select-Object SkuPartNumber,*Units

    #Creates an empy array to build append custom ps objects
    $licenseReport = @()

    #For each subscribed SKU found
    foreach($sku in $licensesRequest){

        #Crates a new ps custom object to store 3 attributes as follows
        $objLicense = [PSCustomObject] @{
            "Sku" = $sku.SkuPartNumber
            "EnabledUnits" = $sku.PrepaidUnits.Enabled
            "ConsumedUnits" = $sku.ConsumedUnits            
        }
        #Append current object in the array
        $licenseReport += $objLicense
    }
    #If Export parameter selected by user then exports the array into a csv file
    if($Export -eq $true){
        $licenseReport | Export-Csv "$installDir\LicenseReport.csv" -NoTypeInformation
    }
    #If no Export parameter then print out in the screen
    else{
        return $licenseReport | Format-Table
    }
}

#Function to get teams usage report from Graph Reports API
Function Get-TeamsUsageReport{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]$AppId,
        [Parameter(Mandatory=$true)]$TenantId,
        [Parameter(Mandatory=$true)]$ClientSecret,
        #[Parameter(Mandatory=$true)]$ReportMode,
        [Parameter(Mandatory=$true)]
        [ValidateSet("Department","Domain","officeLocation")]
        [array]$TeamsReportGroupByAttributes,
        [Parameter(Mandatory=$false)]
        [ValidateSet("D7","D30","D90","D180")]
        $TimeSpan = "D30"
    )

    #Register in a variable the start datetime for statistics purposes
    $stopWatchStart = Get-Date
    Write-Log -Status "Info" -Message "Teams Usage Report execution started at $($stopWatchStart)"
    
    #Uses EscapeDataString function to prevent an issue that replaces all + sign in the client secret string with a blank space
    $ClientSecret = [System.Uri]::EscapeDataString($ClientSecret)

    #Get an Azure AD token using app reg info
    $accessToken = (Get-AzureADToken -AppId $AppId -TenantId $TenantId -ClientSecret $ClientSecret).access_token
    
    #Define summary report name by department
    $m365UsageReportSummaryNameByDepartment = "M365UsageReport_Teams_Summary_ByDepartment"
    #Detailed report path
    $summaryReportPathByDepartment = $installDir + "\$($m365UsageReportSummaryNameByDepartment)_$(Get-Date -Format 'dd-MM-yyyy_hh-mm-ss').csv"
    #Define summary report name by domain name
    $m365UsageReportSummaryNameByDomainName = "M365UsageReport_Teams_Summary_ByDomainName"
    #Detailed report path
    $summaryReportPathByDomainName = $installDir + "\$($m365UsageReportSummaryNameByDomainName)_$(Get-Date -Format 'dd-MM-yyyy_hh-mm-ss').csv"
    #Define summary report name by officeLocation
    $m365UsageReportSummaryNameByofficeLocation = "M365UsageReport_Teams_Summary_ByOfficeLocation"
    #Detailed report path
    $summaryReportPathByofficeLocation = $installDir + "\$($m365UsageReportSummaryNameByofficeLocation)_$(Get-Date -Format 'dd-MM-yyyy_hh-mm-ss').csv"
    #Define detailed report name
    $m365UsageReportDetailedName = "M365UsageReport_Detailed"
    #Detailed report path
    $detailedReportPath = $installDir + "\$($m365UsageReportDetailedName)_$(Get-Date -Format 'dd-MM-yyyy_hh-mm-ss').csv"
    
    Write-Log -Status "Info" -Message "Starting the request for reports: Teams User Activity Detail and Office 365 Active User Detail"
    #Send graph api requests against reports API to get teams reports considering a 30 days time span
    $teamsUserActivityUserDetail = (Send-GraphRequest -Method Get -BearerToken $accessToken -Path "/reports/getTeamsUserActivityUserDetail(period='$TimeSpan')")|ConvertFrom-Csv
    Write-Log -Status "Info" -Message "Finished the collection of Teams User Activity Detail report"
    $office365ActiveUserDetail = (Send-GraphRequest -Method Get -BearerToken $accessToken -Path "/reports/getOffice365ActiveUserDetail(period='$TimeSpan')")|ConvertFrom-Csv
    Write-Log -Status "Info" -Message "Finished the collection of Office 365 Active User Detail report"

    Write-Log -Status "Info" -Message "Starting the request for all users in Azure AD"
    #Send graph api request against users api to get UPN and Department in order to parse department agains users in the reports collected above
    $users = Send-GraphRequest -Method Get -BearerToken $accessToken -Path "/users?`$select=userPrincipalName,accountEnabled,city,companyName,country,department,jobTitle,officeLocation,postalCode,state,streetAddress,usageLocation&`$top=999"
    Write-Log -Status "Info" -Message "Finish the collection of all users in Azure AD"

    #Beta function - multi-thread
    Write-Log -Status "Info" -Message "Start to parse all reportings into one"
    New-M365UsageParseJob -UserList $users -teamsUserActivityUserDetail $teamsUserActivityUserDetail -office365ActiveUserDetail $office365ActiveUserDetail
    #Run the function to join all temporary files generated by the multi-thread function New-M365UsageParseJob into the final report
    Join-TemporaryFiles -ReportPath $detailedReportPath
    Write-Log -Status "Info" -Message "Finished the user report parsing"
    
    #Get all users in final report
    $parsedUserList = Import-Csv $detailedReportPath

    Function New-TeamsDepartmentScorecard{

        Write-Log -Status "Info" -Message "Start to extract unique deparments"
        #Extract unique department strings from users endpoint result
        $departments = ($users | Select-Object Department -Unique).department
        Write-Log -Status "Info" -Message "Finish to extract unique departments"
        $screenReportByDepartment = Group-TeamsReportBy -GroupByAttribute "Department" -TeamsUsersList $parsedUserList -UniqueGroupByAttributeList $departments -TimeSpan $TimeSpan
        return $screenReportByDepartment
    }

    Function New-TeamsDomainScoreCard{
        Write-Log -Status "Info" -Message "Starting the request from all domains in Azure AD"
        $domains = Send-GraphRequest -Method Get -BearerToken $accessToken -Path "/domains"
        Write-Log -Status "Info" -Message "Start to extract unique domains"
        #Extract unique domains strings from domains endpoint result
        $domains = ($domains | Select-Object id -Unique).id
        Write-Log -Status "Info" -Message "Finish to extract unique domains"
        $screenReportByDomain = Group-TeamsReportBy -GroupByAttribute "UserPrincipalName" -TeamsUsersList $parsedUserList -UniqueGroupByAttributeList $domains -TimeSpan $TimeSpan
        return $screenReportByDomain
    }

    Function New-TeamsOfficeLocationScoreCard{
        Write-Log -Status "Info" -Message "Start to extract unique officeLocation"
        #Extract unique office locations strings from users endpoint result
        $officeLocation = ($users | Select-Object officeLocation -Unique).officeLocation
        Write-Log -Status "Info" -Message "Finish to extract unique officeLocation"
        $screenReportByOfficeLocation = Group-TeamsReportBy -GroupByAttribute "officeLocation" -TeamsUsersList $parsedUserList -UniqueGroupByAttributeList $officeLocation -TimeSpan $TimeSpan
        return $screenReportByOfficeLocation
    }

    Write-Log -Status "Info" -Message "Start to group objects"

    if(!$TeamsReportGroupByAttributes){
        $screenReportByDepartment = New-TeamsDepartmentScorecard
        $screenReportByDomain = Group-TeamsReportBy -GroupByAttribute "UserPrincipalName" -TeamsUsersList $parsedUserList -UniqueGroupByAttributeList $domains -TimeSpan $TimeSpan
        $screenReportByDepartment | Export-Csv $summaryReportPathByDepartment -NoTypeInformation -Encoding UTF8
        $screenReportByDomain | Export-Csv $summaryReportPathByDomainName -NoTypeInformation -Encoding UTF8
        $screenReportByOfficeLocation | Export-Csv $summaryReportPathByofficeLocation -NoTypeInformation -Encoding UTF8
    }
    else{
        foreach($attribute in $TeamsReportGroupByAttributes){            
            switch ($attribute) {
                "Department" {
                    $screenReportByDepartment = New-TeamsDepartmentScorecard
                    $screenReportByDepartment | Export-Csv $summaryReportPathByDepartment -NoTypeInformation -Encoding UTF8
                }
                "Domain" {
                    $screenReportByDomain = New-TeamsDomainScoreCard
                    $screenReportByDomain | Export-Csv $summaryReportPathByDomainName -NoTypeInformation -Encoding UTF8
                }
                "officeLocation"{
                    $screenReportByOfficeLocation = New-TeamsOfficeLocationScoreCard
                    $screenReportByOfficeLocation | Export-Csv $summaryReportPathByofficeLocation -NoTypeInformation -Encoding UTF8
                }
                Default {}
            }
        }
    }

    <#

    #Group by department users who has teams license
    $usersPerDepartmentWithTeams = $parsedUserList | Where-Object{$_.HasTeamsLicense -eq "TRUE"} | Group-Object Department
    #Group by department users who has no teams license
    $usersPerDepartmentWithoutTeams = $parsedUserList | Where-Object{$_.HasTeamsLicense -ne "TRUE"} | Group-Object Department
    #Group by department users who has activity last date
    $usersPerDepartmentWithActivity = $parsedUserList | Where-Object{$null -ne $_.TeamsLastActivityDate -and $_.TeamsLastActivityDate -ne ""} | Group-Object Department
    #Group by department users who has no activity last date
    $usersPerDepartmentWithoutActivity = $parsedUserList | Where-Object{$null -eq $_.TeamsLastActivityDate-or $_.TeamsLastActivityDate -eq ""} | Group-Object Department
    #Group by department users who has teams meeting count greater than 0
    $usersPerDepartmentWithMeeting = $parsedUserList | Where-Object{$_.MeetingCount -gt 0} | Group-Object Department

    Write-Log -Status "Info" -Message "Finished to group objects"
    Write-Log -Status "Info" -Message "Start to build teams usage score"

    #Creates an array to append ps custom objects
    $screenReportByDepartment = @()

    #For each unique department found in users end point api, uses the grouped objects above to build up a teams usage score
    foreach($department in $departments){
        #Due to comparisons need, if department is blank set it as $null
        if(!$department){
            $department = $null
        }

        #Creates a ps custom object for the current department and count it down each scenario to build the department score
        $obj = [PSCustomObject]@{
            Department = $department
            UserCount = ($users|Where-Object{$_.department -eq $department}|Measure-Object).Count
            HasTeamsLicense = ($usersPerDepartmentWithTeams.group|Where-Object{$_.department -eq $department}|Measure-Object).Count
            HasNoTeamsLicense = ($usersPerDepartmentWithoutTeams.group|Where-Object{$_.department -eq $department}|Measure-Object).Count
            HasTeamsActivity = ($usersPerDepartmentWithActivity.group|Where-Object{$_.department -eq $department}|Measure-Object).Count
            HasNoTeamsActivity = ($usersPerDepartmentWithoutActivity.group|Where-Object{$_.department -eq $department}|Measure-Object).Count
            HasMeeting = ($usersPerDepartmentWithMeeting.group|Where-Object{$_.department -eq $department}|Measure-Object).Count

        }
        #Append the current ps custom object into the array
        $screenReportByDepartment += $obj
    }

    Write-Log -Status "Info" -Message "Finished the build of teams usage score"
    Write-Log -Status "Info" -Message "Report finished using Report Mode: $($ReportMode)"

    #Uses the ReportMode parameter input to define the output action
    switch($ReportMode){
        #Exports both summary and detailed scorecard
        "Export"{
            $screenReportByDepartment | Export-Csv $summaryReportPathByDepartment -NoTypeInformation -Encoding UTF8
            $screenReportByDomain | Export-Csv $summaryReportPathByDomainName -NoTypeInformation -Encoding UTF8
            $screenReportByOfficeLocation | Export-Csv $summaryReportPathByofficeLocation -NoTypeInformation -Encoding UTF8
        }
        #Exports both summary and per user detail scorecard
        "AsJob"{
            $screenReportByDepartment | Export-Csv $summaryReportPathByDepartment -NoTypeInformation -Encoding UTF8
            $screenReportByDomain | Export-Csv $summaryReportPathByDomainName -NoTypeInformation -Encoding UTF8
            $screenReportByOfficeLocation | Export-Csv $summaryReportPathByofficeLocation -NoTypeInformation -Encoding UTF8
        }
    }
    #>


    #Stop the watch and register the time spent to export teams usage report
    $stopWatchStop = Get-Date
    $stopWatchResult = New-TimeSpan -Start $stopWatchStart -End $stopWatchStop
    #Write-Log -Status "Info" -Message "Report finished using Report Mode: $($ReportMode)"
    Write-Log -Status "Info" -Message "Teams Usage Report execution time: $($stopWatchResult.ToString("dd\.hh\:mm\:ss"))"
}

#Exposes the following functions as module cmdlets
Export-ModuleMember -Function Get-M365LicenseSkuReport
Export-ModuleMember -Function Get-TeamsUsageReport
Export-ModuleMember -Function New-M365UsageCollectorAppRegistration
Export-ModuleMember -Function New-M365UsageCollectorJob