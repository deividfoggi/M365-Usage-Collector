<#################################################################################
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
#################################################################################

 .Synopsis
  Collects information related to Teams usage from a M365 tenant.

 .Description
  Display or export to .csv files information related to Teams usage in the contexts of user activity, user activation and licensing.
  This module allows you to either export to .csv files or have the information in the current PowerShell session to customize the output at your will.

 .Example
   # Get license usage report per sku
   Get-LicenseSkuReport

   # Get Teams usage report including summary and a full report per user
   Get-TeamsUsageReport -ClientId 00000000-0000-0000-0000-000000000000 -TenantId 00000000-0000-0000-0000-000000000000 -ClientSecret 00000000-0000-0000-0000-000000000000
  
   # Get Teams usage summary only
   Get-TeamsUsageReport -ReportMode "SummaryOnly" -ClientId 00000000-0000-0000-0000-000000000000 -TenantId 00000000-0000-0000-0000-000000000000 -ClientSecret 00000000-0000-0000-0000-000000000000

   # Get Teams usage per user
   Get-TeamsUsageReport -ReportMode "PerUser" -ClientId 00000000-0000-0000-0000-000000000000 -TenantId 00000000-0000-0000-0000-000000000000 -ClientSecret 00000000-0000-0000-0000-000000000000
   
   # Export summary and per user reports to csv files
   Get-TeamsUsageReport -ReportMode "Export" -ClientId 00000000-0000-0000-0000-000000000000 -TenantId 00000000-0000-0000-0000-000000000000 -ClientSecret 00000000-0000-0000-0000-000000000000
#>

#Check if AzureAD or AzureADPreview is installed
If(!(Get-InstalledModule -Name AzureAD -ErrorAction SilentlyContinue)){
    #Installs default version of AzureAD
    try{
        Install-Module AzureAD -ErrorAction SilentlyContinue
    }
    catch{
        Write-Warning $_.Exception
    }
}
else{
    #If already installed try to import the module
    try{
        Import-Module AzureAD -ErrorAction Stop
    }
    catch{
        Write-Warning $_.Exception
    }
}

#Function to connect to Azure AD
Function ConnectAzureAD{
    #Try to get tenant details as a way to check if AzureAD module is already connected. If not connect throws an exception to be catched and connect to Azure AD
    try{
        Get-AzureADTenantDetail -ErrorAction Stop | Out-Null
        Write-Host "Azure AD already connected" -ForegroundColor Yellow
    }
    catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException]
    {
        #If a 'need auth' exception then connect to Azure AD asking for credentials
        Write-Warning "Connecting to Azure AD"
        Connect-AzureAD
    }
}

#Function to create an application registration in Azure AD to be used to connect to Graph API
Function New-M365UsageCollectorAppRegistration {
        #Connects to Azure AD
        ConnectAzureAD

        #Define permissions needed
        $appPerms = 'Reports.Read.All','User.Read.All'
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
        $appCreated = New-AzureADApplication @appCreationParameters

        #Define start and end date for client secret duration
        $startDate = Get-Date
        $endDate = $startDate.AddYears(3)

        #Creates a new client secret
        $appReg = New-AzureADApplicationPasswordCredential -ObjectId $appCreated.ObjectId -CustomKeyIdentifier "clientSecret" -StartDate ([DateTime]::Now) -EndDate $endDate

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

    #If there is no token in current session
    If(!$global:accessToken){

        #Warns user that a new access token will be requested
        Write-Warning "No token in cache. Acquiring access token from Azure AD."

        $stringUrl = "https://login.microsoftonline.com/" + $tenantId + "/oauth2/v2.0/token"
        $postData = "client_id=" + $AppId + "&scope=https://graph.microsoft.com/.default&client_secret=" + $ClientSecret + "&grant_type=client_credentials"
        try{
            $accessToken = Invoke-RestMethod -Method post -Uri $stringUrl -ContentType "application/x-www-form-urlencoded" -Body $postData  -ErrorAction Stop
            Write-Warning "Access token acquired."
            return $accessToken
        }
        catch{
            $errorDescription = $_ | ConvertFrom-Json
            Write-Warning $errorDescription.error
            Write-Host $errorDescription.error_description -ForegroundColor Yellow
        }
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
Function Get-LicenseSkuReport {
    Param(
        [Parameter(Mandatory=$false)]$Export
    )

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
        $licenseReport | Export-Csv .\LicenseReport.csv -NoTypeInformation
    }
    #If no Export parameter then print out in the screen
    else{
        return $licenseReport|Format-Table
    }
}

#Function to get teams usage report from Graph Reports API
Function Get-TeamsUsageReport{
    Param(
        [Parameter(Mandatory=$true)]$AppId,
        [Parameter(Mandatory=$true)]$TenantId,
        [Parameter(Mandatory=$true)]$ClientSecret,
        [Parameter(Mandatory=$true)]$ReportMode
    )

    #Register in a variable the start datetime for statistics purposes
    $stopWatchStart = Get-Date
    
    #Uses EscapeDataString function to prevent an issue that replaces all + sign in the client secret string with a blank space
    $ClientSecret = [System.Uri]::EscapeDataString($ClientSecret)

    #Get an Azure AD token using app reg info
    $accessToken = (Get-AzureADToken -AppId $AppId -TenantId $TenantId -ClientSecret $ClientSecret).access_token
    
    #Send graph api requests against reports API to get teams reports considering a 30 days time span
    $teamsUserActivityUserDetail = (Send-GraphRequest -Method Get -BearerToken $accessToken -Path "/reports/getTeamsUserActivityUserDetail(period='D30')")|ConvertFrom-Csv
    $office365ActiveUserDetail = (Send-GraphRequest -Method Get -BearerToken $accessToken -Path "/reports/getOffice365ActiveUserDetail(period='D30')")|ConvertFrom-Csv

    #Send graph api request against users api to get UPN and Department in order to parse department agains users in the reports collected above
    $users = Send-GraphRequest -Method Get -BearerToken $accessToken -Path "/users?`$select=UserPrincipalName,Department&`$top=999"

    #Create an empty array to append ps custom objects
    $joinedObjects = @()

    #Incrementation control variable for progress bar
    $i = 1

    #for each user found in graph api users endpoint
    foreach($user in $users){

        #Write the progress bar
        Write-Progress -Activity "Parsing users in report" -Status "Parsing user $i of $($users.length)" -Id 1 -PercentComplete (($i / $users.length)*100)

        #Extract from the teams activity user detail report the current user findings
        $userteamsUserActivityUserDetail = $teamsUserActivityUserDetail | Where-Object{$_.'User Principal Name' -eq $user.UserPrincipalName}

        #Extract from the teams active user detail report the current user findings
        $office365ActiveUserDetailUser = $office365ActiveUserDetail | Where-Object{$_.'User Principal Name' -eq $user.UserPrincipalName}

        #Create a ps custom object to store current user findings
        $userObj = [PSCustomObject] @{
            #Sanitize UserPrincipalName and DisplayName to remove PII
            UserPrincipalName = "Sanitized"
            DisplayName = "Sanitized"

            #Fill the following attributes accordingly
            Department = $user.Department
            IsDeleted = $office365ActiveUserDetailUser.'Is Deleted'
            DeletedDate = $office365ActiveUserDetailUser.'Deleted Date'
            HasExchangeLicense = $office365ActiveUserDetailUser.'Has Exchange License'
            HasOneDriveLicense = $office365ActiveUserDetailUser.'Has OneDrive License'
            HasSharePointLicense = $office365ActiveUserDetailUser.'Has SharePoint License'
            HasSkypeForBusinessLicense = $office365ActiveUserDetailUser.'Has Skype For Business License'
            HasYammerLicense = $office365ActiveUserDetailUser.'Has Yammer License'
            HasTeamsLicense = $office365ActiveUserDetailUser.'Has Teams License'
            ExchangeLastActivityDate = $office365ActiveUserDetailUser.'Exchange Last Activity Date'
            OneDriveLastActivityDate = $office365ActiveUserDetailUser.'OneDrive Last Activity Date'
            SharePointLastActivityDate = $office365ActiveUserDetailUser.'SharePoint Last Activity Date'
            SkypeForBusinessLastActivityDate = $office365ActiveUserDetailUser.'Skype For Business Last Activity Date'
            YammerLastActivityDate = $office365ActiveUserDetailUser.'Yammer Last Activity Date'
            TeamsLastActivityDate = $office365ActiveUserDetailUser.'Teams Last Activity Date'
            ExchangeLicenseAssignDate = $office365ActiveUserDetailUser.'Exchange License Assign Date'
            OneDriveLicenseAssignDate = $office365ActiveUserDetailUser.'OneDrive License Assign Date'
            SharePointLicenseAssignDate = $office365ActiveUserDetailUser.'SharePoint License Assign Date'
            SkypeForBusinessLicenseAssignDate = $office365ActiveUserDetailUser.'Skype For Business License Assign Date'
            YammerLicenseAssignDate = $office365ActiveUserDetailUser.'Yammer License Assign Date'
            TeamsLicenseAssignDate = $office365ActiveUserDetailUser.'Teams License Assign Date'
            AssignedProducts = $office365ActiveUserDetailUser.'Assigned Products'
            LastActivityDate = $userteamsUserActivityUserDetail.'Last Activity Date'
            TeamChatMessageCount = $userteamsUserActivityUserDetail.'Team Chat Message Count'
            PrivateChatMessageCount = $userteamsUserActivityUserDetail.'Private Chat Message Count'
            CallCount = $userteamsUserActivityUserDetail.'Call Count'
            MeetingCount = $userteamsUserActivityUserDetail.'Meeting Count'
            HasOtherAction = $userteamsUserActivityUserDetail.'Has Other Action'
            ReportPeriod = $userteamsUserActivityUserDetail.'Report Period'
        }
        #Append current object into the array
        $joinedObjects += $userObj

        #Increment the progress bar control variable
        $i++
    }

    #Extract unique department strings from users endpoint result
    $departments = ($users | Select-Object Department -Unique).department
    #Group by department users who has teams license
    $usersPerDepartmentWithTeams = $joinedObjects | Where-Object{$_.HasTeamsLicense -eq "TRUE"} | Group-Object Department
    #Group by department users who has no teams license
    $usersPerDepartmentWithoutTeams = $joinedObjects | Where-Object{$_.HasTeamsLicense -ne "TRUE"} | Group-Object Department
    #Group by department users who has activity last date
    $usersPerDepartmentWithActivity = $joinedObjects | Where-Object{$null -ne $_.TeamsLastActivityDate -and $_.TeamsLastActivityDate -ne ""} | Group-Object Department
    #Group by department users who has no activity last date
    $usersPerDepartmentWithoutActivity = $joinedObjects | Where-Object{$null -ne $_.TeamsLastActivityDate -or $_.TeamsLastActivityDate -eq ""} | Group-Object Department
    #Group by department users who has teams meeting count greater than 0
    $usersPerDepartmentWithMeeting = $joinedObjects | Where-Object{$_.MeetingCount -gt 0} | Group-Object Department

    #Creates an array to append ps custom objects
    $screenReport = @()

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
        $screenReport += $obj
    }

    #Uses the ReportMode parameter input to define the output action
    switch($ReportMode){
        #Prints out a scorecard summary
        "SummaryOnly" {
            $TeamsUsageSummary = $screenReport
            return $TeamsUsageSummary
        }
        #Prints out utilization detail per user
        "PerUser" {
            $TeamsUsagePerUser = $joinedObjects
            return $TeamsUsagePerUser
        }
        #Exports both summary and per user detail scorecard
        "Export"{
            $summaryReportName = "TeamsUsageData_Summary.csv"
            $perUserReportName = "TeamsUsageData_PerUser.csv"
            $screenReport | Export-Csv .\TeamsUsageData_Summary.csv -NoTypeInformation
            $joinedObjects | Export-Csv .\TeamsUsageData_PerUser.csv -NoTypeInformation
            Write-Host "Report saved in the following files:
                Summarized report - $((Get-Item $summaryReportName).FullName)
                Per user report - $((Get-Item $perUserReportName).FullName)."
                #Exports to a csv the subscribed license sku report
            Get-LicenseSkuReport -Export $true
        }
    }

    #Stop the watch and register the time spent to export teams usage report
    $stopWatchStop = Get-Date
    $stopWatchResult = New-TimeSpan -Start $stopWatchStart -End $stopWatchStop
    Write-Warning "Execution time: $($stopWatchResult.ToString("dd\.hh\:mm\:ss"))"
}

#Exposes the following functions as module cmdlets
Export-ModuleMember -Function Get-LicenseSkuReport
Export-ModuleMember -Function Get-TeamsUsageReport
Export-ModuleMember -Function New-M365UsageCollectorAppRegistration