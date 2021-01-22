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

#Requisites check
If(!(Get-InstalledModule -Name AzureAD -ErrorAction SilentlyContinue)){
    Install-Module AzureAD
}else{
    Import-Module AzureAD
}

Function ConnectAzureAD{
    try{
        Get-AzureADTenantDetail -ErrorAction Stop | Out-Null
        Write-Host "Azure AD already connected" -ForegroundColor Yellow
    }
    catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException]
    {
        Write-Warning "Connecting to Azure AD"
        Connect-AzureAD
    }
}

Function Create-M365UsageCollectorAppRegistration {

        ConnectAzureAD

        #### Collect all the permissions first ####
        $appPerms = 'Reports.Read.All','User.Read.All'
        $replyUrls = "https://login.microsoftonline.com/common/oauth2/nativeclient"

        $msGraphService = Get-AzureADServicePrincipal -All $true -Filter "DisplayName eq 'Microsoft Graph'"
        $permissions = $msGraphService.AppRoles.Where({$_.Value -in $appPerms})
            
        $msGraphResourceAccess = New-Object -TypeName "Microsoft.Open.AzureAd.Model.RequiredResourceAccess"
        $msGraphResourceAccess.ResourceAppId = $msGraphService.AppId

        foreach($permission in $permissions){
            $appPermissions = new-object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList $permission.Id,"Role"
            $msGraphResourceAccess.ResourceAccess += $appPermissions
        }

        #### Create the app with all the permissions ####
        $appName = "Teams Usage Collector"
        $appCreationParameters = @{
            "DisplayName" = $appName;
            "RequiredResourceAccess" = @($msGraphResourceAccess)
            "ReplyUrls" = $replyUrls
        }

        $appCreated = New-AzureADApplication @appCreationParameters
        $startDate = Get-Date
        $endDate = $startDate.AddYears(3)
        $appReg = New-AzureADApplicationPasswordCredential -ObjectId $appCreated.ObjectId -CustomKeyIdentifier "clientSecret" -StartDate ([DateTime]::Now) -EndDate $endDate

        $objAppReg = [PSCustomObject]@{
            ObjectId = $appCreated.ObjectId
            AppId = $appCreated.AppId
            TenantId = (Get-AzureADTenantDetail).ObjectId
            ClientSecret = $appReg.Value
        }

        #$objAppReg| Export-Csv .\AppRegInfo -NoTypeInformation

        Write-Host "Application $appName created successfully in your tenant. Take not of the following information. If you lost one of them, ask you tenant admin to get it for you in Azure AD:
            AppId: $($objAppReg.AppId)
            TenantId: $($objAppReg.TenantId)
            ClientSecret: $($objAppReg.ClientSecret)
        
        " -Foreground Green
        Write-Host "Azure Admin should consent using the following link:" -Foreground Yellow
        Write-Host ("https://login.microsoftonline.com/$($objAppReg.TenantId)/adminconsent?client_id=$($objAppReg.AppId)&redirect_uri=$($replyUrls)") -Foreground Yellow
}

Function Get-AzureADToken{
    Param(
        [Parameter(Mandatory=$true)]$AppId,
        [Parameter(Mandatory=$true)]$TenantID,
        [Parameter(Mandatory=$true)]$ClientSecret
    )

    if(!$global:accessToken){

        Write-Warning "No token in cache. Acquiring access token from Azure AD."

        $stringUrl = "https://login.microsoftonline.com/" + $tenantId + "/oauth2/v2.0/token"
        $postData = "client_id=" + $AppId + "&scope=https://graph.microsoft.com/.default&client_secret=" + $ClientSecret + "&grant_type=client_credentials"
        try{
            $accessToken = Invoke-RestMethod -Method post -Uri $stringUrl -ContentType "application/x-www-form-urlencoded" -Body $postData -ErrorAction Stop
            return $accessToken
        }
        catch{
            $errorDescription = $_ | ConvertFrom-Json
            Write-Warning $errorDescription.error
            Write-Host $errorDescription.error_description -ForegroundColor Yellow
        }
    }
}

Function Send-GraphRequest{
    Param(
    [Parameter(Mandatory=$true)]$Method,
    [Parameter(Mandatory=$false)]$BearerToken,
    [Parameter(Mandatory=$false)]$Path,
    [Parameter(Mandatory=$false)]$Beta
    )

    If($Beta -eq $true){
        $Uri = "https://graph.microsoft.com/Beta" + $Path
    }
    else{
        $Uri = "https://graph.microsoft.com/v1.0" + $Path
    }

    try{
        $queryResults = @()
        do{
            $request = Invoke-RestMethod -Method $Method -Headers @{Authorization = "Bearer $($bearerToken)"} -Uri $Uri -ContentType 'application/json' -ErrorAction Stop
            if($request.value){
                $queryResults += $request.value
            }
            else{
                $queryResults += $request
            }
            $Uri = $request.'@odata.nextLink'
        } until (!($Uri))
        return $queryResults
    }
    catch{
        $errorDescription = $_ | ConvertFrom-Json
        Write-Warning $errorDescription.error
        Write-Host $errorDescription.error_description -ForegroundColor Yellow
    }
}
Function Get-LicenseSkuReport {

    Param(
        [Parameter(Mandatory=$false)]$Export
    )

    ConnectAzureAD

    $licensesRequest = Get-AzureADSubscribedSku | Select-Object SkuPartNumber,*Units

    $licenseReport = @()

    foreach($sku in $licensesRequest){

        $objLicense = [PSCustomObject] @{
            "Sku" = $sku.SkuPartNumber
            "EnabledUnits" = $sku.PrepaidUnits.Enabled
            "ConsumedUnits" = $sku.ConsumedUnits            
        }
        $licenseReport += $objLicense
    }
    if($Export -eq $true){
        $licenseReport | Export-Csv .\LicenseReport.csv -NoTypeInformation
    }else{
        return $licenseReport|Format-Table
    }
}
Function Get-TeamsUsageReport{

    Param(
        [Parameter(Mandatory=$true)]$AppId,
        [Parameter(Mandatory=$true)]$TenantId,
        [Parameter(Mandatory=$true)]$ClientSecret,
        [Parameter(Mandatory=$true)]$ReportMode
    )

    $accessToken = (Get-AzureADToken -AppId $AppId -TenantId $TenantId -ClientSecret $ClientSecret).access_token
    
    $teamsUserActivityUserDetail = (Send-GraphRequest -Method Get -BearerToken $accessToken -Path "/reports/getTeamsUserActivityUserDetail(period='D30')")|ConvertFrom-Csv
    $office365ActiveUserDetail = (Send-GraphRequest -Method Get -BearerToken $accessToken -Path "/reports/getOffice365ActiveUserDetail(period='D30')")|ConvertFrom-Csv
    $users = Send-GraphRequest -Method Get -BearerToken $accessToken -Path "/users?`$select=UserPrincipalName,Department&`$top=999"

    $joinedObjects = @()

    foreach($user in $users){

        $userteamsUserActivityUserDetail = $teamsUserActivityUserDetail | Where-Object{$_.'User Principal Name' -eq $user.UserPrincipalName}
        $office365ActiveUserDetailUser = $office365ActiveUserDetail | Where-Object{$_.'User Principal Name' -eq $user.UserPrincipalName}

        $userObj = [PSCustomObject] @{
            <#Sanitized attributes#>
            UserPrincipalName = "Sanitized" # $user.UserPrincipalName
            DisplayName = "Sanitized" # $office365ActiveUserDetailUser.'Display Name'
            <#End of Sanitized attributes#>
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
        $joinedObjects += $userObj
    }

    $departments = ($users | Select-Object Department -Unique).department
    $usersPerDepartmentWithTeams = $joinedObjects | Where-Object{$_.HasTeamsLicense -eq "TRUE"} | Group-Object Department
    $usersPerDepartmentWithoutTeams = $joinedObjects | Where-Object{$_.HasTeamsLicense -ne "TRUE"} | Group-Object Department
    $usersPerDepartmentWithActivity = $joinedObjects | Where-Object{$_.TeamsLastActivityDate -ne $null -and $_.TeamsLastActivityDate -ne ""} | Group-Object Department
    $usersPerDepartmentWithoutActivity = $joinedObjects | Where-Object{$_.TeamsLastActivityDate -eq $null -or $_.TeamsLastActivityDate -eq ""} | Group-Object Department
    $usersPerDepartmentWithMeeting = $joinedObjects | Where-Object{$_.MeetingCount -gt 0} | Group-Object Department

    $screenReport = @()

    foreach($department in $departments){
        if(!$department){
            $department = $null
        }

        $obj = [PSCustomObject]@{
            Department = $department
            UserCount = ($users|Where-Object{$_.department -eq $department}|Measure-Object).Count #checar por que as contagens nesse loop estão incorretas
            HasTeamsLicense = ($usersPerDepartmentWithTeams.group|Where-Object{$_.department -eq $department}|Measure-Object).Count
            HasNoTeamsLicense = ($usersPerDepartmentWithoutTeams.group|Where-Object{$_.department -eq $department}|Measure-Object).Count
            HasTeamsActivity = ($usersPerDepartmentWithActivity.group|Where-Object{$_.department -eq $department}|Measure-Object).Count
            HasNoTeamsActivity = ($usersPerDepartmentWithoutActivity.group|Where-Object{$_.department -eq $department}|Measure-Object).Count
            HasMeeting = ($usersPerDepartmentWithMeeting.group|Where-Object{$_.department -eq $department}|Measure-Object).Count

        }
        $screenReport += $obj
    }

    if($Export -eq $true){

    }
    else{
        switch($ReportMode){
            "SummaryOnly" {
                $TeamsUsageSummary = $screenReport
                return $TeamsUsageSummary
            }
            "PerUser" {
                $TeamsUsagePerUser = $joinedObjects
                return $TeamsUsagePerUser
            }
            "Export"{
                $summaryReportName = "TeamsUsageData_Summary.csv"
                $perUserReportName = "TeamsUsageData_PerUser.csv"
                $screenReport | Export-Csv .\TeamsUsageData_Summary.csv -NoTypeInformation
                $joinedObjects | Export-Csv .\TeamsUsageData_PerUser.csv -NoTypeInformation
                Write-Host "Report saved in the following files:
                    Summarized report - $((Get-Item $summaryReportName).FullName)
                    Per user report - $((Get-Item $perUserReportName).FullName)."
            }
        }
    }
    Get-LicenseSkuReport -Export $true
}

Export-ModuleMember -Function Get-LicenseSkuReport
Export-ModuleMember -Function Get-TeamsUsageReport
Export-ModuleMember -Function Create-M365UsageCollectorAppRegistration