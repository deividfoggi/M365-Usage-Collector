#################################################################################
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
  Export to .csv files information related to Teams usage in the contexts of user activity, user activation and licensing.
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

   # To export data to .csv, use the parameter -Export $true in any of the above options
#>

#Requisites check
If(!(Get-InstalledModule -Name MSAL.PS -ErrorAction SilentlyContinue)){
    Install-Module MSAL.PS
}else{
    Import-Module MSAL.PS
}
If(!(Get-InstalledModule -Name AzureAD -ErrorAction SilentlyContinue)){
    Install-Module AzureAD
}else{
    Import-Module AzureAD
}

Function Send-GraphRequest{
    Param(
    [Parameter(Mandatory=$true)]$Method,
    [Parameter(Mandatory=$false)]$BearerToken,
    [Parameter(Mandatory=$false)]$Path,
    [Parameter(Mandatory=$false)]$Json,
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
            $request = Invoke-RestMethod -Method $Method -Headers @{Authorization = "Bearer $($bearerToken)"} -Uri $Uri -ContentType 'application/json' -Body $json -ErrorAction Stop
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
        Write-Warning -Message $_.Exception.Message
    }
}

Function Get-LicenseSkuReport {

    Param(
        [Parameter(Mandatory=$true)]$Export
    )

    if(!(Get-AzureADTenantDetail)){
        Connect-AzureAD
    }

    $licensesRequest = Get-AzureADSubscribedSku | Select-Object SkuPartNumber,*Units

    $licenseReport = @()

    foreach($sku in $licensesRequest){

        $objLicense = [PSCustomObject] @{
            "ConsumedUnits" = $sku.ConsumedUnits
            "EnabledUnits" = $sku.PrepaidUnits.Enabled
            "Sku" = $sku.SkuPartNumber
        }
        $licenseReport += $objLicense
    }
    if($Export -eq $true){
        $licenseReport | Export-Csv .\LicenseReport.csv -NoTypeInformation
    }else{
        return $licenseReport
    }
}


Function Get-TeamsUsageReport{

    Param(
        [Parameter(Mandatory=$true)]$ClientID,
        [Parameter(Mandatory=$true)]$TenantId,
        [Parameter(Mandatory=$true)]$ClientSecret,
        [Parameter(Mandatory=$true)]$ReportMode,
        [Parameter(Mandatory=$false)]$Export
    )

    #Following 3 lines are for test only and should be removed on final version
    $clientId = "36533c7a-40cd-4f71-8362-c121dbc19b8a"
    $clientSecret = (ConvertTo-SecureString "T__.n-jXkom_SM3uP2t2enBB~~dkBVFRd5" -AsPlainText -Force)
    $tenantId = "cdcae3ff-a663-4732-9cf5-1e33db81acf1"
    
    $teamsUserActivityUserDetail = (Send-GraphRequest -Method Get -BearerToken (Get-MsalToken -ClientId $ClientID -ClientSecret $ClientSecret -TenantId $TenantId).AccessToken -Path "/reports/getTeamsUserActivityUserDetail(period='D180')")|ConvertFrom-Csv
    $office365ActiveUserDetail = (Send-GraphRequest -Method Get -BearerToken (Get-MsalToken -ClientId $ClientID -ClientSecret $ClientSecret -TenantId $TenantId).AccessToken -Path "/reports/getOffice365ActiveUserDetail(period='D180')")|ConvertFrom-Csv
    $users = Send-GraphRequest -Method Get -BearerToken (Get-MsalToken -ClientId $ClientID -ClientSecret $ClientSecret -TenantId $TenantId).AccessToken -Path "/users?`$select=UserPrincipalName,Department&`$top=999"

    $joinedObjects = @()

    foreach($user in $users){

        $userteamsUserActivityUserDetail = $teamsUserActivityUserDetail | Where-Object{$_.'User Principal Name' -eq $user.UserPrincipalName}
        $office365ActiveUserDetailUser = $office365ActiveUserDetail | Where-Object{$_.'User Principal Name' -eq $user.UserPrincipalName}

        $userObj = [PSCustomObject] @{
            UserPrincipalName = $user.UserPrincipalName
            Department = $user.Department
            DisplayName = $office365ActiveUserDetailUser.'Display Name'
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
        switch($ReportMode){
            "SummaryOnly" {
                $screenReport | Export-Csv .\TeamsUsageData_Summary.csv -NoTypeInformation
            }
            "PerUser" {
                $joinedObjects | Export-Csv .\TeamsUsageData_PerUser.csv -NoTypeInformation
            }
        }
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
        }
    }
}

Export-ModuleMember -Function Get-LicenseSkuReport
Export-ModuleMember -Function Get-TeamsUsageReport