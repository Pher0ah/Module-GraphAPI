<#
    .Synopsis
    Helps leverage the Microsoft Graph API using PowerShell

    .Description
    Facilitates making Graph API requests using PoweShell

    .Parameter Start
    The first month to display.

    .Parameter End
    The last month to display.

    .Parameter FirstDayOfWeek
    The day of the month on which the week begins.

    .Parameter HighlightDay
    Specific days (numbered) to highlight. Used for date ranges like (25..31).
    Date ranges are specified by the Windows PowerShell range syntax. These dates are
    enclosed in square brackets.

    .Parameter HighlightDate
    Specific days (named) to highlight. These dates are surrounded by asterisks.

    .Example
    # Show a default display of this month.
    Show-Calendar

    .Example
    # Display a date range.
    Show-Calendar -Start "March, 2010" -End "May, 2010"

    .Example
    # Highlight a range of days.
    Show-Calendar -HighlightDay (1..10 + 22) -HighlightDate "December 25, 2008"
#>
#Requires -version 5.0

#region Global_Construct
#
$script:SessionToken = $null
$script:SessionExpiry = $null
$script:gSessionTenant = $null
$script:gApplicationID = $null
$script:gAppSecret = $null
$script:gAPIVersion = 'v1.0'

#endregion Global_Construct

#Testing TO BE DELETED
$script:gSessionTenant = 'generatione.onmicrosoft.com' #"70def4a1-688c-4890-9d1c-3f854ac0ba68"
$script:gApplicationID = 'bf09db13-cbf5-48c1-b720-7c5de5ca7373'
$script:gAppSecret     = 'pVjCgH]@19]f7E:fyfOuMpkAY525m[/K'



<############################################# Public Functions #############################################>



#region Set-APIVersion
#
function Set-APIVersion(){
  param(
    [Parameter(Mandatory=$true)][string] $Version
  )
  
  If(($Version -eq '1.0') -or ($Version -eq 'beta')){
    #Set the API Version to be used in this session
    If($Version -ne 'beta'){$Version = ('v{0}' -f $Version)}
    $script:gAPIVersion = $Version
  }else{
    Write-Host 'ERROR: Version is incorrect please use 1.0 or beta' -ForegroundColor Red
  }

}#endregion Set-APIVersion



#region Get-APIVersion
#
function Get-APIVersion(){
  
  #Return the API version to be used in this session
  return $script:gAPIVersion

}#endregion Get-APIVersion



#region Set-TenantID
#
function Set-TenantID(){
  param(
    [Parameter(Mandatory=$true)][string] $aTenantID
  )
  
  #Set the Tenant Name/ID to be used in this session
  $script:gSessionTenant = $aTenantID

}#endregion Set-TenantID



#region Get-TenantID
#
function Get-TenantID(){
  
  #Return the Tenant Name/ID used in this session
  return $script:gSessionTenant

}#endregion Get-TenantID



#region Set-ApplicationID
#
function Set-ApplicationID(){
  param(
    [Parameter(Mandatory=$true)][string] $aApplicationID
  )
  
  #Set the Azure Application ID to be used in this session
  $script:gApplicationID = $aApplicationID

}#endregion Set-ApplicationID



#region Get-ApplicationID
#
function Get-ApplicationID(){
  
  #Return the Azure Application ID used in this session
  return $script:gApplicationID

}#endregion Get-ApplicationID



#region Set-ApplicationSecret
#
function Set-ApplicationSecret(){
  param(
    [Parameter(Mandatory=$true)][string] $aAppSecret
  )
  
  #Set the Application Secret to be used in this session
  $script:gAppSecret = $aAppSecret

}#endregion Set-ApplicationSecret



#region Get-ApplicationSecret
#
function Get-ApplicationSecret(){
  
  #Return the Application Secret used in this session
  return $script:gAppSecret

}#endregion Get-ApplicationSecret




<############################################# Internal Functions #############################################>



#region Request-Token
# TODO: Add verification for the token life
# 
function Request-Token(){

  If($script:SessionToken -eq $null){

    If($script:gSessionTenant -eq $null){ 
      Write-Host 'ERROR: A TenantID must be set using Set-TenantID before getting a Session OAUTH Token' -ForegroundColor Red
      return $false
    }
  
    If($script:gApplicationID -eq $null){
      Write-Host 'ERROR: An Application ID must be set using Set-ApplicationID before getting a Session OAUTH Token' -ForegroundColor Red
      return $false
    }
    
    If($script:gAppSecret -eq $null){
      write-Host 'ERROR: An Application Secret must be set using Set-ApplicationSecret before getting a Session OAUTH Token' -ForegroundColor Red
      return $false
    }
    
    #Construct URI to login to Microsoft API
    $theURI = ('https://login.microsoftonline.com/{0}/oauth2/v2.0/token' -f $script:gSessionTenant)
    
    #Construct Body
    $theBody = @{
      client_id     = $script:gApplicationID
      scope         = 'https://graph.microsoft.com/.default'
      client_secret = $script:gAppSecret
      grant_type    = 'client_credentials'
    }
    
    #Get OAuth 2.0 Token
    $theRequest = Invoke-WebRequest -Method Post -Uri $theURI -ContentType 'application/x-www-form-urlencoded' -Body $theBody -UseBasicParsing
    
    #Return Access Token
    If($theRequest.statusCode -eq '200'){
      #Get Session Token
      $script:SessionToken = ($theRequest.Content | ConvertFrom-Json).access_token
      
      #Get TOken Epxiry Time
      $aTokenData = $script:SessionToken.Split('.')[1].Replace('-', '+').Replace('_', '/') + "=="
      $aUnixTime = ([System.Text.Encoding]::UTF8.GetString([convert]::FromBase64String($aTokenData))|Convertfrom-Json).exp
      $Script:SessionExpiry = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($aUnixTime))
      
      return $true
    }else{
      Write-Host ("ERROR: Couldn't retrieve token, StatusCode {0} returned" -f $theRequest.statusCode) -ForegroundColor Red
      return $false
      end
    }
  }else{
    If($Script:SessionExpiry -gt (Get-Date)){
      return $true
    }else{
      #Session Expired
      $Script:SessionToken = $null
      $Script:SessionExpiry = $null
      return (Request-Token)
    }
  }
}#endregion Request-Token



#region Get-APIMethod
# Run GET Method
#
function Get-APIMethod() {
  Param(
    [Parameter(Mandatory=$true ,ValueFromPipeline=$true)][String]$aURI,
    [Parameter(Mandatory=$false,ValueFromPipeline=$false)][AllowNull()][String]$aFolder
  )

  #Initialize Variables
  $theResults = [Collections.ArrayList]::new()
  $theJSON = $null

  #Construct the URI to call
  $theURI = ('https://graph.microsoft.com/v1.0/{0}' -f $aURI)
  
  #Check if the token is valid
  if(Request-Token){ 
    do{ #a loop to get all results
    
      #Initialize variables
      $theReturnValue = $null
      $theJSON = $null
      
      try{ #Running the Graph API query
        $theReturnValue = Invoke-WebRequest -Method Get -Uri $theURI -ContentType 'application/json' -Headers @{Authorization = "Bearer $script:SessionToken"} -ErrorAction Stop
      } catch {
        Write-Host ('ERROR: GET-APIMethod failed with a StatusCode {0} and message with message {1}' -f $Error[0].Exception.Response.statusCode, $Error[0].Exception.Message) -ForegroundColor Red
        return $null
        end
      }
    
      $theJSON = ($theReturnValue.content |ConvertFrom-Json)
      $theResults += $theJSON.value
      If($theURI -match 'top'){$theURI = $null}else{$theUri = $theJSON."@odata.nextlink"}
      
    } until (-not $theURI)
  }

  #Write Results to Folder
  If($aFolder){
    # Write Results
    If(Test-Path -Path $aFolder){
      foreach ($aResult in $theResults.value){
        $aResult | ConvertTo-Json -depth 100 | Out-File -FilePath ('{0}\{1}.json' -f ($aFolder), $aResult.DisplayName)
      }
    }else{
      Write-Host -ForegroundColor Red ('ERROR: Path not found {0}' -f ($aFolder))
    }
  }

  return $theResults
}
#endregion Get-APIMethod



<############################################# Graph Functions #############################################>



#region Get-GraphUsers
#
function Get-GraphUsers(){
  Param(
    [Parameter(Mandatory=$true,  ParameterSetName= 'Count', ValueFromPipeline=$false)][switch]$Count,
    [Parameter(Mandatory=$false, ParameterSetName= 'Get'  , ValueFromPipeline=$false)][String]$Limit = '100',
    [Parameter(Mandatory=$false, ParameterSetName= 'Get'  , ValueFromPipeline=$false)][AllowNull()][String]$Folder
  )

  #Get Users information
  If($Limit){
    $theURI = ("users?`$top={0}" -f $Limit)
  }else{
    $theURI = ("users")
  }

  If($Folder){
    $Output = Get-APIMethod -aURI $theURI -aFolder $Folder
  }else{
    $Output = Get-APIMethod -aURI $theURI
  }
  
  If($Count){
    return $Output.count
  }else{
    return $Output
  }
  
}#endregion Get-GraphUsers



#region Get-GraphDeviceConfigs
#
function Get-GraphDeviceConfigs(){
  Param(
    [Parameter(Mandatory=$false, ValueFromPipeline=$false)][String]$Limit = '100',
    [Parameter(Mandatory=$false, ValueFromPipeline=$false)][AllowNull()][String]$Folder
  )

  #Get Device Configurations
  $theURI = ("deviceManagement/deviceConfigurations?`$top={0}" -f $Limit)
  If($Folder){Get-APIMethod $theURI, $Folder}else{Get-APIMethod $theURI}

}#endregion Get-GraphDeviceConfigs



#region Get-GraphAIPPolicies
#
function Get-GraphAIPPolicies(){
  Param(
    [Parameter(Mandatory=$false, ValueFromPipeline=$false)][String]$Limit = '100',
    [Parameter(Mandatory=$false, ValueFromPipeline=$false)][AllowNull()][String]$Folder
  )

  #Get AIP POlicies
  $theURI = ("deviceAppManagement/windowsInformationProtectionPolicies?`$top={0}" -f $Limit)
  If($Folder){Get-APIMethod $theURI, $Folder}else{Get-APIMethod $theURI}

}#endregion Get-GraphAIPPolicies



<############################################# Export Functions #############################################>
#Sets & Gets
Export-ModuleMember -Function Set-APIVersion
Export-ModuleMember -Function Get-APIVersion
Export-ModuleMember -Function Set-TenantID
Export-ModuleMember -Function Get-TenantID
Export-ModuleMember -Function Set-ApplicationID
Export-ModuleMember -Function Get-ApplicationID
Export-ModuleMember -Function Set-ApplicationSecret
Export-ModuleMember -Function Get-ApplicationSecret

#Functions for Graph
Export-ModuleMember -Function Get-GraphUsers
Export-ModuleMember -Function Get-GraphDeviceConfigs
Export-ModuleMember -Function Get-GraphAIPPolicies

