
<#
<-Update-O365Tokens->
.SYNOPSIS
Updates expired access token with the stored refresh token

.DESCRIPTION
access token is valid for 60 minutes and the refresh token is valid for 90 days.

.PARAMETER ClientSecret 
API recource generally this is "https://graph.microsoft.com/"  but this should be managed in the settings file

.PARAMETER AccessTokenName 
The access token name that should be updated

.PARAMETER RefreshTokenName
The access token name that should be used and updated

.PARAMETER ClientId
The Client id which can be found in the registered application on the MSGraph page

.PARAMETER RedirectUri
The redirect URI that is registered in the created application

.PARAMETER TokenDirectory 
Directory the tokens should be saved to
#>
Function Update-O365Tokens {
    [CmdLetBinding()]
    Param(
        [Parameter()][String]$ClientSecret = 'Your MS Graph app secret',
        [Parameter()][String]$AccessTokenName = 'Your Access Token file name',
        [Parameter()][String]$RefreshTokenName = 'Your refresh token file name',
        [Parameter()][String]$ClientId = 'Your MS graph clientid',
        [Parameter()][String]$RedirectUri = 'your MS Graph redirect URI',
        [Parameter()][String]$TokenDirectory = 'C:\tokens\'
    )
    #############################################################################
    # Using refresh token only works if less than 90 days old
    #############################################################################
    $accesstokenpath = Join-Path -Path $TokenDirectory -ChildPath $AccessTokenName
    $refreshtokenpath = Join-Path -Path $TokenDirectory -ChildPath $RefreshTokenName

    $refreshtoken = Get-Content "$($refreshtokenpath)"
    $ClientSecretEncoded = [System.Web.HttpUtility]::UrlEncode($ClientSecret)
    $Body = "grant_type=refresh_token&refresh_token=$refreshtoken&redirect_uri=$redirectUri&client_id=$clientId&client_secret=$clientSecretEncoded"
    
    $Params = @{
        "URI"         = "https://login.microsoftonline.com/common/oauth2/token";
        "Method"      = "Post";
        "ContentType" = "application/x-www-form-urlencoded";
        "Body"        = $Body;
        "ErrorAction" = "Stop";
    }
    $Authorization = Invoke-RestMethod @Params

    $accesstoken = $Authorization.access_token
    $refreshtoken = $Authorization.refresh_token

    if ($accesstoken) {
        $accesstoken | Out-File "$($accesstokenpath)"
        Write-Verbose "Retrieved new accesstoken"
    }
    else {
        Write-Error 'Failed to retrieve Access Token'
    }
 
    if ($refreshtoken) {
        $refreshtoken | Out-File "$($refreshtokenpath)"
        Write-Verbose "Retrieved new accesstoken"
    }
    else {
        Write-Error 'Failed to retrieve Refresh Token'
    }
}

