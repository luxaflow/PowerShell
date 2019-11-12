<#
<-New-O365Tokens->
.SYNOPSIS
Function requests the initially required token from the MSGraph API

.DESCRIPTION
This function is normally run manually on application deployment. The access token is valid for 60 minutes and the refresh token is valid for 90 days.

.PARAMETER Resource 
API recource generally this is "https://graph.microsoft.com/"  but this should be managed in the settings file

.PARAMETER RedirectUri 
The redirect URI that is registered in the created application

.PARAMETER ClientId
The Client id which can be found in the registered application on the MSGraph page

.PARAMETER ClientSecret
The Secret token that is given only once at app creation land otherwise needs to be reset

.PARAMETER Scope
The access scope that is associated with the application access

.PARAMETER AccessTokenName 
naming convention for access tokens 

.PARAMETER RefreshTokenName
naming convention for refresh tokens 

.PARAMETER TokenSavePath
Directory the tokens should be saved to3
#>
Function New-O365Tokens {

    [CmdletBinding()]
    Param(
        [Parameter()][String]$Resource = 'https://graph.microsoft.com/',
        [Parameter()][String]$RedirectUri = 'Your MS graph redirect usi',
        [Parameter()][String]$ClientId = 'Your MS Graph api clientId',
        [Parameter()][String]$ClientSecret = 'Your MS Graph API Secret',
        [Parameter()][String]$Scope = 'User.Read',
        [Parameter()][String]$AccessTokenName = 'access.token',
        [Parameter()][String]$RefreshTokenName = 'refresh.token',
        [Parameter()][String]$TokenSavePath = 'C:\tokens\'
    )
    begin {
        $regex = '(?<=code=)(.*)(?=&)'
        $resourceEncoded = [System.Web.HttpUtility]::UrlEncode($resource)
        $clientSecretEncoded = [System.Web.HttpUtility]::UrlEncode($clientSecret)
        $url = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&redirect_uri=$redirectUri&client_id=$clientID&resource=$resourceEncoded"

        $accesstokenpath = Join-Path -Path $TokenSavePath -ChildPath $AccessTokenName
        $refreshtokenpath = Join-Path -Path $TokenSavePath -ChildPath $RefreshTokenName
        #############################################################################
        # Get AuthCode
        #############################################################################
        Add-Type -AssemblyName System.Windows.Forms

        $Form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width = 440; Height = 640 }
        $Web = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width = 420; Height = 600; Url = ($URL -f ($Scope -join "%20")) }

        $DocComp = {
            $Global:uri = $Web.Url.AbsoluteUri        
            if ($Global:uri -match "error=[^&]*|code=[^&]*") { $Form.Close() }
        }
        $Web.ScriptErrorsSuppressed = $true
        $Web.Add_DocumentCompleted($DocComp)
        $Form.Controls.Add($Web)
        $Form.Add_Shown( { $Form.Activate() })
        $Form.ShowDialog() | Out-Null

        $QueryOutput = [System.Web.HttpUtility]::ParseQueryString($Web.Url.Query)
    
        $Output = @{ }
        foreach ($Key in $QueryOutput.Keys) {
            $Output["$Key"] = $QueryOutput[$Key]
        }
        #############################################################################
        # Extract Access token from the returned URI
        #############################################################################
    
        $AuthCode = ($Uri | Select-string -pattern $regex).Matches[0].Value
        #############################################################################
        # Create Required Body for inital Token request
        #############################################################################
        $body = "grant_type=authorization_code&redirect_uri=$redirectUri&client_id=$clientId&client_secret=$clientSecretEncoded&code=$AuthCode&resource=$resource&prompt=login"
    }
    process {

        #############################################################################
        # Request Token and store it in file
        #############################################################################
        $Params = @{
            "URI"         = "https://login.microsoftonline.com/common/oauth2/token"
            "Method"      = "Post"
            "ContentType" = "application/x-www-form-urlencoded"
            "Body"        = $body
            "ErrorAction" = "Stop"
        }

        $Authorization = Invoke-RestMethod @Params

        $Global:accesstoken = $Authorization.access_token
        $Global:refreshtoken = $Authorization.refresh_token

        if ($accesstoken) {
            $accesstoken | Out-File "$($accesstokenpath)"
            Write-Verbose "Access token retreived"
        }
        else {
            Write-Error 'Failed to retrieve Access Token'
        }
 
        if ($refreshtoken) {
            $refreshtoken | Out-File "$($refreshtokenpath)"
            Write-Verbose "Refresh token retrieved"
        }
        else {
            Write-Error 'Failed to retrieve Refresh Token'
        }
    }
}