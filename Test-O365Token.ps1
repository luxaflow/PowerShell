<#
<-Test-O365Tokens->
.SYNOPSIS
Checks if access token is still valid and if not updates with the refresh token

.DESCRIPTION
Uses the Update-O365Tokens function to update tokens if required

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

.NOTES
requires additional fuctions
Update-O365Tokens
#>
function Test-O365Tokens {
    [CmdletBinding()]
    Param(
        [Parameter()][String]$ClientSecret = 'Your MS Graph Client Secret',
        [Parameter()][String]$AccessTokenName = 'filen name of you access token',
        [Parameter()][String]$RefreshTokenName = 'file name of you refresh token',
        [Parameter()][String]$ClientId = 'Your MS Graph Client Id',
        [Parameter()][String]$RedirectUri = 'your MS Graph redirect URI',
        [Parameter()][String]$TokenDirectory = 'c:\tokens\'
    )

    try {
        $AccessToken = get-content -Path(Join-Path -Path $TokenDirectory -ChildPath $AccessTokenName)
        $Request = @{
            Method  = 'GET'
            Headers = @{
                Authorization  = "Bearer $Accesstoken"
                "Content-Type" = "application/json"
            }
            URI     = 'https://graph.microsoft.com/v1.0/me'
        }

        Invoke-RestMethod @Request
        return $false
    }
    catch {
        
        try {
            Update-O365Tokens 
        }
        catch {
            throw $_
        }
        return $true
    }
}