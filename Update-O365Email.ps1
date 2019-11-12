function Update-O365Email {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true)][String]$AccessToken,
        [Parameter(
            Mandatory = $false)][System.Net.Mail.MailAddress]$Mailbox,
        [Parameter(
            Mandatory = $true)][String]$EmailId,
        [Parameter(
            Mandatory = $true)]$UpdateObject,
        [Parameter()]$ParentURI = 'https://graph.microsoft.com/v1.0/'
    )

    try {
        $Body = $UpdateObject | ConvertTo-Json
    }
    catch {
        throw $_
    }

    if ($Mailbox) {
        $URI = $ParentURI + "users/$Mailbox/messages/$EmailId"
    }
    else {
        $URI = $ParentURI + "me/messages/$EmailId"
    }
    
    $Params = @{
        Method = 'Patch'
        Header = @{
            Authorization  = "Bearer $AccessToken"
            'Content-Type' = 'application/json'
        } 
        Body   = $body
        URI    = $URI

    }

    return Invoke-RestMethod @Params 
}