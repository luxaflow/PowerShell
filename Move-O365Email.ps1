<#
<-Move-O365Email->
.SYNOPSIS
Move email to other folder based of folder id's with a POST method

.DESCRIPTION
Folder is'd can be fetched with the Get-O365MailFolder

.PARAMETER AccessToken 
Full token that is required this is not the path to the token but the content of the token

.PARAMETER Mailbox 
Mailbox the folder should be found in

.PARAMETER EmailId 
The Graph API email id which can be found with the Get-O365Email function

.PARAMETER DestiantionId 
Shpuld be the id for the destionation mail folder

.PARAMETER ParentURI
This is API uri of MSGraph

.EXAMPLE
Update-O365Email -Mailbox 'MS.7x24@sltn.nl' -EmailId <Email id that needs to be patched> -AccessToken <full token> -DestiantionId <targets_id>

will mark the message as read
#>
function Move-O365Email {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true)][String]$AccessToken,
        [Parameter(
            Mandatory = $true)][System.Net.Mail.MailAddress]$Mailbox,
        [Parameter(
            Mandatory = $true)][String]$EmailId,
        [Parameter(
            Mandatory = $true)]$DestiantionId,
        [Parameter()]$ParentURI = 'https://graph.microsoft.com/v1.0/'
    )

    try {
        $Body = [pscustomobject]@{"destinationId" = $DestiantionId } | ConvertTo-Json
    }
    catch {
        throw $_
    }

    if ($Mailbox) {
        $URI = $ParentURI + "users/$Mailbox/messages/$EmailId/move"
    }
    else {
        $URI = $ParentURI + "me/messages/$EmailId/move"
    }
    

    $Params = @{
        Method = 'POST'
        Header = @{
            Authorization  = "Bearer $AccessToken"
            'Content-Type' = 'application/json'
        } 
        Body   = $body
        URI    = $URI

    }

    return Invoke-RestMethod @Params 
}