<#
<-Get-O365Email->
.SYNOPSIS
Gets a folder object from the MSGraph API bases on a folder path

.DESCRIPTION
Folder path uses \ like a UNC path only no \ is required before the first folder.
this will only return the last folder of the path

.PARAMETER AccessToken 
Full token that is required this is not the path to the token but the content of the token

.PARAMETER Mailbox 
Mailbox the folder should be found in

.PARAMETER MailFolderId 
The Graph API folder id which can be found with the Get-O365MailFolder function

.PARAMETER ReceivedDateTime 
When used this will make sure only e-mails asof the given datetime are returned

.PARAMETER ReturnSize
the amount of emails that is allowed to be returned default is 10 and max is 1000

.PARAMETER ParentURI
This is API uri of MSGraph

.EXAMPLE
Get-O365Email -Mailbox 'MS.7x24@sltn.nl' -MailFolderPath 'Inbox\Subfolder\Subfolder' -AccessToken <full token> -ReceivedDateTime (get-date).adddays(-1) -returnsize 500

returns API email objects
#>
Function Get-O365Email {

    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true)][String]$AccessToken,
        [Parameter(
            Mandatory = $false)][System.Net.Mail.MailAddress]$Mailbox,
        [Parameter(
            Mandatory = $true)][String]$MailFolderId,
        [Parameter(
            Mandatory = $false)][Datetime]$ReceivedDateTime,
        [Parameter(
            Mandatory = $false)][Int]$ReturnSize,
        [Parameter()]$ParentURI = 'https://graph.microsoft.com/v1.0/'
    )
    begin {

        if ($ReceivedDateTime) {
            [String]$ReceivedDateTime = $ReceivedDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ')
        }
       
        if (!$Mailbox) {
            $BaseURI = $ParentURI + "me/"            
        }
        else {
            $BaseURI = $ParentURI + "users/$Mailbox/"
        }
        
        if ($ReturnSize -gt 1000) {
            $ReturnSize = 1000
        }
        
    }
    process {
        if ($ReceivedDateTime -and $ReturnSize) {
            $Query = $BaseURI + 'mailfolders/' + $MailFolderId + '/messages?$filter=receivedDateTime gt ' + $ReceivedDateTime + '&$top=' + $ReturnSize
        }
        elseif ($ReceivedDateTime) {
            $Query = $BaseURI + 'mailfolders/' + $MailFolderId + '/messages?$filter=receivedDateTime gt ' + $ReceivedDateTime
        }
        elseif ($ReturnSize) {
            $Query = $BaseURI + 'mailfolders/' + $MailFolderId + '/messages?$top=' + $ReturnSize
        }
        else {
            $Query = $BaseURI + 'mailfolders/' + $MailFolderId + '/messages'
        }

        $params = @{
            Method  = "Get"
            Headers = @{
                Authorization  = "Bearer $AccessToken"
                "Content-Type" = "application/json"
            }
            URI     = $Query;
        }
        
        $Emails = Invoke-RestMethod @params | ForEach-Object { $_.Value } 

    }
    end {
        return $Emails | Sort-Object receivedDateTime -Descending
    }
}
