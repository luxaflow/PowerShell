<#
<-Get-O365MailFolder->
.SYNOPSIS
Gets a folder object from the MSGraph API bases on a folder path

.DESCRIPTION
Folder path uses \ like a UNC path only no \ is required before the first folder.
this will only return the last folder of the path

.PARAMETER Mailbox 
Mailbox the folder should be found in

.PARAMETER AccessToken 
Full token that is rquired this is not the path to the token but the content of the token

.PARAMETER ParentURI
This is API uri of MSGraph

.EXAMPLE
Get-O365MailFolder -Mailbox 'MS.7x24@sltn.nl' -MailFolderPath 'Inbox\Subfolder\Subfolder'

returns API Folder object
#>
Function Get-O365MailFolder {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $false)][System.Net.Mail.MailAddress]$Mailbox,
        [Parameter(
            Mandatory = $true)][String]$MailFolderPath,
        [Parameter(
            Mandatory = $true)][String]$AccessToken,
        [Parameter()][String]$ParentURI = 'https://graph.microsoft.com/v1.0/'
    ) 
    begin {
        if (!$Mailbox) {
            $BaseURI = $ParentURI + "me/"
        }
        else {
            $BaseURI = $ParentURI + "users/$Mailbox"
        }

        $MailFolders = $MailfolderPath.split("\")
    } #end of begin
    process {
        foreach ($Folder in $MailFolders) {
            if ($Folder -ieq "Inbox") {
                [string]$Query = $BaseURI + '/mailFolders?$filter=displayName eq ' + "'" + $Folder + "'"
        
                $Params = @{
                    Method  = "Get"
                    Headers = @{
                        Authorization  = "Bearer $Accesstoken"
                        "Content-Type" = "application/json"
                    }
                    URI     = $Query
                }
                $FolderObject = Invoke-RestMethod @Params | ForEach-Object { $_.Value } 
            }
            else {
                [string]$Query = $BaseURI + '/mailFolders/' + $FolderObject.id + '/childFolders?$filter=displayName eq ' + "'" + $Folder + "'" 
        
                $Params = @{
                    Method  = "Get"
                    Headers = @{
                        Authorization  = "Bearer $Accesstoken"
                        "Content-Type" = "application/json"
                    }
                    URI     = $Query
                }

                $FolderObject = Invoke-RestMethod @Params | ForEach-Object { $_.Value } 
            }
        }
    }
    end {  
        return $FolderObject
    } 
}
