
<#
<-Get-O365Attachment->
.SYNOPSIS
Gets the attachment from a given e-mail and is preffered can be downloaded

.DESCRIPTION
Gets the attachment from a given e-mail id and if required downloads it and can tiemstamp the file for storage purposes

.PARAMETER Mailbox 
Mailbox the folder should be found in

.PARAMETER AccessToken 
Full token that is required this is not the path to the token but the content of the token

.PARAMETER EmailId 
The Graph API Email id which can be found with the Get-O365Email function

.PARAMETER AttachmentFileName 
Will filter fofr attachments only with the given name this uses a -like filter

.PARAMETER TimeStampFile
in case a timestamp needs to be added to the file name this switch can be used. The time stamp is that of the attachment object received time

.PARAMETER Download
A switch that will downlaod the file. if not used the Attachment object will be returned

.PARAMETER SavePath
In case of downloading the fils a directory they should be saved to

.PARAMETER ParentURI
This is API uri of MSGraph

.NOTES
required addtional functions
ConvertTo-TimeStamp
Update-O365Email (optional)
#>
Function Get-O365Attachment {
  
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true)][System.Net.Mail.MailAddress]$Mailbox,
        [Parameter(
            Mandatory = $true)][String]$AccessToken,
        [Parameter(
            Mandatory = $true)][String]$EmailId,
        [Parameter(
            Mandatory = $false)][String]$AttachmentFileName,
        [Parameter(
            Mandatory = $false)][DateTime]$TimeStampFile,
        [Parameter(
            Mandatory = $false)][Switch]$Download,
        [Parameter()][String]$SavePath = 'c:\',
        [Parameter()][String]$ParentURI = 'https://graph.microsoft.com/v1.0/'
            
    )
    begin {
        if (!$Mailbox) {
            $BaseURI = $ParentURI + "me/"            
        }
        else {
            $BaseURI = $ParentURI + "users/$Mailbox/"
        }
    } #End of begin
    process {
    
        $Query = "$BaseURI/messages/$EmailId/attachments"

        $Params = @{
            Method  = "Get";
            Headers = @{
                Authorization  = "Bearer $Accesstoken"
                "Content-Type" = "application/json"
            };
            URI     = $Query;
        }

        $Attachments = Invoke-RestMethod @Params
    }#end of process
    end {
        if ($Download) {
            
            if ($AttachmentFileName) {
                $attachment = $Attachments | Where-Object { $_.name -ilike "*$AttachmentFileName*" }

                if ($TimeStampFile) {
                    $timestamp = ConvertTo-TimeStamp -DateTimeObject $TimeStampFile
                    [String]$AttachmentName = $timestamp + '_' + $Attachment.Name.Replace('FW: ', '').replace('RE: ', '') 

                }
                else {
                    [String]$AttachmentName = $Attachment.Name.Replace('FW: ', '').replace('RE: ', '') 
                }
                
                [String]$Path = Join-Path -Path $SavePath -ChildPath $AttachmentName
                $Content = [System.Convert]::FromBase64String($attachment.ContentBytes)
                Set-Content -Path $Path -Value $Content -Encoding Byte
            }
            else {

                foreach ($attachment in $attachments.value) {
                    if ($TimeStampFile) {
                        $timestamp = ConvertTo-TimeStamp -DateTimeObject $TimeStampFile
                        [String]$AttachmentName = $timestamp + '_' + $Attachment.Name.Replace('FW: ', '').replace('RE: ', '') 

                    }
                    else {
                        [String]$AttachmentName = $Attachment.Name.Replace('FW: ', '').replace('RE: ', '') 
                    }
                
                    [String]$Path = Join-Path -Path $SavePath -ChildPath $AttachmentName
                    $Content = [System.Convert]::FromBase64String($attachment.ContentBytes)
                    Set-Content -Path $Path -Value $Content -Encoding Byte
                }
            }

            # updates email to read in case this is required
            # $Params = @{
            #     AccessToken  = $AccessToken
            #     Mailbox      = $Mailbox
            #     EmailId      = $EmailId
            #     UpdateObject = [PSCustomObject]@{
            #         isRead = $true
            #     }
            # }

            # Update-O365Email @Params | Out-Null
        }
        else {
            return $Attachments
        }
    }
}
