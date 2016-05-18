<#
.Synopsis
   Harvests emails, contacts, tasks, and calendar from MS outlook.
.DESCRIPTION
   Pulls either partial or full contents of a users inbox or sent folder, by default it pulls the inbox.
   Can also be set to pull contents of of the logged on user's outlook calendar, contacts, or tasks.
   It should be noted that this code WILL START OUTLOOK if the application isn't already running.
.NOTES
    Author: Andrew Cole
    Company: Chiron Technology Services, Inc.
    Credits: Large portions of code based on the below post by Ed Wilson, Microsoft's Scripting Guy:
        http://blogs.technet.com/b/heyscriptingguy/archive/2011/05/24/use-powershell-to-export-outlook-calendar-information.aspx
.EXAMPLE
   Get-Outlook

   This will pull a listing of the user's inbox.
.EXAMPLE
   Get-Outlook -calendar

   This will pull the user's outlook calendar events.
.EXAMPLE
   Get-Outlook -Sent -Full

   This will pull the users complete Sent Mail folder, to include all metadata associated with each email.
.EXAMPLE
   Get-Outlook -contacts

   This will pull the user's outlook contacts, to incluse names positions, titles, addresses, phone numbers and emails.
.EXAMPLE
   Get-Outlook -Tasks

   This will pull the user's outlook Tasks, both active and completed.
#>

Function Get-Outlook
{ 
  param(
    [switch]$Sent,
    [switch]$calendar,
    [switch]$contacts,
    [switch]$Tasks,
    [switch]$full
  ) 
    # This section creates a custom COM object we can use to data mine outlook
    Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null 
    $olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]  
    $outlook = new-object -comobject outlook.application 
    $namespace = $outlook.GetNameSpace("MAPI")
    
    # We can now point to the desired folder based on the switch called from the command line, be it tasks, contacts or calendar
    if($Calendar)
    {
        $folder = $namespace.getDefaultFolder($olFolders::olFolderCalendar)
        $folder.items | Select-Object -Property Subject, Start, Duration, Location
    }
    elseif($Contacts)
    {
        $folder = $namespace.getDefaultFolder($olFolders::olFolderContacts)
        $folder.items | Select-Object -Property FullName, JobTitle, companyName, businessAddress, BusinessTelephoneNumber, HomeTelephoneNumber, MobileTelephoneNumber, Email1Address, Email2Address, body
    }
    elseif($Tasks)
    {
        $folder = $namespace.getDefaultFolder($olFolders::olFolderTasks)
        $folder.items | Select-Object -Property Subject, CreationTime, LastModificationTime, StartDate, DateCompleted, DueDate, Categories, Owner, Importance, IsRecurring
    }
    
    # This section will define email harvesting, which requires a bit more information
    else
    {
        # Here we define the target folder, either Inbox or Sent
        if($Sent)
        {
            $targetFolder = 'olFolderSentMail'
        }
        else
        {
            $targetFolder = 'olFolderInbox'
        }
        # This section gets the target folder and determines whether to just pull a basic overview, or the full emails, including all metadata
        $folder = $namespace.getDefaultFolder($olFolders::$targetFolder) 
        if($full)
        {
            # This is a safety feature, to make sure the user realizes how much they may be grabbing with the full switch
            $response = Read-Host "WARNING: Full Output may be VERY large, and take an extended period of time.  Are you sure you want to continue? (Y or N)"
            if($response -eq 'Y')
            {
                $folder.items | Select-Object -Property * 
            }
            else
            {
                Write-Output "Cancelling -full switch, reverting to standard output."
                $folder.items | Format-Table -Property To, Subject, ReceivedTime, SenderName, body
            }
        }
        else
        {
            $folder.items | Format-Table -Property To, Subject, ReceivedTime, SenderName, body
        }
    }
} 