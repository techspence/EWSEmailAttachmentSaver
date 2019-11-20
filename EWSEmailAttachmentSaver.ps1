<#
  .SYNOPSIS
	This script looks for specific emails in an exchange users mailbox, downloads the attachments, 
	then marks those emails as read and moves the messages to a processed folder for archiving.

           Name: EWS Email Attachment Saver
         Author: Spencer Alessi (@techspence)
        License: MIT License
    Assumptions: The 'processed folder' is a subfolder of the root of the users mailbox (e.g. \\email@company.com\ProcessedFolder)
   Requirements: Exchange 2007 or newer
			     			 Exchange Web Services (EWS) Managed API 2.2

  .DESCRIPTION
	In general this script:
		1. Determines the Folder ID of the $processedfolderpath
		2. Finds the correct email messages based on defined search filters (e.g. unread, subject, has attachments)
		3. Copy's the attachments to the appropriate download location(s)
		4. Mark emails as read and move to the processed folder
  
  .NOTES
	The 'processed folder' is a subfolder of the root of the users mailbox (e.g. \\email@company.com\ProcessedFolder). 
	The root of a users mailbox is called the Top Information Store. If your 'processed folder' is a subfolder under any other 
	folder you must change $processedfolderpath and $tftargetidroot appropriately. 

	In this example, the processed folder is a subfolder of the root mailbox: Location: \\\email@company.com\ProcessedFolder

		$processedfolderpath = "/ProcessedFolder"
		$tftargetidroot = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$mailbox)

	In this example, the processed folder is a subfolder of Inbox: Location: \\\email@company.com\Inbox\ProcessedFolder
		
		$processedfolderpath = "/Inbox/ProcessedFolder"
		$tftargetidroot = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$processedfolderpath)
#>

# User defined variables. Change these to fit your needs
$mailbox = "email@company.com"
$user = $env:USERNAME
$reportroot = "c:\Users\$user\Downloads\"
$logpath = "c:\Users\$user\Downloads\logs\"
$logname = "EWSAttachmentSaver-$(get-date -f yyyy-MM-dd).log"
$logfile = $logpath + $logname
$processedfolderpath = "/Processed"
$subjectfilter = "Patch Report"
$datestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")

Function LogWrite
{
	Param ([string]$logstring)
	
	if (!(Test-Path $logpath)) {
		New-Item -ItemType Directory $logpath | Out-Null
	} 
	else { 
		Add-content $logfile -value $logstring
	}
}

Function FindTargetFolder($folderpath){
	$tftargetidroot = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$mailbox)
	$tftargetfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeservice,$tftargetidroot)
  $pfarray = $folderpath.Split("/")
	
	# Loop processed folders path until target folder is found
	for ($i = 1; $i -lt $pfarray.Length; $i++){
		$fvfolderview = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
		$sfsearchfilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$pfarray[$i])
    $findfolderresults = $exchangeservice.FindFolders($tftargetfolder.Id,$sfsearchfilter,$fvfolderview)
		
		if ($findfolderresults.TotalCount -gt 0){
			foreach ($folder in $findfolderresults.Folders){
				$tftargetfolder = $folder				
			}
		}
		else {
			LogWrite "### Error ###"
			Logwrite $datestamp " : Folder Not Found"
			$tftargetfolder = $null
			break
		}	
	}
	$Global:findFolder = $tfTargetFolder
}

Function FindTargetEmail($subject){
	foreach ($email in $foundemails.Items){
		$email.Load()
		$attachments = $email.Attachments

		foreach ($attachment in $attachments){
			$attachment.Load()
			$attachmentname = $attachment.Name.ToString()

			# Find the year based on the attachment name
			$fileyear = $attachment.Name.Split("-")[0]
			$fileyear = $fileyear.Split("_")[1]
			
			# Instead of (01, 02, 03, etc.) I want (1, 2, 3, etc.) for monthly folder names
			# Why? Because... :D
			$filemonth = $attachment.Name.Split("-")[1]
			if ($filemonth -eq "10"){}
			elseif ($filemonth -gt "10"){} 
			else {$filemonth = $filemonth.Split("0")[1]}

			if ($attachmentname -like '*Workstation*'){
				$workstationreportfolder = $reportroot + "Workstations\" + $fileyear + "\" + $filemonth + "\"
				if (!(Test-Path $workstationreportfolder)) {
					LogWrite "$workstationreportfolder not found.."
					LogWrite "Creating it now.."
					New-Item -ItemType Directory $workstationreportfolder | Out-Null
				}
				LogWrite "$attachmentname saved to $workstationreportfolder"
				$file = New-Object System.IO.FileStream(($workstationreportfolder + $attachmentname), [System.IO.FileMode]::Create)
				$file.Write($attachment.Content, 0, $attachment.Content.Length)
				$file.Close()
			} elseif ($attachmentname -like '*Server*'){
				$serverreportfolder = $reportroot + "Servers\" + $fileyear + "\" + $filemonth + "\"
				if (!(Test-Path $serverreportfolder)) {
					LogWrite "$serverreportfolder not found.."
					LogWrite "Creating it now.."
					New-Item -ItemType Directory $serverreportfolder | Out-Null
				}
				LogWrite "$attachmentname saved to $serverreportfolder"	
				$file = New-Object System.IO.FileStream(($serverreportfolder + $attachmentname), [System.IO.FileMode]::Create)	
				$file.Write($attachment.Content, 0, $attachment.Content.Length)
				$file.Close()
			} else {}
		}
	# Mark email as read & move to processed folder
	$email.IsRead = $true
	$email.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
	[VOID]$email.Move($Global:findFolder.Id)
	}
}

LogWrite "DATETIME: $datestamp"
LogWrite "Mailbox: $mailbox"
LogWrite "Report Root: $reportroot"
LogWrite "Processed Folder: $processedfolderpath"
LogWrite "Subject Filter: $subjectfilter"

# Load the EWS Managed API
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)

# Create EWS Service object for the target mailbox name
# Note, ExchangeVersion does not need to match the version of your Exchange server
# You set the version to indicate the lowest level of service you support
$exchangeservice = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
$exchangeservice.UseDefaultCredentials = $true
$exchangeservice.AutodiscoverUrl($mailbox)

# Bind to the Inbox folder of the target mailbox
$inboxfolderid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$mailbox)
$inboxfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeservice,$inboxfolderid)

# Search the Inbox for messages that are: unread, has specific subject AND has attachment(s)
$sfunread = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead, $false)
$sfsubject = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring ([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject, $subjectfilter)
$sfattachment = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, $true)
$sfcollection = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);
$sfcollection.add($sfunread)
$sfcollection.add($sfsubject)
$sfcollection.add($sfattachment)

# Use -ArgumentList 10 to reduce query overhead by viewing the Inbox 10 items at a time
$view = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList 10
$foundemails = $inboxfolder.FindItems($sfcollection,$view)

# Find $processedfolderpath Folder ID
FindTargetFolder($processedfolderpath)

# Process found emails
FindTargetEmail($subject)
