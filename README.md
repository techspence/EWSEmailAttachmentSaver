
# EWS Email Attachment Saver
This script looks for specific emails in an exchange users mailbox, downloads the attachments, then marks those emails as read and moves the messages to a processed folder for archiving.

**Outline**
1. Determines the Folder ID of the _processed folder_
2. Finds the correct email messages based on defined search filters (e.g. unread, subject, has attachments)
3. Copy's the attachments to the appropriate download location(s)
4. Mark emails as read and move to the processed folder

Check out my Blog Post: [Using Powershell and Microsoft EWS Managed API to download attachments in Exchange 2016](https://spenceralessi.com/Using-Powershell-and-Microsoft-EWS-Managed-API-to-download-attachments-in-Exchange-2016)

## Requirements
- Exchange 2007 or newer
- Exchange Web Services (EWS) Managed API 2.2

## Additional Information
The _processed folder_ is a subfolder of the root of the users mailbox (e.g. `\\email@company.com\ProcessedFolder`). The root of a users mailbox is called the _Top Information Store_. If your _processed folder_ is a subfolder under any other folder you must change `$processedfolderpath` and `$tftargetidroot` appropriately.

To quickly view the outlook folder location, right click on a folder in outlook, then click properties.

**Example: processed folder is a subfolder of the root mailbox:** `Location: \\email@company.com\ProcessedFolder`

```Powershell
$processedfolderpath = "/ProcessedFolder"
$tftargetidroot = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$mailbox)
```
**Example, processed folder is a subfolder of Inbox:** `Location: \\email@company.com\Inbox\ProcessedFolder`

```Powershell    
$processedfolderpath = "/Inbox/ProcessedFolder"
$tftargetidroot = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$processedfolderpath)
```

## Future Enhancements
- Create a Windows Service that will run the EWS Attachment Saver on an interval to check for applicable emails

## Credits/Resources
- [Download the Microsoft Exchange Web Services Managed API 2.2 from](http://www.microsoft.com/en-us/download/details.aspx?id=42951)

- [Microsoft EWS Managed API Reference](http://msdn.microsoft.com/en-us/library/jj220535(v=exchg.80).aspx)

- [Using PowerShell and EWS to monitor a mailbox](https://seanonit.wordpress.com/2014/10/29/using-powershell-and-ews-to-monitor-a-mailbox/)

- [EWS Managed API and Powershell How-To series Part 1](https://gsexdev.blogspot.com/2012/01/ews-managed-api-and-powershell-how-to.html)

- [Writing a simple scripted process to download attachments in Exchange 2007/ 2010 using the EWS Managed API](https://gsexdev.blogspot.com/2010/01/writing-simple-scripted-process-to.html)