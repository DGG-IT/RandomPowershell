<#

.SYNOPSIS

Modified By: Dan Gustafson (BlueMountain Capital)
Based On:    https://ingogegenwarth.wordpress.com/
Version:	 1
Changed:	 12.23.2016

.LINK
http://www.get-blog.com/?p=189
http://learn-powershell.net/2012/05/13/using-background-runspaces-instead-of-psjobs-for-better-performance/
http://gsexdev.blogspot.de/

.DESCRIPTION

The purpose of the script is to get for each folder granted permissions for given mailboxes

.PARAMETER EmailAddress

The e-mail address of the mailbox, which will be checked. The script accepts piped objects from Get-Mailbox or Get-Recipient

.PARAMETER Credentials

Credentials you want to use. If omitted current user context will be used.

.PARAMETER Impersonate

Use this switch, when you want to impersonate.

.PARAMETER RootFolder

From where you want to start the search for folders. Default is "MsgFolderRoot". Other posible value is "Root".

.PARAMETER Server

By default the script tries to retrieve the EWS endpoint via Autodiscover. If you want to run the script against a specific server, just provide the name in this parameter. Not the URL!

.PARAMETER Threads

How many threads will be created. Be careful as this is really CPU intensive! By default 20 threads will be created. I limit the number of threads to 40.

.PARAMETER MultiThread

If you want to run the script multi-threaded use this switch. By default the script will use threads.

.PARAMETER MaxResultTime

The timeout for a job, when using multi-threads. Default is 240 seconds.

.PARAMETER TrustAnySSL

Switch to trust any certificate.

.PARAMETER Subject

(Mandatory) The Exact Subject of the message (in quotes).  Partial matches will not be returned.

.PARAMETER Sender

(Optional) If specified, only matches items sent by a specific email address.  Use this if the target email has a generic subject 

.PARAMETER jobName

(Mandatory) Name the job.  A folder will be created as C:\temp\$jobName\ to store the exported email messages.

.PARAMETER reallyDelete

(Mandatory) Whether to actually delete the message or just export the message.  Default is $False

.EXAMPLE

#run the script against a single mailbox
.\Delete-MailboxItemEWS.PS1 -EmailAddress trick@adatum.com

#pipe all the mailboxes on database USEX01-DB01 into the script
Get-Mailbox -Database USEX01-DB01 -ResultSize unlimited | .\Delete-MailboxItemEWS.PS1

#pipe all the mailboxes on database USEX01-DB01 into the script and use 10 threads
Get-Mailbox -Database USEX01-DB01 -ResultSize unlimited | .\Delete-MailboxItemEWS.PS1 -Threads 10 -MultiThread

#pipe all the mailboxes on database USEX01-DB01 into the script and use 10 threads
Get-Mailbox -Database USEX01-DB01 -ResultSize unlimited | .\Delete-MailboxItemEWS.PS1 -Threads 10 -MultiThread

#same as the previous, but this time using different credentials and impersonate
$SA= Get-Credential adatum\serviceaccount
Get-Mailbox -Database USEX01-DB01 -ResultSize unlimited | .\Delete-MailboxItemEWS.PS1 -Threads 10 -MultiThread -Server ex01.adatum.com -UseDefaultCred 0 -Impersonate -Credentials $SA

#to export the result just use the Export-CSV CmdLet
Get-Mailbox -Database USEX01-DB01 -ResultSize unlimited | .\Delete-MailboxItemEWS.PS1 | Export-Csv -NoTypeInformation -Path .\result.csv

.NOTES

Most Common usage will be as follows.

Get-Mailbox -ResultSize unlimited | .\v4_Delete-MailboxItemEWS.PS1 -MultiThread -Threads 20 -Subject "My maLicIous Email Subject - Document" -jobName SPAM-1234 -reallyDelete $True


#>

[CmdletBinding()]
Param (
	#[parameter( ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$true,Mandatory=$true, Position=0)]
 	[parameter(ValueFromPipelineByPropertyName=$true,Mandatory=$true, Position=0)]
	[Alias('PrimarySmtpAddress')]
	$EmailAddress,

	[parameter( Mandatory=$false, Position=1)]
	[System.Management.Automation.PsCredential]$Credentials,

	[parameter( Mandatory=$false, Position=2)]
	[switch]$Impersonate=$false,

	[parameter( Mandatory=$false, Position=3)]
	[ValidateSet("MsgFolderRoot", "Root")]
	[string]$RootFolder="MsgFolderRoot",

	[parameter( Mandatory=$false, Position=4)]
	[string]$Server,

	[parameter( Mandatory=$false, Position=5)]
	[ValidateRange(0,40)]
	[int]$Threads= '20',

	[parameter( Mandatory=$false, Position=6)]
	[switch]$MultiThread=$true,

	[parameter( Mandatory=$false, Position=7)]
	$MaxResultTime='240',

	[parameter( Mandatory=$false, Position=8)]
	[switch]$TrustAnySSL ,

    [parameter( Mandatory=$True, Position=9)]
	[string]$Subject ,

    [parameter( Mandatory=$false, Position=10)]
	[string]$Sender ,

    [parameter( Mandatory=$True, Position=11)]
	[bool]$reallyDelete=$false ,

    [parameter( Mandatory=$True, Position=12)]
	[string]$jobName 

)

Begin {

#initiate runspace and make sure we are using single-threaded apartment STA
$Jobs = @()
$Sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
$RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $Threads,$Sessionstate, $Host)
$RunspacePool.ApartmentState = "STA"
$RunspacePool.Open()
[int]$j='1'
}

Process {

#start function Delete-MailboxItem
function Delete-MailboxItem {
Param(
 	[parameter( ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$true,Mandatory=$true, Position=0)]
	[Alias('PrimarySmtpAddress')]
	[String]$EmailAddress,

	[parameter( Mandatory=$false, Position=1)]
	[System.Management.Automation.PsCredential]$Credentials,

	[parameter( Mandatory=$false, Position=2)]
	[bool]$Impersonate=$false,

	[parameter( Mandatory=$false, Position=3)]
	[ValidateSet("MsgFolderRoot", "Root")]
	[string]$RootFolder="MsgFolderRoot",

	[parameter( Mandatory=$false, Position=4)]
	[string]$Server,

	[parameter( Mandatory=$false, Position=5)]
	[bool]$TrustAnySSL,

	[parameter( Mandatory=$false, Position=6)]
	[int]$ProgressID ,

    [parameter( Mandatory=$true, Position=7)]
	[string]$Subject ,

    [parameter( Mandatory=$false, Position=8)]
	[string]$Sender ,

    [parameter( Mandatory=$true, Position=9)]
	[bool]$reallyDelete ,

    [parameter( Mandatory=$True, Position=10)]
	[string]$jobName 

)

try {
$MailboxName = $EmailAddress
## Load Managed API dll
###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
$EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
if (Test-Path $EWSDLL)
    {
    Import-Module $EWSDLL
    }
else
    {
    "$(get-date -format yyyyMMddHHmmss):"
    "This script requires the EWS Managed API 1.2 or later."
    "Please download and install the current version of the EWS Managed API from"
    "http://go.microsoft.com/fwlink/?LinkId=255472"
    ""
    "Exiting Script."
    exit
    }

## Set Exchange Version
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2

## Create Exchange Service Object
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
#$service.PreAuthenticate = $true
## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials
If ($Credentials) {
	#Credentials Option 1 using UPN for the windows Account
	#$psCred = Get-Credential
	$psCred = $Credentials
	$creds = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString())
	$service.Credentials = $creds
	#$service.TraceEnabled = $true

}
Else {
	#Credentials Option 2
	$service.UseDefaultCredentials = $true
}

If ($TrustAnySSL) {
	## Choose to ignore any SSL Warning issues caused by Self Signed Certificates

	## Code From http://poshcode.org/624
	## Create a compilation environment
	$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
	$Compiler=$Provider.CreateCompiler()
	$Params=New-Object System.CodeDom.Compiler.CompilerParameters
	$Params.GenerateExecutable=$False
	$Params.GenerateInMemory=$True
	$Params.IncludeDebugInformation=$False
	$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

	$TASource=@'
	namespace Local.ToolkitExtensions.Net.CertificatePolicy{
		public class TrustAll : System.Net.ICertificatePolicy {
		public TrustAll() {
		}
		public bool CheckValidationResult(System.Net.ServicePoint sp,
			System.Security.Cryptography.X509Certificates.X509Certificate cert,
			System.Net.WebRequest req, int problem) {
			return true;
		}
		}
	}
'@
	$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
	$TAAssembly=$TAResults.CompiledAssembly

	## We now create an instance of the TrustAll and attach it to the ServicePointManager
	$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
	[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

	## end code from http://poshcode.org/624
}

## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use
If ($Server) {
	#CAS URL Option 2 Hardcoded
	$uri=[system.URI] "https://$server/ews/exchange.asmx"
	$service.Url = $uri
}
Else {
	#CAS URL Option 1 Autodiscover
	$service.AutodiscoverUrl($MailboxName,{$true})
	#"Using CAS Server : " + $Service.url
}

## Optional section for Exchange Impersonation
If ($Impersonate) {
	$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
}


$fidFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)
Write-Host "fidFolderId: $fidFolderId"

$folderidcnt = $fidFolderId

#Define the FolderView used for Export should not be any larger then 1000 folders due to throttling
$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
#Deep Transval will ensure all folders in the search path are returned
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;


$psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent)

$fsfFolderSearchFilter= New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo( [Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, "IPF.Note")

#Create the Search Filter by Subject
$subjectsfSubjectSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo( [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject,$Subject)

#Create the Search Filter Collection and add the Subject-based search filter to the collection
$sfcSearchFilterCollection = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection( [Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
$sfcSearchFilterCollection.Add($subjectsfSubjectSearchFilter)

#If Sender is specified, add the new search filter to the Search Filter Collection
if($Sender) {
$sendersfSenderSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo( [Microsoft.Exchange.WebServices.Data.ItemSchema]::Sender,$Sender)
$sfcSearchFilterCollection.Add($sendersfSenderSearchFilter)
}

$fiResult = $null
#loop through folders, when more than 1000
do {
$fiResult = $Service.FindFolders($folderidcnt,$fsfSearchFilter,$fvFolderView)

[int]$i='1'
ForEach ($Folder in $fiResult) {

	Write-Progress `
			-id $ProgressID `
			-ParentId 1 `
			-Activity "Processing mailbox - $($MailboxName) with $($fiResult.Folders.count) folders" `
			-PercentComplete ( $i / $fiResult.Folders.count * 100) `
			-Status "Remaining: $($fiResult.Folders.count - $i) processing folder: $($Folder.DisplayName)"
	If (($Folder.Displayname -ne 'System') -and ($Folder.Displayname -ne 'Audits')) {
	
        $ItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(10000)
        
        $firFindItemResults = $null
        do {
            $firFindItemResults = $Folder.FindItems($subjectsfSubjectSearchFilter,$ItemView)
            foreach ($item in $firFindItemResults.Items) {
                try {
                    #ExportMessage
                    $messageSubject = $item.Subject
                    $messageId = $item.Id
                    $itemDateTime = $item.DateTimeReceived
	                $global:msMessage = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service,$messageId,$psPropset);
                    $jobPath = "C:\temp\$jobName"
                    if ( -Not (Test-Path $jobPath.trim() )) {
                        New-Item -Path $jobPath -ItemType Directory | Out-Null
                    }
                    $ienItemExportName = ("C:\temp\$jobName\$MailboxName-$MessageID.eml")
		            $ieItemExport = new-object System.IO.FileStream($ienItemExportName, [System.IO.FileMode]::Create)
                    $ieItemExport.Write($global:msMessage.MimeContent.Content, 0, $global:msMessage.MimeContent.Content.Length)
                    $ieItemExport.Close()

                    #Delete if ReallyDelete
                    if ($reallyDelete){
                        
                        #Comment Below to Neuter Script
                        [void]$item.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)
                    }
                }
                catch {
                    Write-warning "Unable to delete item, $($item.subject). $($Error[0].Exception.Message)"
                }
            }
            $ItemView.Offset += $firfindItemResults.Items.Count
        }
        while ($firFindItemResults.MoreAvailable)
        }
        #Write-Progress -id $ProgressID -ParentId 1 -Activity "Processing mailbox - $($MailboxName) with $($fiResult.Folders.count) folders" -Status "Ready" -Completed
    }
    $fvFolderView.Offset += $fiResult.Folders.Count
}
while($fiResult.MoreAvailable -eq $true)

}
catch{
	$Error[0].Exception
}
Write-Progress -id $ProgressID -ParentId 1 -Activity "Processing mailbox - $($MailboxName) with $($fiResult.Folders.count) folders" -Status "Ready" -Completed
#end function Delete-MailboxItem
}

#if multi-threaded create jobs
If ($MultiThread) {
	#create scriptblock from function
	$ScriptBlock = [scriptblock]::Create((Get-ChildItem Function:\Delete-MailboxItem).Definition)
	ForEach($Address in $EmailAddress) {
		try{
			$j++ | Out-Null
			#Write-Host "Adding job for "$Address
			$MailboxName = $Address
			$PowershellThread = [powershell]::Create().AddScript($ScriptBlock).AddParameter('EmailAddress',$MailboxName)
			$PowershellThread.AddParameter('Credentials',$Credentials) | Out-Null
			$PowershellThread.AddParameter('Impersonate',$Impersonate) | Out-Null
			$PowershellThread.AddParameter('RootFolder',$RootFolder) | Out-Null
			$PowershellThread.AddParameter('Server',$Server) | Out-Null
			$PowershellThread.AddParameter('TrustAnySSL',$TrustAnySSL) | Out-Null
			$PowershellThread.AddParameter('ProgressID',$j) | Out-Null
            $PowershellThread.AddParameter('Subject',$Subject) | Out-Null
			$PowershellThread.AddParameter('Sender',$Sender) | Out-Null
			$PowershellThread.AddParameter('reallyDelete',$reallyDelete) | Out-Null
            $PowershellThread.AddParameter('jobName',$jobName) | Out-Null
			$PowershellThread.RunspacePool = $RunspacePool
			$Handle = $PowershellThread.BeginInvoke()
			$Job = "" | Select-Object Handle, Thread, object
			$Job.Handle = $Handle
			$Job.Thread = $PowershellThread
			$Job.Object = $Address #.ToString()
			$Jobs += $Job
		}
		catch {
			$Error[0].Exception
		}
	}
}
#if not mutli-threaded start sequential processing
Else{
	ForEach($Address in $EmailAddress) {
		$MailboxName = $Address
		Delete-MailboxItem -EmailAddress $MailboxName -Credentials $Credentials -Impersonate $Impersonate `
		-RootFolder $RootFolder -Server $Server -TrustAnySSL $TrustAnySSL -Subject $Subject -Sender $Sender -reallyDelete $reallyDelete -jobName $jobName
	}
}
}

End{

#monitor and retrieve the created jobs
If ($MultiThread) {
	$SleepTimer = 200
	$ResultTimer = Get-Date
	While (@($Jobs | Where-Object {$_.Handle -ne $Null}).count -gt 0)  {
	$Remaining = "$($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False}).object)"
		If ($Remaining.Length -gt 60){
			$Remaining = $Remaining.Substring(0,60) + "..."
		}
		Write-Progress `
			-id 1 `
			-Activity "Waiting for Jobs - $($Threads - $($RunspacePool.GetAvailableRunspaces())) of $Threads threads running" `
			-PercentComplete (($Jobs.count - $($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False}).count)) / $Jobs.Count * 100) `
			-Status "$(@($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False})).count) remaining - $Remaining"

		ForEach ($Job in $($Jobs | Where-Object {$_.Handle.IsCompleted -eq $True})){
			$Job.Thread.EndInvoke($Job.Handle)
			$Job.Thread.Dispose()
			$Job.Thread = $Null
			$Job.Handle = $Null
			$ResultTimer = Get-Date
		}
		If (($(Get-Date) - $ResultTimer).totalseconds -gt $MaxResultTime){
			Write-Warning "Child script appears to be frozen for $($Job.Object), try increasing MaxResultTime"
			#Exit
		}
		Start-Sleep -Milliseconds $SleepTimer
	# kill all incomplete threads when hit "CTRL+q"
	If ($Host.UI.RawUI.KeyAvailable) {
		$KeyInput = $Host.UI.RawUI.ReadKey("IncludeKeyUp,NoEcho")
		If (($KeyInput.ControlKeyState -cmatch '(Right|Left)CtrlPressed') -and ($KeyInput.VirtualKeyCode -eq '81')) {
			Write-Host -fore red "Kill all incomplete threads....."
				ForEach ($Job in $($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False})){
					Write-Host -fore yellow "Stopping job $($Job.Object) ...."
					$Job.Thread.Stop()
					$Job.Thread.Dispose()
				}
			Write-Host -fore red "Exit script now!"
			Exit
		}
	}
	}
	# clean-up
	$RunspacePool.Close() | Out-Null
	$RunspacePool.Dispose() | Out-Null
	[System.GC]::Collect()
}
}
