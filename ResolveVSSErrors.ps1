### Declare variables

$WriterReferenceObjects = @()

$WriterReference = @{
	"ASR Writer"							= "VSS"
	"BITS Writer"							= "BITS"
	"COM+ REGDB Writer"						= "VSS"
	"DFS Replication Service Writer"		= "DFSR"
	"DHCP Jet Writer"						= "DHCPServer"
	"FRS Writer"							= "NtFrs"
	"FSRM Writer"							= "srmsvc"
	"IIS Config Writer"						= "AppHostSvc"
	"IIS Metabase Writer"					= "IISADMIN"
	"MailStore VSS Writer"					= "MailStore VSS Writer"
	"Microsoft Exchange Writer"				= "MSExchangeIS"
	"Microsoft Hyper-V VSS Writer"			= "VMMS"
	"MSMQ Writer (MSMQ)"					= "MSMQ"
	"NTDS"									= "Ntds"
	"OSearch VSS Writer"					= "OSearch"
	"OSearch14 VSS Writer"					= "Osearch14"
	"OSearch15 VSS Writer"					= "OSearch15"
	"Pervasive.SQL VSS Writer"				= "Pervasive.SQL (relational),Pervasive.SQL (transactional)"
	"Registry Writer"						= "VSS"
	"Shadow Copy Optimization Writer"		= "VSS"
	"Sharepoint VSS Writer"					= "Sharepoint VSS Writer"
	"SPSearch VSS Writer"					= "SPSearch"
	"SPSearch4 VSS Writer"					= "SPSearch4"
	"SqlServerWriter"						= "SQLWriter"	
	"System Writer"							= "CryptSvc"
	"TermServLicensing"						= "TermServLicensing"
	"WINS Jet Writer"						= "WINS"
	"WMI Writer"							= "Winmgmt"
}


### Functions

function CreateVSSListWriterObjects() {
	$script:AllWriters = @()
	
	vssadmin list writers | 
		Select-String -Pattern "^$" -Context 0,5 | 
		Select-Object -SkipLast 1 |
		ForEach-Object {

			$VssWriter =  $_.Context.Postcontext.Split(':').Trim(" '")
		
			$script:AllWriters += [PsCustomObject] @{
				'WriterName'        = $VssWriter[1]
				'WriterId'          = [Guid]$VssWriter[3]
				'WriterInstanceId'  = [Guid]$VssWriter[5]
				'StateId'           = [int]$VssWriter[7].SubString($VssWriter[7].IndexOf('[') + 1, 1)
				'State'             = $VssWriter[7].SubString($VssWriter[7].IndexOf(']') + 2)
				'LastError'         = $VssWriter[9]
			}  

	}
}

function CreateVSSWriterReferenceObjects() {
	foreach ($Writer in $WriterReference.GetEnumerator()) {
		$props = @{
			WriterName	= $Writer.Name
			ServiceName = $Writer.Value
		}
		
		$script:WriterReferenceObjects += New-Object -TypeName PSObject -Property $props
	}
}

function ResetFailedWriters() {
	$AllWriters | ForEach-Object {
		$CurrentWriter = $_.WriterName
		if (($_.State -ne "Stable") -OR ($_.LastError -ne "No error")) {
			Write-Output "$CurrentWriter needs to be restarted..."
			$ServiceToRestart = ($WriterReferenceObjects | Where-Object { $_.WriterName -eq $CurrentWriter }).ServiceName
			Restart-Service -Name $ServiceToRestart -Force
		} else {
			Write-Output "$CurrentWriter appears to be in good shape."
		}
	}
}

function CheckForRemainingErrors() {
	$script:RemainingErrors = 0

	$AllWriters | ForEach-Object {
		$CurrentWriter = $_.WriterName
		if (($_.State -ne "Stable") -OR ($_.LastError -ne "No error")) {
			Write-Output "$CurrentWriter is still in error state."
			$script:RemainingErrors = 1
		}
	}
}


### Code Logic

CreateVSSListWriterObjects

CreateVSSWriterReferenceObjects

ResetFailedWriters

CreateVSSListWriterObjects

CheckForRemainingErrors

if ($RemainingErrors -eq 0) {
	Write-Output "All VSS writers errors appear to have been resolved."
} else {
	Write-Output "There were VSS writer errors that were not able to be resolved."
}