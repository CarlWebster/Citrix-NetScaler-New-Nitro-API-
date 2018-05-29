#Requires -Version 3.0
#This File is in Unicode format.  Do not edit in an ASCII editor.

<#
.SYNOPSIS
    Creates a complete inventory of a Citrix NetScaler configuration using Microsoft Word.
.DESCRIPTION
	Creates a complete inventory of a Citrix NetScaler configuration using Microsoft Word and PowerShell.
	Creates a Word document named after the Citrix NetScaler Configuration.
	Document includes a Cover Page, Table of Contents and Footer.
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish

	Script requires at least PowerShell version 3 but runs best in version 5.

.PARAMETER NSIP
    NetScaler IP address, could be NSIP or SNIP with management enabled
.PARAMETER Credential
	NetScaler username
	
	Specifies a user name for the NetScaler credential, such as "User01" or "Domain01\User01".

	You are prompted for a password.

	If you omit this parameter, you are prompted for a user name and a password.	
.PARAMETER UseNSSSL
	EXPERIMENTAL: Require SSL/TLS for communication with the NetScaler Nitro API. 
	NOTE: This requires the client to trust the NetScaler's certificate chain.
	Note from Webster: This parameter is disabled at this time.
.PARAMETER CompanyName
	Company Name to use for the Cover Page.  
	Default value is contained in HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated on the 
	computer running the script.
	This parameter has an alias of CN.
	If either registry key does not exist and this parameter is not specified, the report will
	not contain a Company Name on the cover page.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	Only Word 2010, 2013 and 2016 are supported.
	(default cover pages in Word en-US)
	
	Valid input is:
		Alphabet (Word 2010. Works)
		Annual (Word 2010. Doesn't work well for this report)
		Austere (Word 2010. Works)
		Austin (Word 2010/2013/2016. Doesn't work in 2013 or 2016, mostly works in 2010 but 
			Subtitle/Subject & Author fields need to be moved after title box is moved up)
		Banded (Word 2013/2016. Works)
		Conservative (Word 2010. Works)
		Contrast (Word 2010. Works)
		Cubicles (Word 2010. Works)
		Exposure (Word 2010. Works if you like looking sideways)
		Facet (Word 2013/2016. Works)
		Filigree (Word 2013/2016. Works)
		Grid (Word 2010/2013/2016. Works in 2010)
		Integral (Word 2013/2016. Works)
		Ion (Dark) (Word 2013/2016. Top date doesn't fit; box needs to be manually resized or font 
			changed to 8 point)
		Ion (Light) (Word 2013/2016. Top date doesn't fit; box needs to be manually resized or font 
			changed to 8 point)
		Mod (Word 2010. Works)
		Motion (Word 2010/2013/2016. Works if top date is manually changed to 36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2010. Works)
		Puzzle (Word 2010. Top date doesn't fit; box needs to be manually resized or font 
			changed to 14 point)
		Retrospect (Word 2013/2016. Works)
		Semaphore (Word 2013/2016. Works)
		Sideline (Word 2010/2013/2016. Doesn't work in 2013 or 2016, works in 2010)
		Slice (Dark) (Word 2013/2016. Doesn't work)
		Slice (Light) (Word 2013/2016. Doesn't work)
		Stacks (Word 2010. Works)
		Tiles (Word 2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2010. Works)
		ViewMaster (Word 2013/2016. Works)
		Whisp (Word 2013/2016. Works)
		
	Default value is Sideline.
	This parameter has an alias of CP.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	The PDF file is roughly 5X to 10X larger than the DOCX file.
	This parameter requires Microsoft Word to be installed.
	This parameter uses the Word SaveAs PDF capability.
.PARAMETER MSWord
	SaveAs DOCX file
	This parameter is set True if no other output format is selected.
.PARAMETER AddDateTime
	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2016 at 6PM is 2016-06-01_1800.
	Output filename will be ReportName_2016-06-01_1800.docx (or .pdf).
	This parameter is disabled by default.
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
.PARAMETER SmtpServer
	Specifies the optional email server to send the output report. 
.PARAMETER SmtpPort
	Specifies the SMTP port. 
	Default is 25.
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	Default is False.
.PARAMETER From
	Specifies the username for the From email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER To
	Specifies the username for the To email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	Text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	Text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v3_5_Signed.ps1
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v3_5_Signed.ps1 -PDF
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript .\NetScaler_Script_v3_5_Signed.ps1 -CompanyName "Carl Webster Consulting" -CoverPage "Mod" -UserName "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
.EXAMPLE
	PS C:\PSScript .\NetScaler_Script_v3_5_Signed.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v3_5_Signed.ps1 -AddDateTime
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2016 at 6PM is 2016-06-01_1800.
	Output filename will be Script_Template_2016-06-01_1800.docx
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v3_5_Signed.ps1 -PDF -AddDateTime
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2016 at 6PM is 2016-06-01_1800.
	Output filename will be Script_Template_2016-06-01_1800.PDF
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v3_5_Signed.ps1 -Folder \\FileServer\ShareName
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Output file will be saved in the path \\FileServer\ShareName
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v3_5_Signed.ps1 -SmtpServer mail.domain.tld -From XDAdmin@domain.tld -To ITGroup@domain.tld
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will use the email server mail.domain.tld, sending from XDAdmin@domain.tld, sending to ITGroup@domain.tld.
	Script will use the default SMPTP port 25 and will not use SSL.
	If the current user's credentials are not valid to send email, the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v3_5_Signed.ps1 -SmtpServer smtp.office365.com -SmtpPort 587 -UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will use the email server smtp.office365.com on port 587 using SSL, sending from webster@carlwebster.com, sending to ITGroup@carlwebster.com.
	If the current user's credentials are not valid to send email, the user will be prompted to enter valid credentials.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  
	This script creates a Word, PDF, Formatted Text or HTML document.
.NOTES
	NAME: NetScaler_Script_V3_5_Signed.ps1
	VERSION: 3.5
	AUTHOR: Barry Schiffer, Carl Webster, Michael B. Smith, Iain Brighton, Jeff Wouters
	LASTEDIT: October 24, 2016
#>

#region changelog
<#
.COMMENT
    If you find issues with saving the final document or table layout is messed up please use the X86 version of Powershell!
.NetScaler Documentation Script
    NAME: NetScaler_Script_v3_1.ps1
	VERSION NetScaler Script: 3.1
	VERSION Script Template: 2016
	AUTHOR NetScaler script: Barry Schiffer
    AUTHOR NetScaler script functions: Iain Brighton
    AUTHOR Script template: Carl Webster, Michael B. Smith, Iain Brighton, Jeff Wouters
	LASTEDIT: August 25th 2016 

.Release Notes version 3.5
    Most work on version 3.5 has been done by Andy McCullough!
    After the release of version 3.0 in May 2016, which was a major overhaul of the NetScaler documentation script we found a few issues which have been fixed in the update.

    The script is now fully compatible with NetScaler 11.1 released in July 2016.

    * Added NetScaler functionality
    * Added NetScaler 11.1 Features, LSN / RDP Proxy / REP
    * Added Auditing Section
    * Added GSLB Section, vServer / Services / Sites
    * Added Locations Database section to support GSLB configuration using Static proximity.
    * Added additional DNS Records to the NetScaler DNS Section
    * Added RPC Nodes section
    * Added NetScaler SSL Chapter, moved existing functionality and added detailed information
    * Added AppFW Profiles and Policies
    * Added AAA vServers

    Added NetScaler Gateway functionality
    * Updated NSGW Global Settings Client Experience to include new parameters
    * Updated NSGW Global Settings Published Applications to include new parameters
    * Added Section NSGW "Global Settings AAA Parameters"
    * Added SSL Parameters section for NSGW Virtual Servers
    * Added Rewrite Policies section for each NSGW vServer
    * Updated CAG vServer basic configuration section to include new parameters
    * Updated NetScaler Gateway Session Action > Security to include new attributed
    * Added Section NetScaler Gateway Session Action > Client Experience
    * Added Section NetScaler Gateway Policies > NetScaler Gateway AlwaysON Policies
    * Added NSGW Bookmarks
    * Added NSGW Intranet IP's
    * Added NSGW Intranet Applications
    * Added NSGW SSL Ciphers

	Webster's Updates

	* Updated help text to match other documentation scripts
	* Removed all code related to TEXT and HTML output since Barry does not offer those
	* Added support for specifying an output folder to match other documentation scripts
	* Added support for the -Dev and -ScriptInfo parameters to match other documentation scripts
	* Added support for emailing the output file to match other documentation scripts
	* Removed unneeded functions
	* Brought script code in line with the other documentation scripts
	* Temporarily disabled the use of the UseNSSSL parameter
	
.Release Notes version 3
    Overall
        The script has had a major overhaul and is now completely utilizing the Nitro API instead of the NS.Conf.
        The Nitro API offers a lot more information and most important end result is much more predictable. Adding NetScaler functionality is also much easier.
        Added functionality because of Nitro
        * Hardware and license information
        * Complete routing tables including default routes
        * Complete monitoring information including default monitors
        * 
.Release Notes version 2
    Overall
        Test group has grown from 5 to 20 people. A lot more testing on a lot more configs has been done.
        The result is that I've received a lot of nitty gritty bugs that are now solved. To many to list them all but this release is very very stable.
    New Script functionality
        New table function that now utilizes native word tables. Looks a lot better and is way faster
        Performance improvements; over 500% faster
        Better support for multi language Word versions. Will now always utilize cover page and TOC
    New NetScaler functionality:
        NetScaler Gateway
            Global Settings
            Virtual Servers settings and policies
            Policies Session/Traffic
	    NetScaler administration users and groups
        NetScaler Authentication
	        Policies LDAP / Radius
            Actions Local / RADIUS
            Action LDAP more configuration reported and changed table layout
        NetScaler Networking
            Channels
            ACL
        NetScaler Cache redirection
    Bugfixes
        Naming of items with spaces and quotes fixed
        Expressions with spaces, quotes, dashes and slashed fixed
        Grammatical corrections
        Rechecked all settings like enabled/disabled or on/off and corrected when necessary
        Time zone not show correctly when in GMT+....
        A lot more small items
.Release Notes version 1
    Version 1.0 supports the following NetScaler functionality:
	NetScaler System Information
	Version / NSIP / vLAN
	NetScaler Global Settings
	NetScaler Feature and mode state
	NetScaler Networking
	IP Address / vLAN / Routing Table / DNS
	NetScaler Authentication
	Local / LDAP
	NetScaler Traffic Domain
	Assigned Content Switch / Load Balancer / Service  / Server
	NetScaler Monitoring
	NetScaler Certificate
	NetScaler Content Switches
	Assigned Load Balancer / Service  / Server
	NetScaler Load Balancer
	Assigned Service  / Server
	NetScaler Service
	Assigned Server / monitor
	NetScaler Service Group
	Assigned Server / monitor
	NetScaler Server
	NetScaler Custom Monitor
	NetScaler Policy
	NetScaler Action
	NetScaler Profile
#>
#endregion changelog

#region script template
#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Word") ]

Param(
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Switch]$PDF=$False,

	[parameter(Mandatory=$False )] 
	[Switch]$AddDateTime=$False,
	
    [parameter(Mandatory=$true )]
    [string] $NSIP,
    
    [parameter(Mandatory=$false ) ]
    [PSCredential] $Credential = (Get-Credential -Message 'Enter NetScaler credentials'),
	
	## EXPERIMENTAL: Require SSL/TLS, e.g. https://. This requires the client to trust to the NetScaler's certificate.
    [parameter(Mandatory=$false )]
	[System.Management.Automation.SwitchParameter] $UseNSSSL,
    
	[parameter(Mandatory=$False)] 
	[string]$Folder="",
	
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$SmtpServer="",

	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[int]$SmtpPort=25,

	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[switch]$UseSSL=$False,

	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$From="",

	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$To="",

	[parameter(Mandatory=$False)] 
	[Switch]$Dev=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("SI")]
	[Switch]$ScriptInfo=$False
	
	)

#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#Created on June 1, 2016

Set-StrictMode -Version 2

#force -verbose on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$SaveEAPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'
#recommended by webster
#$Error.Clear()

#updated by Webster 23-Apr-2016
If($Null -eq $PDF)
{
	$PDF = $False
}
If($Null -eq $MSWord)
{
	$MSWord = $False
}
If($Null -eq $AddDateTime)
{
	$AddDateTime = $False
}
If($Null -eq $Folder)
{
	$Folder = ""
}
If($Null -eq $SmtpServer)
{
	$SmtpServer = ""
}
If($Null -eq $SmtpPort)
{
	$SmtpPort = 25
}
If($Null -eq $UseSSL)
{
	$UseSSL = $False
}
If($Null -eq $From)
{
	$From = ""
}
If($Null -eq $To)
{
	$To = ""
}
If($Null -eq $Dev)
{
	$Dev = $False
}
If($Null -eq $ScriptInfo)
{
	$ScriptInfo = $False
}

If(!(Test-Path Variable:PDF))
{
	$PDF = $False
}
If(!(Test-Path Variable:MSWord))
{
	$MSWord = $False
}
If(!(Test-Path Variable:AddDateTime))
{
	$AddDateTime = $False
}
If(!(Test-Path Variable:Folder))
{
	$Folder = ""
}
If(!(Test-Path Variable:SmtpServer))
{
	$SmtpServer = ""
}
If(!(Test-Path Variable:SmtpPort))
{
	$SmtpPort = 25
}
If(!(Test-Path Variable:UseSSL))
{
	$UseSSL = $False
}
If(!(Test-Path Variable:From))
{
	$From = ""
}
If(!(Test-Path Variable:To))
{
	$To = ""
}
If(!(Test-Path Variable:Dev))
{
	$Dev = $False
}
If(!(Test-Path Variable:ScriptInfo))
{
	$ScriptInfo = $False
}

If($Dev)
{
	$Error.Clear()
	$Script:DevErrorFile = "$($pwd.Path)\NSInventoryScriptErrors_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
}

If($Null -eq $MSWord)
{
	If($PDF)
	{
		$MSWord = $False
	}
	Else
	{
		$MSWord = $True
	}
}

If($MSWord -eq $False -and $PDF -eq $False)
{
	$MSWord = $True
}

Write-Verbose "$(Get-Date): Testing output parameters"

If($MSWord)
{
	Write-Verbose "$(Get-Date): MSWord is set"
}
ElseIf($PDF)
{
	Write-Verbose "$(Get-Date): PDF is set"
}
Else
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Verbose "$(Get-Date): Unable to determine output parameter"
	If($Null -eq $MSWord)
	{
		Write-Verbose "$(Get-Date): MSWord is Null"
	}
	ElseIf($Null -eq $PDF)
	{
		Write-Verbose "$(Get-Date): PDF is Null"
	}
	Else
	{
		Write-Verbose "$(Get-Date): MSWord is $($MSWord)"
		Write-Verbose "$(Get-Date): PDF is $($PDF)"
	}
	Write-Error "Unable to determine output parameter.  Script cannot continue"
	Exit
}

If($Folder -ne "")
{
	Write-Verbose "$(Get-Date): Testing folder path"
	#does it exist
	If(Test-Path $Folder -EA 0)
	{
		#it exists, now check to see if it is a folder and not a file
		If(Test-Path $Folder -pathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Verbose "$(Get-Date): Folder path $Folder exists and is a folder"
		}
		Else
		{
			#it exists but it is a file not a folder
			Write-Error "Folder $Folder is a file, not a folder.  Script cannot continue"
			Exit
		}
	}
	Else
	{
		#does not exist
		Write-Error "Folder $Folder does not exist.  Script cannot continue"
		Exit
	}
}

#region initialize variables for word html and text
[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

If($MSWord -or $PDF)
{
	#try and fix the issue with the $CompanyName variable
	$Script:CoName = $CompanyName
	Write-Verbose "$(Get-Date): CoName is $($Script:CoName)"
	
	#the following values were attained from 
	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight = 2
	[long]$wdColorGray15 = 14277081
	[long]$wdColorGray05 = 15987699 
	[int]$wdMove = 0
	[int]$wdSeekMainDocument = 0
	[int]$wdSeekPrimaryFooter = 4
	[int]$wdStory = 6
	[long]$wdColorRed = 255
	[int]$wdColorBlack = 0
	[int]$wdWord2007 = 12
	[int]$wdWord2010 = 14
	[int]$wdWord2013 = 15
	[int]$wdWord2016 = 16
	[int]$wdFormatDocumentDefault = 16
	[int]$wdFormatPDF = 17
	#http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/01/how-can-i-right-align-a-single-column-in-a-word-table.aspx
	#http://msdn.microsoft.com/en-us/library/office/ff835817%28v=office.15%29.aspx
	[int]$wdAlignParagraphLeft = 0
	[int]$wdAlignParagraphCenter = 1
	[int]$wdAlignParagraphRight = 2
	#http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
	[int]$wdCellAlignVerticalTop = 0
	[int]$wdCellAlignVerticalCenter = 1
	[int]$wdCellAlignVerticalBottom = 2
	#http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
	[int]$wdAutoFitFixed = 0
	[int]$wdAutoFitContent = 1
	[int]$wdAutoFitWindow = 2
	#http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
	[int]$wdAdjustNone = 0
	[int]$wdAdjustProportional = 1
	[int]$wdAdjustFirstColumn = 2
	[int]$wdAdjustSameWidth = 3

	[int]$PointsPerTabStop = 36
	[int]$Indent0TabStops = 0 * $PointsPerTabStop
	[int]$Indent1TabStops = 1 * $PointsPerTabStop
	[int]$Indent2TabStops = 2 * $PointsPerTabStop
	[int]$Indent3TabStops = 3 * $PointsPerTabStop
	[int]$Indent4TabStops = 4 * $PointsPerTabStop

	# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
	[int]$wdStyleHeading1 = -2
	[int]$wdStyleHeading2 = -3
	[int]$wdStyleHeading3 = -4
	[int]$wdStyleHeading4 = -5
	[int]$wdStyleNoSpacing = -158
	[int]$wdTableGrid = -155

	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/org/codehaus/groovy/scriptom/tlb/office/word/WdLineStyle.html
	[int]$wdLineStyleNone = 0
	[int]$wdLineStyleSingle = 1

	[int]$wdHeadingFormatTrue = -1
	[int]$wdHeadingFormatFalse = 0 
}
#endregion

#region email function
Function SendEmail
{
	Param([string]$Attachments)
	Write-Verbose "$(Get-Date): Prepare to email"
	
	$emailAttachment = $Attachments
	$emailSubject = $Script:Title
	$emailBody = @"
Hello, <br />
<br />
$Script:Title is attached.
"@ 

	If($Dev)
	{
		Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
	}

	$error.Clear()

	If($UseSSL)
	{
		Write-Verbose "$(Get-Date): Trying to send email using current user's credentials with SSL"
		Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
		-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
		-UseSSL *>$Null
	}
	Else
	{
		Write-Verbose  "$(Get-Date): Trying to send email using current user's credentials without SSL"
		Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
		-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To *>$Null
	}

	$e = $error[0]

	If($e.Exception.ToString().Contains("5.7.57"))
	{
		#The server response was: 5.7.57 SMTP; Client was not authenticated to send anonymous mail during MAIL FROM
		Write-Verbose "$(Get-Date): Current user's credentials failed. Ask for usable credentials."

		If($Dev)
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}

		$error.Clear()

		$emailCredentials = Get-Credential -Message "Enter the email account and password to send email"

		If($UseSSL)
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL -credential $emailCredentials *>$Null 
		}
		Else
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-credential $emailCredentials *>$Null 
		}

		$e = $error[0]

		If($? -and $Null -eq $e)
		{
			Write-Verbose "$(Get-Date): Email successfully sent using new credentials"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Email was not sent:"
			Write-Warning "$(Get-Date): Exception: $e.Exception" 
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Email was not sent:"
		Write-Warning "$(Get-Date): Exception: $e.Exception" 
	}
}
#endregion

#region word specific functions
Function SetWordHashTable
{
	Param([string]$CultureCode)

	#optimized by Michael B. Smith
	
	# DE and FR translations for Word 2010 by Vladimir Radojevic
	# Vladimir.Radojevic@Commerzreal.com

	# DA translations for Word 2010 by Thomas Daugaard
	# Citrix Infrastructure Specialist at edgemo A/S

	# CA translations by Javier Sanchez 
	# CEO & Founder 101 Consulting

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish

	[string]$toc = $(
		Switch ($CultureCode)
		{
			'ca-'	{ 'Taula automática 2' ; Break}
			'da-'	{ 'Automatisk tabel 2' ; Break}
			'de-'	{ 'Automatische Tabelle 2' ; Break}
			'en-'	{ 'Automatic Table 2' ; Break}
			'es-'	{ 'Tabla automática 2' ; Break}
			'fi-'	{ 'Automaattinen taulukko 2' ; Break}
			'fr-'	{ 'Sommaire Automatique 2' ; Break}
			'nb-'	{ 'Automatisk tabell 2' ; Break}
			'nl-'	{ 'Automatische inhoudsopgave 2' ; Break}
			'pt-'	{ 'Sumário Automático 2' ; Break}
			'sv-'	{ 'Automatisk innehållsförteckning2' ; Break}
		}
	)

	$Script:myHash                      = @{}
	$Script:myHash.Word_TableOfContents = $toc
	$Script:myHash.Word_NoSpacing       = $wdStyleNoSpacing
	$Script:myHash.Word_Heading1        = $wdStyleheading1
	$Script:myHash.Word_Heading2        = $wdStyleheading2
	$Script:myHash.Word_Heading3        = $wdStyleheading3
	$Script:myHash.Word_Heading4        = $wdStyleheading4
	$Script:myHash.Word_TableGrid       = $wdTableGrid
}

Function GetCulture
{
	Param([int]$WordValue)
	
	#codes obtained from http://support.microsoft.com/kb/221435
	#http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
	$CatalanArray = 1027
	$DanishArray = 1030
	$DutchArray = 2067, 1043
	$EnglishArray = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
	$FinnishArray = 1035
	$FrenchArray = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
	$GermanArray = 1031, 3079, 5127, 4103, 2055
	$NorwegianArray = 1044, 2068
	$PortugueseArray = 1046, 2070
	$SpanishArray = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
	$SwedishArray = 1053, 2077

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish

	Switch ($WordValue)
	{
		{$CatalanArray -contains $_} {$CultureCode = "ca-"; Break}
		{$DanishArray -contains $_} {$CultureCode = "da-"; Break}
		{$DutchArray -contains $_} {$CultureCode = "nl-"; Break}
		{$EnglishArray -contains $_} {$CultureCode = "en-"; Break}
		{$FinnishArray -contains $_} {$CultureCode = "fi-"; Break}
		{$FrenchArray -contains $_} {$CultureCode = "fr-"; Break}
		{$GermanArray -contains $_} {$CultureCode = "de-"; Break}
		{$NorwegianArray -contains $_} {$CultureCode = "nb-"; Break}
		{$PortugueseArray -contains $_} {$CultureCode = "pt-"; Break}
		{$SpanishArray -contains $_} {$CultureCode = "es-"; Break}
		{$SwedishArray -contains $_} {$CultureCode = "sv-"; Break}
		Default {$CultureCode = "en-"}
	}
	
	Return $CultureCode
}

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)
	
	$xArray = ""
	
	Switch ($CultureCode)
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "Diplomàtic", "Exposició",
					"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "Quadrícula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevægElse", "Brusen", "Facet", "Filigran", 
					"Gitter", "Integral", "Ion (lys)", "Ion (mørk)", 
					"Retro", "Semafor", "Sidelinje", "Stribet", 
					"Udsnit (lys)", "Udsnit (mørk)", "Visningsmaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mørk)", "Ion (mørk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"Nålestribet", "Årlig", "Avispapir", "Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Bewegung", "Facette", "Filigran", 
					"Gebändert", "Integral", "Ion (dunkel)", "Ion (hell)", 
					"Pfiff", "Randlinie", "Raster", "Rückblick", 
					"Segment (dunkel)", "Segment (hell)", "Semaphor", 
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
					"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
					"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
					"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
					"Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
					"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
					"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Con bandas", "Cortar (oscuro)", "Cuadrícula", 
					"Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)", 
					"Ion (oscuro)", "Línea lateral", "Movimiento", "Retrospectiva", 
					"Semáforo", "Slice (luz)", "Vista principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
					"Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periódico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kuiskaus", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
					"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
					"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
					"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("À bandes", "Austin", "Facette", "Filigrane", 
					"Guide", "Intégrale", "Ion (clair)", "Ion (foncé)", 
					"Lignes latérales", "Quadrillage", "Rétrospective", "Secteur (clair)", 
					"Secteur (foncé)", "Sémaphore", "ViewMaster", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annuel", "Austère", "Austin", 
					"Blocs empilés", "Classique", "Contraste", "Emplacements de bureau", 
					"Exposition", "Guide", "Ligne latérale", "Moderne", 
					"Mosaïques", "Mots croisés", "Papier journal", "Perspective",
					"Quadrillage", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
					"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
					"Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
					"Integraal", "Ion (donker)", "Ion (licht)", "Raster",
					"Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
					"Terugblik", "Terzijde", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
					"Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
					"Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
					"Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana", 
					"Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
					"Retrospectiva", "Semáforo")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeça", "Transcend")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
					"RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
					"Övergående")
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
						"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
						"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
						"Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
				}
	}
	
	If($xArray -contains $xCP)
	{
		Return $True
	}
	Else
	{
		Return $False
	}
}

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
		Exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}) -ne $Null
	If($wordrunning)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`tPlease close all instances of Microsoft Word before running this report.`n`n"
		Exit
	}
}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-LocalRegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-LocalRegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

Function _SetDocumentProperty 
{
	#jeff hicks
	Param([object]$Properties,[string]$Name,[string]$Value)
	#get the property object
	$prop = $properties | ForEach { 
		$propname=$_.GetType().InvokeMember("Name","GetProperty",$Null,$_,$Null)
		If($propname -eq $Name) 
		{
			Return $_
		}
	} #ForEach

	#set the value
	$Prop.GetType().InvokeMember("Value","SetProperty",$Null,$prop,$Value)
}

Function FindWordDocumentEnd
{
	#return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

Function SetupWord
{
	Write-Verbose "$(Get-Date): Setting up Word"
    
	# Setup word for output
	Write-Verbose "$(Get-Date): Create Word comObject."
	$Script:Word = New-Object -comobject "Word.Application" -EA 0 4>$Null
	
	If(!$? -or $Null -eq $Script:Word)
	{
		Write-Warning "The Word object could not be created.  You may need to repair your Word installation."
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tThe Word object could not be created.  You may need to repair your Word installation.`n`n`t`tScript cannot continue.`n`n"
		Exit
	}

	Write-Verbose "$(Get-Date): Determine Word language value"
	If( ( validStateProp $Script:Word Language Value__ ) )
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language.Value__
	}
	Else
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language
	}

	If(!($Script:WordLanguageValue -gt -1))
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tUnable to determine the Word language value.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}
	Write-Verbose "$(Get-Date): Word language value is $($Script:WordLanguageValue)"
	
	$Script:WordCultureCode = GetCulture $Script:WordLanguageValue
	
	SetWordHashTable $Script:WordCultureCode
	
	[int]$Script:WordVersion = [int]$Script:Word.Version
	If($Script:WordVersion -eq $wdWord2016)
	{
		$Script:WordProduct = "Word 2016"
	}
	ElseIf($Script:WordVersion -eq $wdWord2013)
	{
		$Script:WordProduct = "Word 2013"
	}
	ElseIf($Script:WordVersion -eq $wdWord2010)
	{
		$Script:WordProduct = "Word 2010"
	}
	ElseIf($Script:WordVersion -eq $wdWord2007)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tMicrosoft Word 2007 is no longer supported.`n`n`t`tScript will end.`n`n"
		AbortScript
	}
	Else
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tYou are running an untested or unsupported version of Microsoft Word.`n`n`t`tScript will end.`n`n`t`tPlease send info on your version of Word to webster@carlwebster.com`n`n"
		AbortScript
	}

	#only validate CompanyName if the field is blank
	If([String]::IsNullOrEmpty($Script:CoName))
	{
		Write-Verbose "$(Get-Date): Company name is blank.  Retrieve company name from registry."
		$TmpName = ValidateCompanyName
		
		If([String]::IsNullOrEmpty($TmpName))
		{
			Write-Warning "`n`n`t`tCompany Name is blank so Cover Page will not show a Company Name."
			Write-Warning "`n`t`tCheck HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value."
			Write-Warning "`n`t`tYou may want to use the -CompanyName parameter if you need a Company Name on the cover page.`n`n"
		}
		Else
		{
			$Script:CoName = $TmpName
			Write-Verbose "$(Get-Date): Updated company name to $($Script:CoName)"
		}
	}

	If($Script:WordCultureCode -ne "en-")
	{
		Write-Verbose "$(Get-Date): Check Default Cover Page for $($WordCultureCode)"
		[bool]$CPChanged = $False
		Switch ($Script:WordCultureCode)
		{
			'ca-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línia lateral"
						$CPChanged = $True
					}
				}

			'da-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'de-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Randlinie"
						$CPChanged = $True
					}
				}

			'es-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línea lateral"
						$CPChanged = $True
					}
				}

			'fi-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sivussa"
						$CPChanged = $True
					}
				}

			'fr-'	{
					If($CoverPage -eq "Sideline")
					{
						If($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
						{
							$CoverPage = "Lignes latérales"
							$CPChanged = $True
						}
						Else
						{
							$CoverPage = "Ligne latérale"
							$CPChanged = $True
						}
					}
				}

			'nb-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'nl-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Terzijde"
						$CPChanged = $True
					}
				}

			'pt-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Linha Lateral"
						$CPChanged = $True
					}
				}

			'sv-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidlinje"
						$CPChanged = $True
					}
				}
		}

		If($CPChanged)
		{
			Write-Verbose "$(Get-Date): Changed Default Cover Page from Sideline to $($CoverPage)"
		}
	}

	Write-Verbose "$(Get-Date): Validate cover page $($CoverPage) for culture code $($Script:WordCultureCode)"
	[bool]$ValidCP = $False
	
	$ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode
	
	If(!$ValidCP)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Verbose "$(Get-Date): Word language value $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date): Culture code $($Script:WordCultureCode)"
		Write-Error "`n`n`t`tFor $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	ShowScriptOptions

	$Script:Word.Visible = $False

	#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
	#using Jeff's Demo-WordReport.ps1 file for examples
	Write-Verbose "$(Get-Date): Load Word Templates"

	[bool]$Script:CoverPagesExist = $False
	[bool]$BuildingBlocksExist = $False

	$Script:Word.Templates.LoadBuildingBlocks()
	#word 2010/2013
	$BuildingBlocksCollection = $Script:Word.Templates | Where {$_.name -eq "Built-In Building Blocks.dotx"}

	Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
	$part = $Null

	$BuildingBlocksCollection | 
	ForEach{
		If($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) 
		{
			$BuildingBlocks = $_
		}
	}        

	If($Null -ne $BuildingBlocks)
	{
		$BuildingBlocksExist = $True

		Try 
		{
			$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
		}

		Catch
		{
			$part = $Null
		}

		If($Null -ne $part)
		{
			$Script:CoverPagesExist = $True
		}
	}

	If(!$Script:CoverPagesExist)
	{
		Write-Verbose "$(Get-Date): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "This report will not have a Cover Page."
	}

	Write-Verbose "$(Get-Date): Create empty word doc"
	$Script:Doc = $Script:Word.Documents.Add()
	If($Null -eq $Script:Doc)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn empty Word document could not be created.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If($Null -eq $Script:Selection)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn unknown error happened selecting the entire Word document for default formatting options.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
	#36 = .50"
	$Script:Word.ActiveDocument.DefaultTabStop = 36

	#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
	Write-Verbose "$(Get-Date): Disable grammar and spell checking"
	#bug reported 1-Apr-2014 by Tim Mangan
	#save current options first before turning them off
	$Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
	$Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
	$Script:Word.Options.CheckGrammarAsYouType = $False
	$Script:Word.Options.CheckSpellingAsYouType = $False

	If($BuildingBlocksExist)
	{
		#insert new page, getting ready for table of contents
		Write-Verbose "$(Get-Date): Insert new page, getting ready for table of contents"
		$part.Insert($Script:Selection.Range,$True) | Out-Null
		$Script:Selection.InsertNewPage()

		#table of contents
		Write-Verbose "$(Get-Date): Table of Contents - $($Script:MyHash.Word_TableOfContents)"
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($Script:MyHash.Word_TableOfContents)
		If($Null -eq $toc)
		{
			Write-Verbose "$(Get-Date): "
			Write-Verbose "$(Get-Date): Table of Content - $($Script:MyHash.Word_TableOfContents) could not be retrieved."
			Write-Warning "This report will not have a Table of Contents."
		}
		Else
		{
			$toc.insert($Script:Selection.Range,$True) | Out-Null
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Table of Contents are not installed."
		Write-Warning "Table of Contents are not installed so this report will not have a Table of Contents."
	}

	#set the footer
	Write-Verbose "$(Get-Date): Set the footer"
	[string]$footertext = "Report created by $username"

	#get the footer
	Write-Verbose "$(Get-Date): Get the footer and format font"
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
	#get the footer and format font
	$footers = $Script:Doc.Sections.Last.Footers
	ForEach($footer in $footers) 
	{
		If($footer.exists) 
		{
			$footer.range.Font.name = "Calibri"
			$footer.range.Font.size = 8
			$footer.range.Font.Italic = $True
			$footer.range.Font.Bold = $True
		}
	} #end ForEach
	Write-Verbose "$(Get-Date): Footer text"
	$Script:Selection.HeaderFooter.Range.Text = $footerText

	#add page numbering
	Write-Verbose "$(Get-Date): Add page numbering"
	$Script:Selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

	FindWordDocumentEnd
	Write-Verbose "$(Get-Date):"
	#end of Jeff Hicks 
}

Function UpdateDocumentProperties
{
	Param([string]$AbstractTitle, [string]$SubjectTitle)
	#Update document properties
	If($MSWORD -or $PDF)
	{
		If($Script:CoverPagesExist)
		{
			Write-Verbose "$(Get-Date): Set Cover Page Properties"
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Company" $Script:CoName
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Title" $Script:title
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Author" $username

			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Subject" $SubjectTitle

			#Get the Coverpage XML part
			$cp = $Script:Doc.CustomXMLParts | Where {$_.NamespaceURI -match "coverPageProps$"}

			#get the abstract XML part
			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}

			#set the text
			If([String]::IsNullOrEmpty($Script:CoName))
			{
				[string]$abstract = $AbstractTitle
			}
			Else
			{
				[string]$abstract = "$($AbstractTitle) for $($Script:CoName)"
			}

			$ab.Text = $abstract

			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "PublishDate"}
			#set the text
			[string]$abstract = (Get-Date -Format d).ToString()
			$ab.Text = $abstract

			Write-Verbose "$(Get-Date): Update the Table of Contents"
			#update the Table of Contents
			$Script:Doc.TablesOfContents.item(1).Update()
			$cp = $Null
			$ab = $Null
			$abstract = $Null
		}
	}
}
#endregion

#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}

#region word output functions
Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
#updated 27-Mar-2014 to include font name, font size, italics and bold options
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName=$Null,
	[int]$fontSize=0,
	[bool]$italics=$False,
	[bool]$boldface=$False,
	[Switch]$nonewline)
	
	#Build output style
	[string]$output = ""
	Switch ($style)
	{
		0 {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
		1 {$Script:Selection.Style = $Script:MyHash.Word_Heading1; Break}
		2 {$Script:Selection.Style = $Script:MyHash.Word_Heading2; Break}
		3 {$Script:Selection.Style = $Script:MyHash.Word_Heading3; Break}
		4 {$Script:Selection.Style = $Script:MyHash.Word_Heading4; Break}
		Default {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Script:Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Script:Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Script:Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Script:Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Script:Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Script:Selection.TypeParagraph()
	}
}
#endregion

#region Iain's Word table functions

<#
.Synopsis
	Add a table to a Microsoft Word document
.DESCRIPTION
	This function adds a table to a Microsoft Word document from either an array of
	Hashtables or an array of PSCustomObjects.

	Using this function is quicker than setting each table cell individually but can
	only utilise the built-in MS Word table autoformats. Individual tables cells can
	be altered after the table has been appended to the document (a table reference
	is returned).
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. Column headers will display the key names as defined.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -List

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. No column headers will be added, in a ListView format.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray

	This example adds table to the MS Word document, utilising all note property names
	the array of PSCustomObjects. Column headers will display the note property names.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -Columns FirstName,LastName,EmailAddress

	This example adds a table to the MS Word document, but only using the specified
	key names: FirstName, LastName and EmailAddress. If other keys are present in the
	array of Hashtables they will be ignored.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray -Columns FirstName,LastName,EmailAddress -Headers "First Name","Last Name","Email Address"

	This example adds a table to the MS Word document, but only using the specified
	PSCustomObject note properties: FirstName, LastName and EmailAddress. If other note
	properties are present in the array of PSCustomObjects they will be ignored. The
	display names for each specified column header has been overridden to display a
	custom header. Note: the order of the header names must match the specified columns.
#>

Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Columns = $Null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Headers = $Null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines,
		[Switch] $NoInternalGridLines,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$True)] [int] $Format = 0
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Null -eq $Columns) -and ($Null -ne $Headers)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $Null;
		}
		ElseIf(($Null -ne $Columns) -and ($Null -ne $Headers)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end ElseIf
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
		[System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Null -eq $Columns) 
				{
					## Build the available columns from all available PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end foreach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Null -eq $Columns) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{ 
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end foreach

				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $True);
			$ConvertToTableArguments.Add("ApplyShading", $True);
			$ConvertToTableArguments.Add("ApplyFont", $True);
			$ConvertToTableArguments.Add("ApplyColor", $True);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $True); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $True);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $True);
			$ConvertToTableArguments.Add("ApplyLastColumn", $True);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$Null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$Null,                                          # Modifiers
			$Null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		If(!$List)
		{
			#the next line causes the heading row to flow across page breaks
			$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;
		}

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}
		If($NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleNone;
		}
		If($NoInternalGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}

<#
.Synopsis
	Sets the format of one or more Word table cells
.DESCRIPTION
	This function sets the format of one or more table cells, either from a collection
	of Word COM object cell references, an individual Word COM object cell reference or
	a hashtable containing Row and Column information.

	The font name, font size, bold, italic , underline and shading values can be used.
.EXAMPLE
	SetWordCellFormat -Hashtable $Coordinates -Table $TableReference -Bold

	This example sets all text to bold that is contained within the $TableReference
	Word table, using an array of hashtables. Each hashtable contain a pair of co-
	ordinates that is used to select the required cells. Note: the hashtable must
	contain the .Row and .Column key names. For example:
	@ { Row = 7; Column = 3 } to set the cell at row 7 and column 3 to bold.
.EXAMPLE
	$RowCollection = $Table.Rows.First.Cells
	SetWordCellFormat -Collection $RowCollection -Bold -Size 10

	This example sets all text to size 8 and bold for all cells that are contained
	within the first row of the table.
	Note: the $Table.Rows.First.Cells returns a collection of Word COM cells objects
	that are in the first table row.
.EXAMPLE
	$ColumnCollection = $Table.Columns.Item(2).Cells
	SetWordCellFormat -Collection $ColumnCollection -BackgroundColor 255

	This example sets the background (shading) of all cells in the table's second
	column to red.
	Note: the $Table.Columns.Item(2).Cells returns a collection of Word COM cells objects
	that are in the table's second column.
.EXAMPLE
	SetWordCellFormat -Cell $Table.Cell(17,3) -Font "Tahoma" -Color 16711680

	This example sets the font to Tahoma and the text color to blue for the cell located
	in the table's 17th row and 3rd column.
	Note: the $Table.Cell(17,3) returns a single Word COM cells object.
#>

Function SetWordCellFormat 
{
	[CmdletBinding(DefaultParameterSetName='Collection')]
	Param (
		# Word COM object cell collection reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='Collection', Position=0)] [ValidateNotNullOrEmpty()] $Collection,
		# Word COM object individual cell reference
		[Parameter(Mandatory=$true, ParameterSetName='Cell', Position=0)] [ValidateNotNullOrEmpty()] $Cell,
		# Hashtable of cell co-ordinates
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
		# Word COM object table reference
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=1)] [ValidateNotNullOrEmpty()] $Table,
		# Font name
		[Parameter()] [AllowNull()] [string] $Font = $Null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $Null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] $BackgroundColor = $Null,
		# Force solid background color
		[Switch] $Solid,
		[Switch] $Bold,
		[Switch] $Italic,
		[Switch] $Underline
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
	}

	Process 
	{
		Switch ($PSCmdlet.ParameterSetName) 
		{
			'Collection' {
				ForEach($Cell in $Collection) 
				{
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end foreach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $true; }
				If($Italic) { $Cell.Range.Font.Italic = $true; }
				If($Underline) { $Cell.Range.Font.Underline = 1; }
				If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
				If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
				If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
				If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
			} # end Cell
			'Hashtable' 
			{
				ForEach($Coordinate in $Coordinates) 
				{
					$Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column);
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				}
			} # end Hashtable
		} # end switch
	} # end process
}

<#
.Synopsis
	Sets alternate row colors in a Word table
.DESCRIPTION
	This function sets the format of alternate rows within a Word table using the
	specified $BackgroundColor. This function is expensive (in performance terms) as
	it recursively sets the format on alternate rows. It would be better to pick one
	of the predefined table formats (if one exists)? Obviously the more rows, the
	longer it takes :'(

	Note: this function is called by the AddWordTable function if an alternate row
	format is specified.
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 255

	This example sets every-other table (starting with the first) row and sets the
	background color to red (wdColorRed).
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 39423 -Seed Second

	This example sets every other table (starting with the second) row and sets the
	background color to light orange (weColorLightOrange).
#>

Function SetWordTableAlternateRowColor 
{
	[CmdletBinding()]
	Param (
		# Word COM object table reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)] [ValidateNotNullOrEmpty()] $Table,
		# Alternate row background color
		[Parameter(Mandatory=$true, Position=1)] [ValidateNotNull()] [int] $BackgroundColor,
		# Alternate row starting seed
		[Parameter(ValueFromPipelineByPropertyName=$true, Position=2)] [ValidateSet('First','Second')] [string] $Seed = 'First'
	)

	Process 
	{
		$StartDateTime = Get-Date;
		Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime);

		## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
		If($Seed.ToLower() -eq 'second') 
		{ 
			$StartRowIndex = 2; 
		}
		Else 
		{ 
			$StartRowIndex = 1; 
		}

		For($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2) 
		{ 
			$Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor;
		}

		## I've put verbose calls in here we can see how expensive this functionality actually is.
		$EndDateTime = Get-Date;
		$ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime;
		Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds);
	}
}
#endregion

#region general script functions
Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If((gm -Name $topLevel -InputObject $object))
		{
			If((gm -Name $secondLevel -InputObject $object.$topLevel))
			{
				Return $True
			}
		}
	}
	Return $False
}

Function AbortScript
{
	If($MSWord -or $PDF)
	{
		$Script:Word.quit()
		Write-Verbose "$(Get-Date): System Cleanup"
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
		If(Test-Path variable:global:word)
		{
			Remove-Variable -Name word -Scope Global
		}
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	$ErrorActionPreference = $SaveEAPreference
	Exit
}

Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): AddDateTime     : $($AddDateTime)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Company Name    : $($Script:CoName)"
	}
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Cover Page      : $($CoverPage)"
	}
	Write-Verbose "$(Get-Date): Dev             : $($Dev)"
	If($Dev)
	{
		Write-Verbose "$(Get-Date): DevErrorFile    : $($Script:DevErrorFile)"
	}
	Write-Verbose "$(Get-Date): Filename1       : $($Script:filename1)"
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Filename2       : $($Script:filename2)"
	}
	Write-Verbose "$(Get-Date): Folder          : $($Folder)"
	Write-Verbose "$(Get-Date): From            : $($From)"
	Write-Verbose "$(Get-Date): NSIP            : $($NSIP)"
	Write-Verbose "$(Get-Date): Save As PDF     : $($PDF)"
	Write-Verbose "$(Get-Date): Save As WORD    : $($MSWORD)"
	Write-Verbose "$(Get-Date): ScriptInfo      : $($ScriptInfo)"
	Write-Verbose "$(Get-Date): Smtp Port       : $($SmtpPort)"
	Write-Verbose "$(Get-Date): Smtp Server     : $($SmtpServer)"
	Write-Verbose "$(Get-Date): Title           : $($Script:Title)"
	Write-Verbose "$(Get-Date): To              : $($To)"
	Write-Verbose "$(Get-Date): Use NS SSL      : $($UseNSSSL)"
	Write-Verbose "$(Get-Date): Use SSL         : $($UseSSL)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): User Name       : $($UserName)"
	}
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): OS Detected     : $($Script:RunningOS)"
	Write-Verbose "$(Get-Date): PoSH version    : $($Host.Version)"
	Write-Verbose "$(Get-Date): PSCulture       : $($PSCulture)"
	Write-Verbose "$(Get-Date): PSUICulture     : $($PSUICulture)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Word language   : $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date): Word version    : $($Script:WordProduct)"
	}
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Script start    : $($Script:StartTime)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
}

Function SaveandCloseDocumentandShutdownWord
{
	#bug fix 1-Apr-2014
	#reset Grammar and Spelling options back to their original settings
	$Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
	$Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

	Write-Verbose "$(Get-Date): Save and Close document and Shutdown Word"
	If($Script:WordVersion -eq $wdWord2010)
	{
		#the $saveFormat below passes StrictMode 2
		#I found this at the following two links
		#http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
		#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Verbose "$(Get-Date): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:FileName1, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$saveFormat)
		}
	}
	ElseIf($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
	{
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Verbose "$(Get-Date): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$Script:Doc.SaveAs2([REF]$Script:FileName1, [ref]$wdFormatDocumentDefault)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$wdFormatPDF)
		}
	}

	Write-Verbose "$(Get-Date): Closing Word"
	$Script:Doc.Close()
	$Script:Word.Quit()
	If($PDF)
	{
		[int]$cnt = 0
		While(Test-Path $Script:FileName1)
		{
			$cnt++
			If($cnt -gt 1)
			{
				Write-Verbose "$(Get-Date): Waiting another 10 seconds to allow Word to fully close (try # $($cnt))"
				Start-Sleep -Seconds 10
				$Script:Word.Quit()
				If($cnt -gt 2)
				{
					#kill the winword process

					#find out our session (usually "1" except on TS/RDC or Citrix)
					$SessionID = (Get-Process -PID $PID).SessionId
					
					#Find out if winword is running in our session
					$wordprocess = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}).Id
					If($wordprocess -gt 0)
					{
						Write-Verbose "$(Get-Date): Attempting to stop WinWord process # $($wordprocess)"
						Stop-Process $wordprocess -EA 0
					}
				}
			}
			Write-Verbose "$(Get-Date): Attempting to delete $($Script:FileName1) since only $($Script:FileName2) is needed (try # $($cnt))"
			Remove-Item $Script:FileName1 -EA 0 4>$Null
		}
	}
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global 4>$Null
	}
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	
	#is the winword process still running? kill it

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId

	#Find out if winword is running in our session
	$wordprocess = $Null
	$wordprocess = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}).Id
	If($null -ne $wordprocess -and $wordprocess -gt 0)
	{
		Write-Verbose "$(Get-Date): WinWord process is still running. Attempting to stop WinWord process # $($wordprocess)"
		Stop-Process $wordprocess -EA 0
	}
}

Function SetFileName1andFileName2
{
	Param([string]$OutputFileName)
	
	If($Folder -eq "")
	{
		$pwdpath = $pwd.Path
	}
	Else
	{
		$pwdpath = $Folder
	}

	If($pwdpath.EndsWith("\"))
	{
		#remove the trailing \
		$pwdpath = $pwdpath.SubString(0, ($pwdpath.Length - 1))
	}

	#set $filename1 and $filename2 with no file extension
	If($AddDateTime)
	{
		[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName)"
		If($PDF)
		{
			[string]$Script:FileName2 = "$($pwdpath)\$($OutputFileName)"
		}
	}

	If($MSWord -or $PDF)
	{
		CheckWordPreReq

		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).docx"
			If($PDF)
			{
				[string]$Script:FileName2 = "$($pwdpath)\$($OutputFileName).pdf"
			}
		}

		SetupWord
	}
}
#endregion

#region script end
Function ProcessScriptEnd
{
	Write-Verbose "$(Get-Date): Script has completed"
	Write-Verbose "$(Get-Date): "

	#http://poshtips.com/measuring-elapsed-time-in-powershell/
	Write-Verbose "$(Get-Date): Script started: $($Script:StartTime)"
	Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds",
		$runtime.Days,
		$runtime.Hours,
		$runtime.Minutes,
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Verbose "$(Get-Date): Elapsed time: $($Str)"

	If($Dev)
	{
		If($SmtpServer -eq "")
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
		}
		Else
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}
	}

	If($ScriptInfo)
	{
		$SIFile = "$($pwd.Path)\NSInventoryScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		Out-File -FilePath $SIFile -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Add DateTime   : $($AddDateTime)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Company Name   : $($Script:CoName)" 4>$Null		
		}
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Cover Page     : $($CoverPage)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Dev            : $($Dev)" 4>$Null
		If($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile   : $($Script:DevErrorFile)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Filename1      : $($Script:FileName1)" 4>$Null
		If($PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Filename2      : $($Script:FileName2)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Folder         : $($Folder)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "From           : $($From)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "NSIP           : $($NSIP)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As PDF    : $($PDF)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As WORD   : $($MSWORD)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script Info    : $($ScriptInfo)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Port      : $($SmtpPort)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Server    : $($SmtpServer)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Title          : $($Script:Title)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "To             : $($To)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use NS SSL     : $($UseNSSSL)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use SSL        : $($UseSSL)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "User Name      : $($UserName)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected    : $($Script:RunningOS)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version   : $($Host.Version)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSCulture      : $($PSCulture)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSUICulture    : $($PSUICulture)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word language  : $($Script:WordLanguageValue)" 4>$Null
			Out-File -FilePath $SIFile -Append -InputObject "Word version   : $($Script:WordProduct)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script start   : $($Script:StartTime)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elapsed time   : $($Str)" 4>$Null
	}

	$ErrorActionPreference = $SaveEAPreference
	[gc]::collect()
}
#endregion

#region general script functions
Function ProcessDocumentOutput
{
	If($MSWORD -or $PDF)
	{
		SaveandCloseDocumentandShutdownWord
	}

	$GotFile = $False

	If($PDF)
	{
		If(Test-Path "$($Script:FileName2)")
		{
			Write-Verbose "$(Get-Date): $($Script:FileName2) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName2)"
			Write-Error "Unable to save the output file, $($Script:FileName2)"
		}
	}
	Else
	{
		If(Test-Path "$($Script:FileName1)")
		{
			Write-Verbose "$(Get-Date): $($Script:FileName1) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName1)"
			Write-Error "Unable to save the output file, $($Script:FileName1)"
		}
	}
	
	#email output file if requested
	If($GotFile -and ![System.String]::IsNullOrEmpty( $SmtpServer ))
	{
		If($PDF)
		{
			$emailAttachment = $Script:FileName2
		}
		Else
		{
			$emailAttachment = $Script:FileName1
		}
		SendEmail $emailAttachment
	}
	[gc]::collect()
}
#endregion


#Script begins

$script:startTime = Get-Date
#endregion script template

#region file name and title name
#The function SetFileName1andFileName2 needs your script output filename
#change title for your report
[string]$Script:Title = "NetScaler Documentation $($Script:CoName)"
SetFileName1andFileName2 "NetScaler Documentation"

#endregion file name and title name

#region NetScaler Documentation Script Complete

## Barry Schiffer Use Stopwatch class to time script execution
$sw = [Diagnostics.Stopwatch]::StartNew()

$selection.InsertNewPage()

#region Nitro Functions

function Get-vNetScalerObjectList {
<#
    .SYNOPSIS
        Returns a list of objects available in a NetScaler Nitro API container.
#>
    [CmdletBinding()]
    param (
        # NetScaler Nitro API Container, i.e. nitro/v1/stat/ or nitro/v1/config/
        [Parameter(Mandatory)] [ValidateSet('Stat','Config')] [System.String] $Container
    )
    begin {
        $Container = $Container.ToLower();
    }
    process {
        if ($script:nsSession.UseSSL) { $protocol = 'https'; }
        else { $protocol = 'http'; }
        $uri = '{0}://{1}/nitro/v1/{2}/' -f $protocol, $script:nsSession.Address, $Container;
        $restResponse = InvokevNetScalerNitroMethod -Uri $Uri -Container $Container;
        $methodResponse = '{0}objects' -f $Container.ToLower();
        Write-Output $restResponse.($methodResponse).objects;
    }
} #end function Get-vNetScalerObjectList

function Get-vNetScalerObject {
<#
    .SYNOPSIS
        Returns a NetScaler Nitro API object(s) via its REST API.
#>
    [CmdletBinding()]
    param (
        # NetScaler Nitro API resource type, e.g. /nitro/v1/config/LBVSERVER
        [Parameter(Mandatory)] [Alias('Object','Type')] [System.String] $ResourceType,
        # NetScaler Nitro API resource name, e.g. /nitro/v1/config/lbvserver/MYLBVSERVER
        [Parameter()] [Alias('Name')] [System.String] $ResourceName,
        # NetScaler Nitro API optional attributes, e.g. /nitro/v1/config/lbvserver/mylbvserver?ATTRS=<attrib1>,<attrib2>
        [Parameter()] [System.String[]] $Attribute,
        # NetScaler Nitro API Container, i.e. nitro/v1/stat/ or nitro/v1/config/
        [Parameter()] [ValidateSet('Stat','Config')] [System.String] $Container = 'Config'
    )
    begin {
        $Container = $Container.ToLower();
        $ResourceType = $ResourceType.ToLower();
        $ResourceName = $ResourceName.ToLower();
    }
    process {
        if ($script:nsSession.UseSSL) { $protocol = 'https'; }
        else { $protocol = 'http'; }
        $uri = '{0}://{1}/nitro/v1/{2}/{3}' -f $protocol, $script:nsSession.Address, $Container, $ResourceType;
        if ($ResourceName) { $uri = '{0}/{1}' -f $uri, $ResourceName; }
        if ($Attribute) {
            $attrs = [System.String]::Join(',', $Attribute);
            $uri = '{0}?attrs={1}' -f $uri, $attrs;
        }
        $uri = [System.Uri]::EscapeUriString($uri.ToLower());
        $restResponse = InvokevNetScalerNitroMethod -Uri $Uri -Container $Container;
        if ($null -ne $restResponse.($ResourceType)) { Write-Output $restResponse.($ResourceType); }
        else { Write-Output $restResponse }
    }
} #end function Get-vNetScalerObject

function InvokevNetScalerNitroMethod {
<#
    .SYNOPSIS
        Calls a fully qualified NetScaler Nitro API
    .NOTES
        This is an internal function and shouldn't be called directly
#>
    [CmdletBinding()]
    param (
        # NetScaler Nitro API uniform resource identifier
        [Parameter(Mandatory)] [string] $Uri,
        # NetScaler Nitro API Container, i.e. nitro/v1/stat/ or nitro/v1/config/
        [Parameter(Mandatory)] [ValidateSet('Stat','Config')] [string] $Container
    )
    begin {
        if ($script:nsSession -eq $null) { throw 'No valid NetScaler session configuration.'; }
        if ($script:nsSession.Session -eq $null -or $script:nsSession.Expiry -eq $null) { throw 'Invalid NetScaler session cookie.'; }
        if ($script:nsSession.Expiry -lt (Get-Date)) { throw 'NetScaler session has expired.'; }
    }
    process {
        $irmParameters = @{
            Uri = $Uri;
            Method = 'Get';
            WebSession = $script:nsSession.Session;
            ErrorAction = 'Stop';
        }
        Write-Output (Invoke-RestMethod @irmParameters);
    }
} #end function InvokevNetScalerNitroMethod

function Connect-vNetScalerSession {
<#
    .SYNOPSIS
        Authenticates to the NetScaler and stores a session cookie.
#>
    [CmdletBinding(DefaultParameterSetName='HTTP')]
    [OutputType([Microsoft.PowerShell.Commands.WebRequestSession])]
    param (
        # NetScaler uniform resource identifier
        [Parameter(Mandatory, ParameterSetName='HTTP')]
        [Parameter(Mandatory, ParameterSetName='HTTPS')]
        [System.String] $ComputerName,
        # NetScaler session timeout (seconds)
        [Parameter(ParameterSetName='HTTP')]
        [Parameter(ParameterSetName='HTTPS')]
        [ValidateNotNull()]
        [System.Int32] $Timeout = 3600,
        # NetScaler authentication credentials
        [Parameter(ParameterSetName='HTTP')]
        [Parameter(ParameterSetName='HTTPS')]
        [System.Management.Automation.PSCredential] $Credential = $(Get-Credential -Message "Provide NetScaler credentials for '$ComputerName'";),
        ## EXPERIMENTAL: Require SSL/TLS, e.g. https://. This requires the client to trust to the NetScaler's certificate.
        [Parameter(ParameterSetName='HTTPS')] [System.Management.Automation.SwitchParameter] $UseNSSSL
    )
    process {
        if ($UseNSSSL) { $protocol = 'https'; }
        else { $protocol = 'http'; }
        $script:nsSession = @{ Address = $ComputerName; UseSSL = $UseNSSSL }
        $json = '{{ "login": {{ "username": "{0}", "password": "{1}", "timeout": {2} }} }}';
        $invokeRestMethodParams = @{
            Uri = ('{0}://{1}/nitro/v1/config/login' -f $protocol, $ComputerName);
            Method = 'Post';
            Body = ($json -f $Credential.UserName, $Credential.GetNetworkCredential().Password, $Timeout);
            ContentType = 'application/json';
            SessionVariable = 'nsSessionCookie';
            ErrorAction = 'Stop';
        }
        $restResponse = Invoke-RestMethod @invokeRestMethodParams;
        ## Store the session cookie at the script scope
        $script:nsSession.Session = $nsSessionCookie;
        ## Store the session expiry
        $script:nsSession.Expiry = (Get-Date).AddSeconds($Timeout);
        ## Return the Rest Method response
        Write-Output $restResponse;
    }
} #end function Connect-vNetScalerSession

function Get-vNetScalerObjectCount {
<#
.Synopsis
    Returns an individual NetScaler Nitro API object.
#>
    [CmdletBinding()]
    param (
        # NetScaler Nitro API Object, e.g. /nitro/v1/config/NSVERSION
        [Parameter(Mandatory)] [string] $Object,
        # NetScaler Nitro API Container, i.e. nitro/v1/stat/ or nitro/v1/config/
        [Parameter(Mandatory)] [ValidateSet('Stat','Config')] [string] $Container
    )

    begin {
        ## Check session cookie
        if ($script:nsSession.Session -eq $null) { throw 'Invalid NetScaler session cookie.'; }
    }

    process {
        $uri = 'http://{0}/nitro/v1/{1}/{2}?count=yes' -f $script:nsSession.Address, $Container.ToLower(), $Object.ToLower();
        $restResponse = InvokevNetScalerNitroMethod -Uri $Uri -Container $Container;
        # $objectResponse = '{0}objects' -f $Container.ToLower();
        Write-Output $restResponse.($Object.ToLower());
    }
}

#endregion Nitro Functions

#region NetScaler Connect

## Ensure we can connect to the NetScaler appliance before we spin up Word!
## Connect to the API if there is no session cookie
## Note: repeated logons will result in 'Connection limit to cfe exceeded' errors.
if (-not (Get-Variable -Name nsSession -Scope Script -ErrorAction SilentlyContinue)) { 
    #[ref] $null = Connect-vNetScalerSession -ComputerName $nsip -Credential $Credential -UseSSL:$UseNSSSL -ErrorAction Stop;
    [ref] $null = Connect-vNetScalerSession -ComputerName $nsip -Credential $Credential -ErrorAction Stop;
}
#endregion NetScaler Connect

#region NetScaler chaptercounters
$Chapters = 33
$Chapter = 0
#endregion NetScaler chaptercounters

#region NetScaler feature state
##Getting Feature states for usage later on and performance enhancements by not running parts of the script when feature is disabled
$NSFeatures = Get-vNetScalerObject -Container config -Object nsfeature -Verbose;
If ($NSFEATURES.WL -eq "True") {$FEATWL = "Enabled"} Else {$FEATWL = "Disabled"}
If ($NSFEATURES.SP -eq "True") {$FEATSP = "Enabled"} Else {$FEATSP = "Disabled"}
If ($NSFEATURES.LB -eq "True") {$FEATLB = "Enabled"} Else {$FEATLB = "Disabled"}
If ($NSFEATURES.CS -eq "True") {$FEATCS = "Enabled"} Else {$FEATCS = "Disabled"}
If ($NSFEATURES.CR -eq "True") {$FEATCR = "Enabled"} Else {$FEATCR = "Disabled"}
If ($NSFEATURES.SC -eq "True") {$FEATSC = "Enabled"} Else {$FEATSC = "Disabled"}
If ($NSFEATURES.CMP -eq "True") {$FEATCMP = "Enabled"} Else {$FEATCMP = "Disabled"}
If ($NSFEATURES.PQ -eq "True") {$FEATPQ = "Enabled"} Else {$FEATPQ = "Disabled"}
If ($NSFEATURES.SSL -eq "True") {$FEATSSL = "Enabled"} Else {$FEATSSL = "Disabled"}
If ($NSFEATURES.GSLB -eq "True") {$FEATGSLB = "Enabled"} Else {$FEATGSLB = "Disabled"}
If ($NSFEATURES.HDSOP -eq "True") {$FEATHDSOP = "Enabled"} Else {$FEATHDOSP = "Disabled"}
If ($NSFEATURES.CF -eq "True") {$FEATCF = "Enabled"} Else {$FEATCF = "Disabled"}
If ($NSFEATURES.IC -eq "True") {$FEATIC = "Enabled"} Else {$FEATIC = "Disabled"}
If ($NSFEATURES.SSLVPN -eq "True") {$FEATSSLVPN = "Enabled"} Else {$FEATSSLVPN = "Disabled"}
If ($NSFEATURES.AAA -eq "True") {$FEATAAA = "Enabled"} Else {$FEATAAA = "Disabled"}
If ($NSFEATURES.OSPF -eq "True") {$FEATOSPF = "Enabled"} Else {$FEATOSPF = "Disabled"}
If ($NSFEATURES.RIP -eq "True") {$FEATRIP = "Enabled"} Else {$FEATRIP = "Disabled"}
If ($NSFEATURES.BGP -eq "True") {$FEATBGP = "Enabled"} Else {$FEATBGP = "Disabled"}
If ($NSFEATURES.REWRITE -eq "True") {$FEATREWRITE = "Enabled"} Else {$FEATREWRITE = "Disabled"}
If ($NSFEATURES.IPv6PT -eq "True") {$FEATIPv6PT = "Enabled"} Else {$FEATIPv6PT = "Disabled"}
If ($NSFEATURES.APPFW -eq "True") {$FEATAppFw = "Enabled"} Else {$FEATAppFw = "Disabled"}
If ($NSFEATURES.RESPONDER -eq "True") {$FEATRESPONDER = "Enabled"} Else {$FEATRESPONDER = "Disabled"}
If ($NSFEATURES.HTMLInjection -eq "True") {$FEATHTMLInjection = "Enabled"} Else {$FEATHTMLInjection = "Disabled"}
If ($NSFEATURES.PUSH -eq "True") {$FEATpush = "Enabled"} Else {$FEATpush = "Disabled"}
If ($NSFEATURES.APPFLOW -eq "True") {$FEATAppFlow = "Enabled"} Else {$FEATAppFlow = "Disabled"}
If ($NSFEATURES.CloudBridge -eq "True") {$FEATCloudBridge = "Enabled"} Else {$FEATCloudBridge = "Disabled"}
If ($NSFEATURES.ISIS -eq "True") {$FEATISIS = "Enabled"} Else {$FEATISIS = "Disabled"}
If ($NSFEATURES.CH -eq "True") {$FEATCH = "Enabled"} Else {$FEATCH = "Disabled"}
If ($NSFEATURES.APPQoE -eq "True") {$FEATAppQoE = "Enabled"} Else {$FEATAppQoE = "Disabled"}
If ($NSFEATURES.vPath -eq "True") {$FEATVpath = "Enabled"} Else {$FEATVpath = "Disabled"}
If ($NSFEATURES.contentaccelerator -eq "True") {$FEATcontentaccelerator = "Enabled"} Else {$FEATcontentaccelerator = "Disabled"}
If ($NSFEATURES.rise -eq "True") {$FEATrise = "Enabled"} Else {$FEATrise = "Disabled"}
If ($NSFEATURES.feo -eq "True") {$FEATfeo = "Enabled"} Else {$FEATfeo = "Disabled"}
If ($NSFEATURES.lsn -eq "True") {$FEATlsn = "Enabled"} Else {$FEATlsn = "Disabled"}
If ($NSFEATURES.rdpproxy -eq "True") {$FEATrdpproxy = "Enabled"} Else {$FEATrdpproxy = "Disabled"}
If ($NSFEATURES.rep -eq "True") {$FEATrep = "Enabled"} Else {$FEATrep = "Disabled"}
#endregion NetScaler feature state

#region NetScaler Version

## Get version and build
$NSVersion = Get-vNetScalerObject -Container config -Object nsversion;
$NSVersion1 = ($NSVersion.version -replace 'NetScaler', '').split()
$Version = ($NSVersion1[1] -replace ':', '')
$Build = $($NSVersion1[5] + " " + $nsversion1[6] + " " + $nsversion1[7] -replace ',', '')

## Set script test version
## WIP THIS WORKS ONLY WHEN REGIONAL SETTINGS DIGIT IS SET TO . :)
$ScriptVersion = 11.1
#endregion NetScaler Version

#region NetScaler System Information

#region Basics
WriteWordLine 1 0 "NetScaler Configuration"

$nsconfig = Get-vNetScalerObject -Container config -Object nsconfig;
$nshostname = Get-vNetScalerObject -Container config -Object nshostname;

WriteWordLine 2 0 "NetScaler Version and configuration"
WriteWordLine 0 0 " "
$Params = $null
$Params = @{
    Hashtable = @{
        Name = $NSHOSTNAME.hostname;
        Version = $Version;
        Build = $Build;
        Saveddate = $nsconfig.lastconfigsavetime;
    }
    Columns = "Name","Version","Build","Saveddate";
    Headers = "Host Name","Version","Build","Last Configuration Saved Date";
    Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
    AutoFit = $wdAutoFitContent;
}
$Table = AddWordTable @Params -NoGridLines;

FindWordDocumentEnd;
WriteWordLine 0 0 " "

WriteWordLine 2 0 "NetScaler Edition"
WriteWordLine 0 0 " "
$License = Get-vNetScalerObject -Container config -Object nslicense;
If ($license.isstandardlic -eq $True){$LIC = "Standard"}
If ($license.isenterpriselic -eq $True){$LIC = "Enterprise"}
If ($license.isplatinumlic -eq $True){$LIC = "Platinum"}

$Params = $null
$Params = @{
    Hashtable = @{
        Edition = $LIC
        SSLVPN = $License.f_sslvpn_users;
    }
    Columns = "Edition","SSLVPN";
    Headers = "NetScaler Edition","SSL VPN licenses";
    Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
    AutoFit = $wdAutoFitContent;
}
$Table = AddWordTable @Params -NoGridLines;

FindWordDocumentEnd;
WriteWordLine 0 0 " "

WriteWordLine 2 0 "NetScaler Hardware"
WriteWordLine 0 0 " "
$nshardware = Get-vNetScalerObject -Container config -Object nshardware;

## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $NSHARDWARETable = @(
    @{ Description = "Description"; Value = "Value" }
    @{ Description = "Hardware Description"; Value = $nshardware.hwdescription }
    @{ Description = "Hardware System ID"; Value = $nshardware.sysid }
    @{ Description = "Host ID"; Value = $nshardware.hostid }
    @{ Description = "Host (MAC)"; Value = $nshardware.host }
    @{ Description = "Serial Number"; Value = $nshardware.serialno }
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $NSHARDWARETable;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params -NoGridLines -List;

FindWordDocumentEnd;
$TableRange = $null
$Table = $null      
 
WriteWordLine 0 0 " "
#endregion Basics

#region NetScaler IP
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler IP"

WriteWordLine 2 0 "NetScaler Management IP Address"

$NSIP1 = Get-vNetScalerObject -Container config -Object nsip;
Foreach ($IP in $NSIP1){ ##Lists all NetScaler IPs while we only need NSIP for this one
    If ($IP.Type -eq "NSIP")
        {
        $Params = $null
        $Params = @{
            Hashtable = @{
                NSIP = $IP.ipaddress;
                Subnet = $IP.netmask;
            }
            Columns = "NSIP","Subnet";
            Headers = "NetScaler IP Address","Subnet";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
        $Table = $null
    }
 }
#endregion NetScaler IP

#region NetScaler High Availability

WriteWordLine 2 0 "NetScaler High Availability"
WriteWordLine 0 0 " "
$HANodes = Get-vNetScalerObject -Container config -Object hanode;$

## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $HAH = @();

foreach ($HANODE in $HANodes) {
    $HANODENAME = $HANODE.name
    #Name attribute will not be returned for secondary appliance
    if ([string]::IsNullOrWhitespace($HANODENAME)){
      $HANODENAME = ""
    } 
    $HAH += @{
        HANAME = $HANODENAME;
        HAIP = $HANODE.ipaddress;
        HASTATUS = $HANODE.state;
        HASYNC = $HANODE.hasync;        
        }
        $HANODEname = $null
    }

    if ($HAH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $HAH;
            Columns = "HANAME","HAIP","HASTATUS","HASYNC";
            Headers = "NetScaler Name","IP Address","HA Status","HA Synchronization";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
            }
        $Table = AddWordTable @Params -NoGridLines;
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
        $Table = $null
    }

#endregion NetScaler High Availability

#region NetScaler Global HTTP Parameters
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Global HTTP Parameters"

WriteWordLine 2 0 "NetScaler Global HTTP Parameters"
WriteWordLine 0 0 " "
$nshttpparam = Get-vNetScalerObject -Container config -Object nshttpparam;

$Params = $null
$Params = @{
    Hashtable = @{
        CookieVersion = $nsconfig.cookieversion;
        Drop = $nshttpparam.dropinvalreqs;
    }
    Columns = "CookieVersion","Drop";
    Headers = "Cookie Version","HTTP Drop Invalid Request";
    Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
    AutoFit = $wdAutoFitContent;
}
$Table = AddWordTable @Params -NoGridLines;

FindWordDocumentEnd;
WriteWordLine 0 0 " "
$Table = $null

#endregion NetScaler Global HTTP Parameters

#region NetScaler Global TCP Parameters
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Global TCP Parameters"

WriteWordLine 2 0 "NetScaler Global TCP Parameters"
WriteWordLine 0 0 " "
$nstcpparam = Get-vNetScalerObject -Container config -Object nstcpparam;

$Params = $null
$Params = @{
    Hashtable = @{
        TCP = $nstcpparam.ws;
        SACK = $nstcpparam.SACK;
        NAGLE = $nstcpparam.NAGLE;
    }
    Columns = "TCP","SACK","NAGLE";
    Headers = "TCP Windows Scaling","Selective Acknowledgement","Use Nagle's Algorithm";
    Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
    AutoFit = $wdAutoFitContent;
}
$Table = AddWordTable @Params -NoGridLines;

FindWordDocumentEnd;
WriteWordLine 0 0 " "
$Table = $null
    
#endregion NetScaler Global TCP Parameters

#region NetScaler Global Diameter Parameters

$nsdiameter = Get-vNetScalerObject -Container config -Object nsdiameter; 

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Global Diameter Parameter"

WriteWordLine 2 0 "NetScaler Global Diameter Parameters"
WriteWordLine 0 0 " "
$Params = $null
$Params = @{
    Hashtable = @{
        HOST = $nsdiameter.identity;
        Realm = $nsdiameter.realm;
        Close = $nsdiameter.serverclosepropagation;

    }
    Columns = "HOST","Realm","Close";
    Headers = "Host Identity","Realm","Server Close Propagation";
    Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
    AutoFit = $wdAutoFitContent;
}
$Table = AddWordTable @Params -NoGridLines;

FindWordDocumentEnd;
WriteWordLine 0 0 " "
$Table = $null

#endregion NetScaler Global Diameter Parameters

#region NetScaler Time Zone
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Time zone"
WriteWordLine 2 0 "NetScaler Time Zone"
WriteWordLine 0 0 " "
$Params = $null
$Params = @{
    Hashtable = @{
        TimeZone = $nsconfig.timezone;
    }
    Columns = "TimeZone";
    Headers = "Time Zone";
    Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
    AutoFit = $wdAutoFitContent;
}
$Table = AddWordTable @Params -NoGridLines;
FindWordDocumentEnd;
WriteWordLine 0 0 " "

#endregion NetScaler Time Zone

#region NetScaler Location Database
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Location Database"
WriteWordLine 2 0 "NetScaler Location Database"
WriteWordLine 0 0 " "
$nslocdbs = Get-vNetScalerObject -Container config -Object locationfile;

$LOCDBSH = $null    
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $LOCDBSH = @();

foreach ($nslocdb in $nslocdbs) {

    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be about 400 characters wide!
    $LOCDBSH += @{
            LocationFile = $nslocdb.Locationfile;
            Format = $nslocdb.format;
        }
    }

if ($LOCDBSH.Length -gt 0) {
    $Params = $null
    $Params = @{
        Hashtable = $LOCDBSH;
        Columns = "Locationfile","Format";
        Headers = "Location File", "Format";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
        }

    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;
    WriteWordLine 0 0 " "
    $Table = $null
    } Else {
      WriteWordLine 0 0 " "
      WriteWordLine 0 0 "No Location database has been configured."
      WriteWordLine 0 0 " "
    }


#endregion NetScaler Location Database

#region NetScaler Administration
$selection.InsertNewPage()
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Administration"
WriteWordLine 2 0 "NetScaler System Authentication"
WriteWordLine 0 0 " "

#region Local Administration Users
WriteWordLine 3 0 "NetScaler System Users"
WriteWordLine 0 0 " "
$nssystemusers = Get-vNetScalerObject -Container config -Object systemuser;

$AUTHLOCH = $null    
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $AUTHLOCH = @();

foreach ($nssystemuser in $nssystemusers) {

    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be about 400 characters wide!
    $AUTHLOCH += @{
            LocalUser = $nssystemuser.username;
        }
    }

if ($AUTHLOCH.Length -gt 0) {
    $Params = $null
    $Params = @{
        Hashtable = $AUTHLOCH;
        Columns = "LocalUser";
        Headers = "Local User";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
        }
    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;
    $Table = $null
    }
WriteWordLine 0 0 " "
#endregion Authentication Local Administration Users

#region Authentication Local Administration Groups
WriteWordLine 3 0 "NetScaler System Groups"
WriteWordLine 0 0 " "
$nssystemgroups = Get-vNetScalerObject -Container config -Object systemgroup;
    
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $AUTHGRPH = @();

foreach ($nssystemgroup in $nssystemgroups) {
    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be about 400 characters wide!
    $AUTHGRPH += @{
            SystemGroup = $nssystemgroup.groupname;
        }
    }

if ($AUTHGRPH.Length -gt 0) {
    $Params = $null
    $Params = @{
        Hashtable = $AUTHGRPH;
        Columns = "SystemGroup";
        Headers = "System Group";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
        }
    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;
    $Table = $null
    }
else { WriteWordLine 0 0 "No Local Groups have been configured"}
WriteWordLine 0 0 " "

#endregion Authentication Local Administration Groups

#region RPC Nodes

WriteWordLine 2 0 "NetScaler RPC Nodes"
WriteWordLine 0 0 " "
$rpcnodecounter = Get-vNetScalerObjectCount -Container config -Object nsrpcnode; 
$rpcnodecount = $rpcnodecounter.__count
$rpcnodes = Get-vNetScalerObject -Container config -Object nsrpcnode;

if($rpcnodecounter.__count -le 0) { WriteWordLine 0 0 "No RPC Nodes have been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $RPCCONFIGH = @();

    foreach ($rpcnode in $rpcnodes) {
        $RPCCONFIGH += @{
            IPADDR = $rpcnode.ipaddress;
            SOURCE = $rpcnode.srcip;
            SECURE = $rpcnode.secure;
            }
        }
        if ($RPCCONFIGH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $RPCCONFIGH;
                Columns = "IPADDR","SOURCE","SECURE";
                Headers = "IP Address","Source IP Address", "Secure";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            $Table = $null
            }
        }

#endregion RPC Nodes

#endregion NetScaler Administration

#region NetScaler Features
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Features"

$selection.InsertNewPage()

WriteWordLine 1 0 "NetScaler Features"
WriteWordLine 0 0 " "
If ($Version -gt $ScriptVersion) {
    WriteWordLine 0 0 ""
    WriteWordLine 0 0 "Warning: You are using Citrix NetScaler version $Version, features added since version $ScriptVersion will not be shown."
    WriteWordLine 0 0 ""
    }
#region NetScaler Basic Features
WriteWordLine 2 0 "NetScaler Basic Features"
WriteWordLine 0 0 " "
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $AdvancedConfiguration = @(
    @{ Description = "Feature"; Value = "State" }
	@{ Description = "Application Firewall"; Value = $FEATAppFw }
	@{ Description = "Authentication, Authorization and Auditing"; Value = $FEATAAA }
    @{ Description = "Content Filter"; Value = $FEATCF }
    @{ Description = "Content Switching"; Value = $FEATCS }
    @{ Description = "HTTP Compression"; Value = $FEATCMP }
    @{ Description = "Integrated Caching"; Value = $FEATIC }
    @{ Description = "Load Balancing"; Value = $FEATLB }
    @{ Description = "NetScaler Gateway"; Value = $FEATSSLVPN }
    @{ Description = "Rewrite"; Value = $FEATRewrite }
    @{ Description = "SSL Offloading"; Value = $FEATSSL }
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $AdvancedConfiguration;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params -NoGridLines -List;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null      
 
WriteWordLine 0 0 " "

#endregion NetScaler Basic Features

#region NetScaler Advanced Features
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Advanced Features"

WriteWordLine 2 0 "NetScaler Advanced Features"
WriteWordLine 0 0 " "
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $AdvancedFeatures = @(
    @{ Description = "Feature"; Value = "State" }
	@{ Description = "Web Logging"; Value = $FEATWL }
    @{ Description = "Surge Protection"; Value = $FEATSP }
    @{ Description = "Cache Redirection"; Value = $FEATCR }
    @{ Description = "Sure Connect"; Value = $FEATSC }
    @{ Description = "Priority Queuing"; Value = $FEATPQ }
    @{ Description = "Global Server Load Balancing"; Value = $FEATGSLB }
    @{ Description = "Http DoS Protection"; Value = $FEATHDOSP }
    @{ Description = "Vpath"; Value = $FEATVpath }
    @{ Description = "Integrated Caching"; Value = $FEATIC }
    @{ Description = "OSPF Routing"; Value = $FEATOSPF }
	@{ Description = "RIP Routing"; Value = $FEATRIP }
    @{ Description = "BGP Routing"; Value = $FEATBGP }
    @{ Description = "IPv6 protocol translation "; Value = $FEATIPv6PT }
    @{ Description = "Responder"; Value = $FEATRESPONDER }
    @{ Description = "Edgesight Monitoring HTML Injection"; Value = $FEATHTMLInjection }
    @{ Description = "OSPF Routing"; Value = $FEATOSPF }
    @{ Description = "NetScaler Push"; Value = $FEATPUSH }
    @{ Description = "AppFlow"; Value = $FEATAppFlow }
    @{ Description = "CloudBridge"; Value = $FEATCloudBridge }
    @{ Description = "ISIS Routing"; Value = $FEATISIS }
    @{ Description = "CallHome"; Value = $FEATCH }
    @{ Description = "AppQoE"; Value = $FEATAppQoE }
    @{ Description = "Front End Optimization"; Value = $FEATfeo }
    @{ Description = "Large Scale NAT"; Value = $FEATlsn }
    @{ Description = "RDP Proxy"; Value = $FEATrdpproxy }
    @{ Description = "Reputation"; Value = $FEATrep }
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $AdvancedFeatures;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params -NoGridLines -List;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null      
 
WriteWordLine 0 0 " "

#endregion NetScaler Advanced Features

#endregion NetScaler Features

#region NetScaler Modes
$selection.InsertNewPage()
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Modes"

WriteWordLine 1 0 "NetScaler Modes"
WriteWordLine 0 0 " "
$nsmode = Get-vNetScalerObject -Container config -Object nsmode; 

## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $ADVModes = @(
    @{ Description = "Mode"; Value = "Enabled"}  
    @{ Description = "Fast Ramp"; Value = $nsmode.fr}        
    @{ Description = "Layer 2 mode"; Value = $nsmode.l2}        
    @{ Description = "Use Source IP"; Value = $nsmode.usip}        
    @{ Description = "Client SideKeep-alive"; Value = $nsmode.cka}        
    @{ Description = "TCP Buffering"; Value = $nsmode.TCPB}        
    @{ Description = "MAC-based forwarding"; Value = $nsmode.MBF}
    @{ Description = "Edge configuration"; Value = $nsmode.Edge}        
    @{ Description = "Use Subnet IP"; Value = $nsmode.USNIP}        
    @{ Description = "Use Layer 3 Mode"; Value = $nsmode.L3}        
    @{ Description = "Path MTU Discovery"; Value = $nsmode.PMTUD}
    @{ Description = "Media Classification"; Value = $nsmode.mediaclassification}        
    @{ Description = "Static Route Advertisement"; Value = $nsmode.SRADV}        
    @{ Description = "Direct Route Advertisement"; Value = $nsmode.DRADV}        
    @{ Description = "Intranet Route Advertisement"; Value = $nsmode.IRADV}        
    @{ Description = "Ipv6 Static Route Advertisement"; Value = $nsmode.SRADV6}        
    @{ Description = "Ipv6 Direct Route Advertisement"; Value = $nsmode.DRADV6}        
    @{ Description = "Bridge BPDUs" ; Value = $nsmode.BridgeBPDUs}
    @{ Description = "Rise APBR"; Value = $nsmode.rise_apbr}        
    @{ Description = "Rise RHI" ; Value = $nsmode.rise_rhi}       
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $ADVModes;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params -NoGridLines -List;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null      
WriteWordLine 0 0 " "

$selection.InsertNewPage()

#endregion NetScaler Modes

#region NetScaler Monitoring
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Monitoring"

WriteWordLine 1 0 "NetScaler Monitoring"
WriteWordLine 0 0 " "
WriteWordLine 2 0 "SNMP Community"
WriteWordLine 0 0 " "

$snmpcommunitycounter = Get-vNetScalerObjectCount -Container config -Object snmpcommunity; 
$snmpcommunitycount = $snmpcommunitycounter.__count
$snmpcoms = Get-vNetScalerObject -Container config -Object snmpcommunity;    

if($snmpcommunitycount -le 0) { WriteWordLine 0 0 "No SNMP Community has been configured"} else {

    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $SNMPCOMH = @();

    foreach ($snmpcom in $snmpcoms) {
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be about 400 characters wide!
        $SNMPCOMH += @{
                SNMPCommunity = $snmpcom.communityname;
                Permissions = $snmpcom.permissions;
            }
        }
        if ($SNMPCOMH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $SNMPCOMH;
                Columns = "SNMPCommunity","Permissions";
                Headers = "SNMP Community","Permissions";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            $Table = $null
        }
    }
WriteWordLine 0 0 " "

WriteWordLine 2 0 "SNMP Manager"
WriteWordLine 0 0 " "
$snmpmanagercounter = Get-vNetScalerObjectCount -Container config -Object snmpmanager; 
$snmpmanagercount = $snmpmanagercounter.__count

if($snmpmanagercount -le 0) { WriteWordLine 0 0 "No SNMP Manager has been configured"} else {
    
    $snmpmanagers = Get-vNetScalerObject -Container config -Object snmpmanager; 

    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $SNMPMANSH = @();

    foreach ($snmpmanager in $snmpmanagers) {

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be about 400 characters wide!
        $SNMPMANSH += @{
                SNMPManager = $snmpmanager.ipaddress;
                Netmask = $snmpmanager.netmask;
            }
        }
        if ($SNMPMANSH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $SNMPMANSH;
                Columns = "SNMPManager","Netmask";
                Headers = "SNMP Manager","Netmask";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            $Table = $null
        }
    }
WriteWordLine 0 0 ""

WriteWordLine 2 0 "SNMP Alert"
WriteWordLine 0 0 " "
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $SNMPALERTSH = @();

$snmpalarms = Get-vNetScalerObject -Container config -Object snmpalarm; 

foreach ($snmpalarm in $snmpalarms) {
        
    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be about 400 characters wide!

    $SNMPALERTSH += @{
            Alarm = $snmpalarm.trapname;
            State = $snmpalarm.state;
            Time = $snmpalarm.time;
            TimeOut = $snmpalarm.timeout;
            Severity = $snmpalarm.severity;
            Logging = $snmpalarm.logging;
        }
    }
    if ($SNMPALERTSH.Length -gt 0) {
        $Params = @{
            Hashtable = $SNMPALERTSH;
            Columns = "Alarm","State","Time","TimeOut","Severity","Logging";
            Headers = "NetScaler Alarm","State","Time","Time-Out","Severity","Logging";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
            }
        $Table = AddWordTable @Params -NoGridLines;
        FindWordDocumentEnd;
        $Table = $null
    }

WriteWordLine 0 0 ""


WriteWordLine 2 0 "SNMP Traps"
WriteWordLine 0 0 " "
$snmptrapscounter = Get-vNetScalerObjectCount -Container config -Object snmptrap; 
$snmptrapscount = $snmptrapscounter.__count

$snmptraps = Get-vNetScalerObject -Container config -Object snmptrap; 

if($snmptrapscount -le 0) { WriteWordLine 0 0 "No SNMP Trap has been configured"} else {
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $SNMPTRAPSH = @();

    foreach ($snmptrap in $snmptraps) {

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be about 400 characters wide!
        $SNMPTRAPSH += @{
                Type = $snmptrap.trapclass;
                Destination = $snmptrap.trapdestination;
                Version = $snmptrap.version;
                Port = $snmptrap.destport;
                Name = $snmptrap.communityname;
            }
        }
        if ($SNMPTRAPSH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $SNMPTRAPSH;
                Columns = "Type","Destination","Version","Port","Name";
                Headers = "Type","Trap Destination","Version","Destination Port","Community Name";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            $Table = $null
        }
    }
WriteWordLine 0 0 " "

$selection.InsertNewPage()

#endregion NetScaler Monitoring

#region NetScaler Auditing

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Auditing"

WriteWordLine 1 0 "NetScaler Auditing"
WriteWordLine 0 0 " "

#region Syslog Parameters

WriteWordLine 2 0 "Syslog Parameters"
WriteWordLine 0 0 " "

$syslogparams = Get-vNetScalerObject -Object auditsyslogparams;

## IB - Create an array of hashtables to store our columns.
## IB - about column names as we'll utilise a -List(view)!
[System.Collections.Hashtable[]] $SYSLOGPARAMH = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "Server IP"; Column2 = $syslogparams.serverip; }
    @{ Column1 = "Server Port"; Column2 = $syslogparams.serverport; }
    @{ Column1 = "Date Format"; Column2 = $syslogparams.dateformat; }
    @{ Column1 = "Log level"; Column2 = $syslogserver.loglevel -join ","; }
    @{ Column1 = "Log Facility"; Column2 = $syslogparams.logfacility; }
    @{ Column1 = "Log TCP Messages"; Column2 = $syslogparams.tcp; }
    @{ Column1 = "Log ACL Messages"; Column2 = $syslogparams.acl; }
    @{ Column1 = "TimeZone"; Column2 = $syslogparams.timezone; }
    @{ Column1 = "Log User Defined Messages"; Column2 = $syslogparams.userdefinedauditlog; }
    @{ Column1 = "AppFlow Export"; Column2 = $syslogparams.appflowexport; }
    @{ Column1 = "Log Large Scale NAT Messages"; Column2 = $syslogparams.lsn; }
    @{ Column1 = "Log ALG Messages"; Column2 = $syslogparams.alg;}
    @{ Column1 = "Log Subscriber Session Messages"; Column2 = $syslogparams.subscriberlog; }
    @{ Column1 = "Log DNS Messages"; Column2 = $syslogparams.dns; }
    
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $SYSLOGPARAMH;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
$Table = $null


#endregion Syslog Parameters

#region Syslog Policies

WriteWordLine 2 0 "Syslog Policies"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `tSyslog Policies"

$syslogpolicies = Get-vNetScalerObject -Container config -Object auditsyslogpolicy;

[System.Collections.Hashtable[]] $SYSLOGPOLH = @();
foreach ($syslogpolicy in $syslogpolicies) {
    $SYSLOGPOLH += @{
            NAME = $syslogpolicy.name;
            RULE = $syslogpolicy.rule;
            ACTION = $syslogpolicy.action;
    }
}


    ## IB - Create an array of hashtables to store our columns. Note: If we need the
    ## IB - headers to include spaces we can override these at table creation time.
    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = $SYSLOGPOLH;
        Columns = "NAME","RULE","ACTION";
        Headers = "Policy Name","Rule","Action";
        AutoFit = $wdAutoFitContent;
        Format = -235; ## IB - Word constant for Light List Accent 5
    }

    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;
    WriteWordLine 0 0 " "
    $Table = $null


#endregion Syslog Policies

#region Syslog Actions

WriteWordLine 2 0 "Syslog Servers"
WriteWordLine 0 0 " "

$syslogservers = Get-vNetScalerObject -Container config -Object auditsyslogaction; 

foreach ($syslogserver in $syslogservers) {
  $syslogservername = $syslogserver.name
  WriteWordLine 3 0 "Syslog Server: $syslogservername"
  WriteWordLine 0 0 " "

  [System.Collections.Hashtable[]] $SYSLOGSRVH = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "Server IP"; Column2 = $syslogserver.serverip; }
    @{ Column1 = "Server Domain Name"; Column2 = $syslogserver.serverdomainname; }
    @{ Column1 = "DNS Resolution Retry"; Column2 = $syslogserver.domainresolveretry; }
    @{ Column1 = "LB vServer Name"; Column2 = $syslogserver.lbvservername; }
    @{ Column1 = "Server Port"; Column2 = $syslogserver.serverport; }
    @{ Column1 = "Log level"; Column2 = $syslogserver.loglevel -join ","; }
    @{ Column1 = "Date Format"; Column2 = $syslogserver.dateformat; }
    @{ Column1 = "Log Facility"; Column2 = $syslogserver.logfacility; }
    @{ Column1 = "Time Zone"; Column2 = $syslogserver.timezone; }
    @{ Column1 = "TCP Logging"; Column2 = $syslogserver.tcp; }
    @{ Column1 = "ACL Logging"; Column2 = $syslogserver.acl; }
    @{ Column1 = "User Configurable Log Messages"; Column2 = $syslogserver.userdefinedauditlog; }
    @{ Column1 = "AppFlow Logging"; Column2 = $syslogserver.appflowexport; }
    @{ Column1 = "Large Scale NAT Logging"; Column2 = $syslogserver.lsn; }
    @{ Column1 = "ALG messages Logging"; Column2 = $syslogserver.alg; }
    @{ Column1 = "Subscriber Logging"; Column2 = $syslogserver.subscriberlog; }
    @{ Column1 = "DNS Logging"; Column2 = $syslogserver.dns; }
    @{ Column1 = "Transport Type"; Column2 = $syslogserver.transport; }
    @{ Column1 = "Net Profile"; Column2 = $syslogserver.netprofile; }
    @{ Column1 = "Max Log Data to hold"; Column2 = $syslogserver.maxlogdatasizetohold; }
    
    
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $SYSLOGSRVH;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
$Table = $null

} #end foreach

#endregion Syslog Actions

#endregion NetScaler Auditing

#endregion NetScaler System Information

#region NetScaler Networking

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Networking"

WriteWordLine 1 0 "NetScaler Networking"
WriteWordLine 0 0 " "
#region NetScaler Interfaces

WriteWordLine 2 0 "NetScaler Interfaces"
WriteWordLine 0 0 " "
$InterfaceCounter = Get-vNetScalerObjectCount -Container config -Object interface; 
$InterfaceCount = $InterfaceCounter.__count
$Interfaces = Get-vNetScalerObject -Container config -Object interface;

if($InterfaceCounter.__count -le 0) { WriteWordLine 0 0 "No Interface has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $NICH = @();

    foreach ($Interface in $Interfaces) {        
        $NICH += @{
            InterfaceID = $interface.devicename;
            InterfaceType = $interface.intftype;
            HAMonitoring = $interface.hamonitor;
            State = $interface.state;
            AutoNegotiate = $interface.autoneg;
            }
        }

        if ($NICH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $NICH;
                Columns = "InterfaceID","InterfaceType","HAMonitoring","State","AutoNegotiate";
                Headers = "Interface ID","Interface Type","HA Monitoring","State","Auto Negotiate";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            $Table = $null
            }
        }

#endregion NetScaler Interfaces

#region NetScaler Channels

WriteWordLine 2 0 "NetScaler Channels"
WriteWordLine 0 0 " "
$ChannelCounter = Get-vNetScalerObjectCount -Container config -Object channel; 
$ChannelCount = $ChannelCounter.__count
$Channels = Get-vNetScalerObject -Container config -Object interface;

if($ChannelCounter.__count -le 0) { 
  WriteWordLine 0 0 "No Channels have been configured"
  WriteWordLine 0 0 " "
  } else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $CHANH = @();

    foreach ($Channel in $Channels) {        
        $CHANH += @{
            CHANNEL = $channel.devicename;
            Alias = $CHANNEL.ifalias;
            HA = $channel.hamonitor;
            State = $channel.state;
            Speed = $channel.reqspeed;
            Tagall = $channel.tagall;
            MTU = $channel.mtu;
            }
        }

        if ($CHANH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $CHANH;
                Columns = "CHANNEL","Alias","HA","State","Speed","Tagall";
                Headers = "Channel","Alias","HA Monitoring","State","Speed","Tag all vLAN";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            $Table = $null
            }
        }

#endregion NetScaler Channels

#region NetScaler IP addresses

WriteWordLine 2 0 "NetScaler IP addresses"
WriteWordLine 0 0 " "
$IPs = Get-vNetScalerObject -Container config -Object nsip;

## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $IPADDRESSH = @();

foreach ($IP in $IPs) {
    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be about 400 characters wide!

    $IPADDRESSH += @{
        IPAddress = $IP.ipaddress;
        SubnetMask = $IP.netmask;
        TrafficDomain = $IP.td;
        Type = $IP.type;
        vServer = $IP.vserver;
        MGMT = $IP.mgmtaccess;
        SNMP = $IP.snmp;
    }
}

$Params = $null
$Params = @{
    Hashtable = $IPADDRESSH;
    Columns = "IPAddress","SubnetMask","TrafficDomain","Type","vServer","MGMT","SNMP";
    Headers = "IP Address","Subnet Mask","Traffic Domain","Type","vServer","Management","SNMP";
    Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
    AutoFit = $wdAutoFitContent;
    }
$Table = AddWordTable @Params -NoGridLines;
FindWordDocumentEnd;
WriteWordLine 0 0 " "
$Table = $null
#endregion NetScaler IP addresses

#region NetScaler vLAN

WriteWordLine 2 0 "NetScaler vLANs"
WriteWordLine 0 0 " "
$VLANCounter = Get-vNetScalerObjectCount -Container config -Object vlan; 
$VLANCount = $VLANCounter.__count
$VLANS = Get-vNetScalerObject -Container config -Object vlan;

if($VLANCounter.__count -le 0) { WriteWordLine 0 0 "No vLAN has been configured"} else {
    $vLANH = $null
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $vLANH = @();

    $vLANH += @{
        vLANName = "default";
        VLANID = "1";
        }

    foreach ($VLAN in $VLANS) {
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be about 400 characters wide!
        $vLANH += @{
            vLANName = $VLAN.aliasname;
            VLANID = $VLAN.id;
            }
        }
        
        if ($vLANH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $vLANH;
                Columns = "VLANNAME","VLANID";
                Headers = "vLAN Alias","ID";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            $Table = $null
            }
        }

#endregion NetScaler vLAN

#region routing table

WriteWordLine 2 0 "NetScaler Routing Table"
WriteWordLine 0 0 " "

$nsroute = Get-vNetScalerObject -Container config -Object route;
    
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $ROUTESH = @();

foreach ($ROUTE in $nsroute) {
    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be about 400 characters wide!

    $ROUTESH += @{
        Network = $ROUTE.network;
        Subnet = $ROUTE.netmask;
        Gateway = $ROUTE.gateway;
        Distance = $ROUTE.distance;
        Weight = $ROUTE.weight;
        Cost = $ROUTE.cost;
        TD = $ROUTE.td;
        }
    }

    if ($ROUTESH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $ROUTESH;
            Columns = "Network","Subnet","Gateway","Distance","Weight","Cost","TD";
            Headers = "Network","Subnet","Gateway","Distance","Weight","Cost","Traffic Domain";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
            }
        $Table = AddWordTable @Params -NoGridLines;
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
        $Table = $null
        }

#endregion routing table

#region NetScaler Traffic Domains
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Traffic Domains"

WriteWordLine 2 0 "NetScaler Traffic Domains"
WriteWordLine 0 0 " "

$TDcounter = Get-vNetScalerObjectCount -Container config -Object nstrafficdomain; 
$TDcount = $TDcounter.__count
$TDs = Get-vNetScalerObject -Container config -Object nstrafficdomain;

if($TDcounter.__count -le 0) { WriteWordLine 0 0 "No Traffic Domains have been configured"} else {
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $TDSH = @();

    foreach ($TD in $TDs) {
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be about 400 characters wide!
        $TDSH += @{
            ## IB - This table will only have 1 row so create the nested hashtable inline
            ID = $TD.td;
            Alias = $TD.aliasname;
            vmac = $TD.vmac;
            State = $TD.state;
        }
    }
    if ($TDSH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $TDSH;
            Columns = "ID","Alias","vmac","State";
            Headers = "Traffic Domain ID","Traffic Domain Alias","Traffic Domain vmac","State";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
            }
        $Table = AddWordTable @Params -NoGridLines;
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
        $Table = $null
    }
}
    
#endregion NetScaler Traffic Domains

#region NetScaler DNS Configuration
$selection.InsertNewPage()
WriteWordLine 1 0 "NetScaler DNS Configuration"
WriteWordLine 0 0 " "

#region dns name servers

WriteWordLine 2 0 "NetScaler DNS Name Servers"
WriteWordLine 0 0 " "

$dnsnameservercounter = Get-vNetScalerObjectCount -Container config -Object dnsnameserver; 
$dnsnameservecount = $dnsnameservercounter.__count
$dnsnameservers = Get-vNetScalerObject -Container config -Object dnsnameserver;

if($dnsnameservercounter.__count -le 0) { WriteWordLine 0 0 "No DNS Name Server has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $DNSNAMESERVERH = @();

    foreach ($DNSNAMESERVER in $DNSNAMESERVERS) {
        $DNSNAMESERVERH += @{
            DNSServer = $dnsnameserver.ip;
            State = $dnsnameserver.state;
            Prot = $dnsnameserver.type;
            }
        }

        if ($DNSNAMESERVERH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $DNSNAMESERVERH;
                Columns = "DNSServer","State","Prot";
                Headers = "DNS Name Server","State","Protocol";;
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            $Table = $null
            }
        }
      
#endregion dns name servers

#region DNS Address Records
WriteWordLine 0 0 " "
WriteWordLine 2 0 "NetScaler DNS Address Records"
WriteWordLine 0 0 " "
$dnsaddreccounter = Get-vNetScalerObjectCount -Container config -Object dnsaddrec; 
$dnsaddreccount = $dnsaddreccounter.__count
$dnsaddrecs = Get-vNetScalerObject -Container config -Object dnsaddrec;

if($dnsaddreccounter.__count -le 0) { WriteWordLine 0 0 "No DNS Name Server has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $DNSRECORDCONFIGH = @();

    foreach ($dnsaddrec in $dnsaddrecs) {
        $DNSRECORDCONFIGH += @{
            DNSRecord = $dnsaddrec.hostname;
            IPAddress = $dnsaddrec.ipaddress;
            TTL = $dnsaddrec.ttl;
            AUTHTYPE = $dnsaddrec.authtype;
            }
        }
        if ($DNSRECORDCONFIGH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $DNSRECORDCONFIGH;
                Columns = "DNSRecord","IPAddress","TTL","AUTHTYPE";
                Headers = "DNS Record","IP Address","TTL","Authentication Type";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            $Table = $null
            }
        }

#endregion DNS Address Records

#region DNS AAA Records
WriteWordLine 0 0 " "
WriteWordLine 2 0 "NetScaler DNS AAA Records"
WriteWordLine 0 0 " "
$dnsaaaareccounter = Get-vNetScalerObjectCount -Container config -Object dnsaaaarec; 
$dnsaaaareccount = $dnsaaaareccounter.__count
$dnsaaaarecs = Get-vNetScalerObject -Container config -Object dnsaaaarec;

if($dnsaaaareccounter.__count -le 0) { WriteWordLine 0 0 "No DNS AAA records have been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $DNSRECORDCONFIGH = @();

    foreach ($dnsaaarec in $dnsaaaarecs) {
        $DNSRECORDCONFIGH += @{
            DNSRecord = $dnsaaarec.hostname;
            IPAddress = $dnsaaarec.ipv6address;
            TTL = $dnsaaarec.ttl;
            AUTHTYPE = $dnsaaarec.authtype;
            }
        }
        if ($DNSRECORDCONFIGH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $DNSRECORDCONFIGH;
                Columns = "DNSRecord","IPAddress","TTL","AUTHTYPE";
                Headers = "DNS Record","IP Address","TTL","Authentication Type";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            $Table = $null
            }
        }

#endregion DNS AAA Records

#region DNS CNAME Records
WriteWordLine 0 0 " "
WriteWordLine 2 0 "NetScaler DNS CNAME Records"
WriteWordLine 0 0 " "
$dnscnamereccounter = Get-vNetScalerObjectCount -Container config -Object dnscnamerec; 
$dnscnamereccount = $dnscnamereccounter.__count
$dnscnamerecs = Get-vNetScalerObject -Container config -Object dnscnamerec;

if($dnscnamereccounter.__count -le 0) { WriteWordLine 0 0 "No DNS CNAME records have been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $DNSRECORDCONFIGH = @();

    foreach ($dnscnamerec in $dnscnamerecs) {
        $DNSRECORDCONFIGH += @{
            DNSRecord = $dnscnamerec.aliasname;
            IPAddress = $dnscnamerec.canonicalname;
            TTL = $dnscnamerec.ttl;
            AUTHTYPE = $dnscnamerec.authtype;
            }
        }
        if ($DNSRECORDCONFIGH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $DNSRECORDCONFIGH;
                Columns = "DNSRecord","IPAddress","TTL","AUTHTYPE";
                Headers = "Alias Name","Canonical Name","TTL","Authentication Type";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            $Table = $null
            }
        }

#endregion DNS CNAME Records

#region DNS MX Records
WriteWordLine 0 0 " "
WriteWordLine 2 0 "NetScaler DNS MX Records"
WriteWordLine 0 0 " "
$dnsmxreccounter = Get-vNetScalerObjectCount -Container config -Object dnsmxrec; 
$dnsmxreccount = $dnsmxreccounter.__count
$dnsmxrecs = Get-vNetScalerObject -Container config -Object dnsmxrec;

if($dnsmxreccounter.__count -le 0) { WriteWordLine 0 0 "No DNS MX records have been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $DNSMXRECORDCONFIGH = @();

    foreach ($dnsmxrec in $dnsmxrecs) {
        $DNSMXRECORDCONFIGH += @{
            DOMAIN = $dnsmxrec.domain;
            MX = $dnsmxrec.mx;
            TTL = $dnsmxrec.ttl;
            AUTHTYPE = $dnsmxrec.authtype;
            }
        }
        if ($DNSMXRECORDCONFIGH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $DNSMXRECORDCONFIGH;
                Columns = "DOMAIN","MX","TTL","AUTHTYPE";
                Headers = "Domain","MX","TTL","Authentication Type";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            $Table = $null
            }
        }

#endregion DNS MX Records

#region DNS NS Records
WriteWordLine 0 0 " "
WriteWordLine 2 0 "NetScaler DNS NS Records"
WriteWordLine 0 0 " "
$dnsnsreccounter = Get-vNetScalerObjectCount -Container config -Object dnsnsrec; 
$dnsnsreccount = $dnsnsreccounter.__count
$dnsnsrecs = Get-vNetScalerObject -Container config -Object dnsnsrec;

if($dnsnsreccounter.__count -le 0) { WriteWordLine 0 0 "No DNS NS records have been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $DNSNSRECORDCONFIGH = @();

    foreach ($dnsnsrec in $dnsnsrecs) {
        $DNSNSRECORDCONFIGH += @{
            DOMAIN = $dnsnsrec.domain;
            NS = $dnsnsrec.nameserver;
            TTL = $dnsnsrec.ttl;
            AUTHTYPE = $dnsnsrec.authtype;
            }
        }
        if ($DNSNSRECORDCONFIGH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $DNSNSRECORDCONFIGH;
                Columns = "DOMAIN","NS","TTL","AUTHTYPE";
                Headers = "Domain","NameServer","TTL","Authentication Type";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            $Table = $null
            }
        }

#endregion DNS NS Records

#region DNS SOA Records
WriteWordLine 0 0 " "
WriteWordLine 2 0 "NetScaler DNS SOA Records"
WriteWordLine 0 0 " "
$dnssoareccounter = Get-vNetScalerObjectCount -Container config -Object dnssoarec; 
$dnssoareccount = $dnsnsreccounter.__count
$dnssoarecs = Get-vNetScalerObject -Container config -Object dnssoarec;

if($dnssoareccounter.__count -le 0) { WriteWordLine 0 0 "No DNS SOA records have been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $DNSSOARECORDCONFIGH = @();

    foreach ($dnssoarec in $dnssoarecs) {
        $DNSSOARECORDCONFIGH += @{
            DOMAIN = $dnssoarec.domain;
            ORIGIN = $dnssoarec.originserver;
            CONTACT = $dnssoarec.contact;
            SERIAL = $dnssoarec.serial;
            TTL = $dnssoarec.ttl;
            AUTHTYPE = $dnssoarec.authtype;
            }
        }
        if ($DNSSOARECORDCONFIGH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $DNSSOARECORDCONFIGH;
                Columns = "DOMAIN","ORIGIN","CONTACT","SERIAL","TTL","AUTHTYPE";
                Headers = "Domain","Origin Server", "Admin Contact","Serial Number","TTL","Authentication Type";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            $Table = $null
            }
        }

#endregion DNS SOA Records

#endregion NetScaler DNS Configuration

#region NetScaler ACL
$selection.InsertNewPage()
WriteWordLine 1 0 "NetScaler ACL Configuration"
WriteWordLine 0 0 " "
#region NetScaler Simple ACL

WriteWordLine 2 0 "NetScaler Simple ACL"
WriteWordLine 0 0 " "
$nssimpleaclCounter = Get-vNetScalerObjectCount -Container config -Object nssimpleacl; 
$nssimpleaclCount = $nssimpleaclCounter.__count
$nssimpleacls = Get-vNetScalerObject -Container config -Object nssimpleacl;

if($nssimpleaclCounter.__count -le 0) { WriteWordLine 0 0 "No Simple ACL has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $nssimpleaclH = @();

    foreach ($nssimpleacl in $nssimpleacls) {        
        $nssimpleaclH += @{
            ACLNAME = $nssimpleacl.aclname;
            ACTION = $nssimpleacl.aclaction;
            SOURCEIP = $nssimpleacl.srcip;
            DESTPORT = $nssimpleacl.destport;
            PROT = $nssimpleacl.protocol;
            TD = $nssimpleacl.td;
            }
        }
        if ($nssimpleaclH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $nssimpleaclH;
                Columns = "ACLNAME","ACTION","SOURCEIP","DESTPORT","PROT","TD";
                Headers = "ACL Name","Action","Source IP","Destination Port","Protocol","Traffic Domain";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
        }


#endregion NetScaler Simple ACL IPv4

#region NetScaler Extended ACL
WriteWordLine 0 0 " "
WriteWordLine 2 0 "NetScaler Extended ACL"
WriteWordLine 0 0 " "
$nsaclCounter = Get-vNetScalerObjectCount -Container config -Object nsacl; 
$nsaclCount = $nsaclCounter.__count
$nsacls = Get-vNetScalerObject -Container config -Object nsacl;

if($nsaclCounter.__count -le 0) { WriteWordLine 0 0 "No Extended ACL has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $nsaclH = @();

    foreach ($nsacl in $nsacls) {        
        $nsaclH += @{
            ACLNAME = $nsacl.aclname;
            ACTION = $nsacl.aclaction;
            SOURCEIP = $nsacl.srcipval;
            TD = $nsacl.td;
            }
        }
        if ($nsaclH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $nsaclH;
                Columns = "ACLNAME","ACTION","SOURCEIP","TD";
                Headers = "ACL Name","Action","Source IP","Traffic Domain";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
        }


#endregion NetScaler Extended ACL IPv4
#endregion NetScaler ACL

#endregion NetScaler Networking

#region NetScaler Authentication
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Authentication"

$selection.InsertNewPage()

WriteWordLine 1 0 "NetScaler Authentication"
WriteWordLine 0 0 " "
#region Authentication LDAP Policies
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler LDAP Authentication"
WriteWordLine 2 0 "NetScaler LDAP Policies"
WriteWordLine 0 0 " "
$authpolsldap = Get-vNetScalerObject -Container config -Object authenticationldappolicy;

If (!$authpolsldap) {
WriteWordLine 0 0 "There are no LDAP authentication policies configured. "
}

## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $AUTHLDAPPOLH = @();

foreach ($authpolldap in $authpolsldap) {
                
    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be about 400 characters wide!
    $AUTHLDAPPOLH += @{
            Policy = $authpolldap.name;
            Expression = $authpolldap.rule;
            Action = $authpolldap.reqaction;
    }
}
        
if ($AUTHLDAPPOLH.Length -gt 0) {
    $Params = $null
    $Params = @{
        Hashtable = $AUTHLDAPPOLH;
        Columns = "Policy","Expression","Action";
        Headers = "LDAP Policy","Expression","LDAP Action";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
        }
    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;

    }

WriteWordLine 0 0 " "
$Table = $null

#endregion Authentication LDAP Policies

#region Authentication LDAP
WriteWordLine 2 0 "NetScaler LDAP authentication actions"
WriteWordLine 0 0 " "
$authactsldap = Get-vNetScalerObject -Container config -Object authenticationldapaction;
If (!$authactsldap) {
 WriteWordLine 0 0 "There are no LDAP authentication servers configured. "
}


foreach ($authactldap in $authactsldap) {
    $ACTNAMELDAP = $authactldap.name
    WriteWordLine 3 0 "LDAP Authentication action $ACTNAMELDAP";
    WriteWordLine 0 0 " "
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $LDAPCONFIG = @(
    @{ Description = "Description"; Value = "Configuration"; }
    @{ Description = "LDAP Server IP"; Value = $authactldap.serverip; }
    @{ Description = "LDAP Server Port"; Value = $authactldap.serverport; }
    @{ Description = "LDAP Server Time-Out"; Value = $authactldap.authtimeout; }
    @{ Description = "Validate Certificate"; Value = $authactldap.validateservercert; }
    @{ Description = "LDAP Base OU"; Value = $authactldap.ldapbase; }
    @{ Description = "LDAP Bind DN"; Value = $authactldap.ldapbinddn; }
    @{ Description = "Login Name"; Value = $authactldap.ldaploginname; }
    @{ Description = "Sub Attribute Name"; Value = $authactldap.ssonameattribute; }
    @{ Description = "Security Type"; Value = $authactldap.sectype; }   
    @{ Description = "Password Changes"; Value = $authactldap.passwdchange; }
    @{ Description = "Group attribute name"; Value = $authactldap.groupattrname; }
    @{ Description = "LDAP Single Sign On Attribute"; Value = $authactldap.ssonameattribute; }
    @{ Description = "Authentication"; Value = $authactldap.authentication; }
    @{ Description = "User Required"; Value = $authactldap.requireuser; }
    @{ Description = "LDAP Referrals"; Value = $authactldap.maxldapreferrals; }
    @{ Description = "Nested Group Extraction"; Value = $authactldap.nestedgroupextraction; }
    @{ Description = "Maximum Nesting level"; Value = $authactldap.maxnestinglevel; }
    );

    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = $LDAPCONFIG;
        Columns = "Description","Value";
        AutoFit = $wdAutoFitContent
        Format = -235; ## IB - Word constant for Light List Accent 5
    }
    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -NoGridLines -List;

	FindWordDocumentEnd;
	$TableRange = $Null
	$Table = $Null
    $selection.InsertNewPage()
}

WriteWordLine 0 0 " "
#endregion Authentication LDAP

#region Authentication Radius Policies
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Radius Authentication"
WriteWordLine 2 0 "NetScaler Radius Policies"
WriteWordLine 0 0 " "
$authpolsradius = Get-vNetScalerObject -Container config -Object authenticationradiuspolicy;

If (!$authpolsradius) {
  WriteWordLine 0 0 "There are no RADIUS authentication policies configured."
}
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $AUTHRADIUSPOLH = @();

foreach ($authpolradius in $authpolsradius) {
                
    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be about 400 characters wide!
    $AUTHRADIUSPOLH += @{
            Policy = $authpolradius.name;
            Expression = $authpolradius.rule;
            Action = $authpolradius.reqaction;
    }
}
        
if ($AUTHRADIUSPOLH.Length -gt 0) {
    $Params = $null
    $Params = @{
        Hashtable = $AUTHRADIUSPOLH;
        Columns = "Policy","Expression","Action";
        Headers = "RADIUS Policy","Expression","LDAP Action";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
        }
    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;
}

WriteWordLine 0 0 " "
$Table = $null

#endregion Authentication Radius Policies

#region Authentication RADIUS
WriteWordLine 2 0 "NetScaler Radius authentication actions"
WriteWordLine 0 0 " "
$authactsradius = Get-vNetScalerObject -Container config -Object authenticationradiusaction;
If (!$authactsradius) {
  WriteWordLine 0 0 "There are no RADIUS authentication actions configured."
}
foreach ($authactradius in $authactsradius) {
    $ACTNAMERADIUS = $authactradius.name
    WriteWordLine 3 0 "Radius Authentication action $ACTNAMERADIUS";
    WriteWordLine 0 0 " "
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $RADUIUSCONFIG = @(
    @{ Description = "Description"; Value = "Configuration"; }
    @{ Description = "RADIUS Server IP"; Value = $authactradius.serverip; }
    @{ Description = "RADIUS Server Port"; Value = $authactradius.serverport; }
    @{ Description = "RADIUS Server Time-Out"; Value = $authactradius.authtimeout; }
    @{ Description = "Radius NAS IP"; Value = $authactradius.radnasip; }
    @{ Description = "IP Vendor ID"; Value = $authactradius.ipvendorid; }
    @{ Description = "Accounting"; Value = $authactradius.accounting; }
    @{ Description = "Calling Station ID"; Value = $authactradius.callingstationid; }
    );

    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = $RADUIUSCONFIG;
        Columns = "Description","Value";
        AutoFit = $wdAutoFitContent
        Format = -235; ## IB - Word constant for Light List Accent 5
    }
    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -NoGridLines -List;

	FindWordDocumentEnd;
	$TableRange = $Null
	$Table = $Null
    $selection.InsertNewPage()
}

WriteWordLine 0 0 " "
#endregion Authentication RADIUS

#endregion NetScaler Authentication

#region traffic management


#region NetScaler Content Switches
$Chapter++
$selection.InsertNewPage()
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Content Switching"

WriteWordLine 1 0 "NetScaler Content Switching"
WriteWordLine 0 0 " "
$csvservers = Get-vNetScalerObject -Object csvserver;

If (!$csvservers) {
    WriteWordLine 0 0 "No policies have been configured for this Content Switch"
}

foreach ($ContentSwitch in $csvservers) {
    $csvservername = $ContentSwitch.name
    WriteWordLine 2 0 "Content Switch $csvservername";
    WriteWordLine 0 0 " "
    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be about 400 characters wide!
    $Params = $null
    $Params = @{
        Hashtable = @{
            State = $ContentSwitch.curState;
            Protocol = $ContentSwitch.servicetype;
            Port = $ContentSwitch.port;
            IP = $ContentSwitch.ipv46;
            TrafficDomain = $ContentSwitch.td;
            CaseSensitive = $ContentSwitch.casesensitive;
            DownStateFlush = $ContentSwitch.downstateflush;
            ClientTimeOut = $ContentSwitch.clttimeout;
        }
        Columns = "State","Protocol","Port","IP","TrafficDomain","CaseSensitive","DownStateFlush","ClientTimeOut";
        Headers = "State","Protocol","Port","IP","Traffic Domain","Case Sensitive","Down State Flush","Client Time-Out";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
    }
    $Table = AddWordTable @Params -NoGridLines;

    FindWordDocumentEnd;
    WriteWordLine 0 0 " "
    $Table = $null
    $csvserverbindings = Get-vNetScalerObject -ResourceType csvserver_cspolicy_binding -Name $ContentSwitch.Name;

    WriteWordLine 3 0 "Policies"
    WriteWordLine 0 0 " "
    [System.Collections.Hashtable[]] $ContentSwitchPolicies = @();

        ## IB - Iterate over all Content Switch bindings (uses new function)
        foreach ($CSbinding in $csvserverbindings) {

            $cspolicy = Get-vNetScalerObject -ResourceType cspolicy -Name $CSbinding.policyname; 
            $csaction = Get-vNetScalerObject -ResourceType csaction -Name $cspolicy.action;

            ## IB - Add each Content Switch binding with a policyName to the array
            $ContentSwitchPolicies += @{
                    Policy = $cspolicy.policyname; 
                    Action = $cspolicy.action;
                    Priority = $cspolicy.priority;
                    Rule = $cspolicy.rule;
                    LB = $csaction.targetlbvserver;
                    }
                }      
        if ($ContentSwitchPolicies.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!

            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $ContentSwitchPolicies;
                Columns = "Policy","Action","Priority","Rule","LB";
                Headers = "Policy Name","Action","Priority","Rule","Target Load Balancer";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;

            FindWordDocumentEnd;
            $Table = $null
        } else {
            WriteWordLine 0 0 "No policies have been configured for this Content Switch"
    }
    FindWordDocumentEnd;

    WriteWordLine 0 0 " "
    WriteWordLine 3 0 "Advanced Configuration"

    [System.Collections.Hashtable[]] $AdvancedConfiguration = @(                
        @{ Description = "Description"; Value = "Configuration"; }
        @{ Description = "Apply AppFlow logging"; Value = $ContentSwitch.appflowlog; }
        @{ Description = "Enable or disable user authentication"; Value = $ContentSwitch.authentication; }
        @{ Description = "Enable state updates"; Value = $ContentSwitch.stateupdate; }
        @{ Description = "Route requests to the cache server"; Value = $ContentSwitch.cacheable; }
        @{ Description = "Precedence to use for policies"; Value = $ContentSwitch.precedence; }
        @{ Description = "URL Case sensitive"; Value = $ContentSwitch.casesensitive; }
    );

    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = $AdvancedConfiguration;
        Columns = "Description","Value";
        AutoFit = $wdAutoFitContent
        Format = -235; ## IB - Word constant for Light List Accent 5
    }
    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -NoGridLines -List;

	FindWordDocumentEnd;
	$TableRange = $Null
	$Table = $Null     
    
    FindWordDocumentEnd;
    $selection.InsertNewPage()

} # end if

#endregion NetScaler Content Switches

#region NetScaler Load Balancers
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Load Balancers"

WriteWordLine 1 0 "NetScaler Load Balancing"
WriteWordLine 0 0 " "
$lbvserverscount = Get-vNetScalerObjectCount -Container config -Object lbvserver;
$lbcount = $lbvserverscount.__count
$lbvservers = Get-vNetScalerObject -Container config -Object lbvserver;

if($lbvserverscount.__count -le 0) { WriteWordLine 0 0 "No Load Balancer has been configured"} else {
    
    ## IB - We no longer need to worrying about the number of columns and/or rows.
    ## IB - Need to create a counter of the current row index
    $CurrentRowIndex = 0;

    foreach ($LoadBalancer in $lbvservers) {
        $CurrentRowIndex++;
        $lbvservername = $LoadBalancer.name
        Write-Verbose "$(Get-Date): `tLoad Balancer $CurrentRowIndex/$($lbcount) $lbvservername"   
        WriteWordLine 2 0 "Load Balancer $lbvservername";

        $lbvserverbindings = Get-vNetScalerObject -ResourceType lbvserver_binding -Name $Loadbalancer.Name;

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be about 400 characters wide!
        $Params = $null
        $Params = @{
            Hashtable = @{
                ## IB - This table will only have 1 row so create the nested hashtable inline
                State = $LoadBalancer.downstateflush;
                Protocol = $LoadBalancer.servicetype;
                Port = $LoadBalancer.port;
                IP = $LoadBalancer.ipv46;
                Persistency = $LoadBalancer.persistencetype;
                TrafficDomain = $LoadBalancer.td;
                Method = $LoadBalancer.lbmethod;
                ClientTimeOut = $LoadBalancer.clttimeout;
            }
            Columns = "State","Protocol","Port","IP","Persistency","TrafficDomain","Method","ClientTimeOut";
            Headers = "State","Protocol","Port","IP","Persistency","Traffic Domain","Method","Client Time-Out";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;
        #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
 
        ##Services Table
        WriteWordLine 3 0 "Service and Service Group"
        WriteWordLine 0 0 " "

        #If No Services then record that there are none
        $lbserverservices = $lbvserverbindings.lbvserver_service_binding
        If (!$lbserverservices){
         WriteWordLine 0 0 "No Services are bound to the virtual server."
         WriteWordLine 0 0 " "
        }

        #If No Service Groups then record that there are none
        $lbserverservicegroups = $lbvserverbindings.lbvserver_servicegroup_binding
        If (!$lbserverservicegroups){
         WriteWordLine 0 0 "No Service Groups are bound to the virtual server."
         WriteWordLine 0 0 " "
        }

        
        If ($lbvserverbindings.lbvserver_service_binding.count -gt 0){

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $LoadBalancerServices = @();

            foreach ($lbvservicebind in $lbvserverbindings.lbvserver_service_binding) {
                    $LoadBalancerServices += @{ Service = $lbvservicebind.servicename;}
                }
        
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $LoadBalancerServices;
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;
            FindWordDocumentEnd;
            $Table = $null
            } else { WriteWordLine 0 0 "No Services are bound to the virtual server."}

        If ($lbvserverbindings.lbvserver_servicegroup_binding.count -gt 0){

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $LoadBalancerServices = @();

            foreach ($lbvservicegroupbind in $lbvserverbindings.lbvserver_servicegroup_binding) {
                    $LoadBalancerServices += @{ Servicegroup = $lbvservicegroupbind.servicename;}
                }
        
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $LoadBalancerServices;
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;
            FindWordDocumentEnd;
            $Table = $null
            } else { WriteWordLine 0 0 "No Service Groups are bound to the virtual server."}

    WriteWordLine 0 0 " "
    WriteWordLine 3 0 "Policies"
    WriteWordLine 0 0 " "
        If ($lbvserverbindings.lbvserver_responderpolicy_binding.count -gt 0){

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $LoadBalancerrespolicies = @();
            foreach ($lbrespolicy in $lbvserverbindings.lbvserver_responderpolicy_binding) {
                    $LoadBalancerrespolicies += @{ Responderpolicy = $lbrespolicy.policyname;}
                }
        
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $LoadBalancerrespolicies;
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;
            FindWordDocumentEnd;
            $Table = $null
            } else { 
              WriteWordLine 0 0 "No Responder Policies are bound to the virtual server."
              WriteWordLine 0 0 " "
            }

        If ($lbvserverbindings.lbvserver_rewritepolicy_binding.count -gt 0){

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $LoadBalancerrwpolicies = @();
            foreach ($lbrwpolicy in $lbvserverbindings.lbvserver_rewritepolicy_binding) {
                    $LoadBalancerrwpolicies += @{ Rewritepolicy = $lbrwpolicy.policyname;}
                }
        
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $LoadBalancerrwpolicies;
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;
            FindWordDocumentEnd;
            } else { 
              WriteWordLine 0 0 "No Rewrite Policies are bound to the virtual server."
              WriteWordLine 0 0 " "
            }

    FindWordDocumentEnd;
    WriteWordLine 0 0 " "
    WriteWordLine 3 0 "Redirect URL"
    WriteWordLine 0 0 " "
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $REDIRURLH = @();

    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be about 400 characters wide!
    $REDIRURLH += @{
            REDIRURL = $LoadBalancer.redirurl;
        }
    
    if ($REDIRURLH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $REDIRURLH;
            Columns = "REDIRURL";
            Headers = "Redirection URL";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
            }
        $Table = AddWordTable @Params -NoGridLines;
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
        $Table = $null

    FindWordDocumentEnd;
    } else {WriteWordLine 0 0 "No Redirection URL Configured"}
   
    ##Advanced Configuration   
    WriteWordLine 0 0 " "
    WriteWordLine 3 0 "Advanced Configuration"
    WriteWordLine 0 0 " "
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $AdvancedConfiguration = @(
        @{ Description = "Description"; Value = "Configuration"; }
        @{ Description = "Apply AppFlow logging"; Value = $LoadBalancer.appflowlog; }
        @{ Description = "Enable or disable user authentication"; Value = Test-StringPropertyOnOff $LoadBalancer "-Authentication"; }
        @{ Description = "Authentication virtual server name"; Value = $LoadBalancer.authnvsname; }
        @{ Description = "Name of the Authentication profile"; Value = $LoadBalance.authnprofile; }
        @{ Description = "User authentication with HTTP 401"; Value = $LoadBalancer.authn401; }
        @{ Description = "Time period a persistence session"; Value = $LoadBalancer.timeout; }
        @{ Description = "Backup persistence type"; Value = $LoadBalancer.persistencebackup; }
        @{ Description = "Time period a backup persistence session"; Value = $LoadBalancer.backuppersistencetimeout; }
        @{ Description = "Use priority queuing"; Value = $LoadBalancer.pq; }
        @{ Description = "Use SureConnect"; Value = $LoadBalancer.sc; }
        @{ Description = "Use network address translation"; Value = $LoadBalancer.rtspnat; }
        @{ Description = "Use Layer 2 parameter"; Value = $LoadBalancer.l2conn; }
        @{ Description = "How the NetScaler appliance responds to ping requests"; Value = $LoadBalancer.icmpvsrresponse; }
        @{ Description = "Route cacheable requests to a cache redirection server"; Value = $LoadBalancer.cacheable; }
    );

    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = $AdvancedConfiguration;
        Columns = "Description","Value";
        AutoFit = $wdAutoFitContent
        Format = -235; ## IB - Word constant for Light List Accent 5
    }
    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -NoGridLines -List;
    ## IB - Set the header background and bold font
    #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

	FindWordDocumentEnd;
	$TableRange = $Null
	$Table = $Null

    $selection.InsertNewPage()
    }
}
#endregion NetScaler Load Balancers

#region NetScaler Cache Redirection
$Chapter++

Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Cache Redirection"
WriteWordLine 1 0 "NetScaler Cache Redirection"
WriteWordLine 0 0 " "
$crservercounter = Get-vNetScalerObjectCount -Container config -Object crvserver; 
$crservercount = $crservercounter.__count
$crservers = Get-vNetScalerObject -Container config -Object crvserver;

if($crservercounter.__count -le 0) { WriteWordLine 0 0 "No Cache Redirection has been configured"} else {
    $CurrentRowIndex = 0;

    foreach ($crserver in $crservers) {
        $CurrentRowIndex++;
        $crname = $crserver.name
        Write-Verbose "$(Get-Date): `tCache Redirection $CurrentRowIndex/$($crservercount) $crname"     
        WriteWordLine 2 0 "Cache Redirection Server $crname";

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be about 400 characters wide!
        $Params = $null
        $Params = @{
            Hashtable = @{
                ## IB - This table will only have 1 row so create the nested hashtable inline
                PROT = $crserver.servicetype;
                IP = $crserver.ip;
                Port = $crserver.port;
                CACHETYPE = $crserver.cachetype;
                REDIRECT = $crserver.redirect;
                CLTTIEMOUT = $crserver.clttimeout;
            }
            Columns = "PROT","IP","PORT","CACHETYPE","REDIRECT","CLTTIEMOUT";
            Headers = "PROT","IP","Port","CACHETYPE","REDIRECT","CLTTIEMOUT";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;

        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
        $Table = $null
        }
    }
$selection.InsertNewPage()

#endregion NetScaler Cache Redirection

#region NetScaler Services
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Services"

FindWordDocumentEnd;

WriteWordLine 1 0 "NetScaler Services"
WriteWordLine 0 0 " "
$servicescounter = Get-vNetScalerObjectCount -Container config -Object service; 
$servicescount = $servicescounter.__count
$services = Get-vNetScalerObject -Container config -Object service;

if($servicescounter.__count -le 0) { WriteWordLine 0 0 "No Services have been configured"} else {
    $CurrentRowIndex = 0;

    foreach ($Service in $Services) {

        $CurrentRowIndex++;
        $servicename = $Service.name
        Write-Verbose "$(Get-Date): `tService $CurrentRowIndex/$($servicescount) $servicename"     
        WriteWordLine 2 0 "Service $servicename"
        WriteWordLine 0 0 " "

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be about 400 characters wide!
        $Params = $null
        $Params = @{
            Hashtable = @{
                ## IB - This table will only have 1 row so create the nested hashtable inline
                Protocol = $Service.servicetype;
                Port = $Service.port;
                TD = $Service.td;
                GSLB = $Service.gslb;
                MaximumClients = $Service.maxclient;
                MaximumRequests = $Service.maxreq;
            }
            Columns = "Protocol","Port","TD","GSLB","MaximumClients","MaximumRequests";
            Headers = "Protocol","Port","Traffic Domain","GSLB","Maximum Clients","Maximum Requests";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;

        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
        
        WriteWordLine 3 0 "Monitor"
        WriteWordLine 0 0 " "
        ## Query for a service monitor binding. NOTE: Can access the .ServiceName property with '$lbvserverbinding.ServiceName'
        $svcmonitorbinds = Get-vNetScalerObject -ResourceType service_lbmonitor_binding -Name $Service.Name;

        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $ServiceMonitors = @();

        ## IB - Iterate over all Service bindings (uses new function)
        foreach ($SVCBind in $svcmonitorbinds) {
            $ServiceMonitors += @{ Monitor = $SVCBind.monitor_name; }
        } # end foreach

        if ($ServiceMonitors.Length -gt 0) {
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $ServiceMonitors;                   
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;
                
            FindWordDocumentEnd;
            $Table = $null
        } else {
            WriteWordLine 0 0 "No Monitors have been configured for this Service"
    } # end if

        WriteWordLine 0 0 " "
        WriteWordLine 3 0 "Advanced Configuration"
        WriteWordLine 0 0 " "
        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $AdvancedConfiguration = @(
            @{ Description = "Description"; Value = "Configuration"; }
			@{ Description = "Cache Type"; Value = $service.cachetype; }
			@{ Description = "Maximum Client Requests"; Value = $service.maxclient ; }
			@{ Description = "Monitor health of this service"; Value = $service.healthmonitor ; }
			@{ Description = "Maximum Requests"; Value = $service.maxreq; }
			@{ Description = "Use Transparent Cache"; Value = $service.cacheable ; }
			@{ Description = "Insert the Client IP header"; Value = $service.cip  ; }
			@{ Description = "Name for the HTTP header"; Value = $service.cipheader ; }
			@{ Description = "Use Source IP"; Value = $service.usip; }
            @{ Description = "Path Monitoring"; Value = $service.pathmonitor ; }
			@{ Description = "Individual Path monitoring"; Value = $service.pathmonitorindv ; }
			@{ Description = "Use the proxy port"; Value = $service.useproxyport ; }
			@{ Description = "SureConnect"; Value = $service.sc ; }
			@{ Description = "Surge protection"; Value = $service.sp ; }
			@{ Description = "RTSP session ID mapping"; Value = $service.rtspsessionidremap ; }
			@{ Description = "Client Time-Out"; Value = $service.clttimeout ; }
			@{ Description = "Server Time-Out"; Value = $service.svrtimeout; }
			@{ Description = "Unique identifier for the service"; Value = $service.customserverid; }
			@{ Description = "Enable client keep-alive"; Value = $service.cka; }
			@{ Description = "Enable TCP buffering"; Value = $service.tcpb ; }
            @{ Description = "Enable compression"; Value = $service.cmp }
			@{ Description = "Maximum bandwidth, in Kbps"; Value = $service.maxbandwidth; }
			@{ Description = "Sum of weights of the monitors"; Value = $service.monthreshold ; }
			@{ Description = "Initial state of the service"; Value = $service.svrstate ; }
			@{ Description = "Perform delayed clean-up"; Value = $service.downstateflush ; }
			@{ Description = "Logging of AppFlow information"; Value = $service.appflowlog; }
        );

        ## IB - Create the parameters to pass to the AddWordTable function
        $Params = $null
        $Params = @{
            Hashtable = $AdvancedConfiguration;
            Columns = "Description","Value";
            AutoFit = $wdAutoFitContent
            Format = -235; ## IB - Word constant for Light List Accent 5
        }
        ## IB - Add the table to the document, splatting the parameters
        $Table = AddWordTable @Params -NoGridLines -List;

		FindWordDocumentEnd;
		$TableRange = $Null
		$Table = $Null
        WriteWordLine 0 0 " "

        $selection.InsertNewPage() 
        }
   }

#endregion NetScaler Services

#region NetScaler Service Groups
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Service Groups"

FindWordDocumentEnd;

WriteWordLine 1 0 "NetScaler Service Groups"
WriteWordLine 0 0 " "
$servicegroupscounter = Get-vNetScalerObjectCount -Container config -Object servicegroup; 
$servicegroupscount = $servicegroupscounter.__count
$servicegroups = Get-vNetScalerObject -Container config -Object servicegroup;

if($servicegroupscounter.__count -le 0) { WriteWordLine 0 0 "No Service Groups have been configured"} else {
    $CurrentRowIndex = 0;

    foreach ($Servicegroup in $servicegroups) {
        $CurrentRowIndex++;
        $servicename = $Servicegroup.servicegroupname
        Write-Verbose "$(Get-Date): `tService $CurrentRowIndex/$($servicegroupscount) $servicename"     
        WriteWordLine 2 0 "Service Group $servicename"
        WriteWordLine 0 0 " "
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be about 400 characters wide!
        $Params = $null
        $Params = @{
            Hashtable = @{
                ## IB - This table will only have 1 row so create the nested hashtable inline
                Protocol = $Servicegroup.servicetype;
                Port = $Servicegroup.port;
                MaximumClients = $Servicegroup.maxclient;
                MaximumRequests = $Servicegroup.maxreq;
            }
            Columns = "Protocol","Port","TD","MaximumClients","MaximumRequests";
            Headers = "Protocol","Port","Traffic Domain","Maximum Clients","Maximum Requests";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;

        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
               
        WriteWordLine 3 0 "Servers"
        WriteWordLine 0 0 " "
        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $ServiceGroupServers = @();

        $servicegroupbinds = Get-vNetScalerObject -ResourceType servicegroup_servicegroupmember_binding -Name $servicegroup.servicegroupname;
        foreach ($servicegroupbind in $servicegroupbinds) {
            
            foreach ($svcgroupserver in $servicegroupbind) { 
                $ServiceGroupServers += @{ Server = $svcgroupserver.servername; }
            }
        }

        if ($ServiceGroupServers.Length -gt 0) {
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $ServiceGroupServers;                   
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;
                
            FindWordDocumentEnd;
        } else {
            WriteWordLine 0 0 "No Servers have been configured for this Service Group"
        }   

        WriteWordLine 0 0 " "
        $Table = $null

        WriteWordLine 3 0 "Monitor"
        WriteWordLine 0 0 " "
        $svcmonitorbinds = Get-vNetScalerObject -ResourceType servicegroup_lbmonitor_binding -Name $Service.servicegroupname;

        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $ServiceMonitors = @();

        ## IB - Iterate over all Service bindings (uses new function)
        foreach ($SVCBind in $svcmonitorbinds) {
            $ServiceMonitors += @{ Monitor = $SVCBind.monitor_name; }
        } # end foreach

        if ($ServiceMonitors.Length -gt 0) {
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $ServiceMonitors;                   
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;
                
            FindWordDocumentEnd;
        } else {
            WriteWordLine 0 0 "No Monitor has been configured for this Service"
    } # end if

        WriteWordLine 0 0 " "
        WriteWordLine 3 0 "Advanced Configuration"
        WriteWordLine 0 0 " "
        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $AdvancedConfiguration = @(
            @{ Description = "Description"; Value = "Configuration"; }
			@{ Description = "Cache Type"; Value = $service.cachetype; }
			@{ Description = "Maximum Client Requests"; Value = $service.maxclient ; }
			@{ Description = "Monitor health of this service"; Value = $service.healthmonitor ; }
			@{ Description = "Maximum Requests"; Value = $service.maxreq; }
			@{ Description = "Use Transparent Cache"; Value = $service.cacheable ; }
			@{ Description = "Insert the Client IP header"; Value = $service.cip  ; }
			@{ Description = "Name for the HTTP header"; Value = $service.cipheader ; }
			@{ Description = "Use Source IP"; Value = $service.usip; }
            @{ Description = "Path Monitoring"; Value = $service.pathmonitor ; }
			@{ Description = "Individual Path monitoring"; Value = $service.pathmonitorindv ; }
			@{ Description = "Use the proxy port"; Value = $service.useproxyport ; }
			@{ Description = "SureConnect"; Value = $service.sc ; }
			@{ Description = "Surge protection"; Value = $service.sp ; }
			@{ Description = "RTSP session ID mapping"; Value = $service.rtspsessionidremap ; }
			@{ Description = "Client Time-Out"; Value = $service.clttimeout ; }
			@{ Description = "Server Time-Out"; Value = $service.svrtimeout; }
			@{ Description = "Unique identifier for the service"; Value = $service.customserverid; }
			@{ Description = "Enable client keep-alive"; Value = $service.cka; }
			@{ Description = "Enable TCP buffering"; Value = $service.tcpb ; }
            @{ Description = "Enable compression"; Value = $service.cmp }
			@{ Description = "Maximum bandwidth, in Kbps"; Value = $service.maxbandwidth; }
			@{ Description = "Sum of weights of the monitors"; Value = $service.monthreshold ; }
			@{ Description = "Initial state of the service"; Value = $service.servicegroupeffectivestate ; }
			@{ Description = "Perform delayed clean-up"; Value = $service.downstateflush ; }
			@{ Description = "Logging of AppFlow information"; Value = $service.appflowlog; }
        );

        ## IB - Create the parameters to pass to the AddWordTable function
        $Params = $null
        $Params = @{
            Hashtable = $AdvancedConfiguration;
            Columns = "Description","Value";
            AutoFit = $wdAutoFitContent
            Format = -235; ## IB - Word constant for Light List Accent 5
        }
        ## IB - Add the table to the document, splatting the parameters
        $Table = AddWordTable @Params -NoGridLines -List;

		FindWordDocumentEnd;
		$TableRange = $Null
		$Table = $Null
        WriteWordLine 0 0 " "

        $selection.InsertNewPage() 
        }
   }

#endregion NetScaler Service Groups

#region NetScaler Servers
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Servers"
WriteWordLine 1 0 "NetScaler Servers"
WriteWordLine 0 0 " "
$servercounter = Get-vNetScalerObjectCount -Container config -Object service; 
$servercount = $servercounter.__count
$servers = Get-vNetScalerObject -Container config -Object server;

if($servercounter.__count -le 0) { WriteWordLine 0 0 "No Server has been configured"} else {

    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $ServersH = @();

    foreach ($Server in $Servers) {
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be about 400 characters wide!
        $ServersH += @{
                Server = $server.name;
                IP = $Server.ipaddress;
                TD = $server.td;
                STATE = $server.state;
                #COMMENT = $server.;
            }
        }
        if ($ServersH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $ServersH;
                Columns = "Server","IP","TD","STATE";
                Headers = "Server","IP Address","Traffic Domain","State";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            $Table = $null
            }
        }

$selection.InsertNewPage()    
#endregion NetScaler Servers

#region Global Server Load Balancing

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Global Server Load Balancing"
WriteWordLine 1 0 "NetScaler Global Server Load Balancing"
WriteWordLine 0 0 " "
#region GSLB Parameters

WriteWordLine 2 0 "GSLB Parameters"
Write-Verbose "$(Get-Date): `tGSLB Parameters"
WriteWordLine 0 0 " "
$gslbparameters = Get-vNetScalerObject -Container config -Object gslbparameter;

[System.Collections.Hashtable[]] $GSLBParameterDetails = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "LDNS Entry Timeout"; Column2 = $gslbparameters.ldnsentrytimeout; }
    @{ Column1 = "RTT Tolerance"; Column2 = $gslbparameters.rtttolerance; }
    @{ Column1 = "IPv4 LDNS Mask"; Column2 = $gslbparameters.ldnsmask; }
    @{ Column1 = "IPv6 LDNS Mask"; Column2 = $gslbparameters.v6ldnsmasklen; }
    @{ Column1 = "LDNS Probe Order"; Column2 = $gslbparameters.ldnsprobeorder; }
    @{ Column1 = "Drop LDNS Requests"; Column2 = $gslbparameters.dropldnsreq; }
 
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $GSLBParameterDetails;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List ;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
$Table = $null

#endregion GSLB Parameters

#region GSLB vServers

WriteWordLine 2 0 "GSLB Virtual Servers"
Write-Verbose "$(Get-Date): `tGSLB Virtual Servers"
WriteWordLine 0 0 " "
$gslbvservercounter = Get-vNetScalerObjectCount -Container config -Object gslbvserver; 
$gslbvservercount = $gslbvservercounter.__count
$gslbvservers = Get-vNetScalerObject -Container config -Object gslbvserver

if($gslbvservercount -le 0) { WriteWordLine 0 0 "No GSLB Virtual Servers have been configured"} else {

foreach ($gslbvserver in $gslbvservers) {

$gslbvservername = $gslbvserver.name

WriteWordLine 3 0 "GSLB Virtual Server: $gslbvservername"
WriteWordLine 0 0 " "
## IB - Create an array of hashtables to store our columns.
## IB - about column names as we'll utilise a -List(view)!
[System.Collections.Hashtable[]] $GSLBvServerDetails = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "Service Type"; Column2 = $gslbvserver.servicetype; }
    @{ Column1 = "State"; Column2 = $gslbvserver.state; }
    @{ Column1 = "Status"; Column2 = $gslbvserver.status; }
    @{ Column1 = "IP Type"; Column2 = $gslbvserver.iptype; }
    @{ Column1 = "DNS Record Type"; Column2 = $gslbvserver.dnsrecordtype; }
    @{ Column1 = "Persistence Type"; Column2 = $gslbvserver.persistencetype; }
    @{ Column1 = "Persistence ID"; Column2 = $gslbvserver.persistenceid; }
    @{ Column1 = "Load Balancing Method"; Column2 = $gslbvserver.lbmethod; }
    @{ Column1 = "Backup Load Balancing Method"; Column2 = $gslbvserver.backuplbmethod; }
    @{ Column1 = "Tolerance"; Column2 = $gslbvserver.tolerance; }
    @{ Column1 = "Timeout"; Column2 = $gslbvserver.timeout; }
    @{ Column1 = "Netmask"; Column2 = $gslbvserver.netmask; }
    @{ Column1 = "IPv6 Netmask"; Column2 = $gslbvserver.v6netmasklen; }
    @{ Column1 = "Persistence mask"; Column2 = $gslbvserver.persistmask; }
    @{ Column1 = "IPv6 Persistence mask"; Column2 = $gslbvserver.v6persistmasklen; }
    @{ Column1 = "Bound Services"; Column2 = $gslbvserver.servicename; }
    @{ Column1 = "Weight"; Column2 = $gslbvserver.weight; }
    @{ Column1 = "Domain Name"; Column2 = $gslbvserver.domainname; }
    @{ Column1 = "TTL"; Column2 = $gslbvserver.ttl; }
    @{ Column1 = "Backup IP Address"; Column2 = $gslbvserver.backupip; }
    @{ Column1 = "Cookie Domain"; Column2 = $gslbvserver.cookiedomain; }
    @{ Column1 = "Cookie Timeout"; Column2 = $gslbvserver.cookietimeout; }
    @{ Column1 = "Domain TTL"; Column2 = $gslbvserver.sitedomainttl; }
    @{ Column1 = "Backup vServer"; Column2 = $gslbvserver.backupvserver; }
    @{ Column1 = "Disable Primary when down"; Column2 = $gslbvserver.disableprimaryondown; }
    @{ Column1 = "Dynamic Weight"; Column2 = $gslbvserver.dynamicweight; }
    @{ Column1 = "ISC Weight"; Column2 = $gslbvserver.iscweight; }
    @{ Column1 = "Site Persistence"; Column2 = $gslbvserver.sitepersistence; }
    @{ Column1 = "Comment"; Column2 = $gslbvserver.comment; }
    @{ Column1 = "vServer Bind Service IP"; Column2 = $gslbvserver.vsvrbindsvcip; }
    @{ Column1 = "vServer Bind Service Port"; Column2 = $gslbvserver.vsvrbindsvcport; }
    
    #TODO: Spillover Policies

);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $GSLBvServerDetails;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List ;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
$Table = $null


#region GSLB vServer Service Bindings

WriteWordLine 4 0 "Services"
WriteWordLine 0 0 " "

$GSLBServiceBinds = Get-vNetScalerObject -ResourceType gslbvserver_gslbservice_binding -Name $gslbvservername;



        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $GSLBServices = @();

        ## IB - Iterate over all Service bindings (uses new function)
        foreach ($GSLBServiceBind in $GSLBServiceBinds) {
            $GSLBServices += @{ ServiceName = $GSLBServiceBind.servicename; Weight = $GSLBServiceBind.weight;}
        } # end foreach

        if ($GSLBServices.Length -gt 0) {
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $GSLBServices; 
                Columns = "ServiceName","Weight";
                Headers =  "Service Name", "Service Weight";                  
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            FindWordDocumentEnd;
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;


        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
        $Table = $null
        } else {
          WriteWordLine 0 0 "No GSLB Services have been bound"
        }
        

#endregion GSLB vServer Service Bindings
#region GSLB Domain Bindings

WriteWordLine 4 0 "Domain Bindings"
WriteWordLine 0 0 " "

        $GSLBDomainBinds = Get-vNetScalerObject -ResourceType gslbvserver_domain_binding -Name $gslbvservername;



        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $GSLBDomains = @();

        ## IB - Iterate over all Service bindings (uses new function)
        foreach ($GSLBDomainBind in $GSLBDomainBinds) {
            $GSLBDomains += @{ DomainName = $GSLBDomainBind.domainname; TTL = $GSLBDomainBind.ttl; CookieDomain = $GSLBDomainBind.cookie_domain; CookieTimeout = $GSLBDomainBind.cookietimeout;}
        } # end foreach
        
        if ($GSLBDomains.Length -gt 0) {
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $GSLBDomains; 
                Columns = "DomainName","TTL","CookieDomain","CookieTimeout";
                Headers = "Domain Name", "TTL", "Cookie Domain", "Cookie Timeout";                  
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            FindWordDocumentEnd;
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;


        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
        $Table = $null
        } else {
          WriteWordLine 0 0 "No GSLB Domains have been bound"
          WriteWordLine 0 0 " "
        }
        

#endregion GSLB Domain Bindings

}
}
#endregion GSLB vServers

#region GSLB Services

WriteWordLine 2 0 "GSLB Services"
Write-Verbose "$(Get-Date): `tGSLB Services"
WriteWordLine 0 0 " "
$gslbservicecounter = Get-vNetScalerObjectCount -Container config -Object gslbservice; 
$gslbservicecount = $gslbservicecounter.__count
$gslbservicesall = Get-vNetScalerObject -Container config -Object gslbservice;

if($gslbservicecount -le 0) { WriteWordLine 0 0 "No GSLB Services have been configured"} else {

foreach ($gslbservice in $gslbservicesall) {

$gslbservicename = $gslbservice.servicename

WriteWordLine 3 0 "NetScaler GSLB Service: $gslbservicename"
WriteWordLine 0 0 " "

## IB - Create an array of hashtables to store our columns.
## IB - about column names as we'll utilise a -List(view)!
[System.Collections.Hashtable[]] $GSLBServiceDetails = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "Service Location"; Column2 = $gslbservice.gslb; }
    @{ Column1 = "GSLB Site"; Column2 = $gslbservice.sitename; }
    @{ Column1 = "IP Address"; Column2 = $gslbservice.ipaddress; }
    @{ Column1 = "IP"; Column2 = $gslbservice.ip; }
    @{ Column1 = "Server Name"; Column2 = $gslbservice.servername; }
    @{ Column1 = "Port"; Column2 = $gslbservice.port; }
    @{ Column1 = "Public IP"; Column2 = $gslbservice.publicip; }
    @{ Column1 = "Public Port"; Column2 = $gslbservice.publicport; }
    @{ Column1 = "Max Clients"; Column2 = $gslbservice.maxclient; }
    @{ Column1 = "Max AAA Users"; Column2 = $gslbservice.maxaaausers; }
    @{ Column1 = "Monitor Threshold"; Column2 = $gslbservice.monthreshold; }
    @{ Column1 = "State"; Column2 = $gslbservice.state; }
    @{ Column1 = "Insert Client IP"; Column2 = $gslbservice.cip; }
    @{ Column1 = "Client IP Header"; Column2 = $gslbservice.cipheader; }
    @{ Column1 = "Site Persistence"; Column2 = $gslbservice.sitepersistence; }
    @{ Column1 = "Site Prefix"; Column2 = $gslbservice.siteprefix; }
    @{ Column1 = "Client Timeout"; Column2 = $gslbservice.clttimeout; }
    @{ Column1 = "Server Timeout"; Column2 = $gslbservice.svrtimeout; }
    @{ Column1 = "Preferred Location"; Column2 = $gslbservice.preferredlocation; }
    @{ Column1 = "Maximum bandwidth, in Kbps"; Column2 = $gslbservice.maxbandwidth; }
    @{ Column1 = "Flush active transactions for DOWN service"; Column2 = $gslbservice.downstateflush; }
    @{ Column1 = "CNAME Entry"; Column2 = $gslbservice.cnameentry; }
    @{ Column1 = "Comment"; Column2 = $gslbservice.comment; }
  


);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $GSLBServiceDetails;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List ;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
$Table = $null

#region GSLB Service Monitors


WriteWordLine 4 0 "Monitors"
WriteWordLine 0 0 " "

$GSLBMonitorBinds = Get-vNetScalerObject -ResourceType gslbservice_lbmonitor_binding -Name $gslbservicename;



        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $GSLBMonitors = @();

        ## IB - Iterate over all Service bindings (uses new function)
        foreach ($GSLBMonitorBind in $GSLBMonitorBinds) {
            $GSLBServices += @{ MonitorName = $GSLBMonitorBind.monitor_name; Weight = $GSLBMonitorBind.weight;}
        } # end foreach

        if ($GSLBMonitors.Length -gt 0) {
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $GSLBMonitors; 
                Columns = "MonitorName","Weight";
                Headers =  "Monitor Name", "Weight";                  
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            FindWordDocumentEnd;
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;


        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
        $Table = $null
        } else {
          WriteWordLine 0 0 "No explicit monitors have been bound to the service"
          WriteWordLine 0 0 " "
        }
        

#endregion GSLB Service Monitors

#region GSLB Service DNS View


WriteWordLine 4 0 "DNS Views"
WriteWordLine 0 0 " "

$GSLBDNSViewBinds = Get-vNetScalerObject -ResourceType gslbservice_dnsview_binding -Name $gslbservicename;



        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $GSLBDNSViews = @();

        ## IB - Iterate over all Service bindings (uses new function)
        foreach ($GSLBDNSViewBind in $GSLBDNSViewBinds) {
            $GSLBDNSViews += @{ ViewName = $GSLBDNSViewBind.viewname; ViewIP = $GSLBDNSViewBind.viewip;}
        } # end foreach

        if ($GSLBMonitors.Length -gt 0) {
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $GSLBMonitors; 
                Columns = "ViewName","ViewIP";
                Headers =  "View Name", "View IP";                  
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            FindWordDocumentEnd;
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;


        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
        $Table = $null
        } else {
          WriteWordLine 0 0 "No DNS Views have been bound to the service"
          WriteWordLine 0 0 " "
        }
        

#endregion GSLB Service DNS View

} #end foreach

} #endif


#endregion GSLB Services

#region GSLB Sites

WriteWordLine 2 0 "GSLB Sites"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `tGSLB Sites"

$gslbsitecounter = Get-vNetScalerObjectCount -Container config -Object gslbsite; 
$gslbsitecount = $gslbsitecounter.__count
$gslbsitesall = Get-vNetScalerObject -Container config -Object gslbsite;

if($gslbsitecount -le 0) { WriteWordLine 0 0 "No GSLB Sites have been configured"} else {

foreach ($gslbsite in $gslbsitesall) {

$gslbsitename = $gslbsite.sitename


WriteWordLine 3 0 "NetScaler GSLB Site: $gslbsitename"
WriteWordLine 0 0 " "

## IB - Create an array of hashtables to store our columns.
## IB - about column names as we'll utilise a -List(view)!
[System.Collections.Hashtable[]] $GSLBSiteDetails = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "Site Type"; Column2 = $gslbsite.sitetype; }
    @{ Column1 = "Site IP Address"; Column2 = $gslbsite.siteipaddress; }
    @{ Column1 = "Site Public IP"; Column2 = $gslbsite.publicip; }
    @{ Column1 = "Metric Exchange"; Column2 = $gslbsite.metricexchange; }
    @{ Column1 = "Persistence Exchange"; Column2 = $gslbsite.persistencemepstatus; }
    @{ Column1 = "Network Metric Exchange"; Column2 = $gslbsite.nwmetricexchange; }
    @{ Column1 = "Session Exchange"; Column2 = $gslbsite.sessionexchange; }
    @{ Column1 = "Parent Site"; Column2 = $gslbsite.parentsite; }
    @{ Column1 = "Cluster IP"; Column2 = $gslbsite.clip; }
    @{ Column1 = "Public Cluster IP"; Column2 = $gslbsite.publicclip; }


);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $GSLBSiteDetails;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List ;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
$Table = $null
} #end foreach
} #end if

#endregion GSLB Sites

$selection.InsertNewPage()

#endregion Global Server Load Balancing

#region NetScaler SSL
WriteWordLine 1 0 "NetScaler SSL"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `tNetScaler SSL"

#region SSL Certificates
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters SSL Certificates"

$selection.InsertNewPage()

WriteWordLine 2 0 "SSL Certificates"
WriteWordLine 0 0 " "
$sslcerts = Get-vNetScalerObject -Object sslcertkey;

## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $SSLCERTSH = @();

foreach ($sslcert in $sslcerts) {

    $sslcert1 = Get-vNetScalerObject -ResourceType sslcertkey -Name $sslcert.certkey;
    $subject = $sslcert1.subject
    $subject1 = $subject.Split(',')[-1]
    $sslfqdn = ($subject1 -replace 'CN=', '')
    
$sslcertname = $sslcert.certkey
WriteWordLine 3 0 "SSL Certificate: $sslcertname"
WriteWordLine 0 0 " "

## IB - Create an array of hashtables to store our columns.
## IB - about column names as we'll utilise a -List(view)!
[System.Collections.Hashtable[]] $SSLCertDetails = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "Name"; Column2 = $sslcert.certkey; }
    @{ Column1 = "FQDN"; Column2 = $sslfqdn; }
    @{ Column1 = "Issuer"; Column2 = $sslcert.issuer; }
    @{ Column1 = "Certificate File"; Column2 = $sslcert.cert; }
    @{ Column1 = "Key File"; Column2 = $sslcert.key; }
    @{ Column1 = "Key Size"; Column2 = $sslcert.publickeysize; }
    @{ Column1 = "Valid From"; Column2 = $sslcert.clientcertnotbefore; }
    @{ Column1 = "Valid Until"; Column2 = $sslcert.clientcertnotafter; }
    @{ Column1 = "Days to Expiry"; Column2 = $sslcert.daystoexpiration; }
    @{ Column1 = "Certificate Type"; Column2 = $sslcert.certificatetype; }
    @{ Column1 = "Linked Certificate"; Column2 = $sslcert.linkcertkeyname; }


);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $SSLCertDetails;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List ;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
$Table = $null
}
        

#endregion SSL Certificates


#region SSL Ciphers
WriteWordLine 2 0 "SSL Ciphers"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `tSSL Ciphers"

$SSLCiphers = Get-vNetScalerObject -Container config -Object sslcipher;



        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $SSLCIPHERH = @();

        ## IB - Iterate over all Service bindings (uses new function)
        foreach ($SSLCipher in $SSLCiphers) {
            $SSLCIPHERH += @{ GRPNAME = $SSLCipher.ciphergroupname; DESC = $SSLCipher.description;}
        } # end foreach

        if ($SSLCIPHERH.Length -gt 0) {
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $SSLCIPHERH; 
                Columns = "GRPNAME","DESC";
                Headers =  "Cipher/Group Name", "Description";                  
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            FindWordDocumentEnd;
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;


        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
        $Table = $null
        } else {
          WriteWordLine 0 0 "No SSL Ciphers were returned."
          WriteWordLine 0 0 " "
        }



#endregion SSL Ciphers

#region SSL Services
WriteWordLine 2 0 "SSL Services"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `tSSL Services"

$SSLServices = Get-vNetScalerObject -Container config -Object sslservice;


Foreach ($SSLService in $SSLServices) {

$sslservicename = $sslservice.servicename

WriteWordLine 3 0 "SSL Service: $sslservicename"
WriteWordLine 0 0 " "
[System.Collections.Hashtable[]] $SSLSERVICEH = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "Diffe-Hellman Key Exchange"; Column2 = $SSLService.dh; }
    @{ Column1 = "Diffe-Hellman Key File"; Column2 = $SSLService.dhfile; }
    @{ Column1 = "Diffe-Hellman Refresh Count"; Column2 = $SSLService.dhcount; }
    @{ Column1 = "Enable DH Key Expire Size Limit"; Column2 = $SSLService.dhkeyexpsizelimit; }
    @{ Column1 = "Enable Ephemeral RSA"; Column2 = $SSLService.ersa; }
    @{ Column1 = "Ephemeral RSA Refresh Count"; Column2 = $SSLService.ersacount; }
    @{ Column1 = "Allow session re-use"; Column2 = $SSLService.sessreuse; }
    @{ Column1 = "Session Time-out"; Column2 = $SSLService.sesstimeout; }
    @{ Column1 = "Enable Cipher Redirect"; Column2 = $SSLService.cipherredirect; }
    @{ Column1 = "Cipher Redirect URL"; Column2 = $SSLService.cipherurl; }
    @{ Column1 = "SSLv2 Redirect"; Column2 = $SSLService.sslv2redirect; }
    @{ Column1 = "SSLv2 Redirect URL"; Column2 = $SSLService.sslv2url; }
    @{ Column1 = "Enable Client Authentication"; Column2 = $SSLService.clientauth; }
    @{ Column1 = "Client Certificates"; Column2 = $SSLService.clientcert; }
    @{ Column1 = "SSL Redirect"; Column2 = $SSLService.sslredirect; }
    @{ Column1 = "SSL 2"; Column2 = $SSLService.ssl2; }
    @{ Column1 = "SSL 3"; Column2 = $SSLService.ssl3; }
    @{ Column1 = "TLS 1"; Column2 = $SSLService.tls1; }
    @{ Column1 = "TLS 1.1"; Column2 = $SSLService.tls11; }
    @{ Column1 = "TLS 1.2"; Column2 = $SSLService.tls12; }
    @{ Column1 = "Server Name Indication (SNI)"; Column2 = $SSLService.snienable; }
    @{ Column1 = "Enable Server Authentication"; Column2 = $SSLService.serverauth; }
    @{ Column1 = "Common Name"; Column2 = $SSLService.commonname; }
    @{ Column1 = "PUSH Encryption Trigger"; Column2 = $SSLService.pushenctrigger; }
    @{ Column1 = "Send Close-Notify"; Column2 = $SSLService.sendclosenotify; }
    @{ Column1 = "DTLS Profile"; Column2 = $SSLService.dtlsprofilename; }
    @{ Column1 = "SSL Profile"; Column2 = $SSLService.sslprofile; }

);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $SSLSERVICEH;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
$Table = $null
} #end foreach


#endregion SSL Services

#region SSL Service Groups

WriteWordLine 2 0 "SSL Service Groups"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `tSSL Service Groups"


$SSLServiceGrps = Get-vNetScalerObject -Container config -Object sslservicegroup;

If ($SSLServiceGrps) {

Foreach ($SSLServiceGrp in $SSLServiceGrps) {

$sslservicegrpname = $sslserviceGrp.servicegroupname

WriteWordLine 3 0 "SSL Service Group: $sslservicegrpname"
WriteWordLine 0 0 " "
[System.Collections.Hashtable[]] $SSLSERVICEGRPH = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "Diffe-Hellman Key Exchange"; Column2 = $SSLServiceGrp.dh; }
    @{ Column1 = "Diffe-Hellman Key File"; Column2 = $SSLServiceGrp.dhfile; }
    @{ Column1 = "Diffe-Hellman Refresh Count"; Column2 = $SSLServiceGrp.dhcount; }
    @{ Column1 = "Enable DH Key Expire Size Limit"; Column2 = $SSLServiceGrp.dhkeyexpsizelimit; }
    @{ Column1 = "Enable Ephemeral RSA"; Column2 = $SSLServiceGrp.ersa; }
    @{ Column1 = "Ephemeral RSA Refresh Count"; Column2 = $SSLServiceGrp.ersacount; }
    @{ Column1 = "Allow session re-use"; Column2 = $SSLServiceGrp.sessreuse; }
    @{ Column1 = "Session Time-out"; Column2 = $SSLServiceGrp.sesstimeout; }
    @{ Column1 = "Enable Cipher Redirect"; Column2 = $SSLServiceGrp.cipherredirect; }
    @{ Column1 = "Cipher Redirect URL"; Column2 = $SSLServiceGrp.cipherurl; }
    @{ Column1 = "SSLv2 Redirect"; Column2 = $SSLServiceGrp.sslv2redirect; }
    @{ Column1 = "SSLv2 Redirect URL"; Column2 = $SSLServiceGrp.sslv2url; }
    @{ Column1 = "Enable Client Authentication"; Column2 = $SSLServiceGrp.clientauth; }
    @{ Column1 = "Client Certificates"; Column2 = $SSLServiceGrp.clientcert; }
    @{ Column1 = "SSL Redirect"; Column2 = $SSLServiceGrp.sslredirect; }
    @{ Column1 = "Enable non FIPS ciphers"; Column2 = $SSLServiceGrp.nonfipsciphers; }
    @{ Column1 = "SSL 2"; Column2 = $SSLServiceGrp.ssl2; }
    @{ Column1 = "SSL 3"; Column2 = $SSLServiceGrp.ssl3; }
    @{ Column1 = "TLS 1"; Column2 = $SSLServiceGrp.tls1; }
    @{ Column1 = "TLS 1.1"; Column2 = $SSLServiceGrp.tls11; }
    @{ Column1 = "TLS 1.2"; Column2 = $SSLServiceGrp.tls12; }
    @{ Column1 = "Server Name Indication (SNI)"; Column2 = $SSLServiceGrp.snienable; }
    @{ Column1 = "Enable Server Authentication"; Column2 = $SSLServiceGrp.serverauth; }
    @{ Column1 = "Common Name"; Column2 = $SSLServiceGrp.commonname; }
    @{ Column1 = "OCSP Check"; Column2 = $SSLServiceGrp.ocspcheck; }
    @{ Column1 = "CRL Check"; Column2 = $SSLServiceGrp.crlcheck; }
    @{ Column1 = "Service name"; Column2 = $SSLServiceGrp.servicename; }
    @{ Column1 = "Certificate Authority"; Column2 = $SSLServiceGrp.ca; }
    @{ Column1 = "SNI Certificate"; Column2 = $SSLServiceGrp.snicert; }
    @{ Column1 = "Send Close Notify"; Column2 = $SSLServiceGrp.sendclosenotify; }
    @{ Column1 = "SSL Profile"; Column2 = $SSLService.sslprofile; }

);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $SSLSERVICEGRPH;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
$Table = $null
} #end foreach
} Else {
WriteWordLine 0 0 "No SSL Service Groups have been configured."
WriteWordLine 0 0 " "
}

#endregion SSL Service Groups

#region SSL Profiles

WriteWordLine 2 0 "SSL Profiles"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `tSSL Profiles"

$SSLProfiles = Get-vNetScalerObject -Container config -Object sslprofile;

If ($SSLProfiles.Length -gt 0) {

Foreach ($SSLProfile in $SSLProfiles) {

$sslprofilename = $sslprofile.name

WriteWordLine 3 0 "SSL Profile: $sslprofilename"
WriteWordLine 0 0 " "
[System.Collections.Hashtable[]] $SSLPROFILEH = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "SSL Profile Type"; Column2 = $SSLprofile.sslprofiletype; }
    @{ Column1 = "Diffe-Hellman Key Exchange"; Column2 = $SSLProfile.dh; }
    @{ Column1 = "Diffe-Hellman Key File"; Column2 = $SSLProfile.dhfile; }
    @{ Column1 = "Diffe-Hellman Refresh Count"; Column2 = $SSLProfile.dhcount; }
    @{ Column1 = "Enable DH Key Expire Size Limit"; Column2 = $SSLProfile.dhkeyexpsizelimit; }
    @{ Column1 = "Enable Ephemeral RSA"; Column2 = $SSLProfile.ersa; }
    @{ Column1 = "Ephemeral RSA Refresh Count"; Column2 = $SSLProfile.ersacount; }
    @{ Column1 = "Allow session re-use"; Column2 = $SSLProfile.sessreuse; }
    @{ Column1 = "Session Time-out"; Column2 = $SSLProfile.sesstimeout; }
    @{ Column1 = "Enable Cipher Redirect"; Column2 = $SSLProfile.cipherredirect; }
    @{ Column1 = "Cipher Redirect URL"; Column2 = $SSLProfile.cipherurl; }
    @{ Column1 = "SSLv2 Redirect"; Column2 = $SSLProfile.sslv2redirect; }
    @{ Column1 = "SSLv2 Redirect URL"; Column2 = $SSLProfile.sslv2url; }
    @{ Column1 = "Enable Client Authentication"; Column2 = $SSLProfile.clientauth; }
    @{ Column1 = "Client Certificates"; Column2 = $SSLProfile.clientcert; }
    @{ Column1 = "SSL Redirect"; Column2 = $SSLProfile.sslredirect; }
    @{ Column1 = "Enable non FIPS ciphers"; Column2 = $SSLProfile.nonfipsciphers; }
    @{ Column1 = "SSL 2"; Column2 = $SSLProfile.ssl2; }
    @{ Column1 = "SSL 3"; Column2 = $SSLProfile.ssl3; }
    @{ Column1 = "TLS 1"; Column2 = $SSLProfile.tls1; }
    @{ Column1 = "TLS 1.1"; Column2 = $SSLProfile.tls11; }
    @{ Column1 = "TLS 1.2"; Column2 = $SSLProfile.tls12; }
    @{ Column1 = "Server Name Indication (SNI)"; Column2 = $SSLProfile.snienable; }
    @{ Column1 = "Enable Server Authentication"; Column2 = $SSLProfile.serverauth; }
    @{ Column1 = "Common Name"; Column2 = $SSLProfile.commonname; }
    @{ Column1 = "Push Encryption Trigger"; Column2 = $SSLProfile.pushenctrigger; }
    @{ Column1 = "Insertion Encoding"; Column2 = $SSLProfile.insertionencoding; }
    @{ Column1 = "Deny SSL Renegotiation"; Column2 = $SSLProfile.denysslreneg; }
    @{ Column1 = "Quantumn Size"; Column2 = $SSLProfile.quantumnsize; }
    @{ Column1 = "Strict CA Checks"; Column2 = $SSLProfile.strictcachecks; }
    @{ Column1 = "Drop Requests with no Host Header"; Column2 = $SSLProfile.dropreqwithnohostheader; }
    @{ Column1 = "Use bound CA chain for Client Authentication"; Column2 = $SSLProfile.clientauthuseboundcachain; }
    @{ Column1 = "Send Close Notify"; Column2 = $SSLProfile.sendclosenotify; }

);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $SSLPROFILEH;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
$Table = $null
} #end foreach
} Else {
WriteWordLine 0 0 "No SSL Profiles have been configured."
}

#endregion SSL Profiles

$selection.InsertNewPage()

#endregion NetScaler SSL

#endregion traffic management

#region NetScaler Security
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Security"
WriteWordLine 1 0 "NetScaler Security"
WriteWordLine 0 0 " "

#region AAA
WriteWordLine 2 0 "NetScaler AAA - Application Traffic"
WriteWordLine 0 0 " "

$aaavserverscount = Get-vNetScalerObjectCount -Container config -Object authenticationvserver;
$aaavservers = Get-vNetScalerObject -Container config -Object authenticationvserver;

if($aaavserverscount.__count -le 0) { WriteWordLine 0 0 "No AAA vServer has been configured"} else {


foreach ($aaavserver in $aaavservers) {
        $aaavservername = $aaavserver.name

        WriteWordLine 3 0 "NetScaler AAA Virtual Server: $aaavservername";

#region AAA vServer Basic Config

## IB - Create an array of hashtables to store our columns.
## IB - about column names as we'll utilise a -List(view)!
[System.Collections.Hashtable[]] $AAAVSH = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "IP Address"; Column2 = $aaavserver.ip; }
    @{ Column1 = "Value"; Column2 = $aaavserver.value; }
    @{ Column1 = "Port"; Column2 = $aaavserver.port; }
    @{ Column1 = "Service Type"; Column2 = $aaavserver.servicetype; }
    @{ Column1 = "Type"; Column2 = $aaavserver.type; }
    @{ Column1 = "State"; Column2 = $aaavserver.curstate; }
    @{ Column1 = "Status"; Column2 = $aaavserver.status; }
    @{ Column1 = "Cache Type"; Column2 = $aaavserver.cachetype; }
    @{ Column1 = "Redirect"; Column2 = $aaavserver.redirect; }
    @{ Column1 = "Precedence"; Column2 = $aaavserver.precedence; }
    @{ Column1 = "Redirect URL"; Column2 = $aaavserver.redirecturl; }
    @{ Column1 = "Authentication"; Column2 = $aaavserver.authentication; }
    @{ Column1 = "Authentication Domain"; Column2 = $aaavserver.authenticationdomain; }
    @{ Column1 = "Rule"; Column2 = $aaavserver.rule; }
    @{ Column1 = "Policy Name"; Column2 = $aaavserver.policyname; }
    @{ Column1 = "Policy"; Column2 = $aaavserver.policy; }
    @{ Column1 = "Service Name"; Column2 = $aaavserver.servicename; }
    @{ Column1 = "Weight"; Column2 = $aaavserver.weight; }
    @{ Column1 = "Caching vServer"; Column2 = $aaavserver.cachevserver; }
    @{ Column1 = "Backup vServer"; Column2 = $aaavserver.backupvserver; }
    @{ Column1 = "Client Timeout"; Column2 = $aaavserver.clttimeout; }
    @{ Column1 = "Spillover Method"; Column2 = $aaavserver.somethod; }
    @{ Column1 = "Spillover Threshold"; Column2 = $aaavserver.sothreshold; }
    @{ Column1 = "Spillover Persistence"; Column2 = $aaavserver.sopersistence; }
    @{ Column1 = "Spillover Persistence Timeout"; Column2 = $aaavserver.sopersistencetimeout; }
    @{ Column1 = "Priority"; Column2 = $aaavserver.priority; }
    @{ Column1 = "Downstate Flush"; Column2 = $aaavserver.downstateflush; }
    @{ Column1 = "Disable Primary When Down"; Column2 = $aaavserver.disableprimaryondown; }
    @{ Column1 = "Listen Policy"; Column2 = $aaavserver.listenpolicy; }
    @{ Column1 = "Listen Priority"; Column2 = $aaavserver.listenpriority; }
    @{ Column1 = "TCP Profile Name"; Column2 = $aaavserver.tcpprofilename; }
    @{ Column1 = "HTTP Profile Name"; Column2 = $aaavserver.httpprofilename; }
    @{ Column1 = "Comment"; Column2 = $aaavserver.comment; }
    @{ Column1 = "Enable AppFlow"; Column2 = $aaavserver.appflowlog; }
    @{ Column1 = "Virtual Server Type"; Column2 = $aaavserver.vstype; }
    @{ Column1 = "NetScaler Gateway Name"; Column2 = $aaavserver.ngname; }
    @{ Column1 = "Max Login Attempts"; Column2 = $aaavserver.maxloginattempts; }
    @{ Column1 = "Failed Login Timeout"; Column2 = $aaavserver.failedlogintimeout; }
    @{ Column1 = "Secondary"; Column2 = $aaavserver.secondary; }
    @{ Column1 = "Group Extraction Enabled"; Column2 = $aaavserver.groupextraction; }

    
    
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $AAAVSH;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
$Table = $null
#endregion AAA vServer Basic Config

#region AAA Cert Policies

            
        WriteWordLine 4 0 "Certificate Authentication Policies"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $aaacertpols = Get-vNetScalerObject -ResourceType authenticationvserver_authenticationcertpolicy_binding -name $aaavservername
        $errorcode = $aaacertpols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No Certificate authentication Policies have been configured"} else { #Uses the mentioned error code to determine existency of policy

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $CERTPOLHASH = @(); 

             foreach ($aaacertpol in $aaacertpols) {                
                $CERTPOLHASH += @{
                    Name = $aaacertpol.policy;
                    Secondary = $aaacertpol.secondary ;
                    Priority = $aaacertpol.priority;
                } # end Hashtable 
            }# end foreach

        if ($CERTPOLHASH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $CERTPOLHASH;
                Columns = "Name","Secondary","Priority";
                Headers = "Name","Secondary","Priority";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

            FindWordDocumentEnd;
                
        } else { WriteWordLine 0 0 "No Certificate Policies have been configured"} #endif AUTHPOLHASH.Length
    } #end if 
WriteWordLine 0 0 " "
$Table = $null

#endregion AAA Cert Policies

#region AAA LDAP Policies

            
        WriteWordLine 4 0 "LDAP Authentication Policies"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $aaaldappols = Get-vNetScalerObject -ResourceType authenticationvserver_authenticationldappolicy_binding -name $aaavservername
        $errorcode = $aaaldappols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No LDAP authentication Policies have been configured"} else { #Uses the mentioned error code to determine existency of policy

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $LDAPPOLHASH = @(); 

             foreach ($aaaldappol in $aaaldappols) {                
                $LDAPPOLHASH += @{
                    Name = $aaaldappol.policy;
                    Secondary = $aaaldappol.secondary ;
                    Priority = $aaaldappol.priority;
                } # end Hashtable 
            }# end foreach

        if ($LDAPPOLHASH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $LDAPPOLHASH;
                Columns = "Name","Secondary","Priority";
                Headers = "Name","Secondary","Priority";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

            FindWordDocumentEnd;
                
        } else { WriteWordLine 0 0 "No LDAP Policies have been configured"} #endif AUTHPOLHASH.Length
    } #end if 
WriteWordLine 0 0 " "
$Table = $null

#endregion AAA LDAP Policies

#region AAA Login Schema Policies

            
        WriteWordLine 4 0 "Login Schema Authentication Policies"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $aaalspols = Get-vNetScalerObject -ResourceType authenticationvserver_authenticationloginschemapolicy_binding -name $aaavservername
        $errorcode = $aaalspols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No Login Schema authentication Policies have been configured"} else { #Uses the mentioned error code to determine existency of policy

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $LSPOLHASH = @(); 

             foreach ($aaalspol in $aaalspols) {                
                $LSPOLHASH += @{
                    Name = $aaalspol.policy;
                    Secondary = $aaalspol.secondary ;
                    Priority = $aaalspol.priority;
                } # end Hashtable 
            }# end foreach

        if ($LSPOLHASH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $LSPOLHASH;
                Columns = "Name","Secondary","Priority";
                Headers = "Name","Secondary","Priority";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

            FindWordDocumentEnd;
                
        } else { WriteWordLine 0 0 "No Login Schema Policies have been configured"} #endif AUTHPOLHASH.Length
    } #end if 
WriteWordLine 0 0 " "
$Table = $null

#endregion AAA Login Schema Policies

#region AAA Negotiate Policies

            
        WriteWordLine 4 0 "Negotiate Authentication Policies"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $aaanegpols = Get-vNetScalerObject -ResourceType authenticationvserver_authenticationnegotiatepolicy_binding -name $aaavservername
        $errorcode = $aaanegpols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No Negotiate authentication Policies have been configured"} else { #Uses the mentioned error code to determine existency of policy

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $NEGPOLHASH = @(); 

             foreach ($aaanegpol in $aaanegpols) {                
                $NEGPOLHASH += @{
                    Name = $aaanegpol.policy;
                    Secondary = $aaanegpol.secondary ;
                    Priority = $aaanegpol.priority;
                } # end Hashtable 
            }# end foreach

        if ($NEGPOLHASH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $NEGPOLHASH;
                Columns = "Name","Secondary","Priority";
                Headers = "Name","Secondary","Priority";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

            FindWordDocumentEnd;
                
        } else { WriteWordLine 0 0 "No Negotiate Policies have been configured"} #endif AUTHPOLHASH.Length
    } #end if 
WriteWordLine 0 0 " "
$Table = $null

#endregion AAA Negotiate Policies

#region AAA Radius Policies

            
        WriteWordLine 4 0 "Radius Authentication Policies"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $aaaradpols = Get-vNetScalerObject -ResourceType authenticationvserver_authenticationradiuspolicy_binding -name $aaavservername
        $errorcode = $aaaradpols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No Radius authentication Policies have been configured"} else { #Uses the mentioned error code to determine existency of policy

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $RADPOLHASH = @(); 

             foreach ($aaaradpol in $aaaradpols) {                
                $RADPOLHASH += @{
                    Name = $aaaradpol.policy;
                    Secondary = $aaaradpol.secondary ;
                    Priority = $aaaradpol.priority;
                } # end Hashtable 
            }# end foreach

        if ($RADPOLHASH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $RADPOLHASH;
                Columns = "Name","Secondary","Priority";
                Headers = "Name","Secondary","Priority";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

            FindWordDocumentEnd;
                
        } else { WriteWordLine 0 0 "No Radius Policies have been configured"} #endif AUTHPOLHASH.Length
    } #end if 
WriteWordLine 0 0 " "
$Table = $null

#endregion AAA Radius Policies

#region AAA SAML IDP Policies

            
        WriteWordLine 4 0 "SAML IDP Authentication Policies"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $aaasidppols = Get-vNetScalerObject -ResourceType authenticationvserver_authenticationdamlidppolicy_binding -name $aaavservername
        $errorcode = $aaasidppols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No SAML IDP authentication Policies have been configured"} else { #Uses the mentioned error code to determine existency of policy

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $SIDPPOLHASH = @(); 

             foreach ($aaasidppol in $aaasidppols) {                
                $SIDPPOLHASH += @{
                    Name = $aaasidppol.policy;
                    Secondary = $aaasidppol.secondary ;
                    Priority = $aaasidppol.priority;
                } # end Hashtable 
            }# end foreach

        if ($SIDPPOLHASH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $SIDPPOLHASH;
                Columns = "Name","Secondary","Priority";
                Headers = "Name","Secondary","Priority";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

            FindWordDocumentEnd;
                
        } else { WriteWordLine 0 0 "No SAML IDP Policies have been configured"} #endif AUTHPOLHASH.Length
    } #end if 
WriteWordLine 0 0 " "
$Table = $null

#endregion AAA SAML IDP Policies

#region AAA SAML Policies

            
        WriteWordLine 4 0 "SAML Authentication Policies"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $aaasamlpols = Get-vNetScalerObject -ResourceType authenticationvserver_authenticationsamlpolicy_binding -name $aaavservername
        $errorcode = $aaasamlpols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No SAML authentication Policies have been configured"} else { #Uses the mentioned error code to determine existency of policy

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $SAMLPOLHASH = @(); 

             foreach ($aaasamlpol in $aaasamlpols) {                
                $SAMLPOLHASH += @{
                    Name = $aaasamlpol.policy;
                    Secondary = $aaasamlpol.secondary ;
                    Priority = $aaasamlpol.priority;
                } # end Hashtable 
            }# end foreach

        if ($SAMLPOLHASH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $SAMLPOLHASH;
                Columns = "Name","Secondary","Priority";
                Headers = "Name","Secondary","Priority";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

            FindWordDocumentEnd;
                
        } else { WriteWordLine 0 0 "No SAML Policies have been configured"} #endif AUTHPOLHASH.Length
    } #end if 
WriteWordLine 0 0 " "
$Table = $null

#endregion AAA SAML Policies

#region AAA TACAS Policies

            
        WriteWordLine 4 0 "TACAS Authentication Policies"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $aaatacaspols = Get-vNetScalerObject -ResourceType authenticationvserver_authenticationtacaspolicy_binding -name $aaavservername
        $errorcode = $aaatacaspols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No TACAS authentication Policies have been configured"} else { #Uses the mentioned error code to determine existency of policy

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $TACASPOLHASH = @(); 

             foreach ($aaatacaspol in $aaatacaspols) {                
                $TACASPOLHASH += @{
                    Name = $aaatacaspol.policy;
                    Secondary = $aaatacaspol.secondary ;
                    Priority = $aaatacaspol.priority;
                } # end Hashtable 
            }# end foreach

        if ($TACASPOLHASH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $TACASPOLHASH;
                Columns = "Name","Secondary","Priority";
                Headers = "Name","Secondary","Priority";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

            FindWordDocumentEnd;
                
        } else { WriteWordLine 0 0 "No TACAS Policies have been configured"} #endif AUTHPOLHASH.Length
    } #end if 
WriteWordLine 0 0 " "
$Table = $null

#endregion AAA TACAS Policies

#region AAA WebAuth Policies

            
        WriteWordLine 4 0 "WebAuth Authentication Policies"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $aaawebpols = Get-vNetScalerObject -ResourceType authenticationvserver_authenticationwebauthpolicy_binding -name $aaavservername
        $errorcode = $aaawebpols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No WebAuth authentication Policies have been configured"} else { #Uses the mentioned error code to determine existency of policy

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $WEBPOLHASH = @(); 

             foreach ($aaawebpol in $aaawebpols) {                
                $WEBPOLHASH += @{
                    Name = $aaawebpol.policy;
                    Secondary = $aaawebpol.secondary ;
                    Priority = $aaawebpol.priority;
                } # end Hashtable 
            }# end foreach

        if ($WEBPOLHASH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $WEBPOLHASH;
                Columns = "Name","Secondary","Priority";
                Headers = "Name","Secondary","Priority";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

            FindWordDocumentEnd;
                
        } else { WriteWordLine 0 0 "No WebAuth Policies have been configured"} #endif AUTHPOLHASH.Length
    } #end if 
WriteWordLine 0 0 " "
$Table = $null

#endregion AAA WebAuth Policies

#region SSL Parameters

 WriteWordLine 4 0 "SSL Parameters"
        WriteWordLine 0 0 " "

        $aaasslparameters = Get-vNetScalerObject -ResourceType sslvserver -Name $aaavservername;

## IB - Create an array of hashtables to store our columns.
## IB - about column names as we'll utilise a -List(view)!
[System.Collections.Hashtable[]] $AAASSLPARAMSH = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "Clear Text Port"; Column2 = $aaasslparameters.cleartextport; }
    @{ Column1 = "Diffe-Hellman Key Exchange"; Column2 = $aaasslparameters.dh; }
    @{ Column1 = "Diffe-Hellman Key File"; Column2 = $aaasslparameters.dhfile; }
    @{ Column1 = "Diffe-Hellman Refresh Count"; Column2 = $aaasslparameters.dhcount; }
    @{ Column1 = "Enable DH Key Expire Size Limit"; Column2 = $aaasslparameters.dhkeyexpsizelimit; }
    @{ Column1 = "Enable Ephemeral RSA"; Column2 = $aaasslparameters.ersa; }
    @{ Column1 = "Ephemeral RSA Refresh Count"; Column2 = $aaasslparameters.ersacount; }
    @{ Column1 = "Allow session re-use"; Column2 = $aaasslparameters.sessreuse; }
    @{ Column1 = "Session Time-out"; Column2 = $aaasslparameters.sesstimeout; }
    @{ Column1 = "Enable Cipher Redirect"; Column2 = $aaasslparameters.cipherredirect; }
    @{ Column1 = "Cipher Redirect URL"; Column2 = $aaasslparameters.cipherurl; }
    @{ Column1 = "SSLv2 Redirect"; Column2 = $aaasslparameters.sslv2redirect; }
    @{ Column1 = "SSLv2 Redirect URL"; Column2 = $aaasslparameters.sslv2url; }
    @{ Column1 = "Enable Client Authentication"; Column2 = $aaasslparameters.clientauth; }
    @{ Column1 = "Client Certificates"; Column2 = $aaasslparameters.clientcert; }
    @{ Column1 = "SSL Redirect"; Column2 = $aaasslparameters.sslredirect; }
    @{ Column1 = "SSL 2"; Column2 = $aaasslparameters.ssl2; }
    @{ Column1 = "SSL 3"; Column2 = $aaasslparameters.ssl3; }
    @{ Column1 = "TLS 1"; Column2 = $aaasslparameters.tls1; }
    @{ Column1 = "TLS 1.1"; Column2 = $aaasslparameters.tls11; }
    @{ Column1 = "TLS 1.2"; Column2 = $aaasslparameters.tls12; }
    @{ Column1 = "Server Name Indication (SNI)"; Column2 = $aaasslparameters.snienable; }
    @{ Column1 = "PUSH Encryption Trigger"; Column2 = $aaasslparameters.pushenctrigger; }
    @{ Column1 = "Send Close-Notify"; Column2 = $aaasslparameters.sendclosenotify; }
    @{ Column1 = "DTLS Profile"; Column2 = $aaasslparameters.dtlsprofilename; }
    @{ Column1 = "SSL Profile"; Column2 = $aaasslparameters.sslprofile; }

);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $AAASSLPARAMSH;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
$Table = $null

#endregion SSL Parameters

#region AAA SSL Ciphers             
        WriteWordLine 4 0 "SSL Ciphers"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $aaacipherbinds = Get-vNetScalerObject -ResourceType sslvserver_sslciphersuite_binding -name $aaavservername;
        $errorcode = $aaacipherbinds.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No SSL Ciphers have been bound."} else { #Uses the mentioned error code to determine existency of policy

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $CIPHERSH = @(); 

             foreach ($aaacipherbind in $aaacipherbinds) {                
                $CIPHERSH += @{
                    Name = $aaacipherbind.ciphername;
                    Description = $aaacipherbind.description;

                    
                } # end Hasthable $INTIPSH
            }# end foreach

        if ($CIPHERSH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $CIPHERSH;
                Columns = "Name","Description";
                Headers = "Name","Description";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

            FindWordDocumentEnd;
                
        } else { WriteWordLine 0 0 "No SSL Ciphers have been configured"} #endif AUTHPOLHASH.Length
    } #end if
WriteWordLine 0 0 " "
$Table = $null
#endregion AAA SSL Ciphers



} #endforeach region AAA
} #end if AAA vServers

#region KCD Accounts
 WriteWordLine 3 0 "KCD Accounts"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $kcdaccounts = Get-vNetScalerObject -ResourceType aaakcdaccount;
        $errorcode = $kcdaccounts.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If (($errorcode -ne 1) -or (!$kcdaccounts)) {WriteWordLine 0 0 "No KCD Accounts have been configured."; WriteWordLine 0 0 " ";} else { #Uses the mentioned error code to determine existency of policy

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $KCDH = @(); 

             foreach ($kcdaccount in $kcdaccounts) {           
             $kcdname = $kcdaccount.kcdaccount
             WriteWordLine 3 0 "KCD Account: $kcdname"     
             WriteWordLine 0 0 " "   
                ## IB - Create an array of hashtables to store our columns.
## IB - about column names as we'll utilise a -List(view)!
[System.Collections.Hashtable[]] $KCDH = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "KeyTab File"; Column2 = $kcdaccount.keytab; }
    @{ Column1 = "Principle"; Column2 = $kcdaccount.principle; }
    @{ Column1 = "SPN"; Column2 = $kcdaccount.kcdspn; }
    @{ Column1 = "Realm"; Column2 = $kcdaccount.realmstr; }
    @{ Column1 = "User Realm"; Column2 = $kcdaccount.userrealm; }
    @{ Column1 = "Enterprise Realm"; Column2 = $kcdaccount.enterpriserealm; }
    @{ Column1 = "Delegated User"; Column2 = $kcdaccount.delegateduser; }
    @{ Column1 = "KCD Password"; Column2 = $kcdaccount.kcdpassword; }
    @{ Column1 = "User Certificate"; Column2 = $kcdaccount.usercert; }
    @{ Column1 = "CA Certificate"; Column2 = $kcdaccount.cacert; }
    @{ Column1 = "Service SPN"; Column2 = $kcdaccount.servicespn; }


);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $KCDH;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
$Table = $null
            }# end foreach

        
WriteWordLine 0 0 " "

#endregion KCD Accounts

} #endif region AAA

#endregion AAA

#region AppFW

        WriteWordLine 2 0 "Application Firewall"
        WriteWordLine 0 0 " "

#region AppFW Profiles

        WriteWordLine 3 0 "Application Firewall Profiles"
        WriteWordLine 0 0 " "

                $errorcode = 1 #Set Error code to 1
        $fwprofiles = Get-vNetScalerObject -ResourceType appfwprofile;
        $errorcode = $fwprofiles.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No AppFW Profiles have been configured."} else { #Uses the mentioned error code to determine existency of policy

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $AFWPROFH = @(); 

             foreach ($fwprofile in $fwprofiles) {           
             $fwprofilename = $fwprofile.name
             WriteWordLine 4 0 "Profile: $fwprofilename"     
             WriteWordLine 0 0 " "   
                ## IB - Create an array of hashtables to store our columns.
## IB - about column names as we'll utilise a -List(view)!
[System.Collections.Hashtable[]] $AFWPROFH = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "Profile Type"; Column2 = $fwprofile.type; }
    @{ Column1 = "StartURL Action"; Column2 = $fwprofile.starturlaction -join ","; }
    @{ Column1 = "Content Type Action"; Column2 = $fwprofile.contenttypeaction -join ","; }
    @{ Column1 = "Inspect Content Types"; Column2 = $fwprofile.inspectcontenttypes -join ","; }
    @{ Column1 = "Start URL Closure"; Column2 = $fwprofile.starturlclosure; }
    @{ Column1 = "Deny URL Action"; Column2 = $fwprofile.denyurlaction -join ","; }
    @{ Column1 = "Referer Header Check"; Column2 = $fwprofile.refererheadercheck; }
    @{ Column1 = "Cookie Consistency Action"; Column2 = $fwprofile.cookieconsistencyaction -join ","; }
    @{ Column1 = "Cookie Transformation"; Column2 = $fwprofile.cookietransforms; }
    @{ Column1 = "Cookie Encryption"; Column2 = $fwprofile.cookieencryption; }
    @{ Column1 = "Proxy Cookies"; Column2 = $fwprofile.cookieproxying; }
    @{ Column1 = "Add Cookie Flags"; Column2 = $fwprofile.addcookieflags; }
    @{ Column1 = "Field Consistency Check"; Column2 = $fwprofile.fieldconsistencyaction; }
    @{ Column1 = "Cross Site Request Forgery Tag Check"; Column2 = $fwprofile.csrftagaction -join ","; }
    @{ Column1 = "XSS (Cross-Site Scripting) Check"; Column2 = $fwprofile.crosssitescriptingaction; }
    @{ Column1 = "Transform Cross-Site Scripts"; Column2 = $fwprofile.crosssitescriptingtransformunsafehtml; }
    @{ Column1 = "XSS - Check complete URLs"; Column2 = $fwprofile.crosssitescriptingcheckcompleteurls; }
    @{ Column1 = "SQL Injection Action"; Column2 = $fwprofile.sqlinjectionaction -join ","; }
    @{ Column1 = "SQL Injection - Transform Special Characters"; Column2 = $fwprofile.sqlinjectiontransformspecialchars; }
    @{ Column1 = "SQL Injection - Only check fields with SQL Characters"; Column2 = $fwprofile.sqlinjectiononlycheckfieldswithsqlchars; }
    @{ Column1 = "SQL Injection Type"; Column2 = $fwprofile.sqlinjectiontype; }
    @{ Column1 = "SQL Injection - Check SQL wild characters"; Column2 = $fwprofile.sqlinjectionchecksqlwildchars; }
    @{ Column1 = "Field Format Actions"; Column2 = $fwprofile.fieldformataction; }
    @{ Column1 = "Default Field Format Type"; Column2 = $fwprofile.defaultfieldformattype; }
    @{ Column1 = "Default Field Format minimum length"; Column2 = $fwprofile.defaultfieldformatminlength; }
    @{ Column1 = "Default Field Format maximum length"; Column2 = $fwprofile.defaultfieldformatmaxlength; }
    @{ Column1 = "Buffer Overflow Actions"; Column2 = $fwprofile.bufferoverflowaction; }
    @{ Column1 = "Buffer Overflow - Maximum URL Length"; Column2 = $fwprofile.bufferoverflowmaxurllength; }
    @{ Column1 = "Buffer Overflow - Maximum Header Length"; Column2 = $fwprofile.bufferoverflowmaxheaderlength; }
    @{ Column1 = "Buffer Overflow - Maximum Cookie Length"; Column2 = $fwprofile.bufferoverflowmaxcookielength; }
    @{ Column1 = "Credit Card Action"; Column2 = $fwprofile.creditcardaction -join ","; }
    @{ Column1 = "Credit Card Types to protect"; Column2 = $fwprofile.creditcard -join ","; }
    @{ Column1 = "Maximum number of Credit Cards per page"; Column2 = $fwprofile.creditcardmaxallowed; }
    @{ Column1 = "X-Out Credit Card Numbers"; Column2 = $fwprofile.creditcardxout; }
    @{ Column1 = "Log Credit Card Numbers when matched"; Column2 = $fwprofile.dosecurecreditcardlogging; }
    @{ Column1 = "Request Streaming"; Column2 = $fwprofile.streaming; }
    @{ Column1 = "Trace Status"; Column2 = $fwprofile.trace; }
    @{ Column1 = "Request Content Type"; Column2 = $fwprofile.requestcontenttype; }
    @{ Column1 = "Response Content Type"; Column2 = $fwprofile.responsecontenttype; }
    @{ Column1 = "XML Denial of Service Action"; Column2 = $fwprofile.xmldosaction -join ","; }
    @{ Column1 = "XML Format Action"; Column2 = $fwprofile.xmlformataction -join "," ; }
    @{ Column1 = "XML SQL Injection Action"; Column2 = $fwprofile.xmlsqlinjectionaction -join ","; }
    @{ Column1 = "XML SQL Injection - Only check fields with SQL characters"; Column2 = $fwprofile.xmlsqlinjectiononlycheckfieldswithsqlchars; }
    @{ Column1 = "XML SQL Injection - Type"; Column2 = $fwprofile.xmlsqlinjectiontype; }
    @{ Column1 = "XML SQL Injection - Check fields with SQL Wild characters"; Column2 = $fwprofile.xmlsqlinjectionchecksqlwildchars; }
    @{ Column1 = "XML SQL Injection - Parse Comments"; Column2 = $fwprofile.xmlsqlinjectionparsecomments; }
    @{ Column1 = "XML XSS (Cross-Site Scripting) Action"; Column2 = $fwprofile.xmlxssaction -join ","; }
    @{ Column1 = "XML WSI (Web Services Interoperability) Action"; Column2 = $fwprofile.xmlwsiaction -join ","; }
    @{ Column1 = "XML Attachments Action"; Column2 = $fwprofile.xmlattachmentaction -join ","; }
    @{ Column1 = "XML validation Action"; Column2 = $fwprofile.xmlvalidationaction -join ","; }
    @{ Column1 = "XML Error Object Name"; Column2 = $fwprofile.xmlerrorobject; }
    @{ Column1 = "Custom Settings"; Column2 = $fwprofile.customsettings; }
    @{ Column1 = "Signatures"; Column2 = $fwprofile.signatures; }
    @{ Column1 = "XML SOAP Fault Action"; Column2 = $fwprofile.xmlsoapfaultaction -join ","; }
    @{ Column1 = "Use HTML Error Object"; Column2 = $fwprofile.usehtmlerrorobject; }
    @{ Column1 = "Error URL"; Column2 = $fwprofile.errorurl; }
    @{ Column1 = "HTML Error Object Name"; Column2 = $fwprofile.htmlerrorobject; }
    @{ Column1 = "Log Every Policy Hit"; Column2 = $fwprofile.logeverypolicyhit; }
    @{ Column1 = "Strip Comments"; Column2 = $fwprofile.stripcomments; }
    @{ Column1 = "Strip HTML Comments"; Column2 = $fwprofile.striphtmlcomments; }
    @{ Column1 = "Strip XML Comments"; Column2 = $fwprofile.sttripxmlcomments; }
    @{ Column1 = "Exempt URLS passing the Start URL Closure check from Security Checks"; Column2 = $fwprofile.exemptclosureurlsfromsecuritychecks; }
    @{ Column1 = "Default Character Set"; Column2 = $fwprofile.defaultcharset; }
    @{ Column1 = "Maximum Post Body Size (bytes)"; Column2 = $fwprofile.postbodylimit; }
    @{ Column1 = "Maximum number of file uploads per form submission"; Column2 = $fwprofile.fileuploadmaxnum; }
    @{ Column1 = "Perform Entity encoding for special response characters"; Column2 = $fwprofile.canonicalizehtmlresponse; }
    @{ Column1 = "Enable Form Tagging"; Column2 = $fwprofile.enableformtagging; }
    @{ Column1 = "Perform Sessionless Field Consistency Checks"; Column2 = $fwprofile.sessionlessfieldconsistency; }
    @{ Column1 = "Enable Sessionless URL Closure Checks"; Column2 = $fwprofile.sessionlessurlclosure; }
    @{ Column1 = "Allow Semi-Colon field separator in URL"; Column2 = $fwprofile.semicolonfieldseparator; }
    @{ Column1 = "Exclude Uploaded Files from Checks"; Column2 = $fwprofile.excludefileuploadfromchecks; }
    @{ Column1 = "HTML SQL Injection - Parse Comments"; Column2 = $fwprofile.sqlinjectionparsecomments; }
    @{ Column1 = "Method for handling Percent encoded names"; Column2 = $fwprofile.invalidpercenthandling; }
    @{ Column1 = "Check Request Headers for SQL Injection and XSS"; Column2 = $fwprofile.checkrequestheaders; }
    @{ Column1 = "Optimize Partial Requests"; Column2 = $fwprofile.optimizepartialreqs; }
    @{ Column1 = "URL decode Request Cookies"; Column2 = $fwprofile.urldecoderequestcookies; }
    @{ Column1 = "Comment"; Column2 = $fwprofile.comment; }
    @{ Column1 = "Archive Name"; Column2 = $fwprofile.archivename; }
    @{ Column1 = "State"; Column2 = $fwprofile.state; }


    



);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $AFWPROFH;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "

            }# end foreach

}        
WriteWordLine 0 0 " "

#endregion AppFw Profiles

#region AppFW Policies

        WriteWordLine 3 0 "Application Firewall Policies"
        WriteWordLine 0 0 " "

                $errorcode = 1 #Set Error code to 1
        $fwpolicies = Get-vNetScalerObject -ResourceType appfwpolicy;
        $errorcode = $fwpolicies.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If (($errorcode -ne 1) -or (!$fwpolicies)) {WriteWordLine 0 0 "No AppFW Policies have been configured."} else { #Uses the mentioned error code to determine existency of policy

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $AFWPOLH = @(); 

             foreach ($fwpolicy in $fwpolicies) {           
             $fwpolicyname = $fwpolicy.name
             WriteWordLine 3 0 "Profile: $fwpolicyname"     
             WriteWordLine 0 0 " "   
                ## IB - Create an array of hashtables to store our columns.
## IB - about column names as we'll utilise a -List(view)!
[System.Collections.Hashtable[]] $AFWPOLH = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "Rule"; Column2 = $fwpolicy.rule; }
    @{ Column1 = "Profile Name"; Column2 = $fwolicy.profilename; }
    @{ Column1 = "Comment"; Column2 = $fwpolicy.comment; }
    @{ Column1 = "Log Action"; Column2 = $fwprofile.logaction; }
    

    



);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $AFWPOLH;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
}# end foreach
            }# end foreach

        
WriteWordLine 0 0 " "

#endregion AppFw Policies


#endregion AppFW


#endregion NetScaler Security

#region Citrix NetScaler Gateway

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters Citrix NetScaler (Access) Gateway"
WriteWordLine 1 0 "Citrix NetScaler (Access) Gateway"
WriteWordLine 0 0 " "

#region Citrix NetScaler Gateway CAG Global

WriteWordLine 2 0 "NetScaler Gateway Global Settings"
Write-Verbose "$(Get-Date): `tNetScaler Gateway Global Settings"
WriteWordLine 0 0 " "
#region GlobalNetwork
<#
WriteWordLine 3 0 "Global Settings Network"

## IB - Create an array of hashtables to store our columns. Note: If we need the
## IB - headers to include spaces we can override these at table creation time.
## IB - Create the parameters to pass to the AddWordTable function

ForEach ($LINE in $SetVpnParameter) {
    $Params = $null
    $Params = @{
        Hashtable = @{
            ## IB - Each hashtable is a separate row in the table!
            Wins = Get-StringProperty $LINE "-winsIP" "Not Configured";
            Mapped = Get-StringProperty $LINE "-useMIP" "VPN_SESS_ACT_NS";
            Intranet = Get-StringProperty $LINE "-iipDnsSuffix" "Not Configured";
            Http = Get-StringProperty $LINE "-httpPort" "Not Configured";
            Timeout = Get-StringProperty $LINE "-forcedTimeout" "Not Configured";
        }
        Columns = "Wins","Mapped","Intranet","Http","Timeout";
        Headers = "WINS Server","Mapped IP","Intranet IP","HTTP Ports","Forced Time-out";
        Format = -235; ## IB - Word constant for Light List Accent 5
        AutoFit = $wdAutoFitContent;
    }

    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -NoGridLines;
    }

FindWordDocumentEnd;

WriteWordLine 0 0 " "
#>
#endregion GlobalNetwork

#region GlobalClientExperience
WriteWordLine 3 0 "Global Settings Client Experience"
WriteWordLine 0 0 " "
$cagglobalclient = Get-vNetScalerObject -Object vpnparameter;

## IB - Create an array of hashtables to store our columns.
## IB - about column names as we'll utilise a -List(view)!
[System.Collections.Hashtable[]] $NsGlobalClientExperience = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "Homepage"; Column2 = $cagglobalclient.emailhome; }
    @{ Column1 = "URL for Web Based Email"; Column2 = $cagglobalclient.sesstimeout; }
    @{ Column1 = "Session Time-Out"; Column2 = $cagglobalclient.sesstimeout; }
    @{ Column1 = "Client-Idle Time-Out"; Column2 = $cagglobalclient.clientidletimeoutwarning; }
    @{ Column1 = "Single Sign-On to Web Applications"; Column2 = $cagglobalclient.sso; }
    @{ Column1 = "Credential Index"; Column2 = $cagglobalclient.ssocredential; }
    @{ Column1 = "Single Sign-On with Windows"; Column2 = $cagglobalclient.windowsautologon; }
    @{ Column1 = "Split Tunnel"; Column2 = $cagglobalclient.splittunnel; }
    @{ Column1 = "Local LAN Access"; Column2 = $cagglobalclient.locallanaccess; }
    @{ Column1 = "Plug-in Type"; Column2 = $cagglobalclient.windowsclienttype; }
    @{ Column1 = "Windows Plugin Upgrade"; Column2 = $cagglobalclient.windowspluginupgrade; }
    @{ Column1 = "MAC Plugin Upgrade"; Column2 = $cagglobalclient.macpluginupgrade; }
    @{ Column1 = "Linux Plugin Upgrade"; Column2 = $cagglobalclient.linuxpluginupgrade; }
    @{ Column1 = "AlwaysON Profile Name"; Column2 = $cagglobalclient.alwaysonprofilename; }
    @{ Column1 = "Clientless Access"; Column2 = $cagglobalclient.clientlessvpnmode; }
    @{ Column1 = "Clientless URL Encoding"; Column2 = $cagglobalclient.clientlessmodeurlencoding; }
    @{ Column1 = "Clientless Persistent Cookie"; Column2 = $cagglobalclient.clientlesspersistentcookie; }
    @{ Column1 = "Single Sign-on to Web Applications"; Column2 = $cagglobalclient.sso; }
    @{ Column1 = "Credential Index"; Column2 = $cagglobalclient.ssocredential; }
    @{ Column1 = "KCD Account"; Column2 = $cagglobalclient.kcdaccount; }
    @{ Column1 = "Single Sign-on with Windows"; Column2 = $cagglobalclient.windowsautologon; }
    @{ Column1 = "Client Cleanup Prompt"; Column2 = $cagglobalclient.clientcleanupprompt; }
    @{ Column1 = "UI Theme"; Column2 = $cagglobalclient.uitheme; }
    @{ Column1 = "Login Script"; Column2 = $cagglobalclient.loginscript; }
    @{ Column1 = "Logout Script"; Column2 = $cagglobalclient.logoutscript; }
    @{ Column1 = "Application Token Timeout"; Column2 = $cagglobalclient.apptokentimeout; }
    @{ Column1 = "MDX Token Timeout"; Column2 = $cagglobalclient.mdxtokentimeout; }
    @{ Column1 = "Allow Users to Change Log Levels"; Column2 = $cagglobalclient.clientconfiguration; }
    @{ Column1 = "Allow access to private network IP addresses only"; Column2 = $cagglobalclient.windowsclienttype; }
    @{ Column1 = "Client Choices"; Column2 = $cagglobalclient.clientchoices; }
    @{ Column1 = "Show VPN Plugin icon"; Column2 = $cagglobalclient.iconwithreceiver; }
    
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $NsGlobalClientExperience;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
#endregion GlobalClientExperience

#region GlobalSecurity
WriteWordLine 3 0 "Global Settings Security"
WriteWordLine 0 0 " "
## IB - Create an array of hashtables to store our columns. Note: If we need the
## IB - headers to include spaces we can override these at table creation time.
## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = @{
        ## IB - Each hashtable is a separate row in the table!
        DEFAUTH = $cagglobalclient.defaultauthorizationaction;
        CLISEC = $cagglobalclient.encryptcsecexp;
        SECBRW = $cagglobalclient.securebrowse;
    }
    Columns = "DEFAUTH","CLISEC","SECBRW";
    Headers = "Default Authorization Action","Client Security Encryption","Secure Browse";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params -NoGridLines;
FindWordDocumentEnd;


WriteWordLine 0 0 " "
$Table = $null
#endregion GlobalSecurity

#region GlobalPublishedApps
WriteWordLine 3 0 "Global Settings Published Applications"
WriteWordLine 0 0 " "
## IB - Create an array of hashtables to store our columns. Note: If we need the
## IB - headers to include spaces we can override these at table creation time.
## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = @{
        ICAPROXY = $cagglobalclient.icaproxy;
        WIHOME = $cagglobalclient.wihome;
        WIMODE = $cagglobalclient.wihomeaddresstype;
        SSO = $cagglobalclient.sso;
        RECEIVERHOME = $cagglobalclient.citrixreceiverhome;
        ASA = $cagglobalclient.storefronturl;

    }
    Columns = "ICAPROXY","WIHOME","WIMODE","SSO","RECEIVERHOME", "ASA";
    Headers = "ICA Proxy","Web Interface Address","Web Interface Portal Mode","Single Sign-On Domain", "Receiver Homepage", "Account Services Address";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params -NoGridLines;
FindWordDocumentEnd;


WriteWordLine 0 0 " "
$Table = $null
#endregion GlobalPublishedApps

#region Global STA
WriteWordLine 3 0 "Global Settings Secure Ticket Authority Configuration"
WriteWordLine 0 0 " "
$vpnglobalstascount = Get-vNetScalerObjectCount -Container config -Object vpnglobal_staserver_binding;
$vpnglobalstascounter = $vpnglobalstascount.__count

if($vpnglobalstascounter -le 0) { WriteWordLine 0 0 "No Global Secure Ticket Authority has been configured"} else {
    
    $vpnglobalstas = Get-vNetScalerObject -Container config -Object vpnglobal_staserver_binding;
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $STASH = @();

    foreach ($vpnglobalsta in $vpnglobalstas) {
        $STASH += @{ 
            STA = $vpnglobalsta.staserver; 
            STAAUTHID = $vpnglobalsta.STAAUTHID;
            STAADDRESSTYPE = $vpnglobalsta.staaddresstype;
        }
    } 
    $Params = $null
    $Params = @{
        Hashtable = $STASH;
        Columns = "STA","STAAUTHID","STAADDRESSTYPE";
        Headers = "Secure Ticket Authority","Authentication ID","Address Type";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
    }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            }
FindWordDocumentEnd;
WriteWordLine 0 0 " "
$Table = $null
#endregion Global STA

#region Global AppController
WriteWordLine 3 0 "Global Settings App Controller Configuration"
WriteWordLine 0 0 " "
$vpnglobalappcontrollercount = Get-vNetScalerObjectCount -Container config -Object vpnglobal_appcontroller_binding;
$vpnglobalappcontrollercounter = $vpnglobalappcontrollercount.__count

if($vpnglobalappcontrollercounter -le 0) { WriteWordLine 0 0 "No Global App Controller has been configured"} else {
    
    $vpnglobalappcs = Get-vNetScalerObject -Container config -Object vpnglobal_appcontroller_binding;
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $APPCH = @();

    ## IB - Iterate over all load balancer bindings (uses new function)
    foreach ($vpnglobalappc in $vpnglobalappcs) {
        $APPCH += @{
            APPController = $vpnglobalappc.appController;
        } 
    }

    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = $APPCH;
        AutoFit = $wdAutoFitContent;
        Format = -235; ## IB - Word constant for Light List Accent 5
        }
    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -NoGridLines;
    }
FindWordDocumentEnd;
WriteWordLine 0 0 " "
$Table = $null
#endregion Global AppController

#region GlobalAAAParams
WriteWordLine 3 0 "Global Settings AAA Parameters"
WriteWordLine 0 0 " "
$cagaaa = Get-vNetScalerObject -Object aaaparameter;

## IB - Create an array of hashtables to store our columns.
## IB - about column names as we'll utilise a -List(view)!
[System.Collections.Hashtable[]] $NsGlobalAAAH = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "Maximum number of Users"; Column2 = $cagaaa.maxaaausers; }
    @{ Column1 = "MaxLogin Attempts"; Column2 = $cagaaa.maxloginattempts; }
    @{ Column1 = "NAT IP Address"; Column2 = $cagaaa.aaadnatip; }
    @{ Column1 = "Failed login timeout"; Column2 = $cagaaa.failedlogintimeout; }
    @{ Column1 = "Default Authentication Type"; Column2 = $cagaaa.defaultauthtype; }
    @{ Column1 = "AAA Session Log Levels"; Column2 = $cagaaa.aaasessionloglevel; }
    @{ Column1 = "Enable Static Page Caching"; Column2 = $cagaaa.enablestaticpagecaching; }
    @{ Column1 = "Enable Enhanced Authentication Feedback"; Column2 = $cagaaa.enableenhancedauthfeedback; }
    @{ Column1 = "Enable Session Stickiness"; Column2 = $cagaaa.enablesessionstickiness; }
   
    
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $NsGlobalAAAH;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
$Table = $null
#endregion GlobalAAAParams

#region NetScaler Gateway Intranet Applications
WriteWordLine 2 0 "NetScaler Gateway Intranet Applications";
WriteWordLine 0 0 " ";
$vpnintappscount = Get-vNetScalerObjectCount -Container config -Object vpnintranetapplication;
$vpnintapps = Get-vNetScalerObject -Container config -Object vpnintranetapplication;

if($vpnintappscount.__count -le 0) { WriteWordLine 0 0 "No Intranet Applications have been configured"} else {

    foreach ($vpnintapp in $vpnintapps) {
        $vpnintappname = $vpnintapp.intranetapplication

        WriteWordLine 3 0 "NetScaler Gateway Intranet Application: $vpnintappname";
        WriteWordLine 0 0 " "




## IB - Create an array of hashtables to store our columns.
## IB - about column names as we'll utilise a -List(view)!
[System.Collections.Hashtable[]] $VPNINTAPPH = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "Protocol"; Column2 = $vpnintapp.protocol; }
    @{ Column1 = "Destination IP Address"; Column2 = $vpnintapp.destip; }
    @{ Column1 = "Netmask"; Column2 = $vpnintapp.netmask; }
    @{ Column1 = "IP Range"; Column2 = $vpnintapp.iprange; }
    @{ Column1 = "Hostname"; Column2 = $vpnintapp.hostname; }
    @{ Column1 = "Client Application"; Column2 = $vpnintapp.clientapplication; }
    @{ Column1 = "Spoof IP"; Column2 = $vpnintapp.spoofiip; }
    @{ Column1 = "Destination Port"; Column2 = $vpnintapp.destport; }
    @{ Column1 = "Interception Mode"; Column2 = $vpnintapp.interception; }
    @{ Column1 = "Source IP"; Column2 = $vpnintapp.srcip; }
    @{ Column1 = "Source Port"; Column2 = $vpnintapp.srcprt; }
   
    
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $VPNINTAPPH;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
$Table = $null


}
}
#endregion NetScaler Gateway Intranet Applications

#region NetScaler Gateway Bookmarks
WriteWordLine 2 0 "NetScaler Gateway Bookmarks";
WriteWordLine 0 0 " ";
$vpnurlscount = Get-vNetScalerObjectCount -Container config -Object vpnurl;
$vpnurls = Get-vNetScalerObject -Container config -Object vpnurl;

if($vpnurlscount.__count -le 0) { WriteWordLine 0 0 "No Bookmarks have been configured"} else {

    foreach ($vpnurl in $vpnurls) {
        $vpnurlname = $vpnurl.urlname

        WriteWordLine 3 0 "NetScaler Gateway Bookmark: $vpnurlname";
        WriteWordLine 0 0 " "




## IB - Create an array of hashtables to store our columns.
## IB - about column names as we'll utilise a -List(view)!
[System.Collections.Hashtable[]] $VPNURLH = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = $vpnurl.linkname; }
    @{ Column1 = "URL"; Column2 = $vpnurl.actualurl; }
    @{ Column1 = "Virtual Server Name"; Column2 = $vpnurl.vservername; }
    @{ Column1 = "Clientless Access"; Column2 = $vpnurl.clientlessaccess; }
    @{ Column1 = "Comment"; Column2 = $vpnurl.comment; }
    @{ Column1 = "Icon URL"; Column2 = $vpnurl.iconurl; }
    @{ Column1 = "SSO Type"; Column2 = $vpnurl.ssotype; }
    @{ Column1 = "Application Type"; Column2 = $vpnurl.applicationtype; }
    @{ Column1 = "SAML SSO Profile"; Column2 = $vpnurl.samlssoprofile; }
      
    
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $VPNURLH;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
$Table = $null


}
}
#endregion NetScaler Gateway Bookmarks

$selection.InsertNewPage()

#endregion Citrix NetScaler Gateway CAG Global

#region CAG vServers

$vpnvserverscount = Get-vNetScalerObjectCount -Container config -Object vpnvserver;
$vpnvservers = Get-vNetScalerObject -Container config -Object vpnvserver;

if($vpnvserverscount.__count -le 0) { WriteWordLine 0 0 "No Citrix NetScaler Gateway has been configured"} else {

    foreach ($vpnvserver in $vpnvservers) {
        $vpnvservername = $vpnvserver.name

        WriteWordLine 2 0 "NetScaler Gateway Virtual Server: $vpnvservername";
#region CAG vServer basic configuration



## IB - Create an array of hashtables to store our columns.
## IB - about column names as we'll utilise a -List(view)!
[System.Collections.Hashtable[]] $VPNVSERVERH = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "State"; Column2 = $vpnvserver.state; }
    @{ Column1 = "IP Address"; Column2 = $vpnvserver.ipv46; }
    @{ Column1 = "Port"; Column2 = $vpnvserver.port; }
    @{ Column1 = "Protocol"; Column2 = $vpnvserver.servicetype; }
    @{ Column1 = "RDP Server Profile"; Column2 = $vpnvserver.rdpserverprofilename; }
    @{ Column1 = "Login Once"; Column2 = $vpnvserver.loginonce; }
    @{ Column1 = "Double Hop"; Column2 = $vpnvserver.doublehop; }
    @{ Column1 = "Down State Flush"; Column2 = $vpnvserver.downstateflush; }
    @{ Column1 = "DTLS"; Column2 = $vpnvserver.dtls; }
    @{ Column1 = "AppFlow Logging"; Column2 = $vpnvserver.appflowlog; }
    @{ Column1 = "Maximum Users"; Column2 = $vpnvserver.maxaaausers; }
    @{ Column1 = "Max Login Attempts"; Column2 = $vpnvserver.maxloginattempts; }
    @{ Column1 = "Failed Login Timeout"; Column2 = $vpnvserver.failedlogintimeout; }
    @{ Column1 = "ICA Only"; Column2 = $vpnvserver.icaonly; }
    @{ Column1 = "Enable Authentication"; Column2 = $vpnvserver.authentication; }
    @{ Column1 = "Windows EPA Plugin Upgrade"; Column2 = $vpnvserver.windowsepapluginupgrade; }
    @{ Column1 = "Linux EPA Plugin Upgrade"; Column2 = $vpnvserver.linuxepapluginupgrade; }
    @{ Column1 = "Mac EPA Plugin Upgrade"; Column2 = $vpnvserver.macepapluginupgrade; }
    @{ Column1 = "ICA Proxy Session Migration"; Column2 = $vpnvserver.icaproxysessionmigration; }
    @{ Column1 = "Enable Device Certificates"; Column2 = $vpnvserver.devicecert; }

   
    
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $VPNVSERVERH;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
$Table = $null

#endregion CAG vServer basic configuration


  
    #region CAG Authentication LDAP Policies             
        WriteWordLine 3 0 "Authentication LDAP Policies"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $vpnvserverldappols = Get-vNetScalerObject -ResourceType vpnvserver_authenticationldappolicy_binding -name $vpnvserver.Name
        $errorcode = $vpnvserverldappols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No LDAP Policies have been configured"} else { #Uses the mentioned error code to determine existency of policy

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $AUTHPOLHASH = @(); 

             foreach ($vpnvserverldappol in $vpnvserverldappols) {                
                $AUTHPOLHASH += @{
                    Name = $vpnvserverldappol.policy;
                    Secondary = $vpnvserverldappol.secondary ;
                    Priority = $vpnvserverldappol.priority;
                } # end Hasthable $AUTHPOLH1
            }# end foreach $AUTHPOLS

        if ($AUTHPOLHASH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $AUTHPOLHASH;
                Columns = "Name","Secondary","Priority";
                Headers = "Name","Secondary","Priority";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

            FindWordDocumentEnd;
                
        } else { WriteWordLine 0 0 "No LDAP Policies have been configured"} #endif AUTHPOLHASH.Length
    } #end if no LDAP configures
WriteWordLine 0 0 " "
$Table = $null
#endregion CAG Authentication LDAP Policies  

    #region CAG Authentication Radius Policies             
        WriteWordLine 3 0 "Authentication RADIUS Policies"
        WriteWordLine 0 0 " "
        $errorcode = 1 
        $vpnvserverradiuspols = Get-vNetScalerObject -ResourceType vpnvserver_authenticationradiuspolicy_binding -name $vpnvserver.Name
        $errorcode = $vpnvserverradiuspols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No RADIUS Policies have been configured"} else {

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $AUTHPOLRADHASH = @(); 

             foreach ($vpnvserverradiuspol in $vpnvserverradiuspols) {                
                $AUTHPOLRADHASH += @{
                    Name = $vpnvserverradiuspol.policy;
                    Secondary = $vpnvserverradiuspol.secondary ;
                    Priority = $vpnvserverradiuspol.priority;
                }
            }

        if ($AUTHPOLRADHASH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $AUTHPOLRADHASH;
                Columns = "Name","Secondary","Priority";
                Headers = "Name","Secondary","Priority";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

            FindWordDocumentEnd;
                
        } else { WriteWordLine 0 0 "No RADIUS Policies have been configured"} #endif AUTHPOLHASH.Length
    } #end if no LDAP configures

WriteWordLine 0 0 " "
$Table = $null
#endregion CAG Authentication Radius Policies  
        
    #region CAG Authentication SAML IDP Policies             
        WriteWordLine 3 0 "Authentication SAML IDP Policies"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $vpnvserversamlidppols = Get-vNetScalerObject -ResourceType vpnvserver_authenticationsamlidppolicy_binding -name $vpnvserver.Name
        $errorcode = $vpnvserversamlidppols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No SAML IDP Policies have been configured"} else { #Uses the mentioned error code to determine existency of policy

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $AUTHPOLSAMLIDPHASH = @(); 

             foreach ($vpnvserversamlidppol in $vpnvserversamlidppols) {                
                $AUTHPOLSAMLIDPHASH += @{
                    Name = $vpnvserversamlidppol.policy;
                    Priority = $vpnvserversamlidppol.priority;
                }
            }

        if ($AUTHPOLSAMLIDPHASH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $AUTHPOLSAMLIDPHASH;
                Columns = "Name","Priority";
                Headers = "Name","Priority";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

            FindWordDocumentEnd;
                
        } else { WriteWordLine 0 0 "No SAML IDP Policies have been configured"}
    } 
WriteWordLine 0 0 " "
$Table = $null
#endregion CAG Authentication SAML IDP Policies  
    
    #region CAG Session Policies        
       
        WriteWordLine 3 0 "Session Policies"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $vpnvserversespols = Get-vNetScalerObject -ResourceType vpnvserver_vpnsessionpolicy_binding -name $vpnvserver.Name
        $errorcode = $vpnvserversespols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No Session Policies have been configured"} else { #Uses the mentioned error code to determine existency of policy
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $SESSIONPOLH = @();     

            foreach ($vpnvserversespol in $vpnvserversespols) {                
                $SESSIONPOLH += @{
                    Name = $vpnvserversespol.policy;
                    Priority = $vpnvserversespol.priority;
                }
            }
        }
            
        if ($SESSIONPOLH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $SESSIONPOLH;
                    Columns = "Name","Priority";
                    Headers = "Policy Name","Priority";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;

            FindWordDocumentEnd;
        } else { WriteWordLine 0 0 "No Session Policies have been configured"} #endif SESSIONPOLHASH.Length

WriteWordLine 0 0 " "
$Table = $null
    #endregion CAG Session Policies 
    
    #region CAG STA Policies        
       
        WriteWordLine 3 0 "Secure Ticket Authority"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $vpnvserverstas = Get-vNetScalerObject -ResourceType vpnvserver_staserver_binding -name $vpnvserver.Name
        $errorcode = $vpnvserverstas.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No Secure Ticket Authority has been configured"} else { #Uses the mentioned error code to determine existency of policy
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $STAPOLH = @();     

            foreach ($vpnvserversta in $vpnvserverstas) {                
                $STAPOLH += @{
                    Name = $vpnvserversta.staserver;
                    STAID = $vpnvserversta.staauthid;
                    STATYPE = $vpnvserversta.staaddresstype;
                } 
            }
        }
            
        if ($STAPOLH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $STAPOLH;
                Columns = "Name","STAID","STATYPE";
                Headers = "Secure Ticket Authority","Authentication ID","Address Type";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;

            FindWordDocumentEnd;
        } else { WriteWordLine 0 0 "No STA Policies have been configured"} #

WriteWordLine 0 0 " "
$Table = $null
    #endregion CAG STA Policies 

    #region CAG Cache Policies        
       
        WriteWordLine 3 0 "Cache Policies"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $vpnvservercachepols = Get-vNetScalerObject -ResourceType vpnvserver_cachepolicy_binding -name $vpnvserver.Name
        $errorcode = $vpnvservercachepols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No Session Policies have been configured"} else { #Uses the mentioned error code to determine existency of policy
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $SESSIONCACHEPOLH = @();     

            foreach ($vpnvservercachepol in $vpnvservercachepols) {                
                $SESSIONCACHEPOLH += @{
                    Name = $vpnvservercachepol.policy;
                    Priority = $vpnvservercachepol.priority;
                }
            }
        }
            
        if ($SESSIONCACHEPOLH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $SESSIONCACHEPOLH;
                    Columns = "Name","Priority";
                    Headers = "Policy Name","Priority";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;

            FindWordDocumentEnd;
        } else { WriteWordLine 0 0 "No Cache Policies have been configured"} #

WriteWordLine 0 0 " "
$Table = $null
    #endregion CAG Cache Policies 

    #region CAG Responder Policies        
       
        WriteWordLine 3 0 "Responder Policies"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $vpnvserverrespols = Get-vNetScalerObject -ResourceType vpnvserver_responderpolicy_binding -name $vpnvserver.Name
        $errorcode = $vpnvserverrespols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No Responder Policies have been configured"} else { #Uses the mentioned error code to determine existency of policy
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $RESPOLH = @();     

            foreach ($vpnvserverrespol in $vpnvserverrespols) {                
                $RESPOLH += @{
                    Name = $vpnvserverrespol.policy;
                    Priority = $vpnvserverrespol.priority;
                }
            }
        }
            
        if ($RESPOLH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $RESPOLH;
                    Columns = "Name","Priority";
                    Headers = "Policy Name","Priority";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;

            FindWordDocumentEnd;
        } else { WriteWordLine 0 0 "No Responder Policies have been configured"}

WriteWordLine 0 0 " "
$Table = $null
    #endregion CAG Responder Policies 

        #region CAG Rewrite Policies        
       
        WriteWordLine 3 0 "Rewrite Policies"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $vpnvserverrwpols = Get-vNetScalerObject -ResourceType vpnvserver_rewritepolicy_binding -name $vpnvserver.Name
        $errorcode = $vpnvserverrwpols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No Rewrite Policies have been configured"} else { #Uses the mentioned error code to determine existency of policy
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $RWPOLH = @();     

            foreach ($vpnvserverrwpol in $vpnvserverrwpols) {                
                $RWPOLH += @{
                    Name = $vpnvserverrwpol.policy;
                    Priority = $vpnvserverrwpol.priority;
                }
            }
        }
            
        if ($RWPOLH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $RWPOLH;
                    Columns = "Name","Priority";
                    Headers = "Policy Name","Priority";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;

            FindWordDocumentEnd;
        } else { WriteWordLine 0 0 "No Rewrite Policies have been configured"}

WriteWordLine 0 0 " "
$Table = $null
    #endregion CAG Rewrite Policies 

    #region CAG SSL Configuration        
       
        WriteWordLine 3 0 "SSL Parameters"
        WriteWordLine 0 0 " "

        $cagsslparameters = Get-vNetScalerObject -ResourceType sslvserver -Name $vpnvserver.Name;

## IB - Create an array of hashtables to store our columns.
## IB - about column names as we'll utilise a -List(view)!
[System.Collections.Hashtable[]] $CAGSSLPARAMSH = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "Clear Text Port"; Column2 = $cagsslparameters.cleartextport; }
    @{ Column1 = "Diffe-Hellman Key Exchange"; Column2 = $cagsslparameters.dh; }
    @{ Column1 = "Diffe-Hellman Key File"; Column2 = $cagsslparameters.dhfile; }
    @{ Column1 = "Diffe-Hellman Refresh Count"; Column2 = $cagsslparameters.dhcount; }
    @{ Column1 = "Enable DH Key Expire Size Limit"; Column2 = $cagsslparameters.dhkeyexpsizelimit; }
    @{ Column1 = "Enable Ephemeral RSA"; Column2 = $cagsslparameters.ersa; }
    @{ Column1 = "Ephemeral RSA Refresh Count"; Column2 = $cagsslparameters.ersacount; }
    @{ Column1 = "Allow session re-use"; Column2 = $cagsslparameters.sessreuse; }
    @{ Column1 = "Session Time-out"; Column2 = $cagsslparameters.sesstimeout; }
    @{ Column1 = "Enable Cipher Redirect"; Column2 = $cagsslparameters.cipherredirect; }
    @{ Column1 = "Cipher Redirect URL"; Column2 = $cagsslparameters.cipherurl; }
    @{ Column1 = "SSLv2 Redirect"; Column2 = $cagsslparameters.sslv2redirect; }
    @{ Column1 = "SSLv2 Redirect URL"; Column2 = $cagsslparameters.sslv2url; }
    @{ Column1 = "Enable Client Authentication"; Column2 = $cagsslparameters.clientauth; }
    @{ Column1 = "Client Certificates"; Column2 = $cagsslparameters.clientcert; }
    @{ Column1 = "SSL Redirect"; Column2 = $cagsslparameters.sslredirect; }
    @{ Column1 = "SSL 2"; Column2 = $cagsslparameters.ssl2; }
    @{ Column1 = "SSL 3"; Column2 = $cagsslparameters.ssl3; }
    @{ Column1 = "TLS 1"; Column2 = $cagsslparameters.tls1; }
    @{ Column1 = "TLS 1.1"; Column2 = $cagsslparameters.tls11; }
    @{ Column1 = "TLS 1.2"; Column2 = $cagsslparameters.tls12; }
    @{ Column1 = "Server Name Indication (SNI)"; Column2 = $cagsslparameters.snienable; }
    @{ Column1 = "PUSH Encryption Trigger"; Column2 = $cagsslparameters.pushenctrigger; }
    @{ Column1 = "Send Close-Notify"; Column2 = $cagsslparameters.sendclosenotify; }
    @{ Column1 = "DTLS Profile"; Column2 = $cagsslparameters.dtlsprofilename; }
    @{ Column1 = "SSL Profile"; Column2 = $cagsslparameters.sslprofile; }

);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $CAGSSLPARAMSH;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "

    #endregion CAG SSL Configuration

        #region CAG SSL Ciphers             
        WriteWordLine 3 0 "SSL Ciphers"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $vpncipherbinds = Get-vNetScalerObject -ResourceType sslvserver_sslciphersuite_binding -name $vpnvserver.Name
        $errorcode = $vpncipherbinds.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No SSL Ciphers have been bound."} else { #Uses the mentioned error code to determine existency of policy

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $CIPHERSH = @(); 

             foreach ($vpncipherbind in $vpncipherbinds) {                
                $CIPHERSH += @{
                    Name = $vpncipherbind.ciphername;
                    Description = $vpncipherbind.description;

                    
                } # end Hasthable $INTIPSH
            }# end foreach

        if ($CIPHERSH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $CIPHERSH;
                Columns = "Name","Description";
                Headers = "Name","Description";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

            FindWordDocumentEnd;
                
        } else { WriteWordLine 0 0 "No SSL Ciphers have been configured"} #endif AUTHPOLHASH.Length
    } #end if
WriteWordLine 0 0 " "
#endregion CAG SSL Ciphers

    #region CAG Intranet Applications             
        WriteWordLine 3 0 "Intranet Applications"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $vpnintappbinds = Get-vNetScalerObject -ResourceType vpnvserver_vpnintranetapplication_binding -name $vpnvserver.Name
        $errorcode = $vpnintappbinds.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No Intranet Applications have been bound."} else { #Uses the mentioned error code to determine existency of policy

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $INTAPPSH = @(); 

             foreach ($vpnintappbind in $vpnintappbinds) {                
                $INTAPPSH += @{
                    Name = $vpnintappbind.intranetapplication;
                } # end Hasthable $INTAPPSH
            }# end foreach

        if ($INTAPPSH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $INTAPPSH;
                Columns = "Name";
                Headers = "Name";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

            FindWordDocumentEnd;
                
        } else { WriteWordLine 0 0 "No Intranet Applications have been configured"} #endif AUTHPOLHASH.Length
    } #end if
WriteWordLine 0 0 " "
#endregion CAG Intranet Applications 

    #region CAG Intranet IPs             
        WriteWordLine 3 0 "Intranet IP's"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $vpnintipbinds = Get-vNetScalerObject -ResourceType vpnvserver_vpnintranetip_binding -name $vpnvserver.Name
        $errorcode = $vpnintipbinds.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No Intranet IP's have been bound."} else { #Uses the mentioned error code to determine existency of policy

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $INTIPSH = @(); 

             foreach ($vpnintipbind in $vpnintipbinds) {                
                $INTIPSH += @{
                    Name = $vpnintipbind.intranetip;
                    NetMask = $vpnintipbind.netmask;
                } # end Hasthable $INTIPSH
            }# end foreach

        if ($INTIPSH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $INTIPSH;
                Columns = "Name","NetMask";
                Headers = "Name","NetMask";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

            FindWordDocumentEnd;
                
        } else { WriteWordLine 0 0 "No Intranet IP's have been configured"} #endif AUTHPOLHASH.Length
    } #end if
WriteWordLine 0 0 " "
#endregion CAG Intranet IPs 

    #region CAG Bookmarks             
        WriteWordLine 3 0 "Bookmarks"
        WriteWordLine 0 0 " "
        $errorcode = 1 #Set Error code to 1
        $vpninturlbinds = Get-vNetScalerObject -ResourceType vpnvserver_vpnurl_binding -name $vpnvserver.Name
        $errorcode = $vpninturlbinds.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No Bookmarks's have been bound."} else { #Uses the mentioned error code to determine existency of policy

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $INTURLSH = @(); 

             foreach ($vpninturlbind in $vpninturlbinds) {                
                $INTURLSH += @{
                    Name = $vpninturlbind.urlname;
                    
                } # end Hasthable $INTIPSH
            }# end foreach

        if ($INTURLSH.Length -gt 0) {
            ## IB - Add the table to the document (only if not null!
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $INTURLSH;
                Columns = "Name";
                Headers = "Name";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

            FindWordDocumentEnd;
                
        } else { WriteWordLine 0 0 "No Bookmarks have been configured"} #endif AUTHPOLHASH.Length
    } #end if
WriteWordLine 0 0 " "
#endregion CAG Bookmarks

    $selection.InsertNewPage()
    }
}


#endregion CAG vServers

#region CAG Session Policies
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters Citrix NetScaler (Access) Gateway Policies"
##WriteWordLine 2 0 "NetScaler Gateway Policies"

##WriteWordLine 0 0 " "
WriteWordLine 2 0 "NetScaler Gateway Session Policies"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `tNetScaler Gateway Session Policies"

$vpnsessionpolicies = Get-vNetScalerObject -Container config -Object vpnsessionpolicy;

foreach ($vpnsessionpolicy in $vpnsessionpolicies) {
    $sesspolname = $vpnsessionpolicy.name
    WriteWordLine 3 0 "NetScaler Gateway Session Policy: $sesspolname";
    WriteWordLine 0 0 " "

    ## IB - Create an array of hashtables to store our columns. Note: If we need the
    ## IB - headers to include spaces we can override these at table creation time.
    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = @{
            ## IB - Each hashtable is a separate row in the table!
            NAME = $vpnsessionpolicy.name;
            RULE = $vpnsessionpolicy.rule;
            ACTION = $vpnsessionpolicy.action;
            ACTIVE = $vpnsessionpolicy.activepolicy;
        }
        Columns = "NAME","RULE","ACTION","ACTIVE";
        Headers = "Policy Name","Rule","Action","Active";
        AutoFit = $wdAutoFitContent;
        Format = -235; ## IB - Word constant for Light List Accent 5
    }

    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;
    WriteWordLine 0 0 " "
}

#region alwayson policies
WriteWordLine 0 0 " "
WriteWordLine 2 0 "NetScaler Gateway AlwaysON Policies"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `tNetScaler Gateway AlwaysON Policies"

$vpnalwaysonpolicies = Get-vNetScalerObject -Container config -Object vpnalwaysonprofile;

If (!$vpnalwaysonpolicies) {
    WriteWordLine 0 0 "No AlwaysON Policies have been configured. "
}

foreach ($vpnalwaysonpolicy in $vpnalwaysonpolicies) {
$policynameAO = $vpnalwaysonpolicy.name
    WriteWordLine 3 0 "NetScaler Gateway AlwaysON Policy: $policynameAO";

    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $AOPOLCONFH = @(
    @{ Description = "Description"; Value = "Configuration"; }
    @{ Description = "Location Based VPN"; Value = $vpnalwaysonpolicy.locationbasedvpn; }
    @{ Description = "Client Control"; Value = $vpnalwaysonpolicy.clientcontrol; }
    @{ Description = "Network Access On VPN Failure"; Value = $vpnalwaysonpolicy.networkaccessonvpnfailure; }
    );

    if ($AOPOLCONFH.Length -gt 0){

    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = $AOPOLCONFH;
        Columns = "Description","Value";
        AutoFit = $wdAutoFitContent
        Format = -235; ## IB - Word constant for Light List Accent 5
    }
    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -NoGridLines -List;

	FindWordDocumentEnd;
	$TableRange = $Null
	$Table = $Null
    } else {
    WriteWordLine 0 0 "No AlwaysON Policies have been configured. "
    }
}
#endregion alwayson policies
WriteWordLine 0 0 " "
#endregion CAG Policies

#region CAG Session Actions
WriteWordLine 0 0 " "
WriteWordLine 2 0 "NetScaler Gateway Session Actions"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `tNetScaler Gateway Session Actions"

$vpnsessionactions = Get-vNetScalerObject -Container config -Object vpnsessionaction;

If (!$vpnsessionactions) {

WordWriteLine 0 0 "There are no Netscaler Gateway Session Actions configured."
WordWriteLine 0 0 " "

}

foreach ($vpnsessionaction in $vpnsessionactions) {
    $sessactname = $vpnsessionaction.name
    WriteWordLine 3 0 "NetScaler Gateway Session Action: $sessactname";
    WriteWordLine 0 0 " "
#region ClientExperience

    WriteWordLine 4 0 "Client Experience"
    WriteWordLine 0 0 " "

    ## IB - Add the table to the document, splatting the parameters
    #$Table = AddWordTable @Params -NoGridLines;
    #FindWordDocumentEnd;
    #WriteWordLine 0 0 " "



[System.Collections.Hashtable[]] $VPNACTCEXH = @(
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "Homepage"; Column2 = $vpnsessionaction.emailhome; }
    @{ Column1 = "URL for Web Based Email"; Column2 = $vpnsessionaction.sesstimeout; }
    @{ Column1 = "Session Time-Out"; Column2 = $vpnsessionaction.sesstimeout; }
    @{ Column1 = "Client-Idle Time-Out"; Column2 = $vpnsessionaction.clientidletimeoutwarning; }
    @{ Column1 = "Single Sign-On to Web Applications"; Column2 = $vpnsessionaction.sso; }
    @{ Column1 = "Credential Index"; Column2 = $vpnsessionaction.ssocredential; }
    @{ Column1 = "Single Sign-On with Windows"; Column2 = $vpnsessionaction.windowsautologon; }
    @{ Column1 = "Split Tunnel"; Column2 = $vpnsessionaction.splittunnel; }
    @{ Column1 = "Local LAN Access"; Column2 = $vpnsessionaction.locallanaccess; }
    @{ Column1 = "Plug-in Type"; Column2 = $vpnsessionaction.windowsclienttype; }
    @{ Column1 = "Windows Plugin Upgrade"; Column2 = $vpnsessionaction.windowspluginupgrade; }
    @{ Column1 = "MAC Plugin Upgrade"; Column2 = $vpnsessionaction.macpluginupgrade; }
    @{ Column1 = "Linux Plugin Upgrade"; Column2 = $vpnsessionaction.linuxpluginupgrade; }
    @{ Column1 = "AlwaysON Profile Name"; Column2 = $vpnsessionaction.alwaysonprofilename; }
    @{ Column1 = "Clientless Access"; Column2 = $vpnsessionaction.clientlessvpnmode; }
    @{ Column1 = "Clientless URL Encoding"; Column2 = $vpnsessionaction.clientlessmodeurlencoding; }
    @{ Column1 = "Clientless Persistent Cookie"; Column2 = $vpnsessionaction.clientlesspersistentcookie; }
    @{ Column1 = "Single Sign-on to Web Applications"; Column2 = $vpnsessionaction.sso; }
    @{ Column1 = "Credential Index"; Column2 = $vpnsessionaction.ssocredential; }
    @{ Column1 = "KCD Account"; Column2 = $vpnsessionaction.kcdaccount; }
    @{ Column1 = "Single Sign-on with Windows"; Column2 = $vpnsessionaction.windowsautologon; }
    @{ Column1 = "Client Cleanup Prompt"; Column2 = $vpnsessionaction.clientcleanupprompt; }
    @{ Column1 = "UI Theme"; Column2 = $vpnsessionaction.uitheme; }
    @{ Column1 = "Login Script"; Column2 = $vpnsessionaction.loginscript; }
    @{ Column1 = "Logout Script"; Column2 = $vpnsessionaction.logoutscript; }
    @{ Column1 = "Application Token Timeout"; Column2 = $vpnsessionaction.apptokentimeout; }
    @{ Column1 = "MDX Token Timeout"; Column2 = $vpnsessionaction.mdxtokentimeout; }
    @{ Column1 = "Allow Users to Change Log Levels"; Column2 = $vpnsessionaction.clientconfiguration; }
    @{ Column1 = "Allow access to private network IP addresses only"; Column2 = $vpnsessionaction.windowsclienttype; }
    @{ Column1 = "Client Choices"; Column2 = $vpnsessionaction.clientchoices; }
    @{ Column1 = "Show VPN Plugin icon"; Column2 = $vpnsessionaction.iconwithreceiver; }
   

);


## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $VPNVSERVERH;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
$Table = $null


#endregion ClientExperience

#region Security
    
    WriteWordLine 4 0 "Security"
    WriteWordLine 0 0 " "
    ## IB - Create an array of hashtables to store our columns. Note: If we need the
    ## IB - headers to include spaces we can override these at table creation time.
    ## IB - Create the parameters to pass to the AddWordTable function
    #$Params = $null
    #$Params = @{
    #    Hashtable = @{
    #        ## IB - Each hashtable is a separate row in the table!
    #        DEFAUTH = $vpnsessionaction.defaultauthorizationaction;
    #        SECBRW = $vpnsessionaction.securebrowse;
    #    }
    #    Columns = "DEFAUTH","SECBRW";
    #    Headers = "Default Authorization Action","Secure Browse";
    #    AutoFit = $wdAutoFitContent;
    #    Format = -235; ## IB - Word constant for Light List Accent 5
    #}

    ## IB - Add the table to the document, splatting the parameters
    #$Table = AddWordTable @Params -NoGridLines;
    #FindWordDocumentEnd;
    #WriteWordLine 0 0 " "

    [System.Collections.Hashtable[]] $VPNACTSECH = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "Default Authorization Action"; Column2 = $vpnsessionaction.defaultauthorizationaction; }
    @{ Column1 = "Secure Browse"; Column2 = $vpnsessionaction.securebrowse; }
    @{ Column1 = "Client Security Check String"; Column2 = $vpnsessionaction.clientsecurity; }
    @{ Column1 = "Quarantine Group"; Column2 = $vpnsessionaction.clientsecuritygroup; }
    @{ Column1 = "Error Message"; Column2 = $vpnsessionaction.clientsecuritymessage; }
    @{ Column1 = "Enable Client Security Logging"; Column2 = $vpnsessionaction.clientsecuritylog; }
    @{ Column1 = "Authorization Groups"; Column2 = $vpnsessionaction.authorizationgroup; }
    @{ Column1 = "Groups allowed to login"; Column2 = $vpnsessionaction.allowedlogingroups; }
   

);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $VPNACTSECH;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
$Table = $null

#endregion Security

#region Published Applications  

    WriteWordLine 4 0 "Published Applications"
    WriteWordLine 0 0 " "
    ## IB - Create an array of hashtables to store our columns. Note: If we need the
    ## IB - headers to include spaces we can override these at table creation time.
    ## IB - Create the parameters to pass to the AddWordTable function
    #$Params = $null
    #$Params = @{
    #    Hashtable = @{
    #        ## IB - Each hashtable is a separate row in the table!
    #        ICAPROXY = $vpnsessionaction.icaproxy;
    #        WIMODE = $vpnsessionaction.wiportalmode;
    #        SSO = $vpnsessionaction.sso;
    #    }
    #    Columns = "ICAPROXY","WIMODE","SSO";
    #    Headers = "ICA Proxy","Web Interface Portal Mode","Single Sign-On Domain";
    #    AutoFit = $wdAutoFitContent;
    #    Format = -235; ## IB - Word constant for Light List Accent 5
    #}

    ## IB - Add the table to the document, splatting the parameters
    #$Table = AddWordTable @Params -NoGridLines;
    #FindWordDocumentEnd;
    #WriteWordLine 0 0 " "

        [System.Collections.Hashtable[]] $VPNACTPAH = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "ICA Proxy"; Column2 = $vpnsessionaction.icaproxy; }
    @{ Column1 = "Web Interface Address"; Column2 = $vpnsessionaction.wihome; }
    @{ Column1 = "Web Interface Address Type"; Column2 = $vpnsessionaction.wihomeaddresstype; }
    @{ Column1 = "Single Sign-on Domain"; Column2 = $vpnsessionaction.sso; }
    @{ Column1 = "Citrix Receiver Home Page"; Column2 = $vpnsessionaction.citrixreceiverhome; }
    @{ Column1 = "Account Services Address"; Column2 = $vpnsessionaction.storefronturl; }

   

);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $VPNACTPAH;
    Columns = "Column1","Column2";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

$Table = AddWordTable @Params -List -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
$Table = $null

#end region Published Applications
    $selection.InsertNewPage()
}

    #endregion CAG Session Policies



#endregion CAG Actions

#endregion Citrix NetScaler Gateway

#region NetScaler Monitors
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Monitors"

WriteWordLine 1 0 "NetScaler Monitors"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `t`tTable: Write NetScaler Monitors Table"

$monitorcounter = Get-vNetScalerObjectCount -Container config -Object lbmonitor; 
$monitorcount = $monitorcounter.__count
$monitors = Get-vNetScalerObject -Container config -Object lbmonitor;
   
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $MONITORSH = @();

foreach ($MONITOR in $MONITORS) {
    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be about 400 characters wide!

    $MONITORSH += @{
            NAME = $MONITOR.monitorname;
            Type = $MONITOR.type;
            DestinationPort = $monitor.destport;
            Interval = $monitor.interval;
            TimeOut = $monitor.resptimeout;
            }
        }

    if ($MONITORSH.Length -gt 0) {
        $Params = $null
        $Params = @{
            Hashtable = $MONITORSH;
            Columns = "NAME","Type","DestinationPort","Interval","TimeOut";
            Headers = "Monitor Name","Type","Destination Port","Interval","Time-Out";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
            }
        $Table = AddWordTable @Params -NoGridLines;
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
        }


$selection.InsertNewPage()

#endregion NetScaler Monitors

#region NetScaler Policies
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Policies"

WriteWordLine 1 0 "NetScaler Policies"
WriteWordLine 0 0 " "
#region Pattern Set Policies
WriteWordLine 2 0 "NetScaler Pattern Set Policies"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `tTable: NetScaler Pattern Set Policies"

$pattsetpolicies = Get-vNetScalerObject -Container config -Object policypatset;

[System.Collections.Hashtable[]] $PATSETS = @();
foreach ($pattsetpolicy in $pattsetpolicies) {
    #$pattsetpolicy.name
    $PATSETS += @{
        PATSET = $pattsetpolicy.name; 
    }
}
    
## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $PATSETS;
    Columns = "PATSET";
    Headers = "Pattern Set Policy";
    Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
    AutoFit = $wdAutoFitContent;
    }
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params -NoGridLines;

FindWordDocumentEnd;
WriteWordLine 0 0 " "
#endregion Pattern Set Policies

#region Responder Policies
WriteWordLine 2 0 "NetScaler Responder Policies"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `tTable: NetScaler Responder Policies"

$responderpolicies = Get-vNetScalerObject -Container config -Object responderpolicy;

[System.Collections.Hashtable[]] $RESPPOL = @();
foreach ($responderpolicy in $responderpolicies) {
    $RESPPOL += @{
        RESPOLNAME = $responderpolicy.name;
        RULE = $responderpolicy.rule;
        ACTION = $responderpolicy.action;
        ACTIVE = $responderpolicy.activepolicy;
    }
}
    
## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $RESPPOL;
    Columns = "RESPOLNAME","RULE","ACTION","ACTIVE";
    Headers = "Responder Policy","Rule","Action","Active Policy";
    Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
    AutoFit = $wdAutoFitContent;
    }
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params -NoGridLines;

FindWordDocumentEnd;
WriteWordLine 0 0 " "
#endregion Responder Policies

#region Rewrite Policies
WriteWordLine 2 0 "NetScaler Rewrite Policies"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `tTable: NetScaler Rewrite Policies"

$rewritepolicies = Get-vNetScalerObject -Container config -Object rewritepolicy;

[System.Collections.Hashtable[]] $RWPPOL = @();
foreach ($rewritepolicy in $rewritepolicies) {
    $RWPPOL += @{
        RWPOLNAME = $rewritepolicy.name;
        RULE = $rewritepolicy.rule;
        ACTION = $rewritepolicy.action;
        ACTIVE = $rewritepolicy.activepolicy;
    }
}
    
## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $RWPPOL;
    Columns = "RWPOLNAME","RULE","ACTION","ACTIVE";
    Headers = "Rewrite Policy","Rule","Action","Active Policy";
    Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
    AutoFit = $wdAutoFitContent;
    }
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params -NoGridLines;

FindWordDocumentEnd;
WriteWordLine 0 0 " "
#endregion Rewrite Policies

#Endregion NetScaler Policies

#endregion New functionality here

#region NetScaler Actions
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Actions"

WriteWordLine 1 0 "NetScaler Actions"
WriteWordLine 0 0 " "
#region Responder Action
WriteWordLine 2 0 "NetScaler Responder Action"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `tTable: NetScaler Responder Action"
$responderactions = Get-vNetScalerObject -Container config -Object responderaction;

## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $ACTRESH = @();

foreach ($responderaction in $responderactions) {
             
    $ACTRESH += @{ 
        Responder = $responderaction.name; 
        Type = $responderaction.type;
        Target = $responderaction.target;
        RESPST = $responderaction.responsestatuscode
        }
}

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $ACTRESH;
    Columns = "Responder","Type","Target","RESPST";
    Headers = "Responder Policy","Type","Target","Response Status Code";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
    }
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params -NoGridLines;

FindWordDocumentEnd;
WriteWordLine 0 0 " "

#endregion Responder Action

#region Rewrite Action
WriteWordLine 2 0 "NetScaler Rewrite Action"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `tTable: NetScaler Rewrite Action"

$rewriteactions = Get-vNetScalerObject -Container config -Object rewriteaction;

## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $ACTRWH = @();

foreach ($rewriteaction in $rewriteactions) {
    $ACTRWH += @{ 
        REWRITE = $rewriteaction.name; 
        Type = $rewriteaction.type;
        Target = $rewriteaction.target;
        STRING = $rewriteaction.stringbuilderexpr
        }
}

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $ACTRWH;
    Columns = "REWRITE","Type","Target","STRING";
    Headers = "Rewrite Policy","Type","Target","String";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
    }
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params -NoGridLines;

FindWordDocumentEnd;
WriteWordLine 0 0 " "

$selection.InsertNewPage()

#endregion Rewrite Action

#endregion NetScaler Actions

#region NetScaler Profiles
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Profiles"

WriteWordLine 1 0 "NetScaler Profiles"
WriteWordLine 0 0 " "
#region NetScaler TCP Profiles

WriteWordLine 2 0 "NetScaler TCP Profiles"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `t`tTable: Write NetScaler TCP Profiles Table"

$tcpprofiles = Get-vNetScalerObject -Container config -Object nstcpprofile;

## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $TCPPROFILESH = @();

foreach ($tcpprofile in $tcpprofiles) {
    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be about 400 characters wide!

    $TCPPROFILESH += @{
            TCP = $tcpprofile.name;
            WS = $tcpprofile.ws;
            SACK = $tcpprofile.sack;
            NAGLE = $tcpprofile.NAGLE;
            MSS = $tcpprofile.MSS;
        }
}

if ($TCPPROFILESH.Length -gt 0) {
    $Params = $null
    $Params = @{
        Hashtable = $TCPPROFILESH;
        Columns = "TCP","WS","SACK","NAGLE","MSS";
        Headers = "TCP","WS","SACK","NAGLE","MSS";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
        }
    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;
    WriteWordLine 0 0 " "
}
  
$selection.InsertNewPage()


#endregion NetScaler TCP Profiles

#region NetScaler HTTP Profiles

WriteWordLine 2 0 "NetScaler HTTP Profiles"
WriteWordLine 0 0 " "

Write-Verbose "$(Get-Date): `t`tTable: Write NetScaler HTTP Profiles Table"

$httprofiles = Get-vNetScalerObject -Container config -Object nshttpprofile;

## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $HTTPPROFILESH = @();

foreach ($httprofile in $httprofiles) {
    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be about 400 characters wide!

    $HTTPPROFILESH += @{
            HTTP = $httprofile.name;
            Drop = $httprofile.dropinvalreqs;
            HTTP2 = $httprofile.http2;
        }
}

if ($HTTPPROFILESH.Length -gt 0) {
    $Params = $null
    $Params = @{
        Hashtable = $HTTPPROFILESH;
        Columns = "HTTP","Drop","HTTP2";
        Headers = "HTTP Profile","Drop Invalid Connections","HTTP2";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
        }
    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;
    WriteWordLine 0 0 " "
}
  
$selection.InsertNewPage()


#endregion NetScaler HTTP Profiles

#endregion NetScaler Profiles

#endregion NetScaler Documentation Script Complete

#region script template 2

Write-Verbose "$(Get-Date): Finishing up document"
#end of document processing

###Change the two lines below for your script
$AbstractTitle = "NetScaler Documentation Report"
$SubjectTitle = "NetScaler Documentation Report"
UpdateDocumentProperties $AbstractTitle $SubjectTitle

ProcessDocumentOutput

ProcessScriptEnd
#recommended by webster
#$error
#endregion script template 2

# SIG # Begin signature block
# MIIgCgYJKoZIhvcNAQcCoIIf+zCCH/cCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUuh28FukGcIE5d+hfGRHs8+84
# rq2gghtxMIIDtzCCAp+gAwIBAgIQDOfg5RfYRv6P5WD8G/AwOTANBgkqhkiG9w0B
# AQUFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVk
# IElEIFJvb3QgQ0EwHhcNMDYxMTEwMDAwMDAwWhcNMzExMTEwMDAwMDAwWjBlMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3Qg
# Q0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCtDhXO5EOAXLGH87dg
# +XESpa7cJpSIqvTO9SA5KFhgDPiA2qkVlTJhPLWxKISKityfCgyDF3qPkKyK53lT
# XDGEKvYPmDI2dsze3Tyoou9q+yHyUmHfnyDXH+Kx2f4YZNISW1/5WBg1vEfNoTb5
# a3/UsDg+wRvDjDPZ2C8Y/igPs6eD1sNuRMBhNZYW/lmci3Zt1/GiSw0r/wty2p5g
# 0I6QNcZ4VYcgoc/lbQrISXwxmDNsIumH0DJaoroTghHtORedmTpyoeb6pNnVFzF1
# roV9Iq4/AUaG9ih5yLHa5FcXxH4cDrC0kqZWs72yl+2qp/C3xag/lRbQ/6GW6whf
# GHdPAgMBAAGjYzBhMA4GA1UdDwEB/wQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB0G
# A1UdDgQWBBRF66Kv9JLLgjEtUYunpyGd823IDzAfBgNVHSMEGDAWgBRF66Kv9JLL
# gjEtUYunpyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEAog683+Lt8ONyc3pklL/3
# cmbYMuRCdWKuh+vy1dneVrOfzM4UKLkNl2BcEkxY5NM9g0lFWJc1aRqoR+pWxnmr
# EthngYTffwk8lOa4JiwgvT2zKIn3X/8i4peEH+ll74fg38FnSbNd67IJKusm7Xi+
# fT8r87cmNW1fiQG2SVufAQWbqz0lwcy2f8Lxb4bG+mRo64EtlOtCt/qMHt1i8b5Q
# Z7dsvfPxH2sMNgcWfzd8qVttevESRmCD1ycEvkvOl77DZypoEd+A5wwzZr8TDRRu
# 838fYxAe+o0bJW1sj6W3YQGx0qMmoRBxna3iw/nDmVG3KwcIzi7mULKn+gpFL6Lw
# 8jCCBTAwggQYoAMCAQICEAQJGBtf1btmdVNDtW+VUAgwDQYJKoZIhvcNAQELBQAw
# ZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQ
# d3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBS
# b290IENBMB4XDTEzMTAyMjEyMDAwMFoXDTI4MTAyMjEyMDAwMFowcjELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUg
# U2lnbmluZyBDQTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAPjTsxx/
# DhGvZ3cH0wsxSRnP0PtFmbE620T1f+Wondsy13Hqdp0FLreP+pJDwKX5idQ3Gde2
# qvCchqXYJawOeSg6funRZ9PG+yknx9N7I5TkkSOWkHeC+aGEI2YSVDNQdLEoJrsk
# acLCUvIUZ4qJRdQtoaPpiCwgla4cSocI3wz14k1gGL6qxLKucDFmM3E+rHCiq85/
# 6XzLkqHlOzEcz+ryCuRXu0q16XTmK/5sy350OTYNkO/ktU6kqepqCquE86xnTrXE
# 94zRICUj6whkPlKWwfIPEvTFjg/BougsUfdzvL2FsWKDc0GCB+Q4i2pzINAPZHM8
# np+mM6n9Gd8lk9ECAwEAAaOCAc0wggHJMBIGA1UdEwEB/wQIMAYBAf8CAQAwDgYD
# VR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHkGCCsGAQUFBwEBBG0w
# azAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUF
# BzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVk
# SURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8vY3JsNC5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRw
# Oi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3Js
# ME8GA1UdIARIMEYwOAYKYIZIAYb9bAACBDAqMCgGCCsGAQUFBwIBFhxodHRwczov
# L3d3dy5kaWdpY2VydC5jb20vQ1BTMAoGCGCGSAGG/WwDMB0GA1UdDgQWBBRaxLl7
# KgqjpepxA8Bg+S32ZXUOWDAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823I
# DzANBgkqhkiG9w0BAQsFAAOCAQEAPuwNWiSz8yLRFcgsfCUpdqgdXRwtOhrE7zBh
# 134LYP3DPQ/Er4v97yrfIFU3sOH20ZJ1D1G0bqWOWuJeJIFOEKTuP3GOYw4TS63X
# X0R58zYUBor3nEZOXP+QsRsHDpEV+7qvtVHCjSSuJMbHJyqhKSgaOnEoAjwukaPA
# JRHinBRHoXpoaK+bp1wgXNlxsQyPu6j4xRJon89Ay0BEpRPw5mQMJQhCMrI2iiQC
# /i9yfhzXSUWW6Fkd6fp0ZGuy62ZD2rOwjNXpDd32ASDOmTFjPQgaGLOBm0/GkxAG
# /AeB+ova+YJJ92JuoVP6EpQYhS6SkepobEQysmah5xikmmRR7zCCBT8wggQnoAMC
# AQICEAU1dkJYhLmnBcR7Ly9I4oUwDQYJKoZIhvcNAQELBQAwcjELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2ln
# bmluZyBDQTAeFw0xNjEwMTgwMDAwMDBaFw0xNzEwMjMxMjAwMDBaMHwxCzAJBgNV
# BAYTAlVTMQswCQYDVQQIEwJUTjESMBAGA1UEBxMJVHVsbGFob21hMSUwIwYDVQQK
# ExxDYXJsIFdlYnN0ZXIgQ29uc3VsdGluZywgTExDMSUwIwYDVQQDExxDYXJsIFdl
# YnN0ZXIgQ29uc3VsdGluZywgTExDMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIB
# CgKCAQEAsEaSbYa/BsjQPvGR3Zbaiq/LathAtAbbO4mTyf+zws81cGyKtI4NkNCT
# qPsKORH9hxw8qqf11JVT5smI5GZ+QkuWTfbpbzgCHac6NhOI652N/qUJDyUAEfOu
# Vi+2SoDl4t5Vl9zkB7dQe1YxZmxk0SGNpm7f+B8nkV2aonoKtNsBEMPFzrIIx11T
# YX22BiqO7rJXidcWz6PCNfDtmnMxBJ0yt0HwL/IqfsPlWTpFAKvsy12z22cO5FzG
# cV73to3U3A66QlwUG2lOj98wriSRlZhhMLCoA3QGmGq//oDEmsuamIOVLV/XQtwq
# kKgNQur/01GUubOPH7zcXF943JQgIwIDAQABo4IBxTCCAcEwHwYDVR0jBBgwFoAU
# WsS5eyoKo6XqcQPAYPkt9mV1DlgwHQYDVR0OBBYEFAMO2suSu//T5kHb495F8PQB
# JKCVMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzB3BgNVHR8E
# cDBuMDWgM6Axhi9odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVk
# LWNzLWcxLmNybDA1oDOgMYYvaHR0cDovL2NybDQuZGlnaWNlcnQuY29tL3NoYTIt
# YXNzdXJlZC1jcy1nMS5jcmwwTAYDVR0gBEUwQzA3BglghkgBhv1sAwEwKjAoBggr
# BgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAIBgZngQwBBAEw
# gYQGCCsGAQUFBwEBBHgwdjAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNl
# cnQuY29tME4GCCsGAQUFBzAChkJodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20v
# RGlnaUNlcnRTSEEyQXNzdXJlZElEQ29kZVNpZ25pbmdDQS5jcnQwDAYDVR0TAQH/
# BAIwADANBgkqhkiG9w0BAQsFAAOCAQEAJNk9CcZhUymD42L1jniEuMbgwYRzVgVZ
# 1kxuoENiLyE49tEziZ6W+k42/itW3UV2dgtNCKyy0RpjY1kw5mfbgcAMmgZ5M/d8
# kIHucV0ZGO0PqAlT+JIw3BbCDlvO9aFccCgU99V3XbCUv9IGsFcTgWcI27DO/3/r
# Pau6vGQbkL83cBDt7Gs1Fsz+pTZGg1md26LiN3dKfneyKDY+BtVNDqJulZ9KP6gz
# Z/QgeK8Vrt/TIvkCocmjzx+AHw3n9mwAifKEuF5zzeyTZE21xywV4seJmtWYtP74
# e5dZz0Uc+1on6zqipe50QBiiu5FWlYpOYCsTqHkX4pz4Igt5+qUDFzCCBmowggVS
# oAMCAQICEAMBmgI6/1ixa9bV6uYX8GYwDQYJKoZIhvcNAQEFBQAwYjELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBDQS0xMB4XDTE0
# MTAyMjAwMDAwMFoXDTI0MTAyMjAwMDAwMFowRzELMAkGA1UEBhMCVVMxETAPBgNV
# BAoTCERpZ2lDZXJ0MSUwIwYDVQQDExxEaWdpQ2VydCBUaW1lc3RhbXAgUmVzcG9u
# ZGVyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAo2Rd/Hyz4II14OD2
# xirmSXU7zG7gU6mfH2RZ5nxrf2uMnVX4kuOe1VpjWwJJUNmDzm9m7t3LhelfpfnU
# h3SIRDsZyeX1kZ/GFDmsJOqoSyyRicxeKPRktlC39RKzc5YKZ6O+YZ+u8/0SeHUO
# plsU/UUjjoZEVX0YhgWMVYd5SEb3yg6Np95OX+Koti1ZAmGIYXIYaLm4fO7m5zQv
# MXeBMB+7NgGN7yfj95rwTDFkjePr+hmHqH7P7IwMNlt6wXq4eMfJBi5GEMiN6ARg
# 27xzdPpO2P6qQPGyznBGg+naQKFZOtkVCVeZVjCT88lhzNAIzGvsYkKRrALA76Tw
# iRGPdwIDAQABo4IDNTCCAzEwDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB/wQCMAAw
# FgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwggG/BgNVHSAEggG2MIIBsjCCAaEGCWCG
# SAGG/WwHATCCAZIwKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNv
# bS9DUFMwggFkBggrBgEFBQcCAjCCAVYeggFSAEEAbgB5ACAAdQBzAGUAIABvAGYA
# IAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABlACAAYwBvAG4AcwB0AGkA
# dAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAgAG8AZgAgAHQAaABlACAA
# RABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMAIABhAG4AZAAgAHQAaABlACAA
# UgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAgAEEAZwByAGUAZQBtAGUAbgB0ACAA
# dwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABpAGEAYgBpAGwAaQB0AHkAIABhAG4A
# ZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwByAGEAdABlAGQAIABoAGUAcgBlAGkA
# bgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4wCwYJYIZIAYb9bAMVMB8GA1Ud
# IwQYMBaAFBUAEisTmLKZB+0e36K+Vw0rZwLNMB0GA1UdDgQWBBRhWk0ktkkynUoq
# eRqDS/QeicHKfTB9BgNVHR8EdjB0MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2Vy
# dC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDA4oDagNIYyaHR0cDovL2Ny
# bDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5jcmwwdwYIKwYB
# BQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20w
# QQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2Vy
# dEFzc3VyZWRJRENBLTEuY3J0MA0GCSqGSIb3DQEBBQUAA4IBAQCdJX4bM02yJoFc
# m4bOIyAPgIfliP//sdRqLDHtOhcZcRfNqRu8WhY5AJ3jbITkWkD73gYBjDf6m7Gd
# JH7+IKRXrVu3mrBgJuppVyFdNC8fcbCDlBkFazWQEKB7l8f2P+fiEUGmvWLZ8Cc9
# OB0obzpSCfDscGLTYkuw4HOmksDTjjHYL+NtFxMG7uQDthSr849Dp3GdId0UyhVd
# kkHa+Q+B0Zl0DSbEDn8btfWg8cZ3BigV6diT5VUW8LsKqxzbXEgnZsijiwoc5ZXa
# rsQuWaBh3drzbaJh6YoLbewSGL33VVRAA5Ira8JRwgpIr7DUbuD0FAo6G+OPPcqv
# ao173NhEMIIGzTCCBbWgAwIBAgIQBv35A5YDreoACus/J7u6GzANBgkqhkiG9w0B
# AQUFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVk
# IElEIFJvb3QgQ0EwHhcNMDYxMTEwMDAwMDAwWhcNMjExMTEwMDAwMDAwWjBiMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTEw
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDogi2Z+crCQpWlgHNAcNKe
# VlRcqcTSQQaPyTP8TUWRXIGf7Syc+BZZ3561JBXCmLm0d0ncicQK2q/LXmvtrbBx
# MevPOkAMRk2T7It6NggDqww0/hhJgv7HxzFIgHweog+SDlDJxofrNj/YMMP/pvf7
# os1vcyP+rFYFkPAyIRaJxnCI+QWXfaPHQ90C6Ds97bFBo+0/vtuVSMTuHrPyvAwr
# mdDGXRJCgeGDboJzPyZLFJCuWWYKxI2+0s4Grq2Eb0iEm09AufFM8q+Y+/bOQF1c
# 9qjxL6/siSLyaxhlscFzrdfx2M8eCnRcQrhofrfVdwonVnwPYqQ/MhRglf0HBKIJ
# AgMBAAGjggN6MIIDdjAOBgNVHQ8BAf8EBAMCAYYwOwYDVR0lBDQwMgYIKwYBBQUH
# AwEGCCsGAQUFBwMCBggrBgEFBQcDAwYIKwYBBQUHAwQGCCsGAQUFBwMIMIIB0gYD
# VR0gBIIByTCCAcUwggG0BgpghkgBhv1sAAEEMIIBpDA6BggrBgEFBQcCARYuaHR0
# cDovL3d3dy5kaWdpY2VydC5jb20vc3NsLWNwcy1yZXBvc2l0b3J5Lmh0bTCCAWQG
# CCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMA
# IABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMA
# IABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMA
# ZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkA
# bgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgA
# IABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUA
# IABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAA
# cgBlAGYAZQByAGUAbgBjAGUALjALBglghkgBhv1sAxUwEgYDVR0TAQH/BAgwBgEB
# /wIBADB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRp
# Z2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQu
# Y29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHoweDA6oDig
# NoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNybDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# QXNzdXJlZElEUm9vdENBLmNybDAdBgNVHQ4EFgQUFQASKxOYspkH7R7for5XDStn
# As0wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZIhvcNAQEF
# BQADggEBAEZQPsm3KCSnOB22WymvUs9S6TFHq1Zce9UNC0Gz7+x1H3Q48rJcYaKc
# lcNQ5IK5I9G6OoZyrTh4rHVdFxc0ckeFlFbR67s2hHfMJKXzBBlVqefj56tizfuL
# LZDCwNK1lL1eT7EF0g49GqkUW6aGMWKoqDPkmzmnxPXOHXh2lCVz5Cqrz5x2S+1f
# wksW5EtwTACJHvzFebxMElf+X+EevAJdqP77BzhPDcZdkbkPZ0XN1oPt55INjbFp
# jE/7WeAjD9KqrgB87pxCDs+R1ye3Fu4Pw718CqDuLAhVhSK46xgaTfwqIa1JMYNH
# lXdx3LEbS0scEJx3FMGdTy9alQgpECYxggQDMIID/wIBATCBhjByMQswCQYDVQQG
# EwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNl
# cnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBT
# aWduaW5nIENBAhAFNXZCWIS5pwXEey8vSOKFMAkGBSsOAwIaBQCgQDAZBgkqhkiG
# 9w0BCQMxDAYKKwYBBAGCNwIBBDAjBgkqhkiG9w0BCQQxFgQUwXEu3kNtZTNLAvHm
# +GYgwz4yt3cwDQYJKoZIhvcNAQEBBQAEggEAk6PtOp5pxzM2ZY8r5C/tfnTmwJ66
# CxNObbNaQBG2xgWwprczbCU35l8h7aPYK0r9fqxmr0ilPwzLcXtwdWcc3VTYPOPD
# hDADH8nlO9cdO8OQ/+NEZvbiquO3taBaPLrVoima6qyK/0qM/7ACUCJcZKO1vmEd
# OsXRDXLF3TjBG8vhYqBM+Dq0EQQi6h+A+HxRbHhN5yx/A7rKarRaiYsp38M/nTQk
# x89CjmmK2PbQIVmPU9Pi7oph4Avcrfekk+DOKeoJ65MYC4rArtlmAjW317oRzf0d
# nPjrtKEPc58hZ3WSlBqnM6uVKNBq8Nwep9behfXYOs/J5w6AChrSP3vdV6GCAg8w
# ggILBgkqhkiG9w0BCQYxggH8MIIB+AIBATB2MGIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAf
# BgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMQIQAwGaAjr/WLFr1tXq5hfw
# ZjAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMTYxMDI0MDEyMTMxWjAjBgkqhkiG9w0BCQQxFgQUC/hkRaGP36uM
# NHh+ipa6hkmyo+IwDQYJKoZIhvcNAQEBBQAEggEAlAPy1nN83WNwNQyUOSaLhHqS
# Pc3wDNhyZd2EZeAEKESZgaJFb9+KaUPX90eC7vn/irIY5CWZNh1aJ1wB2UAwrvP2
# ylB4WKibv1d/WTD1hO+oiAmQuuaBCjmE7ar27afPh0xzIddaUGp4Vl51aQa84K44
# 87s2trZT1NUyRFh/0ogqNDzCFMi3cOoh08OTqdFYwwivlriAyqpmqzp0hRJpGHyV
# NLwlemchPRC5qeY4aTp6OdAjJdkIu4qYfHbd8Hg6r8DJKFFWoBdBcssLfnuGZ1VZ
# duikBjhiCr6OfW0vuiGPxUcZdPPNl1Obdg7Vg0fYp4KgeFaAVMRRMSA4gVoNCQ==
# SIG # End signature block
