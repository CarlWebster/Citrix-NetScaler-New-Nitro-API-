#Requires -Version 3.0
#This File is in Unicode format.  Do not edit in an ASCII editor.

#region Support
<#
.COMMENT
    If you find issues with saving the final document or table layout is messed up please use the X86 version of Powershell!
.SYNOPSIS
    Creates a complete inventory of a Citrix NetScaler configuration using Microsoft Word.
.NetScaler Documentation Script
    NAME: NetScaler_Script_v3_0.ps1
	VERSION NetScaler Script: 3.0
	VERSION Script Template: 2016
	AUTHOR NetScaler script: Barry Schiffer
    AUTHOR NetScaler script functions: Iain Brighton
    AUTHOR Script template: Carl Webster, Michael B. Smith, Iain Brighton, Jeff Wouters
	LASTEDIT: May 24th 2016 
.Release Notes version 3
    Overall
        The script has had a major overhaul and is now completely utilizing the Nitro API instead of the NS.Conf.
        The Nitro API offers a lot more information and most important end result is much more predictable. Adding NetScaler functionality is also much easier.
        Added functionality because of Nitro
        * Hardware and license information
        * Complete routing tables including default routes
        * Complete monitoring information including default monitors
        * 
#>


<#
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
.PARAMETER NSIP
    NetScaler IP address, could be NSIP or SNIP with management enabled
.PARAMETER Credential
    NetScaler username/password
.PARAMETER UseSSL
	EXPERIMENTAL: Require SSL/TLS for communication with the NetScaler Nitro API. NOTE: This requires the client to trust to the NetScaler's certificate chain.
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
	(default cover pages in Word en-US)
	Valid input is:
		Alphabet (Word 2007/2010. Works)
		Annual (Word 2007/2010. Doesn't work well for this report)
		Austere (Word 2007/2010. Works)
		Austin (Word 2007/2010/2013. Doesn't work in 2013, mostly works in 2007/2010 but Subtitle/Subject & Author fields need to me moved after title box is moved up)
		Banded (Word 2013. Works)
		Conservative (Word 2007/2010. Works)
		Contrast (Word 2007/2010. Works)
		Cubicles (Word 2007/2010. Works)
		Exposure (Word 2007/2010. Works if you like looking sideways)
		Facet (Word 2013. Works)
		Filigree (Word 2013. Works)
		Grid (Word 2010/2013.Works in 2010)
		Integral (Word 2013. Works)
		Ion (Dark) (Word 2013. Top date doesn't fit, box needs to be manually resized or font changed to 8 point)
		Ion (Light) (Word 2013. Top date doesn't fit, box needs to be manually resized or font changed to 8 point)
		Mod (Word 2007/2010. Works)
		Motion (Word 2007/2010/2013. Works if top date is manually changed to 36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2007/2010. Works)
		Puzzle (Word 2007/2010. Top date doesn't fit, box needs to be manually resized or font changed to 14 point)
		Retrospect (Word 2013. Works)
		Semaphore (Word 2013. Works)
		Sideline (Word 2007/2010/2013. Doesn't work in 2013, works in 2007/2010)
		Slice (Dark) (Word 2013. Doesn't work)
		Slice (Light) (Word 2013. Doesn't work)
		Stacks (Word 2007/2010. Works)
		Tiles (Word 2007/2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2007/2010. Works)
		ViewMaster (Word 2013. Works)
		Whisp (Word 2013. Works)
	Default value is Sideline.
	This parameter has an alias of CP.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	For Word 2007, the Microsoft add-in for saving as a PDF muct be installed.
	For Word 2007, please see http://www.microsoft.com/en-us/download/details.aspx?id=9943
	The PDF file is roughly 5X to 10X larger than the DOCX file.
.PARAMETER Text
	Creates a formatted text file with a .txt extension.
	This parameter is disabled by default.
.PARAMETER MSWord
	SaveAs DOCX file
	This parameter is set True if no other output format is selected.
.PARAMETER HTML
	Creates an HTML file with an .html extension.
	This parameter is disabled by default.
.PARAMETER AddDateTime
	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2014 at 6PM is 2014-06-01_1800.
	Output filename will be ReportName_2014-06-01_1800.docx (or .pdf).
	This parameter is disabled by default.
.PARAMETER Hardware
	Use WMI to gather hardware information on: Computer System, Disks, Processor and Network Interface Cards
	This parameter may require the script be run from an elevated PowerShell session 
	using an account with permission to retrieve hardware information (i.e. Domain Admin or Local Administrator).
	Selecting this parameter will add to both the time it takes to run the script and size of the report.
	This parameter is disabled by default.
.PARAMETER ComputerName
	Specifies a computer to use to run the script against.
	ComputerName can be entered as the NetBIOS name, FQDN, localhost or IP Address.
	If entered as localhost, the actual computer name is determined and used.
	If entered as an IP address, an attempt is made to determine and use the actual computer name.
	Default is localhost.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v2_0.ps1
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v2_0.ps1 -PDF
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v2_0.ps1 -TEXT
	
	Will use all default values and save the document as a formatted text file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v2_0.ps1 -HTML
	
	Will use all default values and save the document as an HTML file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript .\NetScaler_Script_v2_0.ps1 -CompanyName "Carl Webster Consulting" -CoverPage "Mod" -UserName "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
.EXAMPLE
	PS C:\PSScript .\NetScaler_Script_v2_0.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v2_0.ps1 -AddDateTime
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2014 at 6PM is 2014-06-01_1800.
	Output filename will be Script_Template_2014-06-01_1800.docx
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v2_0.ps1 -PDF -AddDateTime
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2014 at 6PM is 2014-06-01_1800.
	Output filename will be Script_Template_2014-06-01_1800.PDF
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v2_0.ps1 -Hardware
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	localhost for running hardware inventory.
	localhost will be replaced by the actual computer name.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v2_0.ps1 -Hardware -ComputerName 192.168.1.51
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	192.168.1.51 for running hardware inventory.
	192.168.1.51 will be replaced by the actual computer name, if possible.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  
	This script creates a Word, PDF, Formatted Text or HTML document.
.NOTES
	NAME: Based on Script Template date 17072014
	VERSION: 17072014
	AUTHOR: Carl Webster, Michael B. Smith, Iain Brighton, Jeff Wouters, Barry Schiffer
	LASTEDIT: July 17, 2014
#>

<#
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
#>
<#
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
#endregion Support

#region script template
#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "WordOrPDF") ]

Param(
    [parameter(
    Position = 0,
    Mandatory=$true )
    ]
    [string] $NSIP,
    
    [parameter(
    Mandatory=$false )
    ]
    [PSCredential] $Credential = (Get-Credential -Message 'Enter NetScaler credentials'),
	
	## EXPERIMENTAL: Require SSL/TLS, e.g. https://. This requires the client to trust to the NetScaler's certificate.
    [parameter(
    Mandatory=$false )	
	]
	[System.Management.Automation.SwitchParameter] $UseSSL,
    
	[parameter(ParameterSetName="WordOrPDF",
	Position = 1, 
	Mandatory=$False )
	] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(ParameterSetName="WordOrPDF",
	Position = 2, 
	Mandatory=$False )
	] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(ParameterSetName="WordOrPDF",
	Position = 3, 
	Mandatory=$False )
	] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

	[parameter(ParameterSetName="WordOrPDF",
	Position = 4, 
	Mandatory=$False )
	] 
	[Switch]$PDF=$False,

	[parameter(ParameterSetName="Text",
	Position = 5, 
	Mandatory=$False )
	] 
	[Switch]$Text=$False,

	[parameter(ParameterSetName="WordOrPDF",
	Position = 5, 
	Mandatory=$False )
	] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="HTML",
	Position = 5, 
	Mandatory=$False )
	] 
	[Switch]$HTML=$False,

	[parameter(
	Position = 6, 
	Mandatory=$False )
	] 
	[Switch]$AddDateTime=$False,
	
	[parameter(
	Position = 7, 
	Mandatory=$False )
	] 
	[Switch]$Hardware=$False,

	[parameter(
	Position = 8, 
	Mandatory=$False )
	] 
	[string]$ComputerName="LocalHost"
	
	)

#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#Created on June 1, 2014

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
If($Null -eq $Text)
{
	$Text = $False
}
If($Null -eq $MSWord)
{
	$MSWord = $False
}
If($Null -eq $HTML)
{
	$HTML = $False
}
If($Null -eq $AddDateTime)
{
	$AddDateTime = $False
}
If($Null -eq $ComputerName)
{
	$ComputerName = "LocalHost"
}
If($Null -eq $Hardware)
{
	$Hardware = $False
}
If($Null -eq $ComputerName)
{
	$ComputerName = "LocalHost"
}

If(!(Test-Path Variable:PDF))
{
	$PDF = $False
}
If(!(Test-Path Variable:Text))
{
	$Text = $False
}
If(!(Test-Path Variable:MSWord))
{
	$MSWord = $False
}
If(!(Test-Path Variable:HTML))
{
	$HTML = $False
}
If(!(Test-Path Variable:AddDateTime))
{
	$AddDateTime = $False
}
If(!(Test-Path Variable:Hardware))
{
	$Hardware = $False
}
If(!(Test-Path Variable:ComputerName))
{
	$ComputerName = "LocalHost"
}

If($MSWord -eq $Null)
{
	If($Text -or $HTML -or $PDF)
	{
		$MSWord = $False
	}
	Else
	{
		$MSWord = $True
	}
}

If($MSWord -eq $False -and $PDF -eq $False -and $Text -eq $False -and $HTML -eq $False)
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
ElseIf($Text)
{
	Write-Verbose "$(Get-Date): Text is set"
}
ElseIf($HTML)
{
	Write-Verbose "$(Get-Date): HTML is set"
}
Else
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Verbose "$(Get-Date): Unable to determine output parameter"
	If($MSWord -eq $Null)
	{
		Write-Verbose "$(Get-Date): MSWord is Null"
	}
	ElseIf($PDF -eq $Null)
	{
		Write-Verbose "$(Get-Date): PDF is Null"
	}
	ElseIf($Text -eq $Null)
	{
		Write-Verbose "$(Get-Date): Text is Null"
	}
	ElseIf($HTML -eq $Null)
	{
		Write-Verbose "$(Get-Date): HTML is Null"
	}
	Else
	{
		Write-Verbose "$(Get-Date): MSWord is $($MSWord)"
		Write-Verbose "$(Get-Date): PDF is $($PDF)"
		Write-Verbose "$(Get-Date): Text is $($Text)"
		Write-Verbose "$(Get-Date): HTML is $($HTML)"
	}
	Write-Error "Unable to determine output parameter.  Script cannot continue"
	Exit
}

#updated by Webster 23-Apr-2016
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

Function GetComputerWMIInfo
{
	Param([string]$RemoteComputerName)
	
	# original work by Kees Baggerman, 
	# Senior Technical Consultant @ Inter Access
	# k.baggerman@myvirtualvision.com
	# @kbaggerman on Twitter
	# http://blog.myvirtualvision.com
	# modified 1-May-2014 to work in trusted AD Forests and using different domain admin credentials	

	#Get Computer info
	Write-Verbose "$(Get-Date): `t`tProcessing WMI Computer information"
	Write-Verbose "$(Get-Date): `t`t`tHardware information"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Computer Information"
		WriteWordLine 0 1 "General Computer"
	}
	ElseIf($Text)
	{
		Line 0 "Computer Information"
		Line 1 "General Computer"
	}
	ElseIf($HTML)
	{
	}
	
	[bool]$GotComputerItems = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_computersystem
	}
	
	Catch
	{
		$Results = $Null
	}
	
	If($? -and $Results -ne $Null)
	{
		$ComputerItems = $Results | Select Manufacturer, Model, Domain, @{N="TotalPhysicalRam"; E={[math]::round(($_.TotalPhysicalMemory / 1GB),0)}}
		$Results = $Null

		ForEach($Item in $ComputerItems)
		{
			OutputComputerItem $Item
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
			Line 2 ""
		}
		ElseIf($HTML)
		{
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results returned for Computer information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results returned for Computer information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results returned for Computer information"
		}
		ElseIf($HTML)
		{
		}
	}
	
	#Get Disk info
	Write-Verbose "$(Get-Date): `t`t`tDrive information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "Drive(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Drive(s)"
	}
	ElseIf($HTML)
	{
	}

	[bool]$GotDrives = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName Win32_LogicalDisk
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Results -ne $Null)
	{
		$drives = $Results | Select caption, @{N="drivesize"; E={[math]::round(($_.size / 1GB),0)}}, 
		filesystem, @{N="drivefreespace"; E={[math]::round(($_.freespace / 1GB),0)}}, 
		volumename, drivetype, volumedirty, volumeserialnumber
		$Results = $Null
		ForEach($drive in $drives)
		{
			If($drive.caption -ne "A:" -and $drive.caption -ne "B:")
			{
				OutputDriveItem $drive
			}
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results returned for Drive information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results returned for Drive information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results returned for Drive information"
		}
		ElseIf($HTML)
		{
		}
	}
	

	#Get CPU's and stepping
	Write-Verbose "$(Get-Date): `t`t`tProcessor information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "Processor(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Processor(s)"
	}
	ElseIf($HTML)
	{
	}

	[bool]$GotProcessors = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_Processor
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Results -ne $Null)
	{
		$Processors = $Results | Select availability, name, description, maxclockspeed, 
		l2cachesize, l3cachesize, numberofcores, numberoflogicalprocessors
		$Results = $Null
		ForEach($processor in $processors)
		{
			OutputProcessorItem $processor
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results returned for Processor information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results returned for Processor information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results returned for Processor information"
		}
		ElseIf($HTML)
		{
		}
	}

	#Get Nics
	Write-Verbose "$(Get-Date): `t`t`tNIC information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "Network Interface(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Network Interface(s)"
	}
	ElseIf($HTML)
	{
	}

	[bool]$GotNics = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_networkadapterconfiguration
	}
	
	Catch
	{
		$Results
	}

	If($? -and $Results -ne $Null)
	{
		$Nics = $Results | Where {$_.ipaddress -ne $Null}
		$Results = $Null

		If($Nics -eq $Null ) 
		{ 
			$GotNics = $False 
		} 
		Else 
		{ 
			$GotNics = !($Nics.__PROPERTY_COUNT -eq 0) 
		} 
	
		If($GotNics)
		{
			ForEach($nic in $nics)
			{
				Try
				{
					$ThisNic = Get-WmiObject -computername $RemoteComputerName win32_networkadapter | Where {$_.index -eq $nic.index}
				}
				
				Catch 
				{
					$ThisNic = $Null
				}
				
				If($? -and $ThisNic -ne $Null)
				{
					OutputNicItem $Nic $ThisNic
				}
				ElseIf(!$?)
				{
					Write-Warning "$(Get-Date): Error retrieving NIC information"
					Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "Error retrieving NIC information" "" $Null 0 $False $True
						WriteWordLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
						WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
						WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
						WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
					}
					ElseIf($Text)
					{
						Line 2 "Error retrieving NIC information"
						Line 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
						Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
						Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
						Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
					}
					ElseIf($HTML)
					{
					}
				}
				Else
				{
					Write-Verbose "$(Get-Date): No results returned for NIC information"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "No results returned for NIC information" "" $Null 0 $False $True
					}
					ElseIf($Text)
					{
						Line 2 "No results returned for NIC information"
					}
					ElseIf($HTML)
					{
					}
				}
			}
		}	
	}
	ElseIf(!$?)
	{
		Write-Warning "$(Get-Date): Error retrieving NIC configuration information"
		Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Error retrieving NIC configuration information" "" $Null 0 $False $True
			WriteWordLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Error retrieving NIC configuration information"
			Line 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results returned for NIC configuration information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results returned for NIC configuration information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results returned for NIC configuration information"
		}
		ElseIf($HTML)
		{
		}
	}
	
	If($MSWORD -or $PDF)
	{
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 ""
	}
	ElseIf($HTML)
	{
	}

	$Results = $Null
	$ComputerItems = $Null
	$Drives = $Null
	$Processors = $Null
	$Nics = $Null
}

Function OutputComputerItem
{
	Param([object]$Item)
	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ItemInformation = @()
		$ItemInformation += @{ Data = "Manufacturer"; Value = $Item.manufacturer; }
		$ItemInformation += @{ Data = "Model"; Value = $Item.model; }
		$ItemInformation += @{ Data = "Domain"; Value = $Item.domain; }
		$ItemInformation += @{ Data = "Total Ram"; Value = "$($Item.totalphysicalram) GB"; }
		$Table = AddWordTable -Hashtable $ItemInformation -Columns Data,Value -List -AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 125;
		$Table.Columns.Item(2).Width = 100;

		$Table.Rows.SetLeftIndent($Indent2TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$TableRange = $Null
		$Table = $Null
		WriteWordLine 0 2 ""
		
	}
	ElseIf($Text)
	{
		Line 2 "Manufacturer`t: " $Item.manufacturer
		Line 2 "Model`t`t: " $Item.model
		Line 2 "Domain`t`t: " $Item.domain
		Line 2 "Total Ram`t: $($Item.totalphysicalram) GB"
		Line 2 ""
	}
	ElseIf($HTML)
	{
	}
}

Function OutputDriveItem
{
	Param([object]$Drive)
	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $DriveInformation = @()
		$DriveInformation += @{ Data = "Caption"; Value = $Drive.caption; }
		$DriveInformation += @{ Data = "Size"; Value = "$($drive.drivesize) GB"; }
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$DriveInformation += @{ Data = "File System"; Value = $Drive.filesystem; }
		}
		$DriveInformation += @{ Data = "Free Space"; Value = "$($drive.drivefreespace) GB"; }
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$DriveInformation += @{ Data = "Volume Name"; Value = $Drive.volumename; }
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			If($drive.volumedirty)
			{
				$tmp = "Yes"
			}
			Else
			{
				$tmp = "No"
			}
			$DriveInformation += @{ Data = "Volume is Dirty"; Value = $tmp; }
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$DriveInformation += @{ Data = "Volume Serial Number"; Value = $Drive.volumeserialnumber; }
		}
		Switch ($drive.drivetype)
		{
			0	{$tmp = "Unknown"}
			1	{$tmp = "No Root Directory"}
			2	{$tmp = "Removable Disk"}
			3	{$tmp = "Local Disk"}
			4	{$tmp = "Network Drive"}
			5	{$tmp = "Compact Disc"}
			6	{$tmp = "RAM Disk"}
			Default {$tmp = "Unknown"}
		}
		$DriveInformation += @{ Data = "Drive Type"; Value = $tmp; }
		$Table = AddWordTable -Hashtable $DriveInformation -Columns Data,Value -List -AutoFit $wdAutoFitContent;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 125;
		$Table.Columns.Item(2).Width = 100;

		$Table.Rows.SetLeftIndent($Indent2TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$TableRange = $Null
		$Table = $Null
		WriteWordLine 0 2 ""
	}
	ElseIf($Text)
	{
		Line 2 "Caption`t`t: " $drive.caption
		Line 2 "Size`t`t: $($drive.drivesize) GB"
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			Line 2 "File System`t: " $drive.filesystem
		}
		Line 2 "Free Space`t: $($drive.drivefreespace) GB"
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			Line 2 "Volume Name`t: " $drive.volumename
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			Line 2 "Volume is Dirty`t: " -nonewline
			If($drive.volumedirty)
			{
				Line 0 "Yes"
			}
			Else
			{
				Line 0 "No"
			}
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			Line 2 "Volume Serial #`t: " $drive.volumeserialnumber
		}
		Line 2 "Drive Type`t: " -nonewline
		Switch ($drive.drivetype)
		{
			0	{Line 0 "Unknown"}
			1	{Line 0 "No Root Directory"}
			2	{Line 0 "Removable Disk"}
			3	{Line 0 "Local Disk"}
			4	{Line 0 "Network Drive"}
			5	{Line 0 "Compact Disc"}
			6	{Line 0 "RAM Disk"}
			Default {Line 0 "Unknown"}
		}
		Line 2 ""
	}
	ElseIf($HTML)
	{
	}
}

Function OutputProcessorItem
{
	Param([object]$Processor)
	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $ProcessorInformation = @()
		$ProcessorInformation += @{ Data = "Name"; Value = $Processor.name; }
		$ProcessorInformation += @{ Data = "Description"; Value = $Processor.description; }
		$ProcessorInformation += @{ Data = "Max Clock Speed"; Value = "$($processor.maxclockspeed) MHz"; }
		If($processor.l2cachesize -gt 0)
		{
			$ProcessorInformation += @{ Data = "L2 Cache Size"; Value = "$($processor.l2cachesize) KB"; }
		}
		If($processor.l3cachesize -gt 0)
		{
			$ProcessorInformation += @{ Data = "L3 Cache Size"; Value = "$($processor.l3cachesize) KB"; }
		}
		If($processor.numberofcores -gt 0)
		{
			$ProcessorInformation += @{ Data = "Number of Cores"; Value = $Processor.numberofcores; }
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$ProcessorInformation += @{ Data = "Number of Logical Processors"; Value = $Processor.numberoflogicalprocessors; }
		}
		Switch ($processor.availability)
		{
			1	{$tmp = "Other"}
			2	{$tmp = "Unknown"}
			3	{$tmp = "Running or Full Power"}
			4	{$tmp = "Warning"}
			5	{$tmp = "In Test"}
			6	{$tmp = "Not Applicable"}
			7	{$tmp = "Power Off"}
			8	{$tmp = "Off Line"}
			9	{$tmp = "Off Duty"}
			10	{$tmp = "Degraded"}
			11	{$tmp = "Not Installed"}
			12	{$tmp = "Install Error"}
			13	{$tmp = "Power Save - Unknown"}
			14	{$tmp = "Power Save - Low Power Mode"}
			15	{$tmp = "Power Save - Standby"}
			16	{$tmp = "Power Cycle"}
			17	{$tmp = "Power Save - Warning"}
			Default	{$tmp = "Unknown"}
		}
		$ProcessorInformation += @{ Data = "Availability"; Value = $tmp; }
		$Table = AddWordTable -Hashtable $ProcessorInformation -Columns Data,Value -List -AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent2TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$TableRange = $Null
		$Table = $Null
		WriteWordLine 0 2 ""
	}
	ElseIf($Text)
	{
		Line 2 "Name`t`t`t: " $processor.name
		Line 2 "Description`t`t: " $processor.description
		Line 2 "Max Clock Speed`t`t: $($processor.maxclockspeed) MHz"
		If($processor.l2cachesize -gt 0)
		{
			Line 2 "L2 Cache Size`t`t: $($processor.l2cachesize) KB"
		}
		If($processor.l3cachesize -gt 0)
		{
			Line 2 "L3 Cache Size`t`t: $($processor.l3cachesize) KB"
		}
		If($processor.numberofcores -gt 0)
		{
			Line 2 "# of Cores`t`t: " $processor.numberofcores
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			Line 2 "# of Logical Procs`t: " $processor.numberoflogicalprocessors
		}
		Line 2 "Availability`t`t: " -nonewline
		Switch ($processor.availability)
		{
			1	{Line 0 "Other"}
			2	{Line 0 "Unknown"}
			3	{Line 0 "Running or Full Power"}
			4	{Line 0 "Warning"}
			5	{Line 0 "In Test"}
			6	{Line 0 "Not Applicable"}
			7	{Line 0 "Power Off"}
			8	{Line 0 "Off Line"}
			9	{Line 0 "Off Duty"}
			10	{Line 0 "Degraded"}
			11	{Line 0 "Not Installed"}
			12	{Line 0 "Install Error"}
			13	{Line 0 "Power Save - Unknown"}
			14	{Line 0 "Power Save - Low Power Mode"}
			15	{Line 0 "Power Save - Standby"}
			16	{Line 0 "Power Cycle"}
			17	{Line 0 "Power Save - Warning"}
			Default	{Line 0 "Unknown"}
		}
		Line 2 ""
	}
	ElseIf($HTML)
	{
	}
}

Function OutputNicItem
{
	Param([object]$Nic, [object]$ThisNic)
	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $NicInformation = @()
		If($ThisNic.Name -eq $nic.description)
		{
			$NicInformation += @{ Data = "Name"; Value = $ThisNic.Name; }
		}
		Else
		{
			$NicInformation += @{ Data = "Name"; Value = $ThisNic.Name; }
			$NicInformation += @{ Data = "Description"; Value = $Nic.description; }
		}
		$NicInformation += @{ Data = "Connection ID"; Value = $ThisNic.NetConnectionID; }
		$NicInformation += @{ Data = "Manufacturer"; Value = $Nic.manufacturer; }
		Switch ($ThisNic.availability)
		{
			1	{$tmp = "Other"}
			2	{$tmp = "Unknown"}
			3	{$tmp = "Running or Full Power"}
			4	{$tmp = "Warning"}
			5	{$tmp = "In Test"}
			6	{$tmp = "Not Applicable"}
			7	{$tmp = "Power Off"}
			8	{$tmp = "Off Line"}
			9	{$tmp = "Off Duty"}
			10	{$tmp = "Degraded"}
			11	{$tmp = "Not Installed"}
			12	{$tmp = "Install Error"}
			13	{$tmp = "Power Save - Unknown"}
			14	{$tmp = "Power Save - Low Power Mode"}
			15	{$tmp = "Power Save - Standby"}
			16	{$tmp = "Power Cycle"}
			17	{$tmp = "Power Save - Warning"}
			Default	{$tmp = "Unknown"}
		}
		$NicInformation += @{ Data = "Availability"; Value = $tmp; }
		$NicInformation += @{ Data = "Physical Address"; Value = $Nic.macaddress; }
		$NicInformation += @{ Data = "IP Address"; Value = $Nic.ipaddress; }
		$NicInformation += @{ Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }
		$NicInformation += @{ Data = "Subnet Mask"; Value = $Nic.ipsubnet; }
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$NicInformation += @{ Data = "DHCP Enabled"; Value = $Nic.dhcpenabled; }
			$NicInformation += @{ Data = "DHCP Lease Obtained"; Value = $dhcpleaseobtaineddate; }
			$NicInformation += @{ Data = "DHCP Lease Expires"; Value = $dhcpleaseexpiresdate; }
			$NicInformation += @{ Data = "DHCP Server"; Value = $Nic.dhcpserver; }
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$NicInformation += @{ Data = "DNS Domain"; Value = $Nic.dnsdomain; }
		}
		If($nic.dnsdomainsuffixsearchorder -ne $Null -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			[int]$x = 1
			WriteWordLine 0 2 "DNS Search Suffixes`t:" -nonewline
			$nicdnsdomainsuffixsearchorder = $nic.dnsdomainsuffixsearchorder
			$tmp = @()
			ForEach($DNSDomain in $nicdnsdomainsuffixsearchorder)
			{
				$tmp += "$($DNSDomain)`r"
			}
			$NicInformation += @{ Data = "DNS Search Suffixes"; Value = $tmp; }
		}
		If($nic.dnsenabledforwinsresolution)
		{
			$tmp = "Yes"
		}
		Else
		{
			$tmp = "No"
		}
		$NicInformation += @{ Data = "DNS WINS Enabled"; Value = $tmp; }
		If($nic.dnsserversearchorder -ne $Null -and $nic.dnsserversearchorder.length -gt 0)
		{
			$nicdnsserversearchorder = $nic.dnsserversearchorder
			$tmp = @()
			ForEach($DNSServer in $nicdnsserversearchorder)
			{
				$tmp += "$($DNSServer)`r"
			}
			$NicInformation += @{ Data = "DNS Servers"; Value = $tmp; }
		}
		Switch ($nic.TcpipNetbiosOptions)
		{
			0	{$tmp = "Use NetBIOS setting from DHCP Server"}
			1	{$tmp = "Enable NetBIOS"}
			2	{$tmp = "Disable NetBIOS"}
			Default	{$tmp = "Unknown"}
		}
		$NicInformation += @{ Data = "NetBIOS Setting"; Value = $tmp; }
		If($nic.winsenablelmhostslookup)
		{
			$tmp = "Yes"
		}
		Else
		{
			$tmp = "No"
		}
		$NicInformation += @{ Data = "WINS: Enabled LMHosts"; Value = $tmp; }
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$NicInformation += @{ Data = "Host Lookup File"; Value = $Nic.winshostlookupfile; }
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$NicInformation += @{ Data = "Primary Server"; Value = $Nic.winsprimaryserver; }
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$NicInformation += @{ Data = "Secondary Server"; Value = $Nic.winssecondaryserver; }
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$NicInformation += @{ Data = "Scope ID"; Value = $Nic.winsscopeid; }
		}
		$Table = AddWordTable -Hashtable $NicInformation -Columns Data,Value -List -AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent2TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$TableRange = $Null
		$Table = $Null
	}
	ElseIf($Text)
	{
		If($ThisNic.Name -eq $nic.description)
		{
			Line 2 "Name`t`t`t: " $ThisNic.Name
		}
		Else
		{
			Line 2 "Name`t`t`t: " $ThisNic.Name
			Line 2 "Description`t`t: " $nic.description
		}
		Line 2 "Connection ID`t`t: " $ThisNic.NetConnectionID
		Line 2 "Manufacturer`t`t: " $ThisNic.manufacturer
		Line 2 "Availability`t`t: " -nonewline
		Switch ($ThisNic.availability)
		{
			1	{Line 0 "Other"}
			2	{Line 0 "Unknown"}
			3	{Line 0 "Running or Full Power"}
			4	{Line 0 "Warning"}
			5	{Line 0 "In Test"}
			6	{Line 0 "Not Applicable"}
			7	{Line 0 "Power Off"}
			8	{Line 0 "Off Line"}
			9	{Line 0 "Off Duty"}
			10	{Line 0 "Degraded"}
			11	{Line 0 "Not Installed"}
			12	{Line 0 "Install Error"}
			13	{Line 0 "Power Save - Unknown"}
			14	{Line 0 "Power Save - Low Power Mode"}
			15	{Line 0 "Power Save - Standby"}
			16	{Line 0 "Power Cycle"}
			17	{Line 0 "Power Save - Warning"}
			Default	{Line 0 "Unknown"}
		}
		Line 2 "Physical Address`t: " $nic.macaddress
		Line 2 "IP Address`t`t: " $nic.ipaddress
		Line 2 "Default Gateway`t`t: " $nic.Defaultipgateway
		Line 2 "Subnet Mask`t`t: " $nic.ipsubnet
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			Line 2 "DHCP Enabled`t`t: " $nic.dhcpenabled
			Line 2 "DHCP Lease Obtained`t: " $dhcpleaseobtaineddate
			Line 2 "DHCP Lease Expires`t: " $dhcpleaseexpiresdate
			Line 2 "DHCP Server`t`t:" $nic.dhcpserver
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			Line 2 "DNS Domain`t`t: " $nic.dnsdomain
		}
		If($nic.dnsdomainsuffixsearchorder -ne $Null -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			[int]$x = 1
			Line 2 "DNS Search Suffixes`t:" -nonewline
			$nicdnsdomainsuffixsearchorder = $nic.dnsdomainsuffixsearchorder
			ForEach($DNSDomain in $nicdnsdomainsuffixsearchorder)
			{
				If($x -eq 1)
				{
					$x = 2
					Line 0 " $($DNSDomain)"
				}
				Else
				{
					Line 5 " $($DNSDomain)"
				}
			}
		}
		Line 2 "DNS WINS Enabled`t: " -nonewline
		If($nic.dnsenabledforwinsresolution)
		{
			Line 0 "Yes"
		}
		Else
		{
			Line 0 "No"
		}
		If($nic.dnsserversearchorder -ne $Null -and $nic.dnsserversearchorder.length -gt 0)
		{
			[int]$x = 1
			Line 2 "DNS Servers`t`t:" -nonewline
			$nicdnsserversearchorder = $nic.dnsserversearchorder
			ForEach($DNSServer in $nicdnsserversearchorder)
			{
				If($x -eq 1)
				{
					$x = 2
					Line 0 " $($DNSServer)"
				}
				Else
				{
					Line 5 " $($DNSServer)"
				}
			}
		}
		Line 2 "NetBIOS Setting`t`t: " -nonewline
		Switch ($nic.TcpipNetbiosOptions)
		{
			0	{Line 0 "Use NetBIOS setting from DHCP Server"}
			1	{Line 0 "Enable NetBIOS"}
			2	{Line 0 "Disable NetBIOS"}
			Default	{Line 0 "Unknown"}
		}
		Line 2 "WINS:"
		Line 3 "Enabled LMHosts`t: " -nonewline
		If($nic.winsenablelmhostslookup)
		{
			Line 0 "Yes"
		}
		Else
		{
			Line 0 "No"
		}
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			Line 3 "Host Lookup File`t: " $nic.winshostlookupfile
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			Line 3 "Primary Server`t`t: " $nic.winsprimaryserver
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			Line 3 "Secondary Server`t: " $nic.winssecondaryserver
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			Line 3 "Scope ID`t`t: " $nic.winsscopeid
		}
	}
	ElseIf($HTML)
	{
	}
}

Function SetWordHashTable
{
	Param([string]$CultureCode)
	$hash = @{}
	    
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

	Switch ($CultureCode)
	{
		'ca-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents' = 'Taula automática 2'
				}
			}

		'da-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents' = 'Automatisk tabel 2'
				}
			}

		'de-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents' = 'Automatische Tabelle 2'
				}
			}

		'en-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents'  = 'Automatic Table 2'
				}
			}

		'es-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents' = 'Tabla automática 2'
				}
			}

		'fi-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents' = 'Automaattinen taulukko 2'
				}
			}

		'fr-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents' = 'Sommaire Automatique 2'
				}
			}

		'nb-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents' = 'Automatisk tabell 2'
				}
			}

		'nl-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents' = 'Automatische inhoudsopgave 2'
				}
			}

		'pt-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents' = 'Sumário Automático 2'
				}
			}

		'sv-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents' = 'Automatisk innehållsförteckning2'
				}
			}

		Default	{$hash.('en-') = @{
					'Word_TableOfContents'  = 'Automatic Table 2'
				}
			}
	}

	$Script:myHash = $hash.$CultureCode

	If($Script:myHash -eq $Null)
	{
		$Script:myHash = $hash.('en-')
	}

	$Script:myHash.Word_NoSpacing = $wdStyleNoSpacing
	$Script:myHash.Word_Heading1 = $wdStyleheading1
	$Script:myHash.Word_Heading2 = $wdStyleheading2
	$Script:myHash.Word_Heading3 = $wdStyleheading3
	$Script:myHash.Word_Heading4 = $wdStyleheading4
	$Script:myHash.Word_TableGrid = $wdTableGrid
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
		{$CatalanArray -contains $_} {$CultureCode = "ca-"}
		{$DanishArray -contains $_} {$CultureCode = "da-"}
		{$DutchArray -contains $_} {$CultureCode = "nl-"}
		{$EnglishArray -contains $_} {$CultureCode = "en-"}
		{$FinnishArray -contains $_} {$CultureCode = "fi-"}
		{$FrenchArray -contains $_} {$CultureCode = "fr-"}
		{$GermanArray -contains $_} {$CultureCode = "de-"}
		{$NorwegianArray -contains $_} {$CultureCode = "nb-"}
		{$PortugueseArray -contains $_} {$CultureCode = "pt-"}
		{$SpanishArray -contains $_} {$CultureCode = "es-"}
		{$SwedishArray -contains $_} {$CultureCode = "sv-"}
		Default {$CultureCode = "en-"}
	}
	
	Return $CultureCode
}

#updated by Webster 23-Apr-2016
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
		$xArray = $Null
		Return $True
	}
	Else
	{
		$xArray = $Null
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

#updated by Webster 23-Apr-2016 (this function can be deleted
#Function CheckWord2007SaveAsPDFInstalled
#{
#	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Installer\Products\000021090B0090400000000000F01FEC) -eq $False)
#	{
#		Write-Host "`n`n`t`tWord 2007 is detected and the option to SaveAs PDF was selected but the Word 2007 SaveAs PDF add-in is not installed."
#		Write-Host "`n`n`t`tThe add-in can be downloaded from http://www.microsoft.com/en-us/download/details.aspx?id=9943"
#		Write-Host "`n`n`t`tInstall the SaveAs PDF add-in and rerun the script."
#		Return $False
#	}
#	Return $True
#}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

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

Function line
#function created by Michael B. Smith, Exchange MVP
#@essentialexchange on Twitter
#http://TheEssentialExchange.com
#for creating the formatted text report
#created March 2011
#updated March 2014
{
	Param( [int]$tabs = 0, [string]$name = '', [string]$value = '', [string]$newline = "`r`n", [switch]$nonewline )
	While( $tabs -gt 0 ) { $Global:Output += "`t"; $tabs--; }
	If( $nonewline )
	{
		$Global:Output += $name + $value
	}
	Else
	{
		$Global:Output += $name + $value + $newline
	}
}
	
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
		0 {$Script:Selection.Style = $myHash.Word_NoSpacing}
		1 {$Script:Selection.Style = $myHash.Word_Heading1}
		2 {$Script:Selection.Style = $myHash.Word_Heading2}
		3 {$Script:Selection.Style = $myHash.Word_Heading3}
		4 {$Script:Selection.Style = $myHash.Word_Heading4}
		Default {$Script:Selection.Style = $myHash.Word_NoSpacing}
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

Function AbortScript
{
	$Script:Word.quit()
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	$ErrorActionPreference = $SaveEAPreference
	Exit
}

Function FindWordDocumentEnd
{
	#return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

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
		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string[]] $Columns = $null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string[]] $Headers = $null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$true)] [int] $Format = 0
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Columns -eq $null) -and ($Headers -ne $null)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $null;
		}
		ElseIf(($Columns -ne $null) -and ($Headers -ne $null)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end elseif
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
        [System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Columns -eq $null) 
				{
					## Build the available columns from all availble PSCustomObject note properties
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
					If($Headers -ne $null) 
					{
                        [ref] $null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        [ref] $null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
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
                    [ref] $null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end foreach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Columns -eq $null) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Headers -ne $null) 
					{ 
                        [ref] $null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
                        [ref] $null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
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
                    [ref] $null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
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
			$ConvertToTableArguments.Add("ApplyBorders", $true);
			$ConvertToTableArguments.Add("ApplyShading", $true);
			$ConvertToTableArguments.Add("ApplyFont", $true);
			$ConvertToTableArguments.Add("ApplyColor", $true);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $true); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $true);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $true);
			$ConvertToTableArguments.Add("ApplyLastColumn", $true);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$null,                                          # Modifiers
			$null,                                          # Culture
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

		#the next line causes the heading row to flow across page breaks
		$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
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
		[Parameter()] [AllowNull()] [string] $Font = $null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] $BackgroundColor = $null,
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
					If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
					If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end foreach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $true; }
				If($Italic) { $Cell.Range.Font.Italic = $true; }
				If($Underline) { $Cell.Range.Font.Underline = 1; }
				If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
				If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
				If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
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
					If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
					If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
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

Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Company Name : $($Script:CoName)"
	Write-Verbose "$(Get-Date): Cover Page   : $($CoverPage)"
	Write-Verbose "$(Get-Date): User Name    : $($UserName)"
	Write-Verbose "$(Get-Date): Save As PDF  : $($PDF)"
	Write-Verbose "$(Get-Date): Save As TEXT : $($TEXT)"
	Write-Verbose "$(Get-Date): Save As WORD : $($MSWORD)"
	Write-Verbose "$(Get-Date): Save As HTML : $($HTML)"
	Write-Verbose "$(Get-Date): Add DateTime : $($AddDateTime)"
	Write-Verbose "$(Get-Date): HW Inventory : $($Hardware)"
	Write-Verbose "$(Get-Date): Filename1    : $($Script:FileName1)"
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Filename2    : $($Script:FileName2)"
	}
	Write-Verbose "$(Get-Date): OS Detected  : $($RunningOS)"
	Write-Verbose "$(Get-Date): PSUICulture  : $($PSUICulture)"
	Write-Verbose "$(Get-Date): PSCulture    : $($PSCulture)"
	Write-Verbose "$(Get-Date): Word version : $($Script:WordProduct)"
	Write-Verbose "$(Get-Date): Word language: $($Script:WordLanguageValue)"
	Write-Verbose "$(Get-Date): PoSH version : $($Host.Version)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Script start : $($Script:StartTime)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
}

Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	if( $object )
	{
		If( ( gm -Name $topLevel -InputObject $object ) )
		{
			If( ( gm -Name $secondLevel -InputObject $object.$topLevel ) )
			{
				Return $True
			}
		}
	}
	Return $False
}

#updated by Webster 23-Apr-2016
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
		If ($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) 
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
	ForEach ($footer in $footers) 
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
				[string]$abstract = "$($AbstractTitle) for $Script:CoName"
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

#updated by Webster 23-Apr-2016
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

#updated by Webster 23-Apr-2016
Function SaveandCloseTextDocument
{
	If($AddDateTime)
	{
		$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	}

	Write-Output $Global:Output | Out-File $Script:Filename1 4>$Null
}

#updated by Webster 23-Apr-2016
Function SaveandCloseHTMLDocument
{
	If($AddDateTime)
	{
		$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).html"
	}
	
	Out-File -FilePath $Script:FileName1 -Append -InputObject "<p></p></body></html>" 4>$Null
}

Function SetFileName1andFileName2
{
	Param([string]$OutputFileName)
	$pwdpath = $pwd.Path

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
	ElseIf($Text)
	{
		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).txt"
		}
	}
	ElseIf($HTML)
	{
		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).html"
		}
	}
}

Function TestComputerName
{
	Param([string]$Cname)
	If(![String]::IsNullOrEmpty($CName)) 
	{
		#get computer name
		#first test to make sure the computer is reachable
		Write-Verbose "$(Get-Date): Testing to see if $($CName) is online and reachable"
		If(Test-Connection -ComputerName $CName -quiet)
		{
			Write-Verbose "$(Get-Date): Server $($CName) is online."
		}
		Else
		{
			Write-Verbose "$(Get-Date): Computer $($CName) is offline"
			$ErrorActionPreference = $SaveEAPreference
			Write-Error "`n`n`t`tComputer $($CName) is offline.`nScript cannot continue.`n`n"
			Exit
		}
	}

	#if computer name is localhost, get actual computer name
	If($CName -eq "localhost")
	{
		$CName = $env:ComputerName
		Write-Verbose "$(Get-Date): Computer name has been renamed from localhost to $($CName)"
		Return $CName
	}

	#if computer name is an IP address, get host name from DNS
	#http://blogs.technet.com/b/gary/archive/2009/08/29/resolve-ip-addresses-to-hostname-using-powershell.aspx
	#help from Michael B. Smith
	$ip = $CName -as [System.Net.IpAddress]
	If($ip)
	{
		$Result = [System.Net.Dns]::gethostentry($ip)
		
		If($? -and $Result -ne $Null)
		{
			$CName = $Result.HostName
			Write-Verbose "$(Get-Date): Computer name has been renamed from $($ip) to $($CName)"
			Return $CName
		}
		Else
		{
			Write-Warning "Unable to resolve $($CName) to a hostname"
		}
	}
	Else
	{
		#computer is online but for some reason $ComputerName cannot be converted to a System.Net.IpAddress
	}
	Return $CName
}

#Script begins

$script:startTime = Get-Date

If($TEXT)
{
	$global:output = ""
}

$ComputerName = TestComputerName $ComputerName
<#
###The function SetFileName1andFileName2 needs your script output filename
SetFileName1andFileName2 "Script_Template"

###change title for your report
[string]$Script:Title = "This is the Report Title"
#>
#endregion script template

#region file name and title name
#The function SetFileName1andFileName2 needs your script output filename
SetFileName1andFileName2 "NetScaler Documentation"

#change title for your report
[string]$Script:Title = "NetScaler Documentation $CoName"
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
        [Parameter(ParameterSetName='HTTPS')] [System.Management.Automation.SwitchParameter] $UseSSL
    )
    process {
        if ($UseSSL) { $protocol = 'https'; }
        else { $protocol = 'http'; }
        $script:nsSession = @{ Address = $ComputerName; UseSSL = $UseSSL }
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
    [ref] $null = Connect-vNetScalerSession -ComputerName $nsip -Credential $Credential -UseSSL:$UseSSL -ErrorAction Stop;
}
#endregion NetScaler Connect

#region NetScaler chaptercounters
$Chapters = 32
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
If ($NSFEATURES.feo -eq "True") {$FEATfeo = "Enabled"} Else {$FEATVfeo = "Disabled"}
#endregion NetScaler feature state

#region NetScaler Version

## Get version and build
$NSVersion = Get-vNetScalerObject -Container config -Object nsversion;
$NSVersion1 = ($NSVersion.version -replace 'NetScaler', '').split()
$Version = ($NSVersion1[1] -replace ':', '')
$Build = $($NSVersion1[5] + " " + $nsversion1[6] + " " + $nsversion1[7] -replace ',', '')

## Set script test version
## WIP THIS WORKS ONLY WHEN REGIONAL SETTINGS DIGIT IS SET TO . :)
$ScriptVersion = 11.0
#endregion NetScaler Version

#region NetScaler System Information

#region Basics
WriteWordLine 1 0 "NetScaler Configuration"

$nsconfig = Get-vNetScalerObject -Container config -Object nsconfig;
$nshostname = Get-vNetScalerObject -Container config -Object nshostname;

WriteWordLine 2 0 "NetScaler Version and configuration"

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
$TableRange = $Null
$Table = $Null      
 
WriteWordLine 0 0 " "
#endregion Basics

#region NetScaler IP
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler IP"

WriteWordLine 2 0 "NetScaler Management IP Address"

$NSIP = Get-vNetScalerObject -Container config -Object nsip;
Foreach ($IP in $NSIP){ ##Lists all NetScaler IPs while we only need NSIP for this one
    If ($IP.Type -eq "NSIP")
        {
        $Params = @{
            Hashtable = @{
                NSIP = $nsconfig.ipaddress;
                Subnet = $nsconfig.netmask;
            }
            Columns = "NSIP","Subnet";
            Headers = "NetScaler IP Address","Subnet";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
    }
 }
#endregion NetScaler IP

#region NetScaler High Availability

WriteWordLine 2 0 "NetScaler High Availability"

$HANodes = Get-vNetScalerObject -Container config -Object hanode;

## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $HAH = @();

foreach ($HANODE in $HANodes) {        
    $HAH += @{
        HANAME = $HANODE.name;
        HAIP = $HANODE.ipaddress;
        HASTATUS = $HANODE.state;
        HASYNC = $HANODE.hasync;        
        }
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
    }

#endregion NetScaler High Availability

#region NetScaler Global HTTP Parameters
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Global HTTP Parameters"

WriteWordLine 2 0 "NetScaler Global HTTP Parameters"
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

#endregion NetScaler Global HTTP Parameters

#region NetScaler Global TCP Parameters
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Global TCP Parameters"

WriteWordLine 2 0 "NetScaler Global TCP Parameters"

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
    
#endregion NetScaler Global TCP Parameters

#region NetScaler Global Diameter Parameters

$nsdiameter = Get-vNetScalerObject -Container config -Object nsdiameter; 

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Global Diameter Parameter"

WriteWordLine 2 0 "NetScaler Global Diameter Parameters"

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

#endregion NetScaler Global Diameter Parameters

#region NetScaler Time Zone
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Time zone"
WriteWordLine 2 0 "NetScaler Time Zone"

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

#region NetScaler Administration
$selection.InsertNewPage()
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Administration"
WriteWordLine 2 0 "NetScaler System Authentication"
WriteWordLine 0 0 " "

#region Local Administration Users
WriteWordLine 3 0 "NetScaler System Users"

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
    }
WriteWordLine 0 0 " "
#endregion Authentication Local Administration Users

#region Authentication Local Administration Groups
WriteWordLine 3 0 "NetScaler System Groups"

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
    }
else { WriteWordLine 0 0 "No Local Group has been configured"}
WriteWordLine 0 0 " "

#endregion Authentication Local Administration Groups

#endregion NetScaler Administration

#region NetScaler Features
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Features"

$selection.InsertNewPage()

WriteWordLine 1 0 "NetScaler Features"

If ($Version -gt $ScriptVersion) {
    WriteWordLine 0 0 ""
    WriteWordLine 0 0 "Warning: You are using Citrix NetScaler version $Version, features added since version $ScriptVersion will not be shown."
    WriteWordLine 0 0 ""
    }
#region NetScaler Basic Features
WriteWordLine 2 0 "NetScaler Basic Features"

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

WriteWordLine 2 0 "SNMP Community"

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
        }
    }
WriteWordLine 0 0 " "

WriteWordLine 2 0 "SNMP Manager"
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
        }
    }
WriteWordLine 0 0 ""

WriteWordLine 2 0 "SNMP Alert"

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
    }

WriteWordLine 0 0 ""


WriteWordLine 2 0 "SNMP Traps"

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
        }
    }
WriteWordLine 0 0 ""

$selection.InsertNewPage()

#endregion NetScaler Monitoring

#endregion NetScaler System Information

#region NetScaler Networking

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Networking"

WriteWordLine 1 0 "NetScaler Networking"

#region NetScaler Interfaces

WriteWordLine 2 0 "NetScaler Interfaces"

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
            }
        }

#endregion NetScaler Interfaces

#region NetScaler Channels

WriteWordLine 2 0 "NetScaler Channels"

$ChannelCounter = Get-vNetScalerObjectCount -Container config -Object channel; 
$ChannelCount = $ChannelCounter.__count
$Channels = Get-vNetScalerObject -Container config -Object interface;

if($ChannelCounter.__count -le 0) { WriteWordLine 0 0 "No Channel has been configured"} else {
    
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
            }
        }

#endregion NetScaler Channels

#region NetScaler IP addresses

WriteWordLine 2 0 "NetScaler IP addresses"
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
#endregion NetScaler IP addresses

#region NetScaler vLAN

WriteWordLine 2 0 "NetScaler vLANs"

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
        }

#endregion routing table

#region NetScaler Traffic Domains
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Traffic Domains"

WriteWordLine 2 0 "NetScaler Traffic Domains"

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
    }
}
    
#endregion NetScaler Traffic Domains

#region NetScaler DNS Configuration
$selection.InsertNewPage()
WriteWordLine 1 0 "NetScaler DNS Configuration"

#region dns name servers

WriteWordLine 2 0 "NetScaler DNS Name Servers"

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
            }
        }
      
#endregion dns name servers

#region DNS Address Records

WriteWordLine 2 0 "NetScaler DNS Address Records"

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
            }
        }

#endregion DNS Address Records

#endregion NetScaler DNS Configuration

#region NetScaler ACL
$selection.InsertNewPage()
WriteWordLine 1 0 "NetScaler ACL Configuration"

#region NetScaler Simple ACL

WriteWordLine 2 0 "NetScaler Simple ACL"

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

WriteWordLine 2 0 "NetScaler Extended ACL"

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

#region Authentication LDAP Policies
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler LDAP Authentication"
WriteWordLine 2 0 "NetScaler LDAP Policies"

$authpolsldap = Get-vNetScalerObject -Container config -Object authenticationldappolicy;

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

#endregion Authentication LDAP Policies

#region Authentication LDAP
WriteWordLine 2 0 "NetScaler LDAP authentication actions"

$authactsldap = Get-vNetScalerObject -Container config -Object authenticationldapaction;

foreach ($authactldap in $authactsldap) {
    $ACTNAMELDAP = $authactldap.name
    WriteWordLine 3 0 "LDAP Authentication action $ACTNAMELDAP";

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

$authpolsradius = Get-vNetScalerObject -Container config -Object authenticationradiuspolicy;

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

#endregion Authentication Radius Policies

#region Authentication RADIUS
WriteWordLine 2 0 "NetScaler Radius authentication actions"

$authactsradius = Get-vNetScalerObject -Container config -Object authenticationradiusaction;

foreach ($authactradius in $authactsradius) {
    $ACTNAMERADIUS = $authactradius.name
    WriteWordLine 3 0 "Radius Authentication action $ACTNAMERADIUS";

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

#region SSL Certificates
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Certificates"

$selection.InsertNewPage()

WriteWordLine 1 0 "NetScaler Certificates"

$sslcerts = Get-vNetScalerObject -Object sslcertkey;

## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $SSLCERTSH = @();

foreach ($sslcert in $sslcerts) {

    $sslcert1 = Get-vNetScalerObject -ResourceType sslcertkey -Name $sslcert.certkey;
    $subject = $sslcert1.subject
    $subject1 = $subject.Split(',')[-1]
    $sslfqdn = ($subject1 -replace 'CN=', '')
                
    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be about 400 characters wide!
    $SSLCERTSH += @{
            FQDN = $sslfqdn;
            KEY = $sslcert.key;
            EXPIRE = $sslcert.daystoexpiration;
            LINK = $sslcert.linkcertkeyname;
    }
}
        
if ($SSLCERTSH.Length -gt 0) {
    $Params = $null
    $Params = @{
        Hashtable = $SSLCERTSH;
        Columns = "FQDN","KEY","EXPIRE","LINK";
        Headers = "FQDN","SSL Key file","Expiration days","SSL Chain";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
        }
    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;
    }

#endregion SSL Certificates

#region NetScaler Content Switches
$Chapter++
$selection.InsertNewPage()
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Content Switches"

WriteWordLine 1 0 "NetScaler Content Switches"

$csvservers = Get-vNetScalerObject -Object csvserver;

foreach ($ContentSwitch in $csvservers) {
    $csvservername = $ContentSwitch.name
    WriteWordLine 2 0 "Content Switch $csvservername";

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

    $csvserverbindings = Get-vNetScalerObject -ResourceType csvserver_cspolicy_binding -Name $ContentSwitch.Name;

    WriteWordLine 3 0 "Policies"
            
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
        } else {
            WriteWordLine 0 0 "No policy has been configured for this Content Switch"
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

            }

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
            }

    WriteWordLine 0 0 " "
    WriteWordLine 3 0 "Policies"

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
            }

    FindWordDocumentEnd;
    WriteWordLine 0 0 " "
    WriteWordLine 3 0 "Redirect URL"
    
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

    FindWordDocumentEnd;
    } else {WriteWordLine 0 0 "No Redirection URL Configured"}
   
    ##Advanced Configuration   
    WriteWordLine 0 0 " "
    WriteWordLine 3 0 "Advanced Configuration"

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
        }
    }
$selection.InsertNewPage()

#endregion NetScaler Cache Redirection

#region NetScaler Services
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Services"

FindWordDocumentEnd;

WriteWordLine 1 0 "NetScaler Services"

$servicescounter = Get-vNetScalerObjectCount -Container config -Object service; 
$servicescount = $servicescounter.__count
$services = Get-vNetScalerObject -Container config -Object service;

if($servicescounter.__count -le 0) { WriteWordLine 0 0 "No Service has been configured"} else {
    $CurrentRowIndex = 0;

    foreach ($Service in $Services) {

        $CurrentRowIndex++;
        $servicename = $Service.name
        Write-Verbose "$(Get-Date): `tService $CurrentRowIndex/$($servicescount) $servicename"     
        WriteWordLine 2 0 "Service $servicename"

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
        } else {
            WriteWordLine 0 0 "No Monitor has been configured for this Service"
    } # end if

        WriteWordLine 0 0 " "
        WriteWordLine 3 0 "Advanced Configuration"

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

$servicegroupscounter = Get-vNetScalerObjectCount -Container config -Object servicegroup; 
$servicegroupscount = $servicegroupscounter.__count
$servicegroups = Get-vNetScalerObject -Container config -Object servicegroup;

if($servicegroupscounter.__count -le 0) { WriteWordLine 0 0 "No Service Group has been configured"} else {
    $CurrentRowIndex = 0;

    foreach ($Servicegroup in $servicegroups) {
        $CurrentRowIndex++;
        $servicename = $Servicegroup.servicegroupname
        Write-Verbose "$(Get-Date): `tService $CurrentRowIndex/$($servicegroupscount) $servicename"     
        WriteWordLine 2 0 "Service Group $servicename"

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
            WriteWordLine 0 0 "No Server has been configured for this Service Group"
        }   

        WriteWordLine 0 0 " "

        WriteWordLine 3 0 "Monitor"

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
            }
        }

$selection.InsertNewPage()    
#endregion NetScaler Servers

#endregion traffic management

#region Citrix NetScaler Gateway

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters Citrix NetScaler (Access) Gateway"
WriteWordLine 1 0 "Citrix NetScaler (Access) Gateway"

#region Citrix NetScaler Gateway CAG Global

WriteWordLine 2 0 "NetScaler Gateway Global Settings"
Write-Verbose "$(Get-Date): `tNetScaler Gateway Global Settings"

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

$cagglobalclient = Get-vNetScalerObject -Object vpnparameter;

## IB - Create an array of hashtables to store our columns.
## IB - about column names as we'll utilise a -List(view)!
[System.Collections.Hashtable[]] $NsGlobalClientExperience = @(
    ## IB - Each hashtable is a separate row in the table!
    @{ Column1 = "Description"; Column2 = "Value"; }
    @{ Column1 = "Session Time-Out"; Column2 = $cagglobalclient.sesstimeout; }
    @{ Column1 = "Client-Idle Time-Out"; Column2 = $cagglobalclient.clientidletimeoutwarning; }
    @{ Column1 = "Clientless URL Encoding"; Column2 = $cagglobalclient.clientlessmodeurlencoding; }
    @{ Column1 = "Clientless Persistent Cookie"; Column2 = $cagglobalclient.clientlesspersistentcookie; }
    @{ Column1 = "Single Sign-On to Web Applications"; Column2 = $cagglobalclient.sso; }
    @{ Column1 = "Credential Index"; Column2 = $cagglobalclient.ssocredential; }
    @{ Column1 = "Single Sign-On with Windows"; Column2 = $cagglobalclient.windowsautologon; }
    @{ Column1 = "Client Cleanup Prompt"; Column2 = $cagglobalclient.clientcleanupprompt; }
    @{ Column1 = "UI Theme"; Column2 = $cagglobalclient.uitheme; }
    @{ Column1 = "Split Tunnel"; Column2 = $cagglobalclient.splittunnel; }
    @{ Column1 = "Local LAN Access"; Column2 = $cagglobalclient.locallanaccess; }
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
#endregion GlobalSecurity

#region GlobalPublishedApps
WriteWordLine 3 0 "Global Settings Published Applications"

## IB - Create an array of hashtables to store our columns. Note: If we need the
## IB - headers to include spaces we can override these at table creation time.
## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = @{
        ICAPROXY = $cagglobalclient.icaproxy;
        WIMODE = $cagglobalclient.wihomeaddresstype;
        SSO = $cagglobalclient.sso;
    }
    Columns = "ICAPROXY","WIMODE","SSO";
    Headers = "ICA Proxy","Web Interface Portal Mode","Single Sign-On Domain";
    AutoFit = $wdAutoFitContent;
    Format = -235; ## IB - Word constant for Light List Accent 5
}

## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params -NoGridLines;
FindWordDocumentEnd;


WriteWordLine 0 0 " "

#endregion GlobalPublishedApps

#region Global STA
WriteWordLine 3 0 "Global Settings Secure Ticket Authority Configuration"

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

#endregion Global STA

#region Global AppController
WriteWordLine 3 0 "Global Settings App Controller Configuration"

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
#endregion Global AppController

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

        ## IB - Create an array of hashtables to store our columns. Note: If we need the
        $Params = $null
        $Params = @{
            Hashtable = @{
                State = $vpnvserver.state;
                Mode = $vpnvserver.icaonly;
                IPAddress = $vpnvserver.ipv46;
                Port = $vpnvserver.port;
                Protocol = $vpnvserver.servicetype;
                MaximumUsers = $vpnvserver.maxaaausers;
                MaxLogin = $vpnvserver.sothreshold;
            }
            Columns = "State","Mode","IPAddress","Port","Protocol","MaximumUsers","MaxLogin";
            Headers = "State","ICA Only","IP Address","Port","Protocol","Maximum Users","Maximum Logons";
            AutoFit = $wdAutoFitContent
            Format = -235; ## IB - Word constant for Light List Accent 5
        }

        ## IB - Add the table to the document, splatting the parameters
        $Table = AddWordTable @Params -NoGridLines;
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
#endregion CAG vServer basic configuration
  
    #region CAG Authentication LDAP Policies             
        WriteWordLine 3 0 "Authentication LDAP Policies"
        $errorcode = 1 #Set Error code to 1
        $vpnvserverldappols = Get-vNetScalerObject -ResourceType vpnvserver_authenticationldappolicy_binding -name $vpnvserver.Name
        $errorcode = $vpnvserverldappols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No LDAP Policy has been configured"} else { #Uses the mentioned error code to determine existency of policy

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
                
        } else { WriteWordLine 0 0 "No LDAP Policy has been configured"} #endif AUTHPOLHASH.Length
    } #end if no LDAP configures
WriteWordLine 0 0 " "
#endregion CAG Authentication LDAP Policies  

    #region CAG Authentication Radius Policies             
        WriteWordLine 3 0 "Authentication RADIUS Policies"
        $errorcode = 1 
        $vpnvserverradiuspols = Get-vNetScalerObject -ResourceType vpnvserver_authenticationradiuspolicy_binding -name $vpnvserver.Name
        $errorcode = $vpnvserverradiuspols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No RADIUS Policy has been configured"} else {

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
                
        } else { WriteWordLine 0 0 "No RADIUS Policy has been configured"} #endif AUTHPOLHASH.Length
    } #end if no LDAP configures

WriteWordLine 0 0 " "
#endregion CAG Authentication Radius Policies  
        
    #region CAG Authentication SAML IDP Policies             
        WriteWordLine 3 0 "Authentication SAML IDP Policies"
        $errorcode = 1 #Set Error code to 1
        $vpnvserversamlidppols = Get-vNetScalerObject -ResourceType vpnvserver_authenticationsamlidppolicy_binding -name $vpnvserver.Name
        $errorcode = $vpnvserversamlidppols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No SAML IDP Policy has been configured"} else { #Uses the mentioned error code to determine existency of policy

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
                
        } else { WriteWordLine 0 0 "No SAML IDP Policy has been configured"}
    } 
WriteWordLine 0 0 " "

#endregion CAG Authentication SAML IDP Policies  
    
    #region CAG Session Policies        
       
        WriteWordLine 3 0 "Session Policies"
        $errorcode = 1 #Set Error code to 1
        $vpnvserversespols = Get-vNetScalerObject -ResourceType vpnvserver_vpnsessionpolicy_binding -name $vpnvserver.Name
        $errorcode = $vpnvserversespols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No Session Policy has been configured"} else { #Uses the mentioned error code to determine existency of policy
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
        } else { WriteWordLine 0 0 "No Session Policy has been configured"} #endif SESSIONPOLHASH.Length

WriteWordLine 0 0 " "
    #endregion CAG Session Policies 
    
    #region CAG STA Policies        
       
        WriteWordLine 3 0 "Secure Ticket Authority"
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
        } else { WriteWordLine 0 0 "No STA Policy has been configured"} #

WriteWordLine 0 0 " "
    #endregion CAG STA Policies 

    #region CAG Cache Policies        
       
        WriteWordLine 3 0 "Cache Policies"
        $errorcode = 1 #Set Error code to 1
        $vpnvservercachepols = Get-vNetScalerObject -ResourceType vpnvserver_cachepolicy_binding -name $vpnvserver.Name
        $errorcode = $vpnvservercachepols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No Session Policy has been configured"} else { #Uses the mentioned error code to determine existency of policy
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
        } else { WriteWordLine 0 0 "No Cache Policy has been configured"} #

WriteWordLine 0 0 " "
    #endregion CAG Cache Policies 

    #region CAG Responder Policies        
       
        WriteWordLine 3 0 "Responder Policies"
        $errorcode = 1 #Set Error code to 1
        $vpnvserverrespols = Get-vNetScalerObject -ResourceType vpnvserver_responderpolicy_binding -name $vpnvserver.Name
        $errorcode = $vpnvserverrespols.errorcode #Set Errorcode to the actual error, if no error exists it will clear the value

        If ($errorcode -ne 1) {WriteWordLine 0 0 "No Responder Policy has been configured"} else { #Uses the mentioned error code to determine existency of policy
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
        } else { WriteWordLine 0 0 "No Responder Policy has been configured"}

WriteWordLine 0 0 " "
    #endregion CAG Responder Policies 

    $selection.InsertNewPage()
    }
}

#endregion CAG vServers

#region CAG Session Policies
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters Citrix NetScaler (Access) Gateway Policies"
WriteWordLine 1 0 "NetScaler Gateway Policies"

WriteWordLine 0 0 " "
WriteWordLine 2 0 "NetScaler Gateway Session Policies"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `tNetScaler Gateway Session Policies"

$vpnsessionpolicies = Get-vNetScalerObject -Container config -Object vpnsessionpolicy;

foreach ($vpnsessionpolicy in $vpnsessionpolicies) {
    $sesspolname = $vpnsessionpolicy.name
    WriteWordLine 3 0 "NetScaler Gateway Session Policy: $sesspolname";

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
#endregion CAG Policies

#region CAG Session Actions
WriteWordLine 0 0 " "
WriteWordLine 2 0 "NetScaler Gateway Session Actions"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `tNetScaler Gateway Session Actions"

$vpnsessionactions = Get-vNetScalerObject -Container config -Object vpnsessionaction;

foreach ($vpnsessionaction in $vpnsessionactions) {
    $sessactname = $vpnsessionaction.name
    WriteWordLine 3 0 "NetScaler Gateway Session Action: $sessactname";

#region Security
    
    WriteWordLine 4 0 "Security"

    ## IB - Create an array of hashtables to store our columns. Note: If we need the
    ## IB - headers to include spaces we can override these at table creation time.
    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = @{
            ## IB - Each hashtable is a separate row in the table!
            DEFAUTH = $vpnsessionaction.defaultauthorizationaction;
            SECBRW = $vpnsessionaction.securebrowse;
        }
        Columns = "DEFAUTH","SECBRW";
        Headers = "Default Authorization Action","Secure Browse";
        AutoFit = $wdAutoFitContent;
        Format = -235; ## IB - Word constant for Light List Accent 5
    }

    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;
    WriteWordLine 0 0 " "
#endregion Security

#region Published Applications  

    WriteWordLine 4 0 "Published Applications"

    ## IB - Create an array of hashtables to store our columns. Note: If we need the
    ## IB - headers to include spaces we can override these at table creation time.
    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = @{
            ## IB - Each hashtable is a separate row in the table!
            ICAPROXY = $vpnsessionaction.icaproxy;
            WIMODE = $vpnsessionaction.wiportalmode;
            SSO = $vpnsessionaction.sso;
        }
        Columns = "ICAPROXY","WIMODE","SSO";
        Headers = "ICA Proxy","Web Interface Portal Mode","Single Sign-On Domain";
        AutoFit = $wdAutoFitContent;
        Format = -235; ## IB - Word constant for Light List Accent 5
    }

    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;
    WriteWordLine 0 0 " "

#end region Published Applications
    $selection.InsertNewPage()
}

    #endregion CAG Session Policies

#region Client Experience
<#
    $vpnsessionaction.homepage 
    $vpnsessionaction.splitdns
    $vpnsessionaction.splittunnel
    $vpnsessionaction.clientchoices
    $vpnsessionaction.clientlessvpnmode
    $vpnsessionaction.clientlessmodeurlencoding
#>
#endregion Client Experience

#endregion CAG Actions

#endregion Citrix NetScaler Gateway

#region NetScaler Monitors
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Monitors"

WriteWordLine 1 0 "NetScaler Monitors"

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

#region Pattern Set Policies
WriteWordLine 2 0 "NetScaler Pattern Set Policies"

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

#region Responder Action
WriteWordLine 2 0 "NetScaler Responder Action"

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

#region NetScaler TCP Profiles

WriteWordLine 2 0 "NetScaler TCP Profiles"

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
$AbstractTitle = "Script Template Report"
$SubjectTitle = "Sample Script Template Report"
UpdateDocumentProperties $AbstractTitle $SubjectTitle

If($MSWORD -or $PDF)
{
    SaveandCloseDocumentandShutdownWord
}
ElseIf($Text)
{
    SaveandCloseTextDocument
}
ElseIf($HTML)
{
    SaveandCloseHTMLDocument
}

Write-Verbose "$(Get-Date): Script has completed"
Write-Verbose "$(Get-Date): "

If($PDF)
{
	If(Test-Path "$($Script:FileName2)")
	{
		Write-Verbose "$(Get-Date): $($Script:FileName2) is ready for use"
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
	}
	Else
	{
		Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName1)"
		Write-Error "Unable to save the output file, $($Script:FileName1)"
	}
}

Write-Verbose "$(Get-Date): "

#http://poshtips.com/measuring-elapsed-time-in-powershell/
Write-Verbose "$(Get-Date): Script started: $($Script:StartTime)"
Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
$runtime = $(Get-Date) - $Script:StartTime
$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
	$runtime.Days, `
	$runtime.Hours, `
	$runtime.Minutes, `
	$runtime.Seconds,
	$runtime.Milliseconds)
Write-Verbose "$(Get-Date): Elapsed time: $($Str)"
$runtime = $Null
$Str = $Null
$ErrorActionPreference = $SaveEAPreference
#recommended by webster
#$error
#endregion script template 2
# SIG # Begin signature block
# MIIgCgYJKoZIhvcNAQcCoIIf+zCCH/cCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUm1TUZQHEqPUjDf4KMplAiz6d
# KcCgghtxMIIDtzCCAp+gAwIBAgIQDOfg5RfYRv6P5WD8G/AwOTANBgkqhkiG9w0B
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
# AQICEAmkTdj/HQvKi5Whef7gyA8wDQYJKoZIhvcNAQELBQAwcjELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2ln
# bmluZyBDQTAeFw0xNTEwMjkwMDAwMDBaFw0xNjExMDIxMjAwMDBaMHwxCzAJBgNV
# BAYTAlVTMQswCQYDVQQIEwJUTjESMBAGA1UEBxMJVHVsbGFob21hMSUwIwYDVQQK
# ExxDYXJsIFdlYnN0ZXIgQ29uc3VsdGluZywgTExDMSUwIwYDVQQDExxDYXJsIFdl
# YnN0ZXIgQ29uc3VsdGluZywgTExDMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIB
# CgKCAQEA35g9yG7Fh7/h1rbQmW2x6BmWEWCBw6qwOKfXDJyMMeSunAKZ+rnBYX3K
# T1ERQYMYi2/tK1/hNcgW3ja6sSqwEWBde/nLmqdkzMJb2pUPGUhVP0ZMO7KCS8oz
# Ed5FPpT4Hete/8OQyGKTdU16Ne2xhWzgVvKP1g0zLXJojIWYB4+kKOY2OCl8oPhX
# LwMlQEraFUz39JDkwumteT2/MEjORclAAJ+odAk9R1jjOD5p5GzLRi27vDrBUDq2
# wNsHgejZrq4mbyLiNqdZnFKUeQCzCF8YF32U9E0O+fdhY4QvTM2Jdtusz1d/IIz/
# JqM2AjkDkEXUMK6nQ3015j9yoOAQiQIDAQABo4IBxTCCAcEwHwYDVR0jBBgwFoAU
# WsS5eyoKo6XqcQPAYPkt9mV1DlgwHQYDVR0OBBYEFLdZN8kA2rYz8RkS85RNuO4I
# GxMHMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzB3BgNVHR8E
# cDBuMDWgM6Axhi9odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVk
# LWNzLWcxLmNybDA1oDOgMYYvaHR0cDovL2NybDQuZGlnaWNlcnQuY29tL3NoYTIt
# YXNzdXJlZC1jcy1nMS5jcmwwTAYDVR0gBEUwQzA3BglghkgBhv1sAwEwKjAoBggr
# BgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAIBgZngQwBBAEw
# gYQGCCsGAQUFBwEBBHgwdjAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNl
# cnQuY29tME4GCCsGAQUFBzAChkJodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20v
# RGlnaUNlcnRTSEEyQXNzdXJlZElEQ29kZVNpZ25pbmdDQS5jcnQwDAYDVR0TAQH/
# BAIwADANBgkqhkiG9w0BAQsFAAOCAQEAGz9cEmjU3FosI30XHF355vqavCPByB2F
# TYvGpToMODFnVKul0dQjbF9CWWNeuknYfVjmYBKOgBaFkF/eAy4yfk41tmZZnN9D
# j4Ngenvbrx7ZJqC/ZMNgoIM7un1WLrqZKS5tOaFpBwaEeAIzfU9dHHE27zchIoAJ
# x5aDQbnP6SVWitxa/jGa78b9pDslLpv7Pm4KAEv5d2NYiQ7nhvHShFnWY6wMNBTE
# i+q5rSNcm4TzYsyYSoYT+bGs21vvSAlMSKlvsL0oMWLHMdsMKtC+1Wp2sE4Fshdt
# 9K8DBkl33XhdprC2KabgZa6GTz5NA/rV4FW6oDUidts19XbWIjlB7DCCBmowggVS
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
# aWduaW5nIENBAhAJpE3Y/x0LyouVoXn+4MgPMAkGBSsOAwIaBQCgQDAZBgkqhkiG
# 9w0BCQMxDAYKKwYBBAGCNwIBBDAjBgkqhkiG9w0BCQQxFgQUiwqNK4kVMSCRHs4T
# kd2uBA7Nv18wDQYJKoZIhvcNAQEBBQAEggEAm9jq2CVejmu9TrYoR98cpCT1n3NN
# MqcyHYpn7ySedUoHPjwyXNr63cqPRp9WwPTarcnfCYhGmeBAMsCOusq6y3y8Ar4K
# gOSOHkEDsoEdZAziSL7INapZ9p24IekjEdk7wgDkAwblMl7Jl8Zewc5kq7EfGxyX
# LHrO5aWehdhS6A+wdznoouYpJm4YO0xiCKtpeuR/pTQC1H5pZ59fhbEy5TQdApSF
# Jl8aSdOsoZ6k0IZLjpLnH9HVK73v0lQIXEpsvz+DVyIB5d/8MZND485OuAxgnlGO
# 0JM/Z3vlsHdY2UMK2QYZ4nprQHD0+3U6QtciTnOJsf7vBkF0PrcUJQD9paGCAg8w
# ggILBgkqhkiG9w0BCQYxggH8MIIB+AIBATB2MGIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAf
# BgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMQIQAwGaAjr/WLFr1tXq5hfw
# ZjAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMTYwNTIyMjE0MjE5WjAjBgkqhkiG9w0BCQQxFgQUOpmHL/JKr+sS
# mPLlxU5aj673LFswDQYJKoZIhvcNAQEBBQAEggEAFoffgnGzklKy/rkWW6LOJqZZ
# Dy9XPVmapwMTNJ7qrOcwGKBoENlJbFqNX2E1APTnbL2DFHBf/8o+tEZzMIka64pQ
# KYRk3L5miW316bqg4ZG/JbkIzjfORfnZEm+pmKMfx5HpXIUp6NjJ6UoRn4oUEPfE
# AC50r3tJoJ7CwpwhnFJfT2PTYQKUuZ9CnvBEQSYVHfxOhO3vDkO5if5ETmCZNVYm
# 8KlNd66Obo34YmHMRetaRk0covWbb7/fuSzLGmo9nJf0RruP4T4QLLJRR46gGFgu
# 5rUapmhiqWCuKnXPTPAvaBS9Pncb63SKnjgTsXvNHZIwjXbMp6aKchL4+ZWzbg==
# SIG # End signature block
