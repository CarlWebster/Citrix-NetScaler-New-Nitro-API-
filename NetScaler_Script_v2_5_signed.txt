#Requires -Version 3.0
#This File is in Unicode format.  Do not edit in an ASCII editor.

#region Support

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
	This parameter is reserved for a future update and no output is created at this time.
.PARAMETER MSWord
	SaveAs DOCX file
	This parameter is set True if no other output format is selected.
.PARAMETER HTML
	Creates an HTML file with an .html extension.
	This parameter is disabled by default.
	This parameter is reserved for a future update and no output is created at this time.
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
	PS C:\PSScript > .\NetScaler_Script_v2_5.ps1
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Barry Schiffer" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Barry Schiffer"
	$env:username = Administrator

	Barry Schiffer for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v2_5.ps1 -PDF
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Barry Schiffer" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Barry Schiffer"
	$env:username = Administrator

	Barry Schiffer for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v2_5.ps1 -TEXT
	
	Will use all default values and save the document as a formatted text file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Barry Schiffer" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Barry Schiffer"
	$env:username = Administrator

	Barry Schiffer for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v2_5.ps1 -HTML
	
	Will use all default values and save the document as an HTML file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Barry Schiffer" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Barry Schiffer"
	$env:username = Administrator

	Barry Schiffer for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript .\NetScaler_Script_v2_5.ps1 -CompanyName "Barry Schiffer Consulting" -CoverPage "Mod" -UserName "Barry Schiffer"

	Will use:
		Barry Schiffer Consulting for the Company Name.
		Mod for the Cover Page format.
		Barry Schiffer for the User Name.
.EXAMPLE
	PS C:\PSScript .\NetScaler_Script_v2_5.ps1 -CN "Barry Schiffer Consulting" -CP "Mod" -UN "Barry Schiffer"

	Will use:
		Barry Schiffer Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Barry Schiffer for the User Name (alias UN).
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v2_5.ps1 -AddDateTime
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Barry Schiffer" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Barry Schiffer"
	$env:username = Administrator

	Barry Schiffer for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2014 at 6PM is 2014-06-01_1800.
	Output filename will be Script_Template_2014-06-01_1800.docx
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v2_5.ps1 -PDF -AddDateTime
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Barry Schiffer" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Barry Schiffer"
	$env:username = Administrator

	Barry Schiffer for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2014 at 6PM is 2014-06-01_1800.
	Output filename will be Script_Template_2014-06-01_1800.PDF
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v2_5.ps1 -Hardware
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Barry Schiffer" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Barry Schiffer"
	$env:username = Administrator

	Barry Schiffer for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	localhost for running hardware inventory.
	localhost will be replaced by the actual computer name.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_Script_v2_5.ps1 -Hardware -ComputerName 192.168.1.51
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Barry Schiffer" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Barry Schiffer"
	$env:username = Administrator

	Barry Schiffer for the Company Name.
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
	NAME: NetScaler_Script_v2_5_unsigned.ps1
	VERSION: 16122014
	AUTHOR: Carl Webster, Michael B. Smith, Iain Brighton, Jeff Wouters, Barry Schiffer
	LASTEDIT: December 16, 2014
#>

#endregion Support

#region script template
#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "WordOrPDF") ]

Param(
	[parameter(ParameterSetName="WordOrPDF",
	Position = 0, 
	Mandatory=$False )
	] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(ParameterSetName="WordOrPDF",
	Position = 1, 
	Mandatory=$False )
	] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(ParameterSetName="WordOrPDF",
	Position = 2, 
	Mandatory=$False )
	] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

	[parameter(ParameterSetName="WordOrPDF",
	Position = 3, 
	Mandatory=$False )
	] 
	[Switch]$PDF=$False,

	[parameter(ParameterSetName="Text",
	Position = 4, 
	Mandatory=$False )
	] 
	[Switch]$Text=$False,

	[parameter(ParameterSetName="WordOrPDF",
	Position = 4, 
	Mandatory=$False )
	] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="HTML",
	Position = 4, 
	Mandatory=$False )
	] 
	[Switch]$HTML=$False,

	[parameter(
	Position = 5, 
	Mandatory=$False )
	] 
	[Switch]$AddDateTime=$False,
	
	[parameter(
	Position = 6, 
	Mandatory=$False )
	] 
	[Switch]$Hardware=$False,

	[parameter(
	Position = 7, 
	Mandatory=$False )
	] 
	[string]$ComputerName="LocalHost"
	
	)
	
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#Created on June 1, 2014

<#
.NetScaler Documentation Script
    NAME: NetScaler_Script_v2_5.ps1
	VERSION NetScaler Script: 2.5
	VERSION Script Template: 16122014
	AUTHOR NetScaler script: Barry Schiffer
    AUTHOR NetScaler script functions: Iain Brighton
    AUTHOR Script template: Carl Webster, Michael B. Smith, Iain Brighton, Jeff Wouters
	LASTEDIT: December 16, 2014  
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

Set-StrictMode -Version 2

#force -verbose on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$SaveEAPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'

If($PDF -eq $Null)
{
	$PDF = $False
}
If($Text -eq $Null)
{
	$Text = $False
}
If($MSWord -eq $Null)
{
	$MSWord = $False
}
If($HTML -eq $Null)
{
	$HTML = $False
}
If($AddDateTime -eq $Null)
{
	$AddDateTime = $False
}
If($Hardware -eq $Null)
{
	$Hardware = $False
}
If($ComputerName -eq $Null)
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

If($MSWord -or $PDF)
{
	#try and fix the issue with the $CompanyName variable
	$CoName = $CompanyName
	Write-Verbose "$(Get-Date): CoName is $($CoName)"
	
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
	[int]$wdColorRed = 255
	[int]$wdColorBlack = 0
	[int]$wdWord2007 = 12
	[int]$wdWord2010 = 14
	[int]$wdWord2013 = 15
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

	[string]$RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption
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

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)
	
	$xArray = ""
	
	Switch ($CultureCode)
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2013)
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabet", "Anual", "Conservador", "Contrast",
					"Cubicles", "Diplomàtic", "En mosaic", "Exposició", "Línia lateral",
					"Mod", "Moviment", "Piles", "Sobri", "Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2013)
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabet", "Årlig", "BevægElse", "Eksponering",
					"Enkel", "Firkanter", "Fliser", "Gåde", "Kontrast",
					"Mod", "Nålestribet", "Overskrid", "Sidelinje", "Stakke",
					"Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2013)
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alphabet", "Bewegung", "Durchscheinend", "Herausgestellt",
					"Jährlich", "Kacheln", "Kontrast", "Kubistisch", "Modern",
					"Nadelstreifen", "Puzzle", "Randlinie", "Raster", "Schlicht", "Stapel",
					"Traditionell")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013)
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Conservative", "Contrast",
					"Cubicles", "Exposure", "Mod", "Motion", "Pinstripes", "Puzzle",
					"Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2013)
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Conservador",
					"Contraste", "Cubículos", "Exposición", "Línea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Pilas", "Puzzle",
					"Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2013)
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Aakkoset", "Alttius", "Kontrasti", "Kuvakkeet ja tiedot",
					"Liike" , "Liituraita" , "Mod" , "Palapeli", "Perinteinen", "Pinot",
					"Sivussa", "Työpisteet", "Vuosittainen", "Yksinkertainen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("ViewMaster", "Secteur (foncé)", "Sémaphore",
					"Rétrospective", "Ion (foncé)", "Ion (clair)", "Intégrale",
					"Filigrane", "Facette", "Secteur (clair)", "À bandes", "Austin",
					"Guide", "Whisp", "Lignes latérales", "Quadrillage")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Mosaïques", "Ligne latérale", "Annuel", "Perspective",
					"Contraste", "Emplacements de bureau", "Moderne", "Blocs empilés",
					"Rayures fines", "Austère", "Transcendant", "Classique", "Quadrillage",
					"Exposition", "Alphabet", "Mots croisés", "Papier journal", "Austin", "Guide")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alphabet", "Annuel", "Austère", "Blocs empilés", "Blocs superposés",
					"Classique", "Contraste", "Exposition", "Guide", "Ligne latérale", "Moderne",
					"Mosaïques", "Mots croisés", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013)
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabet", "Årlig", "Avlukker", "BevegElse", "Engasjement",
					"Enkel", "Fliser", "Konservativ", "Kontrast", "Mod", "Puslespill",
					"Sidelinje", "Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013)
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Bescheiden", "Beweging",
					"Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks", "Krijtstreep",
					"Mod", "Puzzel", "Stapels", "Tegels", "Terzijde", "Werkplekken")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013)
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Baias", "Conservador",
					"Contraste", "Exposição", "Ladrilhos", "Linha Lateral", "Listras", "Mod",
					"Pilhas", "Quebra-cabeça", "Transcendente")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013)
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabetmönster", "Årligt", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Övergående", "Plattor", "Pussel", "RörElse",
					"Sidlinje", "Sobert", "Staplat")
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid", "Integral",
						"Ion (Dark)", "Ion (Light)", "Motion", "Retrospect", "Semaphore",
						"Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster", "Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
					ElseIf($xWordVersion -eq $wdWord2007)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Conservative", "Contrast",
						"Cubicles", "Exposure", "Mod", "Motion", "Pinstripes", "Puzzle",
						"Sideline", "Stacks", "Tiles", "Transcend")
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

Function CheckWord2007SaveAsPDFInstalled
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Installer\Products\000021090B0090400000000000F01FEC) -eq $False)
	{
		Write-Host "`n`n`t`tWord 2007 is detected and the option to SaveAs PDF was selected but the Word 2007 SaveAs PDF add-in is not installed."
		Write-Host "`n`n`t`tThe add-in can be downloaded from http://www.microsoft.com/en-us/download/details.aspx?id=9943"
		Write-Host "`n`n`t`tInstall the SaveAs PDF add-in and rerun the script."
		Return $False
	}
	Return $True
}

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

Function SetupWord
{
	Write-Verbose "$(Get-Date): Setting up Word"
    
	# Setup word for output
	Write-Verbose "$(Get-Date): Create Word comObject.  If you are not running Word 2007, ignore the next message."
	$Script:Word = New-Object -comobject "Word.Application" -EA 0

	If(!$? -or $Script:Word -eq $Null)
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
	If($Script:WordVersion -eq $wdWord2013)
	{
		$Script:WordProduct = "Word 2013"
	}
	ElseIf($Script:WordVersion -eq $wdWord2010)
	{
		$Script:WordProduct = "Word 2010"
	}
	ElseIf($Script:WordVersion -eq $wdWord2007)
	{
		$Script:WordProduct = "Word 2007"
	}
	Else
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tYou are running an untested or unsupported version of Microsoft Word.`n`n`t`tScript will end.`n`n`t`tPlease send info on your version of Word to webster@carlwebster.com`n`n"
		AbortScript
	}

	If($PDF -and $Script:WordVersion -eq $wdWord2007)
	{
		Write-Verbose "$(Get-Date): Verify the Word 2007 Save As PDF add-in is installed"
		If(CheckWord2007SaveAsPDFInstalled)
		{
			Write-Verbose "$(Get-Date): The Word 2007 Save As PDF add-in is installed"
		}
		Else
		{
			AbortScript
		}
	}

	#only validate CompanyName if the field is blank
	If([String]::IsNullOrEmpty($CoName))
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
						If($Script:WordVersion -eq $wdWord2013)
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
	If($Script:WordVersion -eq $wdWord2007)
	{
		$BuildingBlocksCollection = $Script:Word.Templates | Where {$_.name -eq "Building Blocks.dotx"}
	}
	Else
	{
		#word 2010/2013
		$BuildingBlocksCollection = $Script:Word.Templates | Where {$_.name -eq "Built-In Building Blocks.dotx"}
	}

	Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
	$part = $Null

	$BuildingBlocksCollection | 
	ForEach{
		If ($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) 
		{
			$BuildingBlocks = $_
		}
	}        

	If($BuildingBlocks -ne $Null)
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

		If($part -ne $Null)
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
	If($Script:Doc -eq $Null)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn empty Word document could not be created.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If($Script:Selection -eq $Null)
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
		Write-Verbose "$(Get-Date): Table of Contents - $($myHash.Word_TableOfContents)"
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($myHash.Word_TableOfContents)
		If($toc -eq $Null)
		{
			Write-Verbose "$(Get-Date): "
			Write-Verbose "$(Get-Date): Table of Content - $($myHash.Word_TableOfContents) could not be retrieved."
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

Function SaveandCloseDocumentandShutdownWord
{
	#bug fix 1-Apr-2014
	#reset Grammar and Spelling options back to their original settings
	$Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
	$Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

	Write-Verbose "$(Get-Date): Save and Close document and Shutdown Word"
	If($Script:WordVersion -eq $wdWord2007)
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
		Write-Verbose "$(Get-Date): Running Word 2007 and detected operating system $($RunningOS)"
		If($RunningOS.Contains("Server 2008 R2") -or $RunningOS.Contains("Server 2012"))
		{
			$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type] 
			$Script:Doc.SaveAs($Script:FileName1, $SaveFormat)
			If($PDF)
			{
				Write-Verbose "$(Get-Date): Now saving as PDF"
				$SaveFormat = $wdFormatPDF
				$Script:Doc.SaveAs($Script:FileName2, $SaveFormat)
			}
		}
		Else
		{
			#works for Server 2008 and Windows 7
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
			$Script:Doc.SaveAs([REF]$Script:FileName1, [ref]$SaveFormat)
			If($PDF)
			{
				Write-Verbose "$(Get-Date): Now saving as PDF"
				$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
				$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$saveFormat)
			}
		}
	}
	ElseIf($Script:WordVersion -eq $wdWord2010)
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
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:FileName1, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$saveFormat)
		}
	}
	ElseIf($Script:WordVersion -eq $wdWord2013)
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
		Write-Verbose "$(Get-Date): Running Word 2013 and detected operating system $($RunningOS)"
		#$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs2([REF]$Script:FileName1, [ref]$wdFormatDocumentDefault)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			#$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$wdFormatPDF)
		}
	}

	Write-Verbose "$(Get-Date): Closing Word"
	$Script:Doc.Close()
	$Script:Word.Quit()
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Deleting $($Script:FileName1) since only $($Script:FileName2) is needed"
		Remove-Item $Script:FileName1
	}
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global
	}
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
}

Function SaveandCloseTextDocument
{
	If($AddDateTime)
	{
		$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	}

	Write-Output $Global:Output | Out-File $Script:Filename1
}

Function SaveandCloseHTMLDocument
{
	If($AddDateTime)
	{
		$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).html"
	}
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

#region NetScaler Documentation Functions

<#
.SYNOPSIS
   Get a named property value from a string.
.DESCRIPTION
   Returns a case-insensitive property from a string, assuming the property is
   named before the actual property value and is separated by a space. For
   example, if the specified SearchString contained "-property1 <value1>
   -property2 <value2>�, searching for "-Property1" would return "<value1>".
.PARAMETER SearchString
   String to search for the specified property name.
.PARAMETER PropertyName
   The property name to search the SearchString for.
.PARAMETER Default
   If the property is not found returns the specified string. This parameter is
   optional and if not specified returns $null (by default) if the property is
   not found.
.PARAMETER RemoveQuotes
    Removes quotes from returned property values if present.
.PARAMETER ReplaceEscapedQuotes
    Replaces escaped quotes (\") with quotes (") from the returned property values
    if present. Note: This is generally used for display purposes only.
.EXAMPLE
   Get-StringProperty -SearchString $StringToSearch -PropertyName "-property1"

   This command searches the $StringToSearch variable for the presence of the property
   "-property1" and returns its value, if found. If the property name is not found,
   the default $null will be returned.
.EXAMPLE
   Get-StringProperty $StringToSearch "-property3" "Not found"

   This command searches the $StringToSearch variable for the presence of the property
   "-property3" and returns its value, if found. If the property name is not found,
   the "Not Found" string will be returned.
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>


function Get-StringProperty {
    [CmdletBinding(HelpUri = 'http://virtualengine.co.uk/2014/searching-for-string-properties-with-powershell/')]
    [OutputType([String])]
    Param (
        # String to search
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
        [ValidateNotNullOrEmpty()] [Alias("Search")] [string] $SearchString,
        # String property name to search for
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=1)]
        [ValidateNotNullOrEmpty()] [Alias("Name","Property")] [string] $PropertyName,
        # Default return value for missing values
        [Parameter(ValueFromPipelineByPropertyName=$true, Position=2)]
        [AllowNull()] [String] $Default = $null,
        # String delimiter, default to one or more spaces
        [Parameter(ValueFromPipelineByPropertyName=$true, Position=3)]
        [ValidateNotNullOrEmpty()] [string] $Delimiter = ' ',
        # Remove quotes from quoted strings
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Alias("NoQuotes")] [Switch] $RemoveQuotes,
        # Replace escaped quotes with quotes
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Switch] $ReplaceEscapedQuotes
    )

    Process {
        # First replace escaped quotes with '�'
        $SearchString = $SearchString.Replace('\"', "�");
        # Locate and replace quotes with '^^' and quoted spaces '^' to aid with parsing, until there are none left
        while ($SearchString.Contains('"')) {
            # Store the right-hand side temporarily, skipping the first quote
            $searchStringRight = $SearchString.Substring($SearchString.IndexOf('"') +1);
            # Extract the quoted text from the original string
            $quotedString = $SearchString.Substring($SearchString.IndexOf('"'), $searchStringRight.IndexOf('"') +2);
            # Replace the quoted text, replacing spaces with '^' and quotes with '^^'
            $SearchString = $SearchString.Replace($quotedString, $quotedString.Replace(" ", "^").Replace('"', "^^"));
        }
 
        # Split the $SearchString based on one or more blank spaces
        $stringComponents = $SearchString.Split($Delimiter,[StringSplitOptions]'RemoveEmptyEntries'); 
        for ($i = 0; $i -le $stringComponents.Length; $i++) {
            # The standard Powershell CompareTo method is case-sensitive
            if ([string]::Compare($stringComponents[$i], $PropertyName, $True) -eq 0) {
                # Check that we're not over the array boundary
                if ($i+1 -le $stringComponents.Length) {
                    # Restore any escaped quotation marks and spaces
                    $propertyValue = $stringComponents[$i+1].Replace("^^", '"').Replace("^", " ");
                    # Remove quotes
                    if ($RemoveQuotes) { $propertyValue = $propertyValue.Trim('"'); }
                    # Replace escaped quotes
                    if ($ReplaceEscapedQuotes) { return $propertyValue.Replace('�','"'); }
                    else { return $propertyValue.Replace('�','\"'); }
                }
            }
        }
        # If nothing has been found or we're over the array boundary, return the default value
        return $Default;
    }
}


<#
.SYNOPSIS
   Get an array of properies from a delimited string.
.DESCRIPTION
   The Get-StringProperySplit cmdlet returns an array of space-separated 
   strings from the source string, accounting for quoted text and escaped
   quotations.

   A single string can be returned with the -Index parameter. This parameter
   shortcuts allows replacing calls like '(Get-StringPropertySplit
    -SearchString $Source)[3]' with 'Get-StringPropertySplit -SearchString
    $Source -Index 3'
.PARAMETER SearchString
   String to search for the specified property name.
.PARAMETER Delimter
   The delimiter string/char to use. Regular expressions are supported and defaults to ' +'.
.PARAMETER RemoveQuotes
    Removes quotes from returned property values if present.
.PARAMETER ReplaceEscapedQuotes
    Replaces escaped quotes (\") with quotes (") from the returned property values if present.
.PARAMETER Index
    Returns the [string] at the specified index rather than a [string[]]
.EXAMPLE
   Get-StringPropertySplit -SearchString $StringToSearch

   This command returns an array of strings for all space-delimited values in the
   $StringToSearch variable, accounting for quoted strings and escaped quotes.
.EXAMPLE
   Get-StringPropertySplit $StringToSearch -RemoveQuotes

   This command returns an array of strings for all space-delimited values in the
   $StringToSearch variable, accounting for quoted strings and escaped quotes. All
   quotation marks are removed from quoted strings.
.EXAMPLE
   Get-StringPropertySplit $StringToSearch -ReplaceEscapedQuotes -Index 2

   This command returns a single string for the space-delimied value at array
   index 2 (third element), accounting for quoted strings and escaped quotes. The
   return string will have all escaped quotes '\"' replaced with '"'.
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>

function Get-StringPropertySplit {
    [CmdletBinding(HelpUri = 'http://virtualengine.co.uk/2014/searching-for-string-properties-with-powershell/')]
    [OutputType([String[]])]
    Param (
        # String to search
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
        [ValidateNotNullOrEmpty()] [Alias("Search")] [string] $SearchString,
        # String delimiter, default to one or more spaces
        [Parameter(ValueFromPipelineByPropertyName=$true, Position=1)]
        [ValidateNotNullOrEmpty()] [string] $Delimiter = ' +',
        # Remove quotes from quoted strings
        [Parameter(ValueFromPipelineByPropertyName=$true)] [Switch] $RemoveQuotes,
        # Replace escaped quotes with quotes
        [Parameter(ValueFromPipelineByPropertyName=$true)] [Switch] $ReplaceEscapedQuotes,
        # Return the specified index
        [Parameter(ValueFromPipelineByPropertyName=$true)] [int] $Index = -1
    )

    Process {
        # First replace escaped quotes with '�'
        $SearchString = $SearchString.Replace('\"', '�');
        
        while ($SearchString.Contains('"')) {
            # Store the right-hand side temporarily, skipping the first quote
            $searchStringRight = $SearchString.Substring($SearchString.IndexOf('"') +1);
            # Extract the quoted text from the original string
            $quotedString = $SearchString.Substring($SearchString.IndexOf('"'), $searchStringRight.IndexOf('"') +2);
            # Replace the quoted text, replacing spaces with '^' and quotes with '^^'
            $SearchString = $SearchString.Replace($quotedString, $quotedString.Replace(' ', '^').Replace('"', '^^'));
        }

        $stringArray = $SearchString.Split($Delimiter,[StringSplitOptions]'RemoveEmptyEntries'); 
        # Replace all escaped characters
        for ($i = 0; $i -lt $StringArray.Length; $i++) { 
            $stringArray[$i] = $stringArray[$i].Replace('^^', '"').Replace('^', ' ');
            # Remove quotes
            if ($RemoveQuotes) { $stringArray[$i] = $stringArray[$i].Trim('"'); }
            # Replace escaped quotes
            if ($ReplaceEscapedQuotes) { $stringArray[$i] = $stringArray[$i].Replace('�','"'); }
            else { $stringArray[$i] = $stringArray[$i].Replace('�','\"'); }
        }

        if ($Index -ne -1) { return $stringArray[$Index]; }
        else { return $stringArray; }
    }
}

<#
.SYNOPSIS
   Gets the NetScaler expression from the specified string
.DESCRIPTION
   This cmdlet returns a NetScaler expression that is escaped with 'q/' and is
   terminated by '/'. If a NetScaler expression is not found, $null is returned.
.PARAMETER SearchString
   String to search for the specified property name.
EXAMPLE
   Get-StringProperty -SearchString $StringToSearch

   This command searches the $StringToSearch variable for the presence of the a
   NetScaler expression, i.e. q/ .. /
.EXAMPLE
   Get-StringProperty $StringToSearch "-property3" "Not found"

   This command searches the $StringToSearch variable for the presence of the property
   "-property3" and returns its value, if found. If the property name is not found,
   the "Not Found" string will be returned.
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>

function Get-NetScalerExpression {
    [CmdletBinding()]
    Param (
        # String to search
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
        [ValidateNotNullOrEmpty()] [Alias("Search")] [string] $SearchString
    )

    Process {
        $searchStringLeftPosition = $SearchString.IndexOf('q/');
        if ($searchStringLeftPosition -eq -1) { return $null; }
        $SearchString = $SearchString.Replace('q/', '��');

        $searchStringRightPosition = $SearchString.IndexOf('/');
        if ($searchStringRightPosition -eq -1) { return $null; }
        $SearchString = $SearchString.Replace('/', '�');

        $NetScalerExpression = $SearchString.Substring($searchStringLeftPosition, (($searchStringRightPosition +1)- $searchStringLeftPosition));
        return $NetScalerExpression.Replace('��','q/').Replace('�','/');
    }
}

<#
.SYNOPSIS
   Test for a named property value in a string.
.DESCRIPTION
   Tests for the presence of a property value in a string and returns a boolean
   value. For example, if the specified SearchString contained "-property1
   -property2 <value2>�, searching for "-Property1" or "-Property2" would return
   $true, but searching for "-Property3" would return $false
.PARAMETER SearchString
   String to search for the specified property name.
.PARAMETER PropertyName
   The property name to search the SearchString for.
.EXAMPLE
   Test-StringProperty -SearchString $StringToSearch -PropertyName "-property1"

   This command searches the $StringToSearch variable for the presence of the property
   "-property1". If the property name is found it returns $true. If the property name
   is not found, it will return $false.
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>

function Test-StringProperty {
    [CmdletBinding(HelpUri = 'http://virtualengine.co.uk/2014/searching-for-string-properties-with-powershell/')]
    [OutputType([bool])]
    Param (
        # String to search
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
        [ValidateNotNullOrEmpty()] [Alias("Search")] [string] $SearchString,
        # String property name to search for
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=1)]
        [ValidateNotNullOrEmpty()] [Alias("Name","Property")] [string] $PropertyName
    )

    Process {
        # Split the $SearchString based on one or more blank spaces
        $stringComponents = $SearchString.Split(' +',[StringSplitOptions]'RemoveEmptyEntries'); 
        for ($i = 0; $i -le $stringComponents.Length; $i++) {
            # The standard Powershell CompareTo method is case-sensitive
            if ([string]::Compare($stringComponents[$i], $PropertyName, $True) -eq 0) { return $true; }
        }
        # If nothing has been found or we're over the array boundary, return the default value
        return $false;
    }
}

<#
.SYNOPSIS
    Tests whether a named property in a string exists and returns either Yes
    ($true) or No ($false)
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>

function Test-StringPropertyYesNo([string]$SearchString, [string]$SearchProperty)
{
    if (Test-StringProperty $SearchString $SearchProperty) { return "Yes"; }
    else { return "No"; }
}

<#
.SYNOPSIS
    Tests whether a named property in a string does not exist and returns
    either Yes ($false) or No ($true)
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>

function Test-NotStringPropertyYesNo([string]$SearchString, [string]$SearchProperty)
{
    if (-not (Test-StringProperty $SearchString $SearchProperty)) { return "Yes"; }
    else { return "No"; }
}

<#
.SYNOPSIS
    Tests whether a named property in a string exists and returns either Enabled
    ($true) or Disabled ($false)
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>

function Test-StringPropertyEnabledDisabled([string]$SearchString, [string]$SearchProperty)
{
    if (Test-StringProperty $SearchString $SearchProperty) { return "Enabled"; }
    else { return "Disabled"; }
}

<#
.SYNOPSIS
    Tests whether a named property in a string exists and returns either Disabled
    ($true) or Enabled ($false)
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>

function Test-NotStringPropertyEnabledDisabled([string]$SearchString, [string]$SearchProperty)
{
    if (-not (Test-StringProperty $SearchString $SearchProperty)) { return "Enabled"; }
    else { return "Disabled"; }
}

<#
.SYNOPSIS
    Tests whether a named property in a string exists and returns either On
    ($true) or Off ($false)
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>

function Test-StringPropertyOnOff([string]$SearchString, [string]$SearchProperty)
{
    if (Test-StringProperty $SearchString $SearchProperty) { return "On"; }
    else { return "Off"; }
}

<#
.SYNOPSIS
    Tests whether a named property in a string exists and returns either Off
    ($true) or On ($false)
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>

function Test-NotStringPropertyOnOff([string]$SearchString, [string]$SearchProperty)
{
    if (-not (Test-StringProperty $SearchString $SearchProperty)) { return "On"; }
    else { return "Off"; }
}

<#
.SYNOPSIS
    Returns all strings that include the specified string $PropertyName parameter from 
    the array of string values passed into the $SearchString parameter.
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>

function Get-StringWithProperty {
    [CmdletBinding(DefaultParameterSetName='PropertyName')]
    [OutputType([string[]])]
    Param (
        [Parameter(Mandatory=$true, Position=0)] [ValidateNotNullOrEmpty()] [string[]] $SearchString,
        [Parameter(Mandatory=$true, Position=1, ParameterSetName='PropertyName')] [ValidateNotNullOrEmpty()] [string] $PropertyName,
        [Parameter(Mandatory=$true, Position=1, ParameterSetName='Like')] [ValidateNotNullOrEmpty()] [string] $Like
    )

    Begin {
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
        ## Check that we have a wildcard, somewhere
        if ($PSCmdlet.ParameterSetName -eq 'Like') {
            if (!$Like.Contains('*')) {
                ## If not do we need to add '*' or ' *'?
                if ($Like.EndsWith(' ')) {
                    Write-Warning "Get-StringWithProperty: No wildcard specified and '*' was appended.";
                    $Like += "*";
                } else {
                    Write-Warning "Get-StringWithProperty: No wildcard specified and ' *' was appended.";
                    $Like += " *";
                } # end if
            }
        }
	}

    Process {
        $MatchingStrings = @();
        foreach ($String in $SearchString) {
            switch ($PSCmdlet.ParameterSetName) {
                'PropertyName' {
                    if (Test-StringProperty -SearchString $String -PropertyName $PropertyName) {
                        $MatchingStrings += $String;
                    } # end if
                } # end propertyname
                'Like' {
                    if ($String -like $Like) {
                        $MatchingStrings += $String;
                    } #end if
                } # end like
            } # end switch
        } #end foreach
        return ,$MatchingStrings;
    } # end process
}

#endregion NetScaler Documentation Functions

#region NetScaler documentation pre-requisites

$Scriptver = 1
$SourceFileName = "ns.conf";

## Iain Brighton - Try and resolve the ns.conf file in the current working directory
if(Test-Path (Join-Path ((Get-Location).ProviderPath) $SourceFileName)) 
{
	$SourceFile = Join-Path ((Get-Location).ProviderPath) $SourceFileName; 
}
else 
{
	## Otherwise try the script's directory
	if(Test-Path (Join-Path (Split-Path $MyInvocation.MyCommand.Path) $SourceFileName)) 
	{
		$SourceFile = Join-Path (Split-Path $MyInvocation.MyCommand.Path) $SourceFileName; 
	}
	else 
	{
		throw "Cannot locate a NetScaler ns.conf file in either the working or script directory."; 
	}
}

#added by Carl Webster 24-May-2014
If(!$?)
{
	Write-Error "`n`n`t`tCannot locate a NetScaler ns.conf file in either the working or script directory.`n`n`t`tScript cannot continue.`n`n"
	AbortScript
}

Write-Verbose "$(Get-Date): NetScaler file : $SourceFile"

## We read the file in once as each Get-Content call goes to disk and also creates a new string[]
$File = Get-Content $SourceFile

#added by Carl Webster 24-May-2014
If(!$? -or $File -eq $Null)
{
	Write-Error "`n`n`t`tUnable to read the NetScaler ns.conf file.`n`n`t`tScript cannot continue.`n`n"
	AbortScript
}
#endregion NetScaler documentation pre-requisites

#region NetScaler Create Collections

## Create smart and smaller collections for faster processing of the script.

$Add = Get-StringWithProperty -SearchString $File -Like 'add *';
$Set = Get-StringWithProperty -SearchString $File -Like 'set *';
$Bind = Get-StringWithProperty -SearchString $File -Like 'bind *';
$Enable = Get-StringWithProperty -SearchString $File -Like 'enable *';
$SetNS = Get-StringWithProperty -SearchString $Set -Like 'set ns *';
$SetNSTCPPARAM = Get-StringWithProperty -SearchString $SetNS -Like 'set ns tcpParam *';
$SetVpnParameter = Get-StringWithProperty -SearchString $Set -Like 'set vpn parameter *';
$SETAAAPREAUTH = Get-StringWithProperty -SearchString $Set -Like 'set aaa preauthenticationparameter *';
$AddCSPOLICY = Get-StringWithProperty -SearchString $Add -Like 'add cs policy *';
$ContentSwitches = Get-StringWithProperty -SearchString $Add -Like 'add cs vserver *';
$ContentSwitchBind = Get-StringWithProperty -SearchString $Bind -Like 'bind cs vserver *';
$CACHEREDIRS = Get-StringWithProperty -SearchString $Add -Like 'add cr vserver *';
$LoadBalancers = Get-StringWithProperty -SearchString $Add -Like 'add lb vserver *';
$LoadBalancerBind = Get-StringWithProperty -SearchString $Bind -Like 'bind lb vserver *';
$ServiceGroups = Get-StringWithProperty -SearchString $Add -Like 'add servicegroup *';
$ServiceGroupBind = Get-StringWithProperty -SearchString $Bind -Like 'bind servicegroup *';
$ServiceS = Get-StringWithProperty -SearchString $Add -Like 'add service *';
$ServiceBind = Get-StringWithProperty -SearchString $Bind -Like 'bind service *';
$Servers = Get-StringWithProperty -SearchString $Add -Like 'add server *';
$MONITORS = Get-StringWithProperty -SearchString $Add -Like 'add lb monitor *';
$NICS = Get-StringWithProperty -SearchString $Set -Like 'set interface *';
$CHANNELS = Get-StringWithProperty -SearchString $Set -Like 'set channel *';
$SIMPLEACLS = Get-StringWithProperty -SearchString $Add -Like 'add ns simpleacl *';
$IPList = Get-StringWithProperty -SearchString $Add -Like 'add ns ip *';
$VLANS = Get-StringWithProperty -SearchString $Add -Like 'add vlan *';
$VLANSBIND = Get-StringWithProperty -SearchString $BIND -Like 'bind vlan *';
$ROUTES = Get-StringWithProperty -SearchString $Add -Like 'add route *';
$AccessGateways = Get-StringWithProperty -SearchString $Add -Like 'add vpn vserver *';
$DNSNAMESERVERS = Get-StringWithProperty -SearchString $Add -Like 'add dns nameServer *';
$DNSRECORDCONFIGS = Get-StringWithProperty -SearchString $Add -Like 'add dns addRec *';
$CERTS = Get-StringWithProperty -SearchString $Add -Like 'add ssl certKey *';
$CERTBINDS = Get-StringWithProperty -SearchString $Bind -Like 'bind ssl vserver *';
$AUTHLDAPACTS = Get-StringWithProperty -SearchString $Add -Like 'add authentication ldapAction*';
$AUTHLDAPPOLS = Get-StringWithProperty -SearchString $Add -Like 'add authentication ldapPolicy*';
$AUTHRADIUSS = Get-StringWithProperty -SearchString $Add -Like 'add authentication radiusAction*';
$AUTHRADPOLS = Get-StringWithProperty -SearchString $Add -Like 'add authentication radiusPolicy*';
$AUTHGRPS = Get-StringWithProperty -SearchString $Add -Like 'add system group *';
$AUTHLOCS = Get-StringWithProperty -SearchString $Add -Like 'add system user *';
$AUTHLOCUSERS = Get-StringWithProperty -SearchString $Add -Like 'add authentication localPolicy *';
$CAGSESSIONPOLS = Get-StringWithProperty -SearchString $ADD -Like "add vpn sessionPolicy *";
$CAGSESSIONACTS = Get-StringWithProperty -SearchString $ADD -Like "add vpn sessionAction *";
$BINDVPNVSERVER = Get-StringWithProperty -SearchString $BIND -Like "bind vpn vserver *";
$CAGURLPOLS = Get-StringWithProperty -SearchString $ADD -Like "add vpn url *";
#endregion NetScaler Create Collections

#region NetScaler chaptercounters
$Chapters = 32
$Chapter = 0
#endregion NetScaler chaptercounters

#region NetScaler feature state
##Getting Feature states for usage later on and performance enhancements by not running parts of the script when feature is disabled
$Enable | foreach { 
    if ($_ -like 'enable ns feature *') {
        If ($_.Contains("WL") -eq "True") {$FEATWL = "Enabled"} Else {$FEATWL = "Disabled"}
        If ($_.Contains(" SP ") -eq "True") {$FEATSP = "Enabled"} Else {$FEATSP = "Disabled"}
        If ($_.Contains("LB") -eq "True") {$FEATLB = "Enabled"} Else {$FEATLB = "Disabled"}
        If ($_.Contains("CS") -eq "True") {$FEATCS = "Enabled"} Else {$FEATCS = "Disabled"}
        If ($_.Contains("CR") -eq "True") {$FEATCR = "Enabled"} Else {$FEATCR = "Disabled"}
        If ($_.Contains("SC") -eq "True") {$FEATSC = "Enabled"} Else {$FEATSC = "Disabled"}
        If ($_.Contains("CMP") -eq "True") {$FEATCMP = "Enabled"} Else {$FEATCMP = "Disabled"}
        If ($_.Contains("PQ") -eq "True") {$FEATPQ = "Enabled"} Else {$FEATPQ = "Disabled"}
        If ($_.Contains("SSL") -eq "True") {$FEATSSL = "Enabled"} Else {$FEATSSL = "Disabled"}
        If ($_.Contains("GSLB") -eq "True") {$FEATGSLB = "Enabled"} Else {$FEATGSLB = "Disabled"}
        If ($_.Contains("HDOSP") -eq "True") {$FEATHDSOP = "Enabled"} Else {$FEATHDOSP = "Disabled"}
        If ($_.Contains("CF") -eq "True") {$FEATCF = "Enabled"} Else {$FEATCF = "Disabled"}
        If ($_.Contains("IC") -eq "True") {$FEATIC = "Enabled"} Else {$FEATIC = "Disabled"}
        If ($_.Contains("SSLVPN") -eq "True") {$FEATSSLVPN = "Enabled"} Else {$FEATSSLVPN = "Disabled"}
        If ($_.Contains("AAA") -eq "True") {$FEATAAA = "Enabled"} Else {$FEATAAA = "Disabled"}
        If ($_.Contains("OSPF") -eq "True") {$FEATOSPF = "Enabled"} Else {$FEATOSPF = "Disabled"}
        If ($_.Contains("RIP") -eq "True") {$FEATRIP = "Enabled"} Else {$FEATRIP = "Disabled"}
        If ($_.Contains("BGP") -eq "True") {$FEATBGP = "Enabled"} Else {$FEATBGP = "Disabled"}
        If ($_.Contains("REWRITE") -eq "True") {$FEATREWRITE = "Enabled"} Else {$FEATREWRITE = "Disabled"}
        If ($_.Contains("IPv6PT") -eq "True") {$FEATIPv6PT = "Enabled"} Else {$FEATIPv6PT = "Disabled"}
        If ($_.Contains("AppFw") -eq "True") {$FEATAppFw = "Enabled"} Else {$FEATAppFw = "Disabled"}
        If ($_.Contains("RESPONDER") -eq "True") {$FEATRESPONDER = "Enabled"} Else {$FEATRESPONDER = "Disabled"}
        If ($_.Contains("HTMLInjection") -eq "True") {$FEATHTMLInjection = "Enabled"} Else {$FEATHTMLInjection = "Disabled"}
        If ($_.Contains("push") -eq "True") {$FEATpush = "Enabled"} Else {$FEATpush = "Disabled"}
        If ($_.Contains("AppFlow") -eq "True") {$FEATAppFlow = "Enabled"} Else {$FEATAppFlow = "Disabled"}
        If ($_.Contains("CloudBridge") -eq "True") {$FEATCloudBridge = "Enabled"} Else {$FEATCloudBridge = "Disabled"}
        If ($_.Contains("ISIS") -eq "True") {$FEATISIS = "Enabled"} Else {$FEATISIS = "Disabled"}
        If ($_.Contains("CH") -eq "True") {$FEATCH = "Enabled"} Else {$FEATCH = "Disabled"}
        If ($_.Contains("AppQoE") -eq "True") {$FEATAppQoE = "Enabled"} Else {$FEATAppQoE = "Disabled"}
        If ($_.Contains("Vpath") -eq "True") {$FEATVpath = "Enabled"} Else {$FEATVpath = "Disabled"}
        }
    }
#endregion NetScaler feature state

#region NetScaler Version

## Get version and build
$File | foreach { 
   if ($_ -like '#NS*') {
      $Y = ($_ -replace '#NS', '').split()
      $Version = $($Y[0]) 
      $Build = $($Y[2])
    }
}  

## Set script test version
## WIP THIS WORKS ONLY WHEN REGIONAL SETTINGS DIGIT IS SET TO . :)
$ScriptVersion = 10.1
#endregion NetScaler Version

#region NetScaler System Information

#region Basics
WriteWordLine 2 0 "NetScaler Basic"

$SETNS | Foreach {
    if ($_ -like 'set ns hostname *') {
        $NSHOSTNAMEPropertyArray = Get-StringPropertySplit -SearchString ($_ -Replace 'set ns' ,'') -RemoveQuotes;
        $NSHOSTNAME = $NSHOSTNAMEPropertyArray[1];
        }
    }

$SAVEDDATESTRING = Get-StringWithProperty -SearchString $File -Like '# Last *';
$SAVEDDATESTRING | foreach { 
    $SAVEDDATE = ($_ -Replace '# Last modified by `save config`,' ,'') ;
    }

$Params = $null
$Params = @{
    Hashtable = @{
        Name = $NSHOSTNAME
        Version = $Version;
        Build = $Build;
        Saveddate = $SAVEDDATE
    }
    Columns = "Name","Version","Build","Saveddate";
    Headers = "Host Name","Version","Build","Last Configuration Saved Date";
    Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
    AutoFit = $wdAutoFitContent;
}
$Table = AddWordTable @Params -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
#endregion Basics

#region NetScaler IP
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler IP"

WriteWordLine 2 0 "NetScaler Management IP Address"

$SetNS | foreach {
   if ($_ -like 'set ns config -IPAddress *') {
        $Params = $null
        $NSIP = Get-StringProperty $_ "-IPAddress";
        $Params = @{
            Hashtable = @{
                NSIP = Get-StringProperty $_ "-IPAddress";
                Subnet = Get-StringProperty $_ "-netmask";
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

#region NetScaler Global HTTP Parameters
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Global HTTP Parameters"

WriteWordLine 2 0 "NetScaler Global HTTP Parameters"
$SetNS | foreach {
    if ($_ -like 'set ns param *') {
        $IP = Get-StringProperty $_ "-cookieversion" "0";
        } else { $IP = "0" }
    if ($_ -like 'set ns httpParam *') {
        $DROP = Test-StringPropertyOnOff $_ "-dropInvalReqs";
        } else { $DROP = "On" }
    }

$Params = $null
$Params = @{
    Hashtable = @{
        CookieVersion = $IP;
        Drop = $DROP;
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

$SETNS | foreach {
   if ($_ -like 'set ns tcpParam *') {
        $TCP = Test-StringPropertyEnabledDisabled $_ "-WS";
        $SACK = Test-StringPropertyEnabledDisabled $_ "-SACK";
        $NAGLE = Test-StringPropertyEnabledDisabled $_ "-nagle";
    } else {
        $TCP = "Disabled";
        $SACK = "Disabled";
        $NAGLE = "Disabled";
    }
}

$Params = $null
$Params = @{
    Hashtable = @{
        TCP = $TCP;
        SACK = $SACK;
        NAGLE = $NAGLE;
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
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Global Diameter Parameter"

WriteWordLine 2 0 "NetScaler Global Diameter Parameters"
$SetNS | foreach {
   if ($_ -like 'set ns diameter *') {
        $Params = $null
        $Params = @{
            Hashtable = @{
                HOST = Get-StringProperty $_ "-identity" "NA";
                Realm = Get-StringProperty $_ "-realm" "NA";
                Close = Get-StringProperty $_ "-serverClosePropagation" "No";
            }
            Columns = "HOST","Realm","Close";
            Headers = "Host Identity","Realm","Server Close Propagation";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;

        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
        }
    }

#endregion NetScaler Global Diameter Parameters

#region NetScaler Time Zone
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Time zone"
WriteWordLine 2 0 "NetScaler Time Zone"

$Setns | foreach {  
    if ($_ -like 'set ns param *') {
        $TIMEZONE = Get-StringProperty $_ "-timezone" "Coordinated Universal Time" -RemoveQuotes;
        } else {$TIMEZONE = "Coordinated Universal Time"; }
    }

$Params = $null
$Params = @{
    Hashtable = @{
        TimeZone = $TIMEZONE;
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

#region NetScaler Management vLAN
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler management vLAN"

WriteWordLine 2 0 "NetScaler Management vLAN"

$SETNSDIAMETER = Get-StringWithProperty -SearchString $SetNS -Like 'set ns config -nsvlan *';
if($SETNSDIAMETER.Length -le 0) { WriteWordLine 0 0 "No Management vLAN has been configured"} else {
    $SetNS | foreach {
       if ($_ -like 'set ns config -nsvlan *') {
            $Params = $null
            $Params = @{
                Hashtable = @{
                    ## IB - This table will only have 1 row so create the nested hashtable inline
                    ID = Get-StringProperty $_ "-nsvlan";
                    INTERFACE = Get-StringProperty $_ "-ifnum";
                    Tagged = Get-StringProperty $_ "-tagged";
                }
                Columns = "ID","INTERFACE","Tagged";
                Headers = "vLAN ID","Interface","Tagged";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
            }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            }
        }
    }
WriteWordLine 0 0 " "
#endregion NetScaler Management vLAN

#region NetScaler High Availability
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters High Availability"

WriteWordLine 2 0 "NetScaler High Availability"

$ADDHANODE = Get-StringWithProperty -SearchString $Add -Like 'add HA node *';
if($ADDHANODE.Length -le 0) { WriteWordLine 0 0 "High Availability has not been configured"} else {
    
    $NSRPCH = $null
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $NSRPCH = @();
    
    $NSRPCH += @{
        NSNODE = $NSIP; #NSIP Variable set in chapter NetScaler IP Address
    }

    $ADDHANODE | foreach {
        $HAPropertyArray = Get-StringPropertySplit -SearchString ($_ -Replace 'add HA node' ,'') -RemoveQuotes;
        $NSRPCH += @{
            NSNODE = $HAPropertyArray[1];
        }
    }

    if ($NSRPCH.Length -gt 1) {
        $Params = $null
        $Params = @{
            Hashtable = $NSRPCH;
            Columns = "NSNODE";
            Headers = "NetScaler Node";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
            }
        $Table = AddWordTable @Params -NoGridLines;

        FindWordDocumentEnd;
    } else {WriteWordLine 0 0 "High Availability has not been configured"}
}
WriteWordLine 0 0 " "

#endregion NetScaler High Availability

#region NetScaler Administration
$selection.InsertNewPage()
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Administration"
WriteWordLine 2 0 "NetScaler System Authentication"
WriteWordLine 0 0 " "

#region Local Administration Users
WriteWordLine 3 0 "NetScaler System Users"

$AUTHLOCH = $null    
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $AUTHLOCH = @();

$AUTHLOCH += @{
    LocalUser = "nsroot";
    }

foreach ($AUTHLOC in $AUTHLOCS) {
    ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
    $AUTHLOCPropertyArray = Get-StringPropertySplit -SearchString ($AUTHLOC -Replace 'add system user' ,'') -RemoveQuotes;

    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be able 400 characters wide!
    $AUTHLOCH += @{
            LocalUser = $AUTHLOCPropertyArray[0];
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

if($AUTHGRPS.Length -le 0) { WriteWordLine 0 0 "No Local Group has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $AUTHGRPH = @();

    foreach ($AUTHGRP in $AUTHGRPS) {
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $AUTHGRPPropertyArray = Get-StringPropertySplit -SearchString ($AUTHGRP -Replace 'add system' ,'') -RemoveQuotes;
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $AUTHGRPH += @{
                LocalUser = $AUTHGRPPropertyArray[1];
            }
        }

        if ($AUTHGRPH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $AUTHGRPH;
                Columns = "LocalUser";
                Headers = "Local User";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            }
        }
WriteWordLine 0 0 " "
#endregion Authentication Local Administration Groups

#endregion NetScaler Administration

#region NetScaler Admin Partitions
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Admin Partitions"

WriteWordLine 2 0 "NetScaler Admin Partitions"

if($AdminPartitions.Length -le 0) { WriteWordLine 0 0 "No Admin Partitions have been configured"} else {
       
        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $APH = @();
        
        foreach ($AdminPartition in $AdminPartitions) {
            ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
            $AdminPartitionPropertyArray = Get-StringPropertySplit -SearchString ($AdminPartition -Replace 'add ns partition' ,'') -RemoveQuotes;
            $AdminPartitionDisplayNameWithQoutes = Get-StringProperty $AdminPartition "partition";
            
            ## IB - Create parameters for the hashtable so that we can splat them otherwise the
            ## IB - command will be able 400 characters wide!

            $APBindMatches = Get-StringWithProperty -SearchString $AdminPartitionsBind -Like "bind ns partition $AdminPartitionDisplayNameWithQoutes *";

            $APBindMatches | foreach {
                $vlanap = Get-StringProperty $_ "-vlan";
                }

            $APH += @{
                ID = Get-StringProperty $AdminPartition "-partitionid";
                APNAME = $AdminPartitionPropertyArray[0];
                vLAN = $vlanap;
                MinBand = Get-StringProperty $AdminPartition "-minBandwidth" "10240";
                MaxBand = Get-StringProperty $AdminPartition "-maxBandwidth" "10240";
                Maxconn = Get-StringProperty $AdminPartition "-maxConn" "1024";
                Maxmem = Get-StringProperty $AdminPartition "-maxMemLimit" "10";
                }
            }

            if ($APH.Length -gt 0) {
                $Params = $null
                $Params = @{
                    Hashtable = $APH;
                    Columns = "ID","APNAME","vLAN","MinBand","MaxBand","Maxconn","Maxmem";
                    Headers = "ID","Name","vLAN","Minimum Bandwidth","Maximum Bandwidth","Maximum Connections","Maximum Memory";
                    Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                    AutoFit = $wdAutoFitContent;
                    }
                $Table = AddWordTable @Params -NoGridLines;
                FindWordDocumentEnd;
                WriteWordLine 0 0 " "
                }
            }
    
#endregion NetScaler Admin Partitions

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

If ($Version -gt $ScriptVersion) {
    WriteWordLine 0 0 ""
    WriteWordLine 0 0 "Warning: You are using Citrix NetScaler version $Version, modes added since version $ScriptVersion will not be shown."
    WriteWordLine 0 0 ""
    }

$Enable | foreach {  
    if ($_ -like 'enable ns mode *') {
        If ($_.Contains("FR") -eq "True") {$FR = "Enabled"} Else {$FR = "Disabled"}
        If ($_.Contains("L2") -eq "True") {$L2 = "Enabled"} Else {$L2 = "Disabled"}
        If ($_.Contains("USIP") -eq "True") {$USIP = "Enabled"} Else {$USIP = "Disabled"}
        If ($_.Contains("CKA") -eq "True") {$CKA = "Enabled"} Else {$CKA = "Disabled"}
        If ($_.Contains("TCPB") -eq "True") {$TCPB = "Enabled"} Else {$TCPB = "Disabled"}
        If ($_.Contains("MBF") -eq "True") {$MBF = "Enabled"} Else {$MBF = "Disabled"}
        If ($_.Contains("Edge") -eq "True") {$Edge = "Enabled"} Else {$Edge = "Disabled"}
        If ($_.Contains("USNIP") -eq "True") {$USNIP = "Enabled"} Else {$USNIP = "Disabled"}
        If ($_.Contains("PMTUD") -eq "True") {$PMTUD = "Enabled"} Else {$PMTUD = "Disabled"}
        If ($_.Contains("SRADV") -eq "True") {$SRADV = "Enabled"} Else {$SRADV = "Disabled"}
        If ($_.Contains("DRADV") -eq "True") {$DRADV = "Enabled"} Else {$DRADV = "Disabled"}
        If ($_.Contains("IRADV") -eq "True") {$IRADV = "Enabled"} Else {$IRADV = "Disabled"}
        If ($_.Contains("SRADV6") -eq "True") {$SRADV6 = "Enabled"} Else {$SRADV6 = "Disabled"}
        If ($_.Contains("DRADV6") -eq "True") {$DRADV6 = "Enabled"} Else {$DRADV6 = "Disabled"}
        If ($_.Contains("BridgeBPDUs") -eq "True") {$BridgeBPDUs = "Enabled"} Else {$BridgeBPDUs = "Disabled"}
        If ($_.Contains("L3") -eq "True") {$L3 = "Enabled"} Else {$L3 = "Disabled"}

        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $ADVModes = @(
            @{ Description = "Mode"; Value = "State"}  
            @{ Description = "Fast Ramp"; Value = $FR}        
            @{ Description = "Layer 2 mode"; Value = $L2}        
            @{ Description = "Use Source IP"; Value = $USIP}        
            @{ Description = "Client SideKeep-alive"; Value = $CKA}        
            @{ Description = "TCP Buffering"; Value = $TCPB}        
            @{ Description = "MAC-based forwarding"; Value = $MBF}
            @{ Description = "Edge configuration"; Value = $Edge}        
            @{ Description = "Use Subnet IP"; Value = $USNIP}        
            @{ Description = "Use Layer 3 Mode"; Value = $L3}        
            @{ Description = "Path MTU Discovery"; Value = $PMTUD}        
            @{ Description = "Static Route Advertisement"; Value = $SRADV}        
            @{ Description = "Direct Route Advertisement"; Value = $DRADV}        
            @{ Description = "Intranet Route Advertisement"; Value = $IRADV}        
            @{ Description = "Ipv6 Static Route Advertisement"; Value = $SRADV6}        
            @{ Description = "Ipv6 Direct Route Advertisement"; Value = $DRADV6}        
            @{ Description = "Bridge BPDUs" ; Value = $BridgeBPDUs}        
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
        }
    }

$selection.InsertNewPage()

#endregion NetScaler Modes

#region NetScaler Web Interface
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Web Interface"

WriteWordLine 1 0 "NetScaler Web Interface"

$WI = Get-StringWithProperty -SearchString $File -Like 'install wi package *';

if($WI.Length -le 0) { WriteWordLine 0 0 "Citrix Web Interface has not been installed"} else { WriteWordLine 0 0 "Citrix Web Interface has been installed"}

$selection.InsertNewPage()

#endregion NetScaler Web Interface

#endregion NetScaler System Information

#region NetScaler Monitoring
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Monitoring"

WriteWordLine 1 0 "NetScaler Monitoring"

WriteWordLine 2 0 "SNMP Community"

$SNMPCOMS = Get-StringWithProperty -SearchString $Add -Like 'add snmp community *';

if($SNMPCOMS.Length -le 0) { WriteWordLine 0 0 "No SNMP Community has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $SNMPCOMH = @();

    foreach ($SNMPCOM in $SNMPCOMS) {
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $SNMPCOMPropertyArray = Get-StringPropertySplit -SearchString ($SNMPCOM -Replace 'add snmp community' ,'') -RemoveQuotes;

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $SNMPCOMH += @{
                SNMPCommunity = $SNMPCOMPropertyArray[0];
                Permission = $SNMPCOMPropertyArray[1];
            }
        }
        if ($SNMPCOMH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $SNMPCOMH;
                Columns = "SNMPCommunity","Permission";
                Headers = "SNMP Community","Permission";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
        }
    }
WriteWordLine 0 0 " "

WriteWordLine 2 0 "SNMP Manager"

$SNMPMANS = Get-StringWithProperty -SearchString $Add -Like 'add snmp manager *';

if($SNMPMANS.Length -le 0) { WriteWordLine 0 0 "No SNMP Manager has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $SNMPMANSH = @();

    foreach ($SNMPMAN in $SNMPMANS) {

        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $SNMPMANPropertyArray = Get-StringPropertySplit -SearchString ($SNMPMAN -Replace 'add snmp manager' ,'') -RemoveQuotes;

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $SNMPMANSH += @{
                SNMPManager = $SNMPMANPropertyArray[0];
                Netmask = Get-StringProperty $SNMPMAN "-netmask" "255.255.255.255";
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

$SNMPALERTS = Get-StringWithProperty -SearchString $Set -Like 'set snmp alarm *';

if($SNMPALERTS.Length -le 0) { WriteWordLine 0 0 "No SNMP Alert has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $SNMPALERTSH = @();

    foreach ($SNMPALERT in $SNMPALERTS) {
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $SNMPALERTPropertyArray = Get-StringPropertySplit -SearchString ($SNMPALERT -Replace 'set snmp alarm' ,'') -RemoveQuotes;

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!

        $SNMPALERTSH += @{
                Alarm = $SNMPALERTPropertyArray[0];
                State = Test-NotStringPropertyEnabledDisabled $SNMPALERT "-state";
                Time = Get-StringProperty $SNMPALERT "-time" "0";
                TimeOut = Get-StringProperty $SNMPALERT "-timeout" "NA";
            }
        }
        if ($SNMPALERTSH.Length -gt 0) {
            $Params = @{
                Hashtable = $SNMPALERTSH;
                Columns = "Alarm","State","Time","TimeOut";
                Headers = "NetScaler Alarm","State","Time","Time-Out";
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

#region networking

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Network Configuration"

WriteWordLine 1 0 "NetScaler Networking"

#region NetScaler IP addresses

WriteWordLine 2 0 "NetScaler IP addresses"
if($IPLIST.Length -le 0) { WriteWordLine 0 0 "No IP Address has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $IPADDRESSH = @();

    foreach ($IP in $IPLIST) {
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $IPPropertyArray = Get-StringPropertySplit -SearchString ($IP -Replace 'add ns ip ' ,'') -RemoveQuotes;
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!

        $IPADDRESSH += @{
                IPAddress = $IPPropertyArray[0];
                SubnetMask = $IPPropertyArray[1];
                TrafficDomain = Get-StringProperty $IP "-td" "0";
                Management = Test-StringPropertyEnabledDisabled $IP "-mgmtAccess";
                vServer = Test-NotStringPropertyEnabledDisabled $IP "-vServer";
                GUI = Test-NotStringPropertyEnabledDisabled $IP "-gui";
                SNMP = Test-NotStringPropertyEnabledDisabled $IP "-snmp";
                Telnet = Test-NotStringPropertyEnabledDisabled $IP "-telnet";
            }
        }

        if ($IPADDRESSH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $IPADDRESSH;
                Columns = "IPAddress","SubnetMask","TrafficDomain","Management","vServer","GUI","SNMP","Telnet";
                Headers = "IP Address","Subnet Mask","Traffic Domain","Management","vServer","GUI","SNMP","Telnet";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
        }

#endregion NetScaler IP addresses

#region NetScaler Interfaces

WriteWordLine 2 0 "NetScaler Interfaces"

if($NICS.Length -le 0) { WriteWordLine 0 0 "No network interface has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $NICH = @();

    foreach ($NIC in $NICS) {
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        ## IB - Store the Load Balancer as native and escaped. This is needed by LBBIND etc.
        
        $NICDisplayName = Get-StringProperty $NIC "interface" -RemoveQuotes;
        
        $NICH += @{
            InterfaceID = $NICDisplayName;
            InterfaceType = Get-StringProperty $NIC "-intftype" -RemoveQuotes;
            HAMonitoring = Test-NotStringPropertyOnOff $NIC "-haMonitor";
            State = Test-NotStringPropertyOnOff $NIC "-state";
            AutoNegotiate = Test-NotStringPropertyOnOff $NIC "-autoneg";
            Tag = Test-StringPropertyOnOff $NIC "-tagall";
            }
        }

        if ($NICH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $NICH;
                Columns = "InterfaceID","InterfaceType","HAMonitoring","State","AutoNegotiate","Tag";
                Headers = "Interface ID","Interface Type","HA Monitoring","State","Auto Negotiate","Tag All vLAN";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
        }

#endregion NetScaler Interfaces

#region NetScaler vLAN

WriteWordLine 2 0 "NetScaler vLANs"

if($VLANS.Length -le 0) { WriteWordLine 0 0 "No vLAN has been configured"} else {
    $vLANH = $null
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $vLANH = @();

    foreach ($VLAN in $VLANS) {
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!

        $VLANDisplayName = Get-StringProperty $VLAN "vlan" -RemoveQuotes;
        $VLANBINDMatches = $null 
        $VLANBIND = $null
        
        $VLANBINDMatches = Get-StringWithProperty -SearchString $VLANSBIND -Like "bind vlan $VLANDisplayName *";

        foreach ($VLANBIND in $VLANBINDMatches) {
            $INT1 = Get-StringProperty $VLANBIND "-ifnum";
            }

        $vLANH += @{
            vLANID = $VLANDisplayName;
            Interface1 = $INT1;
            }
        }

        if ($vLANH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $vLANH;
                Columns = "vLANID","Interface1","Interface2","Interface3","Interface4","Interface5";
                Headers = "vLAN ID","Interface","Interface","Interface","Interface","Interface";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
        }

#endregion NetScaler vLAN

#region NetScaler Network Channel
$selection.InsertNewPage()
WriteWordLine 2 0 "NetScaler Network Channel"

if($CHANNELS.Length -le 0) { WriteWordLine 0 0 "No network channel has been configured"} else {
    
    foreach ($CHANNEL in $CHANNELS) {  
        $CHANNELDisplayName = Get-StringProperty $CHANNEL "channel" -RemoveQuotes;
        
        WriteWordLine 3 0 "Network Channel $CHANNELDisplayName"
        $Params = $null
        $Params = @{
            Hashtable = @{
            CHANNEL = $CHANNELDisplayName;
            Alias = Get-StringProperty $CHANNEL "-ifalias" "Not Configured";
            HA = Test-NotStringPropertyOnOff $CHANNEL "haMonitor";
            State = Get-StringProperty $CHANNEL "-state" "Enabled";
            Speed = Get-StringProperty $CHANNEL "-speed" "Auto";            Tagall = Test-StringPropertyOnOff $CHANNEL "-Tagall";
            MTU = Get-StringProperty $CHANNEL "-mtu" "1500";
            }
        Columns = "CHANNEL","Alias","HA","State","Speed","Tagall";
        Headers = "Channel","Alias","HA Monitoring","State","Speed","Tag all vLAN";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
    
        $NICH = $NULL
        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $NICH = @();

        foreach ($NIC in $NICS) {
            $CHANNELNIC = Get-StringProperty $NIC "-ifnum";
            If ($CHANNELNIC -eq $CHANNELDisplayName) {
                ## IB - Create parameters for the hashtable so that we can splat them otherwise the
                ## IB - command will be able 400 characters wide!
                $NICDisplayName = Get-StringProperty $NIC "interface" -RemoveQuotes;
                $NICH += @{
                    InterfaceID = $NICDisplayName;
                    }
                }
            }

        if ($NICH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $NICH;
                Columns = "InterfaceID";
                Headers = "Interface ID";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
        }

        if($VLANS.Length -le 0) { WriteWordLine 0 0 "No vLAN has been configured for this Network Channel"} else {
            $vLANH = $null
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $vLANH = @();

            foreach ($VLAN in $VLANSBIND) {
                $CHANNELVLAN = $null
                $CHANNELVLAN = Get-StringProperty $VLAN "-ifnum";
                If ($CHANNELVLAN -eq $CHANNELDisplayName) {      
                    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
                    ## IB - command will be able 400 characters wide!

                    $VLANDisplayName = Get-StringProperty $VLAN "vlan" -RemoveQuotes;
                    $VLANH += @{
                        VLANID = $VLANDisplayName;
                        }
                    }
                }

            if ($VLANH.Length -gt 0) {
                $Params = $null
                $Params = @{
                    Hashtable = $VLANH;
                    Columns = "VLANID";
                    Headers = "VLAN ID";
                    Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                    AutoFit = $wdAutoFitContent;
                    }
                $Table = AddWordTable @Params -NoGridLines;
                FindWordDocumentEnd;
                WriteWordLine 0 0 " "
                } else { WriteWordLine 0 0 "No vLAN has been configured for this Network Channel"}
            }
        }
    }

WriteWordLine 0 0 " "
$selection.InsertNewPage()
#endregion NetScaler Network Channel

#region routing table

WriteWordLine 2 0 "NetScaler Routing Table"

WriteWordLine 0 0 "The NetScaler documentation script only documents manually added route table entries."
WriteWordLine 0 0 " "

if($ROUTES.Length -le 0) { WriteWordLine 0 0 "No route has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $ROUTESH = @();

    foreach ($ROUTE in $ROUTES) {
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $ROUTEPropertyArray = Get-StringPropertySplit -SearchString ($ROUTE -Replace 'add route ' ,'') -RemoveQuotes;
        $ROUTESH += @{
            Network = $ROUTEPropertyArray[0];
            Subnet = $ROUTEPropertyArray[1];
            Gateway = $ROUTEPropertyArray[2];
            Distance = Get-StringProperty $ROUTE "-distance" "0" -RemoveQuotes;
            Weight = Get-StringProperty $ROUTE "-weight" "1" -RemoveQuotes;
            Cost = Get-StringProperty $ROUTE "-cost" "0" -RemoveQuotes;
            }
        }

        if ($ROUTESH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $ROUTESH;
                Columns = "Network","Subnet","Gateway","Distance","Weight","Cost";
                Headers = "Network","Subnet","Gateway","Distance","Weight","Cost";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
        }

#endregion routing table

#region NetScaler DNS Configuration
$selection.InsertNewPage()
WriteWordLine 2 0 "NetScaler DNS Configuration"

#region dns records

WriteWordLine 3 0 "NetScaler DNS Name Servers"
if($DNSNAMESERVERS.Length -le 0) { WriteWordLine 0 0 "No DNS Name Server has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $DNSNAMESERVERH = @();

    foreach ($DNSNAMESERVER in $DNSNAMESERVERS) {
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $DNSSERVERDisplayName = $null
        $DNSSERVERDisplayName = Get-StringProperty $DNSNAMESERVER "nameServer" -RemoveQuotes;
        $DNSNAMESERVERH += @{
            DNSServer = $DNSSERVERDisplayName;
            State = Get-StringProperty $DNSNAMESERVER "-state" "Enabled" -RemoveQuotes;
            Prot = Get-StringProperty $DNSNAMESERVER "-type" "UDP" -RemoveQuotes;
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
      
#endregion dns records

#region DNS Address Records

WriteWordLine 3 0 "NetScaler DNS Address Records"

if($DNSRECORDCONFIGS.Length -le 0) { WriteWordLine 0 0 "No DNS Address Record has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $DNSRECORDCONFIGH = @();

    foreach ($DNSRECORDCONFIG in $DNSRECORDCONFIGS) {
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $DNSRECORDCONFIGPropertyArray = Get-StringPropertySplit -SearchString ($DNSRECORDCONFIG -Replace 'add dns addRec ' ,'') -RemoveQuotes;
        $DNSRECORDCONFIGH += @{
            DNSRecord = $DNSRECORDCONFIGPropertyArray[0];
            IPAddress = $DNSRECORDCONFIGPropertyArray[1];
            TTL = Get-StringProperty $DNSRECORDCONFIG "-TTL";
            }
        }

        if ($DNSRECORDCONFIGH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $DNSRECORDCONFIGH;
                Columns = "DNSRecord","IPAddress","TTL";
                Headers = "DNS Record","IP Address","TTL";
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
WriteWordLine 2 0 "NetScaler ACL Configuration"

#region NetScaler Simple ACL IPv4

WriteWordLine 3 0 "NetScaler Simple ACL IPv4"

if($SIMPLEACLS.Length -le 0) { WriteWordLine 0 0 "No Simple ACL has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $SIMPLEACLSH = @();

    foreach ($SIMPLEACL in $SIMPLEACLS) {
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $SIMPLEACLPropertyArray = Get-StringPropertySplit -SearchString ($SIMPLEACL -Replace 'add ns simpleacl' ,'') -RemoveQuotes;

        $SIMPLEACLSH += @{
            ACLNAME = $SIMPLEACLPropertyArray[0];
            ACTION = $SIMPLEACLPropertyArray[1];
            SOURCEIP = Get-StringProperty $SIMPLEACL "-srcIP";
            DESTPORT = Get-StringProperty $SIMPLEACL "-destPort";
            PROT = Get-StringProperty $SIMPLEACL "-protocol";            }
        }

        if ($SIMPLEACLSH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $SIMPLEACLSH;
                Columns = "ACLNAME","ACTION","SOURCEIP","DESTPORT","PROT";
                Headers = "ACL Name","Action","Source IP","Destination Port","Protocol";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
        }

#endregion NetScaler Simple ACL IPv4

$selection.InsertNewPage()
#endregion NetScaler ACL

#endregion networking

#region NetScaler Traffic Domains
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Traffic Domains"

WriteWordLine 1 0 "NetScaler Traffic Domains"
##No function yet for routing table per TD

$TD = Get-StringWithProperty -SearchString $Add -Like 'add ns trafficDomain *';

if($TD.Length -le 0) { WriteWordLine 0 0 "No Traffic Domains have been configured"} else {
    WriteWordLine 0 0 "Only documents one assigned vLAN just like VLAN and Interface WIP"
    $TD | foreach {
        $Bind | foreach {  
            if ($_ -like 'bind ns trafficDomain *') {
                $vLAN = Get-StringProperty $_ "-vlan"
                }
            }
       
        $TDPropertyArray = Get-StringPropertySplit -SearchString ($_ -Replace 'add ns trafficDomain' ,'') -RemoveQuotes;        
        ## IB - Store the Load Balancer as native and escaped. This is needed by LBBIND etc.
        $TDDisplayName = Get-StringProperty $_ "trafficDomain" -RemoveQuotes;
        $TDName = $TDDisplayName.Trim();
        
        WriteWordLine 2 0 "Traffic Domain $TDDisplayName"
        WriteWordLine 0 0 " "
        Write-Verbose "$(Get-Date): `tTraffic Domain $TDDisplayName"

        $TDDisplayName = Get-StringProperty $_ "trafficdomain" -RemoveQuotes;
        $TDName = $TDDisplayName.Trim();

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $Params = $null
        $Params = @{
            Hashtable = @{
                ## IB - This table will only have 1 row so create the nested hashtable inline
                ID = $TDPropertyArray[0];
                Alias = Get-StringProperty $_ "-aliasName";
                vLAN = $vLAN;
            }
            Columns = "ID","Alias","vLAN";
            Headers = "Traffic Domain ID","Traffic Domain Alias","Traffic Domain vLAN";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;
        #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

        FindWordDocumentEnd;

        WriteWordLine 0 0 " "
        
        WriteWordLine 4 0 "Content Switch"        
    
        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $CSTDH = @();

        $ContentSwitches | foreach {
            if ((Get-StringProperty $_ "-td") -eq $TDPropertyArray[0]) {
                $CSTDPropertyArray = Get-StringPropertySplit -SearchString ($_ -Replace 'add cs vserver' ,'') -RemoveQuotes;
                $CSTDH = @{
                    ContentSwitch = $CSTDPropertyArray[0]
                }
            }
        }

        if ($CSTDH.Length -gt 0) {
            $Params = @{
                Hashtable = $CSTDH;
                Columns = "ContentSwitch";
                Headers = "Content Switch";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
        } else {
                WriteWordLine 0 0 "No Content Switch been configured for this Traffic Domain"
            } # end if        
        
        WriteWordLine 4 0 "Load Balancer"      
  
        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $LBTDH = @();

        $LoadBalancers | foreach {
            if ((Get-StringProperty $_ "-td") -eq $TDPropertyArray[0]) {
                $LBTDPropertyArray = Get-StringPropertySplit -SearchString ($_ -Replace 'add lb vserver' ,'') -RemoveQuotes;
                $LBTDH += @{
                    LBTD = $LBTDPropertyArray[0];
                }
            }
        }
        
        if ($LBTDH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $LBTDH;
                Columns = "LBTD";
                Headers = "Load Balancer";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
        } else {
                WriteWordLine 0 0 "No Load Balancer been configured for this Traffic Domain"
            } # end if

        WriteWordLine 4 0 "Services"      
    
        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $SVCTDH = @();

        $Services | foreach {
            if ((Get-StringProperty $_ "-td") -eq $TDPropertyArray[0]) {
                $SVCTDPropertyArray = Get-StringPropertySplit -SearchString ($_ -Replace 'add service' ,'') -RemoveQuotes;
                $SVCTDH += @{
                    SVCTD = $SVCTDPropertyArray[0];
                }
            }
        }
        
        if ($SVCTDH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $SVCTDH;
                Columns = "SVCTD";
                Headers = "Service";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
        } else {
                WriteWordLine 0 0 "No Service has been configured for this Traffic Domain"
            } # end if
  
        WriteWordLine 4 0 "Servers"      
  
        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $SVRTDH = @();

        $Servers | foreach {
            if ((Get-StringProperty $_ "-td") -eq $TDPropertyArray[0]) {
                $SVRTDPropertyArray = Get-StringPropertySplit -SearchString ($_ -Replace 'add server' ,'') -RemoveQuotes;
                $SVRTDH += @{
                    SVRTD = $SVRTDPropertyArray[0];
                }
            }
        }
        
        if ($SVRTDH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $SVRTDH;
                Columns = "SVRTD";
                Headers = "Server";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
        } else {
                WriteWordLine 0 0 "No Server has been configured for this Traffic Domain"
            } # end if        

        $selection.InsertNewPage()
        }
    }
    
#endregion NetScaler Traffic Domains

#region NetScaler Authentication
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Authentication"

$selection.InsertNewPage()

WriteWordLine 1 0 "NetScaler Authentication"

#region Local Users
WriteWordLine 2 0 "NetScaler Local Users"
if($AUTHLOCS.Length -le 0) { WriteWordLine 0 0 "No Local User has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $AUTHLOCUSRH = @();

    foreach ($AUTHLOCUSER in $AUTHLOCUSERS) {
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $AUTHLOCUUSERPropertyArray = Get-StringPropertySplit -SearchString ($AUTHLOCUSER -Replace 'add authentication localPolicy' ,'') -RemoveQuotes;

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $AUTHLOCUSRH += @{
                LocalUser = $AUTHLOCUUSERPropertyArray[0];
                Expression = $AUTHLOCUUSERPropertyArray[1];
            }
        }
        if ($AUTHLOCUSRH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $AUTHLOCUSRH;
                Columns = "LocalUser","Expression";
                Headers = "Local User","Expression";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            }
        }
WriteWordLine 0 0 " "
#endregion Authentication Local Users

#region Authentication LDAP Policies
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler LDAP Authentication"
WriteWordLine 2 0 "NetScaler LDAP Policies"

if($AUTHLDAPPOLS.Length -le 0) { WriteWordLine 0 0 "No LDAP Policy has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $AUTHLDAPPOLH = @();

    foreach ($AUTHLDAPPOL in $AUTHLDAPPOLS) {
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $AUTHLDAPPOLPropertyArray = Get-StringPropertySplit -SearchString ($AUTHLDAPPOL -Replace 'add authentication ldapPolicy' ,'') -RemoveQuotes;
                
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $AUTHLDAPPOLH += @{
                Policy = $AUTHLDAPPOLPropertyArray[0];
                Expression = $AUTHLDAPPOLPropertyArray[1];
                Action = $AUTHLDAPPOLPropertyArray[2];
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
        }
WriteWordLine 0 0 " "

#endregion Authentication LDAP Policies

#region Authentication LDAP
WriteWordLine 2 0 "NetScaler LDAP authentication actions"

if($AUTHLDAPACTS.Length -le 0) { WriteWordLine 0 0 "No LDAP Authentication action has been configured"} else {
    $CurrentRowIndex = 0;
    foreach ($AUTHLDAP in $AUTHLDAPACTS) {
        
        $CurrentRowIndex++;
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $AUTHLDAPPropertyArray = Get-StringPropertySplit -SearchString ($AUTHLDAP -Replace 'add authentication ldapAction' ,'') -RemoveQuotes;
        
        ## IB - Store the Load Balancer as native and escaped. This is needed by LBBIND etc.
        $AUTHLDAPDisplayName = Get-StringProperty $AUTHLDAP "ldapAction" -RemoveQuotes;
        $AUTHLDAPName = $AUTHLDAPDisplayName.Trim();
    
        Write-Verbose "$(Get-Date): `tLDAP Authentication $CurrentRowIndex/$($AUTHLDAPS.Length) $AUTHLDAPDisplayName"     
        WriteWordLine 3 0 "LDAP Authentication action $AUTHLDAPDisplayName";

        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $LDAPCONFIG = @(
            @{ Description = "Description"; Value = "Configuration"; }
            @{ Description = "LDAP Server IP"; Value = Get-StringProperty $AUTHLDAP "-serverIP"; }
            @{ Description = "LDAP Server Port"; Value = Get-StringProperty $AUTHLDAP "-serverPort" "389"; }
            @{ Description = "LDAP Server Time-Out"; Value = Get-StringProperty $AUTHLDAP "-authTimeout" "3"; }
            @{ Description = "Validate Certificate"; Value = Test-StringPropertyYesNo $AUTHLDAP "-validateServerCert"; }
            @{ Description = "LDAP Base OU"; Value = Get-StringProperty $AUTHLDAP "-ldapbase" -RemoveQuotes; }
            @{ Description = "LDAP Bind DN"; Value = Get-StringProperty $AUTHLDAP "-ldapBindDn" -RemoveQuotes; }
            @{ Description = "Login Name"; Value = Get-StringProperty $AUTHLDAP "-ldapLoginName"; }
            @{ Description = "Sub Attribute Name"; Value = Get-StringProperty $AUTHLDAP "-subAttributeName"; }
            @{ Description = "Security Type"; Value = Get-StringProperty $AUTHLDAP "-secType" "Default Setting"; }
            @{ Description = "Password Changes"; Value = Get-StringProperty $AUTHLDAP "-passwdChange" "Default Setting"; }
            @{ Description = "Search Filter"; Value = Get-StringProperty $AUTHLDAP "-searchFilter" "Not Configured" -RemoveQuotes;}
            @{ Description = "Group attribute name"; Value = Get-StringProperty $AUTHLDAP "-groupAttrName"; }
            @{ Description = "LDAP Single Sign On Attribute"; Value = Get-StringProperty $AUTHLDAP "-ssoNameAttribute" "Not Configured"; }            @{ Description = "Authentication"; Value = Test-StringPropertyEnabledDisabled $AUTHLDAP "-authentication"; }            @{ Description = "User Required"; Value = Test-StringPropertyYesNo $AUTHLDAP "-requireUser"; }            @{ Description = "LDAP Referrals"; Value = Test-StringPropertyOnOff $AUTHLDAP "-followReferrals"; }
            @{ Description = "Nested Group Extraction"; Value = Test-NotStringPropertyOnOff $AUTHLDAP "-nestedGroupExtraction"; }
            @{ Description = "Maximum Nesting level"; Value = Get-StringProperty $AUTHLDAP "-maxNestingLevel" "Not Configured"; }
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
}
WriteWordLine 0 0 " "
#endregion Authentication LDAP

#region Authentication RADIUS Policies
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Radius Authentication"
WriteWordLine 2 0 "NetScaler Radius Policies"

if($AUTHRADPOLS.Length -le 0) { WriteWordLine 0 0 "No Radius Policy has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $AUTHRADPOLH = @();

    foreach ($AUTHRADPOL in $AUTHRADPOLS) {
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $AUTHRADPOLPropertyArray = Get-StringPropertySplit -SearchString ($AUTHRADPOL -Replace 'add authentication radiusPolicy' ,'') -RemoveQuotes;
      
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $AUTHRADPOLH += @{
                Policy = $AUTHRADPOLPropertyArray[0];
                Expression = $AUTHRADPOLPropertyArray[1];
                Action = $AUTHRADPOLPropertyArray[2];
            }
        }

        if ($AUTHRADPOLH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $AUTHRADPOLH;
                Columns = "Policy","Expression","Action";
                Headers = "Radius Policy","Expression","Radius Action";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            }
        }
WriteWordLine 0 0 " "
#endregion Authentication RADIUS Policies

#region Authentication RADIUS
WriteWordLine 2 0 "NetScaler RADIUS authentication action"

if($AUTHRADIUSS.Length -le 0) { WriteWordLine 0 0 "No RADIUS Authentication actions has been configured"} else {
    
    $CurrentRowIndex = 0;

    foreach ($AUTHRADIUS in $AUTHRADIUSS) {
        $CurrentRowIndex++;
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $AUTHRADIUSPropertyArray = Get-StringPropertySplit -SearchString ($AUTHRADIUS -Replace 'add authentication radiusAction' ,'') -RemoveQuotes;
        
        ## IB - Store the Load Balancer as native and escaped. This is needed by LBBIND etc.
        $AUTHRADIUSDisplayName = Get-StringProperty $AUTHRADIUS "radiusAction" -RemoveQuotes;
        $AUTHRADIUSName = $AUTHRADIUSDisplayName.Trim();

        Write-Verbose "$(Get-Date): `tRADIUS Authentication $CurrentRowIndex/$($AUTHRADIUSS.Length) $AUTHRADIUSDisplayName"     
        WriteWordLine 3 0 "RADIUS Authentication action $AUTHRADIUSDisplayName";
        
        $RADIUSCONFIG = $null
        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $RADIUSCONFIG = @(
            @{ Description = "Description"; Value = "Configuration"; }
            @{ Description = "RADIUS Server IP"; Value = Get-StringProperty $AUTHRADIUS "-serverIP"; }
            @{ Description = "RADIUS Server Port"; Value = Get-StringProperty $AUTHRADIUS "-serverPort" "1812"; }
            @{ Description = "RADIUS Server Time-Out"; Value = Get-StringProperty $AUTHRADIUS "-authTimeout" "3"; }
            @{ Description = "Radius NAS IP"; Value = Test-StringPropertyEnabledDisabled $AUTHRADIUS "-radNASip"; }
            @{ Description = "Radius NAS ID"; Value = Get-StringProperty $AUTHRADIUS "-radNASid" "Not Configured"; }
            @{ Description = "Radius Vendor ID"; Value = Get-StringProperty $AUTHRADIUS "-radVendorID" "Not Configured"; }
            @{ Description = "Radius Attribute Type"; Value = Get-StringProperty $AUTHRADIUS "-radAttributeType" "Not Configured"; }
            @{ Description = "IP Vendor ID"; Value = Get-StringProperty $AUTHRADIUS "-ipVendorID" "Not Configured"; }
            @{ Description = "IP Attribute Type"; Value = Get-StringProperty $AUTHRADIUS "-ipAttributeType" "Not Configured"; }
            @{ Description = "Accounting"; Value = Test-StringPropertyOnOff $AUTHRADIUS "-accounting"; }
            @{ Description = "Password Vendor ID"; Value = Get-StringProperty $AUTHRADIUS "-pwdVendorID" "Not Configured"; }
            @{ Description = "Password Attribute Type"; Value = Get-StringProperty $AUTHRADIUS "-pwdAttributeType" "Not Configured"; }
            @{ Description = "Default Authentication Group"; Value = Get-StringProperty $AUTHRADIUS "-defaultAuthenticationGroup" "Not Configured"; }            @{ Description = "Calling Station ID"; Value = Test-StringPropertyEnabledDisabled $AUTHRADIUS "-callingstationid"; }
            );

        ## IB - Create the parameters to pass to the AddWordTable function
        $Params = $null
        $Params = @{
            Hashtable = $RADIUSCONFIG;
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
}
#endregion Authentication RADIUS

$selection.InsertNewPage()
#endregion NetScaler Authentication

#region NetScaler Certificates
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters Certificates"

WriteWordLine 1 0 "NetScaler Certificates"

$CurrentRowIndex = 0;
if($CERTS.Length -le 0) { WriteWordLine 0 0 "No Certificate has been configured"} else {
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $CERTSH = @();
    
    $CERTS | foreach {
        $CurrentRowIndex++;
        $CERTPropertyArray = Get-StringPropertySplit -SearchString ($_ -Replace 'add ssl certKey' ,'') -RemoveQuotes;

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $CERTSH += @{
            Certificate = $CERTPropertyArray[0]
            CertificateFile = Get-StringProperty $_ "-cert" -RemoveQuotes;
            CertificateKey = Get-StringProperty $_ "-key" -RemoveQuotes;
            Inform = Get-StringProperty $_ "-inform" "NA" -RemoveQuotes;
            }
        }
        if ($CERTSH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $CERTSH;
                Columns = "Certificate","CertificateFile","CertificateKey","Inform";
                Headers = "Certificate","Certificate File","Certificate Key","Inform";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;

            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
    }

$selection.InsertNewPage()

#endregion NetScaler Certificates

#region traffic management

#region NetScaler Content Switches
$Chapter++

Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Content Switches"

WriteWordLine 1 0 "NetScaler Content Switches"

if($ContentSwitches.Length -le 0) { WriteWordLine 0 0 "No Content Switch has been configured"} else {
    $CurrentRowIndex = 0;

    foreach ($ContentSwitch in $ContentSwitches) {
        $CurrentRowIndex++;
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $ContentSwitchPropertyArray = Get-StringPropertySplit -SearchString ($ContentSwitch -Replace 'add cs vserver' ,'') -RemoveQuotes;
        
        ## IB - Store the Load Balancer as native and escaped. This is needed by LBBIND etc.
        $ContentSwitchDisplayNameWithQuotes = Get-StringProperty $ContentSwitch "vserver";
        $ContentSwitchDisplayName = Get-StringProperty $ContentSwitch "vserver" -RemoveQuotes;
        $ContentSwitchName = $ContentSwitchDisplayName.Trim();

        Write-Verbose "$(Get-Date): `tContent Switch $CurrentRowIndex/$($ContentSwitches.Length) $ContentSwitchDisplayName"     
        WriteWordLine 2 0 "Content Switch $ContentSwitchDisplayName";

        If (Test-StringProperty $ContentSwitch "-state") {$STATE = "Disabled"} else {$STATE = "Enabled"}

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $Params = $null
        $Params = @{
            Hashtable = @{
                State = $STATE;
                Protocol = $ContentSwitchPropertyArray[1];
                Port = $ContentSwitchPropertyArray[3];
                IP = $ContentSwitchPropertyArray[2];
                TrafficDomain = Get-StringProperty $ContentSwitch "-td" "0 (Default)";
                CaseSensitive = Get-StringProperty $ContentSwitch "-caseSensitive";
                DownStateFlush = Get-StringProperty $ContentSwitch "-downStateFlush" "Least Connection";
                ClientTimeOut = Get-StringProperty $ContentSwitch "-cltTimeout" "NA";
            }
            Columns = "State","Protocol","Port","IP","TrafficDomain","CaseSensitive","DownStateFlush","ClientTimeOut";
            Headers = "State","Protocol","Port","IP","Traffic Domain","Case Sensitive","Down State Flush","Client Time-Out";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;

        FindWordDocumentEnd;
        WriteWordLine 0 0 " "

        WriteWordLine 3 0 "Policies"

        $ContentSwitchBindMatches = Get-StringWithProperty -SearchString $ContentSwitchBind -Like "bind cs vserver $ContentSwitchDisplayNameWithQuotes *";

        ## Check if we have any specific Content Switch bind matches
        if ($ContentSwitchBindMatches -eq $null -or $ContentSwitchBindMatches.Length -le 0) {
            WriteWordLine 0 0 "No Policy has been configured for this Content Switch"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $ContentSwitchPolicies = @();

            ## IB - Iterate over all Content Switch bindings (uses new function)
            foreach ($CSBind in $ContentSwitchBindMatches) {
                ## IB - Add each Content Switch binding with a policyName to the array
                if (Test-StringProperty -SearchString $CSBind -PropertyName "-policyName") {
                    ## Retrieve the service name from the Content Switch property array (position 3)

                    $AddCsPolicy | foreach {
                        if ($_ -like "add cs policy $(Get-StringProperty $CSBIND "-policyName") *") {
		                    $CSPOLRULE = Get-StringProperty $_ "-rule" -removequotes;
                          }
                        }

                    $ContentSwitchPolicies += @{
                        Policy = Get-StringProperty $CSBIND "-policyName"; 
                        "Load Balancer" = Get-StringProperty $CSBIND "-targetLBVserver";
                        Priority = Get-StringProperty $CSBIND "-priority";
                        Rule = $CSPOLRULE;
                        }
                    }
                } # end foreach

            if ($ContentSwitchPolicies.Length -gt 0) {
                ## IB - Add the table to the document (only if not null!

                ## IB - Create the parameters to pass to the AddWordTable function
                $Params = $null
                $Params = @{
                    Hashtable = $ContentSwitchPolicies;
                    Columns = "Policy","Load Balancer","Priority","Rule";
                    Headers = "Policy Name","Load Balancer","Priority","Rule";
                    AutoFit = $wdAutoFitContent
                    Format = -235; ## IB - Word constant for Light List Accent 5
                }
                ## IB - Add the table to the document, splatting the parameters
                $Table = AddWordTable @Params -NoGridLines;

                FindWordDocumentEnd;
            } else {
                WriteWordLine 0 0 "No policy has been configured for this Content Switch"
            } # end if
        } #end if
        
        ##Table Redirect URL
        WriteWordLine 3 0 "Redirect URL"
            
        if (Test-StringProperty $ContentSwitch "-redirectURL") {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $ContentSwitchRedirects = @(
                @{ RedirectURL = Get-StringProperty $ContentSwitch "-redirectURL" -RemoveQuotes; }
            );

            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = @{
                    RedirectURL = Get-StringProperty $ContentSwitch "-redirectURL" -RemoveQuotes; 
                }
                Columns = "RedirectURL";
                Headers = "Redirect URL";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;

            FindWordDocumentEnd;
            } else { WriteWordLine 0 0 "No Redirect URL has been configured for this Content Switch"; }
            
            WriteWordLine 0 0 " "
            WriteWordLine 3 0 "Advanced Configuration"

            WriteWordLine 0 0 "Need to recheck if all options are correct"

             ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $AdvancedConfiguration = @(                
                @{ Description = "Description"; Value = "Configuration"; }
                @{ Description = "Comment"; Value = Get-StringProperty $ContentSwitch "-comment" "No comment"; }
                @{ Description = "Apply AppFlow logging"; Value = Test-NotStringPropertyEnabledDisabled $ContentSwitch "-appflowLog"; }
                @{ Description = "Name of the TCP profile"; Value = Get-StringProperty $ContentSwitch "-tcpProfileName" "None"; }
                @{ Description = "Name of the HTTP profile"; Value = Get-StringProperty $ContentSwitch "-httpProfileName" "None"; }
                @{ Description = "Name of the NET profile"; Value = Get-StringProperty $ContentSwitch "-netProfile" "None"; }
                @{ Description = "Name of the DB profile"; Value = Get-StringProperty $ContentSwitch "-dbProfileName" "None"; }
                @{ Description = "Enable or disable user authentication"; Value = Test-StringPropertyOnOff $ContentSwitch "-Authentication"; }
                @{ Description = "Authentication virtual server FQDN"; Value = Get-StringProperty $ContentSwitch "-AuthenticationHost" "NA"; }
                @{ Description = "Name of the Authentication profile"; Value = Get-StringProperty $ContentSwitch "-authnProfile" "None"; }
                @{ Description = "Syntax expression identifying traffic"; Value = Test-StringPropertyOnOff $ContentSwitch "-Authentication"; }
                @{ Description = "Priority of the Listener Policy"; Value = Get-StringProperty $ContentSwitch "-AuthenticationHost" "NA"; }
                @{ Description = "Name of the backup virtual server"; Value = Get-StringProperty $ContentSwitch "-authnProfile" "None"; }
                @{ Description = "Enable state updates"; Value = Get-StringProperty $ContentSwitch "-Listenpolicy" "None"; }
                @{ Description = "Route requests to the cache server"; Value = Get-StringProperty $ContentSwitch "-Listenpriority" "101 (Maximum Value)"; }
                @{ Description = "Precedence to use for policies"; Value = Get-StringProperty $ContentSwitch "-backupVServer" "NA"; }
                @{ Description = "URL Case sensitive"; Value = Get-StringProperty $ContentSwitch "-timeout" "2 (Default Value)"; }
                @{ Description = "Type of spillover"; Value = Get-StringProperty $ContentSwitch "-persistenceBackup" "None"; }
                @{ Description = "Maintain source-IP based persistence"; Value = Get-StringProperty $ContentSwitch "-backupPersistenceTimeout" "2 (Default Value)"; }
                @{ Description = "Action if spillover is to take effect"; Value = Test-StringPropertyOnOff $ContentSwitch "-pq"; }
                @{ Description = "State of port rewrite HTTP redirect"; Value = Test-StringPropertyOnOff $ContentSwitch "-sc"; }
                @{ Description = "Continue forwarding to backup vServer"; Value = Test-StringPropertyOnOff $ContentSwitch "-rtspNat"; }
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
        }
    }
$selection.InsertNewPage()

#endregion NetScaler Content Switches

#region NetScaler Cache Redirection
$Chapter++

Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Cache Redirection"

WriteWordLine 1 0 "NetScaler Cache Redirection"

if($CACHEREDIRS.Length -le 0) { WriteWordLine 0 0 "No Cache Redirection has been configured"} else {
    $CurrentRowIndex = 0;

    foreach ($CACHEREDIR in $CACHEREDIRS) {
        $CurrentRowIndex++;
        $CACHEREDIRPropertyArray = Get-StringPropertySplit -SearchString ($CACHEREDIR -Replace 'add cr vserver' ,'') -RemoveQuotes;
        
        ## IB - Store the Load Balancer as native and escaped. This is needed by LBBIND etc.
        $CACHEREDIRDisplayNameWithQuotes = Get-StringProperty $CACHEREDIR "vserver";
        $CACHEREDIRDisplayName = Get-StringProperty $CACHEREDIR "vserver" -RemoveQuotes;

        Write-Verbose "$(Get-Date): `tCache Redirection $CurrentRowIndex/$($ContentSwitches.Length) $CACHEREDIRDisplayName"     
        WriteWordLine 2 0 "Cache Redirection $CACHEREDIRDisplayName";

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $Params = $null
        $Params = @{
            Hashtable = @{
                ## IB - This table will only have 1 row so create the nested hashtable inline
                NAME = $CACHEREDIRPropertyArray[0];
                PROT = $CACHEREDIRPropertyArray[1];
                IP = $CACHEREDIRPropertyArray[2];
                CACHETYPE = Get-StringProperty $CACHEREDIR "-cacheType" "0 (Default)";
                REDIRECT = Get-StringProperty $CACHEREDIR "-redirect";
                CLTTIEMOUT = Get-StringProperty $CACHEREDIR "-cltTimeout";
                DNSVSERVER = Get-StringProperty $CACHEREDIR "-dnsVserverName";
            }
            Columns = "NAME","PROT","IP","CACHETYPE","REDIRECT","CLTTIEMOUT","DNSVSERVER";
            Headers = "NAME","PROT","IP","CACHETYPE","REDIRECT","CLTTIEMOUT","DNSVSERVER";
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

#region NetScaler Load Balancers
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Load Balancers"

WriteWordLine 1 0 "NetScaler Load Balancing"

if($LoadBalancers.Length -le 0) { WriteWordLine 0 0 "No Load Balancer has been configured"} else {
    ## IB - We no longer need to worrying about the number of columns and/or rows.
    ## IB - Need to create a counter of the current row index
    $CurrentRowIndex = 0;

    foreach ($LoadBalancer in $LoadBalancers) {

        $CurrentRowIndex++;
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $LoadBalancerPropertyArray = Get-StringPropertySplit -SearchString ($LoadBalancer -Replace 'add lb vserver' ,'') -RemoveQuotes;
        
        ## IB - Store the Load Balancer as native and escaped. This is needed by LBBIND etc.
        $LoadBalancerDisplayName = Get-StringProperty $LoadBalancer "vserver" -RemoveQuotes;
        $LoadBalancerDisplayNameWithQoutes = Get-StringProperty $LoadBalancer "vserver";
        $LoadBalancerName = $LoadBalancerDisplayName.Trim();

        Write-Verbose "$(Get-Date): `tLoad Balancer $CurrentRowIndex/$($LoadBalancers.Length) $LoadBalancerDisplayName"     
        WriteWordLine 2 0 "Load Balancer $LoadBalancerDisplayName";
        
        If (Test-StringProperty $LoadBalancer "-state") {$STATE = "Disabled"} else {$STATE = "Enabled"}

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $Params = $null
        $Params = @{
            Hashtable = @{
                ## IB - This table will only have 1 row so create the nested hashtable inline
                State = $STATE;
                Protocol = $LoadBalancerPropertyArray[1];
                Port = $LoadBalancerPropertyArray[3];
                IP = $LoadBalancerPropertyArray[2];
                Persistency = Get-StringProperty $LoadBalancer "-persistenceType";
                TrafficDomain = Get-StringProperty $LoadBalancer "-td" "0 (Default)";
                Method = Get-StringProperty $LoadBalancer "-lbmethod" "Least Connection";
                ClientTimeOut = Get-StringProperty $LoadBalancer "-cltTimeout" "NA";
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

        $LoadBalancerBindMatches = Get-StringWithProperty -SearchString $LoadbalancerBind -Like "bind lb vserver $LoadBalancerDisplayNameWithQoutes *";
        ## Check if we have any specific load balancer bind matches
        if ($LoadBalancerBindMatches -eq $null -or $LoadBalancerBindMatches.Length -le 0) {
            WriteWordLine 0 0 "No Service (Group) has been configured for this Load Balancer"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $LoadBalancerServices = @();

            ## IB - Iterate over all load balancer bindings (uses new function)
            foreach ($LBBind in $LoadBalancerBindMatches) {
                ## IB - Add each load balancer binding with a policyName to the array
                if (-not (Test-StringProperty -SearchString $LBBind -PropertyName "-policyName")) {
                    ## Retrieve the service name from the load balancer property array (position 3)
                    $LoadBalancerServices += @{ Service = (Get-StringPropertySplit $LBBind)[4]; }
                }
            } # end foreach

            if ($LoadBalancerServices.Length -gt 0) {
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
            } else {
                WriteWordLine 0 0 "No Service (Group) has been configured for this Load Balancer"
            } # end if
        }
        FindWordDocumentEnd;

        WriteWordLine 0 0 " "
        WriteWordLine 3 0 "Policies"

        if ($LoadBalancerBind.Length -le 0) {
            WriteWordLine 0 0 "No Policy has been configured for this Load Balancer"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $LoadBalancerPolicies = @();

            ## IB - Iterate over all load balancer bindings (uses new function)
            foreach ($LBBind in (Get-StringWithProperty -SearchString $LoadBalancerBind -Like "bind lb vserver $LoadBalancerDisplayNameWithQoutes *")) {
                ## IB - Add each load balancer binding with a policyName to the array
                if (Test-StringProperty -SearchString $LBBind -PropertyName "-policyName") {
                    $LoadBalancerPolicies += @{
                        Name = Get-StringProperty $LBBind "-policyName" -RemoveQuotes;
                        Priority = Get-StringProperty $LBBind "-priority" "NA";
                        Type = Get-StringProperty $LBBind "-type" "NA";
                        Expression = Get-StringProperty $LBBind "-gotoPriorityExpression" "NA"; }
                } # end if
            } # end foreach

            if ($LoadBalancerPolicies.Length -gt 0) {
                ## IB - Add the table to the document (only if not null!

                ## IB - Create the parameters to pass to the AddWordTable function
                $Params = $null
                $Params = @{
                    Hashtable = $LoadBalancerPolicies;
                    Columns = "Name","Priority","Type","Expression";
                    Headers = "Policy Name","Priority","Policy Type","GoTo Expression";
                    AutoFit = $wdAutoFitContent
                    Format = -235; ## IB - Word constant for Light List Accent 5
                }
                ## IB - Add the table to the document, splatting the parameters
                $Table = AddWordTable @Params -NoGridLines;
                ## IB - Set the header background and bold font
                #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

                FindWordDocumentEnd;
            } else {
                WriteWordLine 0 0 "No Policy has been configured for this Load Balancer"
            } # end if
        } #end if

        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
        WriteWordLine 3 0 "Redirect URL"

        if (Test-StringProperty $LoadBalancer "-redirectURL") {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $LoadBalancerRedirects = @(
                @{ RedirectURL = Get-StringProperty $LoadBalancer "-redirectURL" -RemoveQuotes; }
            );

            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = @{
                    RedirectURL = Get-StringProperty $LoadBalancer "-redirectURL" -RemoveQuotes; 
                }
                Columns = "RedirectURL";
                Headers = "Redirect URL";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

            FindWordDocumentEnd;
        } else { WriteWordLine 0 0 "No Redirect URL has been configured for this Load Balancer"; }
        
        ##Advanced Configuration   
        WriteWordLine 0 0 " "
        WriteWordLine 3 0 "Advanced Configuration"

        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $AdvancedConfiguration = @(
            @{ Description = "Description"; Value = "Configuration"; }
            @{ Description = "Comment"; Value = Get-StringProperty $LoadBalancer "-comment" "No comment"; }
            @{ Description = "Apply AppFlow logging"; Value = Test-NotStringPropertyEnabledDisabled $LoadBalancer "-appflowLog"; }
            @{ Description = "Name of the TCP profile"; Value = Get-StringProperty $LoadBalancer "-tcpProfileName" "None"; }
            @{ Description = "Name of the HTTP profile"; Value = Get-StringProperty $LoadBalancer "-httpProfileName" "None"; }
            @{ Description = "Name of the NET profile"; Value = Get-StringProperty $LoadBalancer "-netProfile" "None"; }
            @{ Description = "Name of the DB profile"; Value = Get-StringProperty $LoadBalancer "-dbProfileName" "None"; }
            @{ Description = "Enable or disable user authentication"; Value = Test-StringPropertyOnOff $LoadBalancer "-Authentication"; }
            @{ Description = "Authentication virtual server FQDN"; Value = Get-StringProperty $LoadBalancer "-AuthenticationHost" "NA"; }
            @{ Description = "Authentication virtual server name"; Value = Get-StringProperty $LoadBalancer "-authnVsname" "NA"; }
            @{ Description = "Name of the Authentication profile"; Value = Get-StringProperty $LoadBalancer "-authnProfile" "None"; }
            @{ Description = "User authentication with HTTP 401"; Value = Test-StringPropertyOnOff $LoadBalancer "-authn401"; }
            @{ Description = "Syntax expression identifying traffic"; Value = Get-StringProperty $LoadBalancer "-Listenpolicy" "None"; }
            @{ Description = "Priority of the Listener Policy"; Value = Get-StringProperty $LoadBalancer "-Listenpriority" "101 (Maximum Value)"; }
            @{ Description = "Name of the backup virtual server"; Value = Get-StringProperty $LoadBalancer "-backupVServer" "NA"; }
            @{ Description = "Time period a persistence session"; Value = Get-StringProperty $LoadBalancer "-timeout" "2 (Default Value)"; }
            @{ Description = "Backup persistence type"; Value = Get-StringProperty $LoadBalancer "-persistenceBackup" "None"; }
            @{ Description = "Time period a backup persistence session"; Value = Get-StringProperty $LoadBalancer "-backupPersistenceTimeout" "2 (Default Value)"; }
            @{ Description = "Use priority queuing"; Value = Test-StringPropertyOnOff $LoadBalancer "-pq"; }
            @{ Description = "Use SureConnect"; Value = Test-StringPropertyOnOff $LoadBalancer "-sc"; }
            @{ Description = "Use network address translation"; Value = Test-StringPropertyOnOff $LoadBalancer "-rtspNat"; }
            @{ Description = "Redirection mode for load balancing"; Value = Get-StringProperty $LoadBalancer "-m" "IP Based"; }
            @{ Description = "Use Layer 2 parameter"; Value = Test-StringPropertyOnOff $LoadBalancer "-l2Conn"; }
            @{ Description = "TOS ID of the virtual server"; Value = Get-StringProperty $LoadBalancer "-tosId" "0 (Default)"; }
            @{ Description = "Expression against which traffic is evaluated"; Value = Get-StringProperty $LoadBalancer "-rule" "None"; }
            @{ Description = "Perform load balancing on a per-packet basis"; Value = Test-StringPropertyEnabledDisabled $LoadBalancer "-sessionless"; }
            @{ Description = "How the NetScaler appliance responds to ping requests"; Value = Get-StringProperty $LoadBalancer "-icmpVsrResponse" "NS_VSR_PASSIVE (Default)"; }
            @{ Description = "Route cacheable requests to a cache redirection server"; Value = Test-StringPropertyYesNo $LoadBalancer "-cacheable"; }
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

#region NetScaler Services
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Services"

FindWordDocumentEnd;

WriteWordLine 1 0 "NetScaler Services"

if($Services.Length -le 0) { WriteWordLine 0 0 "No Service has been configured"} else {
    $CurrentRowIndex = 0;

    foreach ($Service in $Services) {
        $CurrentRowIndex++;
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $ServicePropertyArray = Get-StringPropertySplit -SearchString ($Service -Replace 'add service' ,'') -RemoveQuotes;
        
        ## IB - Store the Load Balancer as native and escaped. This is needed by LBBIND etc.
        $ServiceDisplayNameWithQuotes = Get-StringProperty $Service "service";
        $ServiceDisplayName = Get-StringProperty $Service "service" -RemoveQuotes;
        $ServiceName = $ServiceDisplayName.Trim();
    
        Write-Verbose "$(Get-Date): `tService $CurrentRowIndex/$($Services.Length) $ServiceDisplayName"     
        WriteWordLine 2 0 "Service $ServiceDisplayName"

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $Params = $null
        $Params = @{
            Hashtable = @{
                ## IB - This table will only have 1 row so create the nested hashtable inline
                Server = $ServicePropertyArray[1];
                Protocol = $ServicePropertyArray[2];
                Port = $ServicePropertyArray[3];
                TD = Get-StringProperty $Service "-td" "0 (Default)";
                GSLB = Get-StringProperty $Service "-gslb" "NA";
                MaximumClients = Get-StringProperty $Service "-maxClient" "NA";
                MaximumRequests = Get-StringProperty $Service "-maxreq" "NA";
            }
            Columns = "Server","Protocol","Port","TD","GSLB","MaximumClients","MaximumRequests";
            Headers = "Server","Protocol","Port","Traffic Domain","GSLB","Maximum Clients","Maximum Requests";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;

        FindWordDocumentEnd;
        WriteWordLine 0 0 " "

        WriteWordLine 3 0 "Monitor"

        $ServiceBindMatches = Get-StringWithProperty -SearchString $ServiceBind -Like "bind service $ServiceDisplayNameWithQuotes *";
        ## Check if we have any specific Service bind matches
        if ($ServiceBindMatches -eq $null -or $ServiceBindMatches.Length -le 0) {
            WriteWordLine 0 0 "No Monitor has been configured for this Service"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $ServiceMonitors = @();

            ## IB - Iterate over all Service bindings (uses new function)
            foreach ($SVCBind in $ServiceBindMatches) {
                if (Test-StringProperty -SearchString $SVCBind -PropertyName "-monitorName") {
                    $ServiceMonitors += @{ Monitor = Get-StringProperty $SVCBIND "-monitorName"; }
                }
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
            }
        } # end if

        WriteWordLine 0 0 " "
        WriteWordLine 3 0 "Advanced Configuration"

        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $AdvancedConfiguration = @(
            @{ Description = "Description"; Value = "Configuration"; }
            @{ Description = "Clear text port"; Value = Get-StringProperty $Service "-clearTextPort" "NA" ; }
			@{ Description = "Cache Type"; Value = Get-StringProperty $Service "-cacheType" "NA" ; }
			@{ Description = "Maximum Client Requests"; Value = Get-StringProperty $Service "-maxClient" "4294967294 (Maximum Value)" ; }
			@{ Description = "Monitor health of this service"; Value = Test-NotStringPropertyYesNo $Service "-healthMonitor" ; }
			@{ Description = "Maximum Requests"; Value = Get-StringProperty $Service "-maxreq" "65535 (Maximum Value)" ; }
			@{ Description = "Use Transparent Cache"; Value = Test-StringPropertyYesNo $Service "-cacheable" ; }
			@{ Description = "Insert the Client IP header"; Value = Get-StringProperty $Service "-cip" "DISABLED"  ; }
			##@{ Description = "Name for the HTTP header"; Value = Get-StringProperty $Service "-cipHeader" "NA" ; }
			@{ Description = "Use Source IP"; Value = Test-NotStringPropertyYesNo $Service "-usip" ; }
            @{ Description = "Path Monitoring"; Value = Test-StringPropertyYesNo $Service "-pathMonitor" ; }
			@{ Description = "Individual Path monitoring"; Value = Test-StringPropertyYesNo $Service "-pathMonitorIndv" ; }
			@{ Description = "Use the proxy port"; Value = Test-StringPropertyYesNo $Service "-useproxyport" ; }
			@{ Description = "SureConnect"; Value = Test-StringPropertyOnOff $Service "-sc" ; }
			@{ Description = "Surge protection"; Value = Test-NotStringPropertyOnOff $Service "-sp" ; }
			@{ Description = "RTSP session ID mapping"; Value = Test-StringPropertyOnOff $Service "-rtspSessionidRemap" ; }
			@{ Description = "Client Time-Out"; Value = Get-StringProperty $Service "-cltTimeout" "31536000 (Maximum Value)" ; }
			@{ Description = "Server Time-Out"; Value = Get-StringProperty $Service "-svrTimeout" "3153600 (Maximum Value)" ; }
			@{ Description = "Unique identifier for the service"; Value = Get-StringProperty $Service "-CustomServerID" "None" -RemoveQuotes; }
			@{ Description = "The identifier for the service"; Value = Get-StringProperty $Service "-serverID" "None" ; }
			@{ Description = "Enable client keep-alive"; Value = Test-NotStringPropertyYesNo $Service "-CKA" ; }
			@{ Description = "Enable TCP buffering"; Value = Test-NotStringPropertyYesNo $Service "-TCPB" ; }
            @{ Description = "Enable compression"; Value = Test-StringPropertyYesNo $Service "-CMP" ; }
			@{ Description = "Maximum bandwidth, in Kbps"; Value = Get-StringProperty $Service "-maxBandwidth" "4294967287 (Maximum Value)" ; }
			@{ Description = "Use Layer 2 mode"; Value = Test-StringPropertyYesNo $Service "-accessDown" ; }
			@{ Description = "Sum of weights of the monitors"; Value = Get-StringProperty $Service "-monThreshold" "65535 (Maximum Value)" ; }
			@{ Description = "Initial state of the service"; Value = Test-NotStringPropertyEnabledDisabled $Service "-state" ; }
			@{ Description = "Perform delayed clean-up"; Value = Test-NotStringPropertyEnabledDisabled $Service "-downStateFlush" ; }
			@{ Description = "TCP profile"; Value = Get-StringProperty $Service "-tcppProfileName" "NA" ; }
			@{ Description = "HTTP profile"; Value = Get-StringProperty $Service "-httpProfileName" "NA" ; }
			@{ Description = "A numerical identifier"; Value = Get-StringProperty $Service "-hashId" "NA" ; }
			@{ Description = "Comment about the service"; Value = Get-StringProperty $Service "-comment" "NA"; }
			@{ Description = "Logging of AppFlow information"; Value = Test-NotStringPropertyEnabledDisabled $Service "-appflowLog" ; }
			@{ Description = "Network profile"; Value = Get-StringProperty $Service "-netProfile" "NA" ; }
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

if($ServiceGroups.Length -le 0) { WriteWordLine 0 0 "No Service Group has been configured"} else {
    $CurrentRowIndex = 0;

    foreach ($Servicegroup in $ServiceGroups) {
        $CurrentRowIndex++;
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $ServicegroupPropertyArray = Get-StringPropertySplit -SearchString ($Servicegroup -Replace 'add servicegroup' ,'') -RemoveQuotes;
        ## IB - Store the Load Balancer as native and escaped. This is needed by LBBIND etc.
        $ServicegroupDisplayNameWithQuotes = Get-StringProperty $Servicegroup "servicegroup";
        $ServicegroupDisplayName = Get-StringProperty $Servicegroup "servicegroup" -RemoveQuotes;
        $ServicegroupName = $ServicegroupDisplayName.Trim();
        
        Write-Verbose "$(Get-Date): `tService Group $CurrentRowIndex/$($ServiceGroups.Length) $ServicegroupDisplayName"     
        WriteWordLine 2 0 "Service Group $ServicegroupDisplayName"

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $Params = $null
        $Params = @{
            Hashtable = @{
                ## IB - This table will only have 1 row so create the nested hashtable inline
                Server = $ServicegroupPropertyArray[0];
                Protocol = $ServicegroupPropertyArray[1];
                Port = $ServicegroupPropertyArray[3];
                TD = Get-StringProperty $Servicegroup "-td" "0 (Default)";
                GSLB = Get-StringProperty $Servicegroup "-gslb" "NA";
                MaximumClients = Get-StringProperty $Servicegroup "-maxClient" "NA";
                MaximumRequests = Get-StringProperty $Servicegroup "-maxreq" "NA";
            }
            Columns = "Server","Protocol","Port","TD","GSLB","MaximumClients","MaximumRequests";
            Headers = "Server","Protocol","Port","Traffic Domain","GSLB","Maximum Clients","Maximum Requests";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;

        FindWordDocumentEnd;
        WriteWordLine 0 0 " "

        $ServicegroupBindMatches = Get-StringWithProperty -SearchString $ServiceGroupBind -Like "bind serviceGroup $ServicegroupDisplayNameWithQuotes *";

        WriteWordLine 3 0 "Servers"

        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $ServiceGroupServers = @();

        ## IB - Iterate over all Service bindings (uses new function)
        foreach ($Server in $Servers) {
            $ServerSGPropertyArray = Get-StringPropertySplit -SearchString ($Server -Replace 'add server' ,'') -RemoveQuotes;    
            $SRVSGNAME = $ServerSGPropertyArray[0];
            
            foreach ($SVCGroupBind in $ServiceGroupBindMatches) {
                $SGBINDNAME = Get-StringPropertySplit -SearchString ($SVCGroupBind -Replace 'bind serviceGroup' ,'') -RemoveQuotes;    
                $SGBINDNAME = $SGBINDNAME[1]

                If ($SRVSGNAME -eq $SGBINDNAME) {
                    $ServiceGroupServers += @{ Server = $SRVSGNAME; }
                }
            }
        }
        $ServiceGroupServers
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

        if ($ServicegroupBindMatches -eq $null -or $ServicegroupBindMatches.Length -le 0) {
            WriteWordLine 0 0 "No Monitor has been configured for this Service Group"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $ServiceGroupMonitors = @();

            ## IB - Iterate over all Service bindings (uses new function)
            foreach ($SVCGroupBind in $ServiceGroupBindMatches) {

                if (Test-StringProperty -SearchString $SVCGroupBind -PropertyName "-monitorName") {
                    $ServiceGroupMonitors += @{ Monitor = Get-StringProperty $SVCGroupBind "-monitorName"; }
                }
            } # end foreach

            if ($ServiceGroupMonitors.Length -gt 0) {
                ## IB - Create the parameters to pass to the AddWordTable function
                $Params = $null
                $Params = @{
                    Hashtable = $ServiceGroupMonitors;                   
                    AutoFit = $wdAutoFitContent;
                    Format = -235; ## IB - Word constant for Light List Accent 5
                }
                ## IB - Add the table to the document, splatting the parameters
                $Table = AddWordTable @Params -NoGridLines;
                ## IB - Set the header background and bold font
                #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;
                
                FindWordDocumentEnd;
            } else {
                WriteWordLine 0 0 "No Monitor has been configured for this Service Group"
        }   
        } # end if

        WriteWordLine 0 0 " "
        WriteWordLine 3 0 "Advanced Configuration"

        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $AdvancedConfiguration = @(
            @{ Description = "Description"; Value = "Configuration"; }
            @{ Description = "Clear text port"; Value = Get-StringProperty $ServiceGroup "-clearTextPort" "NA" ; }
			@{ Description = "Cache Type"; Value = Get-StringProperty $ServiceGroup "-cacheType" "NA" ; }
			@{ Description = "Maximum Client Requests"; Value = Get-StringProperty $ServiceGroup "-maxClient" "4294967294 (Maximum Value)" ; }
			@{ Description = "Monitor health of this Service Group"; Value = Test-NotStringPropertyYesNo $ServiceGroup "-healthMonitor" ; }
			@{ Description = "Maximum Requests"; Value = Get-StringProperty $ServiceGroup "-maxreq" "65535 (Maximum Value)" ; }
			@{ Description = "Use Transparent Cache"; Value = Test-StringPropertyYesNo $ServiceGroup "-cacheable" ; }
			@{ Description = "Insert the Client IP header"; Value = Get-StringProperty $ServiceGroup "-cip" "NA"  ; }
			@{ Description = "Name for the HTTP header"; Value = Get-StringProperty $ServiceGroup "-cipHeader" "NA" ; }
			@{ Description = "Use Source IP"; Value = Test-StringPropertyYesNo $ServiceGroup "-usip" ; }
            @{ Description = "Path Monitoring"; Value = Test-StringPropertyYesNo $ServiceGroup "-pathMonitor" ; }
			@{ Description = "Individual Path monitoring"; Value = Test-StringPropertyYesNo $ServiceGroup "-pathMonitorIndv" ; }
			@{ Description = "Use the proxy port"; Value = Test-StringPropertyYesNo $ServiceGroup "-useproxyport" ; }
			@{ Description = "SureConnect"; Value = Test-StringPropertyOnOff $ServiceGroup "-sc" ; }
			@{ Description = "Surge protection"; Value = Test-StringPropertyOnOff $ServiceGroup "-sp" ; }
			@{ Description = "RTSP session ID mapping"; Value = Test-StringPropertyOnOff $ServiceGroup "-rtspSessionidRemap" ; }
			@{ Description = "Client Time-Out"; Value = Get-StringProperty $ServiceGroup "-cltTimeout" "31536000 (Maximum Value)" ; }
			@{ Description = "Server Time-Out"; Value = Get-StringProperty $ServiceGroup "-svrTimeout" "3153600 (Maximum Value)" ; }
			@{ Description = "Unique identifier for the Service Group"; Value = Get-StringProperty $ServiceGroup "-CustomServerID" "None" ; }
			@{ Description = "The identifier for the Service Group"; Value = Get-StringProperty $ServiceGroup "-serverID" "None" ; }
			@{ Description = "Enable client keep-alive"; Value = Test-StringPropertyYesNo $ServiceGroup "-CKA" ; }
			@{ Description = "Enable TCP buffering"; Value = Test-StringPropertyYesNo $ServiceGroup "-TCPB" ; }
            @{ Description = "Enable compression"; Value = Test-StringPropertyYesNo $ServiceGroup "-CMP" ; }
			@{ Description = "Maximum bandwidth, in Kbps"; Value = Get-StringProperty $ServiceGroup "-maxBandwidth" "4294967287 (Maximum Value)" ; }
			@{ Description = "Use Layer 2 mode"; Value = Test-StringPropertyYesNo $ServiceGroup "-accessDown" ; }
			@{ Description = "Sum of weights of the monitors"; Value = Get-StringProperty $ServiceGroup "-monThreshold" "65535 (Maximum Value)" ; }
			@{ Description = "Initial state of the Service Group"; Value = Test-NotStringPropertyEnabledDisabled $ServiceGroup "-state" ; }
			@{ Description = "Perform delayed clean-up"; Value = Test-NotStringPropertyEnabledDisabled $ServiceGroup "-downStateFlush" ; }
			@{ Description = "TCP profile"; Value = Get-StringProperty $ServiceGroup "-tcppProfileName" "NA" ; }
			@{ Description = "HTTP profile"; Value = Get-StringProperty $ServiceGroup "-httpProfileName" "NA" ; }
			@{ Description = "A numerical identifier"; Value = Get-StringProperty $ServiceGroup "-hashId" "NA" ; }
			@{ Description = "Comment about the ServiceGroup"; Value = Get-StringProperty $ServiceGroup "-comment" "NA"; }
			@{ Description = "Logging of AppFlow information"; Value = Test-NotStringPropertyEnabledDisabled $ServiceGroup "-appflowLog" ; }
			@{ Description = "Network profile"; Value = Get-StringProperty $ServiceGroup "-netProfile" "NA" ; }
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
$selection.InsertNewPage() 
#endregion NetScaler Service Groups

#region NetScaler Servers
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Servers"
WriteWordLine 1 0 "NetScaler Servers"

if($Servers.Length -le 0) { WriteWordLine 0 0 "No Server has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $ServersH = @();

    foreach ($Server in $Servers) {
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $ServerPropertyArray = Get-StringPropertySplit -SearchString ($Server -Replace 'add Server' ,'') -RemoveQuotes;

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!

        $ServersH += @{
                Server = $ServerPropertyArray[0];
                IP = $ServerPropertyArray[1];
                TD = Get-StringProperty $Server "-td" "0 (Default)";
                STATE = Test-NotStringPropertyEnabledDisabled $Server "-state";
                COMMENT = Get-StringProperty $Server "-comment" "No Comment" -RemoveQuotes;
            }
        }
        if ($ServersH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $ServersH;
                Columns = "Server","IP","TD","STATE","COMMENT";
                Headers = "Server","IP Address","Traffic Domain","State","Comment";
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

#region Citrix NetScaler Gateway CAG Global...

WriteWordLine 2 0 "NetScaler Gateway Global Settings"
Write-Verbose "$(Get-Date): `tNetScaler Gateway Global Settings"
#region GlobalNetwork

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
#endregion GlobalNetwork

#region GlobalClientExperience
WriteWordLine 3 0 "Global Settings Client Experience"

ForEach ($LINE in $SetVpnParameter) {

    ## IB - Create an array of hashtables to store our columns.
    ## IB - about column names as we'll utilise a -List(view)!
    [System.Collections.Hashtable[]] $NsGlobalClientExperience = @(
        ## IB - Each hashtable is a separate row in the table!
        @{ Column1 = "Description"; Column2 = "Value"; }
        @{ Column1 = "Home Page"; Column2 = Get-StringProperty $LINE "-homePage" "Not Configured"; }
        @{ Column1 = "URL for Web Based E-mail"; Column2 = Get-StringProperty $LINE "-useMIP" "Not Configured"; }
        @{ Column1 = "Split Tunnel"; Column2 = Get-StringProperty $LINE "-splitTunnel" "Off"; }
        @{ Column1 = "Session Time-Out"; Column2 = Get-StringProperty $LINE "-sessTimeout" "0"; }
        @{ Column1 = "Client-Idle Time-Out"; Column2 = Get-StringProperty $LINE "-clientIdleTimeout" "0"; }
        @{ Column1 = "Plug-in Type"; Column2 = Get-StringProperty $LINE "-epaClientType" "AGENT"; }
        @{ Column1 = "Clientless Access"; Column2 = Get-StringProperty $LINE "-clientlessVpnMode" "Off"; }
        @{ Column1 = "Clientless URL Encoding"; Column2 = Get-StringProperty $LINE "-clientlessModeUrlEncoding" "VPN_SESS_ACT_CVPN_ENC_OPAQUE"; }
        @{ Column1 = "Clientless Persistent Cookie"; Column2 = Get-StringProperty $LINE "-clientlessPersistentCookie" "Deny"; }
        @{ Column1 = "Single Sign-On to Web Applications"; Column2 = Get-StringProperty $LINE "-SSO" "Off"; }
        @{ Column1 = "Credential Index"; Column2 = Get-StringProperty $LINE "-ssoCredential" "Primary"; }
        @{ Column1 = "KCD Account"; Column2 = Get-StringProperty $LINE "-kcdAccount" "Not Configured"; }
        @{ Column1 = "Single Sign-On with Windows"; Column2 = Get-StringProperty $LINE "-windowsAutoLogon" "Off"; }
        @{ Column1 = "Client Cleanup Prompt"; Column2 = Get-StringProperty $LINE "-forceCleanup" "Off"; }
        @{ Column1 = "UI Theme"; Column2 = Get-StringProperty $LINE "-UITHEME" "DEFAULT"; }
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
    }
FindWordDocumentEnd;

$NsGlobalClientExperience = $null;

WriteWordLine 0 0 " "
#endregion GlobalClientExperience

#region GlobalSecurity
WriteWordLine 3 0 "Global Settings Security"

ForEach ($LINE in $SetVpnParameter) {

    ## IB - Create an array of hashtables to store our columns. Note: If we need the
    ## IB - headers to include spaces we can override these at table creation time.
    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = @{
            ## IB - Each hashtable is a separate row in the table!
            DEFAUTH = Get-StringProperty $LINE "-defaultAuthorizationAction" "DENY";
            CLISEC = Get-StringProperty $LINE "-encryptCsecExp" "Disabled";
            SECBRW = Get-StringProperty $LINE "-SecureBrowse" "Enabled";
        }
        Columns = "DEFAUTH","CLISEC","SECBRW";
        Headers = "Default Authorization Action","Client Security Encryption","Secure Browse";
        AutoFit = $wdAutoFitContent;
        Format = -235; ## IB - Word constant for Light List Accent 5
    }

    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;
    }

WriteWordLine 0 0 " "
#endregion GlobalSecurity

#region GlobalPublishedApps
WriteWordLine 3 0 "Global Settings Published Applications"

ForEach ($LINE in $SetVpnParameter) {

    ## IB - Create an array of hashtables to store our columns. Note: If we need the
    ## IB - headers to include spaces we can override these at table creation time.
    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = @{
            ICAPROXY = Get-StringProperty $LINE "-icaProxy" "OFF";
            WIADDR = Get-StringProperty $LINE "-wihome" "Not Configured" -RemoveQuotes;
            WIMODE = Get-StringProperty $LINE "-wiPortalMode" "NORMAL";
            SSO = Get-StringProperty $LINE "-ntDomain" "Not Configured";
            HOME = Get-StringProperty $LINE "-citrixReceiverHome" "Not Configured";
            ACCSVC = Get-StringProperty $LINE "-storefronturl" "Not Configured";
        }
        Columns = "ICAPROXY","WIADDR","WIMODE","SSO","HOME","ACCSVC";
        Headers = "ICA Proxy","Web Interface addres","Web Interface Portal Mode","Single Sign-On Domain","Citrix Receiver Home Page","Account Services Address";
        AutoFit = $wdAutoFitContent;
        Format = -235; ## IB - Word constant for Light List Accent 5
    }

    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;
    }

WriteWordLine 0 0 " "

#endregion GlobalPublishedApps

#region Global PREAUTH
WriteWordLine 3 0 "Global Settings Pre-Authentication Settings"
if($SETAAAPREAUTH.Length -le 0) { $SETAAAPREAUTH = "123"}

ForEach ($AAAAUTH in $SETAAAPREAUTH) {
    ## IB - Create an array of hashtables to store our columns. Note: If we need the
    ## IB - headers to include spaces we can override these at table creation time.
    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = @{
            ## IB - This table will only have 1 row so create the nested hashtable inline
            ACTION = Get-StringProperty $AAAAUTH "-preauthenticationaction" "ALLOW";
            PROC1 = Get-StringProperty $AAAAUTH "-killProcess" "Not Configured";
            FILES1 = Get-StringProperty $AAAAUTH "-deletefiles" "Not Configured";
            Expr1 = Get-StringProperty $AAAAUTH "-rule" "Not Configured";
        }
        Columns = "ACTION","PROC1","FILES1","Expr1";
        Headers = "Action","Processes to be cancelled","Files to be deleted","Expression";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
    }
    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;
    }

WriteWordLine 0 0 " "
#endregion Global PREAUTH

#region GlobalAuthentication
WriteWordLine 3 0 "Global Settings Authentication Settings"

$Set | foreach {  
    if ($_ -like 'set aaa parameter *') {
        ## IB - Create an array of hashtables to store our columns. Note: If we need the
        ## IB - headers to include spaces we can override these at table creation time.
        ## IB - Create the parameters to pass to the AddWordTable function
        $Params = $null
        $Params = @{
            Hashtable = @{
                ## IB - Each hashtable is a separate row in the table!
                MAXUSR = Get-StringProperty $_ "-maxAAAUsers" "1";
                NATIP = Get-StringProperty $_ "-aaadnatIp" "Default Setting";
                MAXLOG = Get-StringProperty $_ "-maxLoginAttempts" "Unlimited";
                FAILTO = Get-StringProperty $_ "-failedLoginTimeout" "Default Setting";
                ENSTAT = Get-StringProperty $_ "-enableStaticPageCaching" "Enabled";
                ENADV = Get-StringProperty $_ "-enableEnhancedAuthFeedback" "Disabled";
                DEFAUTH = Get-StringProperty $_ "-defaultAuthType" "Local Authentication";
            }
            Columns = "MAXUSR","NATIP","MAXLOG","FAILTO","ENSTAT","ENADV","DEFAUTH";
            Headers = "Maximum Number of Users","NAT IP Address","Maximum login Attempts","Failed Login Timeout","Enable Static Caching","Enable advanced authentication feedback","Default Authentication Type";
            AutoFit = $wdAutoFitContent;
            Format = -235; ## IB - Word constant for Light List Accent 5
        }
        ## IB - Add the table to the document, splatting the parameters
        $Table = AddWordTable @Params -NoGridLines;
        FindWordDocumentEnd;
        }
    }

FindWordDocumentEnd;

WriteWordLine 0 0 " "
#endregion GlobalAuthentication

#region Global STA
WriteWordLine 3 0 "Global Settings Secure Ticket Authority Configuration"
$STAMATCHES = Get-StringWithProperty -SearchString $Bind -Like "bind vpn global -staServer *";

if ($STAMATCHES -eq $null -or $STAMatches.Length -le 0) {
            WriteWordLine 0 0 "No Secure Ticket Authority has been configured"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $STAS = @();

                ## IB - Iterate over all load balancer bindings (uses new function)
            foreach ($STALINE in $STAMATCHES) {
                $STAS += @{ STA = Get-StringProperty $STALINE "-staServer" -RemoveQuotes; }
                } # end foreach
            $Params = $null
            $Params = @{
                Hashtable = $STAS;
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
                }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            }
FindWordDocumentEnd;
WriteWordLine 0 0 " "

#endregion Global STA

#region Global AppController
WriteWordLine 3 0 "Global Settings App Controller Configuration"
$APPCMATCHES = Get-StringWithProperty -SearchString $Bind -Like "bind vpn global -appController *";

if ($APPCMATCHES -eq $null -or $APPCMatches.Length -le 0) {
            WriteWordLine 0 0 "No App Controller has been configured"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $APPCS = @();

                ## IB - Iterate over all load balancer bindings (uses new function)
            foreach ($APPCLINE in $APPCMATCHES) {
                $APPCS += @{ "APP Controller" = Get-StringProperty $APPCLINE "-appController" -RemoveQuotes; }
                } # end foreach

            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $APPCS;
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

#endregion CAG Global

#region CAG vServers

if($AccessGateways.Length -le 0) { WriteWordLine 0 0 "No Citrix NetScaler Gateway has been configured"} else {
    $CurrentRowIndex = 0;
    foreach ($AccessGateway in $AccessGateways) {
        $CurrentRowIndex++;

        ## IB - Store the Load Balancer as native and escaped. This is needed by LBBIND etc.
        $vServerDisplayName = Get-StringProperty $AccessGateway "vserver";
        $vServerDisplayNameNoQuotes = Get-StringProperty $AccessGateway "vserver" -RemoveQuotes;
        $vServerName = $vServerDisplayName.Trim();

        WriteWordLine 2 0 "NetScaler Gateway Virtual Server: $(Get-StringProperty $AccessGateway "vserver" -RemoveQuotes)";
        Write-Verbose "$(Get-Date): `tNetScaler Gateway $CurrentRowIndex/$($AccessGateways.Length) : $vServerDisplayNameNoQuotes";

#region CAG vServer basic configuration

        $AGPropertyArray = Get-StringPropertySplit -SearchString ($AccessGateway -Replace 'add vpn vserver' ,'') -RemoveQuotes;

        ## IB - Create an array of hashtables to store our columns. Note: If we need the
        $Params = $null
        $Params = @{
            Hashtable = @{
                State = Get-StringProperty $AccessGateway "-state" "Enabled";
                Mode = Test-NotStringPropertyOnOff $AccessGateway "-icaOnly";
                IPAddress = $AGPropertyArray[2];
                Port = $AGPropertyArray[3];
                Protocol = $AGPropertyArray[1];
                MaximumUsers = Get-StringProperty $AccessGateway "-maxAAAUsers" "Unlimited";
                MaxLogin = Get-StringProperty $AccessGateway "-maxLoginAttempts" "Unlimited";
            }
            Columns = "State","Mode","IPAddress","Port","Protocol","MaximumUsers","MaxLogin";
            Headers = "State","Smart Access","IP Address","Port","Protocol","Maximum Users","Maximum Logons";
            AutoFit = $wdAutoFitContent
            Format = -235; ## IB - Word constant for Light List Accent 5
        }

        ## IB - Add the table to the document, splatting the parameters
        $Table = AddWordTable @Params -NoGridLines;
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
#endregion CAG vServer basic configuration

#region CAG certificate
        
        WriteWordLine 3 0 "Certificates"
        $CAGSERVERCERTBINDS = Get-StringWithProperty -SearchString $CERTBINDS -Like "bind ssl vserver $vServerDisplayName -certkeyName *";
        if($CAGSERVERCERTBINDS.Length -le 0) { WriteWordLine 0 0 "No Certificate has been configured for this NetScaler Gateway vServer"} else { 
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $CAGCERTSH = @();
            
            Foreach ($CERT in $CAGSERVERCERTBINDS) {$CAGCERTSH += @{ Certificate = Get-StringProperty $CERT "-certkeyName" -RemoveQuotes; }}

            if ($CAGCERTSH.Length -gt 0) {
                ## IB - Create the parameters to pass to the AddWordTable function
                $Params = $null
                $Params = @{
                    Hashtable = $CAGCERTSH;
                    AutoFit = $wdAutoFitContent;
                    Format = -235; ## IB - Word constant for Light List Accent 5
                }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " " 
            }
        }
#endregion CAG certificate

#region CAG vServer policies
        $CAGPOLS = Get-StringWithProperty -SearchString $BINDVPNVSERVER -Like "bind vpn vserver $vServerDisplayName *"; 
    
    #region CAG Authentication LDAP Policies        
        
        WriteWordLine 3 0 "Authentication LDAP Policies"
        
        if($AUTHLDAPPOLS.Length -le 0) { WriteWordLine 0 0 "No LDAP Policy has been configured"} else { 

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $AUTHPOLHASH = @();

            ## BS 1. First we get all LDAP Policies on the NetScaler system
            foreach ($AUTHLDAPPOL in $AUTHLDAPPOLS) {
                $CAGAUTHPOLPropertyArray = Get-StringPropertySplit -SearchString ($AUTHLDAPPOL -Replace 'add authentication ldapPolicy' ,'') ;
                $CAGAUTHPOLPropertyArrayNoQuotes = Get-StringPropertySplit -SearchString ($AUTHLDAPPOL -Replace 'add authentication ldapPolicy' ,'') -RemoveQuotes ;
                
                ## BS 2. Then we determine the policy name for each of the LDAP Policies
                $POLICYNAME = $CAGAUTHPOLPropertyArray[0];
                $LDAPPolicyDisplayName = Get-StringProperty $AUTHLDAPPOL "ldapPolicy" -RemoveQuotes;

                ## BS 3. Now we find out if this specific LDAP policy is bound to this specific CAG vServer
                $AUTHPOLS = Get-StringWithProperty -SearchString $CAGPOLS -Like "bind vpn vserver $vServerDisplayName -policy $POLICYNAME*";

                ## BS 4. If we have an LDAP Policy bound to this specific CAG vServer then we 
                foreach ($AUTHPOL in $AUTHPOLS) {                
                    $PRIMARY = Test-StringProperty $AUTHPOL -PropertyName "-secondary";
                    If ($PRIMARY -eq $True) {$PRIMARY = "Secondary"} else {$PRIMARY = "Primary"}

                    $AUTHPOLHASH += @{
                        Name = $LDAPPolicyDisplayName;
                        Action = $CAGAUTHPOLPropertyArrayNoQuotes[2];
                        Expr = $CAGAUTHPOLPropertyArrayNoQuotes[1];
                        Primary = $PRIMARY ;
                        Priority = Get-StringProperty $AUTHPOL "-priority";
                    } # end Hasthable $AUTHPOLH1
                }# end foreach $AUTHPOLS
            } #end foreach AUTHLDAPPOLS

            if ($AUTHPOLHASH.Length -gt 0) {
                ## IB - Add the table to the document (only if not null!
                ## IB - Create the parameters to pass to the AddWordTable function
                $Params = $null
                $Params = @{
                    Hashtable = $AUTHPOLHASH;
                    Columns = "Name","Action","Expr","Primary","Priority";
                    Headers = "Policy Name","Policy Action","Expression","Primary","Priority";
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
        
        WriteWordLine 3 0 "Authentication Radius Policies"
        if($AUTHRADPOLS.Length -le 0) { WriteWordLine 0 0 "No Radius Policy has been configured"} else {        
            $AUTHPOLRADHASH = $null
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $AUTHPOLRADHASH = @();

            ## BS 1. First we get all RADIUS Policies on the NetScaler system
            foreach ($AUTHRADPOL in $AUTHRADPOLS) {
                $CAGAUTHRADPOLPropertyArray = Get-StringPropertySplit -SearchString ($AUTHRADPOL -Replace 'add authentication radiusPolicy' ,'') ;
                $CAGAUTHRADPOLPropertyArrayNoQuotes = Get-StringPropertySplit -SearchString ($AUTHRADPOL -Replace 'add authentication radiusPolicy' ,'') -RemoveQuotes ;
                
                ## BS 2. Then we determine the policy name for each of the LDAP Policies
                $POLICYNAME = $null
                $POLICYNAME = $CAGAUTHRADPOLPropertyArray[0];

                ## BS 3. Now we find out if this specific RADIUS policy is bound to this specific CAG vServer
                $AUTHRADPOLSBIND = Get-StringWithProperty -SearchString $CAGPOLS -Like "bind vpn vserver $vServerDisplayName -policy $POLICYNAME*";

                ## BS 4. If we have an RADIUS Policy bound to this specific CAG vServer then we 
                $AUTHPOL = $null
                foreach ($AUTHPOL in $AUTHRADPOLSBIND) {                
                    $PRIMARY = Test-StringProperty $AUTHPOL -PropertyName "-secondary";
                    If ($PRIMARY -eq $True) {$PRIMARY = "Secondary"} else {$PRIMARY = "Primary"}
                    
                    $AUTHPOLRADHASH += @{
                        Name = $CAGAUTHRADPOLPropertyArrayNoQuotes[0];
                        Action = $CAGAUTHRADPOLPropertyArrayNoQuotes[2];
                        Expr = $CAGAUTHRADPOLPropertyArrayNoQuotes[1];
                        Primary = $PRIMARY ;
                        Priority = Get-StringProperty $AUTHPOL "-priority";
                    } # end Hasthable $AUTHPOLRADHASH
                }# end foreach $AUTHPOLS
            } #end foreach AUTHRADPOLS

            if ($AUTHPOLRADHASH.Length -gt 0) {
                ## IB - Add the table to the document (only if not null!
                ## IB - Create the parameters to pass to the AddWordTable function
                $Params = $null
                $Params = @{
                    Hashtable = $AUTHPOLRADHASH;
                    Columns = "Name","Action","Expr","Primary","Priority";
                    Headers = "Policy Name","Policy Action","Expression","Primary","Priority";
                    AutoFit = $wdAutoFitContent
                    Format = -235; ## IB - Word constant for Light List Accent 5
                }
                ## IB - Add the table to the document, splatting the parameters
                $Table = AddWordTable @Params -NoGridLines;

                FindWordDocumentEnd;
            } else { WriteWordLine 0 0 "No Radius Policy has been configured"} #endif AUTHPOLHASH.Length
        } #End If no policies configured
    WriteWordLine 0 0 " "
    #endregion CAG Authentication Radius Policies  
    
    #region CAG Session Policies        
       
        WriteWordLine 3 0 "Session Policies"
        if($CAGSESSIONPOLS.Length -le 0) { WriteWordLine 0 0 "No Session Policy has been configured"} else { 

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $SESSIONPOLH = @();

            ## BS 1. First we get all Session Policies on the NetScaler system
            foreach ($CAGSESSIONPOL in $CAGSESSIONPOLS) {
                $CAGSESSIONPOLPropertyArray = $Null
                $CAGSESSIONPOLPropertyArray = Get-StringPropertySplit -SearchString ($CAGSESSIONPOL -Replace 'add vpn' ,'') ;
                $CAGSESSIONPOLPropertyArrayNoQuotes = Get-StringPropertySplit -SearchString ($CAGSESSIONPOL -Replace 'add vpn' ,'') -RemoveQuotes ;
                ## BS 2. Then we determine the policy name for each of the Session Policies
                $POLICYNAME = $CAGSESSIONPOLPropertyArray[1];
                
                ## BS 3. Now we find out if this specific Session policy is bound to this specific CAG vServer
                $SESSIONPOLS = Get-StringWithProperty -SearchString $CAGPOLS -Like "bind vpn vserver $vServerDisplayName -policy $POLICYNAME*";

                ## BS 4. If we have an Session Policy bound to this specific CAG vServer then we 
                foreach ($SESSIONPOL in $SESSIONPOLS) {                
                    $SESSIONPOLH += @{
                        Name = Get-StringProperty $SESSIONPOL "-policy" -RemoveQuotes;
                    } # end Hasthable $SESSIONPOLH
                }# end foreach $SESSIONPOLS
            } #end foreach SESSIONPOLS
            
            if ($SESSIONPOLH.Length -gt 0) {
                ## IB - Add the table to the document (only if not null!
                ## IB - Create the parameters to pass to the AddWordTable function
                $Params = $null
                $Params = @{
                    Hashtable = $SESSIONPOLH;
                        Columns = "Name";
                        Headers = "Policy Name";
                    AutoFit = $wdAutoFitContent
                    Format = -235; ## IB - Word constant for Light List Accent 5
                }
                ## IB - Add the table to the document, splatting the parameters
                $Table = AddWordTable @Params -NoGridLines;

                FindWordDocumentEnd;
            } else { WriteWordLine 0 0 "No Session Policy has been configured"} #endif SESSIONPOLHASH.Length
        } #end if no Session policy configures
WriteWordLine 0 0 " "
    #endregion CAG Session Policies 
        
    #region CAG URL Bookmarks   
       
        WriteWordLine 3 0 "URL Bookmarks "

        if($CAGURLPOLS.Length -le 0) { WriteWordLine 0 0 "No URL Bookmark has been configured"} else { 
            $URLPOLH = $null
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $URLPOLH = @();

            ## BS 1. First we get all URL  Policies on the NetScaler system
            foreach ($CAGURLPOL in $CAGURLPOLS) {
                $CAGURLPOLPropertyArray = Get-StringPropertySplit -SearchString ($CAGURLPOL -Replace 'add vpn' ,'') ;
                $CAGURLPOLPropertyArrayNoQuotes = Get-StringPropertySplit -SearchString ($CAGURLPOL -Replace 'add vpn' ,'') -RemoveQuotes ;
                ## BS 2. Then we determine the policy name for each of the URL Policies
                $POLICYNAME = $CAGURLPOLPropertyArray[1];
                
                ## BS 3. Now we find out if this specific URL policy is bound to this specific CAG vServer
                $URLPOLS = Get-StringWithProperty -SearchString $CAGPOLS -Like "bind vpn vserver $vServerDisplayName -urlName $POLICYNAME*";

                $TEXTTODISPLAY = $CAGURLPOLPropertyArrayNoQuotes[2]
                $URL = $CAGURLPOLPropertyArrayNoQuotes[3]

                ## BS 4. If we have an URL Policy bound to this specific CAG vServer then we 
                foreach ($URLPOL in $URLPOLS) {                
                    $URLPOLH += @{
                        Name = Get-StringProperty $URLPOL "-urlName" -RemoveQuotes;
                        Text = $TEXTTODISPLAY;
                        URL = $URL;
                        CLIENTLESSACCESS = Get-StringProperty $CAGURLPOL "-clientlessAccess" "Off" -RemoveQuotes;
                        Comment = Get-StringProperty $CAGURLPOL "-comment" "No Comment" -RemoveQuotes;
                    } # end Hasthable $URLPOLH
                }# end foreach $SESSIONPOLS
            } #end foreach SESSIONPOLS
            if ($URLPOLH.Length -gt 0) {
                ## IB - Add the table to the document (only if not null!
                ## IB - Create the parameters to pass to the AddWordTable function
                $Params = $null
                $Params = @{
                    Hashtable = $URLPOLH;
                        Columns = "Name","Text","URL","CLIENTLESSACCESS","Comment";
                        Headers = "Policy Name","Text to display","URL","Clientless Access","Comment";
                    AutoFit = $wdAutoFitContent
                    Format = -235; ## IB - Word constant for Light List Accent 5
                }
                ## IB - Add the table to the document, splatting the parameters
                $Table = AddWordTable @Params -NoGridLines;

                FindWordDocumentEnd;
            } else { WriteWordLine 0 0 "No URL has been configured"} #endif SESSIONPOLHASH.Length
        } #end if no Session policy configures
    WriteWordLine 0 0 " "
    #endregion CAG URL Policies 

    #region CAG STA Configuration

    WriteWordLine 3 0 "Secure Ticket Authority Configuration"
    $vServerSTAs = Get-StringWithProperty -SearchString $BINDVPNVSERVER -Like "bind vpn vserver $vServerDisplayName -staServer*";    
    if($vServerSTAs.Length -gt 0) {
        $vServerSTAH = $null
        
        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $vServerSTAH = @();
 
        foreach ($vServerSTA in $vServerSTAs) {
        $vServerSTAH += @{
            Name = Get-StringProperty $vServerSTA "-staServer" -RemoveQuotes;
            } 
        }

        if ($vServerSTAH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $vServerSTAH;
                    Columns = "Name";
                    Headers = "Security Ticket Authority";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
        } else {
            WriteWordLine 0 0 "No specific Secure Ticket Authority has been configured for this virtual server"
        }
    WriteWordLine 0 0 " "
    } # if($vServerSTAs.Length
    #endregion CAG STA Configuration
#endregion CAG vServer policies
          
        $selection.InsertNewPage()
    } #end foreach AccessGateway
} #end if Accessgateway.Length

#endregion CAG vServers

#region CAG Policies
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters Citrix NetScaler (Access) Gateway Policies"
WriteWordLine 1 0 "NetScaler Gateway Policies"

#region CAG Session Policies
WriteWordLine 0 0 " "
WriteWordLine 2 0 "NetScaler Gateway Session Policies"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `tNetScaler Gateway Session Policies"

$CurrentRowIndex = 0;

ForEach ($CAGSESSIONACT in $CAGSESSIONACTS) {
    $CurrentRowIndex++
    WriteWordLine 3 0 "NetScaler Gateway Session Policy: $(Get-StringProperty $CAGSESSIONACT "sessionAction" -RemoveQuotes)";
    Write-Verbose "$(Get-Date): `t`tNetScaler Gateway $CurrentRowIndex/$($CAGSESSIONACTS.Length) : $(Get-StringProperty $CAGSESSIONACT "sessionAction" -RemoveQuotes)";
#region Security
    
    WriteWordLine 4 0 "Security"

    ## IB - Create an array of hashtables to store our columns. Note: If we need the
    ## IB - headers to include spaces we can override these at table creation time.
    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = @{
            ## IB - Each hashtable is a separate row in the table!
            DEFAUTH = Get-StringProperty $CAGSESSIONACT "-defaultAuthorizationAction" "DENY";
            CLISEC = Get-StringProperty $CAGSESSIONACT "-encryptCsecExp" "Disabled";
            SECBRW = Get-StringProperty $CAGSESSIONACT "-SecureBrowse" "Disabled";
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
            ICAPROXY = Get-StringProperty $CAGSESSIONACT "-icaProxy" "Global Configuration";
            WIADDR = Get-StringProperty $CAGSESSIONACT "-wihome" "Global Configuration" -RemoveQuotes;
            WIMODE = Get-StringProperty $CAGSESSIONACT "-wiPortalMode" "Global Configuration";
            SSO = Get-StringProperty $CAGSESSIONACT "-ntDomain" "Global Configuration";
            HOME = Get-StringProperty $CAGSESSIONACT "-citrixReceiverHome" "Global Configuration";
            ACCSVC = Get-StringProperty $CAGSESSIONACT "-storefronturl" "Global Configuration";
        }
        Columns = "ICAPROXY","WIADDR","WIMODE","SSO","HOME","ACCSVC";
        Headers = "ICA Proxy","Web Interface addres","Web Interface Portal Mode","Single Sign-On Domain","Citrix Receiver Home Page","Account Services Address";
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

#endregion CAG Session Policies

#endregion CAG Policies

#endregion Citrix NetScaler Gateway

#region NetScaler Monitors
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Monitors"

WriteWordLine 1 0 "NetScaler Custom Monitors"

Write-Verbose "$(Get-Date): `t`tTable: Write NetScaler Monitors Table"

if($MONITORS.Length -le 0) { WriteWordLine 0 0 "No Custom Monitor has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $MONITORSH = @();

    foreach ($MONITOR in $MONITORS) {

        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $MONITORPropertyArray = Get-StringPropertySplit -SearchString ($MONITOR -Replace 'add lb monitor ' ,'') -RemoveQuotes;
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!

        $MONITORSH += @{
                NAME = $MONITORPropertyArray[0];
                Protocol = $MONITORPropertyArray[1];
                HTTPRequest = Get-StringProperty $MONITOR "-httpRequest" "NA";
                DestinationIP = Get-StringProperty $MONITOR "-destIP" "NA";
                DestinationPort = Get-StringProperty $MONITOR "-destPort" "NA";
                Interval = Get-StringProperty $MONITOR "-interval" "NA";
                ResponseCode = Get-StringProperty $MONITOR "-respCode" "NA";
                TimeOut = Get-StringProperty $MONITOR "-resptimeout" "NA";
                SitePath = Get-StringProperty $MONITOR "-sitePath" "NA";
                }
            }

        if ($MONITORSH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $MONITORSH;
                Columns = "NAME","Protocol","HTTPRequest","DestinationIP","DestinationPort","Interval","ResponseCode","TimeOut","SitePath";
                Headers = "Monitor Name","Protocol","HTTP Request","Destination IP","Destination Port","Interval","Response Code","Time-Out","SitePath";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
        }

$selection.InsertNewPage()

#endregion NetScaler Monitors

#region NetScaler Policies
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Policies"

WriteWordLine 1 0 "NetScaler Policies"

## Work in Progress: Binding to actions and binding to vServers

#Policy Pattern Set
WriteWordLine 2 0 "NetScaler Custom Pattern Set Policies"

Write-Verbose "$(Get-Date): `tTable: NetScaler Custom Pattern Set Policies"
$PATSET1 = Get-StringWithProperty -SearchString $Add -Like 'add policy patset *';

if ($PATSET1 -eq $null -or $PATSET1.Length -le 0) {
        WriteWordLine 0 0 "No Custom Pattern Set Policy has been configured"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $PATSETS = @();

            foreach ($PAT in $PATSET1) {
                $Y = ($PAT -replace 'add policy patset ', '').split()
                $PATSETS += @{ "Pattern Set Policy" = "$Y"; }
                } # end foreach

            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $PATSETS;
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
                }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            }
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "

WriteWordLine 2 0 "NetScaler Custom Responder Policies"

Write-Verbose "$(Get-Date): `tTable: NetScaler Custom Responder Policies"
$POLICY = Get-StringWithProperty -SearchString $Add -Like 'add responder policy *';

if ($POLICY -eq $null -or $POLICY.Length -le 0) {
        WriteWordLine 0 0 "No Custom Responder Policy has been configured"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $POLICIESH = @();

            ## IB - Iterate over all load balancer bindings (uses new function)
            foreach ($POL in $POLICY) {
                $Y = Get-StringPropertySplit $POL �RemoveQuotes
                $POLICIESH += @{ "Responder Policy" = $Y[3]; }
                } # end foreach

            ## IB - Create the parameters to pass to the AddWordTable function
            $PARAMS = $null
            $Params = @{
                Hashtable = $POLICIESH;
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
                }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            }
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "

WriteWordLine 2 0 "NetScaler Custom Rewrite Policies"

Write-Verbose "$(Get-Date): `tTable: NetScaler Custom Rewrite Policies"

$POLRW = Get-StringWithProperty -SearchString $Add -Like 'add rewrite policylabel *';

if ($POLRW -eq $null -or $POLRW.Length -le 0) {
        WriteWordLine 0 0 "No Custom Rewrite Policy has been configured"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $POLRWH = @();

            foreach ($POL in $POLRW) {
                $Y = Get-StringPropertySplit $POL �RemoveQuotes
                $POLRWH += @{ "Rewrite Policy" = $Y[3]; }
                } # end foreach

            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $POLRWH;
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
                }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;
            }
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "

$selection.InsertNewPage()

#endregion NetScaler Policies

#region NetScaler Actions
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Actions"

WriteWordLine 1 0 "NetScaler Actions"

## Work in Progress: Binding to policies

WriteWordLine 2 0 "NetScaler Custom Pattern Set Action"

Write-Verbose "$(Get-Date): `tTable: NetScaler Custom Pattern Set Action"
$ACTPATSET1 = Get-StringWithProperty -SearchString $Add -Like 'add action patset *';

if ($ACTPATSET1 -eq $null -or $ACTPATSET1.Length -le 0) {
        WriteWordLine 0 0 "No Custom Pattern Set Action has been configured"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $ACTPATSETS = @();

            foreach ($ACTPAT in $ACTPATSET1) {
                $Y = ($ACTPAT -replace 'add action patset ', '').split()
                $ACTPATSETS += @{ "Pattern Set Policy" = "$Y"; }
                } # end foreach

            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $ACTPATSETS;
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
                }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            }
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "

WriteWordLine 0 0 " "

WriteWordLine 2 0 "NetScaler Responder Action"

Write-Verbose "$(Get-Date): `tTable: NetScaler Custom Responder Action"
$ACTRES = Get-StringWithProperty -SearchString $Add -Like 'add responder action *';

if ($ACTRES -eq $null -or $ACTRES.Length -le 0) {
        WriteWordLine 0 0 "No Custom Responder Action has been configured"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $ACTRESH = @();

            ## IB - Iterate over all load balancer bindings (uses new function)
            foreach ($POL in $ACTRES) {
                $Y = Get-StringPropertySplit $POL �RemoveQuotes                
                $ACTRESH += @{ 
                    Responder = $Y[3]; 
                    Rule = $Y[4];
                    Undefined = $Y[5];
                    }
                } # end foreach

            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $ACTRESH;
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
                }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            }
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "

WriteWordLine 2 0 "NetScaler Custom Rewrite Action"

Write-Verbose "$(Get-Date): `tTable: NetScaler Custom Rewrite Action"
$ACTRW = Get-StringWithProperty -SearchString $Add -Like 'add rewrite action *';

if ($ACTRW -eq $null -or $ACTRW.Length -le 0) {
        WriteWordLine 0 0 "No Custom Rewrite Policy has been configured"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $ACTRWH = @();

            ## IB - Iterate over all load balancer bindings (uses new function)
            foreach ($POL in $ACTRW) {
                $Y = Get-StringPropertySplit $POL �RemoveQuotes
                $ACTRWH += @{ 
                    Rewrite = $Y[3]; 
                    Rule = $Y[4];
                    Undefined = $Y[5];
                    }
                } # end foreach

            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $ACTRWH;
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
                }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            }
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "

$selection.InsertNewPage()

#endregion NetScaler Actions

#region NetScaler Profiles
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Profiles"

WriteWordLine 1 0 "NetScaler Profiles"

WriteWordLine 2 0 "NetScaler Custom TCP Profiles"

Write-Verbose "$(Get-Date): `t`tTable: Write NetScaler TCP Profiles Table"

$TCPPROFILES = Get-StringWithProperty -SearchString $Set -Like 'set ns tcpProfile*';

if($TCPPROFILES.Length -le 0) { WriteWordLine 0 0 "No Custom TCP Profiles has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $TCPPROFILESH = @();

    foreach ($TCPPROFILE in $TCPPROFILES) {

        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $TCPPROFILEPropertyArray = Get-StringPropertySplit -SearchString ($TCPPROFILE -Replace 'set ns tcpProfile' ,'') -RemoveQuotes;
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!

        $TCPPROFILESH += @{
                TCP = $TCPPROFILEPropertyArray[0];
                WS = Get-StringProperty $TCPPROFILE "-WS" "NA";
                SACK = Get-StringProperty $TCPPROFILE "-SACK" "NA";
                NAGLE = Get-StringProperty $TCPPROFILE "-NAGLE" "NA";
                MSS = Get-StringProperty $TCPPROFILE "-MSS" "NA";
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
        }
    
WriteWordLine 2 0 "NetScaler Custom HTTP Profiles"

Write-Verbose "$(Get-Date): `t`tTable: Write NetScaler HTTP Profiles Table"

$HTTPPROFILES = Get-StringWithProperty -SearchString $Add -Like 'add ns httpProfile*';

if($HTTPPROFILES.Length -le 0) { WriteWordLine 0 0 "No Custom HTTP Profiles has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $HTTPPROFILESH = @();

    foreach ($HTTPPROFILE in $HTTPPROFILES) {
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $HTTPPROFILEPropertyArray = Get-StringPropertySplit -SearchString ($HTTPPROFILE -Replace 'add ns httpProfile' ,'') -RemoveQuotes;
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!

        $HTTPPROFILESH += @{
                HTTP = $HTTPPROFILEPropertyArray[0];
                Drop = Get-StringProperty $HTTPPROFILE "-dropInvalReqs" "Disabled";
                SPDY = Get-StringProperty $HTTPPROFILE "-spdy" "Disabled";
            }
        }
        if ($HTTPPROFILESH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $HTTPPROFILESH;
                Columns = "HTTP","Drop","SPDY";
                Headers = "HTTP Profile","Drop Invalid Connections","SPDY";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
        }
$selection.InsertNewPage()

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
#endregion script template 2
# SIG # Begin signature block
# MIIgAAYJKoZIhvcNAQcCoIIf8TCCH+0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUZtrAfu+2DoKbuD2/BzFlIq7h
# +i6gghtnMIIDtzCCAp+gAwIBAgIQDOfg5RfYRv6P5WD8G/AwOTANBgkqhkiG9w0B
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
# /AeB+ova+YJJ92JuoVP6EpQYhS6SkepobEQysmah5xikmmRR7zCCBTUwggQdoAMC
# AQICEA9d61FpHWqo9d26u4syvuEwDQYJKoZIhvcNAQELBQAwcjELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2ln
# bmluZyBDQTAeFw0xNDEwMTQwMDAwMDBaFw0xNTEwMTkxMjAwMDBaMHwxCzAJBgNV
# BAYTAlVTMQswCQYDVQQIEwJUTjESMBAGA1UEBxMJVHVsbGFob21hMSUwIwYDVQQK
# ExxDYXJsIFdlYnN0ZXIgQ29uc3VsdGluZywgTExDMSUwIwYDVQQDExxDYXJsIFdl
# YnN0ZXIgQ29uc3VsdGluZywgTExDMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIB
# CgKCAQEAspvOyrygBQYA7knTpuj540TxFQ7GcC8GNR4kcjQtWsNdPRmP4id4h70e
# BmFgdI1DM5xvZgKG32ULLAxW8Trhucn3Au+Zyjt01Hsc0HIvti3Zeuqzzahjz3Um
# AHwhLOXeVDCx049X7G4EmTncSUDwfIWoiTcAoIXqpinVHZ2rdRApNQOag2j7zbK/
# piK/0aPTS1t6Uf3Glu/wYYCZmCUxYX3LkyOitpexrdiFW33ZloRW8A5efJxMrDr3
# G2wYSYQkYFjOfoYwTIRW4pqbG29EemTMLW9sydA85jX0XECtkV4WSioydjq4+pJm
# TF0bAPTF2cC3GR4MyOdeM8ehzJx1FQIDAQABo4IBuzCCAbcwHwYDVR0jBBgwFoAU
# WsS5eyoKo6XqcQPAYPkt9mV1DlgwHQYDVR0OBBYEFD6/rjYxtxvJ2+E96CA0EIM7
# Maw+MA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzB3BgNVHR8E
# cDBuMDWgM6Axhi9odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVk
# LWNzLWcxLmNybDA1oDOgMYYvaHR0cDovL2NybDQuZGlnaWNlcnQuY29tL3NoYTIt
# YXNzdXJlZC1jcy1nMS5jcmwwQgYDVR0gBDswOTA3BglghkgBhv1sAwEwKjAoBggr
# BgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzCBhAYIKwYBBQUH
# AQEEeDB2MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wTgYI
# KwYBBQUHMAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFNI
# QTJBc3N1cmVkSURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMBAf8EAjAAMA0GCSqG
# SIb3DQEBCwUAA4IBAQBS64rxnZ6S6z3wL2LEi4s9wjKQfPSWMi/tmj8pLm3I82vy
# DckX4p9Nwsh8T1k1PvPh37Q3HquoIHtdEZBFYfDjAwtWl9GFzS5gZrMHfdnlBO1b
# dZw2vx6+qHEuy+9jVjtndIJPYOtf1FrpuOvY5Ya+idd5wHXfrJXVS95WmLCVCuCe
# Jv6mWclemL3S5t0aX9NgRVH7jcfjvh4jcoFjSMvt3irnJZBZ0a18+3CQPNoa6UC9
# QIuVZH0Oq2RvtPpQwCcl2onBqDHOphX/rD4lmR7KubYcQS910Uxmpv03KvfWpYP6
# a+lM+UzUz7zC340f/jiJvmpX1ZXQjpI2InbFeu4zMIIGajCCBVKgAwIBAgIQAwGa
# Ajr/WLFr1tXq5hfwZjANBgkqhkiG9w0BAQUFADBiMQswCQYDVQQGEwJVUzEVMBMG
# A1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEw
# HwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTEwHhcNMTQxMDIyMDAwMDAw
# WhcNMjQxMDIyMDAwMDAwWjBHMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNl
# cnQxJTAjBgNVBAMTHERpZ2lDZXJ0IFRpbWVzdGFtcCBSZXNwb25kZXIwggEiMA0G
# CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCjZF38fLPggjXg4PbGKuZJdTvMbuBT
# qZ8fZFnmfGt/a4ydVfiS457VWmNbAklQ2YPOb2bu3cuF6V+l+dSHdIhEOxnJ5fWR
# n8YUOawk6qhLLJGJzF4o9GS2ULf1ErNzlgpno75hn67z/RJ4dQ6mWxT9RSOOhkRV
# fRiGBYxVh3lIRvfKDo2n3k5f4qi2LVkCYYhhchhoubh87ubnNC8xd4EwH7s2AY3v
# J+P3mvBMMWSN4+v6GYeofs/sjAw2W3rBerh4x8kGLkYQyI3oBGDbvHN0+k7Y/qpA
# 8bLOcEaD6dpAoVk62RUJV5lWMJPzyWHM0AjMa+xiQpGsAsDvpPCJEY93AgMBAAGj
# ggM1MIIDMTAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8E
# DDAKBggrBgEFBQcDCDCCAb8GA1UdIASCAbYwggGyMIIBoQYJYIZIAYb9bAcBMIIB
# kjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzCCAWQG
# CCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMA
# IABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMA
# IABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMA
# ZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkA
# bgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgA
# IABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUA
# IABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAA
# cgBlAGYAZQByAGUAbgBjAGUALjALBglghkgBhv1sAxUwHwYDVR0jBBgwFoAUFQAS
# KxOYspkH7R7for5XDStnAs0wHQYDVR0OBBYEFGFaTSS2STKdSip5GoNL9B6Jwcp9
# MH0GA1UdHwR2MHQwOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydEFzc3VyZWRJRENBLTEuY3JsMDigNqA0hjJodHRwOi8vY3JsNC5kaWdpY2Vy
# dC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDB3BggrBgEFBQcBAQRrMGkw
# JAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcw
# AoY1aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElE
# Q0EtMS5jcnQwDQYJKoZIhvcNAQEFBQADggEBAJ0lfhszTbImgVybhs4jIA+Ah+WI
# //+x1GosMe06FxlxF82pG7xaFjkAneNshORaQPveBgGMN/qbsZ0kfv4gpFetW7ea
# sGAm6mlXIV00Lx9xsIOUGQVrNZAQoHuXx/Y/5+IRQaa9YtnwJz04HShvOlIJ8Oxw
# YtNiS7Dgc6aSwNOOMdgv420XEwbu5AO2FKvzj0OncZ0h3RTKFV2SQdr5D4HRmXQN
# JsQOfxu19aDxxncGKBXp2JPlVRbwuwqrHNtcSCdmyKOLChzlldquxC5ZoGHd2vNt
# omHpigtt7BIYvfdVVEADkitrwlHCCkivsNRu4PQUCjob4489yq9qjXvc2EQwggbN
# MIIFtaADAgECAhAG/fkDlgOt6gAK6z8nu7obMA0GCSqGSIb3DQEBBQUAMGUxCzAJ
# BgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5k
# aWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBD
# QTAeFw0wNjExMTAwMDAwMDBaFw0yMTExMTAwMDAwMDBaMGIxCzAJBgNVBAYTAlVT
# MRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5j
# b20xITAfBgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMTCCASIwDQYJKoZI
# hvcNAQEBBQADggEPADCCAQoCggEBAOiCLZn5ysJClaWAc0Bw0p5WVFypxNJBBo/J
# M/xNRZFcgZ/tLJz4FlnfnrUkFcKYubR3SdyJxArar8tea+2tsHEx6886QAxGTZPs
# i3o2CAOrDDT+GEmC/sfHMUiAfB6iD5IOUMnGh+s2P9gww/+m9/uizW9zI/6sVgWQ
# 8DIhFonGcIj5BZd9o8dD3QLoOz3tsUGj7T++25VIxO4es/K8DCuZ0MZdEkKB4YNu
# gnM/JksUkK5ZZgrEjb7SzgaurYRvSISbT0C58Uzyr5j79s5AXVz2qPEvr+yJIvJr
# GGWxwXOt1/HYzx4KdFxCuGh+t9V3CidWfA9ipD8yFGCV/QcEogkCAwEAAaOCA3ow
# ggN2MA4GA1UdDwEB/wQEAwIBhjA7BgNVHSUENDAyBggrBgEFBQcDAQYIKwYBBQUH
# AwIGCCsGAQUFBwMDBggrBgEFBQcDBAYIKwYBBQUHAwgwggHSBgNVHSAEggHJMIIB
# xTCCAbQGCmCGSAGG/WwAAQQwggGkMDoGCCsGAQUFBwIBFi5odHRwOi8vd3d3LmRp
# Z2ljZXJ0LmNvbS9zc2wtY3BzLXJlcG9zaXRvcnkuaHRtMIIBZAYIKwYBBQUHAgIw
# ggFWHoIBUgBBAG4AeQAgAHUAcwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQByAHQA
# aQBmAGkAYwBhAHQAZQAgAGMAbwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBjAGUA
# cAB0AGEAbgBjAGUAIABvAGYAIAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAgAEMA
# UAAvAEMAUABTACAAYQBuAGQAIAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQAGEA
# cgB0AHkAIABBAGcAcgBlAGUAbQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBtAGkA
# dAAgAGwAaQBhAGIAaQBsAGkAdAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBjAG8A
# cgBwAG8AcgBhAHQAZQBkACAAaABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBlAHIA
# ZQBuAGMAZQAuMAsGCWCGSAGG/WwDFTASBgNVHRMBAf8ECDAGAQH/AgEAMHkGCCsG
# AQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29t
# MEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8v
# Y3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqg
# OKA2hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURS
# b290Q0EuY3JsMB0GA1UdDgQWBBQVABIrE5iymQftHt+ivlcNK2cCzTAfBgNVHSME
# GDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEARlA+
# ybcoJKc4HbZbKa9Sz1LpMUerVlx71Q0LQbPv7HUfdDjyslxhopyVw1Dkgrkj0bo6
# hnKtOHisdV0XFzRyR4WUVtHruzaEd8wkpfMEGVWp5+Pnq2LN+4stkMLA0rWUvV5P
# sQXSDj0aqRRbpoYxYqioM+SbOafE9c4deHaUJXPkKqvPnHZL7V/CSxbkS3BMAIke
# /MV5vEwSV/5f4R68Al2o/vsHOE8Nxl2RuQ9nRc3Wg+3nkg2NsWmMT/tZ4CMP0qqu
# AHzunEIOz5HXJ7cW7g/DvXwKoO4sCFWFIrjrGBpN/CohrUkxg0eVd3HcsRtLSxwQ
# nHcUwZ1PL1qVCCkQJjGCBAMwggP/AgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAv
# BgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EC
# EA9d61FpHWqo9d26u4syvuEwCQYFKw4DAhoFAKBAMBkGCSqGSIb3DQEJAzEMBgor
# BgEEAYI3AgEEMCMGCSqGSIb3DQEJBDEWBBQG7Wvdjgf8RqTUmYetM/tB2Kgo3zAN
# BgkqhkiG9w0BAQEFAASCAQCunbmgccGMrn2ShfwTMcd+2jYCSt2Qxi8llPQeO3r6
# EIVmurnuic2fLp4BwuN/6Vj+Ko0krfHKpPtJj/n6T40QvHl1CyFZectuV1xQuWCy
# lp2y5PR/e8O6LTgB6KWQ65AHE5U0HULQMr5/nRCgCLdFk/vjTL+Q0PvKvEFTnI25
# dkA8Z7QoPMStp6Fq2SAqmP+SMIrnUjbWLMZVMoqTbHiOSEJ02xxkiyiCwZdOIlJq
# gDHJW7LR3PUS3A3RO+cC8PDSg7ZpZDs5cJn5hlWMgZS++nKjUVBjfj4lY4zkGB6H
# Pr/pCDhKQUnTtCpiG5RFw5u4iYw174Wux32VyIQhWkkooYICDzCCAgsGCSqGSIb3
# DQEJBjGCAfwwggH4AgEBMHYwYjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGln
# aUNlcnQgQXNzdXJlZCBJRCBDQS0xAhADAZoCOv9YsWvW1ermF/BmMAkGBSsOAwIa
# BQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0x
# NDEyMTcwMDQ4NDVaMCMGCSqGSIb3DQEJBDEWBBRo/GpjCLN1UDdh2lZx9lZJqW2f
# zTANBgkqhkiG9w0BAQEFAASCAQCFaIp3Ui+MwAiX9sCDje3I/uIH9eSuB3IND3FQ
# oNSCsjZNxD7nhFl+jKYMuvmOm+2ojzJUCaxA9HtRqXABTNdi9xNTXezB3FhyMt5w
# C1NPHe193rx8cRyxVqpSLZq2zdMyfxrVF4aZVaKm8Etdv9sVJQbynGpl+dlNSYMK
# f0XQXLg4berwxBR4tq7haMg8vOSjmeeSBW7GD/ff4tMJQrjW1RGVQyopn8w4h1Gq
# 2Ba9XoRv/zbmQmw9Mqk5r3ecaJ9abWe+fX9GxBsLNa1MjOBM9jGhc9FhZWKo4wW/
# pUY7JspVa6PfOzfMYzjMWum8CT7wU1w4RO/V11BEI5LRrqlI
# SIG # End signature block
