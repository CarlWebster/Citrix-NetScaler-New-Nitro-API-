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
.PARAMETER CompanyName
	Company Name to use for the Cover Page.  
	Default value is contained in HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated on the 
	computer running the script.
	This parameter has an alias of CN.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	(default cover pages in Word en-US)
	Valid input is:
		Alphabet (Word 2007/2010. Works)
		Annual (Word 2007/2010. Doesn't really work well for this report)
		Austere (Word 2007/2010. Works)
		Austin (Word 2010/2013. Doesn't work in 2013, mostly works in 2007/2010 but Subtitle/Subject & Author fields need to me moved after title box is moved up)
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
	Default value is Motion.
	This parameter has an alias of CP.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	The PDF file is roughly 5X to 10X larger than the DOCX file.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_v1.ps1
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Motion for the Cover Page format.
	Administrator for the User Name.
	The computer running the script for the AdminAddress.
.EXAMPLE
	PS C:\PSScript > .\NetScaler_v1.ps1 -PDF
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Motion for the Cover Page format.
	Administrator for the User Name.
	The computer running the script for the AdminAddress.
.EXAMPLE
	PS C:\PSScript .\NetScaler_V1.ps1 -CompanyName "Carl Webster Consulting" -CoverPage "Mod" -UserName "Carl Webster" -AdminAddress DDC01

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
		Controller named DDC01 for the AdminAddress.
.EXAMPLE
	PS C:\PSScript .\NetScaler_V1.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
		The computer running the script for the AdminAddress.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word document.
.NOTES
	NAME: NetScaler_V1.ps1
	VERSION: 1.0.1
	AUTHOR: Barry Schiffer
	LASTEDIT: May 12, 2014
#>
#endregion Support

#region Word Setup

#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding( SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Standard" ) ]

Param(
	[parameter(ParameterSetName="Standard",
	Position = 0, 
	Mandatory=$false )
	] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(ParameterSetName="Standard",
	Position = 1, 
	Mandatory=$false )
	] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(ParameterSetName="Standard",
	Position = 2, 
	Mandatory=$false )
	] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

	[parameter(ParameterSetName="Standard",
	Position = 3, 
	Mandatory=$false )
	] 
	[Switch]$PDF=$False)
	
#Version 1.0.1 script May 12, 2014
#Minor bug fix release
#	Load Balancer: Changed the chapter name "Services" to "Services and Service Groups". Thanks to Carl Behrent for the heads up!
#	Authentication Local Groups: Changed logic for the Group Name. Thanks to Erik Spicer for the heads up!
#	Script will no longer terminate if the CompanyName registry key is empty and the CompanyName parameter is not used
#	Warning and Error messages are now offset so they are more easily seen and read
#Known Issue
#	Authentication Local Groups: It will sometimes report an extra -option in the name field. This will be fixed soon.

#Version 1.0 script
#originally released to the Citrix community on May 6, 2014


#force -verbose on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
If($PDF -eq $Null)
{
	$PDF = $False
}

#info@barryschiffer.com
#@BarrySchiffer on Twitter
#http://www.barryschiffer.com
#Created on May 1st, 2014

Set-StrictMode -Version 2

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
[int]$wdSaveFormatPDF = 17
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

[string]$RunningOS = (Get-WmiObject -class Win32_OperatingSystem).Caption

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

Switch ($PSCulture.Substring(0,3))
{
	'ca-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents' = 'Taula automática 2';
			}
		}

	'da-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents' = 'Automatisk tabel 2';
			}
		}

	'de-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents' = 'Automatische Tabelle 2';
			}
		}

	'en-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents'  = 'Automatic Table 2';
			}
		}

	'es-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents' = 'Tabla automática 2';
			}
		}

	'fi-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents' = 'Automaattinen taulukko 2';
			}
		}

	'fr-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents' = 'Sommaire Automatique 2';
			}
		}

	'nb-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents' = 'Automatisk tabell 2';
			}
		}

	'nl-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents' = 'Automatische inhoudsopgave 2';
			}
		}

	'pt-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents' = 'Sumário Automático 2';
			}
		}

	'sv-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents' = 'Automatisk innehållsförteckning2';
			}
		}

	Default	{$hash.('en-US') = @{
				'Word_TableOfContents'  = 'Automatic Table 2';
			}
		}
}

$myHash = $hash.$PSCulture

If($myHash -eq $Null)
{
	$myHash = $hash.('en-US')
}

$myHash.Word_NoSpacing = $wdStyleNoSpacing
$myHash.Word_Heading1 = $wdStyleheading1
$myHash.Word_Heading2 = $wdStyleheading2
$myHash.Word_Heading3 = $wdStyleheading3
$myHash.Word_Heading4 = $wdStyleheading4
$myHash.Word_TableGrid = $wdTableGrid

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP)
	
	$xArray = ""
	
	Switch ($PSCulture.Substring(0,3))
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana", "Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)", "Sector (fosc)", "Semàfor", "Visualització", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador", "Contrast", "Cubicles", "Diplomàtic", "Exposició",
					"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari", "Perspectiva", "Piles", "Quadrícula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabet", "Anual", "Conservador", "Contrast", "Cubicles", "Diplomàtic", "En mosaic",
					"Exposició", "Línia lateral", "Mod", "Moviment", "Piles", "Sobri", "Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran", "Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)", "Udsnit (mørk)", "Ion (mørk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter", "Overskrid", "Alfabet", "Kontrast", "Stakke",
					"Fliser", "Gåde", "Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel", "Nålestribet", "Årlig", "Avispapir", "Tradionel")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabet", "Årlig", "BevægElse", "Eksponering", "Enkel", "Firkanter", "Fliser", "Gåde", "Kontrast",
					"Mod", "Nålestribet", "Overskrid", "Sidelinje", "Stakke", "Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)", "Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung", "Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend", "Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie", "Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alphabet", "Bewegung", "Durchscheinend", "Herausgestellt", "Jährlich", "Kacheln", "Kontrast",
					"Kubistisch", "Modern", "Nadelstreifen", "Puzzle", "Randlinie", "Raster", "Schlicht", "Stapel", "Traditionell")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid", "Integral", "Ion (Dark)", "Ion (Light)", "Motion",
					"Retrospect", "Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative", "Contrast", "Cubicles", "Exposure", "Grid",
					"Mod", "Motion", "Newsprint", "Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Conservative", "Contrast", "Cubicles", "Exposure", "Mod", "Motion",
					"Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin", "Slice (luz)", "Faceta", "Semáforo",
					"Retrospectiva", "Cuadrícula", "Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador", "Contraste", "Cuadrícula",
					"Cubículos", "Exposición", "Línea lateral", "Moderno", "Mosaicos", "Movimiento", "Papel periódico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Conservador", "Contraste", "Cubículos", "Exposición",
					"Línea lateral", "Moderno", "Mosaicos", "Movimiento", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)", "Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin", "Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti", "Laatikot", "Liike", "Liituraita", "Mod",
					"Osittain peitossa", "Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko", "Ruudut", "Sanomalehtipaperi",
					"Sivussa", "Vuotuinen", "Ylitys")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Aakkoset", "Alttius", "Kontrasti" ,"Kuvakkeet ja tiedot" ,"Liike" ,"Liituraita" ,"Mod" ,"Palapeli",
					"Perinteinen", "Pinot", "Sivussa", "Työpisteet", "Vuosittainen", "Yksinkertainen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("ViewMaster","Secteur (foncé)","Sémaphore","Rétrospective","Ion (foncé)","Ion (clair)","Intégrale",
					"Filigrane","Facette","Secteur (clair)","À bandes", "Austin", "Guide", "Whisp", "Lignes latérales", "Quadrillage")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Mosaïques", "Ligne latérale", "Annuel", "Perspective", "Contraste", "Emplacements de bureau",
					"Moderne","Blocs empilés", "Rayures fines", "Austère", "Transcendant", "Classique", "Quadrillage", "Exposition",
					"Alphabet", "Mots croisés", "Papier journal", "Austin","Guide")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alphabet", "Annuel", "Austère", "Blocs empilés", "Blocs superposés", "Classique", "Contraste",
					"Exposition","Guide", "Ligne latérale", "Moderne", "Mosaïques", "Mots croisés", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran", "Integral", "Ion (lys)", "Ion (mørk)",
					"Retrospekt", "Rutenett", "Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker", "BevegElse", "Engasjement", "Enkel", "Fliser",
					"Konservativ", "Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje", "Smale striper", "Stabler",
					"Transcenderende")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabet", "Årlig", "Avlukker", "BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Puslespill", "Sidelinje", "Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept", "Integraal", "Ion (donker)", "Ion (licht)",
					"Raster", "Segment (Light)", "Semafoor", "Slice (donker)", "Spriet", "Terugblik", "Terzijde", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden", "Beweging", "Blikvanger", "Contrast", "Eenvoudig",
					"Jaarlijks", "Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief", "Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Bescheiden", "Beweging", "Blikvanger", "Contrast", "Eenvoudig",
					"Jaarlijks", "Krijtstreep", "Mod", "Puzzel", "Stapels", "Tegels", "Terzijde", "Werkplekken")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra", "Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete",
					"Filigrana", "Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral", "Retrospectiva", "Semáforo")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias", "Conservador", "Contraste", "Exposição",
					"Grade", "Ladrilhos", "Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas", "Quebra-cabeça", "Transcend")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Baias", "Conservador", "Contraste", "Exposição",
					"Ladrilhos", "Linha Lateral", "Listras", "Mod", "Pilhas", "Quebra-cabeça", "Transcendente")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)", "Jon (mörkt)", "Knippe", "Rutnät",
					"RörElse", "Sektor (ljus)", "Sektor (mörk)", "Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt", "Kontrast", "Kritstreck", "Kuber",
					"Perspektiv", "Plattor", "Pussel", "Rutnät", "RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
					"Övergående")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabetmönster", "Årligt", "Enkelt", "Exponering", "Konservativt", "Kontrast", "Kritstreck",
					"Kuber", "Övergående", "Plattor", "Pussel", "RörElse", "Sidlinje", "Sobert", "Staplat")
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid", "Integral", "Ion (Dark)", "Ion (Light)", "Motion",
						"Retrospect", "Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster", "Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative", "Contrast", "Cubicles", "Exposure", "Grid",
						"Mod", "Motion", "Newsprint", "Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
					ElseIf($xWordVersion -eq $wdWord2007)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Conservative", "Contrast", "Cubicles", "Exposure", "Mod", "Motion",
						"Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
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
		Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
		Exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}) -ne $Null
	If($wordrunning)
	{
		Write-Host "`n`n`t`tPlease close all instances of Microsoft Word before running this report.`n`n"
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
	[string]$fontName=$null,
	[int]$fontSize=0,
	[bool]$italics=$false,
	[bool]$boldface=$false,
	[Switch]$nonewline)
	
	#Build output style
	[string]$output = ""
	Switch ($style)
	{
		0 {$Selection.Style = $myHash.Word_NoSpacing}
		1 {$Selection.Style = $myHash.Word_Heading1}
		2 {$Selection.Style = $myHash.Word_Heading2}
		3 {$Selection.Style = $myHash.Word_Heading3}
		4 {$Selection.Style = $myHash.Word_Heading4}
		Default {$Selection.Style = $myHash.Word_NoSpacing}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Selection.TypeParagraph()
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
	$Word.quit()
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | out-null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	Exit
}

[string]$Title = "NetScaler Configuration - $Companyname"
[string]$filename1 = "NetScaler.docx"
If($PDF)
{
	[string]$filename2 = "NetScaler.pdf"
}

CheckWordPreReq
$script:startTime = Get-Date

Write-Verbose "$(Get-Date): Setting up Word"

# Setup word for output
Write-Verbose "$(Get-Date): Create Word comObject.  If you are not running Word 2007, ignore the next message."
$Word = New-Object -comobject "Word.Application" -EA 0

If(!$? -or $Word -eq $Null)
{
	Write-Warning "The Word object could not be created.  You may need to repair your Word installation."
	Write-Error "`n`n`t`tThe Word object could not be created.`n`n`t`tYou may need to repair your Word installation.`n`n`t`tScript cannot continue.`n`n"
	Exit
}

[int]$WordVersion = [int] $Word.Version
If($WordVersion -eq $wdWord2013)
{
	$WordProduct = "Word 2013"
}
ElseIf($WordVersion -eq $wdWord2010)
{
	$WordProduct = "Word 2010"
}
ElseIf($WordVersion -eq $wdWord2007)
{
	$WordProduct = "Word 2007"
}
Else
{
	Write-Error "`n`n`t`tYou are running an untested or unsupported version of Microsoft Word.`n`n`t`tScript will end.`n`n`t`tPlease send info on your version of Word to webster@carlwebster.com`n`n"
	AbortScript
}

Write-Verbose "$(Get-Date): Running Microsoft $WordProduct"

If($PDF -and $WordVersion -eq $wdWord2007)
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

Write-Verbose "$(Get-Date): Validate company name"
#only validate CompanyName if the field is blank
If([String]::IsNullOrEmpty($CompanyName))
{
	$CompanyName = ValidateCompanyName
	If([String]::IsNullOrEmpty($CompanyName))
	{
		Write-Warning "`n`n`t`tCompany Name is blank so Cover Page will not show a Company Name."
		Write-Warning "`n`t`tCheck HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value."
		Write-Warning "`n`t`tYou may want to use the -CompanyName parameter is you need a Company Name on the cover page.`n`n"
	}
}

Write-Verbose "$(Get-Date): Check Default Cover Page for language specific version"
[bool]$CPChanged = $False
Switch ($PSCulture.Substring(0,3))
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
				If($WordVersion -eq $wdWord2013)
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

Write-Verbose "$(Get-Date): Validate cover page"
[bool]$ValidCP = ValidateCoverPage $WordVersion $CoverPage
If(!$ValidCP)
{
	Write-Error "`n`n`t`tFor $WordProduct, $CoverPage is not a valid Cover Page option.`n`n`t`tScript cannot continue.`n`n"
	AbortScript
}

Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): Company Name : $CompanyName"
Write-Verbose "$(Get-Date): Cover Page   : $CoverPage"
Write-Verbose "$(Get-Date): User Name    : $UserName"
Write-Verbose "$(Get-Date): Save As PDF  : $PDF"
Write-Verbose "$(Get-Date): Title        : $Title"
Write-Verbose "$(Get-Date): Filename1    : $filename1"
If($PDF)
{
	Write-Verbose "$(Get-Date): Filename2    : $filename2"
}
Write-Verbose "$(Get-Date): OS Detected  : $RunningOS"
Write-Verbose "$(Get-Date): PSUICulture  : $PSUICulture"
Write-Verbose "$(Get-Date): PSCulture    : $PSCulture "
Write-Verbose "$(Get-Date): Word version : $WordProduct"
Write-Verbose "$(Get-Date): Word language: $($Word.Language)"
Write-Verbose "$(Get-Date): PoSH version : $($Host.Version)"
Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): Script start : $($Script:StartTime)"
Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): "

$Word.Visible = $False

#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
#using Jeff's Demo-WordReport.ps1 file for examples
#down to $configlog = $False is from Jeff Hicks
Write-Verbose "$(Get-Date): Load Word Templates"

[bool]$CoverPagesExist = $False
[bool]$BuildingBlocksExist = $False

$word.Templates.LoadBuildingBlocks()
If($WordVersion -eq $wdWord2007)
{
	$BuildingBlocks = $word.Templates | Where {$_.name -eq "Building Blocks.dotx"}
}
Else
{
	#word 2010/2013
	$BuildingBlocks = $word.Templates | Where {$_.name -eq "Built-In Building Blocks.dotx"}
}

Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"

If($BuildingBlocks -ne $Null)
{
	$BuildingBlocksExist = $True

	Try 
		{$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)}

	Catch
		{$part = $Null}

	If($part -ne $Null)
	{
		$CoverPagesExist = $True
	}
}

#cannot continue if cover page does not exist
If(!$CoverPagesExist)
{
	Write-Verbose "$(Get-Date): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
	Write-Error "`n`n`t`tCover Pages are not installed or the Cover Page $($CoverPage) does not exist.`n`n`t`tScript cannot continue.`n`n"
	Write-Verbose "$(Get-Date): Closing Word"
	AbortScript
}

Write-Verbose "$(Get-Date): Create empty word doc"
$Doc = $Word.Documents.Add()
If($Doc -eq $Null)
{
	Write-Verbose "$(Get-Date): "
	Write-Error "`n`n`t`tAn empty Word document could not be created.`n`n`t`tScript cannot continue.`n`n"
	AbortScript
}

$Selection = $Word.Selection
If($Selection -eq $Null)
{
	Write-Verbose "$(Get-Date): "
	Write-Error "`n`n`t`tAn unknown error happened selecting the entire Word document for default formatting options.`n`n`t`tScript cannot continue.`n`n"
	AbortScript
}

#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
#36 = .50"
$Word.ActiveDocument.DefaultTabStop = 36

#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
Write-Verbose "$(Get-Date): Disable grammar and spell checking"
#bug reported 1-Apr-2014 by Tim Mangan
#save current options first before turning them off
$CurrentGrammarOption = $Word.Options.CheckGrammarAsYouType
$CurrentSpellingOption = $Word.Options.CheckSpellingAsYouType
$Word.Options.CheckGrammarAsYouType = $False
$Word.Options.CheckSpellingAsYouType = $False

If($BuildingBlocksExist)
{
	#insert new page, getting ready for table of contents
	Write-Verbose "$(Get-Date): Insert new page, getting ready for table of contents"
	$part.Insert($selection.Range,$True) | out-null
	$selection.InsertNewPage()

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
		$toc.insert($selection.Range,$True) | out-null
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
$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
#get the footer and format font
$footers = $doc.Sections.Last.Footers
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
$selection.HeaderFooter.Range.Text = $footerText

#add page numbering
Write-Verbose "$(Get-Date): Add page numbering"
$selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

#return focus to main document
Write-Verbose "$(Get-Date): Return focus to main document"
$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

#move to the end of the current document
Write-Verbose "$(Get-Date): Move to the end of the current document"
Write-Verbose "$(Get-Date)"
$selection.EndKey($wdStory,$wdMove) | Out-Null
#end of Jeff Hicks

#endregion Word Setup

#region NetScaler Documentation Build

#region NetScaler Documentation Functions

##############################################################################
#.SYNOPSIS
# Get named property from a string.
#
#.DESCRIPTION
# Returns a case-insensitive property from a string, assuming the property is
# named before the actual property value and is separated by a space. For
# example, if the specified SearchString contained
# "-property1 <value1> -property2 <value2>�, searching # for "-Property1"
# would return "<value1>".
##############################################################################
function Get-StringProperty([string]$SearchString, [string]$SearchProperty, [string]$EmptyString = "Undefined")
{
    # Locate and replace quotes with '^^' and quoted spaces '^' to aid with parsing, until there are none left
    While ($SearchString.Contains('"')) {
        # Store the right-hand side temporarily, skipping the first quote
        $SearchStringRight = $SearchString.Substring($SearchString.IndexOf('"') +1);
        # Extract the quoted text from the original string
        $QuotedString = $SearchString.Substring($SearchString.IndexOf('"'), $SearchStringRight.IndexOf('"') +2);
        # Replace the quoted text, replacing spaces with '^' and quotes with '^^'
        $SearchString = $SearchString.Replace($QuotedString, $QuotedString.Replace(" ", "^").Replace('"', "^^"));
    }

    # Split the $SearchString based on one or more blank spaces
    $StringComponents = $SearchString.Split(' +',[StringSplitOptions]'RemoveEmptyEntries'); 
    For($i = 0; $i -le $StringComponents.Length; $i++) {
        # The standard Powershell CompareTo method is case-sensitive
        If([string]::Compare($StringComponents[$i], $SearchProperty, $True) -eq 0) {
            # Check that we're not over the array boundary
            If($i+1 -le $StringComponents.Length) {
                # Restore any escaped quotation marks and spaces
                # If you wanted to trim quotation marks you could use this instead:
                #  return $StringComponents[$i+1].Replace("^^", '"').Replace("^", " ").Trim('"');
                return $StringComponents[$i+1].Replace("^^", '"').Replace("^", " ");
            }
        }
    }
    # If nothing has been found or we're over the array boundary, return the $EmptyString value
    return $EmptyString;
}

##############################################################################
#.SYNOPSIS
# Tests whether a named property in a string exists
#
#.DESCRIPTION
# Returns a boolean value if a string property is present. For example, if
# the specified SearchString contained "-property1 -property2 <value2>�,
# searching for "-Property1" or "-Property2" would return true, but searching
# for "-Property3" would return false
##############################################################################
function Test-StringProperty([string]$SearchString, [string]$SearchProperty)
{
    # Split the $SearchString based on one or more blank spaces
    $StringComponents = $SearchString.Split(' +',[StringSplitOptions]'RemoveEmptyEntries'); 
    for ($i = 0; $i -le $StringComponents.Length; $i++) {
        # The standard Powershell CompareTo method is case-sensitive
        If([string]::Compare($StringComponents[$i], $SearchProperty, $True) -eq 0) { return $true; }
    }
    # No value found so return false
    return $false;
}

##############################################################################
#.SYNOPSIS
# Tests whether a named property in a string exists and returns either Yes
# ($true) or No ($false)
##############################################################################
function Test-StringPropertyYesNo([string]$SearchString, [string]$SearchProperty)
{
    if (Test-StringProperty $SearchString $SearchProperty) { return "Yes"; }
    else { return "No"; }
}

##############################################################################
#.SYNOPSIS
# Tests whether a named property in a string exists and returns either Yes
# ($false) or No ($true)
##############################################################################
function Test-NotStringPropertyYesNo([string]$SearchString, [string]$SearchProperty)
{
    if (-not (Test-StringProperty $SearchString $SearchProperty)) { return "Yes"; }
    else { return "No"; }
}

##############################################################################
#.SYNOPSIS
# Tests whether a named property in a string exists and returns either Enabled
# ($true) or Disabled ($false)
##############################################################################
function Test-StringPropertyEnabledDisabled([string]$SearchString, [string]$SearchProperty)
{
    if (Test-StringProperty $SearchString $SearchProperty) { return "Enabled"; }
    else { return "Disabled"; }
}

##############################################################################
#.SYNOPSIS
# Tests whether a named property in a string exists and returns either Disabled
# ($true) or Enabled ($false)
##############################################################################
function Test-NotStringPropertyEnabledDisabled([string]$SearchString, [string]$SearchProperty)
{
    if (-not (Test-StringProperty $SearchString $SearchProperty)) { return "Enabled"; }
    else { return "Disabled"; }
}

##############################################################################
#.SYNOPSIS
# Tests whether a named property in a string exists and returns either On
# ($true) or Off ($false)
##############################################################################
function Test-StringPropertyOnOff([string]$SearchString, [string]$SearchProperty)
{
    if (Test-StringProperty $SearchString $SearchProperty) { return "On"; }
    else { return "Off"; }
}

##############################################################################
#.SYNOPSIS
# Tests whether a named property in a string exists and returns either Off
# ($true) or On ($false)
##############################################################################
function Test-NotStringPropertyOnOff([string]$SearchString, [string]$SearchProperty)
{
    if (-not (Test-StringProperty $SearchString $SearchProperty)) { return "On"; }
    else { return "Off"; }
}

#endregion NetScaler Documentation Functions

#region NetScaler Documentation Setup

$Scriptver = 1
$SourceFileName = "ns.conf";

## Iain Brighton - Try and resolve the ns.conf file in the current working directory
if (Test-Path (Join-Path ((Get-Location).ProviderPath) $SourceFileName)) {
    $SourceFile = Join-Path ((Get-Location).ProviderPath) $SourceFileName; }
else {
    ## Otherwise try the script's directory
    if (Test-Path (Join-Path (Split-Path $MyInvocation.MyCommand.Path) $SourceFileName)) {
        $SourceFile = Join-Path (Split-Path $MyInvocation.MyCommand.Path) $SourceFileName; }
    else {
        throw "Cannot locate a NetScaler ns.conf file in either the working or script directory."; }
}

Write-Verbose "$(Get-Date): NetScaler file : $SourceFile"

## Iain Brighton - Set the output locations to the current working directory
$filename1 = Join-Path ((Get-Location).ProviderPath) $filename1;
if ($PDF) { 
    $filename2 = Join-Path ((Get-Location).ProviderPath) $filename2;
    Write-Verbose "$(Get-Date): Target Word file : $filename1, PDF file : $filename2";
    }
    else { Write-Verbose "$(Get-Date): Target Word file : $filename1"; }

## We read the file in once as each Get-Content call goes to disk and also creates a new string[]
$File = Get-Content $SourceFile

## Create collections for faster processing of ns conf.
$Set = $File | Where { $_ -Like "Set *" }
$SetNS = $Set | Where { $_ -Like "Set ns *" }
$Add = $File | Where { $_ -Like "Add *" }
$Bind = $File | Where { $_ -Like "Bind *" }
$Enable = $File | Where { $_ -Like "Enable *" }
$ContentSwitch = $Add | Where { $_ -Like "add cs vserver *" }
$Loadbalancer = $Add | Where { $_ -Like "add lb vserver *" }
$LoadbalancerBind = $Bind | Where { $_ -Like "bind lb vserver *" }
$ServiceGroup = $Add | Where { $_ -Like "add servicegroup *" }
$Service = $Add | Where { $_ -Like "add service *" }
$ServiceBind = $Bind | Where { $_ -Like "bind service *" }
$Server = $Add | Where { $_ -Like "add server *" }
$Monitor = $Add | Where { $_ -Like "add lb monitor *" }
$IPList = $Add | Where { $_ -Like "add ns ip *" }

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

## Barry Schiffer add chapter counter
$Chapter = 0
$Chapters = 26

## Barry Schiffer Use Stopwatch class to time script execution
$sw = [Diagnostics.Stopwatch]::StartNew()

$selection.InsertNewPage()
<#
WriteWordLine 1 0 "Support statement"

WriteWordLine 0 0 "This is version $Scriptver of the Citrix NetScaler documentation script"
WriteWordLine 0 0 " "

WriteWordLine 0 0 "This version of the script does not offer support for the following Citrix NetScaler Features but are on our priority list to be added in the next version"
WriteWordLine 0 0 "Citrix NetScaler Gateway"
WriteWordLine 0 0 "Authentication - Radius / TACACS and Cert"
WriteWordLine 0 0 "SNMP Traps / Groups / Users / Views"
WriteWordLine 0 0 "AppFlow Configuration"
WriteWordLine 0 0 "Network Configration includes Traffic Domain / vLAN / IP / Interfaces, all else not yet" 
WriteWordLine 0 0 "Citrix Web Interface configuration, we do however document install state"
WriteWordLine 0 0 "Load Balancing Persistency Groups, all else is documented"
WriteWordLine 0 0 "CloudBridge Connector"
WriteWordLine 0 0 "AAA Configuration"
WriteWordLine 0 0 "NTP Servers"
WriteWordLine 0 0 " "

WriteWordLine 0 0 "This version of the script does not offer support for the following Citrix NetScaler Features and will be added only if requested. Use e-mail address info@barryschiffer.com for requests"
WriteWordLine 0 0 " "
WriteWordLine 0 0 "Cache redirection Configuration"
WriteWordLine 0 0 "GSLB Configuration"
WriteWordLine 0 0 "Integrated Caching"
WriteWordLine 0 0 "Application Firewall Configuration"
WriteWordLine 0 0 "Protection Features Configuration"
WriteWordLine 0 0 "Database Profiles and Users"
WriteWordLine 0 0 "Auditing Configuration"
WriteWordLine 0 0 "AppExpert, we do document the result of loaded App Templates like Content Switches and Load Balancers"
#>
#endregion NetScaler Documentation Setup

#region NetScaler System Information
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler System Information"

$selection.InsertNewPage()
WriteWordLine 1 0 "NetScaler System Information"
WriteWordLine 2 0 "NetScaler Version"

## NetScaler Version
$TableRange = $doc.Application.Selection.Range
[int]$Columns = 2
[int]$Rows = 3
Write-Verbose "$(Get-Date): `t`tTable: Processing NetScaler Version"
$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$table.Style = "Table Grid"
$table.Borders.InsideLineStyle = $wdLineStyleNone
$table.Borders.OutsideLineStyle = $wdLineStyleSingle
$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,1).Range.Font.Bold = $True
$Table.Cell(1,1).Range.Text = "NetScaler Version"
$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,2).Range.Font.Bold = $True
$Table.Cell(1,2).Range.Text = " "	

$Table.Cell(2,1).Range.Font.size = 9
$Table.Cell(2,1).Range.Text = "Major Version"
$Table.Cell(3,1).Range.Font.size = 9
$Table.Cell(3,1).Range.Text = "Build"
$Table.Cell(2,2).Range.Font.size = 9
$Table.Cell(2,2).Range.Text = "$Version"
$Table.Cell(3,2).Range.Font.size = 9
$Table.Cell(3,2).Range.Text = "$Build"
 
$table.AutoFitBehavior($wdAutoFitContent)        

#return focus back to document
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

#move to the end of the current document
$selection.EndKey($wdStory,$wdMove) | Out-Null

WriteWordLine 0 0 " "

#endregion NetScaler System Information

#region NetScaler IP
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler IP"

## NetScaler IP

WriteWordLine 2 0 "NetScaler Management IP Address"

$TableRange = $doc.Application.Selection.Range
[int]$Columns = 2
[int]$Rows = 3
$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$table.Style = "Table Grid"
$table.Borders.InsideLineStyle = $wdLineStyleNone
$table.Borders.OutsideLineStyle = $wdLineStyleSingle
$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,1).Range.Font.Bold = $True
$Table.Cell(1,1).Range.Text = "NetScaler IP Address"
$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,2).Range.Font.Bold = $True
$Table.Cell(1,2).Range.Text = " "	

$SetNS | foreach {
   if ($_ -like 'set ns config -IPAddress *') {
      $Table.Cell(2,1).Range.Font.size = 9
	  $Table.Cell(2,1).Range.Text = "NetScaler IP"
      $Table.Cell(3,1).Range.Font.size = 9
	  $Table.Cell(3,1).Range.Text = "Subnet"
      $Table.Cell(2,2).Range.Font.size = 9
	  $Table.Cell(2,2).Range.Text = Get-StringProperty $_ "-IPAddress";
      $Table.Cell(3,2).Range.Font.size = 9
	  $Table.Cell(3,2).Range.Text = Get-StringProperty $_ "-netmask";
    }
}     
$table.AutoFitBehavior($wdAutoFitContent)        

#return focus back to document
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

#move to the end of the current document
$selection.EndKey($wdStory,$wdMove) | Out-Null

WriteWordLine 0 0 " "

#endregion NetScaler IP

#region NetScaler Time Zone
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Time zone"

WriteWordLine 2 0 "NetScaler Time Zone"

$TableRange = $doc.Application.Selection.Range
[int]$Columns = 1
[int]$Rows = 2
$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$table.Style = "Table Grid"
$table.Borders.InsideLineStyle = $wdLineStyleNone
$table.Borders.OutsideLineStyle = $wdLineStyleSingle
$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,1).Range.Font.Bold = $True
$Table.Cell(1,1).Range.Text = "NetScaler Time zone"

$Setns | foreach {  
    if ($_ -like 'set ns param *') {
        $Table.Cell(2,1).Range.Font.size = 9
	    $Table.Cell(2,1).Range.Text = Get-StringProperty $_ "-timezone";
        }
    }

$table.AutoFitBehavior($wdAutoFitContent)        

#return focus back to document
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

#move to the end of the current document
$selection.EndKey($wdStory,$wdMove) | Out-Null

WriteWordLine 0 0 " "

#endregion NetScaler Time Zone

#region NetScaler Global HTTP Parameters
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Global HTTP Parameters"

WriteWordLine 2 0 "NetScaler Global HTTP Parameters"

$TableRange = $doc.Application.Selection.Range
[int]$Columns = 2
[int]$Rows = 2

$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$table.Style = "Table Grid"
$table.Borders.InsideLineStyle = $wdLineStyleNone
$table.Borders.OutsideLineStyle = $wdLineStyleSingle
$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,1).Range.Font.Bold = $True
$Table.Cell(1,1).Range.Text = "Cookie Version"
$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,2).Range.Font.Bold = $True
$Table.Cell(1,2).Range.Text = "HTTP Drop Invalid Request"

$TestCookie = 0
$Table.Cell(2,1).Range.Font.size = 9
$Setns | foreach {  
    if ($_ -like 'set ns param *') {
        $TestCookie = 1
	    $Table.Cell(2,1).Range.Text = Get-StringProperty $_ "-cookieversion" "0";
        }
    }

if ($TestCookie -eq 0) { $Table.Cell(2,1).Range.Text = "0" }

$TestDrop = 0
$Table.Cell(2,1).Range.Font.size = 9
$Setns | foreach {  
    if ($_ -like 'set ns httpParam *') {
        $TestDrop = 1
	    $Table.Cell(2,2).Range.Text = Test-StringPropertyOnOff $_ "-dropInvalReqs";
        }
    }
if ($TestDrop -eq 0) { $Table.Cell(2,2).Range.Text = "Off" }

$table.AutoFitBehavior($wdAutoFitContent)        

#return focus back to document
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

#move to the end of the current document
$selection.EndKey($wdStory,$wdMove) | Out-Null

WriteWordLine 0 0 " "

#endregion NetScaler Global HTTP Parameters

#region NetScaler Global TCP Parameters
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Global TCP Parameters"

WriteWordLine 2 0 "NetScaler Global TCP Parameters"

$TableRange = $doc.Application.Selection.Range
[int]$Columns = 3
[int]$Rows = 2
$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$table.Style = "Table Grid"
$table.Borders.InsideLineStyle = $wdLineStyleNone
$table.Borders.OutsideLineStyle = $wdLineStyleSingle
$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,1).Range.Font.Bold = $True
$Table.Cell(1,1).Range.Text = "TCP Windows Scaling"
$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,2).Range.Font.Bold = $True
$Table.Cell(1,2).Range.Text = "Selective Acknowledgement"
$Table.Cell(1,3).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,3).Range.Font.Bold = $True
$Table.Cell(1,3).Range.Text = "Use Nagle's Algorithm"

$TestTCPParam = 0
$Table.Cell(2,1).Range.Font.size = 9
$Table.Cell(2,2).Range.Font.size = 9
$Table.Cell(2,3).Range.Font.size = 9
$Setns | foreach {  
    if ($_ -like 'set ns tcpParam *') {
	    $TestTCPParam = 1
        $Table.Cell(2,1).Range.Text = Test-StringPropertyEnabledDisabled $_ "-WS";
	    $Table.Cell(2,2).Range.Text = Test-StringPropertyEnabledDisabled $_ "-SACK";
	    $Table.Cell(2,3).Range.Text = Test-StringPropertyEnabledDisabled $_ "-nagle";
        }
}
if ($TestTCPParam -eq 0) { $Table.Cell(2,1).Range.Text = "Disabled" }
if ($TestTCPParam -eq 0) { $Table.Cell(2,2).Range.Text = "Disabled" }
if ($TestTCPParam -eq 0) { $Table.Cell(2,3).Range.Text = "Disabled" }

$table.AutoFitBehavior($wdAutoFitContent)        

#return focus back to document
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

#move to the end of the current document
$selection.EndKey($wdStory,$wdMove) | Out-Null

WriteWordLine 0 0 " "

#endregion NetScaler Global TCP Parameters

#region NetScaler Global Diameter Parameters
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Global Diameter Parameter"

WriteWordLine 2 0 "NetScaler Global Diameter Parameters"

$TableRange = $doc.Application.Selection.Range
[int]$Columns = 2
[int]$Rows = 2
$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$table.Style = "Table Grid"
$table.Borders.InsideLineStyle = $wdLineStyleNone
$table.Borders.OutsideLineStyle = $wdLineStyleSingle
$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,1).Range.Font.Bold = $True
$Table.Cell(1,1).Range.Text = "Host Identity"
$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,2).Range.Font.Bold = $True
$Table.Cell(1,2).Range.Text = "Realm"

$Setns | foreach {  
    if ($_ -like 'set ns diameter *') {
        $Table.Cell(2,1).Range.Font.size = 9
	    $Table.Cell(2,1).Range.Text = Get-StringProperty $_ "-identity" "NA";
        $Table.Cell(2,2).Range.Font.size = 9
	    $Table.Cell(2,2).Range.Text = Get-StringProperty $_ "-realm" "NA";
        }
}

$table.AutoFitBehavior($wdAutoFitContent)        

#return focus back to document
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

#move to the end of the current document
$selection.EndKey($wdStory,$wdMove) | Out-Null

WriteWordLine 0 0 " "

#endregion NetScaler Global Diameter Parameters

#region NetScaler Management vLAN
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler management vLAN"

WriteWordLine 2 0 "NetScaler Management vLAN"

$Rows = 1

$SetNS | foreach { 
    if ($_ -like 'set ns config -nsvlan *') {
        $Rows++
    }
}

If ($rows -eq 1) {WriteWordLine 0 0 "No Management vLAN has been assigned"} Else {

    $TableRange = $doc.Application.Selection.Range
    [int]$Columns = 2
    [int]$Rows = 3
    $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
    $table.Style = "Table Grid"
	$table.Borders.InsideLineStyle = $wdLineStyleNone
	$table.Borders.OutsideLineStyle = $wdLineStyleSingle

    $ROWC = 0
    do {
        $ROWC++
        $Table.Cell($ROWC,1).Shading.BackgroundPatternColor = $wdColorGray15
        $Table.Cell($ROWC,1).Range.Font.Bold = $True
        $Table.Cell($ROWC,1).Range.Font.size = 9
        }
    while ($ROWC -le $ROWS -1)

	$Table.Cell(1,1).Range.Text = "vLAN ID"
	$Table.Cell(2,1).Range.Text = "Interface"
	$Table.Cell(3,1).Range.Text = "Tagged"
    
    $ROWC1 = 0
    do {
        $ROWC1++
        $Table.Cell($ROWC1,2).Range.Font.size = 9
        }
    while ($ROWC1 -le $ROWS -1)

    $SetNS | foreach { 
       if ($_ -like 'set ns config -nsvlan *') {
	      $Table.Cell(1,2).Range.Text = Get-StringProperty $_ "-nsvlan";
	      $Table.Cell(2,2).Range.Text = Get-StringProperty $_ "-ifnum";
	      $Table.Cell(3,2).Range.Text = Get-StringProperty $_ "-tagged";
          }
    }
    $table.AutoFitBehavior($wdAutoFitContent)        

    #return focus back to document
    $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

    #move to the end of the current document
    $selection.EndKey($wdStory,$wdMove) | Out-Null
    WriteWordLine 0 0 "Table: Management vLAN "
}
WriteWordLine 0 0 " "

#endregion NetScaler Management vLAN

#region NetScaler High Availability
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters High Availibility"

WriteWordLine 2 0 "NetScaler High Availibility"

$Rows = 1
$SetNS | foreach {  
    if ($_ -like 'set ns rpcNode *') {
        $Rows++
        }
}

If ($ROWS -eq 2) {WriteWordLine 0 0 "High availability is not Configured"} Else {
    $TableRange = $doc.Application.Selection.Range
    [int]$Columns = 2
    [int]$Rows = $Rows
    $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
    $table.Style = "Table Grid"
	$table.Borders.InsideLineStyle = $wdLineStyleNone
	$table.Borders.OutsideLineStyle = $wdLineStyleSingle
    $xRow = 1
    $ROWCOUNT = $Rows + 1
    do {
        $xRow++
        $Table.Cell($xRow,1).Range.Font.size = 9
        $Table.Cell($xRow,2).Range.Font.size = 9
        }
    while ($xRow -le $ROWCOUNT)

    $Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
    $Table.Cell(1,1).Range.Font.Bold = $True
    $Table.Cell(1,1).Range.Text = "NetScaler Node"
    $Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
    $Table.Cell(1,2).Range.Font.Bold = $True
    $Table.Cell(1,2).Range.Text = "Source IP"
    $xRow = 1

    $Setns | foreach {  
        if ($_ -like 'set ns rpcNode *') {
            $xRow++
            $Y = ($_ -replace 'set ns rpcNode ', '').split()
            $Table.Cell($xRow,1).Range.Text = $Y[0];
            $Table.Cell($xRow,2).Range.Text = Get-StringProperty $_ "-srcIP";
             }
        }
    #return focus back to document
    $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

    #move to the end of the current document
    $selection.EndKey($wdStory,$wdMove) | Out-Null

    WriteWordLine 0 0 ""
    }

WriteWordLine 0 1 " "

#endregion NetScaler High Availability

#region NetScaler Features
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Features"

$selection.InsertNewPage()

WriteWordLine 2 0 "NetScaler Features"

If ($Version -gt $ScriptVersion) {
    WriteWordLine 0 0 ""
    WriteWordLine 0 0 "Warning: You are using Citrix NetScaler version $Version, features added since version $ScriptVersion will not be shown."
    WriteWordLine 0 0 ""
    }

WriteWordLine 3 0 "NetScaler Basic Features"

## Features

$TableRange = $doc.Application.Selection.Range
[int]$Columns = 2
[int]$Rows = 11

$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$table.Style = "Table Grid"
$table.Borders.InsideLineStyle = $wdLineStyleNone
$table.Borders.OutsideLineStyle = $wdLineStyleSingle

$xRow = 1
$ROWCOUNT = $Rows + 1
do {
    $xRow++
    $Table.Cell($xRow,1).Range.Font.size = 9
    $Table.Cell($xRow,2).Range.Font.size = 9
    If($xRow % 2 -eq 0) {
	    $Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray05
        $Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray05
        }
    }
while ($xRow -le $ROWCOUNT)

$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,1).Range.Font.Bold = $True
$Table.Cell(1,1).Range.Font.size = 11
$Table.Cell(1,1).Range.Text = "Feature"
$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,2).Range.Font.Bold = $True
$Table.Cell(1,2).Range.Font.size = 11
$Table.Cell(1,2).Range.Text = "State"

$Table.Cell(2,1).Range.Text = "Application Firewall"
$Table.Cell(2,2).Range.Text = $FEATAppFw
$Table.Cell(3,1).Range.Text = "Authentication, Authorization and Auditing"
$Table.Cell(3,2).Range.Text = $FEATAAA
$Table.Cell(4,1).Range.Text = "Content Filter"
$Table.Cell(4,2).Range.Text = $FEATCF
$Table.Cell(5,1).Range.Text = "Content Switching"
$Table.Cell(5,2).Range.Text = $FEATCS
$Table.Cell(6,1).Range.Text = "HTTP Compression"
$Table.Cell(6,2).Range.Text = $FEATCMP
$Table.Cell(7,1).Range.Text = "Integrated Caching"
$Table.Cell(7,2).Range.Text = $FEATIC
$Table.Cell(8,1).Range.Text = "Load Balancing"
$Table.Cell(8,2).Range.Text = $FEATLB
$Table.Cell(9,1).Range.Text = "NetScaler Gateway"
$Table.Cell(9,2).Range.Text = $FEATSSLVPN
$Table.Cell(10,1).Range.Text = "Rewrite"
$Table.Cell(10,2).Range.Text = $FEATRewrite
$Table.Cell(11,1).Range.Text = "SSL Offloading"
$Table.Cell(11,2).Range.Text = $FEATSSL

$table.AutoFitBehavior($wdAutoFitContent)
         
#return focus back to document
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

#move to the end of the current document
$selection.EndKey($wdStory,$wdMove) | Out-Null

WriteWordLine 0 0 ""

#endregion NetScaler Features

#region NetScaler Advanced Features
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Advanced Features"

WriteWordLine 3 0 "NetScaler Advanced Features"
  
## Features

$TableRange = $doc.Application.Selection.Range
[int]$Columns = 2
[int]$Rows = 23

$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$table.Style = "Table Grid"
$table.Borders.InsideLineStyle = $wdLineStyleNone
$table.Borders.OutsideLineStyle = $wdLineStyleSingle

$xRow = 1
do {
    $xRow++
    $Table.Cell($xRow,1).Range.Font.size = 9
    $Table.Cell($xRow,2).Range.Font.size = 9
    If($xRow % 2 -eq 0) {
	    $Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray05
        $Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray05
        }
    }
while ($xRow -le $ROWS)

$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,1).Range.Font.Bold = $True
$Table.Cell(1,1).Range.Font.size = 11
$Table.Cell(1,1).Range.Text = "Feature"
$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,2).Range.Font.Bold = $True
$Table.Cell(1,2).Range.Font.size = 11
$Table.Cell(1,2).Range.Text = "State"

$Table.Cell(2,1).Range.Text = "Web Logging"
$Table.Cell(2,2).Range.Text = $FEATWL
$Table.Cell(3,1).Range.Text = "Surge Protection"
$Table.Cell(3,2).Range.Text = $FEATSP
$Table.Cell(4,1).Range.Text = "Cache Redirection"
$Table.Cell(4,2).Range.Text = $FEATCR
$Table.Cell(5,1).Range.Text = "Sure Connect"
$Table.Cell(5,2).Range.Text = $FEATSC
$Table.Cell(6,1).Range.Text = "Priority Queuing"
$Table.Cell(6,2).Range.Text = $FEATPQ
$Table.Cell(7,1).Range.Text = "Global Server Load Balancing"
$Table.Cell(7,2).Range.Text = $FEATGSLB
$Table.Cell(8,1).Range.Text = "Http DoS Protection"
$Table.Cell(8,2).Range.Text = $FEATHDOSP
$Table.Cell(9,1).Range.Text = "Vpath"
$Table.Cell(9,2).Range.Text = $FEATVpath
$Table.Cell(10,1).Range.Text = "Integrated Caching"
$Table.Cell(10,2).Range.Text = $FEATIC
$Table.Cell(11,1).Range.Text = "OSPF Routing"
$Table.Cell(11,2).Range.Text = $FEATOSPF
$Table.Cell(12,1).Range.Text = "RIP Routing"
$Table.Cell(12,2).Range.Text = $FEATRIP
$Table.Cell(13,1).Range.Text = "BGP Routing"
$Table.Cell(13,2).Range.Text = $FEATBGP
$Table.Cell(14,1).Range.Text = "IPv6 protocol translation "
$Table.Cell(14,2).Range.Text = $FEATIPv6PT
$Table.Cell(15,1).Range.Text = "Responder"
$Table.Cell(15,2).Range.Text = $FEATRESPONDER
$Table.Cell(16,1).Range.Text = "HTML Injection"
$Table.Cell(16,2).Range.Text = $FEATHTMLInjection
$Table.Cell(17,1).Range.Text = "OSPF Routing"
$Table.Cell(17,2).Range.Text = $FEATOSPF
$Table.Cell(18,1).Range.Text = "NetScaler Push"
$Table.Cell(18,2).Range.Text = $FEATPUSH
$Table.Cell(19,1).Range.Text = "AppFlow"
$Table.Cell(19,2).Range.Text = $FEATAppFlow
$Table.Cell(20,1).Range.Text = "CloudBridge"
$Table.Cell(20,2).Range.Text = $FEATCloudBridge
$Table.Cell(21,1).Range.Text = "ISIS Routing"
$Table.Cell(21,2).Range.Text = $FEATISIS
$Table.Cell(22,1).Range.Text = "CallHome"
$Table.Cell(22,2).Range.Text = $FEATCH
$Table.Cell(23,1).Range.Text = "AppQoE"
$Table.Cell(23,2).Range.Text = $FEATAppQoE

$table.AutoFitBehavior($wdAutoFitContent)         
#return focus back to document
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

#move to the end of the current document
$selection.EndKey($wdStory,$wdMove) | Out-Null

WriteWordLine 0 0 ""

#endregion NetScaler Advanced Features

#region NetScaler Modes
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Modes"

WriteWordLine 2 0 "NetScaler Modes"

If ($Version -gt $ScriptVersion) {
    WriteWordLine 0 0 ""
    WriteWordLine 0 0 "Warning: You are using Citrix NetScaler version $Version, modes added since version $ScriptVersion will not be shown."
    WriteWordLine 0 0 ""
    }
  
## Features

$TableRange = $doc.Application.Selection.Range
[int]$Columns = 2
[int]$Rows = 17
Write-Verbose "$(Get-Date): `t`tTable: Write Modes"
$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$table.Style = "Table Grid"
$table.Borders.InsideLineStyle = $wdLineStyleNone
$table.Borders.OutsideLineStyle = $wdLineStyleSingle

$xRow = 1
$ROWCOUNT = $Rows + 1
do {
    $xRow++
    $Table.Cell($xRow,1).Range.Font.size = 9
    $Table.Cell($xRow,2).Range.Font.size = 9
    If($xRow % 2 -eq 0) {
	    $Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray05
        $Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray05
        }
    }
while ($xRow -le $ROWCOUNT)

$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,1).Range.Font.Bold = $True
$Table.Cell(1,1).Range.Text = "Mode"
$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,2).Range.Font.Bold = $True
$Table.Cell(1,2).Range.Text = "State"

$Enable | foreach {  
    if ($_ -like 'enable ns mode *') {
        
        $Table.Cell(2,1).Range.Text = "Fast Ramp"
        If ($_.Contains("FR") -eq "True") {$Table.Cell(2,2).Range.Text = "Enabled"} Else {$Table.Cell(2,2).Range.Text = "Disabled"}
        $Table.Cell(3,1).Range.Text = "Layer 2 mode"
        If ($_.Contains("L2") -eq "True") {$Table.Cell(3,2).Range.Text = "Enabled"} Else {$Table.Cell(3,2).Range.Text = "Disabled"}
        $Table.Cell(4,1).Range.Text = "Use Source IP"
        If ($_.Contains("USIP") -eq "True") {$Table.Cell(4,2).Range.Text = "Enabled"} Else {$Table.Cell(4,2).Range.Text = "Disabled"}
        $Table.Cell(5,1).Range.Text = "Client Keep-alive"
        If ($_.Contains("CKA") -eq "True") {$Table.Cell(5,2).Range.Text = "Enabled"} Else {$Table.Cell(5,2).Range.Text = "Disabled"}
        $Table.Cell(6,1).Range.Text = "TCP Buffering"
        If ($_.Contains("TCPB") -eq "True") {$Table.Cell(6,2).Range.Text = "Enabled"} Else {$Table.Cell(6,2).Range.Text = "Disabled"}
        $Table.Cell(7,1).Range.Text = "MAC-based forwarding"
        If ($_.Contains("MBF") -eq "True") {$Table.Cell(7,2).Range.Text = "Enabled"} Else {$Table.Cell(7,2).Range.Text = "Disabled"}
        $Table.Cell(8,1).Range.Text = "Edge configuration"
        If ($_.Contains("Edge") -eq "True") {$Table.Cell(8,2).Range.Text = "Enabled"} Else {$Table.Cell(8,2).Range.Text = "Disabled"}
        $Table.Cell(9,1).Range.Text = "Use Subnet IP"
        If ($_.Contains("USNIP") -eq "True") {$Table.Cell(9,2).Range.Text = "Enabled"} Else {$Table.Cell(9,2).Range.Text = "Disabled"}
        $Table.Cell(10,1).Range.Text = "Use Layer 3 Mode"
        If ($_.Contains("USNIP") -eq "True") {$Table.Cell(10,2).Range.Text = "Enabled"} Else {$Table.Cell(10,2).Range.Text = "Disabled"}
        $Table.Cell(11,1).Range.Text = "Path MTU Discovery"
        If ($_.Contains("PMTUD") -eq "True") {$Table.Cell(11,2).Range.Text = "Enabled"} Else {$Table.Cell(11,2).Range.Text = "Disabled"}
        $Table.Cell(12,1).Range.Text = "Static Route Advertisement"
        If ($_.Contains("SRADV") -eq "True") {$Table.Cell(12,2).Range.Text = "Enabled"} Else {$Table.Cell(12,2).Range.Text = "Disabled"}
        $Table.Cell(13,1).Range.Text = "Direct Route Advertisement"
        If ($_.Contains("DRADV") -eq "True") {$Table.Cell(13,2).Range.Text = "Enabled"} Else {$Table.Cell(13,2).Range.Text = "Disabled"}
        $Table.Cell(14,1).Range.Text = "Intranet Route Advertisement"
        If ($_.Contains("IRADV") -eq "True") {$Table.Cell(14,2).Range.Text = "Enabled"} Else {$Table.Cell(14,2).Range.Text = "Disabled"}
        $Table.Cell(15,1).Range.Text = "Ipv6 Static Route Advertisement"
        If ($_.Contains("SRADV6") -eq "True") {$Table.Cell(15,2).Range.Text = "Enabled"} Else {$Table.Cell(15,2).Range.Text = "Disabled"}
        $Table.Cell(16,1).Range.Text = "Ipv6 Direct Route Advertisement"
        If ($_.Contains("DRADV6") -eq "True") {$Table.Cell(16,2).Range.Text = "Enabled"} Else {$Table.Cell(16,2).Range.Text = "Disabled"}
        $Table.Cell(17,1).Range.Text = "Bridge BPDUs"
        If ($_.Contains("BridgeBPDUs") -eq "True") {$Table.Cell(17,2).Range.Text = "Enabled"} Else {$Table.Cell(17,2).Range.Text = "Disabled"}
        }
    }

$table.AutoFitBehavior($wdAutoFitContent)
#return focus back to document
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

#move to the end of the current document
$selection.EndKey($wdStory,$wdMove) | Out-Null

$selection.InsertNewPage()

#endregion NetScaler Modes

#region NetScaler Network Configuration
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Network Configuration"

WriteWordLine 1 0 "NetScaler Networking"

WriteWordLine 2 0 "NetScaler IP addresses"

$ROWS = 1
$IPLIST | foreach {
   $ROWS++
   }

$TableRange = $doc.Application.Selection.Range
[int]$Columns = 10
[int]$Rows = $ROWS

Write-Verbose "$(Get-Date): `t`tTable: Write IP table"
$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$table.Style = "Table Grid"
$table.Borders.InsideLineStyle = $wdLineStyleSingle
$table.Borders.OutsideLineStyle = $wdLineStyleSingle

$COL = 0
do {
    $Col++
    $Table.Cell(1,$Col).Shading.BackgroundPatternColor = $wdColorGray15
    $Table.Cell(1,$Col).Range.Font.Bold = $True
    $Table.Cell(1,$Col).Range.Font.size = 9
    }
while ($Col -le $Columns -1)

$Table.Cell(1,1).Range.Text = "IP Address"
$Table.Cell(1,2).Range.Text = "Subnet Mask"
$Table.Cell(1,3).Range.Text = "Type"
$Table.Cell(1,4).Range.Text = "Traffic Domain"    
$Table.Cell(1,5).Range.Text = "Management"
$Table.Cell(1,6).Range.Text = "vServer"
$Table.Cell(1,7).Range.Text = "GUI"
$Table.Cell(1,8).Range.Text = "SNMP"
$Table.Cell(1,9).Range.Text = "Telnet"
$Table.Cell(1,10).Range.Text = "Dynamic Routing"
$xRow = 1

$IPLIST | foreach {
    $xROW++
    $COL = 0
    do {
        $COL++
        $Table.Cell($xRow,$COL).Range.Font.size = 8
        }
    while ($Col -le $Columns -1)
    
    $Y = ($_ -replace 'add ns ip ', '').split()
    $Table.Cell($xRow,1).Range.Text = $Y[0];
    $Table.Cell($xRow,2).Range.Text = $Y[1];
    $Table.Cell($xRow,3).Range.Text = Get-StringProperty $_ "-type" "NA";
    $Table.Cell($xRow,4).Range.Text = Get-StringProperty $_ "-td" "0 (Default)";
    $Table.Cell($xRow,5).Range.Text = Test-StringPropertyEnabledDisabled $_ "-mgmtAccess";
    $Table.Cell($xRow,6).Range.Text = Test-NotStringPropertyEnabledDisabled $_ "-vServer";
    $Table.Cell($xRow,7).Range.Text = Test-StringPropertyEnabledDisabled $_ "-gui";
    $Table.Cell($xRow,8).Range.Text = Test-NotStringPropertyEnabledDisabled $_ "-snmp";
    $Table.Cell($xRow,9).Range.Text = Test-StringPropertyEnabledDisabled $_ "-telnet";
    $Table.Cell($xRow,10).Range.Text = Test-StringPropertyEnabledDisabled $_ "-dynamicRouting";
    }

$table.AutoFitBehavior($wdAutoFitContent)

#return focus back to document
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

#move to the end of the current document
$selection.EndKey($wdStory,$wdMove) | Out-Null

WriteWordLine 0 1 " "
###################################################################
Write-Verbose "$(Get-Date): `Processing NetScaler Interfaces"
WriteWordLine 2 0 "NetScaler Interfaces"

$ROWS = 1
$SET | foreach {
    if ($_ -like 'set interface *') {
        $ROWS++
        }
    }

$TableRange = $doc.Application.Selection.Range
[int]$Columns = 6
[int]$Rows = $ROWS

Write-Verbose "$(Get-Date): `t`tTable: Write Interface table"
$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$table.Style = "Table Grid"
$table.Borders.InsideLineStyle = $wdLineStyleSingle
$table.Borders.OutsideLineStyle = $wdLineStyleSingle

$COL = 0
do {
    $Col++
    $Table.Cell(1,$Col).Shading.BackgroundPatternColor = $wdColorGray15
    $Table.Cell(1,$Col).Range.Font.Bold = $True
    $Table.Cell(1,$Col).Range.Font.size = 9
    }
while ($Col -le $Columns -1)

$Table.Cell(1,1).Range.Text = "Interface ID"
$Table.Cell(1,2).Range.Text = "Interface Type"
$Table.Cell(1,3).Range.Text = "HA Monitoring"
$Table.Cell(1,4).Range.Text = "State"
$Table.Cell(1,5).Range.Text = "Auto Negotiate"
$Table.Cell(1,6).Range.Text = "Tag All vLAN"
$xRow = 1

$SET | foreach {
    if ($_ -like 'set interface *') {
        $xROW++
        $COL = 0
        do {
            $COL++
            $Table.Cell($xRow,$COL).Range.Font.size = 9
            }
        while ($Col -le $Columns -1)
    
        $Y = ($_ -replace 'set interface ', '').split()
        $Table.Cell($xRow,1).Range.Text = Get-StringProperty $_ "-ifnum";
        $Table.Cell($xRow,2).Range.Text = Get-StringProperty $_ "-intftype";
        $Table.Cell($xRow,3).Range.Text = Test-NotStringPropertyOnOff $_ "-haMonitor";
        $Table.Cell($xRow,4).Range.Text = Test-NotStringPropertyOnOff $_ "-state";
        $Table.Cell($xRow,5).Range.Text = Test-NotStringPropertyOnOff $_ "-autoneg";
        $Table.Cell($xRow,6).Range.Text = Test-NotStringPropertyOnOff $_ "-tagall";
        }
    }

$table.AutoFitBehavior($wdAutoFitContent)

#return focus back to document
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

#move to the end of the current document
$selection.EndKey($wdStory,$wdMove) | Out-Null

WriteWordLine 0 0 " "

###################################################################

WriteWordLine 2 0 "NetScaler vLANs"

$ROWS = 1
$Add | foreach {
   if ($_ -like 'add vlan *') {
   $ROWS++
   }
}

$TableRange = $doc.Application.Selection.Range
[int]$Columns = 6
[int]$Rows = $ROWS

Write-Verbose "$(Get-Date): `t`tTable: Write vLAN configuration"
$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$table.Style = "Table Grid"
$table.Borders.InsideLineStyle = $wdLineStyleSingle
$table.Borders.OutsideLineStyle = $wdLineStyleSingle

$COL = 0
do {
    $COL++
    $Table.Cell(1,$COL).Shading.BackgroundPatternColor = $wdColorGray15
    $Table.Cell(1,$COL).Range.Font.Bold = $True
    }
while ($Col -le $Columns -1)

$Table.Cell(1,1).Range.Text = "vLAN"
$Table.Cell(1,2).Range.Text = "Interface"
$Table.Cell(1,3).Range.Text = "Interface"
$Table.Cell(1,4).Range.Text = "Interface"
$Table.Cell(1,5).Range.Text = "Interface"
$Table.Cell(1,6).Range.Text = "Interface"
$xRow = 1         

$Add | foreach {
    if ($_ -like 'add vlan *') {
        $xRow++
        $Z = ($_ -replace 'add vlan ', '').split()
        $Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
        $Table.Cell($xRow,1).Range.Font.size = 9
	    $Table.Cell($xRow,1).Range.Text = $($Z[0])
            
        $VLANBIND = "bind vlan $($Z[0]) *"
        $NR = 2
            $Bind | foreach { 
                if ($_ -like $VLANBIND) {
                    if ($_.Contains('ifnum')){
                        $Table.Cell($xRow,$NR).Range.Font.size = 9
			            $Table.Cell($xRow,$NR).Range.Text = Get-StringProperty $_ "-ifnum";
                        $NR++
                        }
                    }
                }

        If ($NR -le 6) {
            $Table.Cell($xRow,$NR).Range.Font.size = 9
		    $Table.Cell($xRow,$NR).Range.Text = "NA";
            }
        }
    $table.AutoFitBehavior($wdAutoFitContent) 
    }

$table.AutoFitBehavior($wdAutoFitContent)

#return focus back to document
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

#move to the end of the current document
$selection.EndKey($wdStory,$wdMove) | Out-Null

WriteWordLine 0 1 " "

########################################################################

## Routing table
## It seems that the auto generated routes are not in the ns.conf. Need to build this based on other information like SNIP

WriteWordLine 2 0 "NetScaler Routing Table"
WriteWordLine 0 0 " "

WriteWordLine 0 0 "Default routes are not in the ns.conf. Need to work on auto creating them based on MIP and SNIP"
WriteWordLine 0 0 " "

$ROWS = 1
$Add | foreach { 
   if ($_ -like 'add route *') {
   $ROWS = $ROWS+1
   }
}

$TableRange = $doc.Application.Selection.Range
[int]$Columns = 3
[int]$Rows = $ROWS
Write-Verbose "$(Get-Date): `t`tTable: Write Routing tables"
$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$table.Style = "Table Grid"
$table.Borders.InsideLineStyle = $wdLineStyleNone
$table.Borders.OutsideLineStyle = $wdLineStyleSingle

$COL = 0
do {
    $COL++
    $Table.Cell(1,$COL).Shading.BackgroundPatternColor = $wdColorGray15
    $Table.Cell(1,$COL).Range.Font.Bold = $True
    If($xRow % 2 -eq 0) {
	    $Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray05
        $Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray05
        $Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray05
        }
    }
while ($Col -le $Columns -1)

$Table.Cell(1,1).Range.Text = "Network"
$Table.Cell(1,2).Range.Text = "Subnet"
$Table.Cell(1,3).Range.Text = "Gateway"
$xRow = 1         

$Add | foreach { 
   if ($_ -like 'add route *') {
      $xRow++
      $Y = ($_ -replace 'add route ', '').split()
        If($xRow % 2 -eq 0) {
			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray05
            $Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray05
			
            }
			$Table.Cell($xRow,1).Range.Font.size = 9
			$Table.Cell($xRow,1).Range.Text = $Y[0]
            $Table.Cell($xRow,2).Range.Font.size = 9
			$Table.Cell($xRow,2).Range.Text = $Y[1]
            $Table.Cell($xRow,3).Range.Font.size = 9
			$Table.Cell($xRow,3).Range.Text = $Y[2]
		}	        
}

$table.AutoFitBehavior($wdAutoFitContent)

#return focus back to document
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

#move to the end of the current document
$selection.EndKey($wdStory,$wdMove) | Out-Null

WriteWordLine 0 0 " "

########################################################################
$selection.InsertNewPage()
WriteWordLine 2 0 "NetScaler DNS Server Records"

Write-Verbose "$(Get-Date): `Processing NetScaler DNS Server Record Configuration"

$ROWS = 1
$Add | foreach { 
   if ($_ -like 'add dns nsRec *') {
   $ROWS = $ROWS+1
   }
}

$TableRange = $doc.Application.Selection.Range
[int]$Columns = 2
[int]$Rows = $ROWS
Write-Verbose "$(Get-Date): `t`tTable: Write DNS Server Records"
$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$table.Style = "Table Grid"
$table.Borders.InsideLineStyle = $wdLineStyleNone
$table.Borders.OutsideLineStyle = $wdLineStyleSingle
$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,1).Range.Font.Bold = $True
$Table.Cell(1,1).Range.Text = "DNS Server Record"
$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,2).Range.Font.Bold = $True
$Table.Cell(1,2).Range.Text = "TTL"
$xRow = 1         

$Add | foreach { 
   if ($_ -like 'add dns nsRec *') {
      $xRow++
      $Y = ($_ -replace 'add dns nsRec ', '').split()
        If($xRow % 2 -eq 0) {
			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray05
            $Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray05
        }
			$Table.Cell($xRow,1).Range.Font.size = 9
			$Table.Cell($xRow,1).Range.Text = $Y[1]
			$Table.Cell($xRow,2).Range.Font.size = 9
			$Table.Cell($xRow,2).Range.Text = Get-StringProperty $_ "-TTL";
		}
		$table.AutoFitBehavior($wdAutoFitContent)        
}

#return focus back to document
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

#move to the end of the current document
$selection.EndKey($wdStory,$wdMove) | Out-Null

WriteWordLine 0 0 " "

########################################################################

Write-Verbose "$(Get-Date): `Processing NetScaler DNS Record Configuration"

WriteWordLine 2 0 "NetScaler DNS Records"

$ROWS = 1
$Add | foreach { 
   if ($_ -like 'add dns addRec *') {
   $ROWS = $ROWS+1
   }
}

$TableRange = $doc.Application.Selection.Range
[int]$Columns = 5
[int]$Rows = $ROWS
Write-Verbose "$(Get-Date): `t`tTable: Write DNS Records"
$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$table.Style = "Table Grid"
$table.Borders.InsideLineStyle = $wdLineStyleNone
$table.Borders.OutsideLineStyle = $wdLineStyleSingle
$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,1).Range.Font.Bold = $True
$Table.Cell(1,1).Range.Text = "DNS Record"
$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,2).Range.Font.Bold = $True
$Table.Cell(1,2).Range.Text = "IP Address"
$Table.Cell(1,3).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell(1,3).Range.Font.Bold = $True
$Table.Cell(1,3).Range.Text = "TTL"

$xRow = 1         

$Add | foreach { 
    if ($_ -like 'add dns addRec *') {
        $xRow++
        $Y = ($_ -replace 'add dns addRec ', '').split()
        If($xRow % 2 -eq 0) {
		    $Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray05
            $Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray05
            $Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray05

            }
		$Table.Cell($xRow,1).Range.Font.size = 9
		$Table.Cell($xRow,1).Range.Text = $Y[0]
        $Table.Cell($xRow,2).Range.Font.size = 9
        $Table.Cell($xRow,2).Range.Text = $Y[1]
        $Table.Cell($xRow,3).Range.Font.size = 9
		$Table.Cell($xRow,3).Range.Text = Get-StringProperty $_ "-TTL";
		}	        
    }

$table.AutoFitBehavior($wdAutoFitContent)

#return focus back to document
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

#move to the end of the current document
$selection.EndKey($wdStory,$wdMove) | Out-Null

#endregion NetScaler Network Configuration

#region NetScaler Authentication
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Authentication"

$selection.InsertNewPage()

WriteWordLine 1 0 "NetScaler Authentication"

WriteWordLine 0 0 " "

WriteWordLine 2 0 "NetScaler Local Users"

$Rows = 1
$Add | foreach {
    if ($_ -like 'add system user *') {
        $Rows++
    }
}

If ($Rows -eq 1) {WriteWordLine 0 0 "No Local Users configured"} else {

    $TableRange = $doc.Application.Selection.Range
    [int]$Columns = 1
    [int]$Rows = $Rows
    Write-Verbose "$(Get-Date): `t`tTable: Write Local User"
    $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
    $table.Style = "Table Grid"
	$table.Borders.InsideLineStyle = $wdLineStyleSingle
	$table.Borders.OutsideLineStyle = $wdLineStyleSingle
    $Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
    $Table.Cell(1,1).Range.Font.Bold = $True
    $Table.Cell(1,1).Range.Text = "Local User"

    $xRow = 1

    $Add | foreach {
        if ($_ -like 'add system user *') {
            $xRow++
            $X = ($_ -replace 'add system user ' , '').split()
            $Table.Cell($xRow,1).Range.Font.size = 9
		    $Table.Cell($xRow,1).Range.Text = $X[0]
            }
        }

    $table.AutoFitBehavior($wdAutoFitContent)        
           
    #return focus back to document
    $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

    #move to the end of the current document
    $selection.EndKey($wdStory,$wdMove) | Out-Null

    WriteWordLine 0 0 "Table: Local Users"
    }

WriteWordLine 2 0 "NetScaler Local Groups"

$Rows = 1
$Add | foreach {
    if ($_ -like 'add system group *') {
        $Rows++
    }
}

WriteWordLine 0 0 " "

If ($Rows -eq 1) {WriteWordLine 0 0 "No Local Groups configured"} else {
    $TableRange = $doc.Application.Selection.Range
    [int]$Columns = 2
    [int]$Rows = $Rows
    Write-Verbose "$(Get-Date): `t`tTable: Write Local Group"
    $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
    $table.Style = "Table Grid"
	$table.Borders.InsideLineStyle = $wdLineStyleSingle
	$table.Borders.OutsideLineStyle = $wdLineStyleSingle
        
    $COL = 0
    do {
        $COL++
        $Table.Cell(1,$COL).Shading.BackgroundPatternColor = $wdColorGray15
        $Table.Cell(1,$COL).Range.Font.Bold = $True
        }
    while ($Col -le $Columns -1)
    
    $Table.Cell(1,1).Range.Text = "Local Group"
    $Table.Cell(1,2).Range.Text = "Encrypted"
    $xRow = 1

    $Add | foreach {
        if ($_ -like 'add system group *') {
            $X = ($_ -replace 'add system group ' , '').split()
            $xRow++
            $Table.Cell($xRow,1).Range.Font.size = 9
		    $Table.Cell($xRow,1).Range.Text = $X
        }
    }
    $table.AutoFitBehavior($wdAutoFitContent)

    #return focus back to document
    $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

    #move to the end of the current document
    $selection.EndKey($wdStory,$wdMove) | Out-Null

    WriteWordLine 0 0 "Table: Local Groups"
    }

WriteWordLine 0 0 " "

$Rows = 1
$Add | foreach {
    if ($_ -like 'add authentication ldapPolicy*') {
        $Rows++
    }
}

If ($Rows -eq 1) {WriteWordLine 0 0 "No LDAP configured"} else {

    $Add | foreach {
       if ($_ -like 'add authentication ldapPolicy*') {
            $Y = ($_ -replace 'add authentication ldapPolicy ', '').split()
            $LDAPSERVER1 = "add authentication ldapAction $($Y[2]) *"
            $LDAPSERVER2 = "add authentication ldapAction $($Y[2])"
        
            $Add | foreach {
                if ($_ -like $LDAPSERVER1) {
                    $X = ($_ -replace $LDAPSERVER2 , '').split()
                    WriteWordLine 2 0 "LDAP Authentication $($Y[0])"
                    WriteWordLine 0 0 " "
                    $TableRange = $doc.Application.Selection.Range
                    [int]$Columns = 2
                    [int]$Rows = 7
                    Write-Verbose "$(Get-Date): `t`tTable: Write LDAP Authentication $($Y[0])"
		            $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		            $table.Style = "Table Grid"
					$table.Borders.InsideLineStyle = $wdLineStyleNone
					$table.Borders.OutsideLineStyle = $wdLineStyleSingle

                    $ROWC = 0
                        do {
                            $ROWC++
                            $Table.Cell($ROWC,1).Shading.BackgroundPatternColor = $wdColorGray15
                            $Table.Cell($ROWC,1).Range.Font.Bold = $True
                            }
                        while ($ROWC -le $ROWS-1)
                        
		            $Table.Cell(1,1).Range.Text = "IP"
		            $Table.Cell(2,1).Range.Text = "Base OU"
		            $Table.Cell(3,1).Range.Text = "Bind DN"
		            $Table.Cell(4,1).Range.Text = "Loginname"
		            $Table.Cell(5,1).Range.Text = "Sub Attribute Name"
		            $Table.Cell(6,1).Range.Text = "Security Type"
		            $Table.Cell(7,1).Range.Text = "Password Changes"

                    $ROWC = 0
                    do {
                        $ROWC++
                        $Table.Cell($ROWC,2).Range.Font.size = 9
                        }
                    while ($ROWC -le $ROWS-1)

			        $Table.Cell(1,2).Range.Text = Get-StringProperty $_ "-serverIP"
			        $Table.Cell(2,2).Range.Text = Get-StringProperty $_ "-ldapbase"
			        $Table.Cell(3,2).Range.Text = Get-StringProperty $_ "-ldapBindDn"
			        $Table.Cell(4,2).Range.Text = Get-StringProperty $_ "-ldapLoginName"
			        $Table.Cell(5,2).Range.Text = Get-StringProperty $_ "-subAttributeName"
			        $Table.Cell(6,2).Range.Text = Get-StringProperty $_ "-secType"
			        $Table.Cell(7,2).Range.Text = Get-StringProperty $_ "-passwdChange"
                    }

            }
            $table.AutoFitBehavior($wdAutoFitContent)        
            #return focus back to document
		    $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

		    #move to the end of the current document
		    $selection.EndKey($wdStory,$wdMove) | Out-Null

            WriteWordLine 0 0 "Table: NetScaler LDAP Policy $($Y[0])"
            WriteWordLine 0 0 " "
            }
        }
    }

$selection.InsertNewPage()

#endregion NetScaler Authentication

#region NetScaler Web Interface
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Web Interface"

WriteWordLine 1 0 "NetScaler Web Interface"

$CHECKUSAGE = 0
$File | foreach { 
    if ($_ -like 'install wi package *') {
        $CHECKUSAGE = 1
        }
    }
if ($CHECKUSAGE -ne 1) {WriteWordLine 0 0 "Citrix Web Interface has not been installed"}
if ($CHECKUSAGE -eq 1) {WriteWordLine 0 0 "Citrix Web Interface has been installed"}

$selection.InsertNewPage()

#endregion NetScaler Web Interface

#region NetScaler Traffic Domains
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Traffic Domains"

WriteWordLine 1 0 "NetScaler Traffic Domains"
WriteWordLine 0 0 " "

##No function yet for routing table per TD

$Rows = 1
$Add | foreach {
    if ($_ -like 'add ns trafficDomain *') {
        $Rows++
        }
    }

if($ROWS -eq 1) { WriteWordLine 0 0 "No Traffic Domains have been configured"} else {
    $Add | foreach {  
        if ($_ -like 'add ns trafficDomain *') {
            $Y = ($_ -replace 'add ns trafficDomain ', '').split()
            WriteWordLine 2 0 "Traffic Domain $($Y[0])"
        
            WriteWordLine 0 0 " "
        
            $TableRange = $doc.Application.Selection.Range
            [int]$Columns = 2
            [int]$Rows = 3

            Write-Verbose "$(Get-Date): `t`tTable: Write Traffic Domain $($Y[0])"
		    $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		    $table.Style = "Table Grid"
			$table.Borders.InsideLineStyle = $wdLineStyleNone
			$table.Borders.OutsideLineStyle = $wdLineStyleSingle

            $Row = 0
            do {
                $Row++
		        $Table.Cell($Row,1).Shading.BackgroundPatternColor = $wdColorGray15
		        $Table.Cell($Row,1).Range.Font.Bold = $True
                }
            while ($Row -le $Rows-1)
            $Row = 0
            do {
                $Row++
                $Table.Cell($Row,2).Range.Font.size = 9

                }
            while ($Row -le $Rows-1)


		    $Table.Cell(1,1).Range.Text = "Traffic Domain ID"
		    $Table.Cell(2,1).Range.Text = "Traffic Domain Alias"
		    $Table.Cell(3,1).Range.Text = "Traffic Domain vLAN"    
            $Table.Cell(1,2).Range.Text = $Y[0]
		    $Table.Cell(2,2).Range.Text = Get-StringProperty $_ "-aliasName"
            

            $Bind | foreach {  
                if ($_ -like 'bind ns trafficDomain *') {
                    $vLAN = Get-StringProperty $_ "-vlan"
                    }
                }

		    $Table.Cell(3,2).Range.Text = $vLAN
            
            $table.AutoFitBehavior($wdAutoFitContent)        
            #return focus back to document
            $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

            #move to the end of the current document
            $selection.EndKey($wdStory,$wdMove) | Out-Null

            WriteWordLine 0 0 "Table: NetScaler Traffic Domain $($Y[0])"
            WriteWordLine 0 0 " "
            
            ##TD Content Switch
            WriteWordLine 4 0 "Content Switch"
            
            $Rows = 1
            $ContentSwitch | foreach {
                if ((Get-StringProperty $_ "-td") -eq $Y[0]) {
                    $Rows++
                    $Rows
                    }
                }

            if($ROWS -eq 1) { WriteWordLine 0 0 "No Content Switch has been configured for Traffic Domain $($Y[0])"} else {
                $TableRange = $doc.Application.Selection.Range
                [int]$Columns = 1
                [int]$Rows = $Rows

		        $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		        $table.Style = "Table Grid"
				$table.Borders.InsideLineStyle = $wdLineStyleNone
				$table.Borders.OutsideLineStyle = $wdLineStyleSingle
		        $Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		        $Table.Cell(1,1).Range.Font.Bold = $True
		        $Table.Cell(1,1).Range.Text = "Content Switch"
                $xRow = 1               
                
                $ContentSwitch | foreach {
                    if ((Get-StringProperty $_ "-td") -eq $Y[0]) {
                        If ($_.Contains('-td $($Y[0])')) {
                            $Rows++
                            $Z = ($_ -replace 'add cs vserver ', '').split()
                            
                            $xRow++
                            $Table.Cell($xRow,1).Range.Font.siZe = 9
		                    $Table.Cell($xRow,1).Range.Text = $Z[0]
                        
                            $table.AutoFitBehavior($wdAutoFitContent)        
                            #return focus back to document
                            $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

                            #move to the end of the current document
                            $selection.EndKey($wdStory,$wdMove) | Out-Null
                            }
                        }
                    WriteWordLine 0 0 "Table: Traffic Domain $($Y[0]) Load Balancer"
                    }
                }
            WriteWordLine 0 0 " "
          
            ##TD Load Balancer
            WriteWordLine 4 0 "Load Balancer"
            $Rows = 1
            $LoadBalancer | foreach {
                if ((Get-StringProperty $_ "-td") -eq $Y[0]){ 
                    $Rows++
                    }
                }

            if($ROWS -eq 1) { WriteWordLine 0 0 "No Load Balancer has been configured for Traffic Domain $($Y[0])"} else {

                $TableRange = $doc.Application.Selection.Range
                [int]$Columns = 1
                [int]$Rows = $Rows

		        $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		        $table.Style = "Table Grid"
				$table.Borders.InsideLineStyle = $wdLineStyleNone
				$table.Borders.OutsideLineStyle = $wdLineStyleSingle
		        $Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		        $Table.Cell(1,1).Range.Font.Bold = $True
		        $Table.Cell(1,1).Range.Text = "Load Balancer"
                $xRow = 1
                
                $LoadBalancer | foreach {
                    if ((Get-StringProperty $_ "-td") -eq $Y[0]){
                        $Z = ($_ -replace 'add lb vserver ', '').split()
                        $xRow++
                        If($xRow % 2 -eq 0) {
	                        $Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray05
                            }
                                                
                        $Table.Cell($xRow,1).Range.Font.size = 9
		                $Table.Cell($xRow,1).Range.Text = $($Z[0])
                        
                        $table.AutoFitBehavior($wdAutoFitContent)        
                        #return focus back to document
                        $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

                        #move to the end of the current document
                        $selection.EndKey($wdStory,$wdMove) | Out-Null   
                        }
                    }
                WriteWordLine 0 0 "Table: Traffic Domain $($Y[0]) Load Balancer"
                }
            
            WriteWordLine 0 0 " "
        
            ##TD Services
            WriteWordLine 4 0 "Service"
            $Rows = 1
            $Service | foreach {
                if ((Get-StringProperty $_ "-td") -eq $Y[0]){
                    $Rows++
                    }
                }

            if($ROWS -eq 1) { WriteWordLine 0 0 "No Service has been configured for Traffic Domain $($Y[0])"} else {
                $TableRange = $doc.Application.Selection.Range
                [int]$Columns = 1
                [int]$Rows = $Rows

		        $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		        $table.Style = "Table Grid"
				$table.Borders.InsideLineStyle = $wdLineStyleNone
				$table.Borders.OutsideLineStyle = $wdLineStyleSingle
                $Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		        $Table.Cell(1,1).Range.Font.Bold = $True
		        $Table.Cell(1,1).Range.Text = "Service"
                
                $xRow = 1
                
                $Service | foreach {
                    if ((Get-StringProperty $_ "-td") -eq $Y[0]) {
                        $xRow++
                        $Z = ($_ -replace 'add service ', '').split()
                        If($xRow % 2 -eq 0) {
	                        $Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray05
                            }	
                        $Table.Cell($xRow,1).Range.Font.size = 9
		                $Table.Cell($xRow,1).Range.Text = $Z[0]
                        
                        $table.AutoFitBehavior($wdAutoFitContent)        
                        #return focus back to document
                        $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

                        #move to the end of the current document
                        $selection.EndKey($wdStory,$wdMove) | Out-Null
                        }
                    }
                WriteWordLine 0 0 "Table: Traffic Domain $($Y[0]) Services"
                }
            WriteWordLine 0 0 " "
                        
            ##TD Servers
            WriteWordLine 4 0 "Server"
            $Rows = 1
            $Server | foreach {
                if ((Get-StringProperty $_ "-td") -eq $Y[0]) {
                    $Rows++
                    }
                }

            if($ROWS -eq 1) { WriteWordLine 0 0 "No Server has been configured for Traffic Domain $($Y[0])"} else {
                $TableRange = $doc.Application.Selection.Range
                [int]$Columns = 1
                [int]$Rows = $Rows

		        $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		        $table.Style = "Table Grid"
				$table.Borders.InsideLineStyle = $wdLineStyleNone
				$table.Borders.OutsideLineStyle = $wdLineStyleSingle
		        $Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		        $Table.Cell(1,1).Range.Font.Bold = $True
		        $Table.Cell(1,1).Range.Text = "Server"
                
                $xRow = 1
                
                $Server | foreach {
                    if ((Get-StringProperty $_ "-td") -eq $Y[0]) {
                        $xRow++
                        $Z = ($_ -replace 'add server ', '').split()
                        If($xRow % 2 -eq 0) {
	                        $Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray05
                            }                        
                        $Table.Cell($xRow,1).Range.Font.size = 9
		                $Table.Cell($xRow,1).Range.Text = $Z[0]
                        
                        $table.AutoFitBehavior($wdAutoFitContent)        
                        #return focus back to document
                        $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

                        #move to the end of the current document
                        $selection.EndKey($wdStory,$wdMove) | Out-Null
                        }
                    }
                WriteWordLine 0 0 "Table: Traffic Domain $($Y[0]) Servers"
                }       
            WriteWordLine 0 0 " "
            }
        }
    }
 
$selection.InsertNewPage()

#endregion NetScaler Traffic Domains

#region NetScaler Monitoring
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Monitoring"

WriteWordLine 1 0 "NetScaler Monitoring"

WriteWordLine 2 0 "SNMP Community"

## add snmp community
$ROWS = 1
$Add | foreach { 
   if ($_ -like 'add snmp community *') {
   $ROWS = $ROWS+1
   }
}

If ($Rows -eq 1) {WriteWordLine 0 0 "No SNMP Community configured"} else {

    $TableRange = $doc.Application.Selection.Range
    [int]$Columns = 2
    [int]$Rows = $ROWS
    Write-Verbose "$(Get-Date): `t`tTable: Write SNMP Communities"
    $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
    $table.Style = "Table Grid"
	$table.Borders.InsideLineStyle = $wdLineStyleNone
	$table.Borders.OutsideLineStyle = $wdLineStyleSingle
    $Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
    $Table.Cell(1,1).Range.Font.Bold = $True
    $Table.Cell(1,1).Range.Text = "SNMP Community"
    $Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
    $Table.Cell(1,2).Range.Font.Bold = $True
    $Table.Cell(1,2).Range.Text = "Permission"
    $xRow = 1

    $Add | foreach { 
        if ($_ -like 'add snmp community *') {
            $Y = ($_ -replace 'add snmp community ', '').split()
            $xRow++
            $Table.Cell($xRow,1).Range.Font.size = 9
		    $Table.Cell($xRow,1).Range.Text = $Y[0]
            $Table.Cell($xRow,2).Range.Font.size = 9
		    $Table.Cell($xRow,2).Range.Text = $Y[1]
            If($xRow % 2 -eq 0) {
	            $Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray05
                $Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray05
                }
            }
        }

    $table.AutoFitBehavior($wdAutoFitContent)

    #return focus back to document
    $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

    #move to the end of the current document
    $selection.EndKey($wdStory,$wdMove) | Out-Null
    WriteWordLine 0 0 "Table: NetScaler SNMP Communities"
    }
    WriteWordLine 0 0 ""
      
## add snmp Manager
WriteWordLine 2 0 "SNMP Manager"

$ROWS = 1
$Add | foreach { 
   if ($_ -like 'add snmp manager *') {
   $ROWS = $ROWS+1
   }
}

If ($Rows -eq 1) {WriteWordLine 0 0 "No SNMP Manager configured"} else {

    $TableRange = $doc.Application.Selection.Range
    [int]$Columns = 2
    [int]$Rows = $ROWS
    Write-Verbose "$(Get-Date): `t`tTable: Write SNMP Manager"
    $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
    $table.Style = "Table Grid"
	$table.Borders.InsideLineStyle = $wdLineStyleNone
	$table.Borders.OutsideLineStyle = $wdLineStyleSingle
    $Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
    $Table.Cell(1,1).Range.Font.Bold = $True
    $Table.Cell(1,1).Range.Text = "SNMP Manager"
    $Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
    $Table.Cell(1,2).Range.Font.Bold = $True
    $Table.Cell(1,2).Range.Text = "Netmask"
    $xRow = 1

    $Add | foreach { 
        if ($_ -like 'add snmp manager *') {
            $Y = ($_ -replace 'add snmp manager ', '').split()
            $xRow++
            If($xRow % 2 -eq 0) {
	            $Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray05
	            $Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray05
                }
            $Table.Cell($xRow,1).Range.Font.size = 9
		    $Table.Cell($xRow,1).Range.Text = $Y[0]
            $Table.Cell($xRow,2).Range.Font.size = 9
		    $Table.Cell($xRow,2).Range.Text = Get-StringProperty $_ "-netmask";
            }
        }

    $table.AutoFitBehavior($wdAutoFitContent)

    #return focus back to document
    $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

    #move to the end of the current document
    $selection.EndKey($wdStory,$wdMove) | Out-Null
    WriteWordLine 0 0 "Table: NetScaler SNMP Managers"
    }
WriteWordLine 0 0 ""

## add snmp Alerts
WriteWordLine 2 0 "SNMP Alert"
$ROWS = 1
$Set | foreach { 
   if ($_ -like 'set snmp alarm *') {
   $ROWS = $ROWS+1
   }
}

If ($Rows -eq 1) {WriteWordLine 0 0 "No SNMP Alarms Configured"} else {

    $TableRange = $doc.Application.Selection.Range
    [int]$Columns = 4
    [int]$Rows = $ROWS
    Write-Verbose "$(Get-Date): `t`tTable: Write NetScaler Alarms"
    $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
    $table.Style = "Table Grid"
	$table.Borders.InsideLineStyle = $wdLineStyleSingle
	$table.Borders.OutsideLineStyle = $wdLineStyleSingle
    $Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
    $Table.Cell(1,1).Range.Font.Bold = $True
    $Table.Cell(1,1).Range.Text = "NetScaler Alarm"
    $Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
    $Table.Cell(1,2).Range.Font.Bold = $True
    $Table.Cell(1,2).Range.Text = "State"
    $Table.Cell(1,3).Shading.BackgroundPatternColor = $wdColorGray15
    $Table.Cell(1,3).Range.Font.Bold = $True
    $Table.Cell(1,3).Range.Text = "Time"
    $Table.Cell(1,4).Shading.BackgroundPatternColor = $wdColorGray15
    $Table.Cell(1,4).Range.Font.Bold = $True
    $Table.Cell(1,4).Range.Text = "Time-Out"
    $xRow = 1

    $Set | foreach { 
        if ($_ -like 'set snmp alarm *') {
            $Y = ($_ -replace 'set snmp alarm ', '').split()
            $xRow++
            $Table.Cell($xRow,1).Range.Font.size = 9
		    $Table.Cell($xRow,1).Range.Text = $($Y[0])
            $Table.Cell($xRow,2).Range.Font.size = 9
		    $Table.Cell($xRow,2).Range.Text = Test-NotStringPropertyEnabledDisabled $_ "-state";
            $Table.Cell($xRow,3).Range.Font.size = 9
            $Table.Cell($xRow,3).Range.Text = Get-StringProperty $_ "-time" "0";
            $Table.Cell($xRow,4).Range.Font.size = 9
            $Table.Cell($xRow,4).Range.Text = Get-StringProperty $_ "-timeout" "NA";
        }
    }

    $table.AutoFitBehavior($wdAutoFitContent)

    #return focus back to document
    $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

    #move to the end of the current document
    $selection.EndKey($wdStory,$wdMove) | Out-Null
    WriteWordLine 0 0 "Table: NetScaler Alarms"
    }
WriteWordLine 0 0 ""

$selection.InsertNewPage()

#endregion NetScaler Monitoring

#region NetScaler Certificates
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters Certificates"

WriteWordLine 1 0 "NetScaler Certificates"

Write-Verbose "$(Get-Date): `t`tTable: Write Certificates"

$ROWS = 1
$Add | foreach { 
    if ($_ -like 'add ssl certKey *') {
        $ROWS++
    }
}

If ($ROWS -eq 1) {WriteWordLine 0 0 "No Certificates installed"} Else {
    $Add | foreach { 
        if ($_ -like 'add ssl certKey *') {
            Write-Verbose "$(Get-Date): `t`tTable: Write Certificate $($Y[0])"
            $Y = ($_ -replace 'add ssl certKey ', '').split()
            $TableRange = $doc.Application.Selection.Range
            [int]$Columns = 2
            [int]$Rows = 4

            $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
            $table.Style = "Table Grid"
			$table.Borders.InsideLineStyle = $wdLineStyleNone
			$table.Borders.OutsideLineStyle = $wdLineStyleSingle
            $Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
            $Table.Cell(1,1).Range.Font.Bold = $True
            $Table.Cell(1,1).Range.Text = "Certificate"
            $Table.Cell(2,1).Shading.BackgroundPatternColor = $wdColorGray15
            $Table.Cell(2,1).Range.Font.Bold = $True
            $Table.Cell(2,1).Range.Text = "Certificate File"
            $Table.Cell(3,1).Shading.BackgroundPatternColor = $wdColorGray15
            $Table.Cell(3,1).Range.Font.Bold = $True
            $Table.Cell(3,1).Range.Text = "Certificate Key"
            $Table.Cell(4,1).Shading.BackgroundPatternColor = $wdColorGray15
            $Table.Cell(4,1).Range.Font.Bold = $True
            $Table.Cell(4,1).Range.Text = "Inform"
        
            $Table.Cell(1,2).Range.Font.size = 9
		    $Table.Cell(1,2).Range.Text = $($Y[0])
            $Table.Cell(2,2).Range.Font.size = 9
            $Table.Cell(3,2).Range.Font.size = 9
            $Table.Cell(4,2).Range.Font.size = 9
            $Table.Cell(2,2).Range.Text = Get-StringProperty $_ "-cert"
            $Table.Cell(3,2).Range.Text = Get-StringProperty $_ "-key"
            $Table.Cell(4,2).Range.Text = Get-StringProperty $_ "-inform" "NA";
            $table.AutoFitBehavior($wdAutoFitFixed)

            #return focus back to document
            $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

            #move to the end of the current document
            $selection.EndKey($wdStory,$wdMove) | Out-Null
            WriteWordLine 0 0 "Table: NetScaler Certificate $($Y[0]) "
            WriteWordLine 0 0 ""
        }
    }
}

$selection.InsertNewPage()

#endregion NetScaler Certificates

#region NetScaler Content Switches
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Content Switches"

WriteWordLine 1 0 "NetScaler Content Switches"

WriteWordLine 0 0 " "

$ROWS=1
$ContentSwitch | foreach { 
    if ($_ -like 'add cs vserver *') {
    $Rows++
    }
}
$RowsC = 0
$RowsTotal = $Rows
if($ROWS -eq 1) { WriteWordLine 0 0 "No Content Switches have been configured"} else {
    $ContentSwitch | foreach { 
        if ($_ -like 'add cs vserver *') {
            $RowsC++
            $Y = ($_ -replace 'add cs vserver ', '').split()
      
            WriteWordLine 2 0 "Content Switch $($Y[0]) "
      
            $TableRange = $doc.Application.Selection.Range
            [int]$Columns = 8
            [int]$Rowst = 2
            Write-Verbose "$(Get-Date): `t`tTable: $RowsC/$RowsTotal Content Switch Table $($Y[0])"
		    $Table = $doc.Tables.Add($TableRange, $Rowst, $Columns)
		    $table.Style = "Table Grid"
			$table.Borders.InsideLineStyle = $wdLineStyleSingle
			$table.Borders.OutsideLineStyle = $wdLineStyleSingle
            
            $Col = 0
            do {
                $Col++
                $Table.Cell(1,$Col).Shading.BackgroundPatternColor = $wdColorGray15
	            $Table.Cell(1,$Col).Range.Font.Bold = $True
                }
            while ($Col -le $Columns -1)

		    $Table.Cell(1,1).Range.Text = "State"
		    $Table.Cell(1,2).Range.Text = "Protocol"
		    $Table.Cell(1,3).Range.Text = "Port"
		    $Table.Cell(1,4).Range.Text = "IP"
		    $Table.Cell(1,5).Range.Text = "Traffic Domain"
		    $Table.Cell(1,6).Range.Text = "Case Sensitive"
		    $Table.Cell(1,7).Range.Text = "Down State Flush"
		    $Table.Cell(1,8).Range.Text = "Client Time-Out"

            $Table.Cell(2,1).Range.Font.size = 9
		    $Table.Cell(2,1).Range.Text = Test-NotStringPropertyEnabledDisabled $_ "-state";
            $Table.Cell(2,2).Range.Font.size = 9
		    $Table.Cell(2,2).Range.Text = $Y[1]
            $Table.Cell(2,3).Range.Font.size = 9
		    $Table.Cell(2,3).Range.Text = $Y[3]
            $Table.Cell(2,4).Range.Font.size = 9
		    $Table.Cell(2,4).Range.Text = $Y[2]
            $Table.Cell(2,6).Range.Font.size = 9
		    $Table.Cell(2,6).Range.Text = Get-StringProperty $_ "-caseSensitive";
            $Table.Cell(2,5).Range.Font.size = 9
		    $Table.Cell(2,5).Range.Text = Get-StringProperty $_ "-td" "0 (Default)";
            $Table.Cell(2,7).Range.Font.size = 9
		    $Table.Cell(2,7).Range.Text = Test-NotStringPropertyEnabledDisabled $_ "-downStateFlush";
            $Table.Cell(2,8).Range.Font.size = 9
		    $Table.Cell(2,8).Range.Text = Get-StringProperty $_ "-cltTimeout" "NA";
        
            $table.AutoFitBehavior($wdAutoFitContent)

            #return focus back to document
		    $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

		    #move to the end of the current document
		    $selection.EndKey($wdStory,$wdMove) | Out-Null
            WriteWordLine 0 0 "Table: Basic Configuration"

            WriteWordLine 0 0 " "

            $CSNAMEBIND = "bind cs vserver $($Y[0]) *"

            WriteWordLine 0 0 " "
            
            ##CS Policy Table
            $ROWS = 1
            $Bind | foreach {
                if ($_ -like $CSNAMEBIND) {
                    if (Test-StringProperty $_ "-policyName") {
                        $ROWS++
                        }
                    }
                }
            
            if($ROWS -eq 1) { WriteWordLine 0 0 "No Policies have been configured for this Content Switch"} else {

                $TableRange = $doc.Application.Selection.Range
                [int]$Columns2 = 4
                [int]$Rows = $ROWS
                $Table = $doc.Tables.Add($TableRange, $Rows, $Columns2)
		        $table.Style = "Table Grid"
				$table.Borders.InsideLineStyle = $wdLineStyleSingle
				$table.Borders.OutsideLineStyle = $wdLineStyleSingle

                $Col2 = 0
                do {
                    $Col2++
                    $Table.Cell(1,$Col2).Shading.BackgroundPatternColor = $wdColorGray15
	                $Table.Cell(1,$Col2).Range.Font.Bold = $True
                    }
                while ($Col2 -le $Columns2 -1)
		        $Table.Cell(1,1).Range.Text = "Policy"
		        $Table.Cell(1,2).Range.Text = "Load Balancer"
		        $Table.Cell(1,3).Range.Text = "Priority"
		        $Table.Cell(1,4).Range.Text = "Rule"
                $xRow = 1
               
                $Bind | foreach {
                    if ($_ -like $CSNAMEBIND) {
                        if (Test-StringProperty $_ "-policyName") {
                            $xRow++                  
                            $Table.Cell($xRow,1).Range.Font.size = 9
		                    $Table.Cell($xRow,1).Range.Text = Get-StringProperty $_ "-policyName";
                            $Table.Cell($xRow,2).Range.Font.size = 9
		                    $Table.Cell($xRow,2).Range.Text = Get-StringProperty $_ "-targetLBVserver";
                            $Table.Cell($xRow,3).Range.Font.size = 9
		                    $Table.Cell($xRow,3).Range.Text = Get-StringProperty $_ "-priority";
                        
                            $CSPOLADD = "add cs policy $(Get-StringProperty $_ "-policyName") *"
                                $File | foreach {
                                    if ($_ -like $CSPOLADD) {
                                        $Table.Cell($xRow,4).Range.Font.size = 9
		                                $Table.Cell($xRow,4).Range.Text = Get-StringProperty $_ "-rule";
                                        }
                                    }
                            }
                        }
                    }

                $table.AutoFitBehavior($wdAutoFitContent)    

                #return focus back to document
		        $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

    		    #move to the end of the current document
		        $selection.EndKey($wdStory,$wdMove) | Out-Null
                WriteWordLine 0 0 "Table: Content Switch $($Y[0]) Configuration"
                WriteWordLine 0 0 " "
                }
            
            WriteWordLine 0 0 " "
        
            ##Table Redirect URL
            WriteWordLine 4 0 "Redirect URL"
            $REDIR = Get-StringProperty $_ "-redirectURL" "NA";

            if (Test-StringProperty $_ "-redirectURL") {
                $TableRange = $doc.Application.Selection.Range
                [int]$Columns = 1
                [int]$Rows = 2
                $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		        $table.Style = "Table Grid"
				$table.Borders.InsideLineStyle = $wdLineStyleSingle
				$table.Borders.OutsideLineStyle = $wdLineStyleSingle
                $Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		        $Table.Cell(1,1).Range.Font.Bold = $True
		        $Table.Cell(1,1).Range.Text = "Redirect URL"
                $Table.Cell(2,1).Range.Font.size = 9
		        $Table.Cell(2,1).Range.Text = Get-StringProperty $_ "-redirectURL";

                $table.AutoFitBehavior($wdAutoFitContent)    

                #return focus back to document
		        $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

		        #move to the end of the current document
		        $selection.EndKey($wdStory,$wdMove) | Out-Null
                WriteWordLine 0 0 "Table: Content Switch Redirection URL"
                WriteWordLine 0 0 " "
            }
            else { WriteWordLine 0 0 "No Redirect URL has been configured for this Content Switch"; }   
    
            ##Advanced Configuration   
            WriteWordLine 4 0 "Advanced Configuration"
            if ($(Get-StringProperty $_ "-comment") -ne $null) {$X = Get-StringProperty $_ "-comment" "No comment";
            WriteWordLine 0 0 "Comment`t`t`t`t: $(Get-StringProperty "-comment" "No comment")";
            WriteWordLine 0 0 "Apply AppFlow logging`t`t`t: $(Test-NotStringPropertyEnabledDisabled $_ "-appflowLog")";
            WriteWordLine 0 0 "Name of the TCP profile`t`t`t: $(Get-StringProperty $_ "-tcpProfileName" "None")";
            WriteWordLine 0 0 "Name of the HTTP profile`t`t: $(Get-StringProperty $_ "-httpProfileName" "None")";
            WriteWordLine 0 0 "Name of the NET profile`t`t`t: $(Get-StringProperty $_ "-netProfile" "None")";
            WriteWordLine 0 0 "Name of the DB profile`t`t`t: $(Get-StringProperty $_ "-dbProfileName" "None")";
            WriteWordLine 0 0 "Enable or disable user authentication`t: $(Test-StringPropertyOnOff $_ "-Authentication")";
            WriteWordLine 0 0 "Authentication virtual server fqdn`t: $(Get-StringProperty $_ "-AuthenticationHost" "NA")";
            WriteWordLine 0 0 "Name of the Authentication profile`t: $(Get-StringProperty $_ "-authnProfile" "None")";
            WriteWordLine 0 0 "Syntax expression identifying traffic`t: $(Get-StringProperty $_ "-Listenpolicy" "None")"
            WriteWordLine 0 0 "Priority of the Listener Policy`t`t: $(Get-StringProperty $_ "-Listenpriority" "101 (Maximum Value)")";
            WriteWordLine 0 0 "Name of the backup virtual server`t: $(Get-StringProperty $_ "-backupVserver" "NA")";
            WriteWordLine 0 0 "Enable state updates`t`t`t: $(Test-StringPropertyEnabledDisabled $_ "-stateupdate")";
            WriteWordLine 0 0 "Route requests to the cache server`t: $(Test-StringPropertyYesNo $_ "-cacheable")";
            WriteWordLine 0 0 "precedence to use for policies`t`t: $(Get-StringProperty $_ "-precedence" "CS_PRIORITY_RULE (Default)")";
            WriteWordLine 0 0 "URL Case sensitive`t`t`t: $(Test-NotStringPropertyOnOff $_ "-caseSensitive")";
            WriteWordLine 0 0 "Type of spillover`t`t`t: $(Get-StringProperty $_ "-soMethod" "None")";
            WriteWordLine 0 0 "Maintain source-IP based persistence`t: Minutes $(Get-StringProperty $_ "-soPersistence" "None")";
            WriteWordLine 0 0 "Action if spillover is to take effect`t: $(Get-StringProperty $_ "-soBackupAction" "NA")";
            WriteWordLine 0 0 "State of port rewrite HTTP redirect`t: $(Test-NotStringPropertyEnabledDisabled $_ "-redirectPortRewrite")";
            WriteWordLine 0 0 "Continue forwarding to backup vServer`t: $(Test-StringPropertyEnabledDisabled $_ "-disablePrimaryOnDown")";
            }
        }
    }
    $selection.InsertNewPage()
}

$selection.InsertNewPage()

#endregion NetScaler Content Switches

#region NetScaler Load Balancers
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Load Balancers"

WriteWordLine 1 0 "NetScaler Load Balancing"

$Rows = 1
$LoadBalancer | foreach {
    $Rows++
    }

$RowsC = 0
$RowsTotal = $Rows

if($ROWS -eq 1) { WriteWordLine 0 0 "No Load Balancer have been configured"} else {
    $LoadBalancer | foreach { 
        $Y = ($_ -replace 'add lb vserver ', '').split()
        $RowsC++
            WriteWordLine 2 0 "Load Balancer $($Y[0]) "
            $TableRange = $doc.Application.Selection.Range
            [int]$Columns = 8
            [int]$Rowst = 2
            Write-Verbose "$(Get-Date): `t`tTable: $RowsC/$RowsTotal Write Load Balancer Table $($Y[0])"
	        $Table = $doc.Tables.Add($TableRange, $Rowst, $Columns)
	        $table.Style = "Table Grid"
			$table.Borders.InsideLineStyle = $wdLineStyleSingle
			$table.Borders.OutsideLineStyle = $wdLineStyleSingle
            $Col = 0
            do {
                $Col++
                $Table.Cell(1,$Col).Shading.BackgroundPatternColor = $wdColorGray15
	            $Table.Cell(1,$Col).Range.Font.Bold = $True
                }
            while ($Col -le $Columns -1)
	        $Table.Cell(1,1).Range.Text = "State"
	        $Table.Cell(1,2).Range.Text = "Protocol"
	        $Table.Cell(1,3).Range.Text = "Port"
	        $Table.Cell(1,4).Range.Text = "IP"
	        $Table.Cell(1,5).Range.Text = "PERSISTENCY"
	        $Table.Cell(1,6).Range.Text = "Traffic Domain"
	        $Table.Cell(1,7).Range.Text = "Method"
	        $Table.Cell(1,8).Range.Text = "Client Time-Out"
        
            If (Test-StringProperty $_ "-state") {$STATE = "Disabled"} else {$STATE = "Enabled"}
            
            $Rowsetup = 0
            $Col = 0
            do {
                $Col++
                $Table.Cell(2,$Col).Range.Font.size = 9
                }
            while ($Col -le $Columns -1)

	        $Table.Cell(2,1).Range.Text = $STATE
	        $Table.Cell(2,2).Range.Text = $($Y[1])
	        $Table.Cell(2,3).Range.Text = $($Y[3])
	        $Table.Cell(2,4).Range.Text = $($Y[2])
	        $Table.Cell(2,5).Range.Text = Get-StringProperty $_ "-persistenceType";
	        $Table.Cell(2,6).Range.Text = Get-StringProperty $_ "-td" "0 (Default)";
	        $Table.Cell(2,7).Range.Text = Get-StringProperty $_ "-lbmethod" "Least Connection";
	        $Table.Cell(2,8).Range.Text = Get-StringProperty $_ "-cltTimeout" "NA";
        
            $table.AutoFitBehavior($wdAutoFitContent)

            #return focus back to document
	        $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

	        #move to the end of the current document
	        $selection.EndKey($wdStory,$wdMove) | Out-Null
            WriteWordLine 0 0 "Table: Load Balancer Configuration"
            WriteWordLine 0 0 " "

            $LBNAMEBIND = "bind lb vserver $($Y[0]) *"
            ##Services Table
            WriteWordLine 4 0 "Services and Service Groups"
            $ROWS2 = 1
            $LoadbalancerBind | foreach {
                if ($_ -like $LBNAMEBIND) {
                    if (-not (Test-StringProperty $_ "-policyName")) {
                        $ROWS2++
                        }
                    }
                }
            
            if($ROWS2 -eq 1) { WriteWordLine 0 0 "No Services have been configured for this Load Balancer"} else {

                $TableRange = $doc.Application.Selection.Range
                [int]$Columns2 = 1
                [int]$Rows2 = $ROWS2
                $Table = $doc.Tables.Add($TableRange, $Rows2, $Columns2)
		        $table.Style = "Table Grid"
				$table.Borders.InsideLineStyle = $wdLineStyleSingle
				$table.Borders.OutsideLineStyle = $wdLineStyleSingle
                $Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		        $Table.Cell(1,1).Range.Font.Bold = $True
		        $Table.Cell(1,1).Range.Text = "Service"
                $xRow2 = 1
               
                $LoadbalancerBind | foreach {
                    if ($_ -like $LBNAMEBIND) {
                        if (-not (Test-StringProperty $_ "-policyName")) {
                            $xRow2++
                            $G = $_.split()
                            $Table.Cell($xRow2,1).Range.Font.size = 9
		                    $Table.Cell($xRow2,1).Range.Text = $G[4]
                            }
                        }
                    }

                $table.AutoFitBehavior($wdAutoFitContent)    

                #return focus back to document
		        $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

		        #move to the end of the current document
		        $selection.EndKey($wdStory,$wdMove) | Out-Null
                WriteWordLine 0 0 "Table: Load Balancer Services"
            }
            WriteWordLine 0 0 " "
            WriteWordLine 4 0 "Policies"
            $ROWS3 = 1
            $LoadbalancerBind | foreach {
                if ($_ -like $LBNAMEBIND) {
                    if (Test-StringProperty $_ "-policyName") {
                        $ROWS3++
                        }
                    }
                }
            WriteWordLine 0 0 " "
                    
            if($ROWS3 -eq 1) { WriteWordLine 0 0 "No Policies have been configured for this Load Balancer"} else {

                $TableRange = $doc.Application.Selection.Range
                [int]$Columns3 = 4
                [int]$Rows3 = $Rows3
                $Table = $doc.Tables.Add($TableRange, $Rows3, $Columns3)
		        $table.Style = "Table Grid"
				$table.Borders.InsideLineStyle = $wdLineStyleSingle
				$table.Borders.OutsideLineStyle = $wdLineStyleSingle
                
                $Col3 = 0
                do {
                    $Col3++
                    $Table.Cell(1,$Col3).Shading.BackgroundPatternColor = $wdColorGray15
	                $Table.Cell(1,$Col3).Range.Font.Bold = $True
                    }
                while ($Col3 -le $Columns3 -1)

		        $Table.Cell(1,1).Range.Text = "Policy Name"
		        $Table.Cell(1,2).Range.Text = "Policy Name"
		        $Table.Cell(1,3).Range.Text = "Policy Type"
		        $Table.Cell(1,4).Range.Text = "GoTo Expression"
                $xRow3 = 1
    
                $LoadbalancerBind | foreach {
                    if ($_ -like $LBNAMEBIND) {
                        if (Test-StringProperty $_ "-policyName") {
                            $xRow3++
                    
                            $Table.Cell($xRow3,1).Range.Font.size = 9
		                    $Table.Cell($xRow3,1).Range.Text = Get-StringProperty $_ "-policyName";
                            $Table.Cell($xRow3,2).Range.Font.size = 9
		                    $Table.Cell($xRow3,2).Range.Text = Get-StringProperty $_ "-priority";
                            $Table.Cell($xRow3,3).Range.Font.size = 9
		                    $Table.Cell($xRow3,3).Range.Text = Get-StringProperty $_ "-type";
                            $Table.Cell($xRow3,4).Range.Font.size = 9
		                    $Table.Cell($xRow3,4).Range.Text = Get-StringProperty $_ "-gotoPriorityExpression";
                            }
                        }
                    }    
                $table.AutoFitBehavior($wdAutoFitContent)    

                #return focus back to document
		        $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

		        #move to the end of the current document
		        $selection.EndKey($wdStory,$wdMove) | Out-Null
                WriteWordLine 0 0 "Table: Load Balancer Policies"
                }
            WriteWordLine 0 0 " "
            WriteWordLine 4 0 "Redirect URL"

            if (Test-StringProperty $_ "-redirectURL") {
                $TableRange = $doc.Application.Selection.Range
                [int]$Columns4 = 1
                [int]$Rows4 = 2
                $Table = $doc.Tables.Add($TableRange, $Rows4, $Columns4)
		        $table.Style = "Table Grid"
				$table.Borders.InsideLineStyle = $wdLineStyleSingle
				$table.Borders.OutsideLineStyle = $wdLineStyleSingle
                $Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		        $Table.Cell(1,1).Range.Font.Bold = $True
		        $Table.Cell(1,1).Range.Text = "Redirect URL"
                $Table.Cell(2,1).Range.Font.size = 9
		        $Table.Cell(2,1).Range.Text = Get-StringProperty $_ "-redirectURL";

                $table.AutoFitBehavior($wdAutoFitContent)    

                #return focus back to document
		        $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

		        #move to the end of the current document
		        $selection.EndKey($wdStory,$wdMove) | Out-Null
                WriteWordLine 0 0 "Table: Load Balancer Redirection URL"
                WriteWordLine 0 0 " "
            }
            else { WriteWordLine 0 0 "No Redirect URL has been configured for this Load Balancer"; }
        ##Advanced Configuration   

        WriteWordLine 4 0 "Advanced Configuration"
        WriteWordLine 0 0 "Comment`t`t`t`t: $(Get-StringProperty $_ "-comment" "No comment")";
        WriteWordLine 0 0 "Apply AppFlow logging`t`t`t: $(Test-NotStringPropertyEnabledDisabled $_ "-appflowLog")";
        WriteWordLine 0 0 "Name of the TCP profile`t`t`t: $(Get-StringProperty $_ "-tcpProfileName" "None")";
        WriteWordLine 0 0 "Name of the HTTP profile`t`t: $(Get-StringProperty $_ "-httpProfileName" "None")";
        WriteWordLine 0 0 "Name of the NET profile`t`t`t: $(Get-StringProperty $_ "-netProfile" "None")";
        WriteWordLine 0 0 "Name of the DB profile`t`t`t: $(Get-StringProperty $_ "-dbProfileName" "None")";
        WriteWordLine 0 0 "Enable or disable user authentication`t: $(Test-StringPropertyOnOff $_ "-Authentication")";
        WriteWordLine 0 0 "Authentication virtual server fqdn`t: $(Get-StringProperty $_ "-AuthenticationHost" "NA")";
        WriteWordLine 0 0 "Authentication virtual server name`t: $(Get-StringProperty $_ "-authnVsname" "NA")";
        WriteWordLine 0 0 "Name of the Authentication profile`t: $(Get-StringProperty $_ "-authnProfile" "None")"; 
        WriteWordLine 0 0 "User authentication with HTTP 401`t: $(Test-StringPropertyOnOff $_ "-authn401")";
        WriteWordLine 0 0 "Syntax expression identifying traffic`t: $(Get-StringProperty $_ "-Listenpolicy" "None")";
        WriteWordLine 0 0 "Priority of the Listener Policy`t`t: $(Get-StringProperty $_ "-Listenpriority" "101 (Maximum Value)")";
        WriteWordLine 0 0 "Name of the backup virtual server`t: $(Get-StringProperty $_ "-backupVServer" "NA")";
        WriteWordLine 0 0 "Time period a persistence session`t: $(Get-StringProperty $_ "-timeout" "2 (Default Value)")";
        WriteWordLine 0 0 "Backup persistence type`t`t: $(Get-StringProperty $_ "-persistenceBackup" "None")";
        WriteWordLine 0 0 "Time period a persistence session`t: $(Get-StringProperty $_ "-backupPersistenceTimeout" "2 (Default Value)")";
        WriteWordLine 0 0 "Use priority queuing`t`t`t: $(Test-StringPropertyOnOff $_ "-pq")";
        WriteWordLine 0 0 "Use SureConnect`t`t`t: $(Test-StringPropertyOnOff $_ "-sc")";
        WriteWordLine 0 0 "Use network address translation`t: $(Test-StringPropertyOnOff $_ "-rtspNat")";
        WriteWordLine 0 0 "Redirection mode for load balancing`t: $(Get-StringProperty $_ "-m" "NSFWD_IP (Default)")";
        WriteWordLine 0 0 "Use Layer 2 parameters`t`t`t: $(Test-StringPropertyOnOff $_ "-l2Conn")";
        WriteWordLine 0 0 "TOS ID of the virtual server`t`t: $(Get-StringProperty $_ "-tosId" "1 (Default)")";
        WriteWordLine 0 0 "Expression against which traffic is evaluated`t`t: $(Get-StringProperty $_ "-rule" "None")";
        WriteWordLine 0 0 "Perform load balancing on a per-packet basis`t`t: $(Test-StringPropertyEnabledDisabled $_ "-sessionless")";
        WriteWordLine 0 0 "How the NetScaler appliance responds to ping requests`t: $(Get-StringProperty $_ "-icmpVsrResponse" "NS_VSR_PASSIVE (Default)")";
        WriteWordLine 0 0 "Route cacheable requests to a cache redirection server`t: $(Test-StringPropertyYesNo $_ "-cacheable")";

<# Seems like a lot of information turned off for now!
        if ($(Get-StringProperty $_ "-downStateFlush") -ne $null) {$X = "Disabled"} else {$X = "Enabled"}
        WriteWordLine 0 0 "Flush all active transactions associated with a virtual server whose state transitions from UP to DOWN`t: $X"
        if ($(Get-StringProperty $_ "-dns64") -ne $null) {$X = "Enabled"} else {$X = "Disabled"}
        WriteWordLine 0 0 "Dns64 on lbvserver`t`t: $X"
        if ($(Get-StringProperty $_ "-bypassAAAA") -ne $null) {$X = "Yes"} else {$X = "No"}
        WriteWordLine 0 0 "While resolving DNS64 query AAAA queries are not sent to back end dns server`t: $X"
        if ($(Get-StringProperty $_ "-RecursionAvailable") -ne $null) {$X = "Yes"} else {$X = "No"}
        WriteWordLine 0 0 "This option causes the DNS replies from this vserver to have the RA bit turned on`t: $X"
        if ($(Get-StringProperty $_ "-range") -ne $null) {$X = $(Get-StringProperty $_ "-range")} else {$X = "1 (Minimum Value)"}
        WriteWordLine 0 0 "Number of IP addresses that the appliance must generate and assign to the virtual server`t: $X"
        if ($(Get-StringProperty $_ "-cookieName") -ne $null) {$X = $(Get-StringProperty $_ "-cookieName")} else {$X = "NA"}
        WriteWordLine 0 0 "Cookie name for COOKIE peristence type`t: $X"
        if ($(Get-StringProperty $_ "-resRule") -ne $null) {$X = $(Get-StringProperty $_ "-resRule")} else {$X = "None"}
        WriteWordLine 0 0 "Default syntax expression specifying which part of a server's response to use for creating rule based persistence sessions`t: $X"
        if ($(Get-StringProperty $_ "-persistMask") -ne $null) {$X = $(Get-StringProperty $_ "-persistMask")} else {$X = "0xFFFFFFFF (Default Value)"}
        WriteWordLine 0 0 "Persistence mask for IP based persistence types`t: $X"
        if ($(Get-StringProperty $_ "-v6persistmasklen") -ne $null) {$X = $(Get-StringProperty $_ "-v6persistmasklen")} else {$X = "128 (Default Value)"}
        WriteWordLine 0 0 "Persistence mask for IP based persistence types (IPv6)`t: $X"
        if ($(Get-StringProperty $_ "-dataLength") -ne $null) {$X = $(Get-StringProperty $_ "-dataLength")} else {$X = "1 (Default)"}
        WriteWordLine 0 0 "Length of the token to be extracted`t: $X"
        if ($(Get-StringProperty $_ "-dataOffset") -ne $null) {$X = $(Get-StringProperty $_ "-dataOffset")} else {$X = "25400 (Default)"}
        WriteWordLine 0 0 "Offset to be considered when extracting a token`t: $X"
        if ($(Get-StringProperty $_ "-connfailover") -ne $null) {$X = "Enabled"} else {$X = "Disabled"}
        WriteWordLine 0 0 "Mode in which the connection failover feature must operate`t: $X"
        if ($(Get-StringProperty $_ "-soMethod") -ne $null) {$X = $(Get-StringProperty $_ "-soMethod")} else {$X = "None"}
        WriteWordLine 0 0 "Type of threshold that triggers spillover`t: $X"
        if ($(Get-StringProperty $_ "-soPersistence") -ne $null) {$X = "Enabled"} else {$X = "Disabled"}
        WriteWordLine 0 0 "If spillover occurs, maintain source IP address based persistence`t: $X"
        if ($(Get-StringProperty $_ "-soPersistenceTimeOut") -ne $null) {$X = $(Get-StringProperty $_ "-soPersistenceTimeOut")} else {$X = "2 (Default)"}
        WriteWordLine 0 0 "Timeout for spillover persistence, in minutes`t: $X"
        if ($(Get-StringProperty $_ "-healthThreshold") -ne $null) {$X = $(Get-StringProperty $_ "-healthThreshold")} else {$X = "NA"}
        WriteWordLine 0 0 "Threshold in percent of active services below which vserver state is made down`t: $X"
        if ($(Get-StringProperty $_ "-soThreshold") -ne $null) {$X = $(Get-StringProperty $_ "-soThreshold")} else {$X = "NA"}
        WriteWordLine 0 0 "Threshold at which spillover occurs`t: $X"
        if ($(Get-StringProperty $_ "-soBackupAction") -ne $null) {$X = $(Get-StringProperty $_ "-soBackupAction")} else {$X = "NA"}
        WriteWordLine 0 0 "Action to be performed if spillover is to take effect`t: $X"
        if ($(Get-StringProperty $_ "-redirectPortRewrite") -ne $null) {$X = "Enabled"} else {$X = "Disabled"}
        WriteWordLine 0 0 "Rewrite the port and change the protocol to ensure successful HTTP redirects from services`t: $X"
        if ($(Get-StringProperty $_ "-disablePrimaryOnDown") -ne $null) {$X = "Enabled"} else {$X = "Disabled"}
        WriteWordLine 0 0 "If the primary virtual server goes down, do not allow it to return to primary status until manually enabled`t: $X"
        if ($(Get-StringProperty $_ "-insertVserverIPPort") -ne $null) {$X = $(Get-StringProperty $_ "-insertVserverIPPort")} else {$X = "Off"}
        WriteWordLine 0 0 "Insert an HTTP header`t`t: $X"
        if ($(Get-StringProperty $_ "-push") -ne $null) {$X = "Enabled"} else {$X = "Disabled"}
        WriteWordLine 0 0 "Process traffic with the push virtual server`t: $X"
        if ($(Get-StringProperty $_ "-pushVserver") -ne $null) {$X = $(Get-StringProperty $_ "-pushVserver")} else {$X = "NA"}
        WriteWordLine 0 0 "Name of the virtual server to which the server pushes updates`t: $X"
        if ($(Get-StringProperty $_ "-pushLabel") -ne $null) {$X = $(Get-StringProperty $_ "-pushLabel")} else {$X = "None"}
        WriteWordLine 0 0 "Expression for extracting a label from the server's response`t: $X"
        if ($(Get-StringProperty $_ "-pushMultiClients") -ne $null) {$X = "Yes"} else {$X = "No"}
        WriteWordLine 0 0 "Allow multiple Web 2.0 connections from the same client to connect to the virtual server and expect updates`t: $X"
        if ($(Get-StringProperty $_ "-newServiceRequest") -ne $null) {$X = $(Get-StringProperty $_ "-newServiceRequest")} else {$X = "(Default)"}
        WriteWordLine 0 0 "Number of requests by which to increase the load on a new service`t: $X"
        if ($(Get-StringProperty $_ "-newServiceRequestIncrementInterval") -ne $null) {$X = $(Get-StringProperty $_ "-newServiceRequestIncrementInterval")} else {$X = "(Default)"}
        WriteWordLine 0 0 "Interval, in seconds, between successive increments in the load on a new service`t: $X"
        if ($(Get-StringProperty $_ "-minAutoscaleMembers") -ne $null) {$X = $(Get-StringProperty $_ "-minAutoscaleMembers")} else {$X = "Disabled"}
        WriteWordLine 0 0 "Minimum number of members expected to be present when vserver is used in Autoscale`t: $X"
        if ($(Get-StringProperty $_ "-maxAutoscaleMembers") -ne $null) {$X = $(Get-StringProperty $_ "-maxAutoscaleMembers")} else {$X = "Disabled"}
        WriteWordLine 0 0 "Maximum number of members expected to be present when vserver is used in Autoscale`t: $X"
        if ($(Get-StringProperty $_ "-persistAVPno") -ne $null) {$X = $(Get-StringProperty $_ "-persistAVPno")} else {$X = "1 (Default)"}
        WriteWordLine 0 0 "Persist AVP number for Diameter Persistency`t: $X"
        if ($(Get-StringProperty $_ "-skippersistency") -ne $null) {$X = $(Get-StringProperty $_ "-skippersistency")} else {$X = "NS_DONT_SKIPPERSIST (Default)"}
        WriteWordLine 0 0 "This argument decides the behavior incase the service which is selected from an existing persistence session has reached threshold`t: $X"
        if ($(Get-StringProperty $_ "-macmodeRetainvlan") -ne $null) {$X = "Enabled"} else {$X = "Disabled"}
        WriteWordLine 0 0 "Retain vlan information of incoming packet`t: $X"
        if ($(Get-StringProperty $_ "-dbsLb") -ne $null) {$X = "Enabled"} else {$X = "Disabled"}
        WriteWordLine 0 0 "For enabling database specific load-balacing`t: $X"
        #>
        $selection.InsertNewPage()
        
        }
    
    }

#endregion NetScaler Load Balancers

#region NetScaler Services
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Services"

WriteWordLine 1 0 "NetScaler Service and Service Groups"
WriteWordLine 2 0 "NetScaler Service"
WriteWordLine 0 0 " "

$Rows = 1
$Service | foreach {
    $Rows++
    }

$RowsC = 0
$RowsTotal = $Rows
if($ROWS -eq 1) { WriteWordLine 0 0 "No Services have been configured"} else {
    $Service | foreach { 
        $RowsC++
        $Y = ($_ -replace 'add service ', '').split()
        WriteWordLine 2 0 "Service $($Y[0]) "
      
        $TableRange = $doc.Application.Selection.Range
        [int]$Columns = 7
        [int]$Rowst = 2
        Write-Verbose "$(Get-Date): `t`tTable: $RowsC/$RowsTotal Write Service Tables $($Y[0])"
		$Table = $doc.Tables.Add($TableRange, $Rowst, $Columns)
		$table.Style = "Table Grid"
		$table.Borders.InsideLineStyle = $wdLineStyleSingle
		$table.Borders.OutsideLineStyle = $wdLineStyleSingle

        $Col = 0
        do {
            $Col++
            $Table.Cell(1,$Col).Shading.BackgroundPatternColor = $wdColorGray15
	        $Table.Cell(1,$Col).Range.Font.Bold = $True
            }
        while ($Col -le $Columns -1)

		$Table.Cell(1,1).Range.Text = "Server"
		$Table.Cell(1,2).Range.Text = "Protocol"
		$Table.Cell(1,3).Range.Text = "Port"
		$Table.Cell(1,4).Range.Text = "Traffic Domain"
		$Table.Cell(1,5).Range.Text = "GSLB"
		$Table.Cell(1,6).Range.Text = "Maximum Clients"
		$Table.Cell(1,7).Range.Text = "Maximum Requests"

        $Table.Cell(2,1).Range.Font.size = 9
		$Table.Cell(2,1).Range.Text = $Y[1]
        $Table.Cell(2,2).Range.Font.size = 9
		$Table.Cell(2,2).Range.Text = $Y[2]
        $Table.Cell(2,3).Range.Font.size = 9
		$Table.Cell(2,3).Range.Text = $Y[3]
        $Table.Cell(2,4).Range.Font.size = 9
		$Table.Cell(2,4).Range.Text = Get-StringProperty $_ "-td" "0 (Default)";
        $Table.Cell(2,5).Range.Font.size = 9
		$Table.Cell(2,5).Range.Text = Get-StringProperty $_ "-gslb" "NA";
        $Table.Cell(2,6).Range.Font.size = 9
		$Table.Cell(2,6).Range.Text = Get-StringProperty $_ "-maxClient" "NA";
        $Table.Cell(2,7).Range.Font.size = 9
        $Table.Cell(2,7).Range.Text = Get-StringProperty $_ "-maxreq" "NA";
       
        $table.AutoFitBehavior($wdAutoFitContent)

        #return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
        WriteWordLine 0 0 "Table: Service basic configuration"
        WriteWordLine 0 0 " "

        $LBNAMEBIND1 = "bind service $($Y[0]) -monitorName *"

        $ROWS2 = 1
        $ServiceBind | foreach {
                if ($_ -like $LBNAMEBIND1) {
                    $ROWS2++
                    }
                }

        WriteWordLine 0 0 " "
        WriteWordLine 4 0 "Monitoring"
        if ($ROWS2 -eq 1) {WriteWordLine 0 0 "No Monitors have been configured for this service"} else {
            
            $TableRange = $doc.Application.Selection.Range
            [int]$Columns2 = 1
            [int]$Rows2 = $ROWS2
                
            $Table = $doc.Tables.Add($TableRange, $Rows2, $Columns2)
		    $table.Style = "Table Grid"
			$table.Borders.InsideLineStyle = $wdLineStyleSingle
			$table.Borders.OutsideLineStyle = $wdLineStyleSingle
            $Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		    $Table.Cell(1,1).Range.Font.Bold = $True
		    $Table.Cell(1,1).Range.Text = "Monitor"
            $xRow2 = 1

            $ServiceBind | foreach { 
                if ($_ -like $LBNAMEBIND1) {          
                    $xRow2++
                    $Table.Cell($xRow2,1).Range.Font.size = 9
		            $Table.Cell($xRow2,1).Range.Text = Get-StringProperty $_ "-monitorName";
                    }              
                }
        
            $table.AutoFitBehavior($wdAutoFitContent)

            #return focus back to document
		    $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

		    #move to the end of the current document
		    $selection.EndKey($wdStory,$wdMove) | Out-Null
            WriteWordLine 0 0 "Table: Service Monitor"    
            }

        WriteWordLine 0 0 " "
        WriteWordLine 0 0 " "
        WriteWordLine 4 0 "Advanced Configuration"
        WriteWordLine 0 0 "Clear text port`t`t`t: $(Get-StringProperty $_ "-clearTextPort" "NA")";
        WriteWordLine 0 0 "Cache Type `t`t`t: $(Get-StringProperty $_ "-cacheType" "NA")";
        WriteWordLine 0 0 "Maximum Client Requests`t: $(Get-StringProperty $_ "-maxClient" "4294967294 (Maximum Value)")";
        WriteWordLine 0 0 "Monitor health of this service`t: $(Test-NotStringPropertyYesNo $_ "-healthMonitor")";
        WriteWordLine 0 0 "Maximum Requests`t`t: $(Get-StringProperty $_ "-maxreq" "65535 (Maximum Value)")";
        WriteWordLine 0 0 "Use Transparent Cache`t`t: $(Test-StringPropertyYesNo $_ "-cacheable")";
        WriteWordLine 0 0 "Insert the Client IP header`t: $(Get-StringProperty $_ "-cip" "NA")";
        WriteWordLine 0 0 "Name for the HTTP header`t: $(Get-StringProperty $_ "-cipHeader" "NA")";
        WriteWordLine 0 0 "Use Source IP`t`t`t: $(Test-StringPropertyYesNo $_ "-usip")";     
        WriteWordLine 0 0 "Path Monitoring`t`t: $(Test-StringPropertyYesNo $_ "-pathMonitor")";
        WriteWordLine 0 0 "Individual Path monitoring`t: $(Test-StringPropertyYesNo $_ "-pathMonitorIndv")";
        WriteWordLine 0 0 "Use the proxy port`t`t: $(Test-StringPropertyYesNo $_ "-useproxyport")";
        WriteWordLine 0 0 "SureConnect`t`t`t: $(Test-StringPropertyOnOff $_ "-sc")";
        WriteWordLine 0 0 "surge protection`t`t: $(Test-StringPropertyOnOff $_ "-sp")";
        WriteWordLine 0 0 "RTSP session ID mapping`t: $(Test-StringPropertyOnOff $_ "-rtspSessionidRemap")";
        WriteWordLine 0 0 "Client Time-Out`t`t`t: $(Get-StringProperty $_ "-cltTimeout" "31536000 (Maximum Value)")";
        WriteWordLine 0 0 "Server Time-Out`t`t: $(Get-StringProperty $_ "-svrTimeout" "3153600 (Maximum Value)")";
        WriteWordLine 0 0 "Unique identifier for the service`t: $(Get-StringProperty $_ "-CustomServerID" "None")";
        WriteWordLine 0 0 "The identifier for the service.`t: $(Get-StringProperty $_ "-serverID" "None")";
        WriteWordLine 0 0 "Enable client keep-alive`t`t: $(Test-StringPropertyYesNo $_ "-CKA")";
        WriteWordLine 0 0 "Enable TCP buffering`t`t: $(Test-StringPropertyYesNo $_ "-TCPB")";
        WriteWordLine 0 0 "Enable compression`t`t: $(Test-StringPropertyYesNo $_ "-CMP")";
        WriteWordLine 0 0 "Maximum bandwidth, in Kbps`t: $(Get-StringProperty $_ "-maxBandwidth" "4294967287 (Maximum Value)")";
        WriteWordLine 0 0 "Use Layer 2 mode`t`t: $(Test-StringPropertyYesNo $_ "-accessDown")";
        WriteWordLine 0 0 "Sum of weights of the monitors`t: $(Get-StringProperty $_ "-monThreshold" "65535 (Maximum Value)")";
        WriteWordLine 0 0 "Initial state of the service`t: $(Test-NotStringPropertyEnabledDisabled $_ "-state")";
        WriteWordLine 0 0 "Perform delayed clean-up`t: $(Test-NotStringPropertyEnabledDisabled $_ "-downStateFlush")";
        WriteWordLine 0 0 "TCP profile`t`t`t: $(Get-StringProperty $_ "-tcppProfileName" "NA")";
        WriteWordLine 0 0 "HTTP profile`t`t`t: $(Get-StringProperty $_ "-httpProfileName" "NA")";
        WriteWordLine 0 0 "A numerical identifier`t`t: $(Get-StringProperty $_ "-hashId" "NA")";
        WriteWordLine 0 0 "Comment about the service`t: $(Get-StringProperty $_ "-comment" "NA")";
        WriteWordLine 0 0 "Logging of AppFlow information`t: $(Test-NotStringPropertyEnabledDisabled $_ "-appflowLog")";
        WriteWordLine 0 0 "Network profile`t`t`t: $(Get-StringProperty $_ "-netProfile" "NA")";

        WriteWordLine 0 0 " "

        $selection.InsertNewPage() 
        WriteWordLine 0 0 " "
        }
   }

#endregion NetScaler Services

#region NetScaler Service Groups
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Service Groups"
$selection.InsertNewPage()

WriteWordLine 2 0 "NetScaler Service Groups"
WriteWordLine 0 0 " "

$Rows = 1
$ServiceGroup | foreach {
    $Rows++
    }
$RowsC = 0
$RowsTotal = $Rows
if($ROWS -eq 1) { WriteWordLine 0 0 "No Service Groups have been configured"} else {
    $ServiceGroup | foreach { 
        $RowsC++
        $Y = ($_ -replace 'add serviceGroup ', '').split()
      
        WriteWordLine 2 0 "Service Group $($Y[0]) "
      
        $TableRange = $doc.Application.Selection.Range
        [int]$Columns = 4
        [int]$Rowst = 2
        
        Write-Verbose "$(Get-Date): `t`tTable: $RowsC/$RowsTotal Write ServiceGroup Table $($Y[0])"
		$Table = $doc.Tables.Add($TableRange, $Rowst, $Columns)
		$table.Style = "Table Grid"
		$table.Borders.InsideLineStyle = $wdLineStyleSingle
		$table.Borders.OutsideLineStyle = $wdLineStyleSingle
        $xRow = 1
        
        $COL = 0
        do {
            $COL++
		    $Table.Cell(1,$COL).Shading.BackgroundPatternColor = $wdColorGray15
		    $Table.Cell(1,$COL).Range.Font.Bold = $True
            }
        while ($COL -le $Columns -1)

        $Table.Cell(1,1).Range.Text = "State"
		$Table.Cell(1,2).Range.Text = "Service Type"
		$Table.Cell(1,3).Range.Text = "Traffic Domain"
        $Table.Cell(1,4).Range.Text = "Use Source IP"
        
        $COL = 0
        do {
            $COL++
            $Table.Cell(2,$col).Range.Font.size = 9
            }
        while ($COL -le $Columns -1)

        $Table.Cell(2,1).Range.Text = Test-StringPropertyEnabledDisabled $_ "-state";
        $Table.Cell(2,2).Range.Text = $Y[1];
		$Table.Cell(2,3).Range.Text = Get-StringProperty $_ "-td" "0 (Default)";
        $Table.Cell(2,4).Range.Text = Test-NotStringPropertyYesNo $_ "-usip";
        
        $table.AutoFitBehavior($wdAutoFitContent)

        #return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
        WriteWordLine 0 0 "Table: Service basic configuration"
        WriteWordLine 0 0 " "

        WriteWordLine 4 0 "Monitoring"
      
        $LBNAMEBIND1 = "bind service $($Y[0]) -monitorName *"

        $ROWS2 = 1
        $ServiceBind | foreach {
            if ($_ -like $LBNAMEBIND1) {
                $ROWS2++
                }
            }

        if ($ROWS2 -eq 1) {WriteWordLine 0 0 "No Monitors have been configured for this service"} else {

            $TableRange = $doc.Application.Selection.Range
            [int]$Columns2 = 1
            [int]$Rows2 = $ROWS2
                
            $Table = $doc.Tables.Add($TableRange, $Rows2, $Columns2)
		    $table.Style = "Table Grid"
			$table.Borders.InsideLineStyle = $wdLineStyleSingle
			$table.Borders.OutsideLineStyle = $wdLineStyleSingle
            $Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		    $Table.Cell(1,1).Range.Font.Bold = $True
		    $Table.Cell(1,1).Range.Text = "Monitor"
            $xRow2 = 1

            $ServiceBind | foreach { 
                if ($_ -like $LBNAMEBIND1) {          
                    $xRow2++
                    $Table.Cell($xRow2,1).Range.Font.size = 9
		            $Table.Cell($xRow2,1).Range.Text = Get-StringProperty $_ "-monitorName"
                }              
            }
        
            $table.AutoFitBehavior($wdAutoFitContent)

            #return focus back to document
		    $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

		    #move to the end of the current document
		    $selection.EndKey($wdStory,$wdMove) | Out-Null
            WriteWordLine 0 0 "Table: Service Monitor"    
            }
        WriteWordLine 0 0 " " 

        WriteWordLine 4 0 "Advanced Configuration"

        WriteWordLine 0 0 "Global Server Load Balancing`t: $(Test-StringPropertyOnOff $_ "-gslb")";
        WriteWordLine 0 0 "Clear text port`t`t`t: $(Get-StringProperty $_ "-clearTextPort" "NA")";
        WriteWordLine 0 0 "Cache Type `t`t`t: $(Get-StringProperty $_ "-cacheType" "NA")";
        WriteWordLine 0 0 "Maximum Client Requests`t: $(Get-StringProperty $_ "-maxClient" "4294967294 (Maximum Value)")";
        WriteWordLine 0 0 "Monitor health of this service`t: $(Test-NotStringPropertyYesNo $_ "-healthMonitor")";
        WriteWordLine 0 0 "Maximum Requests`t`t: $(Get-StringProperty $_ "-maxreq" "65535 (Maximum Value)")";
        WriteWordLine 0 0 "Use Transparent Cache`t`t: $(Test-StringPropertyYesNo $_ "-cacheable")";
        WriteWordLine 0 0 "Insert the Client IP header`t: $(Get-StringProperty $_ "-cip" "NA")";
        WriteWordLine 0 0 "Name for the HTTP header`t: $(Get-StringProperty "-cipHeader" "NA")";
        WriteWordLine 0 0 "Use Source IP`t`t`t: $(Test-StringPropertyYesNo $_ "-usip")";
        WriteWordLine 0 0 "Path Monitoring`t`t: $(Test-StringPropertyYesNo $_ "-pathMonitor")";
        WriteWordLine 0 0 "Individual Path monitoring`t: $(Test-StringPropertyYesNo $_ "-pathMonitorIndv")";
        WriteWordLine 0 0 "Use the proxy port`t`t: $(Test-StringPropertyYesNo $_ "-useproxyport")";
        WriteWordLine 0 0 "SureConnect`t`t`t: $(Test-StringPropertyOnOff $_ "-sc")";
        WriteWordLine 0 0 "surge protection`t`t: $(Test-StringPropertyOnOff $_ "-sp")";
        WriteWordLine 0 0 "RTSP session ID mapping`t: $(Test-StringPropertyOnOff $_ "-rtspSessionidRemap")";
        WriteWordLine 0 0 "Client Time-Out`t`t`t: $(Get-StringProperty $_ "-cltTimeout" "31536000 (Maximum Value)")";
        WriteWordLine 0 0 "Server Time-Out`t`t: $(Get-StringProperty $_ "-svrTimeout" "31536000 (Maximum Value)")";
        WriteWordLine 0 0 "Unique identifier for the service`t: $(Get-StringProperty $_ "-CustomServiceID" "None")";
        WriteWordLine 0 0 "The identifier for the service.`t: $(Get-StringProperty $_ "-serverID" "None")";
        WriteWordLine 0 0 "Enable client keep-alive`t`t: $(Test-StringPropertyYesNo $_ "-CKA")";
        WriteWordLine 0 0 "Enable TCP buffering`t`t: $(Test-StringPropertyYesNo $_ "-TCPB")";
        WriteWordLine 0 0 "Enable compression`t`t: $(Test-StringPropertyYesNo $_ "-CMP")";
        WriteWordLine 0 0 "Maximum bandwidth, in Kbps`t: $(Get-StringProperty $_ "-maxBandwidth" "4294967287 (Maximum Value)")";
        WriteWordLine 0 0 "Use Layer 2 mode`t`t: $(Test-StringPropertyYesNo $_ "-accessDown")";
        WriteWordLine 0 0 "Sum of weights of the monitors`t: $(Get-StringProperty $_ "-monThreshold" "65535 (Maximum Value)")";
        WriteWordLine 0 0 "Initial state of the service`t: $(Test-NotStringPropertyEnabledDisabled $_ "-state")";
        WriteWordLine 0 0 "Perform delayed clean-up`t: $(Test-NotStringPropertyEnabledDisabled $_ "-downStateFlush")";
        WriteWordLine 0 0 "TCP profile`t`t`t: $(Get-StringProperty $_ "-tcpProfileName" "NA")";
        WriteWordLine 0 0 "HTTP profile`t`t`t: $(Get-StringProperty $_ "-httpProfileName" "NA")";
        WriteWordLine 0 0 "A numerical identifier`t`t: $(Get-StringProperty $_ "-hashId" "NA")";
        WriteWordLine 0 0 "Comment`t`t`t: $(Get-StringProperty $_ "-comment" "NA")";
        WriteWordLine 0 0 "Logging of AppFlow information`t: $(Test-NotStringPropertyEnabledDisabled $_ "-appflowLog")";
        WriteWordLine 0 0 "Network profile`t`t`t: $(Get-StringProperty $_ "-netProfile" "NA")";

        WriteWordLine 0 0 " "

        $selection.InsertNewPage()
        WriteWordLine 0 0 " "
   }
}

#endregion NetScaler Service Groups

#region NetScaler Servers
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Servers"

$selection.InsertNewPage()

WriteWordLine 1 0 "NetScaler Servers"

$Rows = 1
$Server | foreach { 
    $Rows++
    }

If ($Rows -eq 1) {WriteWordLine 0 0 "No Servers have been configured"} Else {
    
    $TableRange = $doc.Application.Selection.Range
    [int]$Columns = 5
    [int]$Rows = $Rows
    $xRow = 1
    
    Write-Verbose "$(Get-Date): `t`tTable: Write Server Table"
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = "Table Grid"
	$table.Borders.InsideLineStyle = $wdLineStyleSingle
	$table.Borders.OutsideLineStyle = $wdLineStyleSingle
    $Col = 0
    do {
        $Col++
        $Table.Cell(1,$Col).Shading.BackgroundPatternColor = $wdColorGray15
	    $Table.Cell(1,$Col).Range.Font.Bold = $True
        }
    while ($Col -le $Columns -1)	

	$Table.Cell(1,1).Range.Text = "Server"
	$Table.Cell(1,2).Range.Text = "IP Address"
	$Table.Cell(1,3).Range.Text = "Traffic Domain"
	$Table.Cell(1,4).Range.Text = "State"
	$Table.Cell(1,5).Range.Text = "Comment"

    $Server | foreach { 
        $Y = ($_ -replace 'add server ', '').split()
        $xRow++  
        $Table.Cell($xRow,1).Range.Font.size = 9
		$Table.Cell($xRow,1).Range.Text = $Y[0]
        $Table.Cell($xRow,2).Range.Font.size = 9
        $Table.Cell($xRow,2).Range.Text = $Y[1]
        $Table.Cell($xRow,3).Range.Font.size = 9
        $Table.Cell($xRow,3).Range.Text = Get-StringProperty $_ "-td" "0 (Default)";
        $Table.Cell($xRow,4).Range.Font.size = 9
        $Table.Cell($xRow,4).Range.Text = Test-NotStringPropertyEnabledDisabled $_ "-state";
        $Table.Cell($xRow,5).Range.Font.size = 9
        $Table.Cell($xRow,5).Range.Text = Get-StringProperty $_ "-comment" "No Comments";
        }
    
    $table.AutoFitBehavior($wdAutoFitContent)

    #return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
    WriteWordLine 0 0 "Table: Server configuration"
    WriteWordLine 0 0 " "
    }
            
#endregion NetScaler Servers

#region NetScaler Monitors
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Monitors"

$selection.InsertNewPage()

WriteWordLine 1 0 "Custom NetScaler Monitors"

$ROWS = 1
$Monitor | foreach {
    $Rows++
    }

if ($ROWS -eq 1) {WriteWordLine 0 0 "No custom monitors have been configured for this service"} else {

    $TableRange = $doc.Application.Selection.Range
    [int]$Columns = 8
    [int]$Rows = $ROWS
    $xRow = 1
    
    Write-Verbose "$(Get-Date): `t`tTable: Write Custom Monitors table"
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = "Table Grid"
	$table.Borders.InsideLineStyle = $wdLineStyleSingle
	$table.Borders.OutsideLineStyle = $wdLineStyleSingle

    $Col = 0
    do {
        $Col++
        $Table.Cell(1,$Col).Shading.BackgroundPatternColor = $wdColorGray15
	    $Table.Cell(1,$Col).Range.Font.Bold = $True
        $Table.Cell(1,$Col).Range.Font.size = 8
        }
    while ($Col -le $Columns -1)

	$Table.Cell(1,1).Range.Text = "Monitor Name"
	$Table.Cell(1,2).Range.Text = "Protocol"
	$Table.Cell(1,3).Range.Text = "HTTP Request"
	$Table.Cell(1,4).Range.Text = "Destination IP"
	$Table.Cell(1,5).Range.Text = "Destination Port"
	$Table.Cell(1,6).Range.Text = "Interval"
	$Table.Cell(1,7).Range.Text = "Response Code"
	$Table.Cell(1,8).Range.Text = "Time-Out"

    $Monitor | foreach { 
        $xRow++
        $Y = ($_ -replace 'add lb monitor ', '').split()
        $Table.Cell($xRow,1).Range.Font.size = 8
		$Table.Cell($xRow,1).Range.Text = $Y[0]
        $Table.Cell($xRow,2).Range.Font.size = 8
        $Table.Cell($xRow,2).Range.Text = $Y[1]
        $Table.Cell($xRow,3).Range.Font.size = 8
        $Table.Cell($xRow,4).Range.Font.size = 8
        $Table.Cell($xRow,5).Range.Font.size = 8
        $Table.Cell($xRow,6).Range.Font.size = 8
        $Table.Cell($xRow,7).Range.Font.size = 8
        $Table.Cell($xRow,8).Range.Font.size = 8
            
        $Table.Cell($xRow,3).Range.Text = Get-StringProperty $_ "-httpRequest" "NA";
        $Table.Cell($xRow,4).Range.Text = Get-StringProperty $_ "-destIP" "NA";
        $Table.Cell($xRow,5).Range.Text = Get-StringProperty $_ "-destPort" "NA";
        $Table.Cell($xRow,6).Range.Text = Get-StringProperty $_ "-interval" "NA";
        $Table.Cell($xRow,7).Range.Text = Get-StringProperty $_ "-respCode" "NA";
        $Table.Cell($xRow,8).Range.Text = Get-StringProperty $_ "-resptimeout" "NA";
        }
    
    $table.AutoFitBehavior($wdAutoFitContent)

    #return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
    WriteWordLine 0 0 "Table: Custom Monitors"
    WriteWordLine 0 0 " "
    }

$selection.InsertNewPage()

#endregion NetScaler Monitors

#region NetScaler Policies
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Policies"

WriteWordLine 1 0 "NetScaler Policies"

## Work in Progress: Binding to actions and binding to vServers

WriteWordLine 2 0 "NetScaler Pattern Set Policies"

#Policy Pattern Set
$ROWS = 1
$Add | foreach {
        if ($_ -like 'add policy patset*') {
            $Rows++
        }
}

if ($ROWS -eq 1) {WriteWordLine 0 0 "No Pattern Set Policies have been configured"} else {
    $TableRange = $doc.Application.Selection.Range
    [int]$Columns = 1
    [int]$Rows = $ROWS
    $xRow = 1

    Write-Verbose "$(Get-Date): `t`tTable: Write Pattern Set Policies table"
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = "Table Grid"
	$table.Borders.InsideLineStyle = $wdLineStyleSingle
	$table.Borders.OutsideLineStyle = $wdLineStyleSingle
	$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell(1,1).Range.Font.Bold = $True
	$Table.Cell(1,1).Range.Text = "Pattern Set Policy"

    $Add | foreach {
        if ($_ -like 'add policy patset*') {
            $xRow++
            $Y = ($_ -replace 'add policy patset', '').split()
            $Table.Cell($xRow,1).Range.Font.size = 9
		    $Table.Cell($xRow,1).Range.Text = $Y[1]
            }
        }
    $table.AutoFitBehavior($wdAutoFitContent)

    #return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
    WriteWordLine 0 0 "Table: NetScaler Pattern Set Policies"
    WriteWordLine 0 0 " "
    }

WriteWordLine 2 0 "NetScaler Responder Policies"

#Policy Type Responder
$ROWS = 1
$Add | foreach {
    if ($_ -like 'add responder policy*') {
        $ROWS++
        }
    }

if ($ROWS -eq 1) {WriteWordLine 0 0 "No Responder Policies have been configured"} else {
    $TableRange = $doc.Application.Selection.Range
    [int]$Columns = 1
    [int]$Rows = $ROWS
    $xRow = 1

    Write-Verbose "$(Get-Date): `t`tTable: Responder Policies table"
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = "Table Grid"
	$table.Borders.InsideLineStyle = $wdLineStyleSingle
	$table.Borders.OutsideLineStyle = $wdLineStyleSingle
	$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell(1,1).Range.Font.Bold = $True
	$Table.Cell(1,1).Range.Text = "Responder Policy"

    $Add | foreach {
        if ($_ -like 'add responder policy*') {
            $xRow++
            $Y = ($_ -replace 'add responder policy ', '').split()
            $Table.Cell($xRow,1).Range.Font.size = 9
		    $Table.Cell($xRow,1).Range.Text = $Y[0]
            }
        }
    $table.AutoFitBehavior($wdAutoFitContent)

    #return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
    WriteWordLine 0 0 "Table: NetScaler Responder Policies"
    WriteWordLine 0 0 " "
    }

WriteWordLine 2 0 "NetScaler Rewrite Policies"

#Policy Type Rewrite
$ROWS = 1
$Add | foreach {
    if ($_ -like 'add rewrite policy *') {
        $ROWS++
        }
    }

if ($ROWS -eq 1) {WriteWordLine 0 0 "No Rewrite Policies have been configured"} else {
    $TableRange = $doc.Application.Selection.Range
    [int]$Columns = 3
    [int]$Rows = $ROWS
    $xRow = 1

    Write-Verbose "$(Get-Date): `t`tTable: Rewrite Policies table"
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = "Table Grid"
	$table.Borders.InsideLineStyle = $wdLineStyleSingle
	$table.Borders.OutsideLineStyle = $wdLineStyleSingle
	
$COL = 0
    do {
        $COL++
        $Table.Cell(1,$COL).Shading.BackgroundPatternColor = $wdColorGray15
        $Table.Cell(1,$COL).Range.Font.Bold = $True
        }
    while ($Col -le $Columns -1)
	$Table.Cell(1,1).Range.Text = "Rewrite Policy"
	$Table.Cell(1,2).Range.Text = "Rule"
	$Table.Cell(1,3).Range.Text = "Undefined"

    $Add | foreach {
        if ($_ -like 'add rewrite policy *') {
            $xRow++
            $Y = ($_ -replace 'add rewrite policy ', '').split()
            $Table.Cell($xRow,1).Range.Font.size = 9
		    $Table.Cell($xRow,1).Range.Text = $Y[0]
            $Table.Cell($xRow,2).Range.Font.size = 9
		    $Table.Cell($xRow,2).Range.Text = $Y[1]
            $Table.Cell($xRow,3).Range.Font.size = 9
		    $Table.Cell($xRow,3).Range.Text = $Y[2]
            }
        }
    $table.AutoFitBehavior($wdAutoFitContent)

    #return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
    WriteWordLine 0 0 "Table: NetScaler Rewrite Policies"
    WriteWordLine 0 0 " "
    }

$selection.InsertNewPage()

#endregion NetScaler Policies

#region NetScaler Actions
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Actions"

WriteWordLine 1 0 "NetScaler Actions"

## Work in Progress: Binding to policies

WriteWordLine 2 0 "NetScaler Pattern Set Action"

#Action Type patset
$ROWS = 1
$Add | foreach {
    if ($_ -like 'add action patset *') {
        $ROWS++
        }
    }

WriteWordLine 0 0 " "

if ($ROWS -eq 1) {WriteWordLine 0 0 "No Pattern Set Actions have been configured"} else {
    $TableRange = $doc.Application.Selection.Range
    [int]$Columns = 3
    [int]$Rows = $ROWS
    $xRow = 1

    Write-Verbose "$(Get-Date): `t`tTable: Pattern Set Action table"
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = "Table Grid"
	$table.Borders.InsideLineStyle = $wdLineStyleSingle
	$table.Borders.OutsideLineStyle = $wdLineStyleSingle
	
$COL = 0
    do {
        $COL++
        $Table.Cell(1,$COL).Shading.BackgroundPatternColor = $wdColorGray15
        $Table.Cell(1,$COL).Range.Font.Bold = $True
        }
    while ($Col -le $Columns -1)
	$Table.Cell(1,1).Range.Text = "Pattern Set"
	$Table.Cell(1,2).Range.Text = "Rule"
	$Table.Cell(1,3).Range.Text = "Undefined"

    $Add | foreach {
        if ($_ -like 'add action patset *') {
            $xRow++
            $Y = ($_ -replace 'add action patset ', '').split()
            $Table.Cell($xRow,1).Range.Font.size = 9
		    $Table.Cell($xRow,1).Range.Text = $Y[1]
            $Table.Cell($xRow,2).Range.Font.size = 9
		    $Table.Cell($xRow,2).Range.Text = $Y[1]
            $Table.Cell($xRow,3).Range.Font.size = 9
		    $Table.Cell($xRow,3).Range.Text = $Y[2]
            }
        }
    $table.AutoFitBehavior($wdAutoFitContent)

    #return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
    WriteWordLine 0 0 "Table: NetScaler Pattern Set Actions"
    WriteWordLine 0 0 " "
    }

WriteWordLine 0 0 " "

#Action Type Responder

WriteWordLine 2 0 "NetScaler Responder Action"

$ROWS = 1
$Add | foreach {
    if ($_ -like 'add responder action*') {
        $ROWS++
        }
    }

if ($ROWS -eq 1) {WriteWordLine 0 0 "No Responder Actions have been configured"} else {
    $TableRange = $doc.Application.Selection.Range
    [int]$Columns = 3
    [int]$Rows = $ROWS
    $xRow = 1

    Write-Verbose "$(Get-Date): `t`tTable: Responder Set Action table"
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = "Table Grid"
	$table.Borders.InsideLineStyle = $wdLineStyleSingle
	$table.Borders.OutsideLineStyle = $wdLineStyleSingle
	$COL = 0
    do {
        $COL++
        $Table.Cell(1,$COL).Shading.BackgroundPatternColor = $wdColorGray15
        $Table.Cell(1,$COL).Range.Font.Bold = $True
        }
    while ($Col -le $Columns -1)
	$Table.Cell(1,1).Range.Text = "Responder"
	$Table.Cell(1,2).Range.Text = "Rule"
	$Table.Cell(1,3).Range.Text = "Undefined"

    $Add | foreach {
        if ($_ -like 'add responder action*') {
            $xRow++
            $Y = ($_ -replace 'add responder action ', '').split()
            $Table.Cell($xRow,1).Range.Font.size = 9
		    $Table.Cell($xRow,1).Range.Text = $Y[0]
            $Table.Cell($xRow,2).Range.Font.size = 9
		    $Table.Cell($xRow,2).Range.Text = $Y[1]
            $Table.Cell($xRow,3).Range.Font.size = 9
		    $Table.Cell($xRow,3).Range.Text = $Y[2]
            }
        }
    $table.AutoFitBehavior($wdAutoFitContent)

    #return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
    WriteWordLine 0 0 "Table: NetScaler Responder Actions"
    WriteWordLine 0 0 " "
    }

#Action Type Rewrite
WriteWordLine 2 0 "NetScaler Rewrite Action"

$ROWS = 1
$Add | foreach {
    if ($_ -like 'add rewrite action*') {
        $ROWS++
        }
    }

if ($ROWS -eq 1) {WriteWordLine 0 0 "No Rewrite Actions have been configured"} else {
    $TableRange = $doc.Application.Selection.Range
    [int]$Columns = 3
    [int]$Rows = $ROWS
    $xRow = 1

    Write-Verbose "$(Get-Date): `t`tTable: Rewrite Set Action table"
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = "Table Grid"
	$table.Borders.InsideLineStyle = $wdLineStyleNone
	$table.Borders.OutsideLineStyle = $wdLineStyleSingle
	$COL = 0
    do {
        $COL++
        $Table.Cell(1,$COL).Shading.BackgroundPatternColor = $wdColorGray15
        $Table.Cell(1,$COL).Range.Font.Bold = $True
        }
    while ($Col -le $Columns -1)

    $Table.Cell(1,1).Range.Text = "Rewrite Action"
	$Table.Cell(1,2).Range.Text = "Rule"
	$Table.Cell(1,3).Range.Text = "Undefined"

    $Add | foreach {
        if ($_ -like 'add rewrite action*') {
            $xRow++
            $Y = ($_ -replace 'add rewrite action ', '').split()
            $Table.Cell($xRow,1).Range.Font.size = 9
		    $Table.Cell($xRow,1).Range.Text = $Y[0]
            $Table.Cell($xRow,2).Range.Font.size = 9
		    $Table.Cell($xRow,2).Range.Text = $Y[1]
            $Table.Cell($xRow,3).Range.Font.size = 9
		    $Table.Cell($xRow,3).Range.Text = $Y[2]
            }
        }
    $table.AutoFitBehavior($wdAutoFitContent)

    #return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
    WriteWordLine 0 0 "Table: NetScaler Rewrite Actions"
    WriteWordLine 0 0 " "
    }

$selection.InsertNewPage()

#endregion NetScaler Actions

#region NetScaler Profiles
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Profiles"

WriteWordLine 1 0 "NetScaler Profiles"

WriteWordLine 2 0 "NetScaler TCP Profiles"

#TCP Profile
$ROWS = 1
$SetNS | foreach {
    if ($_ -like 'set ns tcpProfile*') {
        $ROWS++
        }
    }

if ($ROWS -eq 1) {WriteWordLine 0 0 "No TCP Profiles have been configured"} else {
    $TableRange = $doc.Application.Selection.Range
    [int]$Columns = 5
    [int]$Rows = $ROWS
    $xRow = 1
    Write-Verbose "$(Get-Date): `t`tTable: TCP Profile table"
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = "Table Grid"
	$table.Borders.InsideLineStyle = $wdLineStyleSingle
	$table.Borders.OutsideLineStyle = $wdLineStyleSingle
    
    $COL = 0
    do {
        $COL++
        $Table.Cell(1,$COL).Shading.BackgroundPatternColor = $wdColorGray15
        $Table.Cell(1,$COL).Range.Font.Bold = $True
        }
    while ($Col -le $Columns -1)

	$Table.Cell(1,1).Range.Text = "TCP Profile"
	$Table.Cell(1,2).Range.Text = "WS"
	$Table.Cell(1,3).Range.Text = "SACK"
	$Table.Cell(1,4).Range.Text = "NAGLE"
	$Table.Cell(1,5).Range.Text = "MSS"

    $SetNS | foreach {
        if ($_ -like 'set ns tcpProfile*') {
            $xRow++
            $Y = ($_ -replace 'set ns tcpProfile ', '').split()
            $Table.Cell($xRow,1).Range.Font.size = 9
		    $Table.Cell($xRow,1).Range.Text = $Y[0]
            $Table.Cell($xRow,2).Range.Font.size = 9
            $Table.Cell($xRow,3).Range.Font.size = 9
            $Table.Cell($xRow,4).Range.Font.size = 9
            $Table.Cell($xRow,5).Range.Font.size = 9
            
            $Table.Cell($xRow,2).Range.Text = Get-StringProperty $_ "-WS" "NA";
            $Table.Cell($xRow,3).Range.Text = Get-StringProperty $_ "-SACK" "NA";
            $Table.Cell($xRow,4).Range.Text = Get-StringProperty $_ "-NAGLE" "NA";
            $Table.Cell($xRow,5).Range.Text = Get-StringProperty $_ "-MSS" "NA";
            }
        }
    $table.AutoFitBehavior($wdAutoFitContent)

    #return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
    WriteWordLine 0 0 "Table: NetScaler TCP Profiles"
    WriteWordLine 0 0 " "
    }

#HTTP Profile

WriteWordLine 2 0 "NetScaler HTTP Profiles"

$ROWS = 1
$Add | foreach {
    if ($_ -like 'add ns httpProfile*') {
        $ROWS++
        }
    }

if ($ROWS -eq 1) {WriteWordLine 0 0 "No HTTP Profiles have been configured"} else {
    $TableRange = $doc.Application.Selection.Range
    [int]$Columns = 3
    [int]$Rows = $ROWS
    $xRow = 1
    Write-Verbose "$(Get-Date): `t`tTable: HTTP Profile table"
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = "Table Grid"
	$table.Borders.InsideLineStyle = $wdLineStyleSingle
	$table.Borders.OutsideLineStyle = $wdLineStyleSingle

    $COL = 0
    do {
        $COL++
        $Table.Cell(1,$COL).Shading.BackgroundPatternColor = $wdColorGray15
        $Table.Cell(1,$COL).Range.Font.Bold = $True
        }
    while ($Col -le $Columns -1)

	$Table.Cell(1,1).Range.Text = "HTTP Profile"
	$Table.Cell(1,2).Range.Text = "Drop Invalid Requests"
	$Table.Cell(1,3).Range.Text = "SPDY"

    $Add | foreach {
        if ($_ -like 'add ns httpProfile*') {
            $xRow++
            $Y = ($_ -replace 'add ns httpProfile ', '').split()
            $Table.Cell($xRow,1).Range.Font.size = 9
		    $Table.Cell($xRow,1).Range.Text = $($Y[0])
            $Table.Cell($xRow,2).Range.Font.size = 9
            $Table.Cell($xRow,2).Range.Text = Get-StringProperty $_ "-dropInvalReqs" "Disabled";
            $Table.Cell($xRow,3).Range.Font.size = 9
            $Table.Cell($xRow,3).Range.Text = Get-StringProperty $_ "-spdy" "Disabled";
            }
        }
    $table.AutoFitBehavior($wdAutoFitContent)

    #return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
    WriteWordLine 0 0 "Table: NetScaler HTTP Profiles"
    WriteWordLine 0 0 " "
    }



#endregion NetScaler Profiles

#region Statistics

## Statistics
$sw.Stop()

#endregion Statistics

#endregion NetScaler Documentation Build

#region Finalize Word Document

Write-Verbose "$(Get-Date): Finishing up Word document"
#end of document processing
#Update document properties

If($CoverPagesExist)
{
	Write-Verbose "$(Get-Date): Set Cover Page Properties"
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Company" $CompanyName
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Title" $title
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Subject" "Inventory"
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Author" $username

	#Get the Coverpage XML part
	$cp = $doc.CustomXMLParts | where {$_.NamespaceURI -match "coverPageProps$"}

	#get the abstract XML part
	$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}
	#set the text
	[string]$abstract = "Inventory for $CompanyName"
	$ab.Text = $abstract

	$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "PublishDate"}
	#set the text
	[string]$abstract = (Get-Date -Format d).ToString()
	$ab.Text = $abstract

	Write-Verbose "$(Get-Date): Update the Table of Contents"
	#update the Table of Contents
	$doc.TablesOfContents.item(1).Update()
	$cp = $Null
	$ab = $Null
	$abstract = $Null
}

#bug fix 1-Apr-2014
#reset Grammar and Spelling options back to their original settings
$Word.Options.CheckGrammarAsYouType = $CurrentGrammarOption
$Word.Options.CheckSpellingAsYouType = $CurrentSpellingOption

Write-Verbose "$(Get-Date): Save and Close document and Shutdown Word"
If($WordVersion -eq $wdWord2007)
{
	#Word 2007
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
	}
	Else
	{
		Write-Verbose "$(Get-Date): Saving DOCX file"
	}
	Write-Verbose "$(Get-Date): Running Word 2007 and detected operating system $($RunningOS)"
	If($RunningOS.Contains("Server 2008 R2") -or $RunningOS.Contains("Server 2012"))
	{
		$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type] 
		$doc.SaveAs($filename1, $SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$SaveFormat = $wdSaveFormatPDF
			$doc.SaveAs($filename2, $SaveFormat)
		}
	}
	Else
	{
		#works for Server 2008 and Windows 7
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$doc.SaveAs([REF]$filename1, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$doc.SaveAs([REF]$filename2, [ref]$saveFormat)
		}
	}
}
Else
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
	$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
	$doc.SaveAs([REF]$filename1, [ref]$SaveFormat)
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Now saving as PDF"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
		$doc.SaveAs([REF]$filename2, [ref]$saveFormat)
	}
}

Write-Verbose "$(Get-Date): Closing Word"
$doc.Close()
$Word.Quit()
If($PDF)
{
	Write-Verbose "$(Get-Date): Deleting $($filename1) since only $($filename2) is needed"
	Remove-Item $filename1
}
Write-Verbose "$(Get-Date): System Cleanup"
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | out-null
Remove-Variable -Name word
$SaveFormat = $Null
[gc]::collect() 
[gc]::WaitForPendingFinalizers()
Write-Verbose "$(Get-Date): Script has completed"
Write-Verbose "$(Get-Date): "

If($PDF)
{
	Write-Verbose "$(Get-Date): $($filename2) is ready for use"
}
Else
{
	Write-Verbose "$(Get-Date): $($filename1) is ready for use"
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

#endregion Finalize Word Document
# SIG # Begin signature block
# MIIiywYJKoZIhvcNAQcCoIIivDCCIrgCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUISAbEj2xmV7MDu7j7bRltMLi
# taaggh41MIIDtzCCAp+gAwIBAgIQDOfg5RfYRv6P5WD8G/AwOTANBgkqhkiG9w0B
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
# 8jCCBmowggVSoAMCAQICEAOf7e3LeVuN7TIMiRnwNokwDQYJKoZIhvcNAQEFBQAw
# YjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQ
# d3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBD
# QS0xMB4XDTEzMDUyMTAwMDAwMFoXDTE0MDYwNDAwMDAwMFowRzELMAkGA1UEBhMC
# VVMxETAPBgNVBAoTCERpZ2lDZXJ0MSUwIwYDVQQDExxEaWdpQ2VydCBUaW1lc3Rh
# bXAgUmVzcG9uZGVyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAumlK
# gU1vpRQWqorNZ75Lv8Zpj1gc4HnoHp1YJpjaXNR8o/nbK4wSNsP8+WQGsbvCqJgK
# Fw3hletAtOuWbZi/po95z7yKknttnBgGUdilGFMyAScZYeiEQd/G8OjK/netX9ie
# e4xgb4VcRr1r5w+AzucDw3wxz7dlVcb74JkI5HNa+5fa0Ey+tLbGD38mkqm4/Dju
# tOQ6pEjQTOqpRidbz5IRk5wWp/7SrR8ixR6swXHvvErbAQlE35gcLWe6qIoDM8lR
# tfcCTQmkTf6AXsXXRcN9CKoBM8wz2E8wFuT/IjIu63478PkeMuuVJdLy/m1UhLrV
# 5dTR3RuvvVl7lIUwAQIDAQABo4IDNTCCAzEwDgYDVR0PAQH/BAQDAgeAMAwGA1Ud
# EwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwggG/BgNVHSAEggG2MIIB
# sjCCAaEGCWCGSAGG/WwHATCCAZIwKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRp
# Z2ljZXJ0LmNvbS9DUFMwggFkBggrBgEFBQcCAjCCAVYeggFSAEEAbgB5ACAAdQBz
# AGUAIABvAGYAIAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABlACAAYwBv
# AG4AcwB0AGkAdAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAgAG8AZgAg
# AHQAaABlACAARABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMAIABhAG4AZAAg
# AHQAaABlACAAUgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAgAEEAZwByAGUAZQBt
# AGUAbgB0ACAAdwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABpAGEAYgBpAGwAaQB0
# AHkAIABhAG4AZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwByAGEAdABlAGQAIABo
# AGUAcgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4wCwYJYIZIAYb9
# bAMVMB8GA1UdIwQYMBaAFBUAEisTmLKZB+0e36K+Vw0rZwLNMB0GA1UdDgQWBBRj
# L8nfeZJ7tSPKu+Gk7jN+4+Kd+jB9BgNVHR8EdjB0MDigNqA0hjJodHRwOi8vY3Js
# My5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDA4oDagNIYy
# aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5j
# cmwwdwYIKwYBBQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRENBLTEuY3J0MA0GCSqGSIb3DQEBBQUAA4IBAQCr
# dL1AAEx2FSVXPdMcA/99RchFEmbnKGVg2N87s/oNwawzj/SBuWHxnfuYVdfeR0O6
# gD3xSMw/ZzBWH8700EyEvYeknsXhD6gGXdAvbl7cGejwh+rgTq89bCCOc29+1ocY
# 4IbTmvye6oxy6UEPuHG1OCz4KbLVHKKdG+xfKrjcNyDhy7vw0GxspbPLn0r2VOMm
# ND0uuMErHLf2wz3+0S0eUPSUyPj97nPbSbUb9PX/pZDBORQb2O1xG2qY+/pAmkSp
# KQ5VXni4t6SDw3AB8GZA5a55NOErTQOhLebbVGIY7dUJi6Kq1gzITxq+mSV4aZmJ
# 1FmJ3t+I8NNnXnSlnaZEMIIGkDCCBXigAwIBAgIQBKVRftX3ANDrw0+OjYS9xjAN
# BgkqhkiG9w0BAQUFADBvMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQg
# SW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMS4wLAYDVQQDEyVEaWdpQ2Vy
# dCBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBDQS0xMB4XDTExMDkzMDAwMDAwMFoX
# DTE0MTAwODEyMDAwMFowXDELMAkGA1UEBhMCVVMxCzAJBgNVBAgTAlROMRIwEAYD
# VQQHEwlUdWxsYWhvbWExFTATBgNVBAoTDENhcmwgV2Vic3RlcjEVMBMGA1UEAxMM
# Q2FybCBXZWJzdGVyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAz2g4
# Kup2X6Mscbuq96HnetDDiITbncV1LtQ8Rxf8ZtN00+O/TliIZsWtufMq7GsLj1D8
# ikWfcgWGqMngWMsVYB4vdr1B8aQuHmKWld7W+j8FhKp3l+rNuFviTGa62sR6fEVW
# 1N6lDtJJHpfSIg/FUFfAqOKl0gFc45PU7iWCh08+oG5FJdhZ3WY0SosS1QujKEA4
# riSjeXPV6XSLsAHTE/fmHlGuu7NzJyMUzNNz2gPOFxYupHygbduhM5aAItD6GJ1h
# ajlovRt71tAMyeIPWNjj9B2luXxfRbgO9eufw91uFrXnougBPa7/eQ25YdW3NcGf
# tosYjvVI6Ptw/AaSiQIDAQABo4IDOTCCAzUwHwYDVR0jBBgwFoAUe2jOKarAF75J
# euHlP9an90WPNTIwHQYDVR0OBBYEFMHndyU+4pRT+JRECX9EG4y1laDkMA4GA1Ud
# DwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzBzBgNVHR8EbDBqMDOgMaAv
# hi1odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vYXNzdXJlZC1jcy0yMDExYS5jcmww
# M6AxoC+GLWh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9hc3N1cmVkLWNzLTIwMTFh
# LmNybDCCAcQGA1UdIASCAbswggG3MIIBswYJYIZIAYb9bAMBMIIBpDA6BggrBgEF
# BQcCARYuaHR0cDovL3d3dy5kaWdpY2VydC5jb20vc3NsLWNwcy1yZXBvc2l0b3J5
# Lmh0bTCCAWQGCCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAg
# AHQAaABpAHMAIABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0
# AHUAdABlAHMAIABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABE
# AGkAZwBpAEMAZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABS
# AGUAbAB5AGkAbgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3
# AGgAaQBjAGgAIABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBk
# ACAAYQByAGUAIABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBu
# ACAAYgB5ACAAcgBlAGYAZQByAGUAbgBjAGUALjCBggYIKwYBBQUHAQEEdjB0MCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wTAYIKwYBBQUHMAKG
# QGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENv
# ZGVTaWduaW5nQ0EtMS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQUFAAOC
# AQEAm1zhveo2Zy2lp8UNpR2E2CE8/NvEk0NDLszcBBuMda3N8Du23CikXCgrVvE0
# 3mMaeu/cIMDVU01ityLaqvDuovmTsvAKqaSJNztV9yTeWK9H4+h+35UEIU5TvYLs
# uzEW+rI5M2KcCXR6/LF9ZPmnBf9hHnK44hweHpmDWbo8HPqMatnIo7ideucuDn/D
# BM6s63eTMsFQCPYwte5vxuyVLqodOubLvIOMezZzByrpvJp9+gWAL151CE4qR6xQ
# jpgk5KqSkkkyvl72D+3PhNwZuxZDbZil5PIcrjmaBYoG8wfJzoNrtPFq3aG8dnQr
# xjXJjl+IN1iHYehBAUoBX98EozCCBqMwggWLoAMCAQICEA+oSQYV1wCgviF2/cXs
# bb0wDQYJKoZIhvcNAQEFBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGln
# aUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTExMDIxMTEyMDAwMFoXDTI2MDIx
# MDEyMDAwMFowbzELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZ
# MBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEuMCwGA1UEAxMlRGlnaUNlcnQgQXNz
# dXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EtMTCCASIwDQYJKoZIhvcNAQEBBQADggEP
# ADCCAQoCggEBAJx8+aCPCsqJS1OaPOwZIn8My/dIRNA/Im6aT/rO38bTJJH/qFKT
# 53L48UaGlMWrF/R4f8t6vpAmHHxTL+WD57tqBSjMoBcRSxgg87e98tzLuIZARR9P
# +TmY0zvrb2mkXAEusWbpprjcBt6ujWL+RCeCqQPD/uYmC5NJceU4bU7+gFxnd7XV
# b2ZklGu7iElo2NH0fiHB5sUeyeCWuAmV+UuerswxvWpaQqfEBUd9YCvZoV29+1aT
# 7xv8cvnfPjL93SosMkbaXmO80LjLTBA1/FBfrENEfP6ERFC0jCo9dAz0eotyS+BW
# tRO2Y+k/Tkkj5wYW8CWrAfgoQebH1GQ7XasCAwEAAaOCA0MwggM/MA4GA1UdDwEB
# /wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDAzCCAcMGA1UdIASCAbowggG2MIIB
# sgYIYIZIAYb9bAMwggGkMDoGCCsGAQUFBwIBFi5odHRwOi8vd3d3LmRpZ2ljZXJ0
# LmNvbS9zc2wtY3BzLXJlcG9zaXRvcnkuaHRtMIIBZAYIKwYBBQUHAgIwggFWHoIB
# UgBBAG4AeQAgAHUAcwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQByAHQAaQBmAGkA
# YwBhAHQAZQAgAGMAbwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBjAGUAcAB0AGEA
# bgBjAGUAIABvAGYAIAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAgAEMAUAAvAEMA
# UABTACAAYQBuAGQAIAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQAGEAcgB0AHkA
# IABBAGcAcgBlAGUAbQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBtAGkAdAAgAGwA
# aQBhAGIAaQBsAGkAdAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBjAG8AcgBwAG8A
# cgBhAHQAZQBkACAAaABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBlAHIAZQBuAGMA
# ZQAuMBIGA1UdEwEB/wQIMAYBAf8CAQAweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUF
# BzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6
# Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5j
# cnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmw0LmRp
# Z2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwHQYDVR0OBBYE
# FHtozimqwBe+SXrh5T/Wp/dFjzUyMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA0GCSqGSIb3DQEBBQUAA4IBAQB7ch1k/4jIOsG36eepxIe725SS15BZ
# M/orh96oW4AlPxOPm4MbfEPE5ozfOT7DFeyw2jshJXskwXJduEeRgRNG+pw/alE4
# 3rQly/Cr38UoAVR5EEYk0TgPJqFhkE26vSjmP/HEqpv22jVTT8nyPdNs3CPtqqBN
# ZwnzOoA9PPs2TJDndqTd8jq/VjUvokxl6ODU2tHHyJFqLSNPNzsZlBjU1ZwQPNWx
# HBn/j8hrm574rpyZlnjRzZxRFVtCJnJajQpKI5JA6IbeIsKTOtSbaKbfKX8GuTwO
# vZ/EhpyCR0JxMoYJmXIJeUudcWn1Qf9/OXdk8YSNvosesn1oo6WQsQz/MIIGzTCC
# BbWgAwIBAgIQBv35A5YDreoACus/J7u6GzANBgkqhkiG9w0BAQUFADBlMQswCQYD
# VQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGln
# aWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0Ew
# HhcNMDYxMTEwMDAwMDAwWhcNMjExMTEwMDAwMDAwWjBiMQswCQYDVQQGEwJVUzEV
# MBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29t
# MSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTEwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQDogi2Z+crCQpWlgHNAcNKeVlRcqcTSQQaPyTP8
# TUWRXIGf7Syc+BZZ3561JBXCmLm0d0ncicQK2q/LXmvtrbBxMevPOkAMRk2T7It6
# NggDqww0/hhJgv7HxzFIgHweog+SDlDJxofrNj/YMMP/pvf7os1vcyP+rFYFkPAy
# IRaJxnCI+QWXfaPHQ90C6Ds97bFBo+0/vtuVSMTuHrPyvAwrmdDGXRJCgeGDboJz
# PyZLFJCuWWYKxI2+0s4Grq2Eb0iEm09AufFM8q+Y+/bOQF1c9qjxL6/siSLyaxhl
# scFzrdfx2M8eCnRcQrhofrfVdwonVnwPYqQ/MhRglf0HBKIJAgMBAAGjggN6MIID
# djAOBgNVHQ8BAf8EBAMCAYYwOwYDVR0lBDQwMgYIKwYBBQUHAwEGCCsGAQUFBwMC
# BggrBgEFBQcDAwYIKwYBBQUHAwQGCCsGAQUFBwMIMIIB0gYDVR0gBIIByTCCAcUw
# ggG0BgpghkgBhv1sAAEEMIIBpDA6BggrBgEFBQcCARYuaHR0cDovL3d3dy5kaWdp
# Y2VydC5jb20vc3NsLWNwcy1yZXBvc2l0b3J5Lmh0bTCCAWQGCCsGAQUFBwICMIIB
# Vh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMAIABDAGUAcgB0AGkA
# ZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMAIABhAGMAYwBlAHAA
# dABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMAZQByAHQAIABDAFAA
# LwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkAbgBnACAAUABhAHIA
# dAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgAIABsAGkAbQBpAHQA
# IABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUAIABpAG4AYwBvAHIA
# cABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAAcgBlAGYAZQByAGUA
# bgBjAGUALjALBglghkgBhv1sAxUwEgYDVR0TAQH/BAgwBgEB/wIBADB5BggrBgEF
# BQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBD
# BggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2Ny
# bDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDig
# NoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNybDAdBgNVHQ4EFgQUFQASKxOYspkH7R7for5XDStnAs0wHwYDVR0jBBgw
# FoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZIhvcNAQEFBQADggEBAEZQPsm3
# KCSnOB22WymvUs9S6TFHq1Zce9UNC0Gz7+x1H3Q48rJcYaKclcNQ5IK5I9G6OoZy
# rTh4rHVdFxc0ckeFlFbR67s2hHfMJKXzBBlVqefj56tizfuLLZDCwNK1lL1eT7EF
# 0g49GqkUW6aGMWKoqDPkmzmnxPXOHXh2lCVz5Cqrz5x2S+1fwksW5EtwTACJHvzF
# ebxMElf+X+EevAJdqP77BzhPDcZdkbkPZ0XN1oPt55INjbFpjE/7WeAjD9KqrgB8
# 7pxCDs+R1ye3Fu4Pw718CqDuLAhVhSK46xgaTfwqIa1JMYNHlXdx3LEbS0scEJx3
# FMGdTy9alQgpECYxggQAMIID/AIBATCBgzBvMQswCQYDVQQGEwJVUzEVMBMGA1UE
# ChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMS4wLAYD
# VQQDEyVEaWdpQ2VydCBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBDQS0xAhAEpVF+
# 1fcA0OvDT46NhL3GMAkGBSsOAwIaBQCgQDAZBgkqhkiG9w0BCQMxDAYKKwYBBAGC
# NwIBBDAjBgkqhkiG9w0BCQQxFgQUksNkWkhNEsg6cfs621AqpuXXNHEwDQYJKoZI
# hvcNAQEBBQAEggEAtiU+QD+vmTpX9j5OHaqF7ZouVmBWhK/8bbRQZIZUSAnF+SHE
# e13lI8QrAfWrzCLNii1UVKdJ3QKLVajNZMsW6XeVlejmJmf/makBJTTm4f6Fvr+w
# x9VfKK2DqK5WkObJN5BC7uf7j3OctiAMKum/w0vxYdcOTVUeyo9+nkZNA17/w+XE
# 5veGbJJ2dVyAa7V2u+f17QGiiJCaIUjqg3OwLg/OmI8C2JIj1eP7WnGr65qmdkxv
# xRjobHaQhY8M4MefAu04hswSxvwmD4yFX4sAhG09NRM6UAM4j2sKxl04a5Yu8pFb
# QP5w7YG+uEdpczGI2po9IhzF/s5lO4iXDUAHxKGCAg8wggILBgkqhkiG9w0BCQYx
# ggH8MIIB+AIBATB2MGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0
# IEFzc3VyZWQgSUQgQ0EtMQIQA5/t7ct5W43tMgyJGfA2iTAJBgUrDgMCGgUAoF0w
# GAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTQwNTEy
# MTUxMTU4WjAjBgkqhkiG9w0BCQQxFgQUkZL1ARZmSX+TBnChk7ojHcbXFPYwDQYJ
# KoZIhvcNAQEBBQAEggEARHu7FTCKog0YlgIHjnlEEUO+MtboenNZwu6BRiDZqmKv
# +kbi9VqSXmKY5o67IoEEOb5s+pj4Lg0mX0jOfshHDkZEM2c/drzNeMoskmA7xX17
# vXM3FiMQgRmeMuE9VDsvvGSJrm1aawOf5aOfDjZV2cWq+01KF5LpCe2SCM2S3x4B
# tqP3xdxrYTfFyMx4qZlEMk09R/J5qVQxN4cTEUW3dbq63uoNaXki4Z5Cq7ODYr9E
# GaBWA34Z/y6kIv/178Hw2ffYS5s/Xn9+jwQapf3DftJGAD5fHwPcz9lzuERYWZuv
# IA8Uf1knMFl1BwT4EQeOcGCqRvIWrIoMJCiR7GWjFw==
# SIG # End signature block
