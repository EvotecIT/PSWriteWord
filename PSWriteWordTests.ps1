
<#
get_MarginTop
set_MarginTop
get_MarginBottom
set_MarginBottom
get_MarginLeft
set_MarginLeft
get_MarginRight
set_MarginRight
get_MarginHeader
set_MarginHeader
get_MarginFooter
set_MarginFooter
get_MirrorMargins
set_MirrorMargins
get_PageWidth
set_PageWidth
get_PageHeight
set_PageHeight
get_isProtected
get_PageLayout
get_Headers
get_Footers
get_DifferentOddAndEvenPages
set_DifferentOddAndEvenPages
get_DifferentFirstPage
set_DifferentFirstPage
get_Images
get_CustomProperties
get_CoreProperties
get_Text
get_Paragraphs
get_Lists
get_Tables
get_FootnotesText
get_EndnotesText
get_Bookmarks
get_ParagraphsDeepSearch
get_Sections
get_Hyperlinks
get_Pictures
get_Xml
set_Xml
get_PackagePart
set_PackagePart
GetProtectionType
AddProtection
RemoveProtection
InsertDocument
InsertTable
AddTable
InsertTable
InsertTable
InsertTable
Create
Create
Load
Load
ApplyTemplate
ApplyTemplate
ApplyTemplate
ApplyTemplate
AddImage
AddImage
AddHyperlink
AddHeaders
AddFooters
Save
SaveAs
SaveAs
AddCoreProperty
AddCustomProperty
InsertParagraph
InsertParagraph
InsertParagraph
InsertParagraph
InsertParagraph
InsertParagraph
InsertParagraph
InsertParagraph
InsertParagraphs
InsertEquation
InsertChart
InsertChartAfterParagraph
GetSections
AddList
AddListItem
InsertList
InsertList
InsertList
InsertList
InsertDefaultTableOfContents
InsertTableOfContents
InsertTableOfContents
Copy
AddPasswordProtection
Dispose
SetDirection
FindAll
FindAll
FindUniqueByPattern
ReplaceText
ReplaceText
InsertAtBookmark
RemoveParagraphAt
RemoveParagraph
InsertBookmark
InsertTable
InsertTable
InsertTable
InsertTable
InsertSection
InsertSection
InsertSectionPageBreak
InsertList
RemoveTextInGivenFormat
ValidateBookmarks
ToString
Equals
GetHashCode
GetType
MarginTop
MarginBottom
MarginLeft
MarginRight
MarginHeader
MarginFooter
MirrorMargins
PageWidth
PageHeight
isProtected
PageLayout
Headers
Footers
DifferentOddAndEvenPages
DifferentFirstPage
Images
CustomProperties
CoreProperties
Text
Paragraphs
Lists
Tables
FootnotesText
EndnotesText
Bookmarks
ParagraphsDeepSearch
Sections
Hyperlinks
Pictures
Xml
PackagePart

New-WordDocument
Save-WordDocument
Close-WordDocument
Add-WordText
Add-WordBreak
Set-WordBuiltInProperty
Add-WordCoverPage
Set-WordOrientation
Add-WordTOC
Update-WordTOC
Add-WordTable
Get-WordBuiltinStyle
Get-WordWdTableFormat
Add-WordTemplate
Add-WordPicture


#>


Enum TableDesign {
    Custom
    TableNormal
    TableGrid
    LightShading
    LightShadingAccent1
    LightShadingAccent2
    LightShadingAccent3
    LightShadingAccent4
    LightShadingAccent5
    LightShadingAccent6
    LightList
    LightListAccent1
    LightListAccent2
    LightListAccent3
    LightListAccent4
    LightListAccent5
    LightListAccent6
    LightGrid
    LightGridAccent1
    LightGridAccent2
    LightGridAccent3
    LightGridAccent4
    LightGridAccent5
    LightGridAccent6
    MediumShading1
    MediumShading1Accent1
    MediumShading1Accent2
    MediumShading1Accent3
    MediumShading1Accent4
    MediumShading1Accent5
    MediumShading1Accent6
    MediumShading2
    MediumShading2Accent1
    MediumShading2Accent2
    MediumShading2Accent3
    MediumShading2Accent4
    MediumShading2Accent5
    MediumShading2Accent6
    MediumList1
    MediumList1Accent1
    MediumList1Accent2
    MediumList1Accent3
    MediumList1Accent4
    MediumList1Accent5
    MediumList1Accent6
    MediumList2
    MediumList2Accent1
    MediumList2Accent2
    MediumList2Accent3
    MediumList2Accent4
    MediumList2Accent5
    MediumList2Accent6
    MediumGrid1
    MediumGrid1Accent1
    MediumGrid1Accent2
    MediumGrid1Accent3
    MediumGrid1Accent4
    MediumGrid1Accent5
    MediumGrid1Accent6
    MediumGrid2
    MediumGrid2Accent1
    MediumGrid2Accent2
    MediumGrid2Accent3
    MediumGrid2Accent4
    MediumGrid2Accent5
    MediumGrid2Accent6
    MediumGrid3
    MediumGrid3Accent1
    MediumGrid3Accent2
    MediumGrid3Accent3
    MediumGrid3Accent4
    MediumGrid3Accent5
    MediumGrid3Accent6
    DarkList
    DarkListAccent1
    DarkListAccent2
    DarkListAccent3
    DarkListAccent4
    DarkListAccent5
    DarkListAccent6
    ColorfulShading
    ColorfulShadingAccent1
    ColorfulShadingAccent2
    ColorfulShadingAccent3
    ColorfulShadingAccent4
    ColorfulShadingAccent5
    ColorfulShadingAccent6
    ColorfulList
    ColorfulListAccent1
    ColorfulListAccent2
    ColorfulListAccent3
    ColorfulListAccent4
    ColorfulListAccent5
    ColorfulListAccent6
    ColorfulGrid
    ColorfulGridAccent1
    ColorfulGridAccent2
    ColorfulGridAccent3
    ColorfulGridAccent4
    ColorfulGridAccent5
    ColorfulGridAccent6
    None
}




Set-StrictMode -Version Latest
Clear-Host
$VerbosePreference = "SilentlyContinue"
$DebugPreference = "SilentlyContinue"
#$DebugPreference

Import-Module PSWriteWord -Force


# https://blogs.technet.microsoft.com/heyscriptingguy/2010/11/11/use-powershell-to-work-with-the-net-framework-classes/


#$docxElement = [Xceed.Words.NET.DocXElement] #| Get-Member
#$fonts = [Xceed.Words.NET.Font] #| Get-Member
#$docx = [Xceed.Words.NET.Docx]

#$docx::Create("test.xlsx")
#$docx::Save()

#$sc = $sc = "Xceed.Words.NET.DocX" -as [type]
#[reflection.assembly]::GetAssembly($sc)
#[reflection.assembly]::GetAssembly($sc) | Get-Member | fl *
#[Reflection.Assembly]::LoadFile($AssemblyName)


function Write-Word ($FilePath) {
    #[Xceed.Words.NET.DocX].GetMembers().Name
    $Word = [Xceed.Words.NET.DocX]
    $Test = $Word::Create($FilePath)
    $Test.InsertParagraph("This is a text").FontSize("20") | Out-Null
    $Test.InsertParagraph("Like me like i do").FontSize("21") | Out-Null
    $Test.InsertParagraph("Process").FontSize("22") | Out-Null
    $Test.Save()
    #$Word = New-Object -TypeName Xceed.Words.NET
    #document.InsertParagraph( "Inserting table" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;
}
function Read-Word ($FilePath) {
    $Word = [Xceed.Words.NET.DocX]
    $WordOutput = $Word::Load($FilePath)
    $WordOutput.Text.FontSize
    $WordOutput.Alignment
    $WordOutput.Tables
}


function Add-Paragraph($WordDocument) {

    <#
       public static void SimpleFormattedParagraphs()
    {
      Console.WriteLine( "\tSimpleFormattedParagraphs()" );

      // Create a new document.
      using( DocX document = DocX.Create( ParagraphSample.ParagraphSampleOutputDirectory + @"SimpleFormattedParagraphs.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Formatted paragraphs" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Insert a Paragraph into this document.
        var p = document.InsertParagraph();

        // Append some text and add formatting.
        p.Append( "This is a simple formatted red bold paragraph" )
        .Font( new Font( "Arial" ) )
        .FontSize( 25 )
        .Color( Color.Red )
        .Bold()
        .Append( " containing a blue italic text." ).Font( new Font( "Times New Roman" ) ).Color( Color.Blue ).Italic()
        .SpacingAfter( 40 );

        // Insert another Paragraph into this document.
        var p2 = document.InsertParagraph();

        // Append some text and add formatting.
        p2.Append( "This is a formatted paragraph using spacing," )
        .Font( new Font( "Courier New" ) )
        .FontSize( 10 )
        .Italic()
        .Spacing( 5 )
        .Append( "highlight" ).Highlight( Highlight.yellow ).UnderlineColor( Color.Blue ).CapsStyle( CapsStyle.caps )
        .Append( " and strike through." ).StrikeThrough( StrikeThrough.strike );

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: SimpleFormattedParagraphs.docx\n" );
      }
    }

    /// <summary>
    /// Create a document and add a paragraph with all its lines on a single page.
    /// </summary>
    public static void ForceParagraphOnSinglePage()
    {
      Console.WriteLine( "\tForceParagraphOnSinglePage()" );

      // Create a new document.
      using( DocX document = DocX.Create( ParagraphSample.ParagraphSampleOutputDirectory + @"ForceParagraphOnSinglePage.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Prevent paragraph split" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create a Paragraph that will appear on 1st page.
        var p = document.InsertParagraph( "This is a paragraph on first page.\nLine2\nLine3\nLine4\nLine5\nLine6\nLine7\nLine8\nLine9\nLine10\nLine11\nLine12\nLine13\nLine14\nLine15\nLine16\nLine17\nLine18\nLine19\nLine20\nLine21\nLine22\nLine23\nLine24\nLine25\n" );
        p.FontSize(15).SpacingAfter( 30 );

        // Create a Paragraph where all its lines will appear on a same page.
        var p2 = document.InsertParagraph( "This is a paragraph where all its lines are on the same page. The paragraph does not split on 2 pages.\nLine2\nLine3\nLine4\nLine5\nLine6\nLine7\nLine8\nLine9\nLine10" );
        p2.SpacingAfter( 30 );

        // Indicate that all the paragraph's lines will be on the same page
        p2.KeepLinesTogether();

        // Create a Paragraph that will appear on 2nd page.
        var p3 = document.InsertParagraph( "This is a paragraph on second page.\nLine2\nLine3\nLine4\nLine5\nLine6\nLine7\nLine8\nLine9\nLine10" );

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: ForceParagraphOnSinglePage.docx\n" );
      }
    }

    /// <summary>
    /// Create a document and add a paragraph with all its lines on the same page as the next paragraph.
    /// </summary>
    public static void ForceMultiParagraphsOnSinglePage()
    {
      Console.WriteLine( "\tForceMultiParagraphsOnSinglePage()" );

      // Create a new document.
      using( DocX document = DocX.Create( ParagraphSample.ParagraphSampleOutputDirectory + @"ForceMultiParagraphsOnSinglePage.docx" ) )
      {
        // Add a title.
        document.InsertParagraph( "Keeps Paragraphs on same page" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create a Paragraph that will appear on 1st page.
        var p = document.InsertParagraph( "This is a paragraph on first page.\nLine2\nLine3\nLine4\nLine5\nLine6\nLine7\nLine8\nLine9\nLine10\nLine11\nLine12\nLine13\nLine14\nLine15\nLine16\nLine17\nLine18\nLine19\nLine20\nLine21\nLine22\n" );
        p.FontSize( 15 ).SpacingAfter( 30 );

        // Create a Paragraph where all its lines will appear on a same page as the next paragraph.
        var p2 = document.InsertParagraph( "This is a paragraph where all its lines are on the same page as the next paragraph.\nLine2\nLine3\nLine4\nLine5\nLine6\nLine7\nLine8\nLine9\nLine10" );
        p2.SpacingAfter( 30 );

        // Indicate that this paragraph will be on the same page as the next paragraph.
        p2.KeepWithNextParagraph();

        // Create a Paragraph that will appear on 2nd page.
        var p3 = document.InsertParagraph( "This is a paragraph on second page.\nLine2\nLine3\nLine4\nLine5\nLine6\nLine7\nLine8\nLine9\nLine10" );

        // Indicate that all this paragraph's lines will be on the same page.
        p3.KeepLinesTogether();

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: ForceMultiParagraphsOnSinglePage.docx\n" );
      }
    }
    #>

}
function Add-Section {
    param (
        $WordDocument,
        [switch] $PageBreak
    )
    if ($PageBreak) {
        $WordDocument.InsertSectionPageBreak()
    } else {
        $WordDocument.InsertSection()
    }
}

function RunMeLoad($FilePath = "C:\Users\pklys\Desktop\File1.docx") {
    $Word = [Xceed.Words.NET.DocX]
    $WordOutput = $Word::Load($FilePath)
    foreach ($p in   $WordOutput.Paragraphs) {
        Write-COlor "Paragraph " -Color Yellow
        $p
    }



    #$WordOutput.Text.FontSize
    #$WordOutput.Alignment
    #$WordOutput.Tables
}




function RunMe($ADSnapshot) {
    Write-Color "Start" -Color Red
    # Install-Module -Name ISEScriptingGeek -Force -Verbose -AllowClobber
    #Import-module IseScriptingGeek
    #clear-host
    $WordDocumentPath = "file2.docx"
    $WordDocument = New-WordDocument -FilePath $WordDocumentPath


    #  $toc = $WordDocument.InsertTableOfContents("Table of content", 1)
    #$toc


    # $p = $WordDocument.InsertParagraph("This is my text")
    # $p = $WordDocument.InsertParagraph("This is another text").FontSize(15)



    ### DocX Example
    #Add-Section -WordDocument $WordDocument -PageBreak
    #$ListOfItems = @('Test1', 'Test2', 'Test3', 'Test4', 'Test5')
    #Add-List -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfItems
    #$p = $WordDocument.InsertParagraph("This is another text").FontSize(15)
    #Add-List -WordDocument $WordDocument -ListType Numbered -ListData $ListOfItems
    #$p = $WordDocument.InsertParagraph("This is another text").FontSize(15)


    ### AD Export via Bulleted
    #Add-Section -WordDocument $WordDocument -PageBreak
    $ListType = 'Bulleted' #'Numbered' #
    #$p = $WordDocument.InsertParagraph("Active Directory Root DSE").FontSize(15)
    #Add-List -WordDocument $WordDocument -ListType $ListType -Object $ADSnapshot.RootDSE
    #$p = $WordDocument.InsertParagraph("Active Directory Forest Information").FontSize(15)
    #Add-List -WordDocument $WordDocument -ListType $ListType -Object $ADSnapshot.ForestInformation
    #$p = $WordDocument.InsertParagraph("Active Directory Domain Information").FontSize(15)
    #Add-List -WordDocument $WordDocument -ListType $ListType -Object $ADSnapshot.DomainInformation

    #$p = $WordDocument.InsertParagraph("Active Directory Forest Information").FontSize(15)
    #$p1 = $WordDocument.InsertParagraph('1. This is text')
    #$p1.StyleName = "Heading1"
    #$p1.Alignment = "left"
    #$p1.ListItemType = 'Numbered'

    #$p2 = $WordDocument.InsertParagraph()
    #$p2.StyleName = "Heading1"
    #$p2.AddItem



    $numberList = $WordDocument.AddList("Test1", 0, 'Numbered' )
    $heading1 = $WordDocument.InsertList($numberList)

    #Add-List -WordDocument $WordDocument -ListType $ListType -Object $ADSnapshot.ForestInformation


    #$heading1 = $WordDocument.InsertList($numberList)
    #$WordDocument.AddListItem($numberList, 'Test2');
    #$heading1 = $WordDocument.InsertList($numberList)
    $p = $WordDocument.InsertParagraph("Active Directory Root DSE").FontSize(15)

    $numberList.AddItem($p)

    $p1 = $WordDocument.AddListItem($numberList, 'Test2');

    $numberList.AddItem($p1)
    #Add-WordTable -WordDocument $WordDocument -Table $ADSnapshot.RootDSE -Design "LightShading"

    #$heading1 = $WordDocument.InsertList($numberList)
    #  $WordDocument.AddListItem($numberList, 'Test3');

    # Add-List -WordDocument $WordDocument -ListType $ListType -Object $ADSnapshot.RootDSE

    # $heading1 = $WordDocument.InsertList($numberList)
    # $WordDocument.AddListItem($numberList, 'Test4');

    # Add-List -WordDocument $WordDocument -ListType $ListType -Object $ADSnapshot.DomainInformation

    #$heading1 = $WordDocument.InsertList($numberList)
    $Paragraphs = Get-ParagraphForList $WordDocument $heading1.NumID
    foreach ($p in $Paragraphs) {
        $p.StyleName = 'Heading1'
    }

    Get-Paragraphs -WordDocument $WordDocument

    #$numberList1 = $WordDocument.AddList("Test1", 0, 'Numbered' )
    #$heading2 = $WordDocument.InsertList($numberList1)
    #$heading2.NumID

    # $heading
    # $numberList
    #  $p3 = $WordDocument.InsertParagraph(0, "test", $true)
    #$p3

    #$WordDocument.Paragraphs


    # Paragraph InsertParagraph()
    # Paragraph InsertParagraph( int index, string text, bool trackChanges )
    # Paragraph InsertParagraph( Paragraph p )
    # Paragraph InsertParagraph( int index, Paragraph p )
    # Paragraph InsertParagraph( int index, string text, bool trackChanges, Formatting formatting )
    # Paragraph InsertParagraph( string text )
    # Paragraph InsertParagraph( string text, bool trackChanges )
    # Paragraph InsertParagraph( string text, bool trackChanges, Formatting formatting )

    #$t1 = $WordDocument.AddItem()
    #$t1
    #$t1.StyleName = 'Heading1'
    #$t1.InsertParagraph("Test")


    #.Heading("Heading2")
    #$p.Heading = 'Heading1'

    ### AD Export via Table

    #Add-Section -WordDocument $WordDocument -PageBreak
    ##$p = $WordDocument.InsertParagraph("Active Directory Root DSE").FontSize(15)
    #$p = $WordDocument.InsertParagraph("")
    #Add-WordTable -WordDocument $WordDocument -Table $ADSnapshot.RootDSE -Design "LightShading"
    #$p = $WordDocument.InsertParagraph("Active Directory Forest Information").FontSize(15)
    #$p = $WordDocument.InsertParagraph("")
    #Add-WordTable -WordDocument $WordDocument -Table $ADSnapshot.ForestInformation -Design "LightShading"
    #$p = $WordDocument.InsertParagraph("Active Directory Domain Information").FontSize(15)
    #$p = $WordDocument.InsertParagraph("")
    #Add-WordTable -WordDocument $WordDocument -Table $ADSnapshot.DomainInformation -Design "LightShading"

    <#
        $t = $WordDocument.InsertTable(10, 2)
        $t.Design = "LightShading"
        [Xceed.Words.NET.ListItemType] $ListTypeBulleted = [Xceed.Words.NET.ListItemType] 'Bulleted'
        [Xceed.Words.NET.ListItemType] $ListTypeNumbered = [Xceed.Words.NET.ListItemType] 'Numbered'

        $numberList = $WordDocument.AddList("Berries", 0, 'Numbered')

        $WordDocument.AddListItem($numberList, 'Straberises');
        $WordDocument.AddListItem($numberList, 'Straberises');
        $WordDocument.AddListItem($numberList, 'Straberises');
        $WordDocument.InsertList($numberList)
    #>

    ### DocX Eample
    #Add-Section -WordDocument $WordDocument -PageBreak
    #$Object1 = Get-Process #| Select-Object ProcessName, Site, StartTime
    #Add-WordTable -WordDocument $WordDocument -Table $Object1  #-Design "LightShading"

    #$Object2 = Get-PSDrive
    #Add-WordTable -WordDocument $WordDocument -Table $Object2 -Design "LightShading"


    ### DocX Example Save
    Save-WordDocument -WordDocument $WordDocument -FilePath "C:\Users\pklys\Desktop\File1.docx"

}

function RunMeAD() {
    #Import-Module ActiveDirectory

    #HashTable to save ADReport
    $ADSnapshot = @{}

    $ADSnapshot.RootDSE = $(
        $Info = Get-ADRootDSE

        <#
            configurationNamingContext
            currentTime
            defaultNamingContext
            dnsHostName
            domainControllerFunctionality
            domainFunctionality
            dsServiceName
            forestFunctionality
            highestCommittedUSN
            isGlobalCatalogReady
            isSynchronized
            ldapServiceName
            namingContexts
            rootDomainNamingContext
            schemaNamingContext
            serverName
            subschemaSubentry
            supportedCapabilities
            supportedControl
            supportedLDAPPolicies
            supportedLDAPVersion
            supportedSASLMechanisms
            #>
        #$Info

        $Info | Select-Object `
        @{label = 'Configuration Naming Context'; expression = { $_.configurationNamingContext }},
        defaultNamingContext,
        dnsHostName,
        domainControllerFunctionality,
        domainFunctionality,
        forestFunctionality,
        supportedLDAPPolicies,
        subschemaSubentry,
        supportedLDAPVersion,
        supportedSASLMechanisms

    )
    $ADSnapshot.ForestInformation = $(
        $Innfo = Get-ADForest
        <#
            ApplicationPartitions
            CrossForestReferences
            DomainNamingMaster
            Domains
            ForestMode
            GlobalCatalogs
            Name
            PartitionsContainer
            RootDomain
            SchemaMaster
            Sites
            SPNSuffixes
            UPNSuffixes
        #>

        $Info | Select-Object DomainNamingMaster, Domains, ForestMode, Sites
    )
    $ADSnapshot.DomainInformation = $(Get-ADDomain)
    $ADSnapshot.DomainControllers = $(Get-ADDomainController -Filter *)
    $ADSnapshot.DomainTrusts = (Get-ADTrust -Filter *)
    $ADSnapshot.DefaultPassWordPoLicy = $(Get-ADDefaultDomainPasswordPolicy)
    $ADSnapshot.AuthenticationPolicies = $(Get-ADAuthenticationPolicy -LDAPFilter '(name=AuthenticationPolicy*)')
    $ADSnapshot.AuthenticationPolicySilos = $(Get-ADAuthenticationPolicySilo -Filter 'Name -like "*AuthenticationPolicySilo*"')
    $ADSnapshot.CentralAccessPolicies = $(Get-ADCentralAccessPolicy -Filter *)
    $ADSnapshot.CentralAccessRules = $(Get-ADCentralAccessRule -Filter *)
    $ADSnapshot.ClaimTransformPolicies = $(Get-ADClaimTransformPolicy -Filter *)
    $ADSnapshot.ClaimTypes = $(Get-ADClaimType -Filter *)
    $ADSnapshot.DomainAdministrators = $( Get-ADGroup -Identity $('{0}-512' -f (Get-ADDomain).domainSID) | Get-ADGroupMember -Recursive)
    $ADSnapshot.OrganizationalUnits = $(Get-ADOrganizationalUnit -Filter *)
    $ADSnapshot.OptionalFeatures = $(Get-ADOptionalFeature -Filter *)
    $ADSnapshot.Sites = $(Get-ADReplicationSite -Filter *)
    $ADSnapshot.Subnets = $(Get-ADReplicationSubnet -Filter *)
    $ADSnapshot.SiteLinks = $(Get-ADReplicationSiteLink -Filter *)
    $ADSnapshot.LDAPDNS = $(Resolve-DnsName -Name "_ldap._tcp.$((Get-ADDomain).DNSRoot)" -Type srv)
    $ADSnapshot.KerberosDNS = $(Resolve-DnsName -Name "_kerberos._tcp.$((Get-ADDomain).DNSRoot)" -Type srv)
    return $ADSnapshot

}


$ADSnapshot = RunMeAD
RunMe -ADSnapshot $ADSnapshot
#$ADSnapshot.RootDSE

#RunMeLoad

#$ADSnapshot.RootDSE
#$ADSnapshot.ForestInformation #| Where { $_.Key -ne 'CrossForestReferences' }

#$value = RunMeAD
#$value.ForestInformation