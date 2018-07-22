#
# Module manifest for module 'PSWriteWord'
#
# Generated by: Przemyslaw Klys
#
# Generated on: 03.06.2018
#

@{

    # Script module or binary module file associated with this manifest.
    RootModule        = 'PSWriteWord.psm1'

    # Version number of this module.
    ModuleVersion     = '0.4.7.1'

    # Supported PSEditions
    # CompatiblePSEditions = @()

    # ID used to uniquely identify this module
    GUID              = '6314c78a-d011-4489-b462-91b05ec6a5c4'

    # Author of this module
    Author            = 'Przemyslaw Klys'

    # Company or vendor of this module
    CompanyName       = 'Evotec'

    # Copyright statement for this module
    Copyright         = 'Evotec (c) 2018. All rights reserved.'

    # Description of the functionality provided by this module
    Description       = 'Simple project to create Microsoft Word in PowerShell without having Office installed.'

    # Minimum version of the Windows PowerShell engine required by this module
    # PowerShellVersion = ''

    # Name of the Windows PowerShell host required by this module
    # PowerShellHostName = ''

    # Minimum version of the Windows PowerShell host required by this module
    # PowerShellHostVersion = ''

    # Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
    # DotNetFrameworkVersion = ''

    # Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
    # CLRVersion = ''

    # Processor architecture (None, X86, Amd64) required by this module
    # ProcessorArchitecture = ''

    # Modules that must be imported into the global environment prior to importing this module
    # RequiredModules = @()

    # Assemblies that must be loaded prior to importing this module
    # RequiredAssemblies = @()

    # Script files (.ps1) that are run in the caller's environment prior to importing this module.
    #ScriptsToProcess  = @('Enums\*.ps1')
    ScriptsToProcess  = @( 'Enums\InsertWhere.ps1', 'Enums\Alignment.ps1' , 'Enums\AutoFit.ps1' , 'Enums\BarChart.ps1' , 'Enums\BasicShapes.ps1' , 'Enums\BlockArrowShapes.ps1' , 'Enums\BorderSize.ps1' , 'Enums\BorderStyle.ps1' , 'Enums\CalloutShapes.ps1' , 'Enums\CapsStyle.ps1' , 'Enums\Chart.ps1' , 'Enums\ContainerType.ps1' , 'Enums\CustomPropertyType.ps1' , 'Enums\Direction.ps1' , 'Enums\DocumentTypes.ps1' , 'Enums\EditRestrictions.ps1' , 'Enums\EditType.ps1' , 'Enums\EquationShapes.ps1' , 'Enums\FlowChartShapes.ps1' , 'Enums\HeadingType.ps1' , 'Enums\Highlight.ps1' , 'Enums\HorizontalBorderPosition.ps1' , 'Enums\LineChart.ps1' , 'Enums\LineSpacingType.ps1' , 'Enums\LineSpacingTypeAuto.ps1' , 'Enums\LineType.ps1' , 'Enums\ListItemType.ps1' , 'Enums\MatchFormattingOptions.ps1' , 'Enums\Misc.ps1' , 'Enums\Orientation.ps1' , 'Enums\PageNumberFormat.ps1' , 'Enums\RectangleShapes.ps1' , 'Enums\RunTextType.ps1' , 'Enums\Script.ps1' , 'Enums\SectionBreakType.ps1' , 'Enums\ShadingType.ps1' , 'Enums\StarAndBannerShapes.ps1' , 'Enums\StrikeThrough.ps1' , 'Enums\TabelCellBorderType.ps1' , 'Enums\TableBorderTypes.ps1' , 'Enums\TableCellMarginType.ps1' , 'Enums\TableDesign.ps1' , 'Enums\TableOfContentsSwitches.ps1' , 'Enums\TabStopPositionLeader.ps1' , 'Enums\TextDirection.ps1' , 'Enums\UnderlineStyle.ps1' , 'Enums\VerticalAlignment.ps1' )
    #ScriptsToProcess  = @('Private\PSWriteWordEnum.ps1')

    # Type files (.ps1xml) to be loaded when importing this module
    # TypesToProcess = @()

    # Format files (.ps1xml) to be loaded when importing this module
    # FormatsToProcess = @()

    # Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
    # NestedModules = @()

    # Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
    FunctionsToExport = '*'

    # Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
    CmdletsToExport   = @()

    # Variables to export from this module
    VariablesToExport = @()

    # Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
    AliasesToExport   = @()

    # DSC resources to export from this module
    # DscResourcesToExport = @()

    # List of all modules packaged with this module
    # ModuleList = @()

    # List of all files packaged with this module
    FileList          = 'PSWriteWord.psm1', 'PSWriteWord.psd1'

    # Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
    PrivateData       = @{

        PSData = @{

            # Tags applied to this module. These help with module discovery in online galleries.
            Tags       = @('word', 'docx', 'write', 'PSWord', 'office', 'pswriteword', 'writeword')

            # A URL to the license for this module.
            # LicenseUri = ''

            # A URL to the main website for this project.
            ProjectUri = 'https://github.com/EvotecIT/PSWriteWord'

            # A URL to an icon representing this module.
            # IconUri = ''

            # ReleaseNotes of this module
            # ReleaseNotes = ''

        } # End of PSData hashtable

    } # End of PrivateData hashtable

    # HelpInfo URI of this module
    HelpInfoURI       = ''

    # Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
    # DefaultCommandPrefix = ''

}
