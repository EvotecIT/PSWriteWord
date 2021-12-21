Import-Module .\PSWriteWord.psd1 -Force

$objChart = @(
    [PSCustomObject] @{ Type = 'Passed'; Count1 = 0 }
    [PSCustomObject] @{ Type = 'Failed'; Count1 = 3 }
    [PSCustomObject] @{ Type = 'Skipped'; Count1 = 0 }
)
<#

Documentimo -FilePath $PSScriptRoot\Documentimo-BasicList.docx {
    DocToc -Title 'Table of content'

    DocNumbering -Text 'My document' -Level 0 -Type Numbered -Heading Heading1 {
        DocText -Text 'Test', ' Test2' -Color Yellow
        DocTable -DataTable $objChart -Design ColorfulGrid
    }
    DocNumbering -Text 'AnotherChart' {
        DocChart -Title 'Processes' -DataTable $objChart -Key 'Type' -Value 'Count' -LegendPosition Right
    }
    DocNumbering -Text 'AnotherChart' {
        DocChart -Title 'Processes' -DataTable $objChart -Key 'Type' -Value 'Count' -LegendPosition Right {
            DocChartBar -Name 'Passed' -Value 0
            DocChartBar -Name 'Failed' -Value 3
            DocChartBar -Name 'Skipped' -Value 0
        }
        DocChart -Title 'Processes' -DataTable $objChart -Key 'Type' -Value 'Count' -LegendPosition Right  {
            foreach ($value in $objChart) {
                DocChartBar -Name $value.type -Value $value.count
            }
        }
    }

} -Open
#>

#return

$word = New-WordDocument "C:\lama.docx"

#$lamy = Get-GPOPolicy
$list = [System.Collections.ArrayList]@()

foreach ($lama in $objChart) {


    $chart = Add-WordChartSeries -Names $($lama.Name) -Values $($lama.Links.Count)
    $list.Add($chart)
}

Add-WordBarChart -WordDocument $word -ChartName "GPO " -ChartSeries $list.ForEach( { "$_," }) -ChartLegendPosition Bottom -ChartLegendOverlay $true -BarDirection Column
Save-WordDocument $word -Supress $true -Language 'en-US' -Verbose #-OpenDocument