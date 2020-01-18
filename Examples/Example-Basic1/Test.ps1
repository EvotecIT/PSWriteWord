Import-Module .\PSWriteWord.psd1 -Force


$Table = Get-Process | Select-Object -First 5

$TableForCharts = @(
    [PSCustomObject] @{ Name = 'Test 1'; SomeValue = 1 }
    [PSCustomObject] @{ Name = 'Test 2'; SomeValue = 5 }
    [PSCustomObject] @{ Name = 'Test 3'; SomeValue = 6 }
)

Documentimo -FilePath $PSScriptRoot\Test.docx {
    DocTOC -Title 'Table of content'

    DocNumbering -Text 'My document' -Level 0 -Type Numbered -Heading Heading1 {
        DocText -Text 'Test', ' Test2' -Color Yellow
        DocTable -DataTable $Table -Design ColorfulGrid
    }

    DocNumbering -Text 'Chart' {
        DocChart -Title 'Processes' -DataTable $Table  -Key 'ProcessName' -Value 'Handles'

    }
    DocNumbering -Text 'AnotherChart' {
        DocChart -Title 'Priviliged Group Members' -DataTable $TableForCharts  -Key 'Name' -Value 'SomeValue'
    }
    $Table1 = Get-Process | Select-Object -First 5

    DocTable -DataTable $Table -Design ColorfulGrid
    DocList {
        DocListItem -Text 'Test' -Level 0
        DocListItem -Text 'Test1' -Level 2
    }

    DocList -Type Numbered {
        DocListItem -Text 'Test' -Level 0
        DocListItem -Text 'Test1' -Level 2
    }

    DocTable -DataTable $Table1 -Design ColorfulGrid
} -Open