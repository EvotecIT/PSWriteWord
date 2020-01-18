Import-Module .\Documentimo.psd1 -Force

$TableDesigns = [Xceed.Document.NET.TableDesign].GetEnumNames()

$TableData = @{
    Name          = 'Test'
    Value         = 'Showing design'
    Company       = 'Evotec'
    Module        = 'Documentimo'
    'Best Module' = 'PSWriteWord'
}

Documentimo -FilePath "$PSScriptRoot\TableDesigns.docx" {
    foreach ($Design in $TableDesigns) {
        DocNumbering -Text "Table Design - $Design" -Level 1 -Type Numbered -Heading Heading1 {
            DocTable -DataTable $TableData -Design $Design
        }
    }
} -Open