Import-Module .\PSWriteWord.psd1 -Force

Documentimo -FilePath $PSScriptRoot\Documentimo-BasicList.docx {
    DocList {
        DocListItem -Text 'Test 1' -Level 1
        DocListItem -Text 'Test 1' -Level 2
        DocListItem -Text 'Test 1' -Level 2
    }
    DocText -LineBreak
} -Open