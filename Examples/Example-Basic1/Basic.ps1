Import-Module .\PSWriteWord.psd1 -Force

Documentimo -FilePath $PSScriptRoot\Documentimo-BasicList.docx {
    DocTOC -Title 'Table of content'

    $OUs = Get-ADOrganizationalUnit -Filter * | Select-Object -First 2
    foreach ($OU in $OUs) {
        DocNumbering -Text $OU -Level 1 -Type Numbered -Heading Heading1 {

            $UserInfo = Get-ADUser -Filter * -SearchBase $OU
            foreach ($User in $UserInfo) {
                DocTable -DataTable $User -Design ColorfulGridAccent5 -AutoFit Window
                DocText -LineBreak
            }
        }
    }
    DocText -LineBreak

} -Open