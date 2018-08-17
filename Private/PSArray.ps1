function New-ArrayList {
    [CmdletBinding()]
    param()
    $List = New-Object System.Collections.ArrayList
    <#
    Mathias Rørbo Jessen:
        The pipeline will attempt to unravel the list on assignment,
        so you'll have to either wrap the empty arraylist in an array,
        like above, or call WriteObject explicitly and tell it not to, like so:
        $PSCmdlet.WriteObject($List,$false)
    #>
    return , $List
}
function Add-ToArray {
    [CmdletBinding()]
    param(
        [System.Collections.ArrayList] $List,
        [Object] $Element
    )
    #Write-Verbose "Add-ToArray - Element: $Element"
    $List.Add($Element) > $null
}
function Add-ToArrayAdvanced {
    [CmdletBinding()]
    param(
        [System.Collections.ArrayList] $List,
        [Object] $Element,
        [switch] $SkipNull,
        [switch] $RequireUnique,
        [switch] $FullComparison
    )
    if ($SkipNull -and $Element -eq $null) {
        #Write-Verbose "Add-ToArrayAdvanced - SkipNull used"
        return
    }
    if ($RequireUnique) {
        if ($FullComparison) {
            foreach ($ListElement in $List) {
                if ($ListElement -eq $Element) {
                    $TypeLeft = Get-ObjectType -Object $ListElement
                    $TypeRight = Get-ObjectType -Object $Element
                    if ($TypeLeft.ObjectTypeName -eq $TypeRight.ObjectTypeName) {
                        #Write-Verbose "Add-ToArrayAdvanced - RequireUnique with full comparison used"
                        return
                    }
                }
            }
        } else {
            if ($List -contains $Element) {
                #Write-Verbose "Add-ToArrayAdvanced - RequireUnique on name used"
                return
            }
        }
    }
    #Write-Verbose "Add-ToArrayAdvanced - Adding ELEMENT: $Element"
    $List.Add($Element) > $null
}

function Show-Array {
    [CmdletBinding()]
    param(
        [System.Collections.ArrayList] $List,
        [switch] $WithType
    )
    foreach ($Element in $List) {
        $Type = Get-ObjectType -Object $Element
        if ($WithType) {
            Write-Output "$Element (Type: $($Type.ObjectTypeName))"
        } else {
            Write-Output $Element
        }
    }
}


function Remove-FromArray {
    [CmdletBinding()]
    param(
        [System.Collections.ArrayList] $List,
        [Object] $Element,
        [switch] $LastElement
    )
    if ($LastElement) {
        $LastID = $List.Count - 1
        $List.RemoveAt($LastID) > $null
    } else {
        $List.Remove($Element) > $null
    }
}