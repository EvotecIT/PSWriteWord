function Convert-ObjectToProcess {
    [CmdletBinding()]
    param (
        [Array] $DataTable
    )
    $ObjectType = $DataTable.GetType().Name
    Write-Verbose "Convert-ObjectToProcess - GetType Before Conversion: $ObjectType"
    #$($DataTable.GetType().BaseType.Name)
    #$($DataTable.GetType().Name)
    if ($($DataTable.GetType().BaseType.Name) -eq 'Array' -and $($DataTable.GetType().Name) -eq 'Object[]') {
        Write-Verbose 'Convert-ObjectToProcess - Converting Array of Objects'
        #if ($DataTable.Count -gt 1) {
        $DataTable = $DataTable.ForEach( {[PSCustomObject]$_})
        #}

    }

    $ObjectType = $DataTable.GetType().Name
    Write-Verbose "Convert-ObjectToProcess - Table row count: $($DataTable.Count)"
    Write-Verbose "Convert-ObjectToProcess - Object Type: $ObjectType"
    Write-Verbose "Convert-ObjectToProcess - BaseType.Name: $($DataTable.GetType().BaseType.Name)"
    Write-Verbose "Convert-ObjectToProcess - GetType Before Final Conversion: $ObjectType"
    If ($ObjectType -eq 'Hashtable' -or $ObjectType -eq 'OrderedDictionary' -or $ObjectType -eq 'PSCustomObject') {
        Write-Verbose 'Convert-ObjectToProcess - Skipping select for Hashtable / OrderedDictionary / PSCustomObject'
    } else {
        #if ($ObjectType -eq 'PSCustomObject') {
        #    Write-Verbose 'Convert-ObjectToProcess - Skipping all objects'
        #$DataTable = [rray] ($DataTable | Select-Object *)
        #} else {


        if ($ObjectType -eq 'Collection`1' -and $($DataTable.Count) -eq 1) {
            Write-Verbose 'Convert-ObjectToProcess - Selecting all objects, returning array'
            $DataTable = [array] ($DataTable | Select-Object *)
        } else {
            Write-Verbose 'Convert-ObjectToProcess - Selecting all objects'
            $DataTable = ($DataTable | Select-Object *)
        }
        #}
    }

    $ObjectType = $DataTable.GetType().Name

    Write-Verbose "Convert-ObjectToProcess - GetType After Conversion: $ObjectType"
    return , $DataTable
}