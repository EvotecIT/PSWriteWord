Add-Type -TypeDefinition @"
public enum EditRestrictions {
    none,
    readOnly,
    forms,
    comments,
    trackedChanges
}
"@
