Add-Type -TypeDefinition @"
    public enum BarDirection {
        Column,
        Bar
    }
"@

Add-Type -TypeDefinition @"
    public enum BarGrouping {
        Clustered,
        PercentStacked,
        Stacked,
        Standard
    }
"@
