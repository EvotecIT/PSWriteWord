function Get-ColorFromARGB {
    [CmdletBinding()]
    param(
        [int] $A,
        [int] $R,
        [int] $G,
        [int] $B
    )
    return [system.drawing.color]::FromArgb($A, $R, $G, $B)
}