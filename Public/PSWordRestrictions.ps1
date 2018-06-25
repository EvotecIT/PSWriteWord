function Add-WordProtection {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [EditRestrictions] $EditRestrictions,
        [string] $Password
    )
    if ($Password -eq $null) {
        $WordDocument.AddProtection($EditRestrictions)
    } else {
        $WordDocument.AddPasswordProtection($EditRestrictions, $Password)
    }
}

<#
    /// <summary>
    /// Returns true if any editing restrictions are imposed on this document.
    /// </summary>
    /// <example>
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     if(document.isProtected)
    ///         Console.WriteLine("Protected");
    ///     else
    ///         Console.WriteLine("Not protected");
    ///
    ///     // Save the document.
    ///     document.Save();
    /// }
    /// </code>
    /// </example>
    /// <seealso cref="AddProtection(EditRestrictions)"/>
    /// <seealso cref="RemoveProtection"/>
    /// <seealso cref="GetProtectionType"/>
#>
