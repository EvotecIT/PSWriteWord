Add-Type -TypeDefinition @"
public enum ContainerType {
    None,
    TOC,
    Section,
    Cell,
    Table,
    Header,
    Footer,
    Paragraph,
    Body
}
"@