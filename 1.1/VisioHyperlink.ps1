<#
        .SYNOPSIS 
        Sets the hyperlink on a shape to the given address.

        .DESCRIPTION
        Sets the hyperlink on a shape to the given address.

        .PARAMETER Shape
        The shape you want the hyperlink on

        .PARAMETER Link
        The address of the hyperlink

        .INPUTS
        None. You cannot pipe objects to New-VisioHyperlink.

        .OUTPUTS
        None

        .EXAMPLE
        New-VisioHyperlink -shape $rectangle -link http://google.com
        File.txt

#>
Function New-VisioHyperlink{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param($Shape,
    $Link)
    if($PSCmdlet.ShouldProcess('Visio','Create a hyperlink on a shape')){
        $CurrentPage=Get-VisioPage
        if($Shape -is [string]){
            $Shape=$CurrentPage.Shapes[$Shape]
        }
        $LinkObject=$Shape.AddHyperLink()
        $LinkObject.Address=$Link
    }
}