<#
        .SYNOPSIS 
        Draw a rectangle

        .DESCRIPTION
        Draws a rectangle on the active page with the position/size specified in the parameters

        .PARAMETER X0
        Describe The Left edge of the rectangle (in inches)

        .PARAMETER Y0
        Describe The Top edge of the rectangle (in inches)

        .PARAMETER X1
        Describe The Right edge of the rectangle (in inches)

        .PARAMETER Y1
        Describe The Bottom edge of the rectangle (in inches)

        .INPUTS
        None. You cannot pipe objects to New-VisioRectangle.

        .OUTPUTS
        Visio.Shape

        .EXAMPLE
        $rect = New-VisioRectangle 1 5 2 6
        #draws a rectangle

#>
Function New-VisioRectangle{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param($X0,$Y0,$X1,$Y1)
    if($PSCmdlet.ShouldProcess('Visio','Draw a rectangle on the page')){    
        $p=get-visioPage
        $p.DrawRectangle($X0,$Y0,$X1,$Y1)
    }
}