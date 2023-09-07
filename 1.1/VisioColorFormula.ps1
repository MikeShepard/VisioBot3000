<#
        .SYNOPSIS 
        Outputs a Visio color formula based on the color parameter
        .DESCRIPTION
        Outputs a Visio color formula based on the color parameter
        .PARAMETER Color
        A color you want to use in a Visio diagram
        .INPUTS
        None. You cannot pipe objects to Get-VisioColorFormula.
        .OUTPUTS
        None
        .EXAMPLE
        $formula=Get-VisioColorFormula Red;
 #> 
 function Get-VisioColorFormula{
    [CmdletBinding()]
    [OutputType([System.String])]
    Param([System.Drawing.Color]$color)
    
    return "=rgb($($Color.R),$($Color.G),$($Color.B))"
}