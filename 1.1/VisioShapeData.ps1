<#
        .SYNOPSIS 
        Returns a shape data field from a shape

        .DESCRIPTION
        Returns a shape data field from a shape

        .PARAMETER Shape
        The shape that has the shape data

        .PARAMETER Name
        Which shape data field you want the value from

        .PARAMETER All
        Returns all shape data rather than just one

        .INPUTS
        None. You cannot pipe objects to Get-VisioShapeData.

        .OUTPUTS
        String

        .EXAMPLE
        Get-VisioShapeData -shape $webServer -Name IPAddress
#>
Function Get-VisioShapeData{
    [CmdletBinding()]
    Param($Shape,
    [string]$Name,
    [switch]$All)
    
    if($all){
        0..($shape.Section(243).Count-1) | foreach-object {
            [PSCustomObject]@{Label=$shape.CellsSRC(243,$_,2).Formula.Replace('"','')
                              Name=$shape.CellsSRC(243,$_,2).RowName
                              Value=$shape.CellsSRC(243,$_,0).Formula

                              ValueObject=$shape.CellsSRC(243,$_,0)}
        }
    } else {
        $Shape.Cells("Prop.$Name").Formula.TrimStart('"').TrimEnd('"') 
    }
}

<#
        .SYNOPSIS 
        Sets the value of a shape data field.

        .DESCRIPTION
        Sets the value of a shape data field.

        .PARAMETER Shape
        The shape that has the shape data

        .PARAMETER Name
        The name of the shape data field to set

        .PARAMETER Value
        The value to set the shape data to

        .INPUTS
        None. You cannot pipe objects to Set-VisioShapeData.

        .OUTPUTS
        None

        .EXAMPLE
        Set-VisioShapeData -shape $WebServer -Name IPAddress -Value 10.1.1.5
#>
Function Set-VisioShapeData{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param($Shape,
        $Name,
    $Value)
        if($PSCmdlet.ShouldProcess('Visio','Set a value for a custom shape data element')){
        $Shape.Cells("Prop.$Name").Formula="=`"$value`""
    }
}