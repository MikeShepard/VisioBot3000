<#
        .SYNOPSIS 
        Drops a shape on the page

        .DESCRIPTION
        Drops a shape (provided as a master shape) on the page.  If no X coordinate is given, the shape is positioned relative to the previous shape placed
        The shape is given a name and label.

        .PARAMETER Master
        Either the name of the master (previously registered using Register-VisioShape) or a reference to a master object.

        .PARAMETER X
        The X position used to place the shape (in inches). If this is omitted, the shape is positioned relative to the previous shape placed.

        .PARAMETER Y
        The Y position used to place the shape (in inches). 

        .PARAMETER Name
        The name for the new shape.

        .INPUTS
        None. You cannot pipe objects to Add-Extension.

        .OUTPUTS
        Visio.Shape

        .EXAMPLE
        New-VisioShape MasterShapeName -Label 'My Shape' -x 5 -y 5 -Name MyShape


#>
Function New-VisioShape{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param($Master,$Label,$X,$Y,$Name)
    if($PSCmdlet.ShouldProcess('Visio','Drop shape on the page')){
        if($Master -is [string]){
            $Master=$script:Shapes[$Master]
        }
        if(!$Name){
            $Name=$Label
        }
 
        $p=get-VisioPage
        if($updateMode){
            $DroppedShape=$p.Shapes | Where-Object {$_.Name -eq $Label}
        }
        if(-not (get-variable DroppedShape -Scope Local -ErrorAction Ignore) -or ($null -eq $DroppedShape)){
            if(-not $X){
                $RelativePosition=Get-NextShapePosition
                $X=$RelativePosition.X
                $Y=$RelativePosition.Y
            }
            $DroppedShape=$p.Drop($Master.PSObject.BaseObject,$X,$Y)
            $DroppedShape.Name=$Name
        } else {
            write-verbose "Existing shape <$Label> found"
        }
        $DroppedShape.Text=$Label
        New-Variable -Name $Name -Value $DroppedShape -Scope Global -Force
        write-output $DroppedShape
        $Script:LastDroppedObject=$DroppedShape
    }

}

<#
        .SYNOPSIS 
        Copies a master from a stencil and gives it a name.

        .DESCRIPTION
        Copies a master from a stencil and gives it a name.  Also creates a function with the same name to drop the shape onto the active Visio page.

        .PARAMETER Name
        The name used to refer to the shape

        .PARAMETER StencilName
        Which stencil to get the master from

        .PARAMETER MasterName
        The name of the master in the stencil

        .INPUTS
        None. You cannot pipe objects to Register-VisioShape.

        .OUTPUTS
        None

        .EXAMPLE
        Register-VisioShape -Name Block -StencilName BasicShapes -MasterName Block

#>
Function Register-VisioShape{
    [CmdletBinding()]
    Param([string]$Name,
        [Alias('From')][string]$StencilName,
    [string]$MasterName)
 
    if(!$MasterName){
        $MasterName=$Name
    }
    $newShape=$stencils[$StencilName].Masters | Where-Object {$_.Name -eq $MasterName}
    $script:Shapes[$Name]=$newshape
    $outerName=$Name 
    new-item -Path Function:\ -Name "global`:$outername" -value {param($Label, $X,$Y, $Name) $Shape=get-visioshape $outername; New-VisioShape $Shape $Label $X $Y -name $Name}.GetNewClosure() -force  | out-null
    $script:GlobalFunctions.Add($outerName) | Out-Null
}

<#
        .SYNOPSIS 
        Retrieves a saved shape definition

        .DESCRIPTION
        Retrieves a saved shape definition

        .PARAMETER Name
        Describe Parameter1

        .INPUTS
        None. You cannot pipe objects to Get-VisioShape

        .OUTPUTS
        Visio.Shape

        .EXAMPLE
        Get-VisioShape Block

#>
Function Get-VisioShape{
    [CmdletBinding()]
    Param([string]$Name)
    $script:Shapes[$Name]
}
