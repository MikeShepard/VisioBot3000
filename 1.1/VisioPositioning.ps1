<#
        .SYNOPSIS 
        Sets the next position to place a shape using relative positioning

        .DESCRIPTION
        Sets the next position to place a shape using relative positioning

        .INPUTS
        None.  

        .OUTPUTS
        None

        .EXAMPLE
        Set-NextShapePosition -x 5 -y 6
         

#>
Function Set-NextShapePosition{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param($X,$Y)
    if($PSCmdlet.ShouldProcess("Setting next shape position to ($x,$y)")){
        $script:LastDroppedObject=@{X=$X;y=$Y}
    }
}

<#
        .SYNOPSIS 
        Returns the next position to place a shape using relative positioning

        .DESCRIPTION
        Returns the next position to place a shape using relative positioning

        .INPUTS
        None. You cannot pipe objects to Get-NextShapePosition.

        .OUTPUTS
        HashTable

        .EXAMPLE
        Get-NextShapePosition  
        #returns a hashtable with X and Y position of next shape to place.

#>
Function Get-NextShapePosition{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    Param()
    if($LastDroppedObject -eq 0){
        #nothing dropped yet, start at top-left-ish
        $p=Get-VisioPage
        
        return @{X=1;Y=$p.Pagesheet.Cells('PageHeight').ResultIU-1}
    } elseif ($LastDroppedObject -is [hashtable]) {
        return $LastDroppedObject
    } else {
        if($RelativeOrientation -eq 'Horizontal'){
            $X=$LastDroppedObject.Cells('PinX').ResultIU + $LastDroppedObject.Cells('Width').ResultIU + $script:DeltaX
            $Y=$LastDroppedObject.Cells('PinY').ResultIU 
        } else {
            $X=$LastDroppedObject.Cells('PinX').ResultIU 
            $Y=$LastDroppedObject.Cells('PinY').ResultIU - $LastDroppedObject.Cells('Height').ResultIU - $script:DeltaY
        }
        Return @{X=$X;Y=$Y}
    }
}

<#
        .SYNOPSIS 
        Changes the direction VisioBot3000 uses when placing shapes using relative positioning

        .DESCRIPTION
        Changes the direction VisioBot3000 uses when placing shapes using relative positioning

        .PARAMETER Orientation
        Either vertical or Horizontal

        .INPUTS
        None. You cannot pipe objects to Set-RelativePositionDirection.

        .OUTPUTS
        None

        .EXAMPLE
        Set-RelativePositionDirection Horizontal
        File.txt

#>
Function Set-RelativePositionDirection{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param([ValidateSet('Horizontal','Vertical')]$Orientation)
    if($PSCmdlet.ShouldProcess('Visio','Sets VisioBot''s orientation for relative positioning')){
        $script:RelativeOrientation=$Orientation
    }
}

Function Set-NextShapePositionDelta{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param($DeltaX=$script:DeltaX,$Deltay=$script:Deltay)
    if($PSCmdlet.ShouldProcess("Setting next shape position delta to ($Deltax,$Deltay)")){
        $script:DeltaX=$DeltaX
        $script:DeltaY=$DeltaY
    }
}