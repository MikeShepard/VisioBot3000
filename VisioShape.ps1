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

<#
        .NOTES
        ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

        Function:      New-VisioRackShape
        Created by:    Martin Cooper
        Date:          01/10/2021
        GitHub:        https://github.com/mc1903
        Version:       1.2

        ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

        .SYNOPSIS
        Drops and connects a new rack equipment shape (server, etc) into an existing rack shape.

        .DESCRIPTION
        Drops and connects a new rack equipment shape (server, etc) into an existing rack shape at a specified first (lowest) U space.

        .PARAMETER Master
        Either the name of the new rack equipment shape master (previously registered using Register-VisioShape) or a reference to a master object.

        .PARAMETER Label
        The label for the new rack equipment shape.

        .PARAMETER Name
        The name for the new rack equipment shape. If not provided then the Label will be used as the Name

        .PARAMETER RackLabel
        The name for the existing rack shape.

        .PARAMETER RackVendor
        The vendor name for the existing rack shape.

        Supported Rack Vendors are HPE, Dell, IBM & Cisco.

        Within each vendor only the shapes listed below are working.

            HPE shapes from the HPE-Racks set on VisioCafe - https://www.visiocafe.com/downloads/hp/HPE-Common.zip

                    HPE 22U G2 Adv Rack Front
                    HPE 22U G2 Adv Rack Rear
                    HPE 36U G2 Adv Rack Front
                    HPE 36U G2 Adv Rack Rear
                    HPE 42U 800mm G2 Adv Rack Front
                    HPE 42U 800mm G2 Adv Rack Rear
                    HPE 42U 800mm G2 Ent Rack Front
                    HPE 42U 800mm G2 Ent Rack Rear
                    HPE 42U G2 Adv Rack Front
                    HPE 42U G2 Adv Rack Rear
                    HPE 42U G2 Ent Rack Front
                    HPE 42U G2 Ent Rack Rear
                    HPE 48U 800mm G2 Adv Rack Front
                    HPE 48U 800mm G2 Adv Rack Rear
                    HPE 48U 800mm G2 Ent Rack Front
                    HPE 48U 800mm G2 Ent Rack Rear
                    HPE 48U G2 Adv Rack Front
                    HPE 48U G2 Adv Rack Rear
                    HPE 48U G2 Ent Rack Front
                    HPE 48U G2 Ent Rack Rear
                    50U Ent. Rack

            Dell shapes from the Dell-Racks set on VisioCafe - https://www.visiocafe.com/downloads/dell/Dell-Racks.zip

                    2420 Rack Frame
                    4220 Rack Frame
                    4220D Rack Frame
                    4220W Rack Frame
                    4820 Rack Frame
                    4820D Rack Frame
                    4820W Rack Frame

            IBM shapes from the IBM-Racks set on VisioCafe - https://www.visiocafe.com/downloads/ibm/IBM-Common.zip

                    7014-S11 Rack
                    7014-S25 Rack
                    7014-S00 Rack
                    7014-T00 Rack
                    7014-T42 Rack

            Cisco shapes from the Cisco R-Series set on Cisco.com - https://www.cisco.com/c/dam/assets/prod/visio/visio/racks-cisco-r-series.zip

                    RACK2-UCS Front
                    RACK2-UCS Rear
                    RACK2-UCS2 Front
                    RACK2-UCS2 Rear

        .PARAMETER FirstU
        The first (lowest) U space in which to place the new rack equipment shape.

        .PARAMETER X
        The X position used for initial new rack equipment shape placement (in inches) prior to connecting it to the existing rack shape. If not provided the default X position = 1

        .PARAMETER Y
        The Y position used for initial new rack equipment shape placement (in inches) prior to connecting it to the existing rack shape. If not provided the default Y position = 0.25

        .OUTPUTS
        Visio.Shape

        .EXAMPLE
        New-VisioRackShape -Master HPEDL180Gen108LFF -Label Server1 -RackLabel Rack01 -RackVendor HPE -FirstU 1
#>
Function New-VisioRackShape{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param (
        [Parameter(Mandatory=$true)]$Master,
        [Parameter(Mandatory=$true)]$Label,
        [Parameter(Mandatory=$false)]$Name,
        [Parameter(Mandatory=$true)]$RackLabel,
        [Parameter(Mandatory=$true)]$RackVendor,
        [Parameter(Mandatory=$true)]$FirstU,
        [Parameter(Mandatory=$false)]$X,
        [Parameter(Mandatory=$false)]$Y
    )
    If($PSCmdlet.ShouldProcess('Visio','Drop a new rack equipment shape and connect to an existing rack shape')){
        If($Master -is [string]){
            $Master=$script:Shapes[$Master]
        }
        If(!$Name){
            $Name=$Label
        }

        If(!$X){
            $X=1.00
        }

        If(!$Y){
            $Y=0.25
        }

        $p=Get-VisioPage
        $ExistingRackShape=$p.Shapes | Where-Object {$_.Name -eq $RackLabel}
        If(!$ExistingRackShape){
            Write-Verbose "Existing rack shape $RackLabel was NOT found on the active page $($p.Name). Skipping!"
            Break
        }
        Else {
            Write-Verbose "Existing rack shape $RackLabel was found on the active page $($p.Name)."
            If ($RackVendor -match 'HPE') {
                $FirstUBXY = $FirstU | ForEach-Object { $_.ToString("00") }
                $FirstUEXY = $FirstU | ForEach-Object { $_.ToString("00") }
                $BeginXY = "=PAR(PNT($RackLabel!Connections.U$($FirstUBXY)B.X,$RackLabel!Connections.U$($FirstUBXY)B.Y))"
                $EndXY = "=PAR(PNT($RackLabel!Connections.U$($FirstUEXY)E.X,$RackLabel!Connections.U$($FirstUEXY)E.Y))"
            }
            ElseIf ($RackVendor -match 'Dell') {
                $FirstUBXY = $FirstU*2+3
                $FirstUEXY = $FirstU*2+4
                $BeginXY = "=PAR(PNT($RackLabel!Connections.X$($FirstUBXY),$RackLabel!Connections.Y$($FirstUBXY)))"
                $EndXY = "=PAR(PNT($RackLabel!Connections.X$($FirstUEXY),$RackLabel!Connections.Y$($FirstUEXY)))"
            }
            ElseIf ($RackVendor -match 'IBM') {
                $FirstUBXY = $FirstU*2+3
                $FirstUEXY = $FirstU*2+4
                $BeginXY = "=PAR(PNT($RackLabel!Connections.X$($FirstUBXY),$RackLabel!Connections.Y$($FirstUBXY)))"
                $EndXY = "=PAR(PNT($RackLabel!Connections.X$($FirstUEXY),$RackLabel!Connections.Y$($FirstUEXY)))"
            }
            ElseIf ($RackVendor -match 'Cisco') {
                $FirstUBXY = $FirstU
                $FirstUEXY = $FirstU
                $BeginXY = "=PAR(PNT($RackLabel!Connections.Cab$($FirstUBXY)A.X,$RackLabel!Connections.Cab$($FirstUBXY)A.Y))"
                $EndXY = "=PAR(PNT($RackLabel!Connections.Cab$($FirstUBXY)B.X,$RackLabel!Connections.Cab$($FirstUBXY)B.Y))"
            }
            Else {
                Write-Verbose "Rack Vendor $RackVendor is NOT known/supported. Skipping!"
                Break
            }

        }

        If($updateMode){
            $DroppedShape=$p.Shapes | Where-Object {$_.Name -eq $Label}
        }
        Else {
            Write-Verbose "Rack Vendor $RackVendor"
            Write-Verbose "BeginXY: $BeginXY"
            Write-Verbose "  EndXY: $EndXY"
    
            $DroppedShape=$p.Drop($Master.PSObject.BaseObject,$X,$Y)
            $DroppedShape.Name=$Name
            $DroppedShape.Text=$Label
            $DroppedShape.Cells("BeginX").FormulaU = $BeginXY
            $DroppedShape.Cells("BeginY").FormulaU = $BeginXY
            $DroppedShape.Cells("EndX").FormulaU = $EndXY
            $DroppedShape.Cells("EndY").FormulaU = $EndXY
    
            New-Variable -Name $Name -Value $DroppedShape -Scope Global -Force
            
            If ($PSBoundParameters['Verbose']) {
                Write-Output $DroppedShape
            }
            
            $Script:LastDroppedObject=$DroppedShape
        }

    }

}
