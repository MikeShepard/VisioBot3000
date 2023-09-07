<#
        .SYNOPSIS 
        Creates a new Visio Layer and adds the given objects to it.

        .DESCRIPTION
        Long DescriptionCreates a new Visio Layer and adds the given objects to it.

        .PARAMETER LayerName
        The name for the new layer

        .PARAMETER Contents
        The objects to be included in the layer

        .PARAMETER Preserve
        Whether to preserve the existing layer assignments for these objecfts.

        .INPUTS
        None. You cannot pipe objects to New-VisioLayer.

        .OUTPUTS
        Visio.Layer

        .EXAMPLE
        New-VisioLayer -Layer WebServers -Contents WebServer1,WebServer2 -Preserve
#>
Function New-VisioLayer{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param([string]$LayerName,[Array]$Contents,[switch]$Preserve)
    if($PSCmdlet.ShouldProcess('Visio','Create a layer on the current page')){
        if($Preserve){
            $AddOption=1
        } else {
            $AddOption=0
        }
        $p=$Visio.ActivePage
        $layer=$p.Layers | Where-Object {$_.Name -eq $LayerName}
        if ($null -eq $layer){
            $layer=$p.Layers.Add($LayerName) 
        }
        if ($Contents -is [scriptblock]){
            $Contents = & $Contents
        }
        foreach($item in [array]$Contents){
            if($item -is [string]){
                $item=$p.Shapes[$item] 
            }
            $layer.Add($item,$AddOption)
        }
    }
}