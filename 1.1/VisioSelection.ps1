<#
        .SYNOPSIS 
        Creates a selection object using the given shapes.

        .DESCRIPTION
        Creates a selection object using the given shapes.  If -Visible is passed, the selection is shown in the application.

        .PARAMETER Objects
        The objects to be selected

        .PARAMETER Visible
        Whether the selection is visible in the application

        .INPUTS
        What can be piped in
        None. You cannot pipe objects to New-VisioSelection.

        .OUTPUTS
        Visio.Selection

        .EXAMPLE
        New-VisioSelection -Objects Server1,Server2
#>
Function New-VisioSelection{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param([array]$Objects,[switch]$Visible)
    if($PSCmdlet.ShouldProcess('Visio','Create a selection object')){
        $V=Get-VisioApplication
        $sel=$v.ActiveWindow.Selection
        if($visible){
            $sel=$v.ActiveWindow
        }
        $sel.DeselectAll()
        $CurrentPage=Get-VisioPage
        foreach($o in $objects){
            if($o -is [string]){
                $o=$CurrentPage.Shapes[$o]
            }
            $sel.Select($o,2)
        }
        $sel
    }
}
