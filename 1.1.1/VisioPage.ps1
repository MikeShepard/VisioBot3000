Function New-VisioPage{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param([string]$Name,
    $Visio=$script:Visio)

    if($PSCmdlet.ShouldProcess('Creating a new Visio Page')){
        $page=$Visio.ActiveDocument.Pages.Add( )
        if($Name){
            $page.NameU=$Name 
        }
        $page
    }
}

<#
        .SYNOPSIS 
        Change the active page in Visio

        .DESCRIPTION
        Changes the active page in Visio to the page named in the parameter

        .PARAMETER Name
        Page name in the Visio document which you want to switch to

        .PARAMETER Visio
        Optional reference to a Visio Application (used if writing to multiple diagrams at the same time?)

        .INPUTS
        None. You cannot pipe objects to Set-VisioPage

        .OUTPUTS
        None

        .EXAMPLE
        Set-VisioPage -Page 'Page-3'


#>
function Set-VisioPage{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param([string]$Name,
    $Visio=$script:Visio)
    if($PSCmdlet.ShouldProcess('Switching to a different Visio Page','')){
        $page=get-VisioPage $Name
        $Visio.ActiveWindow.Page=$page 
    }
} 

<#
        .SYNOPSIS 
        Returns a visio page

        .DESCRIPTION
        Returns either the named page or the active page if nothing was named.

        .PARAMETER Name
        The name of the page you want.  If you don't supply a name, the active page will be output.

        .INPUTS
        None. You cannot pipe objects to Get-VisioPage.

        .OUTPUTS
        Visio.Page

        .EXAMPLE
        $activePage=get-VisioPage
        #Returns the active page

        .EXAMPLE
        get-VisioPage 'Page-3'
        #returns the page named 'Page-3'


#>
Function Get-VisioPage{
    [CmdletBinding()]
    Param($Name)
    if ($Name) {
        try {
            $Visio.ActiveDocument.Pages($Name) 
        } catch {
            write-warning "$Name not found"
        }
    } else {
        $Visio.ActivePage
    }
}


<#
        .SYNOPSIS 
        Deletes a page from Visio

        .DESCRIPTION
        Deletes a named page or the active page if no page is named.

        .PARAMETER Name
        The name of the page to remove.  If no page is named, the active page is removed.

        .PARAMETER Parameter2
        Describe Parameter1

        .INPUTS
        What can be piped in
        None. You cannot pipe objects to Remove-VisioPage

        .OUTPUTS
        None

        .EXAMPLE
        Remove-VisioPage 'Page-3'
        #removes page 3

        .EXAMPLE
        Remove-VisioPage
        #removes the active page

#>
Function Remove-VisioPage{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param($Name)
    if($PSCmdlet.ShouldProcess('Removing page named <$Name> or current page','')){
        if ($Name) {
            $Visio.ActiveDocument.Pages($Name).Delete(0)
        } else {
            $Visio.ActivePage.Delete(0)
        }
    }

}


<#
        .SYNOPSIS 
        Switches the page orientation

        .DESCRIPTION
        Set the page orientation to either Landscape or Portrait

        .PARAMETER Landscape
        Changes the page orientation to Landscape

        .PARAMETER Portrait
        Changes the page orientation to Portrait

        .INPUTS
        None. You cannot pipe objects to Set-VisioPageLayout

        .OUTPUTS
        None

        .EXAMPLE
        Set-VisioPageLayout -Portrait

        .EXAMPLE
        Set-VisioPageLayout -Landscape

#>
Function Set-VisioPageLayout{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param([switch]$Landscape,[switch]$Portrait)
    if($PSCmdlet.ShouldProcess('Visio','Switch page layout')){
        if($Landscape){
            $Visio.ActivePage.Shapes['ThePage'].CellsU('PrintPageOrientation')=2
        } else {
            $Visio.ActivePage.Shapes['ThePage'].CellsU('PrintPageOrientation')=1
        }
    }
}