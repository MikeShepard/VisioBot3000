<#
        .SYNOPSIS 
        Adds a path to the Stencil search path list
        .DESCRIPTION
        Adds a path to the Stencil search path list.  When registering a stencil, if you only supply a filename, the function will search 
        in the folders listed in the Stencil search path for a matching file and use the first one found.
        .PARAMETER Path
        The path to add to the search path.
        .INPUTS
        None. You cannot pipe objects to Add-StencilSearchPath.
        .OUTPUTS
        None
        .EXAMPLE
        Add-StencilSearchPath 'C:\temp'
 #> 
 function Add-StencilSearchPath{
    [CmdletBinding()]
    Param([string]$Path)
    $script:StencilSearchPath.Add($Path) | Out-Null
}

<#
        .SYNOPSIS 
        Resets the stencil search path list
        .DESCRIPTION
        Resets the stencil search path list.  When registering a stencil, if you only supply a filename, the function will search 
        in the folders listed in the Stencil search path for a matching file and use the first one found.
        .PARAMETER Path
        The list of paths to set the search path to.
        .INPUTS
        None. You cannot pipe objects to Set-StencilSearchPath.
        .OUTPUTS
        None
        .EXAMPLE
        Set-StencilSearchPath 'C:\temp','C:\Program Files (x86)\Microsoft Office\Office15\Visio Content'
 #> 
 function Set-StencilSearchPath{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param([string[]]$Path)
    if($PSCmdlet.ShouldProcess("Setting Stencil Search Path to $path")){
        $script:StencilSearchPath=$Path
    }
}

<#
        .SYNOPSIS 
        Retrieves the stencil search path
        .DESCRIPTION
        Retrieves the stencil search path.  When registering a stencil, if you only supply a filename, the function will search 
        in the folders listed in the Stencil search path for a matching file and use the first one found.
        .INPUTS
        None. You cannot pipe objects to Get-StencilSearchPath.
        .OUTPUTS
        String[]
        .EXAMPLE
        Get-StencilSearchPath 
 #> 
 function Get-StencilSearchPath{
    [CmdletBinding()]
    Param( )
    $script:StencilSearchPath 
}