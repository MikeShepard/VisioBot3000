<#
        .SYNOPSIS 
        Starts the visio application (possibly hidden) and stores a reference to the application object

        .DESCRIPTION
        Starts the visio application (possibly hidden) and stores a reference to the application object

        .PARAMETER Hide
        Starts Visio without showing the user interface

        .INPUTS
        None. You cannot pipe objects to Add-Extension.

        .OUTPUTS
        None

        .EXAMPLE
        New-VisioApplication
        --Visio Pops up--

        .EXAMPLE
        New-VisioApplication -Hide
        --Nothing seems to happen

#>
Function New-VisioApplication{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param([switch]$Hide,
          [switch]$Passthru)
    if($PSCmdlet.ShouldProcess('Creating a new instance of Visio','')){
        if ($Hide){
            $script:Visio=New-Object -ComObject Visio.InvisibleApp 
        } else {
            $script:Visio = New-Object -ComObject Visio.Application
        }
        if($Passthru){
            $script:Visio
        }
        [System.Collections.ArrayList]$script:StencilSearchPath=@(join-path (join-path $script:Visio.Path 'Visio Content') $script:Visio.Language)
    }
}


<#
        .SYNOPSIS 
        Ouptuts a reference to the Visio application object

        .DESCRIPTION
        Ouptuts a reference to the Visio application object


        .INPUTS
        None. You cannot pipe objects to Add-Extension.

        .OUTPUTS
        Visio.Application
        Visio.InvisibleApp

        .EXAMPLE
        $app=Get-VisioApplication
#>
Function Get-VisioApplication{
    [CmdletBinding()]
    Param()
    if(!(Test-VisioApplication)){
        New-VisioApplication
    }
    return $Visio
} 

<#
        .SYNOPSIS 
        Outputs $true if the stored Visio application object is live

        .DESCRIPTION
        Outputs $true if the stored Visio application object is live

        .INPUTS
        None. You cannot pipe objects to Add-Extension.

        .OUTPUTS
        Boolean

        .EXAMPLE
        If(Test-VisioApplication){ #do something with the application }
#>
Function Test-VisioApplication{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    Param()
    $Script:Visio -and $Script:Visio.Documents
}