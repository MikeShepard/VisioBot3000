<#
        .SYNOPSIS 
        Loads a stencil and gives it a name

        .DESCRIPTION
        Loads a stencil and gives it a name

        .PARAMETER Name
        The name to use to refer to the stencil

        .PARAMETER Path
        The path to the stencil file.  Ignored with -Builtin

        .PARAMETER BuiltIn
        Flags that Path (or Name) refer to a built-in stencil.

        .INPUTS
        None. You cannot pipe objects to Register-VisioStencil.

        .OUTPUTS
        None

        .EXAMPLE
        Register-VisioStencil -Name Containers -Path 'c:\temp\my containers.vssx'

        .EXAMPLE
        Register-VisioStencil -Name Connectors -Builtin

#>
Function Register-VisioStencil{
    [CmdletBinding()]
    Param([string]$Name,
        [Alias('From')][string]$Path,
    [ValidateSet('Backgrounds','Borders','Containers','Callouts','Legends')][string]$BuiltIn)
    if($BuiltIn){
        $stencilID=@('Backgrounds','Borders','Containers','Callouts','Legends').IndexOf($BuiltIn)
        $stencilPath=$Visio.GetBuiltInStencilFile($stencilID,$vis.MSDefault)
        $stencil=$Visio.Documents.OpenEx($stencilPath,$vis.OpenHidden)
         
    } else {
        if($Path -eq (split-path -path $Path -leaf)){
            #if the path is just a filename
            if(-not(test-path $path)){
                #and the filename doesn't exist in the current directory
                foreach($folder in $StencilSearchPath){
                    if (test-path (join-path -Path $folder -ChildPath $path)){
                        $path=join-path -Path $folder -ChildPath $path
                        break
                    }
                }
            }
        }
        if (test-path $path){
            $stencil=$Visio.Documents.OpenEx($Path,$vis.OpenHidden)
        } else {
            write-error "$path not found when registering the stencil"
        }
    }  
    $script:stencils[$Name]=$stencil
}