<#
        .SYNOPSIS 
        Sets the text of multiple objects that have already been created
        .DESCRIPTION
        Sets the text of multiple objects that have already been created. Useful for updating static objects that are part of a template, such as title, author, etc.
        Nested objects can be referenced with a slash (e.g. TitleBox/Title, or TitleBox/Subtitle)
        .PARAMETER Map
        A hashtable mapping names of objects to the text you want them to have.
        .INPUTS
        None. You cannot pipe objects to Set-VisioText.
        .OUTPUTS
        None
        .EXAMPLE
        Set-VisioText -Map @{Title='My first Visio Digram';
        Author='Mike';
        CreatedOn="$(get-date)"}
#> 

function Set-VisioText{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param([Hashtable]$Map)
    
    foreach($key in $Map.Keys){
        $originalKey-$key
        $text=$map[$key]
        $p=Get-VisioPage
        while($key.Contains('/')){
            $prefix,$key=$key.split('/',2)
            $p=$p.Shapes[$prefix]
        }
        $Shape=$p.Shapes[$key]
        if($PSCmdlet.ShouldProcess("Setting $OriginalKey to $text")){
            $Shape.Characters.Text="$text"
        }
    }
} 