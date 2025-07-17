

<#
        .SYNOPSIS 
        Opens a visio document

        .DESCRIPTION
        Opens an existing Visio document, a blank Visio Document 

        .PARAMETER Path
        The path to an existing document, and empty string (to create a blank document) or the path to a Visio template (vstx, etc.)

        .PARAMETER Visio
        Optional reference to a Visio Application (used if writing to multiple diagrams at the same time?)

        .PARAMETER Update
        Switch indicating that we're updating a diagram, potentially created with VisioBot3000

        .INPUTS
        None. You cannot pipe objects to Add-Extension.

        .OUTPUTS
        None

        .EXAMPLE
        Open-VisioDocument 
        --Creates a blank document--

        .EXAMPLE
        Open-VisioDocument .\MySampleVisio.vsdx
        --Opens the named document

        .EXAMPLE
        Open-VisioDocument .\MyVisioTemplate.vstx
        --Creates a Visio template for editing (not a new document based on the template)

#>
Function Open-VisioDocument{
    [CmdletBinding()]
    Param([string]$Path,
        $Visio=$script:Visio,
    [switch]$Update)
    if(!(Test-VisioApplication)){
        New-VisioApplication
        $Visio=$script:Visio
    }
    $documents = $Visio.Documents
    $documents.Add($Path) | out-null
    if($Update){
        $script:updatemode=$True 
    }
}


<#
        .SYNOPSIS 
        Creates a new document

        .DESCRIPTION
        Creates a new document

        .PARAMETER Path
        The path you want to save the document to 

        .PARAMETER From
        The path to a template file to create the new document from 

        .PARAMETER Visio
        Optional reference to a Visio Application (used if writing to multiple diagrams at the same time?)

        .PARAMETER Update
        Switch indicating that we're updating a diagram, potentially created with VisioBot3000

        .INPUTS
        None. You cannot pipe objects to Add-Extension.

        .OUTPUTS
        None

        .EXAMPLE
        New-VisioDocument 
        --Creates a blank document--

        .EXAMPLE
        New-VisioDocument .\MySampleVisio.vsdx
        --Opens the named document

        .EXAMPLE
        New-VisioDocument .\MyVisioTemplate.vstx
        --Creates a new document based on a Visio template  

#> 
function New-VisioDocument{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param([string]$Path,
        [string]$From='',
        $Visio=$script:visio,
        [switch]$Update,
    [switch]$Landscape,[switch]$portrait)
    if($PSCmdlet.ShouldProcess('Creating a new Visio Document','')){
        if(!(Test-VisioApplication)){
            New-VisioApplication
            $Visio=$script:Visio
        }
        if($Update){
            if($From -ne ''){
                Write-Warning 'New-VisioDocument: -From ignored when -Update is present'
            }
            Open-VisioDocument $Path -Update
        } else {
            Open-VisioDocument $From

        }
        
        if($Landscape){
            $Visio.ActiveDocument.DiagramServicesEnabled=8
            $Visio.ActivePage.Shapes['ThePage'].CellsU('PrintPageOrientation')=2
        } elseif ($portrait) {
            $Visio.ActivePage.Shapes['ThePage'].CellsU('PrintPageOrientation')=1
        }
        if($path){
            $Visio.ActiveDocument.SaveAs($Path) | Out-Null
        }
    }
}

<#
        .SYNOPSIS 
        Outputs the active Visio document

        .DESCRIPTION
        Outputs the active Visio document

        .PARAMETER Visio
        Optional reference to a Visio Application (used if writing to multiple diagrams at the same time?)

        .INPUTS
        None. You cannot pipe objects to Get-VisioDocument.

        .OUTPUTS
        visio.Document

        .EXAMPLE
        $doc=Get-VisioDocument

#>
Function Get-VisioDocument{
    [CmdletBinding()]
    Param($Visio=$script:Visio)
    if(!(Test-VisioApplication)){
        New-VisioApplication 
    }
    return $Visio.ActiveDocument
}

<#
        .SYNOPSIS 
        Saves the diagram and optionally exits Visio

        .DESCRIPTION
        Saves the diagram and optionally exits Visio

        .PARAMETER Close
        Whether to exit Visio or not

        .INPUTS
        None. You cannot pipe objects to Complete-Diagram.

        .OUTPUTS
        None

        .EXAMPLE
        Complete-Diagram
#>
Function Complete-VisioDocument{
    [CmdletBinding()]
    Param([switch]$Close)
    if(Test-VisioApplication){
        $script:updateMode=$false
        $Visio.ActiveDocument.Save() 
        if($Close){
            $Visio.Quit()
        }
        foreach($name in $script:GlobalFunctions){
            remove-item -Path "Function`:$name"
        }
    } else {
        Write-Warning 'Visio application is not loaded'
    }
}