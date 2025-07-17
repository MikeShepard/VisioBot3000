<#
.Synopsis
   Returns the value of the DiagramServicesEnabled property of the given document
.DESCRIPTION
   Returns the value of the DiagramServicesEnabled property of the given document
.PARAMETER Document
    The document you wish to get the value of the DiagramServicesEnabled property
.EXAMPLE
    $doc=Get-VisioDocument
    Get-VisioDiagramServices -document $doc
.INPUTS
    You cannot pipe any objects to Get-VisioDiagramServices
.OUTPUTS
    Int
#>
function Get-VisioDiagramServices{
[CmdletBinding()]
Param($Document)

$Document.DiagramServicesEnabled
}

<#
.Synopsis
   Sets the value of the DiagramServicesEnabled property of the given document
.DESCRIPTION
   Sets the value of the DiagramServicesEnabled property of the given document
.PARAMETER Document
    The document you wish to get the value of the DiagramServicesEnabled property
.PARAMETER Value
    The value to set the DiagramServicesEnabled to 
.EXAMPLE
    $doc=Get-VisioDocument
    Set-VisioDiagramServices -document $doc -Value $vis.ServiceAll
.INPUTS
    You cannot pipe any objects to Get-VisioDiagramServices
.OUTPUTS
    Int
#>
function Set-VisioDiagramServices{
[CmdletBinding(SupportsShouldProcess=$true)]
Param($Document,
      [int]$Value)

    if($PSCmdlet.ShouldProcess("Set Visio Diagram Services to $value")){
        $Document.DiagramServicesEnabled=$Value
    }
}