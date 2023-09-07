<#
.Synopsis
   Converts Visio objects to (small) PSCustomObjects
.DESCRIPTION
   Convert-VisioObjectToPSObject creates a nested object (starting with a page or container) that contains just a few of the properties
   of the Visio object (like name, text).
.PARAMETER Object
    The object to convert to a PSCustomObject
.EXAMPLE
    Get-VisioPage | Convert-VisioObjectToPSObject | ConvertTo-Json -depth 10
.INPUTS
    You can pipe any Visio object to Convert-VisioObjectToPSObject.  I would recommend using a page, or a main container object.
.OUTPUTS
    PSCustomObject
#>
function Convert-VisioObjectToPSObject{
[CmdletBinding()]
Param([Parameter(ValueFromPipeline=$true)]$Object) 
    
    $ObjectHash=@{Name=$Object.Name;
                  Text=$Object.Text}
    if(($Object | get-member Master) -and $Object.Master){
       $ObjectHash.Add('Type',$Object.Master.Name)
    }
    if($Object | get-member PageSheet){
        #get the shapes that aren't in containers
        $ObjectHash.Add('Type','Page')
        $containedObjects= $Object.Shapes | Where-Object {$_.MemberOfContainers.Count -eq 0} |foreach-object {Convert-VisioObjectToPSObject $_}
    } elseif ($Object.Style -eq 'Connector'){
        #it's a connector
        $ObjectHash.Add('From',$Object.Connects[1].ToSheet.Name)
        $ObjectHash.Add('To',$Object.Connects[2].ToSheet.Name)
    } elseif($Object | get-member ContainerProperties){
        #get the top-level objects contained
        if($Object.ContainerProperties){
            $containedObjectIDs=$Object.ContainerProperties.GetMemberShapes(16+2) 
            $containedObjects=$ContainedObjectIDs|foreach-object {Convert-VisioObjectToPSObject ($Object.ContainingPage.Shapes | Where-Object id -EQ $_)}
        }
    }
    if($containedObjects){
        $ObjectHash.Add('Contents',$containedObjects)
    }
    [PSCustomObject]$ObjectHash

}