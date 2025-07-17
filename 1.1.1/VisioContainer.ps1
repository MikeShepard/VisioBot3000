<#
        .SYNOPSIS 
        Draws a container

        .DESCRIPTION
        Draws a container around previously drawn shapes on the active page.

        .PARAMETER Name
        The name to assign to the dropped shape

        .PARAMETER Conents
        A scriptblock which, when executed, outputs the objects to be contained in the container

        .PARAMETER Shape
        The master shape to use to draw the container

        .PARAMETER Label
        The text to label the container with

        .INPUTS
        None. You cannot pipe objects to New-VisioContainer.

        .OUTPUTS
        Visio.Shape

        .EXAMPLE
        New-VisioContainer -shape (Get-VisioShape Domain) -label MyDomain -contents {
            New-VisioShape -master WebServer -label BackupServer -x 5 -y 8
        }

#>
Function New-VisioContainer {
    [CmdletBinding(SupportsShouldProcess = $True)]
    Param( [string]$Name,
        [Scriptblock]$Contents,
        $Shape,
        $Label)

    [void]$script:Visio.ActiveDocument.Pages[1]
         
    if (!$Name) {
        $Name = $Label
    }
    if ($shape -is [string]) {
        $shape = Get-VisioShape $shape
    }
    if ($PSCmdlet.ShouldProcess('Visio', 'Drop a container around shapes')) {
        $page = Get-VisioPage
        if ($Contents) {
            [array]$containedObjects = & $Contents
            $firstShape = $containedObjects[0]
            if ($updatemode) {
                $droppedContainer = $page.Shapes | Where-Object {$_.Name -eq $Name}
            }
        
            if (get-variable droppedContainer -Scope Local -ErrorAction Ignore) {
                If ($droppedContainer.ContainerProperties.GetMemberShapes(16 + 2) -notcontains $firstShape.ID) {
                    $droppedcontainer.ContainerProperties.AddMember($firstShape, 2)
                }
            } else {
                $droppedContainer = $page.DropContainer($Shape, $page.Application.ActiveWindow.Selection)
                $droppedContainer.Name = $Name
            } 
            $droppedContainer.ContainerProperties.SetMargin($vis.PageUnits, 0.25)
            $containedObjects | select-object -Skip 1 | foreach-object { 
                if (-not $updatemode -or ($droppedContainer.ContainerProperties.GetMemberShapes(16 + 2) -notcontains $_.ID)) {
                    $droppedcontainer.ContainerProperties.AddMember($_, 1)
                }
            }        
            $droppedContainer.ContainerProperties.FitToContents()
            $droppedContainer.Text = $Label
            $Script:LastDroppedObject = $droppedContainer

            $droppedContainer

        }
        New-Variable -Name $Name -Value $droppedContainer -Scope Global -Force
    }
}

<#
        Copies a master for a container from a stencil and gives it a name.

        .DESCRIPTION
        Copies a master for a container from a stencil and gives it a name.  Also creates a function with the same name to drop the container (with contents) onto the active Visio page.

        .PARAMETER Name
        The name used to refer to the shape

        .PARAMETER StencilName
        Which stencil to get the master from

        .PARAMETER MasterName
        The name of the master in the stencil

        .INPUTS
        None. You cannot pipe objects to Register-VisioContainer.

        .OUTPUTS
        None

        .EXAMPLE
        Register-VisioContainer -Name BasicContainer -StencilName Containers -MasterName Plain

#>
Function Register-VisioContainer {
    [CmdletBinding()]
    Param([string]$Name,
        [Alias('From')][string]$StencilName,
        [string]$MasterName)
 
    if (!$MasterName) {
        $MasterName = $Name
    }
    $newShape = $stencils[$StencilName].Masters | Where-Object {$_.Name -eq $MasterName}
    $script:Shapes[$Name] = $newshape
    $outerName = $Name
    new-item -Path Function:\ -Name "global`:$outername" -value {param($Label, $Contents, $Name) New-VisioContainer -label $Label -contents $Contents -shape $outername -name $Name}.GetNewClosure() -force  | out-null
    $script:GlobalFunctions.Add($outername) | Out-Null
}
