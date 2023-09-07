<#
        .SYNOPSIS 
        Connects two shapes

        .DESCRIPTION
        Creates a connector object between two previously drawn shapes.

        .PARAMETER From
        The shape that the connector will originate from

        .PARAMETER To
        The shape that the connector will end on

        .PARAMETER Name
        The name to assign to the connector shape

        .PARAMETER Color
        The color to draw the connector

        .PARAMETER Arrow
        Determines whether an arrow is drawn on the connector at the final end

        .PARAMETER Bidirectional
        Determines whether an arrow is drawn on the connector at the originating end

        .PARAMETER Label
        The text to be shown on the arrow


        .INPUTS
        None. You cannot pipe objects to New-VisioConnector.

        .OUTPUTS
        Visio.Shape

        .EXAMPLE
        $arrow = New-VisioConnector -From WebServer -To SQLServer -name SQLConnection -Arrow -color Red -label SQL
        File.txt


#>
Function New-VisioConnector{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param([array]$From,
        [array]$To,
        [string]$Name,
        [System.Drawing.Color]$Color,
        [switch]$Arrow,
        [switch]$bidirectional, 
        [string]$Label,
        $Master)
    $ColorFormula=get-VisioColorFormula $color
    if($PSCmdlet.ShouldProcess('Visio','Connect shapes with a connector')){
        $CurrentPage=Get-VisioPage
        foreach($dest in $To){
            foreach($source in $From){
                if($source -is [string]){
                    $source=$CurrentPage.Shapes[$source]
                }
                if($dest -is [string]){
                    $dest=$CurrentPage.Shapes[$dest]
                }

                $CalculatedName='{0}_{1}_{2}' -f $Label,$source.Name,$dest.Name
                if($updatemode){
                    $connector=$CurrentPage.Shapes | Where-Object {$_.Name -eq $CalculatedName}
                }
                if (-not (get-variable Connector -Scope Local -ErrorAction Ignore)){
                    if($Master){
                        if($Master -is [string]){
                            $Master=Get-VisioShape $Master
                            $ConnectorNameToFind=$Master.Name
                        }
                        $source.AutoConnect($dest,0,$Master)
                    } else { 
                        $ConnectorNameToFind='Dynamic Connector'
                        $source.AutoConnect($dest,0)
                    }
                    $connector=$CurrentPage.Shapes($ConnectorNameToFind)| Select-Object -first 1
                    $connector.Name=$CalculatedName
                }
                $connector.Text=$Label
                $connector.CellsU('LineColor').Formula=$ColorFormula
                $connector.CellsSRC(1,23,10) = 16
                $connector.CellsSRC(1,23,19) = 1 

                if($Arrow){
                    $connector.Cells('EndArrow')=5
                    if($bidirectional){ 
                        $connector.Cells(‘BeginArrow')=5
                    } else {
            
                    }
                } else {
                    $connector.Cells('EndArrow')=0
                    $connector.Cells('BeginArrow')=0
                }
                Remove-variable Connector -ErrorAction SilentlyContinue
            }
        }
    }

}
<#
        .SYNOPSIS 
        Saves a "nickname" for a certain style of connector

        .DESCRIPTION
        Saves a "nickname" for a certain style of connector

        .PARAMETER Name
        The name to use to refer to this style of connector

        .PARAMETER Color
        The color to draw the connector in

        .PARAMETER Arrow
        Whether to put an arrow at the end of the connector

        .PARAMETER Bidirectional
        Whether to put an arrow at the beginning of the connector


        .INPUTS
        None. You cannot pipe objects to Register-VisioConnector.

        .OUTPUTS
        None

        .EXAMPLE
        Register-VisioConnector -Name HTTP -Color Black -Ar
#>
Function Register-VisioConnector{
    [CmdletBinding()]
    Param([string]$Name,
        [System.Drawing.Color]$Color,
        [switch]$Arrow,
        [switch]$bidirectional,
    $Master)
    new-item -Path Function:\ -Name "global`:$Name" -value {param($From,$To,$Label) New-VisioConnector -from $From -to $To -name $Name -color $Color -Arrow:$Arrow.IsPresent -bidirectional:$bidirectional.IsPresent $Label -Master $Master}.GetNewClosure() -force  | out-null
    $script:GlobalFunctions.Add($Name) | Out-Null
}

