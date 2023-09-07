Set-StrictMode -Version Latest

#Need System.Drawing for Colors.
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') | out-null


if(-not $PSScriptRoot)
{
    $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
}

#Get public and private function definition files.
$Public  = Get-ChildItem $PSScriptRoot\*.ps1 -ErrorAction SilentlyContinue
#$Private = Get-ChildItem $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue 

#Dot source the files
Foreach($import in @($Public))
{
    Try
    {
        #PS2 compatibility
        if($import.fullname)
        {
            . $import.fullname
        }
    }
    Catch
    {
        Write-Error "Failed to import function $($import.fullname): $_"
    }
}








