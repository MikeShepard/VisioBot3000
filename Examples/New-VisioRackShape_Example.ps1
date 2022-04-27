<#

Download the following Stencils from https://www.visiocafe.com/

	HPE-Common - https://www.visiocafe.com/downloads/hp/HPE-Common.zip
	HPE-ProLiant - https://www.visiocafe.com/downloads/hp/HPE-ProLiant.zip

Unblock both files 

Unzip the contents into C:\Temp

#>

#Install the latest VisioBot3000 Module from PowerShell Gallery
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
Install-Module -Name VisioBot3000

#Import the VisioBot3000 Module
Import-Module -Name VisioBot3000 -Force

#Close down any running Visio session **without saving**
Stop-Process -Name VISIO -ErrorAction SilentlyContinue

#Clear All Variables
Remove-Variable -Name * -ErrorAction SilentlyContinue

#Create a New Visio Session and Blank 'Page-1'
New-VisioDocument

#Register the HPE Stencils
Register-VisioStencil -Name HPE-Racks -Path 'C:\Temp\HPE-Racks.vss'
Register-VisioStencil -Name HPE-ProLiant-DL -Path 'C:\Temp\HPE-ProLiant-DL.vss'

#Register the HPE Shapes
Register-VisioShape -Name HPE42U600EntRackF -StencilName HPE-Racks -MasterName 'HPE 42U G2 Ent Rack Front'
Register-VisioShape -Name HPEDL180Gen108LFF -StencilName HPE-ProLiant-DL -MasterName 'DL180 Gen10 8LFF front'
Register-VisioShape -Name HPEDL580Gen10SFF -StencilName HPE-ProLiant-DL -MasterName 'DL580 Gen10 front'
Register-VisioShape -Name HPEDL360Gen108SFF -StencilName HPE-ProLiant-DL -MasterName 'DL360 Gen10 8SFF front'

#Drop the New Rack Shape
New-VisioShape -Master HPE42U600EntRackF -Label Rack01 -x 1 -y 3

#Drop the New Rack Equipment Shapes into the Existing Rack Shape
New-VisioRackShape -Master HPEDL180Gen108LFF -Label Server1 -RackLabel Rack01 -RackVendor HPE -FirstU 1
New-VisioRackShape -Master HPEDL580Gen10SFF -Label Server2 -RackLabel Rack01 -RackVendor HPE -FirstU 3
New-VisioRackShape -Master HPEDL580Gen10SFF -Label Server3 -RackLabel Rack01 -RackVendor HPE -FirstU 7
New-VisioRackShape -Master HPEDL360Gen108SFF -Label Server4 -RackLabel Rack01 -RackVendor HPE -FirstU 11
New-VisioRackShape -Master HPEDL360Gen108SFF -Label Server5 -RackLabel Rack01 -RackVendor HPE -FirstU 12

