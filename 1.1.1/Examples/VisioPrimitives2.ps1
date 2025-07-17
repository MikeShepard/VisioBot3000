Import-Module VisioBot3000 -Force
 

#start Visio and create a new document
New-VisioApplication
New-VisioDocument C:\temp\TestVisioPrimitives.vsdx 

$doc=Get-VisioDocument
Set-VisioDiagramServices -Document $doc -Value $vis.ServiceAll

#tell Visio what Stencils I want to use and give them "nicknames"
Register-VisioStencil -Name Containers -Path C:\temp\MyContainers.vssx 
Register-VisioStencil -Name Servers -Path SERVER_U.vssx

#pick a master from one of those stencils and give it a nickname
Register-VisioShape -Name WebServer -From Servers -MasterName 'Web Server'
Register-VisioShape -Name DBServer -From Servers -MasterName 'Database Server'



#pick another master (this time a container) and give it a nickname
#note that this is a different cmdlet
Register-VisioContainer -Name Domain -From Containers -MasterName 'Domain'

#draw a container with two items in it
New-VisioContainer -shape Domain -Label MyDomain -contents {
   New-VisioShape -master WebServer -Label PrimaryServer -x 5 -y 5
   New-VisioShape -master DBServer -Label SQL01 -x 5 -y 7
}

#add a connector
New-VisioConnector -from PrimaryServer -to SQL01 -Label SQL -color Red -Arrow