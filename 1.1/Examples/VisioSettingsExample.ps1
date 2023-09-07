Import-Module VisioBot3000 -Force
 

#start Visio and create a new document
New-VisioApplication
New-VisioDocument C:\temp\TestVisioPrimitives.vsdx 

$doc=Get-VisioDocument
Set-VisioDiagramServices -Document $doc -Value $vis.ServiceAll

#adjust path to match the location you put the setting file.
Import-VisioSettings -settings C:\Users\mike\Documents\WindowsPowerShell\modules\VisioBot3000\Examples\DiagramSettings.psd1 

#draw a container with two items in it
Domain   MyDomain {
     WebServer PrimaryServer 
     DBServer SQL01  
}

#add a connector
SQL -from PrimaryServer -To SQL01