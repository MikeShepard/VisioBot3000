Import-Module VisioBot3000 -Force

New-VisioApplication

New-VisioDocument C:\temp\VisioPrimitives1.vsdx 
Register-VisioStencil -Name Containers -Path C:\temp\MyContainers.vssx 
Register-VisioStencil -Name Servers -Path SERVER_U.vssx
Register-VisioShape -Name WebServer -From Servers -MasterName 'Web Server'
Register-VisioContainer -Name Location -From Containers -MasterName 'Location'
Register-VisioContainer -Name Domain -From Containers -MasterName 'Domain'
Register-VisioContainer -Name Logical -From Containers -MasterName 'Logical'

New-VisioContainer -shape Logical -Label MyFarm -contents {
    New-VisioContainer -shape Location -Label MyCity -contents {
        New-VisioContainer -shape (Get-VisioShape Domain) -Label MyDomain -contents {
		    New-VisioShape -master WebServer -Label PrimaryServer -x 5 -y 5
	    }
    }
    New-VisioContainer -shape  Location  -Label DRSite -contents {
        New-VisioContainer Get-VisioShape Domain  -Label MyDomain -contents {
		    New-VisioShape -master WebServer -Label BackupServer -x 5 -y 8
	    }
    }
    New-VisioConnector -From PrimaryServer -To BackupServer -Label SQL -Color Red -Arrow
}