stop-process -Name VISIO -ea SilentlyContinue
remove-item c:\temp\VisioDSL2.vsdx -ea SilentlyContinue
import-module VisioBot3000 -Force

Diagram C:\temp\TestVisio.vsdx
Stencil Containers -From C:\temp\MyContainers.vssx
Stencil Servers -From SERVER_U.vssx
Shape WebServer -From Servers -MasterName 'Web Server'
Container Location -From Containers -MasterName 'Location'
Container Domain -From Containers -MasterName 'Domain'

Domain MyDomain {
	Location MyCity {
		WebServer PrimaryServer 5 5
	}
	Location DRSite {
		WebServer BackupServer 5 8
	}
}
