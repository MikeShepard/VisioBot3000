#clean-up because I'm re-running this over and over
stop-process -Name VISIO -ea SilentlyContinue
import-module VisioBot3000 -Force
#remove-item c:\temp\testvisio4.vsdx -force -silentlycontinue

Diagram C:\temp\TestVisio4.vsdx -Update

# Define shapes, containers, and connectors for the diagram
Stencil Containers -From C:\temp\MyContainers.vssx 
Stencil Servers -From SERVER_U.vssx
Shape WebServer -From Servers -MasterName 'Web Server'
Container Location -From Containers -MasterName 'Location'
Container Domain -From Containers -MasterName 'Domain'
Container Logical -From Containers -MasterName 'Logical'
Connector SQL -Color Red -arrow 

#this is the diagram
Logical MyFarm {
    Location MyCity {
        Domain MyDomain  {
            WebServer PrimaryServer
            WebServer HotSpare

        }
    }
    Location DRSite {
        Domain MyDomain -name SiteB_MyDomain {
            Set-RelativePositionDirection Vertical
		    WebServer BackupServer 
            WebServer DRHotSpare
	    }
    }
}
SQL -From PrimaryServer -To BackupServer 
Hyperlink $BackupServer -link http://google.com

Complete-Diagram 