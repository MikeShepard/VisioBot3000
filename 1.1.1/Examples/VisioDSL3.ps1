#clean-up because I'm re-running this over and over
stop-process -Name VISIO -ea SilentlyContinue
remove-item c:\temp\testvisio3.vsdx -ea SilentlyContinue
import-module VisioBot3000 -Force

Diagram C:\temp\TestVisio3.vsdx

# Define shapes, containers, and connectors for the diagram
Stencil Containers -From C:\temp\MyContainers.vssx
Stencil Servers -From SERVER_U.vssx
Shape WebServer -From Servers -MasterName 'Web Server'
Container Location -From Containers -MasterName 'Location'
Container Domain -From Containers -MasterName 'Domain'
Container Logical -From Containers -MasterName 'Logical'
Connector SQL -Color Red -arrow

#this is the diagram
Set-NextShapePosition -x 3.5 -y 7
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

	    }
    }
}
SQL -From PrimaryServer -To BackupServer
Hyperlink $BackupServer -link http://google.com

Complete-Diagram