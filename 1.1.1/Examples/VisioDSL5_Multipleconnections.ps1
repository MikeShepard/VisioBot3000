#clean-up because I'm re-running this over and over
stop-process -Name VISIO -ea SilentlyContinue
remove-item c:\temp\testvisio5.vsdx -ea SilentlyContinue
import-module VisioBot3000 -Force

Diagram C:\temp\TestVisio5.vsdx  

# Define shapes, containers, and connectors for the diagram
Stencil Servers -From SERVER_U.vssx
Shape WebServer -From Servers -MasterName 'Web Server'
Shape SQLServer -From Servers -MasterName 'Database Server'
Connector SQL -Color Red -arrow 

#this is the diagram
WebServer Web1
WebServer Web2
WebServer Web3
Set-RelativePositionDirection Vertical
SQLServer DB1
SQLServer DB2

SQL -from Web1,Web2,Web3 -to DB1,DB2



Complete-Diagram 