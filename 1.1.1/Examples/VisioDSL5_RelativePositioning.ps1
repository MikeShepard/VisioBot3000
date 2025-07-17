#clean-up because I'm re-running this over and over
stop-process -Name VISIO -ea SilentlyContinue
remove-item c:\temp\testvisio5.vsdx -ea SilentlyContinue
import-module VisioBot3000 -Force

Diagram C:\temp\TestVisio5.vsdx -From C:\temp\IntegrationDiagram.vstx

# Define shapes, containers, and connectors for the diagram
Stencil Servers -From SERVER_U.vssx
Stencil Containers -From C:\temp\MyContainers.vssx 
Shape WebServer -From Servers -MasterName 'Web Server'
Shape SQLServer -From Servers -masterName 'Database Server'
Container Location -From Containers -MasterName 'Location'
Connector SQL -color Red -Arrow 
Set-NextShapePosition -x 3 -y 5.5
Location Datacenter {
       WebServer PrimaryServer
       WebServer SecondaryServer
       WebServer ThirdServer
       Set-RelativePositionDirection Vertical
       SQLServer DBServer
}

SQL -from PrimaryServer,SecondaryServer,ThirdServer -to DBServer

Legend @{
            'Information/CreatedBy/Name'='Mike Shepard - the boss!';
            'Information/LastUpdateBy/Name'='Mike Shepard - the boss2!';
            'Title/Title'='VisioBot3000 DSL Example';
            'Title/SubTitle'='Relative positioning and legend'}


Complete-Diagram 