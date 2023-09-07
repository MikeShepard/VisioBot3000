ipmo VisioBot3000 -Force

Diagram C:\temp\TestVisio_CustomConnector.vsdx

Stencil Servers -Path SERVER_U.vssx
Stencil Connectors -path CONNEC_U.VSSX

Shape WebServer -From Servers -MasterName 'Web Server'
Shape CurveConnector -From Connectors -MasterName 'Curve Connect 1'
Connector Curve -Color Red -Master CurveConnector
WebServer Server1
WebServer Server2

Curve -From Server1 -To Server2  

Complete-Diagram