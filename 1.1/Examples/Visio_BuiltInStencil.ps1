ipmo VisioBot3000 -Force

Diagram C:\temp\TestVisio_BuiltinStencil.vsdx
Add-StencilSearchPath c:\temp 
Stencil Containers -BuiltIn Containers
Stencil Servers -Path SERVER_U.vssx
Shape  WebServer -From Servers -MasterName 'Web Server'

Container Classic -from Containers  


Classic Fred {
        WebServer PrimaryServer
    }