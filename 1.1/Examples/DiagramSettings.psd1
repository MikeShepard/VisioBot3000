@{
StencilPaths='c:\temp'

Stencils=@{Containers='C:\temp\MyContainers.vssx';
           Servers='SERVER_U.vssx'}

Shapes=@{WebServer='Servers','Web Server';
         DBServer='Servers','Database Server'
        }
Containers=@{Domain='Containers','Domain'
            }

Connectors=@{SQL=@{Color='Red';Arrow=$true}}

}