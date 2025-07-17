# VisioBot3000

[![Join the chat at https://gitter.im/MikeShepard/VisioBot3000](https://badges.gitter.im/MikeShepard/VisioBot3000.svg)](https://gitter.im/MikeShepard/VisioBot3000?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge&utm_content=badge)

Simple Visio Automation from Powershell

A module with some useful function definitions to do diagrams in Visio.  

Also exposes a DSL which lets you create diagrams without "scripting".

Primitive Operations
---------------------
* New-VisioApplication  
* Get-VisioApplication  
* New-VisioDocument  
* Open-VisioDocument
* New-VisioPage 
* Set-VisioPage
* Get-VisioPage
* Remove-VisioPage
* New-VisioRectangle
* Set-VisioPageLayout
* New-VisioShape
* New-VisioConnector
* New-VisioContainer
* Register-VisioBuiltInStencil
* Register-VisioStencil
* Register-VisioShape
* Register-VisioContainer
* Register-VisioConnector
* Get-VisioShape
* New-VisioHyperLink
* New-VisioSelection
* Set-VisioShapeData
* Get-VisioShapeData
* Complete-Diagram
* New-VisioLayer
* Set-NextShapePosition
* Get-NextShapePosition
* Set-RelativePositionDirection
* Set-VisioText
* Get-VisioColorFormula
* Add-StencilSearchPath
* Set-StencilSearchPath
* Get-StencilSearchPath

Aliases
-------
* Diagram (New-VisioDocument)
* Stencil (Register-VisioStencil)
* Shape (Register-VisioShape)
* Container (Register-VisioContainer)
* Connector (Register-VisioConnector)

Lots of examples in the "Examples" folder

Example using "Primitives"
--------------------------
    Import-Module VisioBot3000 -Force

    New-VisioApplication

    New-VisioDocument C:\temp\TestVisio3.vsdx 
    Register-VisioStencil -Name Containers -Path C:\temp\MyContainers.vssx 
    Register-VisioStencil -Name Servers -Path C:\temp\SERVER_U.vssx
    Register-VisioShape -Name WebServer -From Servers -MasterName 'Web Server'
    Register-VisioContainer -Name Location -From Containers -MasterName 'Location'
    Register-VisioContainer -Name Domain -From Containers -MasterName 'Domain'
    Register-VisioContainer -Name Logical -From Containers -MasterName 'Logical'

    New-VisioContainer -shape (Get-VisioShape Logical) -label MyFarm -contents {
        New-VisioContainer -shape (Get-VisioShape Location) -label MyCity -contents {
            New-VisioContainer -shape (Get-VisioShape Domain) -label MyDomain -contents {
                New-VisioShape -master WebServer -label PrimaryServer -x 5 -y 5
            }
        }
        New-VisioContainer -shape (Get-VisioShape Location) -label DRSite -contents {
            New-VisioContainer -shape (Get-VisioShape Domain) -label MyDomain -contents {
                New-VisioShape -master WebServer -label BackupServer -x 5 -y 8
            }
        }
        New-VisioConnector -From PrimaryServer -To BackupServer -Label SQL -Color Red -Arrow
    }


Same Example using DSL
-----------------------
    Diagram C:\temp\TestVisio3.vsdx 
    Stencil Containers -From C:\temp\MyContainers.vssx 
    Stencil Servers -From C:\temp\SERVER_U.vssx
    Shape WebServer -From Servers -MasterName 'Web Server'
    Container Location -From Containers -MasterName 'Location'
    Container Domain -From Containers -MasterName 'Domain'
    Container Logical -From Containers -MasterName 'Logical'
    Connector SQL -Color Red -arrow


    Logical MyFarm {
        Location MyCity {
            Domain MyDomain {
                WebServer PrimaryServer 5 5
            }
        }
        Location DRSite {
            Domain MyDomain {
                WebServer BackupServer 5 8
            }
        }
        SQL -From PrimaryServer -To BackupServer
    }