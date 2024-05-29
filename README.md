
for [Español](#español) 

<h2 style="color: Black; font-size: 28px;">From Classroom to Boardroom: A Beginner's Guide to Building Azure Functions for Business Impact</h2>
Welcome to our guide on setting up and using Azure Functions! This material is specifically crafted for those who are new to cloud computing and serverless architectures. Whether you're a student stepping into the world of technology or a hobbyist looking to expand your skills, this guide is intended to provide you with a clear, step-by-step introduction to developing with Azure Functions.

Please note, the instructions and setups described here are designed for academic and learning purposes—they are simplified to facilitate understanding and are not meant for deployment in production environments. Our goal is to help you grasp the fundamentals without getting bogged down in the complexities often faced by experts and professionals in the field.

Additionally, this guide emphasizes learning with value. The practical applications and use cases included not only introduce you to Azure Functions but also demonstrate how these skills can be applied to real-world business scenarios. By focusing on both educational and business aspects, you will see how mastering Azure Functions can lead to impactful solutions in a professional setting.

So, let’s get started on this exciting journey into the world of serverless applications with Azure Functions!
___

<h2 style="color: #4BA3C7; font-size: 28px;">Our Business Goal</h2>
The focus of our upcoming use case is to manage SharePoint Online more effectively by setting version limits on document libraries. Implementing this approach helps control the default storage size, thus avoiding unnecessary expenses that come with purchasing additional storage space. By creating a function to streamline your SharePoint environment, we not only enhance your understanding of Azure Functions but also tackle a prevalent business challenge—optimizing resource usage in cloud storage. This practical application serves as a crucial step towards a more efficient and cost-effective document management system in SharePoint Online. Let's move forward and create a function that embodies these objectives.

___

<h2 style="color: #4BA3C7; font-size: 28px;">Lab Environment Setup</h2>

This lab is created on a Windows 11 system using Visual Studio Code with an SSH connection to an Ubuntu 22 virtual machine hosted on Hyper-V, Microsoft's virtualization component in Windows 11. It is recommended that the Ubuntu operating system and all components installed on it are consistently updated to their latest versions, ensuring a cutting-edge development environment for exploration.

If you need a step-by-step guide to set up this lab, don't hesitate to reach out

___

<h2 style="color: #4BA3C7; font-size: 28px;">Requirements</h2>

### Azure PowerShell Module (Az)

**Purpose:** Required for connecting to Azure services, like Azure Key Vault, and performing authentication.

**Usage:** Uses Connect-AzAccount for authenticating with Azure services.

**Installation:**
~~~ bash
sudo apt-get install -y powershell
~~~

Then within PowerShell
~~~ powershell
Install-Module -Name Az -AllowClobber -Scope CurrentUser
~~~

**Version Check:**
~~~ powershell
Get-InstalledModule -Name Az
~~~
Output
~~~ powershell
PS /home/user/learning/az/Set-Version-Limit-Function> Get-InstalledModule -Name Az            

Version              Name                                Repository           Description
-------              ----                                ----------           -----------
12.0.0               Az                                  PSGallery            Microsoft Azure PowerShell - Cmdlets to manage resources in Azure. This module is comp…
~~~

**Environment:** Local and Cloud

### .NET Core SDK
**Relevance:** Required for local development and testing of Azure Functions, ensuring compatibility with Azure Functions V4.

**Installation:**
~~~ bash
wget https://packages.microsoft.com/config/ubuntu/$(lsb_release -rs)/packages-microsoft-prod.deb -O packages-microsoft-prod.deb
sudo dpkg -i packages-microsoft-prod.deb
sudo apt-get update
sudo apt-get install -y dotnet-sdk-6.0
~~~

Version Check:
~~~ bash
dotnet --version
~~~
Output
~~~ bash
6.0.128
~~~

**Environment:** Local (primarily for testing and development)


### Azure Functions Environment
**Purpose:** Provides serverless compute capabilities to host and execute scripts without managing the underlying infrastructure.

**Installation:**

~~~ bash
sudo npm install -g azure-functions-core-tools@4 --unsafe-perm true
sudo apt-get update
sudo apt-get install azure-functions-core-tools-4
~~~

**Version Check:**
~~~ bash
func --version
~~~ 
Output
~~~ bash
4.0.5801
~~~

**Environment:** Local and Cloud

### SharePoint Online Management Shell (PnP PowerShell)
**Purpose:** Provides cmdlets for managing SharePoint Online, such as connecting to SharePoint and managing document libraries.
Key Commands: Connect-PnPOnline, Set-PnPList.

**Installation:**
~~~ powershell
Install-Module -Name PnP.PowerShell -AllowClobber -Scope CurrentUser
~~~

**Version Check:**
~~~ powershell
Get-InstalledModule -Name PnP.PowerShell
~~~
Output
~~~ powershell
PS /home/user/learning/az/Set-Version-Limit-Function> Get-InstalledModule -Name PnP.PowerShell     

Version              Name                                Repository           Description
-------              ----                                ----------           -----------
2.4.0                PnP.PowerShell                      PSGallery            Microsoft 365 Patterns and Practices PowerShell Cmdlets
~~~

**Environment:** Local and Cloud

### Azure Key Vault

**Purpose:** Securely stores sensitive information such as credentials, used by scripts to authenticate with services like SharePoint Online.

**Environment:** Cloud (primarily for production)
___


<h2 style="color: #4BA3C7; font-size: 28px;">Step-by-step Guide to Creating a PowerShell Function to Limit the Version of a SharePoint Online Document Library</h2>

### Step 1: Initialize a New Azure Functions Project
Open a terminal in VS Code, navigate to the  (e.g. /home/user/learning/az ) where you want to create the project (Set-Version-Limit-Function), and run:

 ~~~ bash
func init Set-Version-Limit-Function --worker-runtime powershell
 ~~~

Navigate into Set-Version-Limit-Function (your project directory created)

 ~~~ bash
cd Set-Version-Limit-Function 
 ~~~

### Step 2: Create a New Function
Create a new function (called Set-Version-Limit) within the function app using the Azure Functions Core Tools.
~~~ bash
func new --name Set-Version-Limit --template "HTTP trigger" --authlevel anonymous
~~~

### Step 3: Develop Set-Version-Limit Function
Open run.ps1 (/home/user/learning/az/Set-Version-Limit-Function/Set-Version-Limit) in VS Code, and replease its contents with this code

~~~ powershell
using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)
Write-Host "PowerShell HTTP trigger function processed a request."

# Import PnP PowerShell Module
Import-Module PnP.PowerShell -ErrorAction Stop
Add-Type -AssemblyName "Microsoft.Extensions.Logging.Abstractions"

# Retrieve site URL from the query parameters
$siteUrl = $Request.Query.siteUrl
if (-not $siteUrl) {
    $response = @{
        status  = "Failed"
        message = "No site URL provided."
    }
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::BadRequest
        Body = ($response | ConvertTo-Json)
    })
    return
}

# Define your SharePoint credentials from environment variables
$username = $env:SP_USERNAME
$password = ConvertTo-SecureString -String $env:SP_PASSWORD -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($username, $password)

# Define version limit from environment variables
$versionLimit = 12

# Initialize success and failure lists
$successLibraries = @()
$failedLibraries = @()

# Define response structure
$response = @{
    status  = "Unknown"
    message = ""
}

# Connect to SharePoint Online
try {
    Connect-PnPOnline -Url $siteUrl -Credentials $cred
    Write-Output "Connected successfully to $siteUrl."
} catch {
    Write-Host "Failed to connect to ${siteUrl}: $_"
    $response.status = "Failed"
    $response.message = "Failed to connect to ${siteUrl}: $($_)"
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::InternalServerError
        Body = ($response | ConvertTo-Json)
    })
    exit
}

# Get all document libraries in the site
try {
    $libraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 }
    if ($libraries.Count -eq 0) {
        Write-Host "No document libraries found at ${siteUrl}."
        $response.status = "Failed"
        $response.message = "No document libraries found at ${siteUrl}."
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::NotFound
            Body = ($response | ConvertTo-Json)
        })
        exit
    }

    # Set the version limit for each document library
    foreach ($library in $libraries) {
        try {
            Set-PnPList -Identity $library -EnableVersioning $true -MajorVersions $versionLimit
            Write-Output "Version limit set to $versionLimit for $($library.Title) at $siteUrl."
            $successLibraries += $library.Title
        } catch {
            Write-Host "Error setting version limits for $($library.Title) at ${siteUrl} : $_"
            $failedLibraries += $library.Title
        }
    }

    $response.status = "Success"
    $response.message = "$($successLibraries.Count) libraries had their version limit set to $versionLimit; $($failedLibraries.Count) libraries failed to set limit."
    $response.successLibraries = $successLibraries
    $response.failedLibraries = $failedLibraries
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body = ($response | ConvertTo-Json)
    })
} catch {
    Write-Host "Error managing document libraries at ${siteUrl}: $_"
    $response.status = "Failed"
    $response.message = "Error managing document libraries at ${siteUrl} : $($_)"
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::InternalServerError
        Body = ($response | ConvertTo-Json)
    })
}

~~~

### Step 4: Run and Test Locally
Test Set-Version-Limit function locally:

~~~ bash
func start
~~~

~~~ bash
Azure Functions Core Tools
Core Tools Version:       4.0.5801 Commit hash: N/A +5ac2f09758b98257e728dd1b5576ce5ea9ef68ff (64-bit)
Function Runtime Version: 4.34.1.22669


Functions:

        Set-Version-Limit: [GET,POST] http://localhost:7071/api/Set-Version-Limit

For detailed output, run func with --verbose flag.
[2024-05-29T08:04:54.025Z] Worker process started and initialized.
~~~

Open a new terminal in VS Code and test the function 

~~~ bash
curl -X POST "http://localhost:7071/api/version-limit?siteUrl=https://softwareone.sharepoint.com/sites/TestMLV"
~~~

~~~ bash
{
  "successLibraries": [
    "Documents",
    "Form Templates",
    "Site Assets",
    "VersionLimitLibrary"
  ],
  "failedLibraries": [
    "Style Library"
  ],
  "status": "Success",
  "message": "4 libraries had their version limit set to 12; 1 libraries failed to set limit."
}%  
~~~

___
## Español

<h2 style="color: Black; font-size: 28px;">De la Teoría a la Práctica: Guía para Principiantes en el Desarrollo de Funciones de Azure con Fines Empresariales</h2>
¡Bienvenido a nuestra guía sobre cómo configurar y usar Funciones de Azure! Este material está diseñado específicamente para aquellos que son nuevos en la computación en la nube y en las arquitecturas sin servidor. Ya sea que seas un estudiante entrando en el mundo de la tecnología o un aficionado buscando expandir tus habilidades, esta guía tiene como objetivo proporcionarte una introducción clara y paso a paso sobre cómo desarrollar con Funciones de Azure.
Por favor, ten en cuenta que las instrucciones y configuraciones descritas aquí están diseñadas para fines académicos y de aprendizaje; están simplificadas para facilitar la comprensión y no están destinadas para su despliegue en entornos de producción. Nuestro objetivo es ayudarte a entender los fundamentos sin que te sientas abrumado por las complejidades que a menudo enfrentan los expertos y profesionales en el campo.

Además, esta guía enfatiza el aprendizaje con valor. Las aplicaciones prácticas y casos de uso incluidos no solo te presentan las Funciones de Azure, sino que también demuestran cómo estas habilidades pueden aplicarse a escenarios empresariales del mundo real. Al enfocarnos tanto en aspectos educativos como empresariales, verás cómo dominar las Funciones de Azure puede llevar a soluciones impactantes en un entorno profesional.

¡Así que empecemos este emocionante viaje en el mundo de las aplicaciones sin servidor con Funciones de Azure!
___

<h2 style="color: #4BA3C7; font-size: 28px;">Nuestro Objetivo Empresarial</h2>
El enfoque de nuestro próximo caso de uso es gestionar SharePoint Online de manera más efectiva estableciendo límites de versión en las bibliotecas de documentos. Implementar este enfoque ayuda a controlar el tamaño de almacenamiento predeterminado, evitando así gastos innecesarios derivados de la compra de espacio de almacenamiento adicional. Al crear una función para optimizar tu entorno de SharePoint, no solo mejoramos tu comprensión de las Funciones de Azure, sino que también abordamos un desafío empresarial común: optimizar el uso de recursos en el almacenamiento en la nube. Esta aplicación práctica sirve como un paso crucial hacia un sistema de gestión de documentos más eficiente y rentable en SharePoint Online. Avancemos y creemos una función que encarne estos objetivos.

___

<h2 style="color: #4BA3C7; font-size: 28px;">Información del Laboratorio</h2>

Este laboratorio se ha creado en un sistema Windows 11 utilizando Visual Studio Code con una conexión SSH a una máquina virtual Ubuntu 22 alojada en Hyper-V, el componente de virtualización de Microsoft en Windows 11. Se recomienda que el sistema operativo Ubuntu y todos los componentes instalados en él se actualicen constantemente a sus últimas versiones, asegurando un entorno de desarrollo de vanguardia para la exploración.

Si necesitas una guía paso a paso para configurar este laboratorio, no dudes en contactarnos.

___

<h2 style="color: #4BA3C7; font-size: 28px;">Requirements</h2>

### Azure PowerShell Module (Az)

**Purpose:** Necesario para conectarse a los servicios de Azure, como Azure Key Vault, y realizar la autenticación.

**Usage:** Utiliza Connect-AzAccount para autenticarse con los servicios de Azure

**Installation:**
~~~ bash
sudo apt-get install -y powershell
~~~

Then within PowerShell
~~~ powershell
Install-Module -Name Az -AllowClobber -Scope CurrentUser
~~~

**Version Check:**
~~~ powershell
Get-InstalledModule -Name Az
~~~
Output
~~~ powershell
PS /home/user/learning/az/Set-Version-Limit-Function> Get-InstalledModule -Name Az            

Version              Name                                Repository           Description
-------              ----                                ----------           -----------
12.0.0               Az                                  PSGallery            Microsoft Azure PowerShell - Cmdlets to manage resources in Azure. This module is comp…
~~~

**Environment:** Local and Cloud

### .NET Core SDK
**Relevance:** Necesario para el desarrollo y prueba local de Funciones de Azure, asegurando la compatibilidad con Azure Functions V4.

**Installation:**
~~~ bash
wget https://packages.microsoft.com/config/ubuntu/$(lsb_release -rs)/packages-microsoft-prod.deb -O packages-microsoft-prod.deb
sudo dpkg -i packages-microsoft-prod.deb
sudo apt-get update
sudo apt-get install -y dotnet-sdk-6.0
~~~

Version Check:
~~~ bash
dotnet --version
~~~
Output
~~~ bash
6.0.128
~~~

**Environment:** Local (principalmente para pruebas y desarrollo)


### Azure Functions Environment
**Purpose:** Proporciona capacidades serverless compute para alojar y ejecutar scripts sin gestionar la infraestructura subyacente.

**Installation:**

~~~ bash
sudo npm install -g azure-functions-core-tools@4 --unsafe-perm true
sudo apt-get update
sudo apt-get install azure-functions-core-tools-4
~~~

**Version Check:**
~~~ bash
func --version
~~~ 
Output
~~~ bash
4.0.5801
~~~

**Environment:** Local and Cloud

### SharePoint Online Management Shell (PnP PowerShell)
**Purpose:** Proporciona cmdlets para gestionar SharePoint Online, como conectarse a SharePoint y gestionar bibliotecas de documentos.
Key Commands: Connect-PnPOnline, Set-PnPList.

**Installation:**
~~~ powershell
Install-Module -Name PnP.PowerShell -AllowClobber -Scope CurrentUser
~~~

**Version Check:**
~~~ powershell
Get-InstalledModule -Name PnP.PowerShell
~~~
Output
~~~ powershell
PS /home/user/learning/az/Set-Version-Limit-Function> Get-InstalledModule -Name PnP.PowerShell     

Version              Name                                Repository           Description
-------              ----                                ----------           -----------
2.4.0                PnP.PowerShell                      PSGallery            Microsoft 365 Patterns and Practices PowerShell Cmdlets
~~~

**Environment:** Local and Cloud

### Azure Key Vault

**Purpose:** Almacena de manera segura información sensible, como credenciales, utilizadas por scripts para autenticarse con servicios como SharePoint Online.

**Environment:** Cloud (principalmente para producciónn)
___


<h2 style="color: #4BA3C7; font-size: 28px;">Guía Paso a Paso para Crear una Función de PowerShell para Limitar la Versión de una Biblioteca de Documentos de SharePoint Online</h2>

### Paso 1: Inicializar un Nuevo Proyecto de Funciones de Azure
Abre una terminal en VS Code, navega al directorio (por ejemplo, /home/user/learning/az) donde deseas crear el proyecto (Set-Version-Limit-Function) y ejecuta::

 ~~~ bash
func init Set-Version-Limit-Function --worker-runtime powershell
 ~~~

Navega a Set-Version-Limit-Function (tu directorio de proyecto creado)

 ~~~ bash
cd Set-Version-Limit-Function 
 ~~~

### Paso 2: Crear una Nueva Función
Crea una nueva función (llamada Set-Version-Limit) dentro de la aplicación de funciones utilizando las Herramientas de Funciones de Azure.
~~~ bash
func new --name Set-Version-Limit --template "HTTP trigger" --authlevel anonymous
~~~

### Paso 3: Desarrollar la Función Set-Version-Limit
Abre run.ps1 (/home/user/learning/az/Set-Version-Limit-Function/Set-Version-Limit) en VS Code, y reemplaza su contenido con este código

~~~ powershell
using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)
Write-Host "PowerShell HTTP trigger function processed a request."

# Import PnP PowerShell Module
Import-Module PnP.PowerShell -ErrorAction Stop
Add-Type -AssemblyName "Microsoft.Extensions.Logging.Abstractions"

# Retrieve site URL from the query parameters
$siteUrl = $Request.Query.siteUrl
if (-not $siteUrl) {
    $response = @{
        status  = "Failed"
        message = "No site URL provided."
    }
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::BadRequest
        Body = ($response | ConvertTo-Json)
    })
    return
}

# Define your SharePoint credentials from environment variables
$username = $env:SP_USERNAME
$password = ConvertTo-SecureString -String $env:SP_PASSWORD -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($username, $password)

# Define version limit from environment variables
$versionLimit = 12

# Initialize success and failure lists
$successLibraries = @()
$failedLibraries = @()

# Define response structure
$response = @{
    status  = "Unknown"
    message = ""
}

# Connect to SharePoint Online
try {
    Connect-PnPOnline -Url $siteUrl -Credentials $cred
    Write-Output "Connected successfully to $siteUrl."
} catch {
    Write-Host "Failed to connect to ${siteUrl}: $_"
    $response.status = "Failed"
    $response.message = "Failed to connect to ${siteUrl}: $($_)"
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::InternalServerError
        Body = ($response | ConvertTo-Json)
    })
    exit
}

# Get all document libraries in the site
try {
    $libraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 }
    if ($libraries.Count -eq 0) {
        Write-Host "No document libraries found at ${siteUrl}."
        $response.status = "Failed"
        $response.message = "No document libraries found at ${siteUrl}."
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::NotFound
            Body = ($response | ConvertTo-Json)
        })
        exit
    }

    # Set the version limit for each document library
    foreach ($library in $libraries) {
        try {
            Set-PnPList -Identity $library -EnableVersioning $true -MajorVersions $versionLimit
            Write-Output "Version limit set to $versionLimit for $($library.Title) at $siteUrl."
            $successLibraries += $library.Title
        } catch {
            Write-Host "Error setting version limits for $($library.Title) at ${siteUrl} : $_"
            $failedLibraries += $library.Title
        }
    }

    $response.status = "Success"
    $response.message = "$($successLibraries.Count) libraries had their version limit set to $versionLimit; $($failedLibraries.Count) libraries failed to set limit."
    $response.successLibraries = $successLibraries
    $response.failedLibraries = $failedLibraries
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body = ($response | ConvertTo-Json)
    })
} catch {
    Write-Host "Error managing document libraries at ${siteUrl}: $_"
    $response.status = "Failed"
    $response.message = "Error managing document libraries at ${siteUrl} : $($_)"
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::InternalServerError
        Body = ($response | ConvertTo-Json)
    })
}

~~~

### Step 4: Run and Test Locally
Test Set-Version-Limit function locally:

~~~ bash
func start
~~~

~~~ bash
Azure Functions Core Tools
Core Tools Version:       4.0.5801 Commit hash: N/A +5ac2f09758b98257e728dd1b5576ce5ea9ef68ff (64-bit)
Function Runtime Version: 4.34.1.22669


Functions:

        Set-Version-Limit: [GET,POST] http://localhost:7071/api/Set-Version-Limit

For detailed output, run func with --verbose flag.
[2024-05-29T08:04:54.025Z] Worker process started and initialized.
~~~

Open a new terminal in VS Code and test the function 

~~~ bash
curl -X POST "http://localhost:7071/api/version-limit?siteUrl=https://softwareone.sharepoint.com/sites/TestMLV"
~~~

~~~ bash
{
  "successLibraries": [
    "Documents",
    "Form Templates",
    "Site Assets",
    "VersionLimitLibrary"
  ],
  "failedLibraries": [
    "Style Library"
  ],
  "status": "Success",
  "message": "4 libraries had their version limit set to 12; 1 libraries failed to set limit."
}%  
~~~

