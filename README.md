<h2 style="color: Black; font-size: 28px;">Infrastructure and Tools Setup for Azure Functions Development
</h2>

Welcome to our guide on setting up and using Azure Functions! This material is specifically crafted for those who are new to cloud computing and serverless architectures. Whether you're a student stepping into the world of technology or a hobbyist looking to expand your skills, this guide is intended to provide you with a clear, step-by-step introduction to developing with Azure Functions. Please note, the instructions and setups described here are designed for academic and learning purposes—they are simplified to facilitate understanding and are not meant for deployment in production environments. Our goal is to help you grasp the fundamentals without getting bogged down in the complexities often faced by experts and professionals in the field. So, let’s get started on this exciting journey into the world of serverless applications with Azure Functions!

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

