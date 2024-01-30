<!-- Back to top link -->
<a name="readme-top"></a>

<!-- NAME -->
# Manage Teams Updates
**Script to manage MS Teams classic updates** 

<!-- ABSTRACT -->
## ABSTRACT 
This script manages Microsoft Teams Classic updates by renaming update.exe and squirrel.exe, and updating the shortcut.

<!-- ABOUT THE PROJECT -->
## DESCRIPTION
The script renames the update.exe and squirrel.exe files in the Microsoft Teams directory
and updates the shortcut target path to point to the renamed update.exe.
    
 <p align="right">(<a href="#readme-top">back to top</a>)</p>
 
<!-- Getting Started -->
## QUICKSTART

### Prerequisites
Get information about
* Windows version ; Version must be 10 or alter
    * _Cmdlet_
    ```powershell
    Get-ComputerInfo
    ```
    * _Environment Class_
    ```powershell
    [Environment]::OSVersion
    ```
* Powershell version ; Version must be 5.1 or later
    * _Cmdlet_
    ```powershell
    Get-Host
    ```
    * _Automatic Variable_
    ```powershll
    $PSVersionTable
    ```
### Installation

Here a some options to install the script and run it

1. Set up a Scheduled Task for Windows
2. Set up a Configuration Item for SCCM
3. Set up a Logon Script for GPO

 <p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- ROADMAP -->
## ROADMAP

| Windows | Linux | MacOS|
| :----: | :---: | :--: |
| In progress | To be decided | To be decided |

- [ ] Windows
    - [x] Script
    - [ ] Cmdlet
   

<p align="right">(<a href="#readme-top">back to top</a>)</p>


<!-- LICENSE -->
## LICENSE

Distributed under the  Unlicense license. See `LICENSE` for more information.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- ACKNOWLEDGMENTS -->
## LINKS
* [Microsoft PowerShell Documentation](https://learn.microsoft.com/en-us/powershell/)
 
<p align="right">(<a href="#readme-top">back to top</a>)</p>
 


<!-- CONTACT -->
## CONTACT

:e-mail: 

<p align="right">(<a href="#readme-top">back to top</a>)</p>
