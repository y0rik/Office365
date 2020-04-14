<#
.SYNOPSIS
    Short description of the script
.DESCRIPTION
    Longer description of the script
.PARAMETER Param1
    Describe Param1 attribute
    !More PARAMETER sections may be added if there is more than 1 parameter
.INPUTS
    Input description of the script
.OUTPUTS
    Output description of the script
.NOTES
  Version:  Current version
  Author:   Author(s)
  Created:  Draft version date
  Updated:  Last update date if you prefer to keep it directly in code
  Other:    Other information if needed
.EXAMPLE
    Put the script typical running scenarios under this section:
    - command example
    - explanation of the command

    !More EXAMPLE sections may be added if necessary
    !Find below EXAMPLE section example

.EXAMPLE
    .\Run-Script.ps1 -RootPath "c:\root" -Recursively -Verbose
    Run the script to start from 'c:\root' path, execute recursively with extended output
#>

# global script settings
# use 'SupportsShouldProcess = $true' if you need support for -WhatIf, -Verbose and -Confirm
# use 'ConfirmImpact = 'High'' if you want to make sure 'Confirm:$true' under any circumstances (ignores $ConfirmPreference)
# use 'PositionalBinding = $false' to turn off parameters' positioning
# more information about CmdletBinding: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions_cmdletbindingattribute
<#
[CmdletBinding(ConfirmImpact=<String>,
DefaultParameterSetName=<String>,
HelpURI=<URI>,
SupportsPaging=<Boolean>,
SupportsShouldProcess=<Boolean>,
PositionalBinding=<Boolean>)]
#>

# script parameters
# uncomment below parameters section if needed, more parameters' customization is available
# more about parameters: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions_advanced_parameters
<#
param (
    [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$Param1
)
#>

# uncomment if powershell version verification is needed, put all powershell requirements below
# more about requirements: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_requires
<#
#Requires -Version PSVersion
#>

# External binaries
# add a .Net assembly
#Add-Type -AssemblyName
# add a Com object
#New-Object -ComObject


# put your global variables below this tag
###GLOBAL DECLARATIONS###

###FUNCTIONS###
# put your functions below this tag

###BODY###
# put main script execution code below this tag

# script self-identification, uncomment if needed
#$scriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
#$scriptName = $MyInvocation.MyCommand.Name


