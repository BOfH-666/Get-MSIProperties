<#
.SYNOPSIS
    Gathering information about a given *.msi file.

.DESCRIPTION
    This command returns information about a give *.msi file like product code, manufacturer, version and name.

.PARAMETER msiPath
    Full path of the *.msi file.

.EXAMPLE
    PS C:\> Get-MSIProperties C:\temp\example.msi

    This command returns the product code, name, manufacturer and version of the given *.msi file"

.NOTES
    Script-Version: 5.0.0
    Author: Olaf.Soyk
    Date: 28.01.2015 - 10:55
    
    ################################
    Helperfunction for UnInstall-MSI
    ################################
    
    code borrowed from http://blog.joefield.co.uk/?p=3
    --------------------------------------------------
    Fetches ProductProperties from MSI file 
#>
function Get-MSIProperties {

    param($msiPath)
    if ($null -eq $msiPath) {
        'Expects full path to MSI file'
        return;
    }

    $installer = New-Object -comObject WindowsInstaller.Installer
    $database = $installer.GetType().InvokeMember("OpenDatabase", [System.Reflection.BindingFlags]::InvokeMethod, $null, $installer, ($msiPath, 0))

    #region ProductVersion
    $view = $database.GetType().InvokeMember("OpenView", [System.Reflection.BindingFlags]::InvokeMethod, $null, $database, "SELECT `Value` FROM `Property` WHERE `Property` = 'ProductVersion'")
    $view.GetType().InvokeMember("Execute", [System.Reflection.BindingFlags]::InvokeMethod, $null, $view, $null)
    $record = $view.GetType().InvokeMember("Fetch", [System.Reflection.BindingFlags]::InvokeMethod, $null, $view, $null)
    $ProductVersion = $record.GetType().InvokeMember("StringData", [System.Reflection.BindingFlags]::GetProperty, $null, $record, 1)
    #endregion

    #region ProductName
    $view = $database.GetType().InvokeMember("OpenView", [System.Reflection.BindingFlags]::InvokeMethod, $null, $database, "SELECT `Value` FROM `Property` WHERE `Property` = 'ProductName'")
    $view.GetType().InvokeMember("Execute", [System.Reflection.BindingFlags]::InvokeMethod, $null, $view, $null)
    $record = $view.GetType().InvokeMember("Fetch", [System.Reflection.BindingFlags]::InvokeMethod, $null, $view, $null)
    $ProductName = $record.GetType().InvokeMember("StringData", [System.Reflection.BindingFlags]::GetProperty, $null, $record, 1)
    #endregion

    #region Manufacturer
    $view = $database.GetType().InvokeMember("OpenView", [System.Reflection.BindingFlags]::InvokeMethod, $null, $database, "SELECT `Value` FROM `Property` WHERE `Property` = 'Manufacturer'")
    $view.GetType().InvokeMember("Execute", [System.Reflection.BindingFlags]::InvokeMethod, $null, $view, $null)
    $record = $view.GetType().InvokeMember("Fetch", [System.Reflection.BindingFlags]::InvokeMethod, $null, $view, $null)
    $Manufacturer = $record.GetType().InvokeMember("StringData", [System.Reflection.BindingFlags]::GetProperty, $null, $record, 1)
    #endregion

    #region ProductCode
    $view = $database.GetType().InvokeMember("OpenView", [System.Reflection.BindingFlags]::InvokeMethod, $null, $database, "SELECT `Value` FROM `Property` WHERE `Property` = 'ProductCode'")
    $view.GetType().InvokeMember("Execute", [System.Reflection.BindingFlags]::InvokeMethod, $null, $view, $null)
    $record = $view.GetType().InvokeMember("Fetch", [System.Reflection.BindingFlags]::InvokeMethod, $null, $view, $null)
    $ProductCode = $record.GetType().InvokeMember("StringData", [System.Reflection.BindingFlags]::GetProperty, $null, $record, 1)
    #endregion
    
    [PSCustomObject]@{
        Manufacturer   = $Manufacturer
        ProductName    = $ProductName
        ProductVersion = $ProductVersion
        ProductCode    = $ProductCode
    }
}



<#
.Synopsis
   Get product- & version-specific information from MSI file
.DESCRIPTION
   Use the MsiInstaller.Installer ComObject to enumerate MSI database specific information
   There are only 5 properties for MSI's that are mandatory.  (According to https://msdn.microsoft.com/en-us/library/windows/desktop/aa370905(v=vs.85).aspx )
   These are:
       ProductCode     - A unique identifier for a specific product release.
       Manufacturer    - Name of the application manufacturer.
       ProductName     - Human readable name of an application.
       ProductVersion  - String format of the product version as a numeric value.
       ProductLanguage - Numeric language identifier (LANGID) for the database.
 
   By default all of these are returned.  This can be modified by using the [-Property] Parameter.
 
.EXAMPLE
PS C:\> Get-MsiInformation -Path "$env:Temp\Installer.msi"
 
Path            : C:\Users\username\AppData\Local\Temp\Installer.msi
ProductCode     : {75BDEFC7-6E84-55FF-C326-CE14E3C889EC}
ProductVersion  : 1.9.492.0
ProductName     : Installer v1.9.0
Manufacturer    : My Company, Inc.
ProductLanguage : 1033
 
This example takes the path as a parameter and returns all fields
 
.EXAMPLE
Get-ChildItem -Path "$env:Temp\1.0.0" -Recurse -File -Include "*.msi" | Get-MsiInformation -Property ProductVersion
 
This example takes multiple paths from a Get-ChildItem query and extracts the information
     
.INPUTS
   [System.IO.File[]] - Single or Array of Paths to interrogate
 
.OUTPUTS
   [System.Management.Automation.PSCustomObject[]] - Contains the Msi File Object and Associated Properties
 
.LINK
   http://blog.kmsigma.com/
 
.LINK
   https://msdn.microsoft.com/en-us/library/windows/desktop/aa370905(v=vs.85).aspx
 
.LINK
   <blockquote class="wp-embedded-content" data-secret="04FLTT0eH3"><a href="http://www.scconfigmgr.com/2014/08/22/how-to-get-msi-file-information-with-powershell/">How to get MSI file information with PowerShell</a></blockquote><iframe class="wp-embedded-content" sandbox="allow-scripts" security="restricted" style="position: absolute; clip: rect(1px, 1px, 1px, 1px);" src="http://www.scconfigmgr.com/2014/08/22/how-to-get-msi-file-information-with-powershell/embed/#?secret=04FLTT0eH3" data-secret="04FLTT0eH3" width="600" height="338" title="&#8220;How to get MSI file information with PowerShell&#8221; &#8212; System Center ConfigMgr" frameborder="0" marginwidth="0" marginheight="0" scrolling="no"></iframe>
 
.NOTES
   Heavily Infuenced by http://www.scconfigmgr.com/2014/08/22/how-to-get-msi-file-information-with-powershell/
 
.FUNCTIONALITY
   Uses ComObjects to Enumerate specific fields in the MSI database
#>
function Get-MsiInformation {
    [CmdletBinding(SupportsShouldProcess = $true, 
        PositionalBinding = $false,
        ConfirmImpact = 'Medium')]
    [Alias('gmsi')]
    #[OutputType([System.Management.Automation.PSCustomObject[]])]
    Param(
        [parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = 'Provide the path to an MSI')]
        [ValidateNotNullOrEmpty()]
        [System.IO.FileInfo[]]$Path,
  
        [parameter(Mandatory = $false)]
        [ValidateSet( 'ProductCode', 'Manufacturer', 'ProductName', 'ProductVersion', 'ProductLanguage' )]
        [string[]]$Property = ( 'ProductCode', 'Manufacturer', 'ProductName', 'ProductVersion', 'ProductLanguage' )
    )
 
    Begin {
        # Do nothing for prep
    }
    Process {
         
        ForEach ( $P in $Path ) {
            if ($pscmdlet.ShouldProcess($P, 'Get MSI Properties')) {            
                try {
                    Write-Verbose -Message "Resolving file information for $P"
                    $MsiFile = Get-Item -Path $P
                    Write-Verbose -Message "Executing on $P"
                     
                    # Read property from MSI database
                    $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
                    $MSIDatabase = $WindowsInstaller.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $WindowsInstaller, @($MsiFile.FullName, 0))
                     
                    # Build hashtable for retruned objects properties
                    $PSObjectPropHash = [ordered]@{File = $MsiFile.FullName }
                    ForEach ( $Prop in $Property ) {
                        Write-Verbose -Message "Enumerating Property: $Prop"
                        $Query = "SELECT Value FROM Property WHERE Property = '$( $Prop )'"
                        $View = $MSIDatabase.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $MSIDatabase, ($Query))
                        $View.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $View, $null)
                        $Record = $View.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $View, $null)
                        $Value = $Record.GetType().InvokeMember('StringData', 'GetProperty', $null, $Record, 1)
  
                        # Return the value to the Property Hash
                        $PSObjectPropHash.Add($Prop, $Value)
 
                    }
                     
                    # Build the Object to Return
                    $Object = @( New-Object -TypeName PSObject -Property $PSObjectPropHash )
                     
                    # Commit database and close view
                    $MSIDatabase.GetType().InvokeMember('Commit', 'InvokeMethod', $null, $MSIDatabase, $null)
                    $View.GetType().InvokeMember('Close', 'InvokeMethod', $null, $View, $null)           
                    $MSIDatabase = $null
                    $View = $null
                }
                catch {
                    Write-Error -Message $_.Exception.Message
                }
                finally {
                    Write-Output -InputObject @( $Object )
                }
            } # End of ShouldProcess If
        } # End For $P in $Path Loop
 
    }
    End {
        # Run garbage collection and release ComObject
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null
        [System.GC]::Collect()
    }
}
<#
End of Function
#>