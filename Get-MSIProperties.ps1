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
    Author: Olaf.Soyk@Computacenter.com
    Date: 28.01.2015 - 10:55
    
    ################################
    Helperfunction for UnInstall-MSI
    ################################
    
    code borrowed from http://blog.joefield.co.uk/?p=3
    --------------------------------------------------
    Fetches ProductProperties from MSI file and provides a global variable with it
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
    
    $GLOBAL:MSIProperties = New-Object PSObject -Property @{
        Manufacturer   = $Manufacturer
        ProductName    = $ProductName
        ProductVersion = $ProductVersion
        ProductCode    = $ProductCode
    }
    $GLOBAL:MSIProperties
}
