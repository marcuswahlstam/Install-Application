
# RegKey for app
$installAppDisplayName = "Check_MK Agent","Check MK Agent"
$installAppVersion = "1.6.0p3"

# Path to installer
$appInstallerPath = "\\corporate\osk\Data\GEMENSAM\DATA\Applikationer\CheckMKClient\1.6.0\check_mk_agent.msi"

# Installer command
$appInstallerArguments = "/i","$appInstallerPath","/qn"

# Update only, only run if previous version is installed ($true or $false)
$UpdateOnly = $true

# DB-information
$DBServer = "SE-OSK-SQL100.corporate.saft.org"
$DBName = "SaftScriptLog"
$DBTable = "UpdateToCheckMK160"

$regKey = "HKLM:\SOFTWARE\Saft Batteries\Scripts\UpdateToCheckMK160"
#$regName = "UpdateToCheckMK160"
#$regValue = "Success" # Change this if you want the script to run again on all servers

###
### Do not edit anything below
###

function Find-InstalledApp
{
    foreach ($app in $installAppDisplayName)
    {
        #64bit
        $foundApp = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | where {$_.DisplayName -like "*$app*"}

        if (!($foundApp))
        {
            $foundApp = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | where {$_.DisplayName -like "*$app*"}
        }

        if ($foundApp){break}
    }
    return $foundApp
}
<#
function Invoke-AppInstall
{
    # Command to install app
    Invoke-Expression $appInstallerCommand
}
#>

function Add-ToReg
{
    param([string] $regValue,[string] $regName,[string]$Type)
    if (Test-Path $regKey)
    {
        if ((Get-Item $regKey -EA Ignore).Property -contains $regName)
        {
            Set-ItemProperty -Path $regKey -Name $regName -Value $regValue -Force | Out-Null
        }
        else
        {
            New-ItemProperty $regKey -Name $regName -Value $regValue -PropertyType $Type -Force | Out-Null
        }
    }
    else
    {
        New-Item -Path $regKey -Force | Out-Null
        New-ItemProperty $regKey -Name $regName -Value $regValue -PropertyType $Type -Force | Out-Null
    }
}

function Add-ToDBLog
{
    param([string] $LogText,[string] $Result)
    $timeNow = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $sqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $sqlConnection.ConnectionString = "Server=$DBServer;Database=$DBName;Integrated Security=True;"
    $sqlConnection.Open()

    $query= "begin tran
            if exists (SELECT * FROM $DBTable WITH (updlock,serializable) WHERE Computername='"+$env:COMPUTERNAME+"')
            begin
                UPDATE $DBTable SET Computername='"+$env:COMPUTERNAME+"', ScriptOutput='"+$LogText+"', ScriptSuccessful='"+$result+"', DateRun='"+$timeNow+"'
                WHERE Computername = '"+$env:COMPUTERNAME+"'
            end
            else
            begin
                INSERT INTO $DBTable (Computername, ScriptOutput, ScriptSuccessful, DateRun)
                VALUES ('"+$env:COMPUTERNAME+"', '"+$LogText+"', '"+$result+"', '"+$timeNow+"')
            end
            commit tran"

    $sqlCommand = New-Object System.Data.SqlClient.SqlCommand($query,$sqlConnection)
    $sqlDS = New-Object System.Data.DataSet
    $sqlDA = New-Object System.Data.SqlClient.SqlDataAdapter($sqlCommand)
    [void]$sqlDA.Fill($sqlDS)

    $sqlConnection.Close()
}

###
### SCRIPT BEGIN
###

if ((Get-ItemProperty $regKey -ErrorAction Ignore).Status -eq "Success")
{
    exit 0
}
elseif ((Get-ItemProperty $regKey -ErrorAction Ignore).Retries -le "2" -and (Get-ItemProperty $regKey -ErrorAction Ignore).Status -eq "Failure")
{
    $retries = (Get-ItemProperty $regKey -ErrorAction Ignore).Retries
    $retries = $retries + 1
    if ($retries -eq $null){$retries = 1}
    Add-ToReg -regName "Retries" -regValue $retries -Type "DWORD"
}
elseif ((Get-ItemProperty $regKey -ErrorAction Ignore).Retries -ge "3" -and (Get-ItemProperty $regKey -ErrorAction Ignore).Status -eq "Failure")
{
    Add-ToDBLog -LogText "Max retries is reached. Will not try anymore." -Result "Failure"
    exit 0
}


$installedAppInformation = Find-InstalledApp

if ($installedAppInformation)
{
    $installedAppVersion = $installedAppInformation.DisplayVersion
    $installedAppDisplayName = $installedAppInformation.DisplayName
    $installedAppUninstallString = ($installedAppInformation.UninstallString -split 'msiexec.exe ')[1] + " /qn"

    if ($installedAppVersion -ne "$installAppVersion")
    {
        # App installed but not correct version
        try{
            #Invoke-AppInstall
            Start-Process -FilePath "msiexec.exe" -ArgumentList "$appInstallerArguments" -Wait -Passthru
            Start-Process -FilePath "msiexec.exe" -ArgumentList "$installedAppUninstallString" -Wait -PassThru
            Add-ToDBLog -LogText "$installedAppDisplayName $installedAppVersion was updated to $installAppVersion" -Result "Success"
            Add-ToReg -regName "Status" -regValue "Success" -Type "String"
        }
        catch{
            Add-ToDBLog -LogText "Update to $installAppDisplayName $installAppVersion failed" -Result "Failure"
            Add-ToReg -regName "Status" -regValue "Failure" -Type "String"
            #Add-ToReg -regName "InstallationRetries" -regValue 
        }
        
    }
    else {
        # App is installed and correct version
        Add-ToDBLog -LogText "$installAppDisplayName $installAppVersion already installed" -Result "Success"
        Add-ToReg -regName "Status" -regValue "Success" -Type "String"
    }

}
else {
    # App not installed
    if ($UpdateOnly -eq $false)
    {
        try{
            Start-Process -FilePath "msiexec.exe" -ArgumentList "$appInstallerArguments" -Wait -Passthru
            Add-ToDBLog -LogText "$installAppDisplayName installed (previous version did not exist)" -Result "Success"
            Add-ToReg -regName "Status" -regValue "Success" -Type "String"
        }
        catch{
            Add-ToDBLog -LogText "New installation of $installAppDisplayName $installAppVersion failed" -Result "Failure"
            Add-ToReg -regName "Status" -regValue "Failure" -Type "String"
        }
    }
}