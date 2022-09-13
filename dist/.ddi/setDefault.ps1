$namespace = "root\CIMV2\NV"    # Namespace of NVIDIA WMI provider
$class = "ProfileManager"       # class to be queried
if($computer -eq $null)         # if not set globally
{
    $computer  = "localhost"    # substitute with remote machine names or IP of the machine
}

$profileType3DApplication = 0   # 3D Application profile type
$profileType3DGlobal      = 1   # 3D Global profile type

function Set-3DGlobalPreset
{
    Param
    (
        [System.Management.ManagementBaseObject]$profileManagerInstance,
        [string]$global3DPreset
    )
    
    $method = "setCurrentProfile3D"

    "Current Global 3D Preset: "+$profileManagerInstance.currentProfile3D
    "Calling $class.$method()"
    $result = Invoke-WmiMethod -Path $profileManagerInstance.__PATH -Name $method -ArgumentList $global3DPreset, $null
    if($result.ReturnValue -eq $true)
    {
        $ret = "succeeded"
    }
    else
    {
        $ret = "failed"
    }
    "$method call for "+$global3DPreset+": $ret"
    "---"
}

function Restore-Defaults3D
{
    Param
    (
        [System.Management.ManagementBaseObject]$profileManagerInstance
    )
    $method = "restoreDefaults3D"
    "Calling ProfileManager.$method()"
    $result = Invoke-WmiMethod -Path $profileManagerInstance.__PATH -Name $method
    if($result.ReturnValue -eq $true)
    {
        $ret = "succeeded"
    }
    else
    {
        $ret = "failed"
    }
    "$method call: $ret" 
    "---"
}



$profileManager = Get-WmiObject -class $class -computername $computer -namespace $namespace
if($profileManager -eq $Null)
{
    "Profile manager instance unavailable"
    return
}

#Detect-ApplicationProfiles

#$global3DPreset = "Workstation App - Advanced Streaming"
#Set-3DGlobalPreset ([ref]$profileManager) $global3DPreset

Restore-Defaults3D([ref]$profileManager)
