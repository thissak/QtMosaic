# I’m trying to script setting the Power Management setting in the Base Profile and the script works on my own computer every time I run it. 
# But, when I take the script over to a different computer the Power Management setting won’t take the value.
# Boiling the script down to the essential command I have the following with the result.

wmic /namespace:nv Path Profile where ID=1975410716 Call setValueById settingId=274197361 value=1

Produces:
Executing (\computername\root\cimv2\nv:Profile.id=1975410716)->setValueById()
Method execution successful.
Out Parameters:
instance of __PARAMETERS
{
    ReturnValue = FALSE;
};

# When doing the same on my own computer the ReturnValue is True because it succeeds. 
# So, why does it work on one computer but not another? Both are nVidia Quadro GPUs both have the NVWMI components installed. 
# I’ll put the full real script below.

Set Environment Variables
$value = 1 # Valid options: 0=Adaptive, 1-Prefer Maximum, 2=nVidia Controlled, 3=Prefer Consistent
$namespace = “root\CIMV2\NV” # Namespace of NVIDIA WMI provider
$Profile3DGlobalName = “Base Profile” # name of Profile to be queried
$settingName = “Power management mode” #name of setting in above profile to be set
$settingTableID = 0 # ID# of 3D profile setting name table
$method = “setValueById”
if($computer -eq $null) # if not set globally
{
    $computer = “localhost” # substitute with remote machine names or IP of the machine
}

Set Base Profile as active profile
$profileManager = Get-WmiObject -class “ProfileManager” -computername $computer -namespace $namespace
if($profileManager -eq $Null)
{
    “nVidia Profile Manager unavailable”
    return
}
$result = Invoke-WmiMethod -Path $profileManager.__PATH -Name setCurrentProfile3D -ArgumentList $Profile3DGlobalName, $null
if(!$result.ReturnValue)
{
    “Unable to set current 3D Profile to $Profile3DGlobalName”
    return
}

Retrieve profile object for $Profile3DGlobalName
$Profile3DGlobalObj = Get-WmiObject -class “Profile” -computername $computer -namespace $namespace | Where-Object { $_.name -eq $Profile3DGlobalName }
if($Profile3DGlobalObj -eq $Null)
{
    “No Profile named $Profile3DGlobalName found.”
    return
}

retrieve Setting table corresponding to 3D Profiles
$3DSettingTable = Get-WmiObject -class “SettingTable” -computername $computer -namespace $namespace | Where-Object {$_.type -eq $settingTableID}
if($3DSettingTable -eq $null)
{
    “WARNING : cannot retrieve a SettingTable instance for 3D profile settings”
    return
}
$settingID = $3DSettingTable.InvokeMethod(“getIdFromName”, $settingName)

Apply the Setting
$result = Invoke-WmiMethod -Path $Profile3DGlobalObj.__PATH -Name $method -ArgumentList $settingID,$value
if($result.ReturnValue)
{
    Write-Host “$Profile3DGlobalObj.__PATH.$method(settingId = $settingID, value = $value) returns $($result.ReturnValue)”
    “Setting $settingName to Prefer Maximum succeeded.”
    return
}
“Setting $settingName to Prefer Maximum failed.”
return




