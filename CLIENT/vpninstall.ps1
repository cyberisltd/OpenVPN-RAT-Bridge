# This script must be run as admin. It installs openvpn silently, creates a bridge using devcon, and adds all physical network adapters. It's loosely based around the following link, with a few more hacks  thttp://stackoverflow.com/questions/17588957/programmatically-create-destroy-network-bridges-with-net-on-windows-7

# This script is designed for Win 7 x64, but you would be able to get it working on other OS versions without too much difficulty.

if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
   Write-Host "Admin privileges are required to run this script" -f red 
   exit 1
}

Write-Host "This PowerShell command is running under the current users context, $env:userdomain\$env:username" -f green  

# Check whether there is already a bridge (miniport) present. If there is, exit, as it's a more complex set up than we expect!
$query = "SELECT * FROM Win32_NetworkAdapter WHERE PNPDeviceID IS NOT NULL and name like '%bridge%'"
if (get-wmiobject -query $query) {
	# A bridge exists
	Write-Host "We've found a bridge, so not proceeding with install".
	exit 2
}

# Do check for openvpn install
if (Test-Path "c:\program files\openvpn\config\") {
	Write-Host "It appears there is an openvpn conf directory, so not proceeding with install."
	exit 3
}

# Make sure your toolset includes devcon (it's not inbuilt, it's part of the WDK)
# Note, you need to create a signed CAT file, and import the cert to both the ROOT store and Trusted publisher store to get it working (instructions followed from http://www.richud.com/wiki/Windows_7_Install_Unsigned_Drivers_CAT_fix)
Write-Host "Adding certificates for the bridge driver..."
certutil.exe -addstore "TrustedPublisher" .\tools\driver.cer
certutil.exe -addstore "ROOT" .\tools\driver.cer

# Now to the actual driver install - note, this doesn't add the adapters to the bridge.
Write-Host "Creating the bridge..."
.\tools\devconx64.exe install ".\tools\bridge_install_win7_x64\netbrdgm.inf" ms_bridgemp

# The next command installs the openvpn certificate so you don't get a GUI prompt when adding the TAP interface
Write-Host "Adding certificates for the TAP driver..."
certutil.exe -addstore "TrustedPublisher" .\tools\openvpninstaller.p7b

# Install openvpn silently
.\tools\openvpn-install-2.3.5-I602-x86_64.exe /S

# Wait for the install to finish, 30 seconds should be plenty.
Start-Sleep -s 30

# Copy the certs across to the right location
copy .\config\* "c:\program files\openvpn\config\"

# Now add the adapters on the host to the newly created bridge, and set to compatibility mode
# This script assumes two adapters, one containing the name 'Network Connection' and one containing the name 'V9', which is the latest version of the openvpn tap adapter. 
$lanadapter = "SELECT PNPDeviceID FROM Win32_NetworkAdapter WHERE PNPDeviceID IS NOT NULL and name like '%Network Connection%'"
$tapadapter = "SELECT PNPDeviceID FROM Win32_NetworkAdapter WHERE PNPDeviceID IS NOT NULL and name like '%V9%'"

# Release the IP addresses currently assigned - you're going to lose connectivity now...hope it comes up!
write-host "Releasing IP addresses as you don't want local adapters retaining IP addresses once the bridge is built. Hopefully you'll regain connectivity in a few seconds!!"
ipconfig /release

# Now bind the adapters to the bridge
# Bind bridge can be found here https://github.com/OurGrid/OurVirt/tree/master/tools/win32/bindbridge
$landev = get-wmiobject -query $lanadapter | Select-Object PNPDeviceID -ExpandProperty PNPDeviceID | out-string 
if ($landev) {
	.\tools\bindbridge.exe ms_bridge $landev.Trim() bind
}
else {
	Write-Host "Could not identify one and only one physical adapter...exiting"
	exit 4
}
$tapdev = get-wmiobject -query $tapadapter | Select-Object PNPDeviceID -ExpandProperty PNPDeviceID | out-string  
if ($tapdev) {
	.\tools\bindbridge.exe ms_bridge $tapdev.Trim() bind
}
else {
	Write-Host "Could not identify one and only one TAP adapter...exiting"
	exit 5
}

# Set all interfaces to compat mode to avoid issues with bridging. This assumes there were no issues building the bridge!
# This is especially important if running on a virtual machine test lab
netsh bridge set adapter 1 forcecompatmode=enable
netsh bridge set adapter 2 forcecompatmode=enable

# Disable and re-enable the bridge (this seems necessary to get it working)
netsh interface set interface "Network Bridge" DISABLED
netsh interface set interface "Network Bridge" ENABLED

# Configure the openvpn service to start automatically, and finally, start the service
Set-Service openvpnservice -startuptype "automatic"
Start-Service openvpnservice
