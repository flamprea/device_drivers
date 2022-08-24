#!/bin/nsh
# ©2007 Frank Lamprea, BladeLogic.
#
# Extended object that will show device driver information on Windows Systems
#
# This should be added as a CENTRALLY EXECUTED Extended Object with the CSV grammar.
# The hostname should be passed as the first paramater ??TARGET.HOST??
#
# Based on script by Mark Jeffrey, BladeLogic.
#

# Define a local temporary directory
LOCAL_TEMP_DIR="/tmp/vbs_$$"
LOCAL_TEMP_FILE="extobj_$$.vbs"

# Create it, if necessary
mkdir -p "$LOCAL_TEMP_DIR"

FULL_PATH_THIS_FILE="$0"
THIS_FILE=`basename "$FULL_PATH_THIS_FILE"`

TARGET_HOST="$1"

if test "$TARGET_HOST" = ""; then
	echo "Incorrect usage."
	echo
	echo "Usage:  $THIS_FILE <Target_Hostname>"
	exit 255
fi

# Find location of embedded script
LENGTH_OF_MAIN_SCRIPT=`cat "$FULL_PATH_THIS_FILE" | grep -n '### END OF MAIN SCRIPT ###' | grep -v grep | cut -d ':' -f1`

# Increment this by one to move past the line identifier
LENGTH_OF_MAIN_SCRIPT=`expr $LENGTH_OF_MAIN_SCRIPT + 1`

# Extract VBScript
tail +$LENGTH_OF_MAIN_SCRIPT "$FULL_PATH_THIS_FILE" > "$LOCAL_TEMP_DIR/$LOCAL_TEMP_FILE"

# Define the CSCript command
CSCRIPT_CMD="cscript /nologo"

# Define a remote temporary directory and create it, if necessary
REMOTE_TEMP_DIR="//$TARGET_HOST/tmp/vbs"
mkdir -p "$REMOTE_TEMP_DIR"

# Copy over our extracted VBScript
cp "$LOCAL_TEMP_DIR/$LOCAL_TEMP_FILE" "$REMOTE_TEMP_DIR"

# Actually CD into the remote temporary directory, to avoid
# Windows pathname problems
cd "$REMOTE_TEMP_DIR"

# Execute the VBScript
nexec -e $CSCRIPT_CMD $LOCAL_TEMP_FILE

# Clean up temporary files
rm -fr "$REMOTE_TEMP_DIR/$LOCAL_TEMP_FILE"
rm -fr "//@/$LOCAL_TEMP_DIR/$LOCAL_TEMP_FILE"
rm -fr "//@/$LOCAL_TEMP_DIR"

exit 0

### END OF MAIN SCRIPT ###
Dim args
Dim objPNPSignedDrivers, SignedDriver, strTextOutput
Dim filter
Dim num


set args = WScript.Arguments
num = args.Count

ComputerName = "localhost"
if num >= 1 then
	filter = args.Item(0)
else
	filter = ""
end if

Set objPNPSignedDrivers = GetObject("winmgmts:\\" & ComputerName).InstancesOf("Win32_PnPSignedDriver")

'Wscript.Echo "Device Class,Description,Device Name,Device ID,INF File Name,Driver Provider Name,Driver Version,INF Manufacturer"
On Error Resume Next

For Each SignedDriver In objPNPSignedDrivers
	On Error GoTo 0
	if filter = "" then
		strTextOutput = SignedDriver.DeviceID & "," & SignedDriver.DeviceClass & ",""" & SignedDriver.Description & """,""" _
		& SignedDriver.DeviceName & """," &  SignedDriver.InfName & ",""" _
		& SignedDriver.DriverProviderName & """," & SignedDriver.DriverVersion & "," _
		& """" & SignedDriver.Manufacturer & """"
		Wscript.echo strTextOutput
	elseif filter = SignedDriver.DeviceClass then
		strTextOutput = SignedDriver.DeviceID & "," & SignedDriver.DeviceClass & ",""" & SignedDriver.Description & """,""" _
		& SignedDriver.DeviceName & """," &  SignedDriver.InfName & ",""" _
		& SignedDriver.DriverProviderName & """," & SignedDriver.DriverVersion & "," _
		& """" & SignedDriver.Manufacturer & """"
		Wscript.echo strTextOutput
	end if
Next
