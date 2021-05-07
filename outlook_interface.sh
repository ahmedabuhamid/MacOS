#!/bin/bash
: <<ABOUT_THIS_SCRIPT
	Written by: Ahmed Abouhamed
	eng.abuhamed@gmail.com
	Last updated: May 6, 2021
	Originally posted: May 6, 2021
	Purpose: Use this script as part of an extension attribute in Jamf
	to report if the current loggedin user is using the new outlook
	interface or using the classic interface ?
	LICENSE: Except where otherwise noted, this work is licensed under
	http://creativecommons.org/licenses/by/4.0/
	
	INSTRUCTIONS

	1) Log in to the Jamf Pro server.
	2) Navigate to JSS Settings > Computer Management > Extension Attributes.
	3) Click the " + " button to create a new extension attribute with these settings:
	   Display Name: Outlook interface
	   Description: Reports the used outlook interface used (New/Classic).
	   Data Type: String
	   Inventory Display: Extension Attributes
	   Input Type: Script
	   Script: < Copy and paste this entire script >
	4) Save the extension attribute.
	5) Use Recon.app or "sudo jamf recon" to inventory a Mac with Office installed.
	6) Result can be viewed individually for each computer or create a smart group with the below criteria:
			a) 'Outlook interface' [LIKE] 'New Outlook'  >> to group all computers with New Outlook
			b) 'Outlook interface' [LIKE] 'Classic Outlook'  >> to group all computers with Classic Outlook
ABOUT_THIS_SCRIPT

user=`stat -f "%Su" /dev/console`

outlook_interface=`sudo -u $user defaults read com.microsoft.Outlook IsRunningNewOutlook`

if [ "$outlook_interface" == "1" ] ; then
	echo "<result>New Outlook</result>"
else
	echo "<result>Classic Outlook</result>"
fi