########################################################################
	
	# Outlook Signature Generator
	
	# Author: Michael Poff
	
	# Purpose: Automatically create outlook signature files to
	# deploy organization wide.

########################################################################

try
{
#import the active directory module which is needed for Get-ADUser
import-module activedirectory
 
$save_location = ' ' #declare where server side signature files save to
Remove-Item $save_location\* -recurse #Remove server side signature files from existing directory
 
#Define the Organizational Unit for users who you want signatures made.
$users = Get-ADUser -filter * -searchbase "OU=Regular Users,OU=Users,OU=CompanyName,DC=domaincontroller,DC=local" -Properties * -credential domain\admin -server domaincontroller.com

# Check last time signatures were created
$Now = Get-Date -format "dd MMM yyyy hh:mm tt"
$srvr_upchck = "\\server\Deployment\EmailSig\Resources\srvr_upchck.txt"

#Check for the srvr_upchck.txt file, and replace with current date/time
if ((test-path -path $srvr_upchck) -eq $True)
{
Remove-Item $srvr_upchck
write-output $Now | add-content $srvr_upchck
}
Else
{
write-output $Now | add-content $srvr_upchck
}

#Signature creation occurs within this loop.
foreach ($user in $users) 
{

	#Take active directory information and turn it into simpler PS variables.
	$name = $User.displayName
	$user_name  = $User.sAMAccountName
	$title = $User.description
	$phone = $User.telephoneNumber
	$email = $User.mail
	$division = $User.physicalDeliveryOfficeName
	$website = $User.wWWHomePage
	$image = $User.Company
	#Update the image links below if neccessary, or remove them if not images needed
	$img_1 = '<img src="http://imagelinkgoeshere.com" alt="Logo" style="width:270px;height:123px;">'
	$img_2 = '<img src="http://otherimagelinkgoeshere.com" alt="Logo" style="width:230px;height:169px;">'
	
	#The output will be whatever your above save location was plus a nested folder indicated by username (Note: It is very important that the folder here is the AD username)
	$output = "$save_location\$user_name\" + $name + ".htm"
	
	#Outlook signatures require a generic resource folder for the signatures to look right, this declares that for each user so it can be copied into the new signature folder.
	$resource_path = "$save_location\$user_name\$name"+"_files\"
	

	
	Write-Host "======================================================"
	Write-Host "Now attempting to create signature Folder For: " $user_name
	Write-Host "======================================================"
	new-item -itemtype Directory -Path $save_location"\$user_name\"
	new-item -itemtype Directory -Path $resource_path
	
#Create Actual Signature
		
	#Format based on what division the user works in, this pulls from the physicalDeliveryOfficeName AD attribute. May or may not apply to your specific situation.	
	if ($division -eq "Division 1") {$sigimg = $img_1} else {$sigimg = $img_2} #Image based on division
	if ($division -eq "Division 2") {$division = ""} #This line will blank out a division in the HTML code. For my purpose this would have resulted in a redundancy so instead of printing it twice I blanked it out.
	if (!$division) {$division = ""} #If no division leave blank
	else {$division + "<br>"} #Otherwise print the division plus an html line break
	
	#This is the signature
	$sigcontent =
@"
	<html>
	</head>
	<span style="Font-Family: Verdana; Font-Size:8pt;">
	<Strong>$Name</Strong> <br>
	$title<br>
	Phone: $phone <br>
	E-Mail: <a href="mailto:$email"> $email </a> <br>
	<br>
	<strong>Company Name</strong><br>
	$division
	Street Address <br>
	City, State Zip Code <br>
	(616) 111-1111: Office<br>
	(616) 222-2222: Fax<br>
	(877) 777-7777: Toll Free<br>
	<a href="$website">$website</a><br>
	$sigimg
	<br>
	<br>
	<p class=MsoNormal style='text-autospace:none'><span style='font-size:8.0pt;
	font-family:"Arial",sans-serif;color:black'>Learn more about <a
	href="www.example.com">Example</a> &amp; <a
	href="www.example.com">Example</a> |
	Subscribe to our <a href="http://www.youtube.com/">You Tube</a>
	channel<o:p></o:p></span></p>
	<p class=MsoNormal><o:p>&nbsp;</o:p></p>
	<p class=MsoNormal style='text-autospace:none'><span style='font-size:8.0pt;
	font-family:"Arial",sans-serif;color:grey'>Notice: This E-mail, and all subsequent replies, are covered by the Electronic Communications Privacy Act, 18 U.S.C. 2510-2521 and is legally privileged. If you received this message by mistake please notify the sender, and delete the message.
	
	</div>
	</body>
	</html>
"@
	write-output $sigcontent | add-content $output
	
	#Create colorschememapping.xml for resources folder
	$cschemecontent =
@"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:clrMap xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
"@
	write-output $cschemecontent | add-content $resource_path"\colorschememapping.xml"

	#Create filelist.xml for resources folder
	$filelistcontent =
@"
<xml xmlns:o="urn:schemas-microsoft-com:office:office">
 <o:MainFile HRef="../$name.htm"/>
 <o:File HRef="themedata.thmx"/>
 <o:File HRef="colorschememapping.xml"/>
 <o:File HRef="filelist.xml"/>
</xml>
"@
	write-output $filelistcontent | add-content $resource_path"\filelist.xml"
	
	#Copy themedata.thmx from resources folder on server
	copy-item \\server\Deployment\EmailSig\Resources\themedata.thmx $resource_path
}
Read-Host -Prompt "Press Enter to exit"
}

catch
{
    Read-Host -Prompt "The above error occurred. Press Enter to exit."
}
