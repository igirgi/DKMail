# DumbMailClient
##Simple mail-sender via Outlook automation

List of possible mail subjects and recipients must be in the file "mailer.config" in UTF8

User should:
* select subj&recipient from drop-down list
* drag files-attachemets
* write text of the message
* press button at the bottom

Settings are at the DKMail.config file in form:
<group -string>;<subject -string>;<mail - string with eMail adress>

For example  
Gases;Helium;helium@sample.com  
Gases;Hydrogen;hydro@sample.com  
Gases;Oxygen;o@sample.com  
Metals;Iron;iron@sample.com  
Metals;Gold;qqq@sample.com  
Radioactive;Uranium;cccc@sample.com  
Radioactive;Radon;ddddd@sample.com  

will be shown as  
Gases  
	Helium  
	Hydrogen  
	Oxygen  
Metals  
	Iron  
	Gold  
Radioactive  
	Uranium  
	Radon  

DKMail.config file can be placed at
1. the same directory with EXE file
2. C:\User\\<user>\AppData\Roaming\DKMail\DKMail\ – for selected user
3. C:\ProgramData\DKMail\DKMail\  - for all users

To use Outlook without confirmations - change security setting of Outlook https://support.microsoft.com/en-us/help/3189806/a-program-is-trying-to-send-an-e-mail-message-on-your-behalf-warning-i
